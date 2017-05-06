#r "Newtonsoft.Json"

using System;
using System.Net;
using Newtonsoft.Json;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using Microsoft.Graph;

private const string TriggerWord = "!odbot";
private const string idaClientId = "[ClientId]";
private const string idaClientSecret = "[ClientSecret]";
private const string idaAuthorityUrl = "https://login.microsoftonline.com/common";
private const string idaMicrosoftGraphUrl = "https://graph.microsoft.com";

// Main entry point for our Azure Function. Listens for webhooks from OneDrive and responds to the webhook with a 204 No Content.
public static async Task<object> Run(HttpRequestMessage req, CloudTable syncStateTable, CloudTable tokenCacheTable, TraceWriter log)
{
    log.Info($"Webhook was triggered!");

    // Handle validation scenario for creating a new webhook subscription
    Dictionary<string, string> qs = req.GetQueryNameValuePairs()
                            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
    if (qs.ContainsKey("validationToken"))
    {
        var token = qs["validationToken"];
        log.Info($"Responding to validationToken: {token}");
        return PlainTextResponse(token);
    }

    // If not the validation scenario, read the body of the request and parse the notification
    string jsonContent = await req.Content.ReadAsStringAsync();
    log.Verbose($"Raw request content: {jsonContent}");
    
    // Since webhooks can be batched together, loop over all the notifications we receive and process them individually.
    // In the real world, this shouldn't be done in the request handler, but rather queued to be done later.
    dynamic data = JsonConvert.DeserializeObject(jsonContent);
    if (data.value != null)
    {
        foreach (var subscription in data.value)
        {
            var clientState = subscription.clientState;
            var resource = subscription.resource;
            string subscriptionId = (string)subscription.subscriptionId;
            log.Info($"Notification for subscription: '{subscriptionId}' Resource: '{resource}', clientState: '{clientState}'");

            // Process the individual subscription information
            bool exists = await ProcessSubscriptionNotificationAsync(subscriptionId, syncStateTable, tokenCacheTable, log);
            if (!exists)
            {
                return req.CreateResponse(HttpStatusCode.Gone);
            }
        }
        return req.CreateResponse(HttpStatusCode.NoContent);
    }

    log.Info($"Request was incorrect. Returning bad request.");
    return req.CreateResponse(HttpStatusCode.BadRequest);
}

// Do the work to retrieve deltas from this subscription and then find any changed Excel files
private static async Task<bool> ProcessSubscriptionNotificationAsync(string subscriptionId, CloudTable syncStateTable, CloudTable tokenCacheTable, TraceWriter log)
{
    // Retrieve our stored state from an Azure Table
    StoredSubscriptionState state = StoredSubscriptionState.Open(subscriptionId, syncStateTable);
    if (state == null)
    {
        log.Info($"Missing data for subscription '{subscriptionId}'.");
        return false;
    }

    log.Info($"Found subscription '{subscriptionId}' with lastDeltaUrl: '{state.LastDeltaToken}'.");

    GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(async (request) => {
        string accessToken = await RetrieveAccessTokenAsync(state.SignInUserId, tokenCacheTable, log);
        request.Headers.TryAddWithoutValidation("Authorization", $"Bearer {accessToken}");
    }));

    // Query for items that have changed since the last notification was received
    List<string> changedExcelFileIds = await FindChangedExcelFilesInOneDrive(state, client, log);

    // Do work on the changed files
    foreach (var file in changedExcelFileIds)
    {
        log.Info($"Processing changes in file: {file}");
        try 
        {
            string sessionId = await StartOrResumeWorkbookSessionAsync(client, file, syncStateTable, log);
            log.Info($"File {file} is using sessionId: {sessionId}");
            await ScanExcelFileForPlaceholdersAsync(client, file, sessionId, log);
        } 
        catch (Exception ex)
        {
            log.Info($"Exception processing file: {ex.Message}");
        }
    }
    
    // Update our saved state for this subscription
    state.Insert(syncStateTable);
    return true;
}

// Use the Excel REST API to look for queries that we can replace with real data
private static async Task ScanExcelFileForPlaceholdersAsync(GraphServiceClient client, string fileId, string workbookSession, TraceWriter log)
{
    const string SheetName = "Sheet1";

    var dataRequest = client.Me.Drive.Items[fileId].Workbook.Worksheets[SheetName].UsedRange().Request();
    if (null != workbookSession)
    {
        dataRequest.Headers.Add(new HeaderOption("workbook-session-id", workbookSession));
    }
    var data = await dataRequest.Select("address,cellCount,columnCount,values").GetAsync();

    var usedRangeId = data.Address;
    var sendPatch = false;
    dynamic range = data.Values;

    for (int rowIndex = 0; rowIndex < range.Count; rowIndex++)
    {
        var rowValues = range[rowIndex];
        for (int columnIndex = 0; columnIndex < rowValues.Count; columnIndex++)
        {
            var value = (string)rowValues[columnIndex];
            if (value.StartsWith($"{TriggerWord} "))
            {
                log.Info($"Found cell [{rowIndex},{columnIndex}] with value: {value} ");
                rowValues[columnIndex] = await ReplacePlaceholderValueAsync(value);
                sendPatch = true;
            }
            else
            {
                // Replace the value with null so we don't overwrite anything on the PATCH
                rowValues[columnIndex] = null;
            }
        }
    }

    if (sendPatch)
    {
        log.Info($"Updating file {fileId} with replaced values.");
        await client.Me.Drive.Items[fileId].Workbook.Worksheets[SheetName].Range(data.Address).Request().PatchAsync(data);
    }
}

// Make a request to retrieve a response based on the input value
private static async Task<string> ReplacePlaceholderValueAsync(string inputValue)
{
    // This is merely an example. A real solution would do something much richer
    if (inputValue.StartsWith($"{TriggerWord} ") && inputValue.EndsWith(" stock quote"))
    {
        // For demo purposes, return a random value instead of the stock quote
        Random rndNum = new Random(int.Parse(Guid.NewGuid().ToString().Substring(0, 8), System.Globalization.NumberStyles.HexNumber));
        return rndNum.Next(20, 100).ToString(); 
    }
    
    return inputValue;
}

// Request the delta stream from OneDrive to find files that have changed between notifications for this account
        // Request the delta stream from OneDrive to find files that have changed between notifications for this account
private static async Task<List<string>> FindChangedExcelFilesInOneDrive(StoredSubscriptionState state, GraphServiceClient client, TraceWriter log)
{
    const string DefaultDeltaToken = idaMicrosoftGraphUrl + "/v1.0/me/drive/root/delta?token=latest";

    // We default to reading the "latest" state of the drive, so we don't have to process all the files in the drive
    // when a new subscription comes in.
    string deltaUrl = DefaultDeltaToken;
    if (!String.IsNullOrEmpty(state.LastDeltaToken))
    {
        deltaUrl = state.LastDeltaToken;
    }

    const int MaxLoopCount = 50;
    List<string> changedFileIds = new List<string>();

    IDriveItemDeltaRequest request = new DriveItemDeltaRequest(deltaUrl, client, null);

    // Only allow reading 50 pages, if we read more than that, we're going to cancel out
    for (int loopCount = 0; loopCount < MaxLoopCount && request != null; loopCount++)
    {
        log.Info($"Making request for '{state.SubscriptionId}' to '{deltaUrl}' ");
        var deltaResponse = await request.GetAsync();

        log.Verbose($"Found {deltaResponse.Count} files changed in this page.");
        try
        {
            var changedExcelFiles = (from f in deltaResponse
                                        where f.File != null && f.Name != null && f.Name.EndsWith(".xlsx") && f.Deleted == null
                                        select f.Id);
            log.Info($"Found {changedExcelFiles.Count()} changed Excel files in this page.");
            changedFileIds.AddRange(changedExcelFiles);
        }
        catch (Exception ex)
        {
            log.Info($"Exception enumerating changed files: {ex.ToString()}");
            throw;
        }

        
        if (null != deltaResponse.NextPageRequest)
        {
            request = deltaResponse.NextPageRequest;
        }
        else if (null != deltaResponse.AdditionalData["@odata.deltaLink"])
        {
            string deltaLink = (string)deltaResponse.AdditionalData["@odata.deltaLink"];
            log.Verbose($"All changes requested, nextDeltaUrl: {deltaLink}");
            state.LastDeltaToken = deltaLink;
            return changedFileIds;
        }
        else
        {
            request = null;
        }
    }

    // If we exit the For loop without returning, that means we read MaxLoopCount pages without finding a deltaToken
    log.Info($"Read through MaxLoopCount pages without finding an end. Too much data has changed.");
    state.LastDeltaToken = DefaultDeltaToken;

    return changedFileIds;
    
}

/// <summary>
/// Ensure that we're working out of a shared session if multiple updates to a file are happening frequently.
/// This improves performance and ensures consistency of the data between requests.
/// </summary>
private static async Task<string> StartOrResumeWorkbookSessionAsync(GraphServiceClient client, string fileId, CloudTable table, TraceWriter log)
{
    const string userId = "1234";

    var fileItem = FileHistory.Open(userId, fileId, table);
    if (null == fileItem)
    {
        log.Info($"No existing Excel session found for file: {fileId}");
        fileItem = FileHistory.CreateNew(userId, fileId);
    }

    if (!string.IsNullOrEmpty(fileItem.ExcelSessionId))
    {
        // Verify session is still available
        TimeSpan lastUsed = DateTime.UtcNow.Subtract(fileItem.LastAccessedDateTime);
        if (lastUsed.TotalMinutes < 5)
        {
            fileItem.LastAccessedDateTime = DateTime.UtcNow;
            try
            {
                // Attempt to update the cache, but if we get a conflict, just ignore it
                fileItem.Insert(table);
            }
            catch { }
            log.Info($"Reusing existing session for file: {fileId}");
            return fileItem.ExcelSessionId;
        }
    }

    string sessionId = null;
    try
    {
        // Create a new workbook session
        var session = await client.Me.Drive.Items[fileId].Workbook.CreateSession(true).Request().PostAsync();
        fileItem.LastAccessedDateTime = DateTime.UtcNow;
        fileItem.ExcelSessionId = session.Id;
        log.Info($"Reusing existing session for file: {fileId}");
        sessionId = session.Id;
    }
    catch { }

    try 
    {
        fileItem.Insert(table);
    }
    catch { }

    return sessionId;
}

private static HttpResponseMessage PlainTextResponse(string text)
{
    HttpResponseMessage response = new HttpResponseMessage()
    {
        StatusCode = HttpStatusCode.OK,
        Content = new StringContent(
                text,
                System.Text.Encoding.UTF8,
                "text/plain"
            )
    };
    return response;
}

// Retrieve a new access token from AAD
private static async Task<string> RetrieveAccessTokenAsync(string signInUserId, CloudTable tokenCacheTable, TraceWriter log) 
{
    log.Verbose($"Retriving new accessToken for signInUser: {signInUserId}");

    var tokenCache = new AzureTableTokenCache(signInUserId, tokenCacheTable);
    var authContext = new AuthenticationContext(idaAuthorityUrl, tokenCache);

    try 
    {
        var userCredential = new UserIdentifier(signInUserId, UserIdentifierType.UniqueId);
        // Don't really store your clientId and clientSecret in your code. Read these from configuration.
        var clientCredential = new ClientCredential(idaClientId, idaClientSecret);
        var authResult = await authContext.AcquireTokenSilentAsync(idaMicrosoftGraphUrl, clientCredential, userCredential);
        return authResult.AccessToken;
    }
    catch (AdalSilentTokenAcquisitionException ex)
    {
        log.Info($"ADAL Error: Unable to retrieve access token: {ex.Message}");
        return null;
    }
}

/*** SHARED CODE STARTS HERE ***/

/// <summary>
/// Persists information about a subscription, userId, and deltaToken state. This class is shared between the Azure Function and the bootstrap project
/// </summary>
public class StoredSubscriptionState : TableEntity
{
    public StoredSubscriptionState()
    {
        this.PartitionKey = "AAA";
    }

    public string SignInUserId { get; set; }
    public string LastDeltaToken { get; set; }
    public string SubscriptionId { get; set; }
    public string ExcelSessionId { get; set; }


    public static StoredSubscriptionState CreateNew(string subscriptionId)
    {
        var newState = new StoredSubscriptionState();
        newState.RowKey = subscriptionId;
        newState.SubscriptionId = subscriptionId;
        return newState;
    }

    public void Insert(CloudTable table)
    {
        TableOperation insert = TableOperation.InsertOrReplace(this);
        table.Execute(insert);
    }

    public static StoredSubscriptionState Open(string subscriptionId, CloudTable table)
    {
        TableOperation retrieve = TableOperation.Retrieve<StoredSubscriptionState>("AAA", subscriptionId);
        TableResult results = table.Execute(retrieve);
        return (StoredSubscriptionState)results.Result;
    }
}

/// <summary>
/// Keep track of file specific information for a short period of time, so we can avoid repeatedly acting on the same file
/// </summary>
public class FileHistory : TableEntity
{
    public FileHistory()
    {
        this.PartitionKey = "BBB";
    }

    public string ExcelSessionId { get; set; }
    public DateTime LastAccessedDateTime { get; set; }

    public static FileHistory CreateNew(string userId, string fileId)
    {
        var newState = new FileHistory();
        newState.RowKey = $"{userId},{fileId}";
        return newState;
    }

    public void Insert(CloudTable table)
    {
        TableOperation insert = TableOperation.InsertOrReplace(this);
        table.Execute(insert);
    }

    public static FileHistory Open(string userId, string fileId, CloudTable table)
    {
        TableOperation retrieve = TableOperation.Retrieve<FileHistory>("BBB", $"{userId},{fileId}");
        TableResult results = table.Execute(retrieve);
        return (FileHistory)results.Result;
    }
}

/// <summary>
/// ADAL TokenCache implementation that stores the token cache in the provided Azure CloudTable instance.
/// This class is shared between the Azure Function and the bootstrap project.
/// </summary>
public class AzureTableTokenCache : TokenCache
{
    private readonly string signInUserId;
    private readonly CloudTable tokenCacheTable;

    private TokenCacheEntity cachedEntity;      // data entity stored in the Azure Table

    public AzureTableTokenCache(string userId, CloudTable cacheTable)
    {
        signInUserId = userId;
        tokenCacheTable = cacheTable;

        this.AfterAccess = AfterAccessNotification;

        cachedEntity = ReadFromTableStorage();
        if (null != cachedEntity)
        {
            Deserialize(cachedEntity.CacheBits);
        }
    }

    private TokenCacheEntity ReadFromTableStorage()
    {
        TableOperation retrieve = TableOperation.Retrieve<TokenCacheEntity>(TokenCacheEntity.PartitionKeyValue, signInUserId);
        TableResult results = tokenCacheTable.Execute(retrieve);
        return (TokenCacheEntity)results.Result;
    }

    private void AfterAccessNotification(TokenCacheNotificationArgs args)
    {
        if (this.HasStateChanged)
        {
            if (cachedEntity == null)
            {
                cachedEntity = new TokenCacheEntity();
            }
            cachedEntity.RowKey = signInUserId;
            cachedEntity.CacheBits = Serialize();
            cachedEntity.LastWrite = DateTime.Now;

            TableOperation insert = TableOperation.InsertOrReplace(cachedEntity);
            tokenCacheTable.Execute(insert);

            this.HasStateChanged = false;
        }
    }

    /// <summary>
    /// Representation of the data stored in the Azure Table
    /// </summary>
    private class TokenCacheEntity : TableEntity
    {
        public const string PartitionKeyValue = "tokenCache";
        public TokenCacheEntity()
        {
            this.PartitionKey = PartitionKeyValue;
        }

        public byte[] CacheBits { get; set; }
        public DateTime LastWrite { get; set; }
    }

}
