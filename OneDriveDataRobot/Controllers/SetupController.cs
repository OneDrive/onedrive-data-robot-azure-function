
namespace OneDriveDataRobot.Controllers
{
    using OneDriveDataRobot.Models;
    using OneDriveDataRobot.Utils;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Threading.Tasks;
    using System.Web.Http;
    using static OneDriveDataRobot.AuthHelper;
    using Microsoft.Graph;

    [Authorize]
    public class SetupController : ApiController
    {
        public async Task<IHttpActionResult> ActivateRobot()
        {
            // Setup a Microsoft Graph client for calls to the graph
            string graphBaseUrl = SettingsHelper.MicrosoftGraphBaseUrl;
            GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(async (req) =>
            {
                // Get a fresh auth token
                var authToken = await AuthHelper.GetUserAccessTokenSilentAsync(graphBaseUrl);
                req.Headers.TryAddWithoutValidation("Authorization", $"Bearer {authToken.AccessToken}");
            }));

            // Get an access token and signedInUserId
            AuthTokens tokens = null;
            try
            {
                tokens = await AuthHelper.GetUserAccessTokenSilentAsync(graphBaseUrl);
            }
            catch (Exception ex)
            {
                return Ok(new DataRobotSetup { Success = false, Error = ex.Message });
            }

            // Check to see if this user is already wired up, so we avoid duplicate subscriptions
            var robotSubscription = StoredSubscriptionState.FindUser(tokens.SignInUserId, AzureTableContext.Default.SyncStateTable);

            var notificationSubscription = new Subscription()
            {
                ChangeType = "updated",
                NotificationUrl = SettingsHelper.NotificationUrl,
                Resource = "/me/drive/root",
                ExpirationDateTime = DateTime.UtcNow.AddDays(3),
                ClientState = "SecretClientState"
            };

            Subscription createdSubscription = null;
            if (null != robotSubscription)
            {
                // See if our existing subscription can be extended to today + 3 days
                try
                {
                    createdSubscription = await client.Subscriptions[robotSubscription.SubscriptionId].Request().UpdateAsync(notificationSubscription);
                }
                catch { }
            }

            if (null == createdSubscription)
            {
                // No existing subscription or we failed to update the existing subscription, so create a new one
                createdSubscription = await client.Subscriptions.Request().AddAsync(notificationSubscription);
            }

            var results = new DataRobotSetup() { 
                SubscriptionId = createdSubscription.Id
            };

            if (robotSubscription == null)
            {
                robotSubscription = StoredSubscriptionState.CreateNew(createdSubscription.Id);
                robotSubscription.SignInUserId = tokens.SignInUserId;
            }

            // Get the latest delta URL
            var latestDeltaResponse = await client.Me.Drive.Root.Delta("latest").Request().GetAsync();
            robotSubscription.LastDeltaToken = latestDeltaResponse.AdditionalData["@odata.deltaLink"] as string;

            // Once we have a subscription, then we need to store that information into our Azure Table
            robotSubscription.Insert(AzureTableContext.Default.SyncStateTable);

            results.Success = true;
            results.ExpirationDateTime = createdSubscription.ExpirationDateTime;

            return Ok(results);
        }

        public async Task<IHttpActionResult> DisableRobot()
        {
            // Setup a Microsoft Graph client for calls to the graph
            string graphBaseUrl = SettingsHelper.MicrosoftGraphBaseUrl;
            GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(async (req) =>
            {
                // Get a fresh auth token
                var authToken = await AuthHelper.GetUserAccessTokenSilentAsync(graphBaseUrl);
                req.Headers.TryAddWithoutValidation("Authorization", $"Bearer {authToken.AccessToken}");
            }));

            // Get an access token and signedInUserId
            AuthTokens tokens = null;
            try
            {
                tokens = await AuthHelper.GetUserAccessTokenSilentAsync(graphBaseUrl);
            }
            catch (Exception ex)
            {
                return Ok(new DataRobotSetup { Success = false, Error = ex.Message });
            }

            // See if the robot was previous activated for the signed in user.
            var robotSubscription = StoredSubscriptionState.FindUser(tokens.SignInUserId, AzureTableContext.Default.SyncStateTable);

            if (null == robotSubscription)
            {
                return Ok(new DataRobotSetup { Success = true, Error = "The robot wasn't activated for you anyway!" });
            }

            // Remove the webhook subscription
            try
            {
                await client.Subscriptions[robotSubscription.SubscriptionId].Request().DeleteAsync();
            }
            catch { }

            // Remove the robotSubscription information
            robotSubscription.Delete(AzureTableContext.Default.SyncStateTable);

            return Ok(new DataRobotSetup { Success = true, Error = "The robot was been deactivated from your account." });
        }
    }
}