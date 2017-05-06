# OneDrive Data Robot Azure Function Sample Code

This project provides an example implementation for connecting Azure Functions to OneDrive to enable your solution to react to changes in files in OneDrive in nearly instantly.

The project consists of two parts:

* An [Azure Function](https://azure.microsoft.com/services/functions/) definition that handles the processing of webhook notifications and the resulting work from those notifications
* An ASP.NET MVC application that activates and deactivated the OneDrive Data Robot for a signed in user.

In this scenario, the benefit of using Azure Function is that the load is required by the data robot is dynamic and hard to predict.
Instead of scaling out an entire web application to handle the load, Azure Functions can scale dynamically based on the load required at any given time.
This provides a cost-savings measure for hosting the application while still ensuring high performance results.

## Getting Started

To get started with the sample, you need to complete the following steps:

1. Register a new application with Azure Active Directory, generate an app password, and provide a redirect URI for the application.
2. Create a new Azure Function and upload the code files in the **AzureFunction** folder into the function's definition.
3. Run the sample project and sign-in with your Office 365 account and activate the data robot by clicking the **Activate** button.
4. Navigate to OneDrive and modify a file (see below for details).
5. Watch the data robot update the file automatically.

### Register a new application

To register a new application with Azure Active Directory, log into the [Azure Portal](https://portal.azure.com).

After logging into the Azure Portal, follow these steps to register the sample application:

1. Navigate to the **Azure Active Directory** module.
2. Select **App registrations** and click **New application registration**.
    1. Type the name of your file handler application.
    2. Ensure **Application Type** is set to **Web app / API**
    3. Enter a sign-on URL for your application, for this sample use `https://localhost:44382`.
    4. Click **Create** to create the app.
3. After the app has been created successfully, select the app from the list of applications. It should be at the bottom of the list.
4. Copy the **Application ID** for the app you registered and paste it into two places:
    * In the Web.config file on the line: `<add key="ida:ClientId" value="[ClientId]" />`
    * In the run.csx file on the line: `private const string idaClientId = "[ClientId]";`
5. Configure the application settings for this sample:
    1. Select **Reply URLs** and ensure that `https://localhost:44382` is listed.
    2. Select **Required Permissions** and then **Add**.
    3. Select **Select an API** and then choose **Microsoft Graph** and click **Select**.
    4. Find the permission **Have full access to user files** and check the box next to it, then click **Select**, and then **Done**.
    5. Select **Keys** and generate a new application key by entering a description for the key, selecting a duration, and then click **Save**. Copy the value of the displayed key since it will only be displayed once. Paste it into two places:
       * In the Web.config file on the line: `<add key="ida:ClientSecret" value="[ClientSecret]" />`
       * In the run.csx file on the line: `private const string idaClientSecret = "[ClientSecret]"`

### Create an Azure Function

To create the Azure Function portion of this sample, you will need to be logged into the [Azure Portal](https://portal.azure.com).

1. Click **New** and select **Function App** in the Azure Portal.
2. Enter a name for your function app, such as datarobot99. The name of your function app must be unique, so you'll need to modify the name to find a unique one.
3. Select your existing Azure subscription, desired resource group, hosting plan, and location for the Azure Function app.
4. Choose to create a new storage for this azure function, and provide a unique name for the storage.
5. Click create to have the Azure Portal create everything for you.

After the required components have been provisioned, click on the **App Services** module in the portal and find the Azure Function App we just created.

#### Register a new Azure Function

To create a new Azure Function application and setup a function for this project:

1. Click the `+` next to **Functions**
2. Select the **Webhook + API** scenario, choose **CSharp** as the language, and then **Create this function**.
3. Copy the code from [run.csx][] and paste it into the code editor and then click **Save**
4. On the right side, select **View files** to expand the files that make up this function
5. Click **Upload** and then navigate to the `project.json` file in the AzureFunction folder and upload it. This file configures the dependencies for the Azure Function, and will add the Azure authentication library and the Microsoft Graph SDK to the function project.
6. Click **Integrate** on the left side, under the **HttpTriggerCSharp1** function name (or if your function has a different name, select **Integrate** under that). Configure your function accordingly:
   1. Select **HTTP (req)** under **Triggers** and configure the values accordingly then click **Save**.
     * **Allowed HTTP methods:** Selected methods
     * **Mode:** Standard
     * **Request parameter name:** req
     * **Route template:** default value (empty)
     * **Authorization level:** Anonymous
     * **Selected HTTP methods:** Uncheck everything except POST
   2. Select **New Input** on the Inputs column and choose **Azure Table Storage**. Set the parameters accordingly:
     * **Table parameter name:** syncStateTable
     * **Table name:** syncState
     * **Storage account connection:** Click **new** and then connect it to the storage connection you created (named something like datarobot99ae20).
     * Leave the other parameters with their default values, and then click **Save**.
   3. Select **New Input** again, and again choose **Azure Table Storage**. Set the parameters accordingly:
     * **Table parameter name:** tokenCacheTable
     * **Table name:** tokenCache
     * **Storage account connection:** Choose the existing storage connection from step 2, something like datarobot99ae20_STORAGE.
     * Leave the other parameters with their default values, and then click **Save**.
7. Click back on the function name in the left navigation column to bring up the code editor.
8. Click **Get function URL** and copy the URL for this function. 
   * Paste this value into the **Web.config** file on the line: `<add key="ida:NotificationUrl" value="[azureFunctionServiceUrl]" />`


### Run the project and sign-in

Now that everything is properly configured, open the web project in Visual Studio and press F5 launch the project in the debugger.

1. Sign in to the data robot project and authorize the application to have access to the data in your OneDrive.
2. After you authorize the data robot, you should see a Subscription ID and Expiration date/time.
   These values are returned from the Microsoft Graph webhook notification subscription that powers the data robot.
   By default the expiration time is 3 days from when the robot is activated.

If no value is returned, check to ensure that the notification URL is correct in the Web.config file, and verify in the Azure Function console that you are seeing a request successfully processed by your function code.


### Navigate to OneDrive and try out the data robot

This sample data robot uses a Web API to insert live stock quotes into Excel files while you are editing them.
To invoke the data robot and ask it to insert a stock quote into a cell, you can do the following:

1. Open your OneDrive ({tenant}.onedrive.com).
2. Click **New** then **Excel Workbook**. A new Excel workbook will open in the Excel web application.
3. The data robot looks for the keyword `!roland` followed by what you are asking for:
   * To retrieve a stock quote, use `!roland MSFT stock quote` where MSFT can be replaced with the stock ticker symbol of your choice.

If the data robot is unable to retrieve a real stock quote, it will make one up so the demo always works.

**Note:** OneDrive webhooks can take up to 5 minutes to be delivered, depending on load and other conditions.
As a result, requests in the workbook may take a few minutes to be updated with real data.

## Related references

For more information about Microsoft Graph API, see [Microsoft Graph](https://graph.microsoft.com).

## License

See [License](LICENSE.txt) for the license agreement convering this sample code.
