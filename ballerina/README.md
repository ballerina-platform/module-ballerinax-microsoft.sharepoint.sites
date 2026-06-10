## Overview

[Microsoft SharePoint](https://www.microsoft.com/en-us/microsoft-365/sharepoint/collaboration) is a cloud-based collaboration and content management platform that enables organizations to create, share, and manage sites, documents, and resources securely across teams and enterprises.

The `ballerinax/microsoft.sharepoint.sites` package offers APIs to connect and interact with the [Microsoft SharePoint API](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0) endpoints, specifically based on [Microsoft Graph REST API v1.0](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0).

## Setup guide

To use the Microsoft SharePoint Sites connector, you must have access to the Microsoft SharePoint API through a [Microsoft Azure developer account](https://portal.azure.com/) and obtain an OAuth 2.0 client credentials (access token). If you do not have a Microsoft account, you can sign up for one [here](https://signup.microsoft.com/).

### Step 1: Create a Microsoft Account and Register an Application

1. Navigate to the [Microsoft Azure portal](https://portal.azure.com/) and sign in with your Microsoft account or create a new one.

2. Microsoft SharePoint Sites API access requires a **Microsoft 365 subscription** (such as Microsoft 365 Business Basic, Business Standard, Business Premium, or an Enterprise plan like E1, E3, or E5). Ensure your organization has an active Microsoft 365 plan that includes SharePoint.

3. Once signed in, navigate to **Azure Active Directory** (now called **Microsoft Entra ID**) from the left-hand navigation panel.

4. In the left menu, select **App registrations**, then click **New registration**.

5. Provide a name for your application, select the appropriate **Supported account types** (e.g., single tenant or multi-tenant), and click **Register**.

### Step 2: Generate Client Credentials (Access Token)

1. After registering your application, you will be taken to the app's **Overview** page. Note down the **Application (client) ID** and the **Directory (tenant) ID**, as these are required for authentication.

2. In the left menu of your registered application, navigate to **Certificates & secrets**.

3. Under the **Client secrets** tab, click **New client secret**.

4. Provide a description and select an expiry duration, then click **Add**.

5. Copy the **Value** of the newly created client secret immediately — this is your client secret.

6. Next, navigate to **API permissions** in the left menu and click **Add a permission**.

7. Select **Microsoft Graph**, then choose **Application permissions**.

8. Search for and add the required SharePoint permissions, such as `Sites.Read.All`, `Sites.ReadWrite.All`, or others as needed by your use case.

9. Click **Grant admin consent** for your organization to activate the permissions.

> **Tip:** You must copy and store the client secret value somewhere safe. It won't be visible again in your Azure portal settings for security reasons.

## Quickstart

To use the `microsoft.sharepoint.sites` connector in your Ballerina application, update the `.bal` file as follows:

### Step 1: Import the module

```ballerina
import ballerinax/microsoft.sharepoint.sites;
```

### Step 2: Instantiate a new connector

1. Create a `Config.toml` file and configure the obtained credentials:

```toml
clientId = "<Your_Client_Id>"
clientSecret = "<Your_Client_Secret>"
tenantId = "<Your_Tenant_Id>"
```

2. Create a `sites:ConnectionConfig` and initialize the client:

```ballerina
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string tenantId = ?;

final sites:Client msClient = check new ({
    auth: {
        clientId,
        clientSecret,
        tokenUrl: "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token",
        scopes: ["https://graph.microsoft.com/.default"]
    }
});
```

### Step 3: Invoke the connector operation

Now, utilize the available connector operations.

#### Create a new list in a site

```ballerina
public function main() returns error? {
    sites:MicrosoftGraphList newList = {
        displayName: "Project Documents",
        list: {
            template: "documentLibrary",
            contentTypesEnabled: true
        }
    };

    sites:MicrosoftGraphList response = check msClient->createLists("contoso.sharepoint.com,abc123,def456", newList);
}
```

### Step 4: Run the Ballerina application

```bash
bal run
```

## Examples

The `microsoft.sharepoint.sites` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples), covering the following use cases:

1. [Site analytics dashboard](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-analytics-dashboard) - Demonstrates how to retrieve and visualize SharePoint site analytics data using the connector.
2. [Site permission audit cleanup](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-permission-audit-cleanup) - Illustrates auditing and cleaning up site permissions to maintain proper access control.
3. [Content governance audit workflow](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/content-governance-audit-workflow) - Demonstrates how to automate content governance auditing workflows across SharePoint sites.
4. [Sharepoint drive inventory](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/sharepoint-drive-inventory) - Illustrates how to enumerate and inventory drives and their contents across SharePoint sites.
