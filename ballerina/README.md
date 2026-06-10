## Overview

[Microsoft SharePoint](https://www.microsoft.com/en-us/microsoft-365/sharepoint/collaboration) is a cloud-based collaboration and content management platform that enables organizations to create, manage, and share sites, documents, and data across teams and enterprises.

The `ballerinax/microsoft.sharepoint.sites` package offers APIs to connect and interact with the [Microsoft SharePoint API](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0) endpoints, specifically based on [Microsoft Graph REST API v1.0](https://learn.microsoft.com/en-us/graph/overview?view=graph-rest-1.0).

## Setup guide

To use the Microsoft SharePoint Sites connector, you must have access to the Microsoft SharePoint API through a [Microsoft Azure developer account](https://portal.azure.com/) and obtain OAuth 2.0 client credentials (client ID, client secret, and tenant ID). If you do not have a Microsoft account, you can sign up for one [here](https://signup.microsoft.com/).

### Step 1: Create a Microsoft Account and Register an Application

1. Navigate to the [Microsoft website](https://www.microsoft.com/) and sign up for an account or log in if you already have one.

2. Go to the [Azure Portal](https://portal.azure.com/) and ensure you have access to an active Azure subscription. Microsoft SharePoint API access requires a **Microsoft 365 Business Basic, Business Standard, Business Premium, or Enterprise (E1, E3, E5)** plan, as the SharePoint API is restricted to users on these plans.

3. In the Azure Portal, navigate to **Azure Active Directory** (now called **Microsoft Entra ID**) from the left-hand menu.

4. Select **App registrations** from the left panel, then click **New registration**.

5. Enter a name for your application, select the appropriate **Supported account types** (e.g., single tenant or multi-tenant), and click **Register**.

### Step 2: Obtain Client Credentials

1. Log in to the [Azure Portal](https://portal.azure.com/) and navigate to **Microsoft Entra ID** > **App registrations**, then select your registered application.

2. From the application overview page, copy and save the **Application (client) ID** and the **Directory (tenant) ID** — you will need these later.

3. In the left panel, select **Certificates & secrets**, then click **New client secret**.

4. Enter a description for the secret, select an expiry duration, and click **Add**.

5. Copy the **Value** of the newly created client secret immediately and store it securely.

6. To grant the required permissions, navigate to **API permissions** in the left panel, click **Add a permission**, select **Microsoft Graph**, and then choose **Application permissions**. Add the necessary SharePoint-related permissions such as `Sites.Read.All` or `Sites.ReadWrite.All` depending on your use case.

7. Click **Grant admin consent** to approve the permissions for your organization.

> **Tip:** You must copy and store the client secret value somewhere safe immediately after creation. It won't be visible again in the Azure Portal for security reasons.

## Quickstart

To use the `microsoft.sharepoint.sites` connector in your Ballerina application, update the `.bal` file as follows:

### Step 1: Import the module

```ballerina
import ballerinax/microsoft.sharepoint.sites as mssites;
```

### Step 2: Instantiate a new connector

1. Create a `Config.toml` file and configure the obtained credentials:

```toml
clientId = "<Your_Client_Id>"
clientSecret = "<Your_Client_Secret>"
tenantId = "<Your_Tenant_Id>"
```

2. Create a `mssites:ConnectionConfig` and initialize the client:

```ballerina
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string tenantId = ?;

final mssites:Client mssitesClient = check new ({
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
    mssites:MicrosoftGraphList newList = {
        displayName: "Project Documents",
        list: {
            template: "documentLibrary",
            contentTypesEnabled: true
        }
    };

    mssites:MicrosoftGraphList response = check mssitesClient->sitesCreateLists("contoso.sharepoint.com,abc123,def456", newList);
}
```

### Step 4: Run the Ballerina application

```bash
bal run
```

## Examples

The `microsoft.sharepoint.sites` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples), covering the following use cases:

1. [Site analytics dashboard](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-analytics-dashboard) - Demonstrates how to retrieve and visualize SharePoint site analytics data using the connector.
2. [Site permission audit cleanup](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-permission-audit-cleanup) - Illustrates how to audit and clean up permissions across SharePoint sites.
3. [Content governance audit workflow](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/content-governance-audit-workflow) - Demonstrates how to automate content governance auditing workflows for SharePoint sites.
4. [Sharepoint drive inventory](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/sharepoint-drive-inventory) - Illustrates how to enumerate and inventory drives and their contents across SharePoint sites.
