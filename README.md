# Ballerina Microsoft SharePoint Connector

[![Build](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/ci.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/ci.yml)
[![Trivy](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/trivy-scan.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/trivy-scan.yml)
[![GraalVM Check](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/build-with-bal-test-graalvm.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/build-with-bal-test-graalvm.yml)
[![GitHub Last Commit](https://img.shields.io/github/last-commit/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/commits/master)
[![GitHub Issues](https://img.shields.io/github/issues/ballerina-platform/ballerina-library/module/microsoft.sharepoint.sites.svg?label=Open%20Issues)](https://github.com/ballerina-platform/ballerina-library/labels/module%microsoft.sharepoint.sites)

## Overview
[Microsoft SharePoint](https://www.microsoft.com/en-us/microsoft-365/sharepoint/collaboration) is a cloud-based collaboration and content management platform that enables organizations to create, share, and manage sites, documents, and resources securely across teams and enterprises.

The `ballerinax/microsoft.sharepoint.sites` package offers APIs to connect and interact with the [Microsoft SharePoint Sites API](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0) endpoints, specifically based on [Microsoft Graph REST API v1.0](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0).

## Setup guide
To use the Microsoft SharePoint Sites connector, you must have access to the Microsoft SharePoint API through a [Microsoft Azure developer account](https://portal.azure.com/) and obtain an OAuth 2.0 client credentials (client ID, client secret, and tenant ID). If you do not have a Microsoft account, you can sign up for one [here](https://signup.microsoft.com/).

### Step 1: Create a Microsoft Account and Register an Application

1. Navigate to the [Microsoft Azure portal](https://portal.azure.com/) and sign up for an account or log in if you already have one.

2. Ensure you have an active Microsoft 365 subscription that includes SharePoint Online (such as Microsoft 365 Business Basic, Business Standard, Business Premium, or any Enterprise E1/E3/E5 plan), as access to the SharePoint Sites API requires a valid Microsoft 365 tenant with SharePoint Online enabled.

3. In the Azure portal, navigate to **Azure Active Directory** (now called **Microsoft Entra ID**) from the left-hand navigation menu.

4. Select **App registrations** and then click **New registration** to register a new application that will be used to access the SharePoint Sites API.

5. Provide a name for your application, select the appropriate supported account types, and click **Register**.

### Step 2: Generate API Credentials (Client ID, Client Secret, and Tenant ID)

1. Log in to the [Microsoft Azure portal](https://portal.azure.com/).

2. Navigate to **Microsoft Entra ID** (formerly Azure Active Directory) from the left-hand navigation menu and select **App registrations**.

3. Select the application you registered in Step 1. On the application's **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID** — these serve as your `clientId` and `tenantId`.

4. In the left-hand menu of your application, select **Certificates & secrets**, then click **New client secret**.

5. Provide a description and select an expiry duration for the secret, then click **Add**. Copy the **Value** of the newly created client secret immediately — this is your `clientSecret`.

6. To grant the necessary permissions for SharePoint Sites access, navigate to **API permissions** in the left-hand menu, click **Add a permission**, select **Microsoft Graph**, and then choose **Application permissions**. Add the required permissions such as `Sites.Read.All` or `Sites.ReadWrite.All` depending on your use case.

7. Click **Grant admin consent** for your organization to activate the permissions.

> **Tip:** You must copy and store the client secret value somewhere safe immediately after creation. It won't be visible again in the Azure portal for security reasons.

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
    sites:List newList = {
        displayName: "Project Tasks",
        list: {
            template: "genericList"
        }
    };

    sites:List response = check msClient->createLists("contoso.sharepoint.com,abc123,def456", newList);
}
```

### Step 4: Run the Ballerina application

```bash
bal run
```

## Examples
The `microsoft.sharepoint.sites` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples), covering the following use cases:

1. [Site analytics dashboard](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-analytics-dashboard) - Demonstrates how to retrieve and visualize SharePoint site analytics data using the connector.
2. [Site permission audit cleanup](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-permission-audit-cleanup) - Illustrates how to audit and clean up site permissions to maintain proper access control.
3. [Content governance audit workflow](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/content-governance-audit-workflow) - Demonstrates how to automate content governance auditing workflows across SharePoint sites.
4. [Sharepoint drive inventory](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/sharepoint-drive-inventory) - Illustrates how to enumerate and compile an inventory of drives and their contents across SharePoint sites.

## Useful Links

* For more information go to the [`microsoft.sharepoint.sites` package](https://central.ballerina.io/ballerinax/microsoft.sharepoint.sites/latest).
* For example demonstrations of the usage, go to [Ballerina By Examples](https://ballerina.io/learn/by-example/).
* Chat live with us via our [Discord server](https://discord.gg/ballerinalang).
* Post all technical questions on Stack Overflow with the [#ballerina](https://stackoverflow.com/questions/tagged/ballerina) tag.
