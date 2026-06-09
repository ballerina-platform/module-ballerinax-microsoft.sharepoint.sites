# Ballerina Microsoft SharePoint Connector

[![Build](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/ci.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/ci.yml)
[![Trivy](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/trivy-scan.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/trivy-scan.yml)
[![GraalVM Check](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/build-with-bal-test-graalvm.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/actions/workflows/build-with-bal-test-graalvm.yml)
[![GitHub Last Commit](https://img.shields.io/github/last-commit/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/commits/master)
[![GitHub Issues](https://img.shields.io/github/issues/ballerina-platform/ballerina-library/module/microsoft.sharepoint.sites.svg?label=Open%20Issues)](https://github.com/ballerina-platform/ballerina-library/labels/module%microsoft.sharepoint.sites)

## Overview

[Microsoft SharePoint](https://www.microsoft.com/en-us/microsoft-365/sharepoint/collaboration) is a cloud-based collaboration and content management platform that enables organizations to create, share, and manage sites, documents, and resources seamlessly across teams and enterprises.

The `ballerinax/microsoft.sharepoint.sites` package offers APIs to connect and interact with the [Microsoft SharePoint API](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0) endpoints, specifically based on [Microsoft Graph REST API v1.0](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0).

## Setup guide

To use the Microsoft SharePoint Sites connector, you must have access to the Microsoft SharePoint API through a [Microsoft Azure developer account](https://portal.azure.com/) and obtain an OAuth 2.0 client credentials (client ID, client secret, and tenant ID). If you do not have a Microsoft account, you can sign up for one [here](https://signup.microsoft.com/).

### Step 1: Create a Microsoft Account and Register an Application

1. Navigate to the [Microsoft Azure portal](https://portal.azure.com/) and sign up for an account or log in if you already have one.

2. Ensure you have an active Microsoft 365 subscription (such as Microsoft 365 Business Basic, Business Standard, Business Premium, or an Enterprise plan like E3/E5), as access to SharePoint Sites and its API requires a valid Microsoft 365 license.

3. Once logged in, navigate to **Azure Active Directory** (now called **Microsoft Entra ID**) from the left-hand navigation panel or by searching for it in the top search bar.

4. In the left sidebar, select **App registrations**, then click **New registration** to register a new application that will be used to access the SharePoint Sites API.

5. Provide a name for your application, select the appropriate supported account types (typically **Accounts in this organizational directory only**), and click **Register**.

### Step 2: Generate Client Credentials (Client ID, Client Secret, and Tenant ID)

1. Log in to the [Azure portal](https://portal.azure.com/) and navigate to **Microsoft Entra ID** > **App registrations**.

2. Select the application you registered in Step 1.

3. On the application's overview page, copy the **Application (client) ID** and the **Directory (tenant) ID** — these are your `clientId` and `tenantId` values.

4. In the left sidebar of your application, select **Certificates & secrets**, then click **New client secret**.

5. Provide a description and select an expiry duration for the secret, then click **Add**.

6. Copy the **Value** of the newly created client secret immediately — this is your `clientSecret`.

7. To grant the necessary permissions for SharePoint Sites, navigate to **API permissions** in the left sidebar, click **Add a permission**, select **Microsoft Graph**, choose **Application permissions**, and add the required permissions such as `Sites.Read.All` or `Sites.ReadWrite.All` depending on your use case.

8. Click **Grant admin consent** to approve the permissions for your organization.

> **Tip:** You must copy and store the client secret value somewhere safe immediately after creation. It won't be visible again in the Azure portal for security reasons.

## Quickstart

To use the `microsoft.sharepoint.sites` connector in your Ballerina application, update the `.bal` file as follows:

### Step 1: Import the module

```ballerina
import ballerina/oauth2;
import ballerinax/microsoft.sharepoint.sites as mssites;
```

### Step 2: Instantiate a new connector

1. Create a `Config.toml` file and configure the obtained credentials:

```toml
clientId = "<Your_Client_Id>"
clientSecret = "<Your_Client_Secret>"
refreshToken = "<Your_Refresh_Token>"
```

2. Create a `mssites:ConnectionConfig` and initialize the client:

```ballerina
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string refreshToken = ?;

final mssites:Client mssitesClient = check new({
    auth: {
        clientId,
        clientSecret,
        refreshToken
    }
});
```

### Step 3: Invoke the connector operation

Now, utilize the available connector operations.

#### Create a new list in a site

```ballerina
public function main() returns error? {
    mssites:MicrosoftGraphList newList = {
        displayName: "Project Tasks",
        list: {
            template: "genericList"
        }
    };

    mssites:MicrosoftGraphList response = check mssitesClient->sitesCreateLists("contoso.sharepoint.com,abc123", newList);
}
```

### Step 4: Run the Ballerina application

```bash
bal run
```

## Examples

The `microsoft.sharepoint.sites` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples), covering the following use cases:

1. [Site analytics dashboard](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-analytics-dashboard) - Demonstrates how to retrieve and visualize site activity and usage analytics using the SharePoint Sites connector.
2. [Site permission audit cleanup](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-permission-audit-cleanup) - Illustrates how to audit existing site permissions and remove unauthorized or outdated access entries.
3. [Content governance audit workflow](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/content-governance-audit-workflow) - Demonstrates how to automate content governance by auditing site content against organizational policies and compliance requirements.
4. [Sharepoint drive inventory](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/sharepoint-drive-inventory) - Illustrates how to enumerate and compile an inventory of all drives and their contents across SharePoint sites.

## Useful Links

- For more information go to the [`microsoft.sharepoint.sites` package](https://central.ballerina.io/ballerinax/microsoft.sharepoint.sites/latest).
- For example demonstrations of the usage, go to [Ballerina By Examples](https://ballerina.io/learn/by-example/).
- Chat live with us via our [Discord server](https://discord.gg/ballerinalang).
- Post all technical questions on Stack Overflow with the [#ballerina](https://stackoverflow.com/questions/tagged/ballerina) tag.
