# Site Analytics Dashboard

This example demonstrates how to retrieve and display a comprehensive analytics dashboard for a Microsoft SharePoint site, including site details, all-time activity statistics, and last seven days activity stats using the SharePoint Sites connector.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.sites/latest) to obtain your Azure AD credentials and configure the required permissions.

2. For this example, create a `Config.toml` file in the example's root directory with your credentials:

```toml
clientId = "<Your Client ID>"
clientSecret = "<Your Client Secret>"
tenantId = "<Your Tenant ID>"
siteId = "<Your SharePoint Site ID>"
```

> **Note:** To find your `siteId`, you can use the Microsoft Graph Explorer at `https://developer.microsoft.com/en-us/graph/graph-explorer` and call `GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{siteName}`.

## Run the Example

Execute the following command to run the example. The script will print the analytics dashboard to the console.

```shell
bal run
```
