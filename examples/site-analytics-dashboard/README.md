# Site Analytics Dashboard

This example demonstrates how to retrieve and display a comprehensive analytics dashboard for a Microsoft SharePoint site, including site details, all-time activity statistics, and last seven days activity statistics using the Microsoft SharePoint Sites connector.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint Sites connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.sites/latest) to obtain the required credentials and configure your Azure AD application with the appropriate permissions.

2. For this example, create a `Config.toml` file in the project root with your credentials:

```toml
clientId = "<Your Client ID>"
clientSecret = "<Your Client Secret>"
tenantId = "<Your Tenant ID>"
siteId = "<Your SharePoint Site ID>"
```

## Run the Example

Execute the following command to run the example. The script will print the analytics dashboard to the console.

```shell
bal run
```
