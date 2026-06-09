# SharePoint Drive Inventory

This example demonstrates how to discover all SharePoint sites in an organization and collect a comprehensive drive inventory for each site, including primary drive details and all document libraries, to support governance reporting and storage capacity planning.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.sites/latest) to register an Azure AD application and obtain the required credentials.

2. Create a `Config.toml` file in the example's root directory and provide your Microsoft Azure AD credentials:

```toml
clientId = "<Your Client ID>"
clientSecret = "<Your Client Secret>"
tenantId = "<Your Tenant ID>"
```

## Run the Example

Execute the following command to run the example. The script will print its progress and the full drive inventory to the console.

```shell
bal run
```
