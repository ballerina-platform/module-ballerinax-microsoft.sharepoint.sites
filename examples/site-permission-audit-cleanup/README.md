# Site Permission Audit Cleanup

This example demonstrates how to audit and clean up SharePoint site permissions by retrieving all permissions, identifying overly broad access grants (such as `owner`, `write`, or `fullControl` roles), and automatically revoking suspicious permissions to enforce the principle of least privilege.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint Sites connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.sites/latest) to register an Azure AD application and obtain the required credentials.

2. For this example, create a `Config.toml` file in the project root with your credentials:

```toml
clientId = "<Your Client ID>"
clientSecret = "<Your Client Secret>"
tenantId = "<Your Tenant ID>"
siteId = "<Your SharePoint Site ID>"
```

## Run the Example

Execute the following command to run the example. The script will print its audit progress and cleanup results to the console.

```shell
bal run
```
