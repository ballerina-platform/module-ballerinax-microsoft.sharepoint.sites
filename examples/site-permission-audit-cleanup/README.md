# Site Permission Audit Cleanup

This example demonstrates how to audit and clean up SharePoint site permissions by retrieving all permissions for a given site, flagging overly broad access grants (such as `owner`, `write`, or `fullcontrol` roles), and automatically revoking the first suspicious permission found to enforce the principle of least privilege.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.sites/latest) to register an Azure AD application and obtain the required credentials.

2. **Configuration**

   Create a `Config.toml` file in the example's root directory and provide your credentials:

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
