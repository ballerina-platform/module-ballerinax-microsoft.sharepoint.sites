# Content Governance Audit Workflow

This example demonstrates how to automate a content governance audit on a Microsoft SharePoint site by retrieving all content types and their column definitions, then enforcing organizational governance standards by automatically adding a mandatory "Compliance Category" column if it is missing.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.sites/latest) to register an Azure AD application and obtain the required credentials.

2. For this example, create a `Config.toml` file in the project root with your credentials:

```toml
tenantId = "<Your Tenant ID>"
clientId = "<Your Client ID>"
clientSecret = "<Your Client Secret>"
siteId = "<Your SharePoint Site ID>"
targetContentTypeId = "<Your Target Content Type ID>"
```

## Run the Example

Execute the following command to run the example. The script will print its progress to the console as it performs each step of the audit workflow.

```shell
bal run
```
