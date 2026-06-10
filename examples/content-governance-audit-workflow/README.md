# Content Governance Audit Workflow

This example demonstrates how to automate a content governance audit on a Microsoft SharePoint site by retrieving all site content types and their column definitions, then enforcing governance standards by automatically adding a mandatory "Compliance Category" column to a specified content type if it is missing.

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

> **Note:**
> - `tenantId` — The Azure Active Directory tenant ID associated with your Microsoft 365 organization.
> - `clientId` — The application (client) ID of your registered Azure AD app.
> - `clientSecret` — The client secret generated for your Azure AD app.
> - `siteId` — The unique identifier of the SharePoint site to audit (e.g., `contoso.sharepoint.com,<site-guid>,<web-guid>`).
> - `targetContentTypeId` — The ID of the specific content type whose columns you want to audit and enforce (e.g., `0x0101...`).

## Run the Example

Execute the following command to run the example. The script will print its progress to the console, listing all content types, inspecting columns, and adding the "Compliance Category" column if it does not already exist.

```shell
bal run
```

### Expected Output

```
=== Content Governance Audit Workflow ===

Step 1: Retrieving all content types defined on the site...
Found 3 content type(s) on the site:
  - ID: 0x01 | Name: Item | Description: ...
  - ID: 0x0101 | Name: Document | Description: ...
  ...

Step 2: Examining column definitions for content type ID: <targetContentTypeId>
Found 5 column(s) in the content type:
  - ID: ... | Name: Title | Required: true
  ...

Step 3: 'Compliance Category' column is missing. Adding mandatory column to enforce governance standards...
Successfully added 'Compliance Category' column:
  - Column ID   : <generated-id>
  - Column Name : Compliance Category
  - Required    : true

=== Content Governance Audit Workflow Completed ===
```
