# Running Tests

## Prerequisites

Before running the tests, ensure you have configured the necessary credentials and environment for the **microsoft.sharepoint.sites** connector. Refer to the connector setup guide for detailed instructions:

[microsoft.sharepoint.sites Connector README](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/blob/main/ballerina/README.md)

---

## Running Tests

The tests support two modes: **mock server** (default, no credentials required) and **live server** (runs against the real Microsoft SharePoint API).

### 1. Mock Server Tests (Default)

No credentials are needed. Simply run:

```bash
bal test
```

The `Config.toml` in the `tests` directory already has `isTestOnLiveServer = false` set as the default.

---

### 2. Live Server Tests

To run tests against the live Microsoft SharePoint API, update the `Config.toml` file in the `tests` directory with your credentials:

```toml
[ballerinax.microsoft.sharepoint.sites]
isTestOnLiveServer = true
clientId = "<your-client-id>"
tenantId = "<your-tenant-id>"
clientSecret = "<your-client-secret>"
siteId = "<your-sharepoint-site-id>"
```

Then run:

```bash
bal test
```

> **Note:** The `siteId` is the full SharePoint site identifier in the format `hostname,siteCollectionId,webId` (e.g., `contoso.sharepoint.com,abc-123,def-456`). You can retrieve it by calling `GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-name}` in [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and copying the `id` field from the response.
