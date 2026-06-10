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

To run tests against the live Microsoft SharePoint API, you need valid credentials.

#### Step 1: Set Environment Variables

**Linux/macOS:**

```bash
export SHAREPOINT_ACCESS_TOKEN="<your-access-token>"
export SHAREPOINT_SITE_ID="<your-sharepoint-site-id>"
```

**Windows (Command Prompt):**

```cmd
set SHAREPOINT_ACCESS_TOKEN=<your-access-token>
set SHAREPOINT_SITE_ID=<your-sharepoint-site-id>
```

**Windows (PowerShell):**

```powershell
$env:SHAREPOINT_ACCESS_TOKEN = "<your-access-token>"
$env:SHAREPOINT_SITE_ID = "<your-sharepoint-site-id>"
```

#### Step 2: Enable Live Server Mode

Update the `Config.toml` file in the `tests` directory:

```toml
[ballerinax.microsoft.sharepoint.sites]
isTestOnLiveServer = true
```

#### Step 3: Run the Tests

```bash
bal test
```

> **Note:** `SHAREPOINT_ACCESS_TOKEN` is a short-lived Bearer token obtained from Microsoft identity platform. You can get one via the OAuth 2.0 client credentials flow using your `clientId`, `clientSecret`, and `tenantId`. `SHAREPOINT_SITE_ID` is the full SharePoint site identifier (e.g., `contoso.sharepoint.com,<site-guid>,<web-guid>`).
