# Running Tests

## Prerequisites

Before running the tests, ensure you have configured the necessary credentials and environment for the `microsoft.sharepoint.sites` connector. Refer to the connector setup guide for detailed instructions:

👉 [microsoft.sharepoint.sites Connector README](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/blob/main/ballerina/README.md)

---

## Running Tests

The tests for this connector require valid Microsoft SharePoint credentials and run against the live SharePoint API.

### 1. Configure Credentials

Create a `Config.toml` file in the `tests` directory with the following content:

```toml
accessToken = "<your-sharepoint-access-token>"
siteId = "<your-sharepoint-site-id>"
```

> **Note:** Replace each placeholder value with your actual credentials before running the tests.

Alternatively, you can provide the credentials using environment variables.

**Linux/macOS:**

```bash
export SHAREPOINT_ACCESS_TOKEN="<your-sharepoint-access-token>"
export SHAREPOINT_SITE_ID="<your-sharepoint-site-id>"
```

**Windows (Command Prompt):**

```cmd
set SHAREPOINT_ACCESS_TOKEN=<your-sharepoint-access-token>
set SHAREPOINT_SITE_ID=<your-sharepoint-site-id>
```

**Windows (PowerShell):**

```powershell
$env:SHAREPOINT_ACCESS_TOKEN = "<your-sharepoint-access-token>"
$env:SHAREPOINT_SITE_ID = "<your-sharepoint-site-id>"
```

---

### 2. Run the Tests

Once credentials are configured, execute the following command from the `tests` directory (or the package root):

```bash
bal test
```

---

### Skipping Live Tests

If you need to skip the live tests (for example, in a CI environment without credentials), set the `LIVE_TEST_DISABLED` environment variable:

**Linux/macOS:**

```bash
export LIVE_TEST_DISABLED=true
bal test
```

**Windows (Command Prompt):**

```cmd
set LIVE_TEST_DISABLED=true
bal test
```

**Windows (PowerShell):**

```powershell
$env:LIVE_TEST_DISABLED = "true"
bal test
```

> When `LIVE_TEST_DISABLED` is set to `true`, all live API tests will be skipped gracefully.
