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
accessKeyId = "<your-access-key-id>"
secretAccessKey = "<your-secret-access-key>"
region = "<your-region>"
```

> **Note:** Replace each placeholder value with your actual credentials before running the tests.

Alternatively, you can provide the credentials using environment variables.

**Linux/macOS:**

```bash
export accessKeyId="<your-access-key-id>"
export secretAccessKey="<your-secret-access-key>"
export region="<your-region>"
```

**Windows (Command Prompt):**

```cmd
set accessKeyId=<your-access-key-id>
set secretAccessKey=<your-secret-access-key>
set region=<your-region>
```

**Windows (PowerShell):**

```powershell
$env:accessKeyId = "<your-access-key-id>"
$env:secretAccessKey = "<your-secret-access-key>"
$env:region = "<your-region>"
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
