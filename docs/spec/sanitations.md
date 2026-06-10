_Author_: Thisaru Guruge \
_Created_: 2026-06-10 \
_Updated_: 2026-06-10 \
_Edition_: Swan Lake

# Sanitation for OpenAPI specification

This document records the sanitation done on top of the official OpenAPI specification from Microsoft SharePoint Sites Connector.
The OpenAPI specification is obtained from the [Microsoft Graph REST API v1.0 OpenAPI description](https://github.com/microsoftgraph/msgraph-metadata/tree/master/openapi/v1.0).
These changes are done in order to improve the overall usability, and as workarounds for some known language limitations.

1. **Moved `/sites` to the base server URL**: All 97 API paths shared the common `/sites` prefix (e.g., `/sites`, `/sites/{siteId}/analytics`). The prefix has been moved to the server URL (`https://graph.microsoft.com/v1.0/sites`), and the paths have been updated accordingly (e.g., `/`, `/{siteId}/analytics`). This eliminates the redundant `sitesSite` prefix that the Ballerina OpenAPI tool would otherwise generate for every method and type name.

## OpenAPI cli command

The following command was used to generate the Ballerina client from the OpenAPI specification. The command should be executed from the repository root directory.

```bash
bal openapi -i docs/spec/aligned_ballerina_openapi.yaml --mode client --license license.txt -o ballerina
```

Note: The license year is hardcoded to 2026.
