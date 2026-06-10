# Examples

The `microsoft.sharepoint.sites` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples), covering use cases like site analytics dashboard, site permission audit cleanup, and content governance audit workflow.

1. [Site analytics dashboard](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-analytics-dashboard) - Build a dashboard that retrieves and displays analytics data for SharePoint sites.

2. [Site permission audit cleanup](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/site-permission-audit-cleanup) - Audit and clean up permissions across SharePoint sites to ensure proper access control.

3. [Content governance audit workflow](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/content-governance-audit-workflow) - Automate a content governance workflow that audits SharePoint site content for compliance.

4. [Sharepoint drive inventory](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.sites/tree/main/examples/sharepoint-drive-inventory) - Generate a comprehensive inventory of all drives and their contents across SharePoint sites.

## Prerequisites

1. Generate Microsoft SharePoint credentials to authenticate the connector as described in the [Setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.sites/latest#setup-guide).

2. For each example, create a `Config.toml` file with the credentials for the example. Here's an example of how your `Config.toml` file should look:

    ```toml
    clientId = "<Client ID>"
    clientSecret = "<Client Secret>"
    tenantId = "<Tenant ID>"
    siteId = "<Site ID>"
    ```

    > **Note:** Not all examples use `siteId`. Refer to each example's `README.md` for the specific configuration required.

## Running an Example

Execute the following commands to build an example from the source:

* To build an example:

    ```bash
    bal build
    ```

* To run an example:

    ```bash
    bal run
    ```
