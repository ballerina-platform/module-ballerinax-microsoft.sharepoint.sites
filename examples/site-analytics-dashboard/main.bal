import ballerina/io;
import ballerinax/microsoft.sharepoint.sites;

configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string tenantId = ?;
configurable string siteId = ?;

public function main() returns error? {
    sites:ConnectionConfig config = {
        auth: <sites:OAuth2ClientCredentialsGrantConfig>{
            clientId: clientId,
            clientSecret: clientSecret,
            tokenUrl: string `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            scopes: ["https://graph.microsoft.com/.default"]
        }
    };

    sites:Client sharepointClient = check new (config);

    io:println("=== SharePoint Site Analytics Dashboard ===");
    io:println("");

    io:println("Step 1: Retrieving site details...");
    sites:MicrosoftGraphSite siteDetails = check sharepointClient->sitesSiteGetSite(
        siteId,
        queries = {
            dollarSelect: ["id", "displayName", "name", "description", "webUrl", "createdDateTime", "lastModifiedDateTime"]
        }
    );

    string siteDetailId = siteDetails?.id ?: "N/A";
    string siteDisplayName = siteDetails?.displayName ?: "N/A";
    string siteName = siteDetails?.name ?: "N/A";
    string siteWebUrl = siteDetails?.webUrl ?: "N/A";
    string siteDescription = siteDetails?.description ?: "N/A";
    string siteCreatedDateTime = siteDetails?.createdDateTime ?: "N/A";
    string siteLastModifiedDateTime = siteDetails?.lastModifiedDateTime ?: "N/A";

    io:println("Site ID: " + siteDetailId);
    io:println("Display Name: " + siteDisplayName);
    io:println("Site Name: " + siteName);
    io:println("Web URL: " + siteWebUrl);
    io:println("Description: " + siteDescription);
    io:println("Created Date: " + siteCreatedDateTime);
    io:println("Last Modified: " + siteLastModifiedDateTime);
    io:println("");

    io:println("Step 2: Fetching all-time analytics summary...");
    sites:MicrosoftGraphItemActivityStat allTimeStats = check sharepointClient->sitesAnalyticsGetAllTime(
        siteId,
        queries = {
            dollarSelect: ["id", "access", "create", "delete", "edit", "move", "startDateTime", "endDateTime", "isTrending", "incompleteData"]
        }
    );

    io:println("=== All-Time Analytics Summary ===");
    string allTimeStatId = allTimeStats?.id ?: "N/A";
    string allTimeStartDateTime = allTimeStats?.startDateTime ?: "N/A";
    string allTimeEndDateTime = allTimeStats?.endDateTime ?: "N/A";
    boolean allTimeIsTrending = allTimeStats?.isTrending ?: false;

    io:println("Stat ID: " + allTimeStatId);
    io:println("Start Date: " + allTimeStartDateTime);
    io:println("End Date: " + allTimeEndDateTime);
    io:println("Is Trending: " + allTimeIsTrending.toString());

    var allTimeAccess = allTimeStats?.access;
    if allTimeAccess is record {} {
        io:println("Access Activity: " + allTimeAccess.toString());
    }

    var allTimeCreate = allTimeStats?.create;
    if allTimeCreate is record {} {
        io:println("Create Activity: " + allTimeCreate.toString());
    }

    var allTimeEdit = allTimeStats?.edit;
    if allTimeEdit is record {} {
        io:println("Edit Activity: " + allTimeEdit.toString());
    }

    var allTimeDelete = allTimeStats?.delete;
    if allTimeDelete is record {} {
        io:println("Delete Activity: " + allTimeDelete.toString());
    }

    var allTimeMove = allTimeStats?.move;
    if allTimeMove is record {} {
        io:println("Move Activity: " + allTimeMove.toString());
    }

    var allTimeIncompleteData = allTimeStats?.incompleteData;
    if allTimeIncompleteData is record {} {
        io:println("Incomplete Data Info: " + allTimeIncompleteData.toString());
    }
    io:println("");

    io:println("Step 3: Fetching last seven days activity stats...");
    sites:MicrosoftGraphItemActivityStat lastSevenDaysStats = check sharepointClient->sitesAnalyticsGetLastSevenDays(
        siteId,
        queries = {
            dollarSelect: ["id", "access", "create", "delete", "edit", "move", "startDateTime", "endDateTime", "isTrending", "incompleteData"]
        }
    );

    io:println("=== Last Seven Days Analytics Summary ===");
    string recentStatId = lastSevenDaysStats?.id ?: "N/A";
    string recentStartDateTime = lastSevenDaysStats?.startDateTime ?: "N/A";
    string recentEndDateTime = lastSevenDaysStats?.endDateTime ?: "N/A";
    boolean recentIsTrending = lastSevenDaysStats?.isTrending ?: false;

    io:println("Stat ID: " + recentStatId);
    io:println("Period Start: " + recentStartDateTime);
    io:println("Period End: " + recentEndDateTime);
    io:println("Is Trending: " + recentIsTrending.toString());

    var recentAccess = lastSevenDaysStats?.access;
    if recentAccess is record {} {
        io:println("Recent Access Activity: " + recentAccess.toString());
    }

    var recentCreate = lastSevenDaysStats?.create;
    if recentCreate is record {} {
        io:println("Recent Create Activity: " + recentCreate.toString());
    }

    var recentEdit = lastSevenDaysStats?.edit;
    if recentEdit is record {} {
        io:println("Recent Edit Activity: " + recentEdit.toString());
    }

    var recentDelete = lastSevenDaysStats?.delete;
    if recentDelete is record {} {
        io:println("Recent Delete Activity: " + recentDelete.toString());
    }

    var recentMove = lastSevenDaysStats?.move;
    if recentMove is record {} {
        io:println("Recent Move Activity: " + recentMove.toString());
    }

    var recentIncompleteData = lastSevenDaysStats?.incompleteData;
    if recentIncompleteData is record {} {
        io:println("Incomplete Data Info: " + recentIncompleteData.toString());
    }
    io:println("");

    io:println("=== Dashboard Summary ===");
    string resolvedDisplayName = siteDetails?.displayName ?: siteDetails?.name ?: siteId;
    string summaryAllTimeStart = allTimeStats?.startDateTime ?: "N/A";
    string summaryAllTimeEnd = allTimeStats?.endDateTime ?: "N/A";
    string summaryRecentStart = lastSevenDaysStats?.startDateTime ?: "N/A";
    string summaryRecentEnd = lastSevenDaysStats?.endDateTime ?: "N/A";
    boolean summaryIsTrending = lastSevenDaysStats?.isTrending ?: false;

    io:println("Site '" + resolvedDisplayName + "' analytics retrieved successfully.");
    io:println("All-time stats period: " + summaryAllTimeStart + " to " + summaryAllTimeEnd);
    io:println("Recent stats period: " + summaryRecentStart + " to " + summaryRecentEnd);
    io:println("Trending status (last 7 days): " + summaryIsTrending.toString());
    io:println("Analytics dashboard generation complete.");
}