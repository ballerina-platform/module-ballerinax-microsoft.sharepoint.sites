// Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
//
// WSO2 LLC. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

import ballerina/io;
import ballerinax/microsoft.sharepoint.sites;

configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string tenantId = ?;
configurable string siteId = ?;

public function main() returns error? {
    sites:ConnectionConfig config = {
        auth: <sites:OAuth2ClientCredentialsGrantConfig>{
            clientId,
            clientSecret,
            tokenUrl: string `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            scopes: ["https://graph.microsoft.com/.default"]
        }
    };

    sites:Client sharepointClient = check new (config);

    io:println("=== SharePoint Site Analytics Dashboard ===");
    io:println("");

    io:println("Step 1: Retrieving site details...");
    sites:MicrosoftGraphSite siteDetails = check sharepointClient->siteGetSite(
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
    sites:MicrosoftGraphItemActivityStat allTimeStats = check sharepointClient->analyticsGetAllTime(
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

    (sites:MicrosoftGraphItemActionStat|record {})? allTimeAccess = allTimeStats?.access;
    if allTimeAccess != () {
        io:println("Access Activity: " + allTimeAccess.toString());
    }

    (sites:MicrosoftGraphItemActionStat|record {})? allTimeCreate = allTimeStats?.create;
    if allTimeCreate != () {
        io:println("Create Activity: " + allTimeCreate.toString());
    }

    (sites:MicrosoftGraphItemActionStat|record {})? allTimeEdit = allTimeStats?.edit;
    if allTimeEdit != () {
        io:println("Edit Activity: " + allTimeEdit.toString());
    }

    (sites:MicrosoftGraphItemActionStat|record {})? allTimeDelete = allTimeStats?.delete;
    if allTimeDelete != () {
        io:println("Delete Activity: " + allTimeDelete.toString());
    }

    (sites:MicrosoftGraphItemActionStat|record {})? allTimeMove = allTimeStats?.move;
    if allTimeMove != () {
        io:println("Move Activity: " + allTimeMove.toString());
    }

    (sites:MicrosoftGraphIncompleteData|record {})? allTimeIncompleteData = allTimeStats?.incompleteData;
    if allTimeIncompleteData != () {
        io:println("Incomplete Data Info: " + allTimeIncompleteData.toString());
    }
    io:println("");

    io:println("Step 3: Fetching last seven days activity stats...");
    sites:MicrosoftGraphItemActivityStat lastSevenDaysStats = check sharepointClient->analyticsGetLastSevenDays(
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

    (sites:MicrosoftGraphItemActionStat|record {})? recentAccess = lastSevenDaysStats?.access;
    if recentAccess != () {
        io:println("Recent Access Activity: " + recentAccess.toString());
    }

    (sites:MicrosoftGraphItemActionStat|record {})? recentCreate = lastSevenDaysStats?.create;
    if recentCreate != () {
        io:println("Recent Create Activity: " + recentCreate.toString());
    }

    (sites:MicrosoftGraphItemActionStat|record {})? recentEdit = lastSevenDaysStats?.edit;
    if recentEdit != () {
        io:println("Recent Edit Activity: " + recentEdit.toString());
    }

    (sites:MicrosoftGraphItemActionStat|record {})? recentDelete = lastSevenDaysStats?.delete;
    if recentDelete != () {
        io:println("Recent Delete Activity: " + recentDelete.toString());
    }

    (sites:MicrosoftGraphItemActionStat|record {})? recentMove = lastSevenDaysStats?.move;
    if recentMove != () {
        io:println("Recent Move Activity: " + recentMove.toString());
    }

    (sites:MicrosoftGraphIncompleteData|record {})? recentIncompleteData = lastSevenDaysStats?.incompleteData;
    if recentIncompleteData != () {
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
