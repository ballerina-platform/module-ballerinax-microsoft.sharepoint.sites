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

public function main() returns error? {
    sites:Client sharepointClient = check new ({
        auth: {
            clientId: clientId,
            clientSecret: clientSecret,
            tokenUrl: "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token",
            scopes: ["https://graph.microsoft.com/.default"]
        }
    });

    io:println("=== SharePoint Site Drive Discovery and Multi-Drive Inventory Workflow ===");
    io:println("");

    io:println("Step 1: Discovering all SharePoint sites in the organization...");
    sites:MicrosoftGraphSiteCollectionResponse siteListResponse = check sharepointClient->siteListSite(
        queries = {
            dollarSearch: "*",
            dollarTop: 10,
            dollarSelect: ["id", "displayName", "webUrl", "name", "description"]
        }
    );

    sites:MicrosoftGraphSite[] allSites = siteListResponse.value ?: [];
    io:println("Total sites discovered: " + allSites.length().toString());
    io:println("");

    if allSites.length() == 0 {
        io:println("No sites found. Exiting workflow.");
        return;
    }

    io:println("--- Discovered Sites ---");
    foreach sites:MicrosoftGraphSite site in allSites {
        string siteId = site?.id ?: "N/A";
        string displayNameVal = site?.displayName ?: "Unnamed Site";
        string siteWebUrl = site?.webUrl ?: "N/A";
        io:println("  Site Name : " + displayNameVal);
        io:println("  Site ID   : " + siteId);
        io:println("  Web URL   : " + siteWebUrl);
        io:println("  ---");
    }
    io:println("");

    io:println("=== Step 2 & 3: Retrieving Drive Inventory for Each Site ===");
    io:println("");

    foreach sites:MicrosoftGraphSite site in allSites {
        string? siteIdOptional = site?.id;
        if siteIdOptional is () {
            io:println("Skipping site with no ID.");
            continue;
        }

        string siteIdVal = siteIdOptional;
        string displayNameVal = site?.displayName ?: "Unnamed Site";
        io:println("Processing site: " + displayNameVal + " (ID: " + siteIdVal + ")");
        io:println("--------------------------------------------------");

        io:println("  Step 2: Fetching primary (default) drive...");
        sites:MicrosoftGraphDrive|error primaryDriveResult = sharepointClient->getDrive(
            siteIdVal,
            queries = {
                dollarSelect: ["id", "name", "driveType", "webUrl", "description", "quota", "owner", "createdDateTime", "lastModifiedDateTime"]
            }
        );

        if primaryDriveResult is error {
            io:println("  Warning: Could not retrieve primary drive for site '" + displayNameVal + "': " + primaryDriveResult.message());
        } else {
            sites:MicrosoftGraphDrive primaryDrive = primaryDriveResult;
            string driveId = primaryDrive?.id ?: "N/A";
            string driveName = primaryDrive?.name ?: "Default Drive";
            string driveType = primaryDrive?.driveType ?: "N/A";
            string driveWebUrl = primaryDrive?.webUrl ?: "N/A";

            io:println("  Primary Drive Details:");
            io:println("    Drive ID    : " + driveId);
            io:println("    Drive Name  : " + driveName);
            io:println("    Drive Type  : " + driveType);
            io:println("    Web URL     : " + driveWebUrl);

            (sites:MicrosoftGraphQuota|record {})? quotaVal = primaryDrive?.quota;
            if quotaVal != () {
                io:println("    Quota Info  : " + quotaVal.toString());
            } else {
                io:println("    Quota Info  : Not available");
            }

            (sites:MicrosoftGraphIdentitySet|record {})? ownerVal = primaryDrive?.owner;
            if ownerVal != () {
                io:println("    Owner Info  : " + ownerVal.toString());
            } else {
                io:println("    Owner Info  : Not available");
            }
        }

        io:println("");

        io:println("  Step 3: Enumerating all drives (document libraries) on this site...");
        sites:MicrosoftGraphDriveCollectionResponse|error allDrivesResult = sharepointClient->listDrives(
            siteIdVal,
            queries = {
                dollarSelect: ["id", "name", "driveType", "webUrl", "description", "quota", "createdDateTime", "lastModifiedDateTime"],
                dollarTop: 20
            }
        );

        if allDrivesResult is error {
            io:println("  Warning: Could not list drives for site '" + displayNameVal + "': " + allDrivesResult.message());
        } else {
            sites:MicrosoftGraphDriveCollectionResponse drivesResponse = allDrivesResult;
            sites:MicrosoftGraphDrive[] allDrives = drivesResponse.value ?: [];
            io:println("  Total drives found on site: " + allDrives.length().toString());
            io:println("");

            if allDrives.length() == 0 {
                io:println("  No additional drives found for this site.");
            } else {
                io:println("  Drive Inventory:");
                int driveIndex = 1;
                foreach sites:MicrosoftGraphDrive drive in allDrives {
                    string driveId = drive?.id ?: "N/A";
                    string driveName = drive?.name ?: "Unnamed Drive";
                    string driveType = drive?.driveType ?: "N/A";
                    string driveWebUrl = drive?.webUrl ?: "N/A";
                    string driveDescription = drive?.description ?: "No description";
                    string createdDateTimeVal = drive?.createdDateTime ?: "N/A";
                    string lastModifiedDateTimeVal = drive?.lastModifiedDateTime ?: "N/A";

                    io:println("    [" + driveIndex.toString() + "] Drive Name        : " + driveName);
                    io:println("        Drive ID          : " + driveId);
                    io:println("        Drive Type        : " + driveType);
                    io:println("        Web URL           : " + driveWebUrl);
                    io:println("        Description       : " + driveDescription);
                    io:println("        Created           : " + createdDateTimeVal);
                    io:println("        Last Modified     : " + lastModifiedDateTimeVal);

                    (sites:MicrosoftGraphQuota|record {})? quotaVal = drive?.quota;
                    if quotaVal != () {
                        io:println("        Quota Details     : " + quotaVal.toString());
                    } else {
                        io:println("        Quota Details     : Not available");
                    }

                    io:println("        ----");
                    driveIndex = driveIndex + 1;
                }
            }
        }

        io:println("");
        io:println("==================================================");
        io:println("");
    }

    io:println("=== SharePoint Drive Discovery and Inventory Workflow Completed Successfully ===");
    io:println("Summary: Processed " + allSites.length().toString() + " site(s) and collected drive inventory for governance reporting and storage capacity planning.");
}
