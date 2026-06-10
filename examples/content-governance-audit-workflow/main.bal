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

configurable string tenantId = ?;
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string siteId = ?;
configurable string targetContentTypeId = ?;

public function main() returns error? {
    sites:Client sharepointClient = check new ({
        auth: {
            clientId: clientId,
            clientSecret: clientSecret,
            tokenUrl: "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token",
            scopes: ["https://graph.microsoft.com/.default"]
        }
    });

    io:println("=== Content Governance Audit Workflow ===");
    io:println("");

    io:println("Step 1: Retrieving all content types defined on the site...");
    sites:MicrosoftGraphContentTypeCollectionResponse contentTypesResponse =
        check sharepointClient->listContentTypes(
            siteId,
            queries = {
                dollarSelect: ["id", "name", "description"]
            }
        );

    sites:MicrosoftGraphContentType[] contentTypes = contentTypesResponse.value ?: [];
    io:println("Found " + contentTypes.length().toString() + " content type(s) on the site:");
    foreach sites:MicrosoftGraphContentType ct in contentTypes {
        string ctId = ct?.id ?: "N/A";
        string ctName = ct?.name ?: "Unnamed";
        string ctDescription = ct?.description ?: "No description";
        io:println("  - ID: " + ctId + " | Name: " + ctName + " | Description: " + ctDescription);
    }
    io:println("");

    io:println("Step 2: Examining column definitions for content type ID: " + targetContentTypeId);
    sites:MicrosoftGraphColumnDefinitionCollectionResponse columnsResponse =
        check sharepointClient->contentTypesListColumns(
            siteId,
            targetContentTypeId,
            queries = {
                dollarSelect: ["id", "name", "displayName", "description", "required", "enforceUniqueValues", "hidden", "indexed"]
            }
        );

    sites:MicrosoftGraphColumnDefinition[] columns = columnsResponse.value ?: [];
    boolean complianceCategoryExists = false;

    io:println("Found " + columns.length().toString() + " column(s) in the content type:");
    foreach sites:MicrosoftGraphColumnDefinition col in columns {
        string colId = col?.id ?: "N/A";
        string colName = col?.name ?: "Unnamed";
        boolean isRequired = col?.required ?: false;
        io:println("  - ID: " + colId + " | Name: " + colName + " | Required: " + isRequired.toString());

        if colName == "Compliance Category" {
            complianceCategoryExists = true;
        }
    }
    io:println("");

    if complianceCategoryExists {
        io:println("Step 3: 'Compliance Category' column already exists. No action needed.");
    } else {
        io:println("Step 3: 'Compliance Category' column is missing. Adding mandatory column to enforce governance standards...");

        sites:MicrosoftGraphChoiceColumn choiceColumn = {
            allowTextEntry: false,
            choices: [
                "Regulatory",
                "Internal Policy",
                "Legal Hold",
                "Public",
                "Confidential",
                "Highly Confidential"
            ],
            displayAs: "dropDownMenu"
        };

        sites:MicrosoftGraphColumnDefinition newColumn = {
            name: "Compliance Category",
            displayName: "Compliance Category",
            description: "Mandatory field to classify document compliance category as required by organizational governance standards.",
            required: true,
            enforceUniqueValues: false,
            hidden: false,
            indexed: true,
            choice: choiceColumn
        };

        sites:MicrosoftGraphColumnDefinition createdColumn =
            check sharepointClient->contentTypesCreateColumns(siteId, targetContentTypeId, newColumn);

        string createdId = createdColumn?.id ?: "N/A";
        string createdName = createdColumn?.name ?: "Unnamed";
        boolean createdRequired = createdColumn?.required ?: false;

        io:println("Successfully added 'Compliance Category' column:");
        io:println("  - Column ID   : " + createdId);
        io:println("  - Column Name : " + createdName);
        io:println("  - Required    : " + createdRequired.toString());
    }

    io:println("");
    io:println("=== Content Governance Audit Workflow Completed ===");
}
