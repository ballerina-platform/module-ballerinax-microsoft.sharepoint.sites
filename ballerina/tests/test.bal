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

import ballerina/os;
import ballerina/test;

configurable boolean isTestOnLiveServer = false;

final boolean isTestOnMockServer = !isTestOnLiveServer;

Client sharepointClient = test:mock(Client);

// Test resource IDs updated during test execution
string testSiteId = MOCK_SITE_ID;
string testColumnId = MOCK_COLUMN_ID;
string testContentTypeId = MOCK_CONTENT_TYPE_ID;
string testPermissionId = MOCK_PERMISSION_ID;
string testOperationId = MOCK_OPERATION_ID;
string testDriveId = MOCK_DRIVE_ID;

@test:BeforeSuite
function initClient() returns error? {
    if isTestOnLiveServer {
        string accessToken = os:getEnv("SHAREPOINT_ACCESS_TOKEN");
        if accessToken.trim().length() == 0 {
            return error("SHAREPOINT_ACCESS_TOKEN environment variable is required for live server tests");
        }
        string liveSiteId = os:getEnv("SHAREPOINT_SITE_ID");
        if liveSiteId.trim().length() == 0 {
            return error("SHAREPOINT_SITE_ID environment variable is required for live server tests");
        }
        sharepointClient = check new ({
            auth: {token: accessToken}
        });
        testSiteId = liveSiteId;
    } else {
        sharepointClient = check new ({
            auth: {token: "mock-access-token"}
        }, "http://localhost:9090");
    }
}

// ========================
// Site Tests
// ========================

@test:Config {}
function testListSites() returns error? {
    MicrosoftGraphSiteCollectionResponse|error result = sharepointClient->siteListSite();
    if result is error {
        test:assertFail(result.message());
    }
    MicrosoftGraphSite[]? sites = result.value;
    test:assertNotEquals(sites, (), "Sites collection should not be null");
    if sites is MicrosoftGraphSite[] {
        test:assertTrue(sites.length() > 0, "Should return at least one site");
    }
}

@test:Config {
    dependsOn: [testListSites]
}
function testGetSite() returns error? {
    MicrosoftGraphSite|error result = sharepointClient->siteGetSite(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertEquals(result?.id, testSiteId, "Site ID should match");
    test:assertNotEquals(result?.displayName, (), "Site should have a display name");
}

@test:Config {
    dependsOn: [testGetSite]
}
function testUpdateSite() returns error? {
    MicrosoftGraphSite payload = {displayName: "Updated Site Name"};
    error? result = sharepointClient->siteUpdateSite(testSiteId, payload);
    test:assertEquals(result, (), "Update should succeed without error");
}

// ========================
// Analytics Tests
// ========================

@test:Config {
    dependsOn: [testGetSite]
}
function testGetAnalytics() returns error? {
    MicrosoftGraphItemAnalytics|error result = sharepointClient->getAnalytics(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Analytics should not be null");
}

@test:Config {
    dependsOn: [testGetAnalytics]
}
function testGetAllTimeAnalytics() returns error? {
    MicrosoftGraphItemActivityStat|error result = sharepointClient->analyticsGetAllTime(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "All-time analytics should not be null");
}

@test:Config {
    dependsOn: [testGetAnalytics]
}
function testGetLastSevenDaysAnalytics() returns error? {
    MicrosoftGraphItemActivityStat|error result = sharepointClient->analyticsGetLastSevenDays(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Last seven days analytics should not be null");
}

@test:Config {
    dependsOn: [testGetAnalytics]
}
function testListItemActivityStats() returns error? {
    MicrosoftGraphItemActivityStatCollectionResponse|error result = sharepointClient->analyticsListItemActivityStats(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Activity stats collection should not be null");
}

// ========================
// Column Tests
// ========================

@test:Config {
    dependsOn: [testGetSite]
}
function testListColumns() returns error? {
    MicrosoftGraphColumnDefinitionCollectionResponse|error result = sharepointClient->listColumns(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    MicrosoftGraphColumnDefinition[]? columns = result.value;
    test:assertNotEquals(columns, (), "Columns collection should not be null");
}

@test:Config {
    dependsOn: [testListColumns]
}
function testCreateColumn() returns error? {
    MicrosoftGraphColumnDefinition payload = {
        displayName: "Test Column",
        description: "A test column"
    };
    MicrosoftGraphColumnDefinition|error result = sharepointClient->createColumns(testSiteId, payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result?.id, (), "Created column should have an ID");
    test:assertEquals(result?.displayName, "Test Column", "Column display name should match");
    string? colId = result?.id;
    if colId is string {
        testColumnId = colId;
    }
}

@test:Config {
    dependsOn: [testCreateColumn]
}
function testGetColumn() returns error? {
    MicrosoftGraphColumnDefinition|error result = sharepointClient->getColumns(testSiteId, testColumnId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertEquals(result.id, testColumnId, "Column ID should match");
}

@test:Config {
    dependsOn: [testGetColumn]
}
function testGetSourceColumn() returns error? {
    MicrosoftGraphColumnDefinition|error result = sharepointClient->columnsGetSourceColumn(testSiteId, testColumnId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Source column should not be null");
}

@test:Config {
    dependsOn: [testGetSourceColumn]
}
function testUpdateColumn() returns error? {
    MicrosoftGraphColumnDefinition payload = {description: "Updated column description"};
    error? result = sharepointClient->updateColumns(testSiteId, testColumnId, payload);
    test:assertEquals(result, (), "Column update should succeed without error");
}

@test:Config {
    dependsOn: [testUpdateColumn]
}
function testDeleteColumn() returns error? {
    error? result = sharepointClient->deleteColumns(testSiteId, testColumnId);
    test:assertEquals(result, (), "Column delete should succeed without error");
}

// ========================
// External Columns Tests
// ========================

@test:Config {
    dependsOn: [testGetSite]
}
function testListExternalColumns() returns error? {
    MicrosoftGraphColumnDefinitionCollectionResponse|error result = sharepointClient->listExternalColumns(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "External columns response should not be null");
}

@test:Config {
    dependsOn: [testListExternalColumns]
}
function testGetExternalColumn() returns error? {
    MicrosoftGraphColumnDefinition|error result = sharepointClient->getExternalColumns(testSiteId, MOCK_COLUMN_ID);
    if result is error && isTestOnLiveServer {
        // External columns may not exist in live environment — acceptable skip
        return;
    }
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result.id, (), "External column should have an ID");
}

// ========================
// Content Type Tests
// ========================

@test:Config {
    dependsOn: [testGetSite]
}
function testListContentTypes() returns error? {
    MicrosoftGraphContentTypeCollectionResponse|error result = sharepointClient->listContentTypes(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    MicrosoftGraphContentType[]? contentTypes = result.value;
    test:assertNotEquals(contentTypes, (), "Content types collection should not be null");
}

@test:Config {
    dependsOn: [testListContentTypes]
}
function testCreateContentType() returns error? {
    MicrosoftGraphContentType payload = {
        name: "Test Content Type",
        description: "A test content type"
    };
    MicrosoftGraphContentType|error result = sharepointClient->createContentTypes(testSiteId, payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result?.id, (), "Created content type should have an ID");
    test:assertEquals(result?.name, "Test Content Type", "Content type name should match");
    string? ctId = result?.id;
    if ctId is string {
        testContentTypeId = ctId;
    }
}

@test:Config {
    dependsOn: [testCreateContentType]
}
function testGetContentType() returns error? {
    MicrosoftGraphContentType|error result = sharepointClient->getContentTypes(testSiteId, testContentTypeId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertEquals(result.id, testContentTypeId, "Content type ID should match");
}

@test:Config {
    dependsOn: [testGetContentType]
}
function testListContentTypeColumns() returns error? {
    MicrosoftGraphColumnDefinitionCollectionResponse|error result = sharepointClient->contentTypesListColumns(testSiteId, testContentTypeId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Content type columns response should not be null");
}

@test:Config {
    dependsOn: [testGetContentType]
}
function testListContentTypeColumnLinks() returns error? {
    MicrosoftGraphColumnLinkCollectionResponse|error result = sharepointClient->contentTypesListColumnLinks(testSiteId, testContentTypeId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Column links response should not be null");
}

@test:Config {
    dependsOn: [testGetContentType]
}
function testListContentTypeBaseTypes() returns error? {
    MicrosoftGraphContentTypeCollectionResponse|error result = sharepointClient->contentTypesListBaseTypes(testSiteId, testContentTypeId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Base types response should not be null");
}

@test:Config {
    dependsOn: [testGetContentType]
}
function testGetContentTypeBase() returns error? {
    MicrosoftGraphContentType|error result = sharepointClient->contentTypesGetBase(testSiteId, testContentTypeId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Base content type should not be null");
}

@test:Config {
    dependsOn: [testGetContentType]
}
function testListContentTypeColumnPositions() returns error? {
    MicrosoftGraphColumnDefinitionCollectionResponse|error result = sharepointClient->contentTypesListColumnPositions(testSiteId, testContentTypeId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Column positions response should not be null");
}

@test:Config {
    dependsOn: [testListContentTypeColumns]
}
function testCreateContentTypeColumn() returns error? {
    MicrosoftGraphColumnDefinition payload = {
        displayName: "CT Test Column"
    };
    MicrosoftGraphColumnDefinition|error result = sharepointClient->contentTypesCreateColumns(testSiteId, testContentTypeId, payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result?.id, (), "Created column should have an ID");
}

@test:Config {
    dependsOn: [testListContentTypeColumnLinks]
}
function testCreateContentTypeColumnLink() returns error? {
    MicrosoftGraphColumnLink payload = {name: "Test Column"};
    MicrosoftGraphColumnLink|error result = sharepointClient->contentTypesCreateColumnLinks(testSiteId, testContentTypeId, payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result.id, (), "Created column link should have an ID");
}

@test:Config {
    dependsOn: [testGetContentType]
}
function testIsContentTypePublished() returns error? {
    BooleanValueResponse|error result = sharepointClient->siteContentTypesContentTypeIsPublished(testSiteId, testContentTypeId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "isPublished response should not be null");
}

@test:Config {
    dependsOn: [testIsContentTypePublished],
    enable: isTestOnMockServer
}
function testPublishContentType() returns error? {
    error? result = sharepointClient->siteContentTypesContentTypePublish(testSiteId, testContentTypeId);
    test:assertEquals(result, (), "Publish should succeed without error");
}

@test:Config {
    dependsOn: [testPublishContentType],
    enable: isTestOnMockServer
}
function testUnpublishContentType() returns error? {
    error? result = sharepointClient->siteContentTypesContentTypeUnpublish(testSiteId, testContentTypeId);
    test:assertEquals(result, (), "Unpublish should succeed without error");
}

@test:Config {
    dependsOn: [testListContentTypes],
    enable: isTestOnMockServer
}
function testAddCopyContentType() returns error? {
    ContentTypesMicrosoftGraphAddCopyBody payload = {
        contentType: "https://contoso.sharepoint.com/sites/source/_api/web/contentTypes/0x01"
    };
    ContentTypeOrNullResponse|error result = sharepointClient->siteContentTypesAddCopy(testSiteId, payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Added content type should not be null");
}

@test:Config {
    dependsOn: [testListContentTypes]
}
function testGetCompatibleHubContentTypes() returns error? {
    CollectionOfContentType|error result = sharepointClient->siteContentTypesGetCompatibleHubContentTypes(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Compatible hub content types should not be null");
}

@test:Config {
    dependsOn: [testGetContentType]
}
function testUpdateContentType() returns error? {
    MicrosoftGraphContentType payload = {description: "Updated description"};
    error? result = sharepointClient->updateContentTypes(testSiteId, testContentTypeId, payload);
    test:assertEquals(result, (), "Content type update should succeed without error");
}

@test:Config {
    dependsOn: [
        testUpdateContentType,
        testListContentTypeColumns,
        testListContentTypeColumnLinks,
        testListContentTypeBaseTypes,
        testGetContentTypeBase,
        testListContentTypeColumnPositions,
        testCreateContentTypeColumn,
        testCreateContentTypeColumnLink,
        testIsContentTypePublished
    ]
}
function testDeleteContentType() returns error? {
    error? result = sharepointClient->deleteContentTypes(testSiteId, testContentTypeId);
    test:assertEquals(result, (), "Content type delete should succeed without error");
}

// ========================
// Drive Tests
// ========================

@test:Config {
    dependsOn: [testGetSite]
}
function testGetDrive() returns error? {
    MicrosoftGraphDrive|error result = sharepointClient->getDrive(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result?.id, (), "Drive should have an ID");
    test:assertNotEquals(result?.driveType, (), "Drive should have a type");
}

@test:Config {
    dependsOn: [testGetSite]
}
function testListDrives() returns error? {
    MicrosoftGraphDriveCollectionResponse|error result = sharepointClient->listDrives(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    MicrosoftGraphDrive[]? drives = result.value;
    test:assertNotEquals(drives, (), "Drives collection should not be null");
    if drives is MicrosoftGraphDrive[] && drives.length() > 0 {
        string? driveId = drives[0].id;
        if driveId is string {
            testDriveId = driveId;
        }
    }
}

@test:Config {
    dependsOn: [testListDrives]
}
function testGetDriveById() returns error? {
    MicrosoftGraphDrive|error result = sharepointClient->getDrives(testSiteId, testDriveId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result.id, (), "Drive should have an ID");
}

// ========================
// Items Tests
// ========================

@test:Config {
    dependsOn: [testGetSite]
}
function testListItems() returns error? {
    MicrosoftGraphBaseItemCollectionResponse|error result = sharepointClient->listItems(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Items response should not be null");
}

@test:Config {
    dependsOn: [testListItems],
    enable: isTestOnMockServer
}
function testGetItem() returns error? {
    MicrosoftGraphBaseItem|error result = sharepointClient->getItems(testSiteId, "mock-item-id");
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result.id, (), "Item should have an ID");
}

// ========================
// Permission Tests
// ========================

@test:Config {
    dependsOn: [testGetSite]
}
function testListPermissions() returns error? {
    MicrosoftGraphPermissionCollectionResponse|error result = sharepointClient->listPermissions(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    MicrosoftGraphPermission[]? permissions = result.value;
    test:assertNotEquals(permissions, (), "Permissions collection should not be null");
}

@test:Config {
    dependsOn: [testListPermissions]
}
function testCreatePermission() returns error? {
    MicrosoftGraphPermission payload = {
        roles: ["read"]
    };
    MicrosoftGraphPermission|error result = sharepointClient->createPermissions(testSiteId, payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result?.id, (), "Created permission should have an ID");
    string? permId = result?.id;
    if permId is string {
        testPermissionId = permId;
    }
}

@test:Config {
    dependsOn: [testCreatePermission]
}
function testGetPermission() returns error? {
    MicrosoftGraphPermission|error result = sharepointClient->getPermissions(testSiteId, testPermissionId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertEquals(result.id, testPermissionId, "Permission ID should match");
}

@test:Config {
    dependsOn: [testGetPermission]
}
function testUpdatePermission() returns error? {
    MicrosoftGraphPermission payload = {roles: ["write"]};
    error? result = sharepointClient->updatePermissions(testSiteId, testPermissionId, payload);
    test:assertEquals(result, (), "Permission update should succeed without error");
}

@test:Config {
    dependsOn: [testUpdatePermission]
}
function testDeletePermission() returns error? {
    error? result = sharepointClient->deletePermissions(testSiteId, testPermissionId);
    test:assertEquals(result, (), "Permission delete should succeed without error");
}

@test:Config {
    dependsOn: [testGetPermission],
    enable: isTestOnMockServer
}
function testGrantPermission() returns error? {
    PermissionIdMicrosoftGraphGrantBody payload = {
        recipients: [{email: "test@contoso.com"}],
        roles: ["read"]
    };
    CollectionOfPermission|error result = sharepointClient->sitePermissionsPermissionGrant(testSiteId, testPermissionId, payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Grant permission response should not be null");
}

// ========================
// Operations Tests
// ========================

@test:Config {
    dependsOn: [testGetSite]
}
function testListOperations() returns error? {
    MicrosoftGraphRichLongRunningOperationCollectionResponse|error result = sharepointClient->listOperations(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Operations response should not be null");
}

@test:Config {
    dependsOn: [testListOperations],
    enable: isTestOnMockServer
}
function testCreateOperation() returns error? {
    MicrosoftGraphRichLongRunningOperation payload = {
        'type: "createContentType",
        status: "notStarted"
    };
    MicrosoftGraphRichLongRunningOperation|error result = sharepointClient->createOperations(testSiteId, payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result.id, (), "Created operation should have an ID");
    string? opId = result.id;
    if opId is string {
        testOperationId = opId;
    }
}

@test:Config {
    dependsOn: [testCreateOperation],
    enable: isTestOnMockServer
}
function testGetOperation() returns error? {
    MicrosoftGraphRichLongRunningOperation|error result = sharepointClient->getOperations(testSiteId, testOperationId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertEquals(result.id, testOperationId, "Operation ID should match");
}

@test:Config {
    dependsOn: [testGetOperation],
    enable: isTestOnMockServer
}
function testUpdateOperation() returns error? {
    MicrosoftGraphRichLongRunningOperation payload = {status: "running"};
    error? result = sharepointClient->updateOperations(testSiteId, testOperationId, payload);
    test:assertEquals(result, (), "Operation update should succeed without error");
}

@test:Config {
    dependsOn: [testUpdateOperation],
    enable: isTestOnMockServer
}
function testDeleteOperation() returns error? {
    error? result = sharepointClient->deleteOperations(testSiteId, testOperationId);
    test:assertEquals(result, (), "Operation delete should succeed without error");
}

// ========================
// Subsite Tests
// ========================

@test:Config {
    dependsOn: [testGetSite]
}
function testListSubsites() returns error? {
    MicrosoftGraphSiteCollectionResponse|error result = sharepointClient->listSites(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Subsites response should not be null");
}

@test:Config {
    dependsOn: [testListSubsites],
    enable: isTestOnMockServer
}
function testGetSubsite() returns error? {
    MicrosoftGraphSite|error result = sharepointClient->getSites(testSiteId, MOCK_SUBSITE_ID);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result.id, (), "Subsite should have an ID");
}

// ========================
// Global Functions
// ========================

@test:Config {}
function testGetAllSites() returns error? {
    CollectionOfSite|error result = sharepointClient->getAllSites();
    if result is error {
        test:assertFail(result.message());
    }
    MicrosoftGraphSite[]? sites = result.value;
    test:assertNotEquals(sites, (), "All sites collection should not be null");
}

@test:Config {}
function testSitesDelta() returns error? {
    CollectionOfSite1|error result = sharepointClient->delta();
    if result is error {
        test:assertFail(result.message());
    }
    MicrosoftGraphSite[]? sites = result.value;
    test:assertNotEquals(sites, (), "Delta sites collection should not be null");
}

@test:Config {
    enable: isTestOnMockServer
}
function testAddSites() returns error? {
    SitesMicrosoftGraphAddBody payload = {
        value: [{id: "mock-site-to-follow"}]
    };
    CollectionOfSite|error result = sharepointClient->add(payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Add sites response should not be null");
}

@test:Config {
    enable: isTestOnMockServer
}
function testRemoveSites() returns error? {
    SitesMicrosoftGraphAddBody payload = {
        value: [{id: "mock-site-to-unfollow"}]
    };
    CollectionOfSite|error result = sharepointClient->remove(payload);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Remove sites response should not be null");
}

// ========================
// Content Type Action Tests (getApplicableContentTypesForList)
// ========================

@test:Config {
    dependsOn: [testGetSite],
    enable: isTestOnMockServer
}
function testGetApplicableContentTypesForList() returns error? {
    CollectionOfContentType|error result = sharepointClient->siteGetApplicableContentTypesForList(testSiteId, MOCK_LIST_ID);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Applicable content types should not be null");
}

// ========================
// getActivitiesByInterval Tests
// ========================

@test:Config {
    dependsOn: [testGetSite],
    enable: isTestOnMockServer
}
function testGetActivitiesByInterval() returns error? {
    CollectionOfItemActivityStat|error result = sharepointClient->siteGetActivitiesByInterval96b0(testSiteId);
    if result is error {
        test:assertFail(result.message());
    }
    test:assertNotEquals(result, (), "Activity stats should not be null");
}