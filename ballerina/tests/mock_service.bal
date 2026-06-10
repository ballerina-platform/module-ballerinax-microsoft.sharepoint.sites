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

import ballerina/http;

const string MOCK_SITE_ID = "mock-site-id";
const string MOCK_COLUMN_ID = "mock-column-id";
const string MOCK_CONTENT_TYPE_ID = "0x01";
const string MOCK_DRIVE_ID = "mock-drive-id";
const string MOCK_PERMISSION_ID = "mock-permission-id";
const string MOCK_OPERATION_ID = "mock-operation-id";
const string MOCK_LIST_ID = "mock-list-id";
const string MOCK_SUBSITE_ID = "mock-subsite-id";
const string MOCK_COLUMN_LINK_ID = "mock-column-link-id";
const string MOCK_ACTIVITY_STAT_ID = "mock-activity-stat-id";
const string MOCK_ITEM_ACTIVITY_STAT_ID = "mock-item-activity-stat-id";

service / on new http:Listener(9090) {
    resource function 'default [string... path](http:Request req) returns http:Response|error {
        return route(path, req.method);
    }
}

isolated function route(string[] path, string method) returns http:Response {
    int len = path.length();

    // / (list sites)
    if len == 0 {
        if method == "GET" {
            return ok(mockSiteCollectionResponse());
        }
        return methodNotAllowed();
    }

    // /{seg0} or /microsoft.graph.*
    if len == 1 {
        string seg0 = path[0];
        if seg0 == "microsoft.graph.getAllSites()" && method == "GET" {
            return ok(mockSiteCollectionResponse());
        }
        if seg0 == "microsoft.graph.add" && method == "POST" {
            return ok(mockSiteCollectionResponse());
        }
        if seg0 == "microsoft.graph.remove" && method == "POST" {
            return ok(mockSiteCollectionResponse());
        }
        if seg0 == "microsoft.graph.delta()" && method == "GET" {
            return ok(mockDeltaSiteResponse());
        }
        // /{siteId}
        if method == "GET" {
            return ok(mockSite());
        }
        if method == "PATCH" {
            return noContent();
        }
        return methodNotAllowed();
    }

    // /{siteId}/{sub}
    if len == 2 {
        string sub = path[1];
        if sub.startsWith("microsoft.graph.getActivitiesByInterval") && method == "GET" {
            return ok(mockActivityStatCollectionResponse());
        }
        if sub.startsWith("microsoft.graph.getApplicableContentTypesForList") && method == "GET" {
            return ok(mockContentTypeCollectionResponse());
        }
        if sub == "columns" {
            if method == "GET" {
                return ok(mockColumnCollectionResponse());
            }
            if method == "POST" {
                return created(mockColumn());
            }
        } else if sub == "contentTypes" {
            if method == "GET" {
                return ok(mockContentTypeCollectionResponse());
            }
            if method == "POST" {
                return created(mockContentType());
            }
        } else if sub == "drive" && method == "GET" {
            return ok(mockDrive());
        } else if sub == "drives" && method == "GET" {
            return ok(mockDriveCollectionResponse());
        } else if sub == "items" && method == "GET" {
            return ok(mockBaseItemCollectionResponse());
        } else if sub == "analytics" && method == "GET" {
            return ok(mockAnalytics());
        } else if sub == "permissions" {
            if method == "GET" {
                return ok(mockPermissionCollectionResponse());
            }
            if method == "POST" {
                return created(mockPermission());
            }
        } else if sub == "operations" {
            if method == "GET" {
                return ok(mockOperationCollectionResponse());
            }
            if method == "POST" {
                return created(mockOperation());
            }
        } else if sub == "sites" && method == "GET" {
            return ok(mockSiteCollectionResponse());
        } else if sub == "lists" {
            if method == "GET" {
                return ok(mockListCollectionResponse());
            }
            if method == "POST" {
                return created(mockList());
            }
        } else if sub == "externalColumns" && method == "GET" {
            return ok(mockColumnCollectionResponse());
        }
        return notFound();
    }

    // /{siteId}/{sub}/{item}
    if len == 3 {
        string sub = path[1];
        string item = path[2];
        if sub == "columns" {
            if method == "GET" {
                return ok(mockColumn());
            }
            if method == "PATCH" || method == "DELETE" {
                return noContent();
            }
        } else if sub == "contentTypes" {
            if item == "microsoft.graph.addCopy" && method == "POST" {
                return created(mockContentType());
            }
            if item == "microsoft.graph.addCopyFromContentTypeHub" && method == "POST" {
                return created(mockContentType());
            }
            if item.startsWith("microsoft.graph.getCompatibleHubContentTypes") && method == "GET" {
                return ok(mockContentTypeCollectionResponse());
            }
            if method == "GET" {
                return ok(mockContentType());
            }
            if method == "PATCH" || method == "DELETE" {
                return noContent();
            }
        } else if sub == "drives" && method == "GET" {
            return ok(mockDrive());
        } else if sub == "permissions" {
            if method == "GET" {
                return ok(mockPermission());
            }
            if method == "PATCH" || method == "DELETE" {
                return noContent();
            }
        } else if sub == "operations" {
            if method == "GET" {
                return ok(mockOperation());
            }
            if method == "PATCH" || method == "DELETE" {
                return noContent();
            }
        } else if sub == "sites" && method == "GET" {
            return ok(mockSite());
        } else if sub == "lists" && method == "GET" {
            return ok(mockList());
        } else if sub == "analytics" {
            if item == "allTime" && method == "GET" {
                return ok(mockActivityStat());
            }
            if item == "lastSevenDays" && method == "GET" {
                return ok(mockActivityStat());
            }
            if item == "itemActivityStats" && method == "GET" {
                return ok(mockActivityStatCollectionResponse());
            }
            if item == "itemActivityStats" && method == "POST" {
                return created(mockActivityStat());
            }
        } else if sub == "items" && method == "GET" {
            return ok(mockBaseItem());
        } else if sub == "externalColumns" && method == "GET" {
            return ok(mockColumn());
        }
        return notFound();
    }

    // /{siteId}/{sub}/{itemId}/{action}
    if len == 4 {
        string sub = path[1];
        string action = path[3];
        if sub == "contentTypes" {
            if action == "microsoft.graph.isPublished()" && method == "GET" {
                return ok({"value": true});
            }
            if action == "microsoft.graph.publish" && method == "POST" {
                return noContent();
            }
            if action == "microsoft.graph.unpublish" && method == "POST" {
                return noContent();
            }
            if action == "columns" {
                if method == "GET" {
                    return ok(mockColumnCollectionResponse());
                }
                if method == "POST" {
                    return created(mockColumn());
                }
            }
            if action == "columnLinks" {
                if method == "GET" {
                    return ok(mockColumnLinkCollectionResponse());
                }
                if method == "POST" {
                    return created(mockColumnLink());
                }
            }
            if action == "columnPositions" && method == "GET" {
                return ok(mockColumnCollectionResponse());
            }
            if action == "baseTypes" && method == "GET" {
                return ok(mockContentTypeCollectionResponse());
            }
            if action == "base" && method == "GET" {
                return ok(mockContentType());
            }
        } else if sub == "columns" {
            if action == "sourceColumn" && method == "GET" {
                return ok(mockColumn());
            }
        } else if sub == "permissions" {
            if action == "microsoft.graph.grant" && method == "POST" {
                return ok(mockPermissionCollectionResponse());
            }
        } else if sub == "analytics" {
            string itemId = path[2];
            if itemId == "itemActivityStats" {
                if method == "GET" {
                    return ok(mockActivityStat());
                }
                if method == "PATCH" || method == "DELETE" {
                    return noContent();
                }
            }
        }
        return notFound();
    }

    // /{siteId}/{sub}/{itemId}/{subSub}/{subItemId}
    if len == 5 {
        string sub = path[1];
        string subSub = path[3];
        if sub == "contentTypes" {
            if subSub == "columns" {
                if method == "GET" {
                    return ok(mockColumn());
                }
                if method == "PATCH" || method == "DELETE" {
                    return noContent();
                }
            }
            if subSub == "columnLinks" {
                if method == "GET" {
                    return ok(mockColumnLink());
                }
                if method == "PATCH" || method == "DELETE" {
                    return noContent();
                }
            }
        }
        return notFound();
    }

    return notFound();
}

isolated function ok(json body) returns http:Response {
    http:Response resp = new;
    resp.statusCode = 200;
    resp.setJsonPayload(body);
    return resp;
}

isolated function created(json body) returns http:Response {
    http:Response resp = new;
    resp.statusCode = 201;
    resp.setJsonPayload(body);
    return resp;
}

isolated function noContent() returns http:Response {
    http:Response resp = new;
    resp.statusCode = 204;
    return resp;
}

isolated function notFound() returns http:Response {
    http:Response resp = new;
    resp.statusCode = 404;
    resp.setJsonPayload({"error": {"code": "ResourceNotFound", "message": "Resource not found"}});
    return resp;
}

isolated function methodNotAllowed() returns http:Response {
    http:Response resp = new;
    resp.statusCode = 405;
    resp.setJsonPayload({"error": {"code": "MethodNotAllowed", "message": "HTTP method not allowed"}});
    return resp;
}

isolated function mockSite() returns json {
    return {
        "id": MOCK_SITE_ID,
        "displayName": "Test SharePoint Site",
        "name": "testsite",
        "webUrl": "https://contoso.sharepoint.com/sites/testsite",
        "createdDateTime": "2024-01-01T00:00:00Z",
        "lastModifiedDateTime": "2024-01-01T00:00:00Z"
    };
}

isolated function mockSiteCollectionResponse() returns json {
    return {"value": [mockSite()]};
}

isolated function mockDeltaSiteResponse() returns json {
    return {
        "value": [mockSite()],
        "@odata.deltaLink": "https://graph.microsoft.com/v1.0/sites/microsoft.graph.delta()?$deltatoken=testtoken"
    };
}

isolated function mockColumn() returns json {
    return {
        "id": MOCK_COLUMN_ID,
        "displayName": "Test Column",
        "description": "A test column",
        "hidden": false,
        "required": false
    };
}

isolated function mockColumnCollectionResponse() returns json {
    return {"value": [mockColumn()]};
}

isolated function mockColumnLink() returns json {
    return {
        "id": MOCK_COLUMN_LINK_ID,
        "name": "Test Column"
    };
}

isolated function mockColumnLinkCollectionResponse() returns json {
    return {"value": [mockColumnLink()]};
}

isolated function mockContentType() returns json {
    return {
        "id": MOCK_CONTENT_TYPE_ID,
        "name": "Test Content Type",
        "description": "A test content type",
        "hidden": false,
        "isBuiltIn": false,
        "readOnly": false,
        "sealed": false
    };
}

isolated function mockContentTypeCollectionResponse() returns json {
    return {"value": [mockContentType()]};
}

isolated function mockDrive() returns json {
    return {
        "id": MOCK_DRIVE_ID,
        "name": "Documents",
        "driveType": "documentLibrary",
        "webUrl": "https://contoso.sharepoint.com/sites/testsite/Shared%20Documents",
        "createdDateTime": "2024-01-01T00:00:00Z"
    };
}

isolated function mockDriveCollectionResponse() returns json {
    return {"value": [mockDrive()]};
}

isolated function mockBaseItem() returns json {
    return {
        "id": "mock-item-id",
        "name": "test-item",
        "createdDateTime": "2024-01-01T00:00:00Z"
    };
}

isolated function mockBaseItemCollectionResponse() returns json {
    return {"value": [mockBaseItem()]};
}

isolated function mockAnalytics() returns json {
    return {"id": "mock-analytics-id"};
}

isolated function mockActivityStat() returns json {
    return {
        "id": MOCK_ACTIVITY_STAT_ID,
        "isTrending": false,
        "startDateTime": "2024-01-01T00:00:00Z",
        "endDateTime": "2024-01-07T23:59:59Z"
    };
}

isolated function mockActivityStatCollectionResponse() returns json {
    return {"value": [mockActivityStat()]};
}

isolated function mockPermission() returns json {
    return {
        "id": MOCK_PERMISSION_ID,
        "roles": ["read"],
        "expirationDateTime": "2025-12-31T00:00:00Z"
    };
}

isolated function mockPermissionCollectionResponse() returns json {
    return {"value": [mockPermission()]};
}

isolated function mockOperation() returns json {
    return {
        "id": MOCK_OPERATION_ID,
        "status": "succeeded",
        "type": "createContentType",
        "percentageComplete": 100.0
    };
}

isolated function mockOperationCollectionResponse() returns json {
    return {"value": [mockOperation()]};
}

isolated function mockList() returns json {
    return {
        "id": MOCK_LIST_ID,
        "displayName": "Test List",
        "name": "testlist",
        "webUrl": "https://contoso.sharepoint.com/sites/testsite/Lists/testlist",
        "createdDateTime": "2024-01-01T00:00:00Z"
    };
}

isolated function mockListCollectionResponse() returns json {
    return {"value": [mockList()]};
}
