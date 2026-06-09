import ballerina/io;
import ballerinax/microsoft.sharepoint.sites;

configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string tenantId = ?;
configurable string siteId = ?;

public function main() returns error? {
    sites:Client sharepointClient = check new ({
        auth: {
            clientId: clientId,
            clientSecret: clientSecret,
            tokenUrl: "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token",
            scopes: ["https://graph.microsoft.com/.default"]
        }
    });

    io:println("=== SharePoint Site Permission Audit and Cleanup Workflow ===");
    io:println("Site ID: " + siteId);
    io:println("");

    io:println("Step 1: Retrieving all permissions for the site...");
    sites:MicrosoftGraphPermissionCollectionResponse permissionCollection =
        check sharepointClient->sitesListPermissions(siteId);

    sites:MicrosoftGraphPermission[] permissions = permissionCollection.value ?: [];
    io:println("Total permissions found: " + permissions.length().toString());
    io:println("");

    if permissions.length() == 0 {
        io:println("No permissions found on the site. Audit complete.");
        return;
    }

    io:println("Step 2: Auditing permissions for overly broad access grants...");
    io:println("-----------------------------------------------------------");

    string suspiciousPermissionId = "";

    foreach sites:MicrosoftGraphPermission permission in permissions {
        string permId = permission.id ?: "unknown-id";
        string[]? roles = permission.roles;
        string rolesStr = roles is string[] ? string:'join(", ", ...roles) : "no roles";

        io:println("Permission ID: " + permId);
        io:println("  Roles: " + rolesStr);

        if roles is string[] {
            foreach string role in roles {
                if role == "owner" || role == "write" || role == "fullControl" {
                    io:println("  [WARNING] Overly broad role detected: '" + role + "' - flagging for review.");
                    if suspiciousPermissionId == "" {
                        suspiciousPermissionId = permId;
                    }
                }
            }
        }
        io:println("");
    }

    if suspiciousPermissionId == "" {
        io:println("No overly permissive access grants detected. Site security posture looks good.");
        return;
    }

    io:println("Suspicious permission identified for detailed inspection: " + suspiciousPermissionId);
    io:println("");

    io:println("Step 2b: Fetching full details of the suspicious permission...");
    sites:MicrosoftGraphPermission suspiciousPermission =
        check sharepointClient->sitesGetPermissions(siteId, suspiciousPermissionId);

    string detailedId = suspiciousPermission.id ?: "unknown";
    string[]? detailedRoles = suspiciousPermission.roles;
    string detailedRolesStr = detailedRoles is string[] ? string:'join(", ", ...detailedRoles) : "no roles";

    io:println("Full Permission Details:");
    io:println("  ID: " + detailedId);
    io:println("  Roles: " + detailedRolesStr);

    var grantedToRaw = suspiciousPermission?.grantedTo;
    if grantedToRaw is sites:MicrosoftGraphIdentitySet {
        sites:MicrosoftGraphIdentitySet grantedTo = grantedToRaw;
        io:println("  Granted To: " + grantedTo.toString());
    }

    string? expirationDateTime = suspiciousPermission?.expirationDateTime;
    if expirationDateTime is string {
        io:println("  Expiration: " + expirationDateTime);
    } else {
        io:println("  Expiration: No expiration set (permanent access - security risk!)");
    }

    io:println("");

    io:println("Step 3: Revoking overly permissive permission to enforce least privilege...");
    io:println("Deleting permission ID: " + suspiciousPermissionId);

    error? deleteResult = sharepointClient->sitesDeletePermissions(siteId, suspiciousPermissionId);

    if deleteResult is error {
        io:println("[ERROR] Failed to delete permission: " + deleteResult.message());
        return deleteResult;
    }

    io:println("[SUCCESS] Permission successfully revoked.");
    io:println("");

    io:println("=== Audit and Cleanup Summary ===");
    io:println("Site ID         : " + siteId);
    io:println("Permissions Audited: " + permissions.length().toString());
    io:println("Permission Revoked : " + suspiciousPermissionId);
    io:println("Security posture has been tightened successfully.");
}