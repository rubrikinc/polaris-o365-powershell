function Get-PolarisToken() {
    <#
    .SYNOPSIS

    Returns an API access token for a given Polaris account.

    .DESCRIPTION

    Returns an API access token for a given Polaris account, taking the URL, username and password.

    .PARAMETER Username
    Polaris username.

    .PARAMETER Password
    Polaris password.

    .PARAMETER Password
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .INPUTS

    None. You cannot pipe objects to Get-PolarisToken.

    .OUTPUTS

    System.String. Get-PolarisToken returns a string containing the access token.

    .EXAMPLE

    PS> $token = Get-PolarisToken -Username $username -Password $password -PolarisURL $url
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String]$Username,
        [Parameter(Mandatory = $True)]
        [String]$Password,
        [Parameter(Mandatory = $True)]
        [String]$PolarisURL
    )
    $headers = @{
        'Content-Type' = 'application/json';
        'Accept'       = 'application/json';
    }
    $payload = @{
        "username" = $Username;
        "password" = $Password;
    }
    # Replace com/ with com to prevent 404 error
    $PolarisURL = $PolarisURL -replace '(.*?)com/(.*)', '$1com$2' 
    $endpoint = $PolarisURL + '/api/session'
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    return $response.access_token
}

function Get-PolarisTokenServiceAccount() {
    <#
    .SYNOPSIS

     Connect to a Polaris Account using a Service Account. This is the recommended connection method. 

    .DESCRIPTION

    Returns an API access token for a given Polaris account. The cmdlet requires a Service Account JSON file stored at ~/.rubrik/polaris-service-account.json.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisTokenServiceAccount.

    .OUTPUTS

    System.String. Get-PolarisTokenServiceAccount returns a string containing the access token required to connect to the Polaris GraphQL API.
 
    .EXAMPLE

    PS> $token = Get-PolarisTokenServiceAccount
    #>


    try {
        $serviceAccountFile = Get-Content -Path "~/.rubrik/polaris-service-account.json" -ErrorAction Stop | ConvertFrom-Json 
    }
    catch {
        $errorMessage = $_.Exception | Out-String

        if($errorMessage.Contains('because it does not exist')) {
            throw "The Service Account JSON secret file was not found. Ensure the file is location at ~/.rubrik/polaris-service-account.json."
        } 
        
        throw $_.Exception
        
    }


    $payload = @{
        grant_type = "client_credentials";
        client_id = $serviceAccountFile.client_id;
        client_secret = $serviceAccountFile.client_secret
    }   

    $missingServiceAccount = @()
    if ($serviceAccountFile.client_id -eq $null) {
        $missingServiceAccount += "'client_id'"
    }

    if ($serviceAccountFile.client_secret -eq $null) {
        $missingServiceAccount += "'client_secret'"
    }

    if ($serviceAccountFile.access_token_uri -eq $null) {
        $missingServiceAccount += "'access_token_uri'"
    }


    if ($missingServiceAccount.count -gt 0){
        throw "The Service Account JSON secret file is missing the required paramaters: $missingServiceAccount"
    }


    $headers = @{
        'Content-Type' = 'application/json';
        'Accept'       = 'application/json';
    }
   
    
    $response = Invoke-RestMethod -Method POST -Uri $serviceAccountFile.access_token_uri -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    return $response.access_token

   
}

function Get-PolarisSLA() {
    <#
    .SYNOPSIS

    Returns the SLA Domains from a given Polaris account.

    .DESCRIPTION

    Returns SLA Domains for a given Polaris account. This can be used to return
    based on a name query, by using the 'Name' parameter.

    .PARAMETER Token
    Polaris access token, get this using the 'Get-PolarisTokenServiceAccount' or 'Get-PolarisToken' command.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .PARAMETER Name
    Optional. The name of the required SLA Domain. If none is provided, all
    SLAs are returned.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisSLA.

    .OUTPUTS

    System.Object. Get-PolarisSLA returns an array containing the ID, Name,
    and Description of the returned SLA Domains.

    .EXAMPLE

    PS> Get-PolarisSLA -Token $token -PolarisURL $url -Name 'Bronze'
    name   id                                   description
    ----   --                                   -----------
    Bronze 00000000-0000-0000-0000-000000000002 Bronze SLA
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String]$Token,
        [Parameter(Mandatory = $True)]
        [String]$PolarisURL,
        [Parameter(Mandatory = $False)]
        [String]$Name
    )

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $payload = @{
        "operationName" = "SLAList";
        "variables"     = @{"first" = 20; "name" = $Name };
        "query"         = "query SLAList(`$after: String, `$first: Int, `$name: String) {
            globalSlaConnection(after: `$after, first: `$first, filter: [{field: NAME, text: `$name}]) {
                edges {
                    node {
                        id
                        name
                    }
                }
                pageInfo {
                    endCursor
                    hasNextPage
                    hasPreviousPage
                }
            }
        }"
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers

    $sla_detail = @()

    foreach ($edge in $response.data.globalSlaConnection.edges) {
        $row = '' | Select-Object name, id, description
        $row.name = $edge.node.name
        $row.id = $edge.node.id
        $sla_detail += $row
    }

    return $sla_detail
}

function Get-PolarisO365Subscriptions {
    <#
    .SYNOPSIS

    Returns all O365 subscriptions from a given Polaris account.

    .DESCRIPTION

    Returns an array of Office 365 subscriptions from a given Polaris account, taking
    an API token, and Polaris URL.

    .PARAMETER Token
    Polaris API Token.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .INPUTS

    None. You cannot pipe objects to Get-PolarisO365Subscriptions.

    .OUTPUTS

    System.Object. Get-PolarisO365Subscriptions returns an array containing the ID, Name,
    status, count of users, count of unprotected users, and SLA details for the
    returned O365 Subscriptions.

    .EXAMPLE

    PS> Get-PolarisO365Subscriptions -Token $token -PolarisURL $url

    name                    : MyO365Subscription
    id                      : 12345678-1234-1234-abcd-123456789012
    status                  : ACTIVE
    usersCount              : 15468
    unprotectedUsersCount   : 1018
    effectiveSlaDomainName  : UNPROTECTED
    configuredSlaDomainName : UNPROTECTED
    effectiveSlaDomainId    : UNPROTECTED
    configuredSlaDomainId   : UNPROTECTED
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String]$Token,
        [Parameter(Mandatory = $True)]
        [String]$PolarisURL
    )

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    # Get a list of all the orgs

    $payload = @{
        "operationName" = "O365OrgList";
        "query"         = "query O365OrgList(`$after: String, `$first: Int) {
            o365Orgs(after: `$after, first: `$first, filter: []) {
                edges {
                    node {
                        id
                    }
                }
                pageInfo {
                    endCursor
                    hasNextPage
                    hasPreviousPage
                }
            }
        }";
        "variables"     = @{
            "after" = $null;
            "first" = $null;
        }
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers

    $org_ids = @()
    foreach ($org in $response.data.o365Orgs.edges) {
        $org_ids += $org.node.id
    }

    # For each org let's get the details

    $org_details = @()

    foreach ($org_id in $org_ids) {
        $payload = @{
            "operationName" = "o365OrgCard";
            "query"         = "query o365OrgCard(`$id: UUID!) {
                o365Org(fid: `$id) {
                    id
                    status
                    name
                    effectiveSlaDomain {
                        name
                        id
                    }
                    configuredSlaDomain {
                        name
                        id
                    }
                    childConnection(filter: []) {
                        count
                    }
                    unprotectedUsersCount
                }
            }";
            "variables"     = @{
                "id" = "$org_id";
            }
        }

        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers

        $row = '' | Select-Object name, id, status, usersCount, unprotectedUsersCount, effectiveSlaDomainName, configuredSlaDomainName, effectiveSlaDomainId, configuredSlaDomainId
        $row.name = $response.data.o365Org.name
        $row.id = $response.data.o365Org.id
        $row.status = $response.data.o365Org.status
        $row.usersCount = $response.data.o365Org.childConnection.count
        $row.unprotectedUsersCount = $response.data.o365Org.unprotectedUsersCount
        $row.effectiveSlaDomainName = $response.data.o365Org.effectiveSlaDomain.name
        $row.configuredSlaDomainName = $response.data.o365Org.configuredSlaDomain.name
        $row.effectiveSlaDomainId = $response.data.o365Org.effectiveSlaDomain.id
        $row.configuredSlaDomainId = $response.data.o365Org.effectiveSlaDomain.name
        $org_details += $row
    }

    return $org_details
}

function Get-PolarisO365Mailboxes() {
    <#
    .SYNOPSIS

    Returns all O365 mailboxes for a given subscription in a given Polaris account.

    .DESCRIPTION

    Returns an array of Office 365 mailboxes from a given subscription and Polaris account, taking
    an API token, Polaris URL, and subscription ID.

    .PARAMETER Token
    Polaris API Token.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .PARAMETER SubscriptionID
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the
    'Get-PolarisO365Subscriptions' command.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisO365Mailboxes.

    .OUTPUTS

    System.Object. Get-PolarisO365Mailboxes returns an array containing the ID, Name,
    email address, and SLA details for the returned O365 mailboxes.

    .EXAMPLE

    PS> Get-PolarisO365Mailboxes -Token $token -PolarisURL $url -SubscriptionId $my_sub.id

    name                        : Milan Kundera
    id                          : 12341234-1234-1234-abcd-123456789012
    userPrincipalName           : milan.kundera@mydomain.onmicrosoft.com
    slaAssignment               : Direct
    effectiveSlaDomainName      : Gold
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String]$Token,
        [Parameter(Mandatory = $True)]
        [String]$PolarisURL,
        [Parameter(Mandatory = $True)]
        [String]$SubscriptionId
    )

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    # get users

    $node_array = @()

    $payload = @{
        "operationName" = "O365MailboxList";
        "query"         = "query O365MailboxList(`$first: Int!, `$after: String, `$orgId: UUID!, `$filter: [Filter!]!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            o365Mailboxes(o365OrgId: `$orgId, after: `$after, first: `$first, filter: `$filter, sortBy: `$sortBy, sortOrder: `$sortOrder) {
                edges {
                    node {
                        id
                        name
                        userPrincipalName
                        effectiveSlaDomain {
                            name
                        }
                        slaAssignment
                    }
                }
                pageInfo {
                    endCursor
                    hasNextPage
                    hasPreviousPage
                }
            }
        }";
        "variables"     = @{
            "after"     = $null;
            "filter"    = @(
                @{
                    "field" = "IS_RELIC";
                    "texts" = @("false");
                };
            )
            "first"     = 100;
            "orgId"     = $SubscriptionId;
            "sortBy"    = "EMAIL_ADDRESS";
            "sortOrder" = "ASC";
        }
    }
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    $node_array += $response.data.o365Mailboxes.edges
    # get all pages of results
    while ($response.data.o365Mailboxes.pageInfo.hasNextPage) {
        $payload.variables.after = $response.data.o365Mailboxes.pageInfo.endCursor
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $node_array += $response.data.o365Mailboxes.edges
    }

    $mailbox_details = @()

    foreach ($node in $node_array) {
        $row = '' | Select-Object name, id, userPrincipalName, slaAssignment, effectiveSlaDomainName
        $row.name = $node.node.name
        $row.id = $node.node.id
        $row.userPrincipalName = $node.node.userPrincipalName
        $row.slaAssignment = $node.node.slaAssignment
        $row.effectiveSlaDomainName = $node.node.effectiveSlaDomain.name
        $mailbox_details += $row
    }

    return $mailbox_details
}


function Get-PolarisO365OneDrives() {
    <#
    .SYNOPSIS
    Returns all O365 OneDrive users for a given subscription in a given Polaris account.
    .DESCRIPTION
    Returns an array of Office OneDrive 365 users from a given subscription and Polaris account, taking
    an API token, Polaris URL, and subscription ID.
    .PARAMETER Token
    Polaris API Token.
    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'
    .PARAMETER SubscriptionID
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the
    'Get-PolarisO365Subscriptions' command.
    .INPUTS
    None. You cannot pipe objects to Get-PolarisO365OneDrives.
    .OUTPUTS
    System.Object. Get-PolarisO365OneDriveUsers returns an array containing the ID, Name,
    email address, and SLA details for the returned O365 OneDrive users.
    .EXAMPLE
    PS> Get-PolarisO365OneDrives -Token $token -PolarisURL $url -SubscriptionId $my_sub.id
    name                   : Milan Kundera
    id                     : 12341234-1234-1234-abcd-123456789012
    emailAddress           : milan.kundera@mydomain.onmicrosoft.com
    slaAssignment          : Direct
    effectiveSlaDomainName : Gold
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Token,
        [Parameter(Mandatory=$True)]
        [String]$PolarisURL,
        [Parameter(Mandatory=$True)]
        [String]$SubscriptionId
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    # get users

    $node_array = @()

    $payload = @{
        "operationName" = "O365OnedriveList";
        "query" = "query O365OnedriveList(`$first: Int!, `$after: String, `$orgId: UUID!, `$filter: [Filter!]!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            o365Onedrives(o365OrgId: `$orgId, after: `$after, first: `$first, filter: `$filter, sortBy: `$sortBy, sortOrder: `$sortOrder) {
                edges {
                    node {
                        id
                        userName
                        name
                        userPrincipalName
                        effectiveSlaDomain {
                            id
                            name
                        }
                        authorizedOperations
                        slaAssignment
                    }
                }
                pageInfo {
                    endCursor
                    hasNextPage
                    hasPreviousPage
                }
            }
        }";
        "variables" = @{
            "after" = $null;
            "filter" = @(
                @{
                    "field" = "IS_RELIC";
                    "texts" = @("false")
                };
            )
            "first" = 100;
            "orgId" = $SubscriptionId;
            "sortBy" = "EMAIL_ADDRESS";
            "sortOrder" = "ASC";
        }

    }
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    $node_array += $response.data.o365Onedrives.edges
    # get all pages of results
    while ($response.data.o365Onedrives.pageInfo.hasNextPage) {
        $payload.variables.after = $response.data.o365Onedrives.pageInfo.endCursor
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $node_array += $response.data.o365Onedrives.edges
    }

    $user_details = @()

    foreach ($node in $node_array) {
        $row = '' | Select-Object name,id,userName,userPrincipalName,slaAssignment,effectiveSlaDomainName
        $row.name = $node.node.name
        $row.id = $node.node.id
        $row.userName = $node.node.userName
        $row.userPrincipalName = $node.node.userPrincipalName
        $row.slaAssignment = $node.node.slaAssignment
        $row.effectiveSlaDomainName = $node.node.effectiveSlaDomain.name
        $user_details += $row
    }

    return $user_details
}

function Get-PolarisO365OneDriveSnapshot() {
    <#
    .SYNOPSIS
    Return the ID and Storage Location for the last OneDrive snapshot taken.

    .DESCRIPTION
    Returns an array that contains the Snapshot ID and Storage Location for the last snapshot taken on a OneDrive account. 
    This information can then be utilized in Restore-PolarisO365OneDrive.

    .PARAMETER Token
    Polaris API Token.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .PARAMETER OneDriveID
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the
    'Get-PolarisO365Subscriptions' command.

    .INPUTS
    None. You cannot pipe objects to Get-PolarisO365OneDriveSnapshot.

    .OUTPUTS
    System.Object. Get-PolarisO365OneDriveSnapshot returns an array containing lastSnapshotId and
    lastSnapshotStorageLocation.
    
    .EXAMPLE
    PS> Get-PolarisO365OneDriveSnapshot -Token $token -PolarisURL $url -OneDriveID $OneDriveID
    lastSnapshotId                       lastSnapshotStorageLocation
    --------------                       ---------------------------
    15e80edc-3211-412d-8cd2-1f5e33c52863                          46
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Token
        [Parameter(Mandatory=$True)]
        [String]$OneDriveId,
        [Parameter(Mandatory=$True)]
        [String]$PolarisURL
        
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $payload = @{
        "operationName" = "O365OnedriveList";
        "query" = "query O365OnedriveList(`$snappableId: UUID!) {
            o365Onedrive(snappableFid: `$snappableId) {
              newestSnapshot {
                id
              }
              snapshotConnection(first: 1, sortOrder: Desc) {
                nodes {
                  sequenceNumber
                }
              }
            }
          }";
        "variables" = @{
            "snappableId" = $OneDriveId;
        }

    }

   
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    
    $row = '' | Select-Object lastSnapshotId,lastSnapshotStorageLocation
    $row.lastSnapshotId = $response.data.o365Onedrive.newestSnapshot.id
    $row.lastSnapshotStorageLocation = $response.data.o365Onedrive.snapshotConnection.nodes.sequenceNumber
    
    return $row
}

function Restore-PolarisO365OneDrive() {
    <#
    .SYNOPSIS
    Restore a Users entire OneDrive
    .DESCRIPTION
    Restore a users entire OneDrive to either it's original location or by created a Download link through Rubrik.
    .PARAMETER Token
    Polaris API Token.
    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'
    .PARAMETER OneDriveId
    The ID of the OneDrive you wish to restore
    .PARAMETER SnapshotId
    The ID of the snapshot you wish to restore.
    .PARAMETER SnapshotStorageLocation
    The ID of the snapshot storage location you wish to restore.
    .PARAMETER RecoveryOption
    The type of restore job you wish to use. Specify Original to restore to the Original OneDrive of Download to create a download link through Rubrik.
    .INPUTS
    None. You cannot pipe objects to Restore-PolarisO365OneDrive.
    .OUTPUTS
    String. The taskchainID of the Restore job which can be used to monitor the jobs progress.
    .EXAMPLE
    PS> Restore-PolarisO365OneDrive -PolarisURL $url -OneDriveId $user.id -SnapshotId $snapshotDetails.lastSnapshotId -SnapshotStorageLocation $snapshotDetails.lastSnapshotStorageLocation -RecoveryOption "Download"
    123594e0-1477-4be8-b6a2-f04174336a98
    #>

    param(
        [Parameter(Mandatory=$True)]
        [ValidateSet("Download", "Original")]
        [String]$RecoveryOption,
        [Parameter(Mandatory=$True)]
        [String]$OneDriveId,
        [Parameter(Mandatory=$True)]
        [String]$SnapshotStorageLocation,
        [Parameter(Mandatory=$True)]
        [String]$SnapshotId,
        [Parameter(Mandatory=$True)]
        [String]$PolarisURL
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'
    
    if ($RecoveryOption -eq "Download") {
        $actionType = "EXPORT_SNAPPABLE"
        
    } else {
        $actionType = "RESTORE_SNAPPABLE"
    }


    $payload = @{
        "operationName" = "O365RestoreOnedriveMutation";
        "query" = "mutation O365RestoreOnedriveMutation(`$filesToRestore: [FileInfo!]!, `$foldersToRestore: [FolderInfo!]!, `$destOnedriveUUID: UUID!, `$sourceOnedriveUUID: UUID!, `$restoreFolderPath: String!, `$actionType: O365RestoreActionType!) {
            restoreO365Snappable(snappableType: ONEDRIVE, sourceSnappableUUID: `$sourceOnedriveUUID, destSnappableUUID: `$destOnedriveUUID, snappableRestoreConfig: {OneDriveRestoreConfig: {FilesToRestore: `$filesToRestore, FoldersToRestore: `$foldersToRestore, RestoreFolderPath: `$restoreFolderPath}}, actionType: `$actionType) {
              taskchainId
            }
          }";
        "variables" = @{
            "foldersToRestore" = @(
                @{
                    "FolderID" = "root";
                    "FolderName" = "OneDrive";
                    "SnapshotID" = $SnapshotId;
                    "FolderSize" = 0;
                    "SnapshotNum" = [int]$SnapshotStorageLocation;

                }
            );
            "filesToRestore" = @();
            "restoreFolderPath" = "";
            "sourceOnedriveUUID" = $OneDriveId;
            "destOnedriveUUID" = $OneDriveId;
            "actionType" = $actionType;
        }

    }
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers

    return $response.data.restoreO365Snappable.taskchainId
    
}

function Get-PolarisO365Mailbox() {
    <#
    .SYNOPSIS

    Returns a filtered list of O365 mailboxes for a given subscription in a given Polaris account.

    .DESCRIPTION

    Returns a filtered list of Office 365 mailboxes from a given subscription and Polaris account, taking
    an API token, Polaris URL, subscription ID, and search string.

    .PARAMETER Token
    Polaris API Token.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .PARAMETER SubscriptionID
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the
    'Get-PolarisO365Subscriptions' command.

    .PARAMETER SearchString
    Search string, used to filter mailbox's name or user principal name.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisO365Mailbox.

    .OUTPUTS

    System.Object. Get-PolarisO365Mailbox returns an array containing the ID, Name,
    email address, and SLA details for the returned O365 mailboxes.

    .EXAMPLE

    PS> Get-PolarisO365Mailbox -Token $token -PolarisURL $url -SubscriptionId $my_sub.id -SearchString 'Milan'

    name                   : Milan Kundera
    id                     : 12341234-1234-1234-abcd-123456789012
    emailAddress           : milan.kundera@mydomain.onmicrosoft.com
    slaAssignment          : Direct
    effectiveSlaDomainName : Gold
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String]$Token,
        [Parameter(Mandatory = $True)]
        [String]$PolarisURL,
        [Parameter(Mandatory = $True)]
        [String]$SubscriptionId,
        [Parameter(Mandatory = $True)]
        [String]$SearchString
    )

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    # get users

    $node_array = @()

    $payload = @{
        "operationName" = "O365MailboxList";
        "query"         = "query O365MailboxList(`$first: Int!, `$after: String, `$orgId: UUID!, `$filter: [Filter!]!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            o365Mailboxes(o365OrgId: `$orgId, after: `$after, first: `$first, filter: `$filter, sortBy: `$sortBy, sortOrder: `$sortOrder) {
                edges {
                    node {
                        id
                        name
                        userPrincipalName
                        effectiveSlaDomain {
                            name
                        }
                        slaAssignment
                    }
                }
                pageInfo {
                    endCursor
                    hasNextPage
                    hasPreviousPage
                }
            }
        }";
        "variables"     = @{
            "after"     = $null;
            "filter"    = @(
                @{
                    "field" = "IS_RELIC";
                    "texts" = @("false");
                },
                @{
                    "field" = "NAME_OR_EMAIL_ADDRESS";
                    "texts" = @($SearchString);
                }
            );
            "first"     = 100;
            "orgId"     = $SubscriptionId;
            "sortBy"    = "EMAIL_ADDRESS";
            "sortOrder" = "ASC";
        }
    }
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    $node_array += $response.data.o365Mailboxes.edges
    # get all pages of results
    while ($response.data.o365Mailboxes.pageInfo.hasNextPage) {
        $payload.variables.after = $response.data.o365Mailboxes.pageInfo.endCursor
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $node_array += $response.data.o365Mailboxes.edges
    }

    $mailbox_details = @()

    foreach ($node in $node_array) {
        $row = '' | Select-Object name, id, userPrincipalName, slaAssignment, effectiveSlaDomainName
        $row.name = $node.node.name
        $row.id = $node.node.id
        $row.userPrincipalName = $node.node.userPrincipalName
        $row.slaAssignment = $node.node.slaAssignment
        $row.effectiveSlaDomainName = $node.node.effectiveSlaDomain.name
        $mailbox_details += $row
    }

    return $mailbox_details
}

function Get-PolarisO365OneDrive() {
    <#
    .SYNOPSIS
    Returns a filtered list of O365 OneDrive users for a given subscription in a given Polaris account.
    .DESCRIPTION
    Returns a filtered list of Office 365 OneDrive users from a given subscription and Polaris account, taking
    an API token, Polaris URL, subscription ID, and search string.
    .PARAMETER Token
    Polaris API Token.
    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'
    .PARAMETER SubscriptionID
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the
    'Get-PolarisO365Subscriptions' command.
    .PARAMETER SearchString
    Search string, used to filter user's name or email address.
    .INPUTS
    None. You cannot pipe objects to Get-PolarisO365MailboxUser.
    .OUTPUTS
    System.Object. Get-PolarisO365OneDrives returns an array containing the ID, Name,
    email address, and SLA details for the returned O365 OneDrive users.
    .EXAMPLE
    PS> Get-PolarisO365OneDrives -Token $token -PolarisURL $url -SubscriptionId $my_sub.id
    name                   : Milan Kundera
    id                     : 12341234-1234-1234-abcd-123456789012
    emailAddress           : milan.kundera@mydomain.onmicrosoft.com
    slaAssignment          : Direct
    effectiveSlaDomainName : Gold
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Token,
        [Parameter(Mandatory=$True)]
        [String]$PolarisURL,
        [Parameter(Mandatory=$True)]
        [String]$SubscriptionId,
        [Parameter(Mandatory=$True)]
        [String]$SearchString
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    # get users

    $node_array = @()

    $payload = @{
        "operationName" = "O365OnedriveList";
        "query" = "query O365OnedriveList(`$first: Int!, `$after: String, `$orgId: UUID!, `$filter: [Filter!]!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            o365Onedrives(o365OrgId: `$orgId, after: `$after, first: `$first, filter: `$filter, sortBy: `$sortBy, sortOrder: `$sortOrder) {
                edges {
                    node {
                        id
                        userName
                        name
                        userPrincipalName
                        effectiveSlaDomain {
                            id
                            name
                        }
                        authorizedOperations
                        slaAssignment
                    }
                }
                pageInfo {
                    endCursor
                    hasNextPage
                    hasPreviousPage
                }
            }
        }";
        "variables" = @{
            "after" = $null;
            "filter" = @(
                @{
                    "field" = "IS_RELIC";
                    "texts" = @("false")
                },
                @{
                    "field" = "NAME_OR_EMAIL_ADDRESS";
                    "texts" = @($SearchString);
                };
            )
            "first" = 100;
            "orgId" = $SubscriptionId;
            "sortBy" = "EMAIL_ADDRESS";
            "sortOrder" = "ASC";
        }

    }
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    $node_array += $response.data.o365Onedrives.edges
    # get all pages of results
    while ($response.data.o365Onedrives.pageInfo.hasNextPage) {
        $payload.variables.after = $response.data.o365Onedrives.pageInfo.endCursor
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $node_array += $response.data.o365Onedrives.edges
    }

    $user_details = @()

    foreach ($node in $node_array) {
        $row = '' | Select-Object name,id,userName,userPrincipalName,slaAssignment,effectiveSlaDomainName
        $row.name = $node.node.name
        $row.id = $node.node.id
        $row.userName = $node.node.userName
        $row.userPrincipalName = $node.node.userPrincipalName
        $row.slaAssignment = $node.node.slaAssignment
        $row.effectiveSlaDomainName = $node.node.effectiveSlaDomain.name
        $user_details += $row
    }

    return $user_details
}

function Get-PolarisO365SharePoint() {
    <#
    .SYNOPSIS
    Returns a filtered list of O365 SharePoint sites and/or document libraries for a given subscription in a given Polaris account.
    .DESCRIPTION
    Returns a filtered list of Office 365 SharePoint sites and/or document libraries from a given subscription and Polaris account, taking
    an API token, Polaris URL, subscription ID, and search string.
    .PARAMETER Token
    Polaris API Token.
    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'
    .PARAMETER SubscriptionId
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the
    'Get-PolarisO365Subscriptions' command.
    .PARAMETER SearchString
    Search string, used to filter site or document library name.
    .PARAMETER Includes
    It indidates if the returned object includes only SharePoint sites, document libraries or both. The value can only be 'SitesOnly', 'DocumentLibrariesOnly'.
    if it's not specified, it returns both sites and document libraries by default.
    .INPUTS
    None. You cannot pipe objects to Get-PolarisO365SharePoint.
    .OUTPUTS
    System.Object. Get-PolarisO365SharePoint returns an array containing the ID, Name,
    and SLA details for the returned O365 SharePoint sites and/or document libraries.
    .EXAMPLE
    PS> Get-PolarisO365SharePoint -Token $token -PolarisURL $url -SubscriptionId $my_sub.id -Includes 'SitesOnly' -SearchString 'test'
    name                   : Milan Kundera
    id                     : 12341234-1234-1234-abcd-123456789012
    type                   : O365Site
    slaAssignment          : Direct
    effectiveSlaDomainName : Gold
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Token,
        [Parameter(Mandatory=$True)]
        [String]$PolarisURL,
        [Parameter(Mandatory=$True)]
        [String]$SubscriptionId,
        [Parameter(Mandatory=$false)]
        [String]$SearchString,
        [Parameter(Mandatory=$false)]
        [ValidateSet("SitesOnly", "DocumentLibrariesOnly")]
        [String]$Includes
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $o365Sites = $null
    $o365SharepointDrives = $null

    $payload = @{
        "query"         = "";
        "variables"     = @{
            "after"     = $null;
            "filter"    = @(
                @{
                    "field" = "IS_RELIC";
                    "texts" = @("false");
                },
                @{
                    "field" = "NAME_OR_EMAIL_ADDRESS";
                    "texts" = @($SearchString);
                }
            );
            "first"     = 100;
            "o365OrgId" = $SubscriptionId;
        }
    }

    $querySites = "o365Sites(after: `$after, o365OrgId: `$o365OrgId, filter: `$filter, first: `$first, sortBy: `$sortBy, sortOrder: `$sortOrder) {
        nodes {
            id
            name
            parentId
            title
            url
            hierarchyLevel
            effectiveSlaDomain {
                id
                name
            }
            objectType
            slaAssignment
        }
        pageInfo {
            endCursor
            hasNextPage
            hasPreviousPage
        }
    }"

    $queryDrives = "o365SharepointDrives(after: `$after, o365OrgId: `$o365OrgId, filter: `$filter, first: `$first, sortBy: `$sortBy, sortOrder: `$sortOrder) {
        nodes {
            id
            naturalId
            name
            parentId
            totalStorageInBytes
            usedStorageInBytes
            effectiveSlaDomain {
                id
                name
            }
            objectType
            slaAssignment
            onDemandSnapshotCount
        }
        pageInfo {
            endCursor
            hasNextPage
            hasPreviousPage
        }
    }"

    if ($Includes -eq "SitesOnly") {
        $payload.query = "query O365SharepointQuery(`$after: String, `$o365OrgId:UUID!, `$filter: [Filter!], `$first: Int!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            $($querySites)
        }"
 
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $o365Sites = $response.data.o365Sites
    } elseif ($Includes -eq "DocumentLibrariesOnly") {
        $payload.query = "query O365SharepointQuery(`$after: String, `$o365OrgId:UUID!, `$filter: [Filter!], `$first: Int!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            $($queryDrives)
        }"

        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $o365SharepointDrives = $response.data.o365SharepointDrives
    } else {
        $payload.query = "query O365SharepointQuery(`$after: String, `$o365OrgId:UUID!, `$filter: [Filter!], `$first: Int!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            $($querySites)
            $($queryDrives)
        }"

        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $o365Sites = $response.data.o365Sites
        $o365SharepointDrives = $response.data.o365SharepointDrives
     }

     $node_array += @()
     $sharepoint_details = @()

     if ($null -ne $o365Sites) {
        $node_array += $o365Sites.nodes
        while ($o365Sites.pageInfo.hasNextPage) {
            $payload.variables.after = $o365Sites.pageInfo.endCursor
            $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
            $o365Sites = $response.data.o365Sites
            $node_array += $o365Sites.nodes
        }
    }

    if ($null -ne $o365SharepointDrives) {
        $node_array += $o365SharepointDrives.nodes
        while ($o365SharepointDrives.pageInfo.hasNextPage) {
            $payload.variables.after = $o365SharepointDrives.pageInfo.endCursor
            $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
            $o365SharepointDrives = $response.data.o365SharepointDrives
            $node_array += $o365SharepointDrives.nodes
        }
    }

    foreach ($node in $node_array) {
        $row = '' | Select-Object name,id,type,slaAssignment,effectiveSlaDomainName
        $row.name = $node.name
        $row.id = $node.id
        $row.type = $node.objectType
        $row.slaAssignment = $node.slaAssignment
        $row.effectiveSlaDomainName = $node.effectiveSlaDomain.name
        $sharepoint_details += $row
    }

    return $sharepoint_details
}

function Set-PolarisO365ObjectSla() {
    <#
    .SYNOPSIS

    Sets the SLA Domain for a selected Office 365 object (mailbox, onedrive or subscription).

    .DESCRIPTION

    Sets the protection for an O365 mailbox, onedrive or subscription in a given Polaris account, taking
    an API token, Polaris URL, object ID, and SLA ID.

    .PARAMETER Token
    Polaris API Token.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .PARAMETER ObjectID
    The object ID(s) for an O365 user or subscription. Can be obtained using 'Get-PolarisO365Mailbox', 'Get-PolarisO365OneDrive',
    'Get-PolarisO365Mailboxes', 'Get-PolarisO365OneDrives', or 'Get-PolarisO365Subscriptions' commands. This can take an array of object IDs.

    .PARAMETER SlaID
    The SLA ID for an SLA Domain. Can be obtained through the 'Get-PolarisSLA' command. Use the string
    'UNPROTECTED' to remove any SLA from this object, or the string 'DONOTPROTECT' to explicitly not protect
    this or any child objects.

    .INPUTS

    None. You cannot pipe objects to Set-PolarisO365ObjectSla.

    .OUTPUTS

    System.String. This returns the string 'Success' if the modification was successful, or throws an
    error if the command is not successful.

    .EXAMPLE

    PS> Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_mailbox.id -SlaID $my_sla.id
    Success

    .EXAMPLE

    PS> Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_onedrive.id -SlaID 'DONOTPROTECT'
    Success

    .EXAMPLE

    PS> Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_subscription.id -SlaID 'UNPROTECTED'
    Success
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String]$Token,
        [Parameter(Mandatory = $True)]
        [String]$PolarisURL,
        [Parameter(Mandatory = $True)]
        [String[]]$ObjectID,
        [Parameter(Mandatory = $True)]
        [String]$SlaID
    )

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $payload = @{
        "operationName" = "AssignSLA";
        "variables"     = @{
            "globalSlaAssignType"  = "protectWithSlaId";
            "globalSlaOptionalFid" = $SlaID;
            "objectIds"            = $ObjectID;
        };
        "query"         = "mutation AssignSLA(`$globalSlaOptionalFid: UUID, `$globalSlaAssignType: SlaAssignTypeEnum!, `$objectIds: [UUID!]!) {
            assignSla(globalSlaOptionalFid: `$globalSlaOptionalFid, globalSlaAssignType: `$globalSlaAssignType, objectIds: `$objectIds) {
                success
            }
        }";
    }

    if ($SlaID -eq 'UNPROTECTED') {
        $payload['variables']['globalSlaOptionalFid'] = $null
        $payload['variables']['globalSlaAssignType'] = 'noAssignment'
    }

    if ($SlaID -eq 'DONOTPROTECT') {
        $payload['variables']['globalSlaOptionalFid'] = $null
        $payload['variables']['globalSlaAssignType'] = 'doNotProtect'
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    if ($response.data.assignSla.success -eq $true) {
        return 'Success'
    }
    else {
        throw 'Issue assigning SLA domain to object'
    }
}


function Get-PolarisJob() {
    <#
    .SYNOPSIS

    Returns the details for a given Polaris job.

    .DESCRIPTION

    Returns detailed information on a Polaris job. 

    .PARAMETER Token
    Polaris access token, get this using the 'Get-PolarisTokenServiceAccount' or 'Get-PolarisToken' command.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .PARAMETER JobId
    The ID of the job you wish to get details on.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisJob.

    .OUTPUTS

    System.Object. Get-PolarisJob returns an array containing the activityConnection, id,
    lastUpdated,  lastActivityType, lastActivityStatus, objectId, objectName, severity, 
    objectType, progress, isCancelable, and startTime of the returned job.

    .EXAMPLE

    PS> Get-PolarisJob -PolarisURL $url -Token $token -JobId $JobId
    activityConnection : {@{activityInfo=; message=Started export of Drew Russell Microsoft 365 onedrive data; status=TaskSuccess; time=10/5/2021 3:45:19 PM; severity=Info}}
    id                 : 7592590
    activitySeriesId   : 1d3594e0-1488-4ba4-b6a2-f04174336a73
    lastUpdated        : 10/5/2021 3:45:19 PM
    lastActivityType   : Recovery
    lastActivityStatus : TaskSuccess
    objectId           : e4092af5-009b-47e9-9412-bcd57924cfa6
    objectName         : Drew Russell
    objectType         : O365Onedrive
    severity           : Info
    progress           : 100%
    isCancelable       : False
    startTime          : 10/5/2021 3:45:19 PM
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String]$Token,
        [Parameter(Mandatory = $True)]
        [String]$PolarisURL,
        [Parameter(Mandatory = $True)]
        [String]$JobId
        
    )

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $payload = @{
        "operationName" = "EventSeriesDetailsQuery";
        "variables"     = @{
            "activitySeriesId"  = $JobId;
            "clusterUuid" = "00000000-0000-0000-0000-000000000000";
        };
        "query"         = "query EventSeriesDetailsQuery(`$activitySeriesId: UUID!, `$clusterUuid: UUID) {
            activitySeries(activitySeriesId: `$activitySeriesId, clusterUuid: `$clusterUuid) {
              activityConnection {
                nodes {
                  activityInfo
                  message
                  status
                  time
                  severity
                }
              }
              id
              fid
              activitySeriesId
              lastUpdated
              lastActivityType
              lastActivityStatus
              objectId
              objectName
              objectType
              severity
              progress
              isCancelable
              startTime
            }
          }"
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    
    $data = $response.data.activitySeries  
    
    $row = '' | Select-Object activityConnection,id,activitySeriesId,lastUpdated,lastActivityType,lastActivityStatus,objectId,objectName,objectType,severity,progress,isCancelable,startTime
    $activityConnection = @()
    foreach ($activity in $data.activityConnection.nodes) {
        $activityRow = '' | Select-Object activityInfo,message,status,time,severity
        $activityRow.activityInfo = $activity.activityInfo;
        $activityRow.message = $activity.message;
        $activityRow.status = $activity.status;
        $activityRow.time = $activity.time;
        $activityRow.severity = $activity.severity;
        $activityConnection += $activityRow
    }

    $row.activityConnection = $activityConnection
  
    $row.id = $data.id
    $row.activitySeriesId = $data.activitySeriesId
    $row.lastUpdated = $data.lastUpdated
    $row.lastActivityType = $data.lastActivityType
    $row.lastActivityStatus = $data.lastActivityStatus
    $row.objectId = $data.objectId
    $row.objectName = $data.objectName
    $row.objectType = $data.objectType
    $row.severity = $data.severity
    $row.progress = $data.progress
    $row.isCancelable = $data.isCancelable
    $row.startTime = $data.startTime

    return $row
}


function Get-PolarisO365EnterpriseApplication() {
    <#
    .SYNOPSIS

    Returns the Enterprise Applications configured on the Polaris Account.

    .DESCRIPTION

    Returns an array for each Enterprise Application configured on the Polaris
    account

    .PARAMETER Token
    Polaris access token, get this using the 'Get-PolarisTokenServiceAccount' or 'Get-PolarisToken' command.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'
.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisO365EnterpriseApplication.

    .OUTPUTS

    System.Object. Get-PolarisO365EnterpriseApplication returns an array 
    containing the appId, subscription, appType, addedAt, appOwner, and
    isAuthenticated for each Microsoft 365 Enterprise Application connected
    to Rubrik.

    .EXAMPLE

    PS> Get-PolarisO365EnterpriseApplication -Token $token -PolarisURL $url
    appId           : 72d1998d-15dc-4388-80de-8731e59aab89
    subscription    : Rubrik Demo
    appType         : TEAMS
    addedAt         : 8/17/2021 1:31:51 PM
    appOwner        : RUBRIK_SAAS
    isAuthenticated : True
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String]$Token,
        [Parameter(Mandatory = $True)]
        [String]$PolarisURL,
    )

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'



    $payload = @{
        "operationName" = "ListO365Apps";
        "variables"     = @{
            "first"  = 40;
            "o365AppFilters" = @();
        };
        "query"         = "query ListO365Apps(`$first: Int!, `$after: String, `$o365AppFilters: [AppFilter!]!, `$o365AppSortByParam: AppSortByParam) {
            listO365Apps(first: `$first, after: `$after, o365AppFilters: `$o365AppFilters, o365AppSortByParam: `$o365AppSortByParam) {
              edges {
                node {
                  appId
                  subscription
                  appType
                  addedAt
                  appOwner
                  isAuthenticated
                }
              }
              pageInfo {
                endCursor
                hasNextPage
                hasPreviousPage
                
              }
            }
          }"
    }

    $node_array = @()


    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    $node_array += $response.data.listO365Apps.edges

    # get all pages of results
    while ($response.data.listO365Apps.pageInfo.hasNextPage) {
        $payload.variables.after = $response.data.listO365Apps.pageInfo.endCursor
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $node_array += $response.data.listO365Apps.edges
    }

    $appArray = @()
    foreach ($app in $node_array) {
        $app = $app.node

        $row = '' | Select-Object appId,subscription,appType,addedAt,appOwner,isAuthenticated
        $row.appId = $app.appId;
        $row.subscription = $app.subscription;
        $row.appType = $app.appType;
        $row.addedAt = $app.addedAt;
        $row.appOwner = $app.appOwner;
        $row.isAuthenticated = $app.isAuthenticated;
        
        $appArray += $row
    }

    return $appArray 
}





