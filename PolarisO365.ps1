function Get-PolarisToken() {
    <#
    .SYNOPSIS

    Returns an API access token for a given Polaris instance.

    .DESCRIPTION

    Returns an API access token for a given Polaris instance, taking the URL, username and password.

    .PARAMETER Username
    Polaris username.

    .PARAMETER Password
    Polaris password.

    .PARAMETER Password
    The URL for the Polaris instance in the form 'https://myurl'

    .INPUTS

    None. You cannot pipe objects to Get-PolarisToken.

    .OUTPUTS

    System.String. Get-PolarisToken returns a string containing the access token.

    .EXAMPLE

    PS> $token = Get-PolarisToken -Username $username -Password $password -PolarisURL $url
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Username,
        [Parameter(Mandatory=$True)]
        [String]$Password,
        [Parameter(Mandatory=$True)]
        [String]$PolarisURL
    )
    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
    }
    $payload = @{
        "username" = $Username;
        "password" = $Password;
    }
    $endpoint = $PolarisURL + '/api/session'
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    return $response.access_token
}

function Get-PolarisSLA() {
    <#
    .SYNOPSIS

    Returns the SLA Domains from a given Polaris instance.

    .DESCRIPTION

    Returns SLA Domains for a given Polaris instance. This can be used to return 
    based on a name query, by using the 'Name' parameter.

    .PARAMETER Token
    Polaris access token, get this using the 'Get-PolarisToken' command.

    .PARAMETER PolarisURL
    The URL for the Polaris instance in the form 'https://myurl'

    .PARAMETER Name
    Optional. The name of the required SLA Domain.

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
        [Parameter(Mandatory=$True)]
        [String]$Token,
        [Parameter(Mandatory=$True)]
        [String]$PolarisURL,
        [Parameter(Mandatory=$False)]
        [String]$Name
    )
    
    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }
    
    $endpoint = $PolarisURL + '/api/graphql'

    $payload = @{
        "operationName" = "SLAList";
        "variables" = @{"first" = 20; "name" = $Name};
        "query" = "query SLAList(`$after: String, `$first: Int, `$name: String) {
            globalSlaConnection(after: `$after, first: `$first, filter: [{field: NAME, text: `$name}]) {
                edges {
                    node {
                        id
                        name
                        description
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
        $row = '' | select name,id,description
        $row.name = $edge.node.name
        $row.id = $edge.node.id
        $row.description = $edge.node.description
        $sla_detail += $row
    }

    return $sla_detail
}

function Get-PolarisO365Subscriptions {
    <#
    .SYNOPSIS

    Returns all O365 subscriptions from a given Polaris instance.

    .DESCRIPTION

    Returns an array of Office 365 subscriptions from a given Polaris instance, taking 
    an API token, and Polaris URL.

    .PARAMETER Token
    Polaris API Token.

    .PARAMETER PolarisURL
    The URL for the Polaris instance in the form 'https://myurl'

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
        [Parameter(Mandatory=$True)]
        [String]$Token,
        [Parameter(Mandatory=$True)]
        [String]$PolarisURL
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }
    
    $endpoint = $PolarisURL + '/api/graphql'
    
    # Get a list of all the orgs
    
    $payload = @{
        "operationName" = "O365OrgList";
        "query" = "query O365OrgList(`$after: String, `$first: Int) {
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
        "variables" = @{
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
            "query" = "query o365OrgCard(`$id: UUID!) {
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
            "variables" = @{
                "id" = "$org_id";
            }
        }
    
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    
        $row = '' | select name,id,status,usersCount,unprotectedUsersCount,effectiveSlaDomainName,configuredSlaDomainName,effectiveSlaDomainId,configuredSlaDomainId
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

function Get-PolarisO365Users() {
    <#
    .SYNOPSIS

    Returns all O365 users for a given subscription in a given Polaris instance.

    .DESCRIPTION

    Returns an array of Office 365 users from a given subscription and Polaris instance, taking 
    an API token, Polaris URL, and subscription ID.

    .PARAMETER Token
    Polaris API Token.

    .PARAMETER PolarisURL
    The URL for the Polaris instance in the form 'https://myurl'

    .PARAMETER SubscriptionID
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the 
    'Get-PolarisO365Subscriptions' command.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisO365Users.

    .OUTPUTS

    System.Object. Get-PolarisO365Users returns an array containing the ID, Name, 
    email address, and SLA details for the returned O365 users.

    .EXAMPLE

    PS> Get-PolarisO365Users -Token $token -PolarisURL $url -SubscriptionId $my_sub.id

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
        "operationName" = "O365UserList";
        "query" = "query O365UserList(`$first: Int!, `$after: String, `$id: UUID!, `$filter: [Filter!]!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            o365Org(fid: `$id) {
                id
                childConnection(first: `$first, filter: `$filter, sortBy: `$sortBy, sortOrder: `$sortOrder, after: `$after) {
                    edges {
                        node {
                            id
                            name
                            emailAddress
                            effectiveSlaDomain {
                                name
                            }
                            authorizedOperations {
                                id
                                operations
                                __typename
                            }
                            childConnection(filter: []) {
                                nodes {
                                    id
                                    name
                                    objectType
                                }
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
            "id" = $SubscriptionId;
            "sortBy" = "EMAIL_ADDRESS";
            "sortOrder" = "ASC";
        }
    }
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    $node_array += $response.data.o365Org.childConnection.edges.node
    # get all pages of results
    while ($response.data.o365Org.childConnection.pageInfo.hasNextPage) {
        $payload.variables.after = $response.data.o365Org.childConnection.pageInfo.endCursor
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $node_array += $response.data.o365Org.childConnection.edges.node
    }

    $user_details = @()

    foreach ($node in $node_array) {
        $row = '' | select name,id,emailAddress,slaAssignment,effectiveSlaDomainName
        $row.name = $node.name
        $row.id = $node.id
        $row.emailAddress = $node.emailAddress
        $row.slaAssignment = $node.slaAssignment
        $row.effectiveSlaDomainName = $node.effectiveSlaDomain.name
        $user_details += $row
    }

    return $user_details
}

function Get-PolarisO365User() {
    <#
    .SYNOPSIS

    Returns a filtered list of O365 users for a given subscription in a given Polaris instance.

    .DESCRIPTION

    Returns a filtered list of Office 365 users from a given subscription and Polaris instance, taking 
    an API token, Polaris URL, subscription ID, and search string.

    .PARAMETER Token
    Polaris API Token.

    .PARAMETER PolarisURL
    The URL for the Polaris instance in the form 'https://myurl'

    .PARAMETER SubscriptionID
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the 
    'Get-PolarisO365Subscriptions' command.

    .PARAMETER SearchString
    Search string, used to filter user's name or email address.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisO365User.

    .OUTPUTS

    System.Object. Get-PolarisO365User returns an array containing the ID, Name, 
    email address, and SLA details for the returned O365 users.

    .EXAMPLE

    PS> Get-PolarisO365User -Token $token -PolarisURL $url -SubscriptionId $my_sub.id -SearchString 'Milan'

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
        "operationName" = "O365UserList";
        "variables" = @{
            "id" = $SubscriptionId;
            "first" = 100;
            "filter" = @(
                @{
                    "field" = "IS_RELIC";
                    "texts" = @("false");
                },
                @{
                    "field" = "NAME_OR_EMAIL_ADDRESS";
                    "texts" = @($SearchString);
                }
            );
            "sortBy" = "EMAIL_ADDRESS";
            "sortOrder" = "ASC";
        };
        "query" = "query O365UserList(`$first: Int!, `$after: String, `$id: UUID!, `$filter: [Filter!]!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            o365Org(fid: `$id) {
                id
                childConnection(first: `$first, filter: `$filter, sortBy: `$sortBy, sortOrder: `$sortOrder, after: `$after) {
                    edges {
                        node {
                            id
                            name
                            emailAddress
                            effectiveSlaDomain {
                                name
                            }
                            authorizedOperations {
                                id
                                operations
                            }
                            childConnection(filter: []) {
                                nodes {
                                    id
                                    name
                                    objectType
                                }
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
            }
        }"
    }
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    $node_array += $response.data.o365Org.childConnection.edges.node
    # get all pages of results
    while ($response.data.o365Org.childConnection.pageInfo.hasNextPage) {
        $payload.variables.after = $response.data.o365Org.childConnection.pageInfo.endCursor
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $node_array += $response.data.o365Org.childConnection.edges.node
    }

    $user_details = @()

    foreach ($node in $node_array) {
        $row = '' | select name,id,emailAddress,slaAssignment,effectiveSlaDomainName
        $row.name = $node.name
        $row.id = $node.id
        $row.emailAddress = $node.emailAddress
        $row.slaAssignment = $node.slaAssignment
        $row.effectiveSlaDomainName = $node.effectiveSlaDomain.name
        $user_details += $row
    }

    return $user_details
}

function Set-PolarisO365ObjectSla() {
    <#
    .SYNOPSIS

    Sets the SLA Domain for a selected Office 365 object (user or subscription).

    .DESCRIPTION

    Sets the protection for an O365 user or subscription in a given Polaris instance, taking 
    an API token, Polaris URL, object ID, and SLA ID.

    .PARAMETER Token
    Polaris API Token.

    .PARAMETER PolarisURL
    The URL for the Polaris instance in the form 'https://myurl'

    .PARAMETER ObjectID
    The object ID(s) for an O365 user or subscription. Can be obtained using 'Get-PolarisO365User', 
    'Get-PolarisO365Users', or 'Get-PolarisO365Subscriptions' commands. This can take an array of object IDs.

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

    PS> Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_user.id -SlaID $my_sla.id
    Success
    
    .EXAMPLE

    PS> Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_user.id -SlaID 'DONOTPROTECT'
    Success

    .EXAMPLE

    PS> Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_subscription.id -SlaID 'UNPROTECTED'
    Success
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Token,
        [Parameter(Mandatory=$True)]
        [String]$PolarisURL,
        [Parameter(Mandatory=$True)]
        [String[]]$ObjectID,
        [Parameter(Mandatory=$True)]
        [String]$SlaID
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $payload = @{
        "operationName" = "AssignSLA";
        "variables" = @{
            "globalSlaAssignType" = "protectWithSlaId";
            "globalSlaOptionalFid" = $SlaID;
            "objectIds" = $ObjectID;
        };
        "query" = "mutation AssignSLA(`$globalSlaOptionalFid: UUID, `$globalSlaAssignType: SlaAssignTypeEnum!, `$objectIds: [UUID!]!) {
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
    if ($response.data.assignSla.success = $true) {
        return 'Success'
    } else {
        throw 'Issue assigning SLA domain to object'
    }
}

