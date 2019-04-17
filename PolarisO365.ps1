function Get-PolarisToken() {
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

function Get-PolarisO365Subscriptions {
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

