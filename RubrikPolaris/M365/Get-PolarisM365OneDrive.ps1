function Get-PolarisM365OneDrive() {
    <#
    .SYNOPSIS
    Returns configuration information for a single OneDrive user.

    .DESCRIPTION
    Returns an array of configuration information for a single OneDrive user.

    .PARAMETER Email
    The EMail address of the OneDrive user. 

    .PARAMETER SubscriptionID
    Optional ID of the Microsoft 365 subscription connect to Rubrik. Only required when more than one subscription is connected to Rubrik.

    .INPUTS
    None. You cannot pipe objects to Get-PolarisM365OneDrives.

    .OUTPUTS
    System.Object. Get-PolarisM365OneDrive returns an array containing the Name, ID,
    User Name, User Principal Name, SLA Assignment, and Effictive SLA Domain Name for a OneDrive user.

    .EXAMPLE
    PS> Get-PolarisM365OneDrive -Email $EmailAddress
    name                   : Drew Russell
    id                     : 643df529-9e90-4af8-b715-d8d2d181690e
    userName               : Drew Russell
    userPrincipalName      : drew.russell@mydomain.onmicrosoft.com
    slaAssignment          : Derived
    effectiveSlaDomainName : Gold
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Email,
        [String]$SubscriptionId,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    $m365Subscription = Get-PolarisM365Subscriptions

    if ($m365Subscription.Count -gt 1) {
        throw "Multiple Microsoft 365 subscriptions have been configured on Rubrik. Please specify the correct subscription through the SubscriptionId variable."
    }

    if ($SubscriptionId -ne $null) {
        $SubscriptionId = $m365Subscription.subscriptionId
    }

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

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
                @{
                    "field" = "NAME_OR_EMAIL_ADDRESS";
                    "texts" = @("$Email")
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
Export-ModuleMember -Function Get-PolarisM365OneDrive

