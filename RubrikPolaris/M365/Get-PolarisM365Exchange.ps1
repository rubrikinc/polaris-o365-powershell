function Get-PolarisM365Exchange() {
    <#
    .SYNOPSIS
    Returns all details of a Exchange user for a given subscription in a given Polaris account.

    .DESCRIPTION
    Returns an array of details for a Microsoft Exchange users from a given subscription.


    .PARAMETER Email
    The email address of the user you wish to look up details for. 

    .PARAMETER SubscriptionId
    The Rubrik Subscription ID for the Microsoft 365 subscription that the user belongs to. 
    This field is only required if multiple subscriptions are connected to Rubrik.
    
    .INPUTS
    None. You cannot pipe objects to Get-PolarisM365Exchange.

    .OUTPUTS
    System.Object. Get-PolarisM365Exchange returns an array containing the name, id,
    userPrincipalName, slaAssignment, effectiveSlaDomainName and subscriptionId the specifified Exchange user.

    .EXAMPLE
    PS> Get-PolarisM365Exchange -Email "demo.user@rubrikdemo.com"
    name                   : Demo User
    id                     : e4fe6903-59e2-3983-b293-31510f9ff2735
    userPrincipalName      : demo.user@rubrikdemo.com
    slaAssignment          : Derived
    effectiveSlaDomainName : Bronze SLA
    subscriptionId         : 5ef7d09f-e393-422d-a06c-9b3fc9c45d27
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Email,
        [Parameter(Mandatory=$False)]
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
        "operationName" = "O365MailboxListQuery";
        "query" = "query O365MailboxListQuery(`$first: Int!, `$after: String, `$orgId: UUID!, `$filter: [Filter!]!, `$sortBy: HierarchySortByField, `$sortOrder: SortOrder) {
            o365Mailboxes(o365OrgId: `$orgId, after: `$after, first: `$first, filter: `$filter, sortBy: `$sortBy, sortOrder: `$sortOrder) {
                edges {
                    node {
                        id
                        name
                        userPrincipalName
                        effectiveSlaDomain {
                            id
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
            "sortBy" = "NAME";
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

    $user_details = @()

    foreach ($node in $node_array) {
        $row = '' | Select-Object name,id,userPrincipalName,slaAssignment,effectiveSlaDomainName, subscriptionId
        $row.name = $node.node.name
        $row.id = $node.node.id
        $row.userPrincipalName = $node.node.userPrincipalName
        $row.slaAssignment = $node.node.slaAssignment
        $row.effectiveSlaDomainName = $node.node.effectiveSlaDomain.name
        $row.subscriptionId = $SubscriptionId
        $user_details += $row
    }

    return $user_details
}
Export-ModuleMember -Function Get-PolarisM365Exchange

