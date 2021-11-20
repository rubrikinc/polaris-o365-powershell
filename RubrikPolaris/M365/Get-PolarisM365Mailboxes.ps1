function Get-PolarisM365Mailboxes() {
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
    'Get-PolarisM365Subscriptions' command.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisM365Mailboxes.

    .OUTPUTS

    System.Object. Get-PolarisM365Mailboxes returns an array containing the ID, Name,
    email address, and SLA details for the returned O365 mailboxes.

    .EXAMPLE

    PS> Get-PolarisM365Mailboxes -Token $token -PolarisURL $url -SubscriptionId $my_sub.id

    name                        : Milan Kundera
    id                          : 12341234-1234-1234-abcd-123456789012
    userPrincipalName           : milan.kundera@mydomain.onmicrosoft.com
    slaAssignment               : Direct
    effectiveSlaDomainName      : Gold
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String]$SubscriptionId,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
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
Export-ModuleMember -Function Get-PolarisM365Mailboxes



