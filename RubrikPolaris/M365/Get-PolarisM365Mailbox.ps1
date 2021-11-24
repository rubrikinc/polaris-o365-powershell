function Get-PolarisM365Mailbox() {
    <#
    .SYNOPSIS

    Returns a filtered list of Microsoft 365 mailboxes for a given subscription in a given Rubrik account.

    .DESCRIPTION

    Returns a filtered list of Microsoft 365 mailboxes from a given subscription and Rubrik account, taking
    an subscription ID, and search string.


    .PARAMETER SubscriptionID
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the
    'Get-PolarisM365Subscriptions' command.

    .PARAMETER SearchString
    Search string, used to filter mailbox's name or user principal name.

    .INPUTS

    SubscriptionId. Value can be piped through Get-PolarisM365Subscriptions.

    .OUTPUTS

    System.Object. Get-PolarisM365Mailbox returns an array containing the ID, Name,
    email address, and SLA details for the returned O365 mailboxes.

    .EXAMPLE

    PS> Get-PolarisM365Mailbox -SubscriptionId $my_sub.id -SearchString 'Milan'

    name                   : Milan Kundera
    subscriptionid         : 12341234-1234-1234-abcd-123456789012
    emailAddress           : milan.kundera@mydomain.onmicrosoft.com
    slaAssignment          : Direct
    effectiveSlaDomainName : Gold

    PS> Get-PolarisM365Subscriptions | Get-PolarisM365Mailbox -SearchString "Milan"

    name                   : Milan Kundera
    subscriptionid         : 12341234-1234-1234-abcd-123456789012
    emailAddress           : milan.kundera@mydomain.onmicrosoft.com
    slaAssignment          : Direct
    effectiveSlaDomainName : Gold
    #>

    param(
        [Parameter(Mandatory = $True,ValueFromPipelineByPropertyName = $True)]
        [String]$SubscriptionId,
        [Parameter(Mandatory = $True)]
        [String]$SearchString,
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
Export-ModuleMember -Function Get-PolarisM365Mailbox

