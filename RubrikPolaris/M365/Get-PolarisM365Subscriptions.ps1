function Get-PolarisM365Subscriptions {
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

    None. You cannot pipe objects to Get-PolarisM365Subscriptions.

    .OUTPUTS

    System.Object. Get-PolarisM365Subscriptions returns an array containing the ID, Name,
    status, count of users, count of unprotected users, and SLA details for the
    returned O365 Subscriptions.

    .EXAMPLE

    PS> Get-PolarisM365Subscriptions -Token $token -PolarisURL $url

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
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
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
Export-ModuleMember -Function Get-PolarisM365Subscriptions
