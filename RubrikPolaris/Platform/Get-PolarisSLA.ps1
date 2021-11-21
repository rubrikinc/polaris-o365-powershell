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
        [Parameter(Mandatory = $False)]
        [String]$Name,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
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
Export-ModuleMember -Function Get-PolarisSLA

