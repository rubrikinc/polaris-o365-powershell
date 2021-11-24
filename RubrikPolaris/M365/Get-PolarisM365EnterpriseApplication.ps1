function Get-PolarisM365EnterpriseApplication() {
    <#
    .SYNOPSIS

    Returns the Enterprise Applications configured on the Polaris Account.

    .DESCRIPTION

    Returns an array for each Enterprise Application configured on the Polaris
    account

    .INPUTS

    None. You cannot pipe objects to Get-PolarisM365EnterpriseApplication.

    .OUTPUTS

    System.Object. Get-PolarisM365EnterpriseApplication returns an array 
    containing the appId, subscription, appType, addedAt, appOwner, and
    isAuthenticated for each Microsoft 365 Enterprise Application connected
    to Rubrik.

    .EXAMPLE

    PS> Get-PolarisM365EnterpriseApplication
    appId           : 72d1998d-15dc-4388-80de-8731e59aab89
    subscription    : Rubrik Demo
    appType         : TEAMS
    addedAt         : 8/17/2021 1:31:51 PM
    appOwner        : RUBRIK_SAAS
    isAuthenticated : True
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
Export-ModuleMember -Function Get-PolarisM365EnterpriseApplication
