function Get-PolarisRole() {
    <#
    .SYNOPSIS

    Returns the details for the configured Role(s) in Polaris

    .DESCRIPTION

    Returns the Name, Description, and ID, and configured Privileges  for Polaris Role(s).

    .PARAMETER Token
    Polaris access token, get this using the 'Connect-Polaris 'command.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .PARAMETER Name
    An optional parameter to filter the results by name.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisRole.

    .OUTPUTS

    System.Object. Get-PolarisRole returns an array containing the Name, Description,
    and ID returned Role(s).

    .EXAMPLE

    PS> Get-PolarisRole 
    Name                        Description                                        ID                                    Privileges
    ----                        -----------                                        --                                    --------
    Engineering                 Engineering Users                                  f02e5e46-3693-4e94-82fc-edcb50bf694e  {ViewGcpCloudAccount, ViewReport}
    Engineering - SQL                                                              d0a55aa2-1023-4c9f-9e1d-844601a2ef37  {ViewGcpCloudAccount, ViewReport}
    Microsoft 365 - Engineering Local account access required for Engineering team 95ba8e14-347e-4597-ac42-608f5384f309  {ViewGcpCloudAccount, ViewReport}

    PS> Get-PolarisRole -Name "Microsoft 365 - Engineering"
    Name                        Description                                        ID                                   Privileges
    ----                        -----------                                        --                                   ----------
    Microsoft 365 - Engineering Local account access required for Engineering team 95ba8e14-347e-4597-ac42-608f5384f309 {ViewGcpCloudAccount, ViewReport}
    #>
    #>

    param(
        [Parameter(Mandatory = $False)]
        [String]$Name = "",
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
        "operationName" = "RolesQuery";
        "variables"     = @{
            "nameSearch"  = $Name;
            "sortBy"      = "Name";
            "sortOrder"  = "Asc";
        };
        "query"         = "query RolesQuery(`$sortBy: RoleFieldEnum, `$sortOrder: SortOrderEnum, `$nameSearch: String) {
          getAllRolesInOrgConnection(sortBy: `$sortBy, sortOrder: `$sortOrder, nameFilter: `$nameSearch) {
            nodes {
              id
              name
              description
              permissions {
                operation
              }
            }
          }
        }"
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    
    $data = $response.data.getAllRolesInOrgConnection.nodes 

    $role_details = @()

    foreach ($role in $data) {
        $prvileges = @()
        $row = '' | Select-Object Name,Description,ID,Privileges
        $row.name = $role.name;
        $row.description = $role.description;
        $row.id = $role.id;
        foreach ($permission in $role.permissions) {
            $prvileges += $permission.operation
        }
        $row.privileges = $prvileges

        $role_details += $row

    } 
  

    return $role_details
}
Export-ModuleMember -Function Get-PolarisRole


