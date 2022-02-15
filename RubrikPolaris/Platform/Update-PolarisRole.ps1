function Update-RubrikRole() {
    <#
    .SYNOPSIS

    Returns the details for a given Polaris job.

    .DESCRIPTION

    Returns detailed information on a Polaris job. 

    .PARAMETER Token
    Polaris access token, get this using the 'Get-PolarisTokenServiceAccount' or 'Get-PolarisToken' command.

    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'

    .PARAMETER JobId
    The ID of the job you wish to get details on.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisJob.

    .OUTPUTS

    System.Object. Get-PolarisJob returns an array containing the activityConnection, id,
    lastUpdated,  lastActivityType, lastActivityStatus, objectId, objectName, severity, 
    objectType, progress, isCancelable, and startTime of the returned job.

    .EXAMPLE

    PS> Get-PolarisJob -PolarisURL $url -Token $token -JobId $JobId
    activityConnection : {@{activityInfo=; message=Started export of Drew Russell Microsoft 365 onedrive data; status=TaskSuccess; time=10/5/2021 3:45:19 PM; severity=Info}}
    id                 : 7592590
    activitySeriesId   : 1d3594e0-1488-4ba4-b6a2-f04174336a73
    lastUpdated        : 10/5/2021 3:45:19 PM
    lastActivityType   : Recovery
    lastActivityStatus : TaskSuccess
    objectId           : e4092af5-009b-47e9-9412-bcd57924cfa6
    objectName         : Drew Russell
    objectType         : O365Onedrive
    severity           : Info
    progress           : 100%
    isCancelable       : False
    startTime          : 10/5/2021 3:45:19 PM
    #>

    param(
        [Parameter(Mandatory = $True)]
        [System.Collections.ArrayList]$Privileges,
        [System.Collections.ArrayList]$ObjectId,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
        
    )

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $permission = @()
    for (privilege in $privileges){
      $permission.Add(
        @{
          "operation" = privilege;
          objectsForHierarchyTypes = @(
            @{
              "objectIds" = $ObjectId;
              "snappableType" = "AllSubHierarchyType"
            }
          )
          "objectIds" = $objectId;
        }
      )
    }

    $payload = @{
        "operationName" = "UpdateRole";
        "variables"     = @{
            "roleId"  = $RoleId;
            "name" = $name;
            "description" = $description;
            "permissions" = $permissions;
        };
        "query"         = "mutation UpdateRole(`$roleId: String!, `$name: String!, `$description: String!, `$permissions: [PermissionInput!]!) {}
          updateRole(roleId: `$roleId, name: `$name, description: `$description, permissions: `$permissions)
          }"
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    
    $data = $response.data.activitySeries  
    
    $row = '' | Select-Object activityConnection,id,activitySeriesId,lastUpdated,lastActivityType,lastActivityStatus,objectId,objectName,objectType,severity,progress,isCancelable,startTime
    $activityConnection = @()
    foreach ($activity in $data.activityConnection.nodes) {
        $activityRow = '' | Select-Object activityInfo,message,status,time,severity
        $activityRow.activityInfo = $activity.activityInfo;
        $activityRow.message = $activity.message;
        $activityRow.status = $activity.status;
        $activityRow.time = $activity.time;
        $activityRow.severity = $activity.severity;
        $activityConnection += $activityRow
    }

    $row.activityConnection = $activityConnection
  
    $row.id = $data.id
    $row.activitySeriesId = $data.activitySeriesId
    $row.lastUpdated = $data.lastUpdated
    $row.lastActivityType = $data.lastActivityType
    $row.lastActivityStatus = $data.lastActivityStatus
    $row.objectId = $data.objectId
    $row.objectName = $data.objectName
    $row.objectType = $data.objectType
    $row.severity = $data.severity
    $row.progress = $data.progress
    $row.isCancelable = $data.isCancelable
    $row.startTime = $data.startTime

    return $row
}
Export-ModuleMember -Function Update-RubrikRole


