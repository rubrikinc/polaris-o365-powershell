function Get-PolarisM365OneDriveSnapshot() {
    <#
    .SYNOPSIS
    Return the ID and Storage Location for the last OneDrive snapshot taken.

    .DESCRIPTION
    Returns an array that contains the Snapshot ID and Storage Location for the last snapshot taken on a OneDrive account. 
    This information can then be utilized in Restore-PolarisM365OneDrive.

    .PARAMETER OneDriveID
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the
    'Get-PolarisM365Subscriptions' command.

    .INPUTS
    None. You cannot pipe objects to Get-PolarisM365OneDriveSnapshot.

    .OUTPUTS
    System.Object. Get-PolarisM365OneDriveSnapshot returns an array containing lastSnapshotId and
    lastSnapshotStorageLocation.
    
    .EXAMPLE
    PS> Get-PolarisM365OneDriveSnapshot -OneDriveID $OneDriveID
    lastSnapshotId                       lastSnapshotStorageLocation
    --------------                       ---------------------------
    15e80edc-3211-412d-8cd2-1f5e33c52863                          46
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$OneDriveId,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
        
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $payload = @{
        "operationName" = "O365OnedriveList";
        "query" = "query O365OnedriveList(`$snappableId: UUID!) {
            o365Onedrive(snappableFid: `$snappableId) {
              newestSnapshot {
                id
              }
              snapshotConnection(first: 1, sortOrder: Desc) {
                nodes {
                  sequenceNumber
                }
              }
            }
          }";
        "variables" = @{
            "snappableId" = $OneDriveId;
        }

    }

   
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    
    $row = '' | Select-Object lastSnapshotId,lastSnapshotStorageLocation
    $row.lastSnapshotId = $response.data.o365Onedrive.newestSnapshot.id
    $row.lastSnapshotStorageLocation = $response.data.o365Onedrive.snapshotConnection.nodes.sequenceNumber
    
    return $row
}
Export-ModuleMember -Function Get-PolarisM365OneDriveSnapshot

