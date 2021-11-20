function Restore-PolarisM365OneDrive() {
    <#
    .SYNOPSIS
    Restore a Users entire OneDrive
    .DESCRIPTION
    Restore a users entire OneDrive to either it's original location or by created a Download link through Rubrik.
    .PARAMETER Token
    Polaris API Token.
    .PARAMETER PolarisURL
    The URL for the Polaris account in the form 'https://$PolarisAccount.my.rubrik.com'
    .PARAMETER OneDriveId
    The ID of the OneDrive you wish to restore
    .PARAMETER SnapshotId
    The ID of the snapshot you wish to restore.
    .PARAMETER SnapshotStorageLocation
    The ID of the snapshot storage location you wish to restore.
    .PARAMETER RecoveryOption
    The type of restore job you wish to use. Specify Original to restore to the Original OneDrive of Download to create a download link through Rubrik.
    .INPUTS
    None. You cannot pipe objects to Restore-PolarisM365OneDrive.
    .OUTPUTS
    String. The taskchainID of the Restore job which can be used to monitor the jobs progress.
    .EXAMPLE
    PS> Restore-PolarisM365OneDrive -PolarisURL $url -OneDriveId $user.id -SnapshotId $snapshotDetails.lastSnapshotId -SnapshotStorageLocation $snapshotDetails.lastSnapshotStorageLocation -RecoveryOption "Download"
    123594e0-1477-4be8-b6a2-f04174336a98
    #>

    param(
        [Parameter(Mandatory=$True)]
        [ValidateSet("Download", "Original")]
        [String]$RecoveryOption,
        [Parameter(Mandatory=$True)]
        [String]$OneDriveId,
        [Parameter(Mandatory=$True)]
        [String]$SnapshotStorageLocation,
        [Parameter(Mandatory=$True)]
        [String]$SnapshotId,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'
    
    if ($RecoveryOption -eq "Download") {
        $actionType = "EXPORT_SNAPPABLE"
        
    } else {
        $actionType = "RESTORE_SNAPPABLE"
    }


    $payload = @{
        "operationName" = "O365RestoreOnedriveMutation";
        "query" = "mutation O365RestoreOnedriveMutation(`$filesToRestore: [FileInfo!]!, `$foldersToRestore: [FolderInfo!]!, `$destOnedriveUUID: UUID!, `$sourceOnedriveUUID: UUID!, `$restoreFolderPath: String!, `$actionType: O365RestoreActionType!) {
            restoreO365Snappable(snappableType: ONEDRIVE, sourceSnappableUUID: `$sourceOnedriveUUID, destSnappableUUID: `$destOnedriveUUID, snappableRestoreConfig: {OneDriveRestoreConfig: {FilesToRestore: `$filesToRestore, FoldersToRestore: `$foldersToRestore, RestoreFolderPath: `$restoreFolderPath}}, actionType: `$actionType) {
              taskchainId
            }
          }";
        "variables" = @{
            "foldersToRestore" = @(
                @{
                    "FolderID" = "root";
                    "FolderName" = "OneDrive";
                    "SnapshotID" = $SnapshotId;
                    "FolderSize" = 0;
                    "SnapshotNum" = [int]$SnapshotStorageLocation;

                }
            );
            "filesToRestore" = @();
            "restoreFolderPath" = "";
            "sourceOnedriveUUID" = $OneDriveId;
            "destOnedriveUUID" = $OneDriveId;
            "actionType" = $actionType;
        }

    }
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers

    return $response.data.restoreO365Snappable.taskchainId
    
}
Export-ModuleMember -Function Restore-PolarisM365OneDrive

