function Restore-PolarisM365OneDrive() {
    <#
    .SYNOPSIS
    Restore a Users entire OneDrive
    
    .DESCRIPTION
    Restore a Users entire OneDrive based on the latest backup.

    .PARAMETER Email
    The Email address of the OneDrive user you wish to restore. 

    .PARAMETER OneDriveId
    The ID of the OneDrive you wish to restore. This value is only needed if the automatic ID lookup by Email fails.
    
    .PARAMETER RecoveryOption
    The type of restore job you wish to use. Original is only supported option.
   
    .INPUTS
    None. You cannot pipe objects to Restore-PolarisM365OneDrive.
   
    .OUTPUTS
    String. The taskchainID of the Restore job which can be used to monitor the jobs progress.
   
    .EXAMPLE
    PS> Restore-PolarisM365OneDrive -Email $emailAddress
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Email,
        [Parameter(Mandatory=$False)]
        [ValidateSet("Original")] # Keep code in place for future use. 
        [String]$RecoveryOption = "Original",
        [Parameter(Mandatory=$False)]
        [String]$OneDriveId,
        [Parameter(Mandatory=$False)]
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    if ($OneDriveId -ne $null) {
        $oneDriveUser = Get-PolarisM365OneDrive -Email $Email 
        if ($null -eq $oneDriveUser) {
            throw "The specified OneDrive user was not found. Please check the email address or manually specify the OneDriveId variable and try again."
        }

        $OneDriveId = $oneDriveUser.id
        
    }

    $snapshot = Get-PolarisM365OneDriveSnapshot -OneDriveID $OneDriveID
    if ($null -eq $snapshot) {
        throw "The specified OneDrive does not have any snapshots to restore from."
    }     

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'
    
    # Keep code in place for future use. 
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
                    "SnapshotID" = $snapshot.lastSnapshotId;
                    "FolderSize" = 0;
                    "SnapshotNum" = [int]$snapshot.lastSnapshotStorageLocation;

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

