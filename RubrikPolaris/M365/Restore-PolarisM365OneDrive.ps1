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
        [Array]$Emails,
        [Parameter(Mandatory=$False)]
        [ValidateSet("Original")] # Keep code in place for future use. 
        [String]$RecoveryOption = "Original",
        [Parameter(Mandatory=$False)]
        [Array]$OneDriveIds,
        [Parameter(Mandatory=$False)]
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

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

    #TODO: Prevent both $Emails and $OneDriveIds from being passed in at the same time.
    Write-Information -Message "Starting the restoration process for $($Emails.Count) OneDrive account(s)."
    Write-Information -Message "Determing the Rubrik ID for each OneDrive users."

    if ($OneDriveIds.Count -eq 0) {
        $OneDriveIds = @()
        foreach ($e in $Emails) {
            $oneDriveUser = Get-PolarisM365OneDrive -Email $e 
            if ($null -eq $oneDriveUser) {
                Write-Warning -Message "The specified OneDrive user was not found. Please check the email address or manually specify the OneDriveId variable and try again."
            }

            $OneDriveIds += $oneDriveUser.id

            # Warm the Search container on the backend to improve API performance below
            $warmPayload = @{
                "operationName" = "WarmO365ObjectSearchCacheMutation";
                "query" = "mutation WarmO365ObjectSearchCacheMutation(`$input: WarmSearchCacheInput!) {
                    warmSearchCache(input: `$input)
                            
                }";
                "variables" = @{
                    "input" = @{
                        "workloadFid" = $oneDriveUser.id;
                    }
                  
                }
           
            }
    
            Invoke-RestMethod -Method POST -Uri $endpoint -Body $($warmPayload | ConvertTo-JSON -Depth 100) -Headers $headers | Out-Null

        }
    }

    #TODO: Conslidate these into a single array
    Write-Information -Message "Starting the restore process for each OneDrive user."
    $taskchainIds = @()
    foreach ($oId in $OneDriveIds) {

        $snapshot = Get-PolarisM365OneDriveSnapshot -OneDriveID $oId

        if ($null -eq $snapshot.lastSnapshotId) {
            Write-Warning -Message "The specified OneDrive does not have any snapshots to restore from."
        } 

        if ($snapshot.isIndexed -eq $false) {
            Write-Warning -Message "The specified OneDrive is not indexed. Please wait for the indexing to complete and try again."
        }

        $payload = @{
            "operationName" = "O365RestoreOnedriveMutation";

            "query" = "mutation O365RestoreOnedriveMutation(`$input: RestoreO365SnappableInput!){
                restoreO365Snappable(input: `$input) {
                    taskchainId
            }
            }";
            "variables" = @{
                "input" =@{
                    "snappableType" = "ONEDRIVE"; 
                    "restoreConfig" = @{
                        "OneDriveRestoreConfig" = @{
                            "filesToRestore" = @(); 
                            "foldersToRestore" = @{
                                "folderId" = "root"; 
                                "folderName" = "OneDrive"; 
                                "folderSize" = 0; 
                                "snapshotId" = $snapshot.lastSnapshotId;
                                "snapshotNum" = [int]$snapshot.lastSnapshotStorageLocation;
    
                            }; 
                            "restoreFolderPath" = "";
                        }
                    }; 
                    "sourceSnappableUuid" = $oId
                    "destinationSnappableUuid" = $oId 
                    "actionType" = "RESTORE_SNAPPABLE"
                }
            }
        
        }

        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    
        $taskchainIds += $response.data.restoreO365Snappable.taskchainId
       

    }

    return $taskchainIds

    
    
}
Export-ModuleMember -Function Restore-PolarisM365OneDrive
