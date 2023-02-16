function Restore-PolarisM365Exchange() {
    <#
    .SYNOPSIS
    Restore a Users entire Exchange
    
    .DESCRIPTION
    Restore a Users entire Exchange based on the latest backup.

    .PARAMETER Emails
    The Email address of the Exchange user you wish to restore. 

    .PARAMETER ExchangeId
    The ID of the Exchange you wish to restore. This value is only needed if the automatic ID lookup by Email fails.
    
    .PARAMETER RecoveryOption
    The type of restore job you wish to use. Original is only supported option.
   
    .INPUTS
    None. You cannot pipe objects to Restore-PolarisM365Exchange.
   
    .OUTPUTS
    String. The taskchainID of the Restore job which can be used to monitor the jobs progress.
   
    .EXAMPLE
    PS> Restore-PolarisM365Exchange -Email $emailAddress
    #>

    param(
        [Parameter(Mandatory=$True)]
        [Array]$Emails,
        [Parameter(Mandatory=$False)]
        [ValidateSet("Original")] # Keep code in place for future use. 
        [String]$RecoveryOption = "Original",
        [Parameter(Mandatory=$False)]
        [Array]$ExchangeIds,
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

    
    #TODO: Prevent both $Emails and $ExchangeIds from being passed.
    Write-Information -Message "Starting the restoration process for $($Emails.Count) Exchange account(s)."
    Write-Information -Message "Determing the Rubrik ID for each Exchange users."
    if ($ExchangeIds.Count -eq 0) {
        $ExchangeIds = @()
        foreach ($e in $Emails) {
            $ExchangeUser = Get-PolarisM365Exchange -Email $e
            #TODO: Create a warning instead of Write-Warning -Messageing an error
            if ($null -eq $ExchangeUser) {
                Write-Warning -Message "The specified $e user was not found. Please check the email address or manually specify the ExchangeId variable and try again."
            } else {
                $ExchangeIds += $ExchangeUser.id

                # Warm the Search container on the backend to improve API performance below
                $warmPayload = @{
                    "operationName" = "WarmO365ObjectSearchCacheMutation";
                    "query" = "mutation WarmO365ObjectSearchCacheMutation(`$input: WarmSearchCacheInput!) {
                        warmSearchCache(input: `$input)
                                
                    }";
                    "variables" = @{
                        "input" = @{
                            "workloadFid" = $ExchangeUser.id;
                        }
                      
                    }
        
                }
        
                Invoke-RestMethod -Method POST -Uri $endpoint -Body $($warmPayload | ConvertTo-JSON -Depth 100) -Headers $headers | Out-Null

            }

    
        }
           
    }

 
    #TODO: Conslidate these into a single array
    $rootFolders = @()
    $snapshots = @()
    $exchangeRestoreToSkip = @()
    Write-Information -Message "Determining the latest snapshot for each Exchange user."
    foreach ($eId in $ExchangeIds) {
        $skipRestore = $false

        $snapshot = Get-PolarisM365ExchangeSnapshot -ExchangeID $eId
        $snapshots += $snapshot
 
        if ($null -eq $snapshot.lastSnapshotId) {
            Write-Warning -Message "Skipping Restore: The specified Exchange mailbox does not have any snapshots to restore from."
            $skipRestore = $true
        } 

        if ($snapshot.isIndexed -eq $false) {
            Write-Warning -Message "Skipping Restore: The specified Exchange mailbox is not indexed. Please wait for the indexing to complete and try again."
            $skipRestore = $true
        }

        if ($skipRestore -eq $false) {
            $rootFolderPayload = @{
                "operationName" = "folderQuery";
                "query" = "query folderQuery(`$snapshotFid: UUID!, `$folderId: String!, `$snappableFid: UUID!, `$orgId: UUID!) {
                    browseFolder(snapshotFid: `$snapshotFid, folderId: `$folderId, snappableFid: `$snappableFid, orgId: `$orgId) {
                        edges {
                          node {
                            id
                          }
                        }
                      }}";
                "variables" = @{
                    "snappableFid" = $eId;
                    "snapshotFid" = $snapshot.lastSnapshotId;
                    "folderId" = "root";
                    "orgId" = $ExchangeUser.subscriptionId;
                }
        
            }
    
        
            $rootFolder =  Invoke-RestMethod -Method POST -Uri $endpoint -Body $($rootFolderPayload | ConvertTo-JSON -Depth 100) -Headers $headers
            $rootFolders += $rootFolder
        } else {
            $exchangeRestoreToSkip += $eId
        }
    
       

    }

    #TODO: Add error handing if any foreach loop fails (i.e prevent a mismatch of IDs) 
    Write-Information -Message "Starting the restore process for each Exchange user."
    $taskchainIds = @()
    foreach ($eId in $exchangeIds){

        if ($exchangeRestoreToSkip.Contains($eId) -eq $false) {
            
            $index = [array][array]::IndexOf($exchangeIds, $eId)

            #TODO: prevent cannot index into a null array.
            $payload = @{
                "operationName" = "O365RestoreMailboxMutation";
                "query" = "mutation O365RestoreMailboxMutation(`$orgId: UUID, `$mailboxId: UUID!, `$restoreConfigs: [RestoreObjectConfig!]!, `$actionType: O365RestoreActionType!) {
                    restoreO365Mailbox(restoreConfig: {mailboxUuid: `$mailboxId, restoreConfigs: `$restoreConfigs, orgUuid: `$orgId, actionType: `$actionType}) {
                      taskchainId
                    }
                  }
                  ";
                "variables" = @{
                    "mailboxId" = $eID;
                    "orgId" = $ExchangeUser.subscriptionId;
                    "restoreConfigs" = @(
                        @{
                            "SnapshotUUID" = $snapshots[$index].lastSnapshotId;
                            "FolderID" = $rootFolders[$index].data.browseFolder.edges[0].node.id;
        
                        }
                    );
                    "actionType" = "RESTORE_SNAPPABLE";
                  
                }
        
            }

           
            $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
            $taskchainIds += $response.data.restoreO365Mailbox.taskchainId
        }


    
    }
  
 
    return $taskchainIds
    
}
Export-ModuleMember -Function Restore-PolarisM365Exchange

