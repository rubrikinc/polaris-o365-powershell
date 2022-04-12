function Restore-PolarisM365Exchange() {
    <#
    .SYNOPSIS
    Restore a Users entire Exchange
    
    .DESCRIPTION
    Restore a Users entire Exchange based on the latest backup.

    .PARAMETER Email
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
    

    if ($ExchangeIds.Count -eq 0) {
        $ExchangeIds = @()
        foreach ($e in $Emails) {
            $ExchangeUser = Get-PolarisM365Exchange -Email $e
            #TODO - Create a warning instead of throwing an error
            if ($null -eq $ExchangeUser) {
                throw "The specified Exchange user was not found. Please check the email address or manually specify the ExchangeId variable and try again."
            }

            $ExchangeIds += $ExchangeUser.id

        }
        
        
    }

    foreach ($eId in $ExchangeIds) {
        # Warm the Search container on the backend to improve API performance below
        $warmPayload = @{
            "operationName" = "WarmO365ObjectSearchCacheMutation";
            "query" = "mutation WarmO365ObjectSearchCacheMutation(`$snappableId: UUID!) {
                        warmSearchCache(snappableFid: `$snappableId)
            }";
            "variables" = @{
                "snappableId" = $eId;
            }

        }

        Invoke-RestMethod -Method POST -Uri $endpoint -Body $($warmPayload | ConvertTo-JSON -Depth 100) -Headers $headers | Out-Null

    }
    #TODO - Conslidat these into a single array
    $rootFolders = @()
    $snapshots = @()
    foreach ($eId in $ExchangeIds) {

        $snapshot = Get-PolarisM365ExchangeSnapshot -ExchangeID $eId
        $snapshots += $snapshot
 
        if ($null -eq $snapshot.lastSnapshotId) {
            #TODO - Create a warning instead of throwing an error
            throw "The specified Exchange mailbox does not have any snapshots to restore from."
        } 

        if ($snapshot.isIndexed -eq $false) {
            #TODO - Create a warning instead of throwing an error
            throw "The specified Exchange mailbox is not indexed. Please wait for the indexing to complete and try again."
        }
    
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

    }

    #TODO - Add error handing if any foreach loop fails (i.e prevent a mismatch of IDs)  
    $taskchainIds = @()
    foreach ($eId in $exchangeIds){
        $index = [array][array]::IndexOf($exchangeIds, $eId)


        # cannot index into a null array.
        $payload = @{
            "operationName" = "O365RestoreMailboxMutation";
            "query" = "mutation O365RestoreMailboxMutation(`$orgId: UUID, `$mailboxId: UUID!, `$restoreConfigs: [RestoreObjectConfig!]!) {
                restoreO365Mailbox(restoreConfig: {mailboxUUID: `$mailboxId, restoreConfigs: `$restoreConfigs, orgUuid: `$orgId}) {
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
              
            }
    
        }
       
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $taskchainIds += $response.data.restoreO365Mailbox.taskchainId

    
    }
  
 
    return $taskchainIds
    
}
Export-ModuleMember -Function Restore-PolarisM365Exchange

