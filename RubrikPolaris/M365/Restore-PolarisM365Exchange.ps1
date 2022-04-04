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
        [String]$Email,
        [Parameter(Mandatory=$False)]
        [ValidateSet("Original")] # Keep code in place for future use. 
        [String]$RecoveryOption = "Original",
        [Parameter(Mandatory=$False)]
        [String]$ExchangeId,
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


    if ($ExchangeId -ne $null) {
        $ExchangeUser = Get-PolarisM365Exchange -Email $Email 
        if ($null -eq $ExchangeUser) {
            throw "The specified Exchange user was not found. Please check the email address or manually specify the ExchangeId variable and try again."
        }

        $ExchangeId = $ExchangeUser.id
        
    }

    # Warm the Search container on the backend to improve API performance below
    $warmPayload = @{
        "operationName" = "WarmO365ObjectSearchCacheMutation";
        "query" = "mutation WarmO365ObjectSearchCacheMutation(`$snappableId: UUID!) {
                    warmSearchCache(snappableFid: `$snappableId)
          }";
        "variables" = @{
            "snappableId" = $ExchangeId;
        }

    }
    Invoke-RestMethod -Method POST -Uri $endpoint -Body $($warmPayload | ConvertTo-JSON -Depth 100) -Headers $headers | Out-Null
    
    $snapshot = Get-PolarisM365ExchangeSnapshot -ExchangeID $ExchangeID
    if ($null -eq $snapshot) {
        throw "The specified Exchange does not have any snapshots to restore from."
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
            "snappableFid" = $ExchangeId;
            "snapshotFid" = $snapshot.lastSnapshotId;
            "folderId" = "root";
            "orgId" = $ExchangeUser.subscriptionId;
        }

    }

    $rootFolder =  Invoke-RestMethod -Method POST -Uri $endpoint -Body $($rootFolderPayload | ConvertTo-JSON -Depth 100) -Headers $headers
  
  
    $payload = @{
        "operationName" = "O365RestoreMailboxMutation";
        "query" = "mutation O365RestoreMailboxMutation(`$orgId: UUID, `$mailboxId: UUID!, `$restoreConfigs: [RestoreObjectConfig!]!) {
            restoreO365Mailbox(restoreConfig: {mailboxUUID: `$mailboxId, restoreConfigs: `$restoreConfigs, orgUuid: `$orgId}) {
              taskchainId
            }
          }
          ";
        "variables" = @{
            "mailboxId" = $ExchangeId;
            "orgId" = $ExchangeUser.subscriptionId;
            "restoreConfigs" = @(
                @{
                    "SnapshotUUID" = $snapshot.lastSnapshotId;
                    "FolderID" = $rootFolder.data.browseFolder.edges[0].node.id;

                }
            );
          
        }

    }

    Write-Output $payload.variables | ConvertTo-JSON -Depth 100

    throw "asdf"
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers

    return $response.data.restoreO365Mailbox.taskchainId
    
}
Export-ModuleMember -Function Restore-PolarisM365Exchange

