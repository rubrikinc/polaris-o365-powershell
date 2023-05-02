function Start-MassRecovery() {
    <#
    .SYNOPSIS
    Mass restore OneDrives for an AD Group 
    
    .DESCRIPTION
    Mass restore entire OneDrives for an AD Group from latest backups before a
    given recovery point in time.

    .PARAMETER Name
    The name of the mass recovery you wish to choose. 

    .PARAMETER RecoveryPoint
    The date time you wish to use to restore closest earlier backups. The format
    is "MM/DD/YY HH:MM:SS"
    
    .PARAMETER SubscriptionName
    The subscription name you wish to mass restore under.
   
    .PARAMETER AdGroupId
    The ID of the AD Group you wish to mass restore.

    .PARAMETER WorkloadType
    The type of workload you wish to mass restore, only "OneDrive" is supported
    right now.

    .INPUTS
    None. You cannot pipe objects to Start-MassRecovery.
   
    .OUTPUTS
    System.Object. The taskchainID, massRecoveryInstanceID, error and jobID for 
    the mass recovery job.
   
    .EXAMPLE
    PS> Start-MassRecovery -Name $name -RecoveryPoint $recoveryPoint -SubscriptionName $subscriptionName -AdGroupId $adGroupId
        -WorkloadType $workloadType
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Name,
        [Parameter(Mandatory=$True)]
        [DateTime]$RecoveryPoint,
        [Parameter(Mandatory=$True)]
        [String]$SubscriptionName,
        [Parameter(Mandatory=$True)]
        [String]$AdGroupId,
        [Parameter(Mandatory=$True)]
        [String]$WorkloadType,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    if (($WorkloadType -ne "OneDrive") -and ($WorkloadType -ne "Exchange")) {
        Write-Host "Error starting mass recovery $Name. The error response is 'Only WorkloadType as OneDrive or Exchange is supported'."
        return
    }

    $snappableToSubSnappableMap = @{
        "OneDrive" = @(
            @{
                SnappableType = "O365_ONEDRIVE";
                SubSnappableType = "NONE";
                NameSuffix="_OneDrive";
            }
        );
        "Exchange" = @(
            @{
                SnappableType = "O365_EXCHANGE";
                SubSnappableType = "O365_MAILBOX";
                NameSuffix="_Mailbox";
            };
            @{
                SnappableType = "O365_EXCHANGE";
                SubSnappableType = "O365_CALENDAR";
                NameSuffix="_Calendar";
            };
            @{
                SnappableType = "O365_EXCHANGE";
                SubSnappableType = "O365_CONTACT";
                NameSuffix="_Contact";
            };
        );
    }

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'
    $rpMilliseconds = ([DateTimeOffset]$RecoveryPoint).ToUnixTimeMilliseconds()

    Write-Information -Message "Starting the mass restoration process for $WorkloadType account(s) under AD Group ID $AdGroupId."
    $snappableToSubSnappableMap[$WorkloadType] | ForEach-Object -Process {
        $recoveryName=$Name+$_.NameSuffix
        $payload = @{
            "operationName" = "StartBulkRecovery";
            "variables"     = @{
                "input" = @{
                    "definition" = @{
                        "name" = $recoveryName;
                        "adGroupSelectorWithRecoverySpec" =  @{
                            "baseInfo" = @{
                                "snappableType" = $_.SnappableType;
                                "subSnappableType" = $_.SubSnappableType;
                                "recoverySpec" = @{
                                    "recoveryPoint" = $rpMilliseconds;
                                    "srcSubscriptionName" = $SubscriptionName;
                                    "targetSubscriptionName" = $SubscriptionName;
                                }
                            };
                            "adGroupId"= $AdGroupId;
                        };
                        "recoveryMode" = "AD_HOC";
                        "failureAction" = "IGNORE_AND_CONTINUE";
                        "recoveryDomain" = "O365";
                    };
                };
            };
            "query" = "mutation StartBulkRecovery(`$input: StartBulkRecoveryInput!) {
                startBulkRecovery(input: `$input) {
                  bulkRecoveryInstanceId
                  taskchainId
                  jobId
                  error
                }
              }";
        }
    
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        if ($null -eq $response) {
            return 
        }
        if ($response.errors) {
            $response = $response.errors[0].message
            Write-Host "Error starting mass recovery $recoveryName. The error response is $($response).`n"
            return
        }
       
        $row = '' | Select-Object massRecoveryInstanceID,taskchainID, jobID, error
        $row.massRecoveryInstanceId = $response.data.startBulkRecovery.bulkRecoveryInstanceId
        $row.taskchainID = $response.data.startBulkRecovery.taskchainId
        $row.jobID = $response.data.startBulkRecovery.jobId
        $row.error = $response.data.startBulkRecovery.error

        Write-Host "Started mass recovery $recoveryName with the following details:"
        Write-Host $row
        Write-Host "`n"
    }

    return
}
Export-ModuleMember -Function Start-MassRecovery
