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

    .PARAMETER ConfiguredGroupName
    The name of the Configured Group you wish to mass restore.

    .PARAMETER WorkloadType
    The type of workload you wish to mass restore, only "OneDrive" is supported
    right now.

    .PARAMETER SubWorkloadType
    The type of sub workload you wish to restore. Only supported for "Exchange"
    workload type, where sub workload types are "Calendar", "Contacts" and "Mailbox"

    .INPUTS
    None. You cannot pipe objects to Start-MassRecovery.
   
    .OUTPUTS
    System.Object. The taskchainID, massRecoveryInstanceID, error and jobID for 
    the mass recovery job.
   
    .EXAMPLE
    PS> Start-MassRecovery -Name $name -RecoveryPoint $recoveryPoint -SubscriptionName $subscriptionName -AdGroupId $adGroupId
        -WorkloadType $workloadType

    PS> Start-MassRecovery -Name $name -RecoveryPoint $recoveryPoint -SubscriptionName $subscriptionName -ConfiguredGroupName $configuredGroupName
        -WorkloadType $workloadType
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$Name,
        [Parameter(Mandatory=$True)]
        [DateTime]$RecoveryPoint,
        [Parameter(Mandatory=$True)]
        [String]$SubscriptionName,
        [Parameter(Mandatory=$False)]
        [String]$AdGroupId,
        [Parameter(Mandatory=$False)]
        [String]$ConfiguredGroupName,
        [Parameter(Mandatory=$True)]
        [ValidateSet("OneDrive", "Exchange", "Sharepoint")]
        [String]$WorkloadType,
        [Parameter(Mandatory=$True)]
        [Boolean]$InplaceRecovery,
        [Parameter(Mandatory=$False)]
        [ValidateSet("Mailbox", "Calendar", "Contacts")]
        [String]$SubWorkloadType,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    if ((($WorkloadType -eq "OneDrive") -or ($WorkloadType -eq "Exchange")) -and ($AdGroupId -eq "")) {
        Write-Host "Error starting mass recovery $Name. AdGroupId should not be empty for OneDrive or Exchange workload type.`n"
        return
    }

    if (($WorkloadType -eq "Sharepoint") -and ($ConfiguredGroupName -eq "")) {
        Write-Host "Error starting mass recovery $Name. ConfiguredGroupName should not be empty for Sharepoint workload type.`n"
        return
    }

    if (($WorkloadType -ne "Exchange") -and ($SubWorkloadType -ne "")) {
        Write-Host "Error starting mass recovery $Name. SubWorkloadType should only be specified for Exchange workload type.`n"
        return
    }

    $InplaceRecoverySpec = @{
        "nameCollisionRule" = "OVERWRITE";
    }
    if ($InplaceRecovery -eq $False) { 
        $InplaceRecoverySpec = $null
    }

    $snappableToSubSnappableMap = @{
        "OneDrive" = @(
            @{
                SnappableType = "O365_ONEDRIVE";
                SubSnappableType = "NONE";
                NameSuffix="OneDrive";
                InplaceRecoverySpec = $null;
            }
        );
        "Exchange" = @(
            @{
                SnappableType = "O365_EXCHANGE";
                SubSnappableType = "O365_MAILBOX";
                NameSuffix="Mailbox";
                InplaceRecoverySpec = $InplaceRecoverySpec;               
            };
            @{
                SnappableType = "O365_EXCHANGE";
                SubSnappableType = "O365_CALENDAR";
                NameSuffix="Calendar";
                InplaceRecoverySpec = $InplaceRecoverySpec;
            };
            @{
                SnappableType = "O365_EXCHANGE";
                SubSnappableType = "O365_CONTACT";
                NameSuffix="Contacts";
                InplaceRecoverySpec = $InplaceRecoverySpec;
            };
        );
        "Sharepoint" = @(
            @{
                SnappableType = "O365_SHAREPOINT";
                SubSnappableType = "NONE";
                NameSuffix="Sharepoint";
                InplaceRecoverySpec = $InplaceRecoverySpec;
            }
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
  
    $subscriptionId = getSubscriptionId($SubscriptionName)
  
    $snappableToSubSnappableMap[$WorkloadType] | Where-Object {
        ($_.NameSuffix -eq $SubWorkloadType) -or ($SubWorkloadType -eq "")
    } | ForEach-Object -Process {
        $recoveryName=$Name+"_"+$_.NameSuffix
        $baseInfo = @{
            "snappableType" = $_.SnappableType;
            "subSnappableType" = $_.SubSnappableType;
            "recoverySpec" = @{
                "recoveryPoint" = $rpMilliseconds;
                "srcSubscriptionId" = $subscriptionId;
                "targetSubscriptionId" = $subscriptionId;
                "inplaceRecoverySpec" = $_.InplaceRecoverySpec;
            }
        }

        if ($WorkloadType -eq "Sharepoint") {
            $o365GroupSelectorWithRecoverySpec = @{
                "baseInfo" = $baseInfo;
                "groupName" = $ConfiguredGroupName;
            }
        } else {
            # right now it is implicit that other workload types would be OneDrive, Exchange
            $o365GroupSelectorWithRecoverySpec = @{
                "baseInfo" = $baseInfo;
                "adGroupId" = $AdGroupId;
            }
        }

        $payload = @{
            "operationName" = "StartBulkRecovery";
            "variables"     = @{
                "input" = @{
                    "definition" = @{
                        "name" = $recoveryName;
                        "o365GroupSelectorWithRecoverySpec" = $o365GroupSelectorWithRecoverySpec;
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
