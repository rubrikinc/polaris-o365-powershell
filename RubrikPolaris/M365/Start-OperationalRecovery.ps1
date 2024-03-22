function Start-OperationalRecovery() {
    <#
    .SYNOPSIS
    Operational restore Exchange for an AD Group 
    
    .DESCRIPTION
    Operational restore Exchange for an AD Group from latest backups before a
    given recovery point in time using the mailboxTimeRange.

    .PARAMETER Name
    The name of the operational recovery you wish to choose.
   
    .PARAMETER RecoveryPoint
    The date time you wish to use to restore closest earlier backups. The format
    is "YYYY-MM-DD HH:MM:SS"
    
    .PARAMETER SubscriptionName
    The subscription name you wish to operational restore under.
   
    .PARAMETER AdGroupId
    The ID of the AD Group you wish to operational restore.

    .PARAMETER WorkloadType
    The type of workload you wish to mass restore, only "Exchange" is supported
    right now.
  
    .PARAMETER MailboxFromTime
    The date time you wish to use to retore the emails received after that. The
    format is "YYYY-MM-DD HH:MM:SS"

    .PARAMETER MailboxUntilTime
    The date time you wish to use to retore the emails received before that. The
    format is "YYYY-MM-DD HH:MM:SS"

    .PARAMETER ArchiveFolderAction
    The Action of archive folder you wish to use to restore mailbox.

    .PARAMETER SubWorkloadType
    The type of sub workload you wish to restore. Only supported for "Exchange"
    workload type, where sub workload types are "Calendar", "Contacts" and "Mailbox"
    
    .INPUTS
    None. You cannot pipe objects to Start-OperationalRecovery.
   
    .OUTPUTS
    System.Object. The taskchainID, massRecoveryInstanceID, error and jobID for 
    the mass recovery job.
   
    .EXAMPLE
    PS> Start-OperationalRecovery -Name $name -RecoveryPoint $recoveryPoint -SubscriptionName $subscriptionName -AdGroupId $adGroupId -WorkloadType $workloadType -MailboxFromTime $mailboxFromTime -MailboxUntilTime $mailboxUntilTime
    
    PS> Start-OperationalRecovery -Name $name -RecoveryPoint $recoveryPoint -SubscriptionName $subscriptionName -AdGroupId $adGroupId -WorkloadType $workloadType -ArchiveFolderAction $archiveFolderAction 
   
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
        [ValidateSet("Exchange")]
        [String]$WorkloadType,
        [Parameter(Mandatory=$False)]
        [DateTime]$MailboxFromTime,
        [Parameter(Mandatory=$False)]
        [DateTime]$MailboxUntilTime,
        [Parameter(Mandatory=$False)]
        [ValidateSet("NO_ACTION", "EXCLUDE_ARCHIVE", "ARCHIVE_ONLY")]
        [String]$ArchiveFolderAction,
        [Parameter(Mandatory=$False)]
        [ValidateSet("Mailbox", "Calendar", "Contacts")]
        [String]$SubWorkloadType,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )
    
    if (($SubWorkloadType -eq "Mailbox") -or ($SubWorkloadType -eq "")) {
        if ($ArchiveFolderAction -eq "") {
            $ArchiveFolderAction = "NO_ACTION"
        }
	if (($MailboxFromTime -eq $null) -and ($MailboxUntilTime -eq $null) -and ($ArchiveFolderAction -eq "NO_ACTION")) {
	    Write-Host "Error starting operational recovery $Name. One of MailboxFromTime, MailboxUntilTime and ArchiveFolderAction should not be empty for Exchange Mailbox type.`n"
            return
	}
    }    

    $calendarFromTime = (Get-Date).AddDays(-14) | Get-Date -format s
    Write-Host "Starting Operational Recovery $Name using MailboxTimeRange fromTime: $MailboxFromTime, untilTime: $MailboxUntilTime and CalendarTime Range fromTime: $calendarFromTime.`n"

    $snappableToSubSnappableMap = @{
        "Exchange" = @(
            @{
                SnappableType = "O365_EXCHANGE";
                SubSnappableType = "O365_MAILBOX";
                NameSuffix="Mailbox";
                OperationalRecoverySpec = @{
                    "mailboxOperationalRecoverySpec" = @{
                        "mailboxTimeRange" = @{
                            "fromTime" = $MailboxFromTime;
                            "untilTime" = $MailboxUntilTime;
                        };
                        "archiveFolderAction" = $ArchiveFolderAction;
		    };
                   "operationalRecoveryStage" = "INITIAL_OPERATIONAL_RECOVERY";
                };
            };
            @{
                SnappableType = "O365_EXCHANGE";
                SubSnappableType = "O365_CALENDAR";
                NameSuffix="Calendar";
                OperationalRecoverySpec = @{
                    "calendarOperationalRecoverySpec" = @{
                        "calendarTimeRange" = @{
                            "fromTime" = $calendarFromTime 
                        }; 
                    };
                   "operationalRecoveryStage" = "INITIAL_OPERATIONAL_RECOVERY";
                };
            };
            @{
                SnappableType = "O365_EXCHANGE";
                SubSnappableType = "O365_CONTACT";
                NameSuffix="Contacts";
                OperationalRecoverySpec = $null;
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

    Write-Information -Message "Starting the operational restoration process for $WorkloadType account(s) under AD Group ID $AdGroupId."
  
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
                "operationalRecoverySpec" = $_.OperationalRecoverySpec
            }
        }

        $o365GroupSelectorWithRecoverySpec = @{
            "baseInfo" = $baseInfo;
            "adGroupId" = $AdGroupId;
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
            Write-Host "Error starting operational recovery $recoveryName. The error response is $($response).`n"
            return
        }

        $row = '' | Select-Object massRecoveryInstanceID,taskchainID, jobID, error
        $row.massRecoveryInstanceId = $response.data.startBulkRecovery.bulkRecoveryInstanceId
        $row.taskchainID = $response.data.startBulkRecovery.taskchainId
        $row.jobID = $response.data.startBulkRecovery.jobId
        $row.error = $response.data.startBulkRecovery.error

        Write-Host "Started operational recovery $recoveryName with the following details:"
        Write-Host $row
        Write-Host "`n"
    }

    return
}
Export-ModuleMember -Function Start-OperationalRecovery
