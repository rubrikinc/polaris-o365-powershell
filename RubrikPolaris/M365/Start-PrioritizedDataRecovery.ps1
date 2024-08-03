function Start-PrioritizedDataRecovery() {
    <#
    .SYNOPSIS
    Prioritized restore Exchange, Sharepoint for a Group 
    
    .DESCRIPTION
    Prioritized restore Exchange, Sharepoint for a Group from latest backups before a
    given recovery point using the time range filter and other specified recovery options.

    .PARAMETER Name
    The name of the prioritized data recovery you wish to choose.
   
    .PARAMETER RecoveryPoint
    The date time you wish to use to restore closest earlier backups. The format
    is "YYYY-MM-DD HH:MM:SS"
    
    .PARAMETER SubscriptionName
    The subscription name you wish to prioritized restore under.
   
    .PARAMETER AdGroupId
    The ID of the AD Group you wish to prioritized restore Exchange.

    .PARAMETER ConfiguredGroupName
    The Name of the Configured Group you wish to prioritized restore Sharepoint.

    .PARAMETER WorkloadType
    The type of workload you wish to prioritized restore, "Exchange" and "Sharepoint"
    are supported right now.
  
    .PARAMETER MailboxFromTime
    The date time you wish to use to retore the emails received after that. The
    format is "YYYY-MM-DD HH:MM:SS"

    .PARAMETER MailboxUntilTime
    The date time you wish to use to retore the emails received before that. The
    format is "YYYY-MM-DD HH:MM:SS"

    .PARAMETER ArchiveFolderAction
    The Action of archive folder you wish to use to restore mailbox.
    
    .PARAMETER SharepointFromTime
    The date time you wish to use to retore the sites received after that. The
    format is "YYYY-MM-DD HH:MM:SS"

    .PARAMETER SharepointUntilTime
    The date time you wish to use to retore the sites received before that. The
    format is "YYYY-MM-DD HH:MM:SS"

    .PARAMETER ShouldSkipItemPermission
    The Action of skip item permission you wish to use to restore site.

    .PARAMETER SiteOwnerEmail
    The site owner email you wish to use when original site owner does not exist any more.
 
    .PARAMETER InplaceRecovery
    The Action of recover objects to original location and overwrite duplicates.

    .PARAMETER SubWorkloadType
    The type of sub workload you wish to restore. Only supported for "Exchange"
    workload type, where sub workload types are "Calendar", "Contacts" and "Mailbox"
    
    .INPUTS
    None. You cannot pipe objects to Start-PrioritizedDataRecovery.
   
    .OUTPUTS
    System.Object. The taskchainID, massRecoveryInstanceID, error and jobID for 
    the mass recovery job.
   
    .EXAMPLE
    PS> Start-PrioritizedDataRecovery -Name $name -RecoveryPoint $recoveryPoint -SubscriptionName $subscriptionName -AdGroupId $adGroupId -WorkloadType Exchange -MailboxFromTime $mailboxFromTime -MailboxUntilTime $mailboxUntilTime -InplaceRecovery $True
    
    PS> Start-PrioritizedDataRecovery -Name $name -RecoveryPoint $recoveryPoint -SubscriptionName $subscriptionName -AdGroupId $adGroupId -WorkloadType Exchange -ArchiveFolderAction $archiveFolderAction -InplaceRecovery $False

    PS> Start-PrioritizedDataRecovery -Name $name -RecoveryPoint $recoveryPoint -SubscriptionName $subscriptionName -ConfiguredGroupName $configuredGroupName -WorkloadType Sharepoint -SharepointFromTime $sharepointFromTime -SharepointUntilTime $sharepointUntilTime -ShouldSkipItemPermission $True -InplaceRecovery $True
   
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
        [ValidateSet("Exchange", "Sharepoint")]
        [String]$WorkloadType,
        [Parameter(Mandatory=$False)]
        [DateTime]$MailboxFromTime,
        [Parameter(Mandatory=$False)]
        [DateTime]$MailboxUntilTime,
        [Parameter(Mandatory=$False)]
        [ValidateSet("NO_ACTION", "EXCLUDE_ARCHIVE", "ARCHIVE_ONLY")]
        [String]$ArchiveFolderAction,
        [Parameter(Mandatory=$False)]
        [DateTime]$SharepointFromTime,
        [Parameter(Mandatory=$False)]
        [DateTime]$SharepointUntilTime,
        [Parameter(Mandatory=$False)]
        [Boolean]$ShouldSkipItemPermission,
        [Parameter(Mandatory=$False)]
        [String]$SiteOwnerEmail,
        [Parameter(Mandatory=$True)]
        [Boolean]$InplaceRecovery,
        [Parameter(Mandatory=$False)]
        [ValidateSet("Mailbox", "Calendar")]
        [String]$SubWorkloadType,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )
    
    if ($WorkloadType -eq "Exchange") {
        if ($AdGroupId -eq "") {
            Write-Host "Error starting prioritized data recovery $Name. AdGroupId should not be empty for Exchange workload type.`n"
            return
        }
        if ($ArchiveFolderAction -eq "") {
            $ArchiveFolderAction = "NO_ACTION"
        }
	if (($SubWorkloadType -ne "Mailbox") -and ($MailboxFromTime -eq $null) -and ($MailboxUntilTime -eq $null) -and ($ArchiveFolderAction -eq "NO_ACTION")) {
	    Write-Host "Error starting prioritized data recovery $Name. One of MailboxFromTime, MailboxUntilTime and ArchiveFolderAction should not be empty for Exchange Mailbox type.`n"
            return
	}
    }

    if ($WorkloadType -eq "Sharepoint") {
        if ($ConfiguredGroupName -eq "") {
            Write-Host "Error starting prioritized data recovery $Name. ConfiguredGroupName should not be empty for Sharepoint workload type.`n"
            return
        }
        if (($SharepointFromTime -eq $null) -and ($SharepointUntilTime -eq $null)) {
            Write-Host "Error starting prioritized data recovery $Name. One of SharepointFromTime, SharepointUntilTime should not be empty for Sharepoint.`n"
            return
        }
        if ($ShouldSkipItemPermission -eq "") {
            Write-Host "Error starting prioritized data recovery $Name. ShouldSkipItemPermission should not be empty for Sharepoint.`n"
            return
        }
    }    

    $calendarFromTime = (Get-Date).AddDays(-14) | Get-Date -format s
    
    if ($WorkloadType -eq "Mailbox") {
        Write-Host "Starting Prioritized Data Recovery $Name using MailboxTimeRange fromTime: $MailboxFromTime, untilTime: $MailboxUntilTime and CalendarTime Range fromTime: $calendarFromTime.`n"
    } elseif ($WorkloadType -eq "Sharepoint") {
        Write-Host "Starting Prioritized Data Recovery $Name using LastModifiedTimeFilter fromTime: $SharepointFromTime, untilTime: $SharepointUntilTime.`n"
    }

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
        "Sharepoint" = @(
            @{
                SnappableType = "O365_SHAREPOINT";
                SubSnappableType = "NONE";
                NameSuffix="Sharepoint";
                OperationalRecoverySpec = @{
                    "sharepointOperationalRecoverySpec" = @{
                        "lastModifiedTimeFilter" = @{
                            "fromTime" = $SharepointFromTime;
                            "untilTime" = $SharepointUntilTime;
                        };
                        "shouldSkipItemPermission" = $ShouldSkipItemPermission;
                        "siteOwnerEmail" = $SiteOwnerEmail;
                    };
                   "operationalRecoveryStage" = "INITIAL_OPERATIONAL_RECOVERY";
                };
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

    Write-Information -Message "Starting the prioritized restoration process for $WorkloadType account(s)."
  
    $subscriptionId = getSubscriptionId($SubscriptionName)
  
    $inplaceRecoverySpec = @{
        "nameCollisionRule" = "OVERWRITE";
    }

    if ($InplaceRecovery -eq $False) {
        $inplaceRecoverySpec = $null
    } 

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
                "operationalRecoverySpec" = $_.OperationalRecoverySpec;
                "inplaceRecoverySpec" = $inplaceRecoverySpec;
            }
        }

        if ($WorkloadType -eq "Sharepoint") {
            $o365GroupSelectorWithRecoverySpec = @{
                "baseInfo" = $baseInfo;
                "groupName" = $ConfiguredGroupName;
            }
        } else {
            # right now it is implicit that other workload types would be Exchange
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
            Write-Host "Error starting prioritized data recovery $recoveryName. The error response is $($response).`n"
            return
        }

        $row = '' | Select-Object massRecoveryInstanceID,taskchainID, jobID, error
        $row.massRecoveryInstanceId = $response.data.startBulkRecovery.bulkRecoveryInstanceId
        $row.taskchainID = $response.data.startBulkRecovery.taskchainId
        $row.jobID = $response.data.startBulkRecovery.jobId
        $row.error = $response.data.startBulkRecovery.error

        Write-Host "Started prioritized data recovery $recoveryName with the following details:"
        Write-Host $row
        Write-Host "`n"
    }

    return
}
Export-ModuleMember -Function Start-PrioritizedDataRecovery
