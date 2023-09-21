function Get-MassRecoveryProgress() {
    <#
    .SYNOPSIS
    
    Returns progress for mass recovery 
    
    .DESCRIPTION
    Returns the progress for the mass recovery with a given mass recovery
    instance ID

    .PARAMETER MassRecoveryInstanceId
    The instance ID of mass recovery you wish to check progress.

    .PARAMETER SubscriptionName
    The subscription name for the mass recovery you wish to check progress.

    .INPUTS
    None. You cannot pipe objects to Get-MassRecoveryProgress.
   
    .OUTPUTS
    System.Object. Get-MassRecoveryProgress returns status, current step, number 
    of failed, in progress, canceled, succeeded, total restore onedrive objects, etc. 
   
    .EXAMPLE
    PS> Get-MassRecoveryProgress -MassRecoveryInstanceId $massRecoveryInstanceId -SubscriptionName $subscriptionName
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$MassRecoveryInstanceId,
        [Parameter(Mandatory=$True)]
        [String]$SubscriptionName,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    if((Test-IsGuid $MassRecoveryInstanceId) -eq $False) {
      Write-Host "Error fetching mass recovery progress from Rubrik. The error response is 'instance ID is not a valid UUID'."
      return
    }

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    Write-Host -Message "Fetching progress for mass recovery with instance ID $MassRecoveryInstanceId."

    $subscriptionId = getSubscriptionId($SubscriptionName)

    $payload = @{
        "operationName" = "BulkRecoveryProgress";
        "variables"     = @{
            "input" = @{
              "bulkRecoveryInstanceId" = $MassRecoveryInstanceId;
              "subscriptionId" = $subscriptionId;
            };
          };
        "query"         = "query BulkRecoveryProgress(`$input: BulkRecoveryProgressInput!) {
            bulkRecoveryProgress(input: `$input) {
              recoveryPlanName
              createTime
              startTime
              endTime
              elapsedTime
              currentStep
              failedObjects
              succeededObjects
              inProgressObjects
              canceledObjects
              objectsWithoutSnapshot
              totalObjects
              groupsProcessed
              totalGroups
              massRecoveryInstanceId: bulkRecoveryInstanceId
              massRecoveryDomain: bulkRecoveryDomain
              taskchainId
              failureActionType
              failureReason
              status
              groupProgresses {
                groupName
                groupId
                groupType
                workloadProgresses {
                  workloadType
                }
              }
            }
          }";
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers

    if ($null -eq $response) {
      return 
    }

    if ($response.errors) {
      $response = $response.errors[0].message
      Write-Host "Error fetching mass recovery progress from Rubrik. The error response is $($response)."
      return
    }

    $respProgress = $response.data.bulkRecoveryProgress

    $respProgress.createTime = getDateTime($respProgress.createTime)
    $respProgress.startTime = getDateTime($respProgress.startTime)
    $respProgress.endTime = getDateTime($respProgress.endTime)
    $respProgress.elapsedTime = getTime($respProgress.elapsedTime)

    if ($respProgress.startTime -eq "") {
      $respProgress.elapsedTime = ""
    }

    if ($respProgress.status -eq "CANCELED") {
      $respProgress.failureReason = ""
      $respProgress | Add-Member NoteProperty canceledObjects($cancelObjects)
    } else {
      $respProgress.PSObject.Properties.Remove('canceledObjects')
    }

    if ($respProgress.status -ne "IN_PROGRESS") {
        $respProgress.PSObject.Properties.Remove('currentStep')
    }

    $respProgress.failureActionType = "IGNORE_AND_CONTINUE"

    $overallProgress = $respProgress
    $groupProgress = $respProgress.groupProgresses[0]
    $workloadProgress = $groupProgress.workloadProgresses[0]

    $overallProgress.PSObject.Properties.Remove('groupProgresses')
    $groupProgress.PSObject.Properties.Remove('workloadProgresses')
 
    Write-Output $($overallProgress | Select-Object *)
    Write-Output $($groupProgress | Select-Object *)
    Write-Output $($workloadProgress | Select-Object *)

    return
}

function getDateTime($unixTimeStamp) {
  if ($null -eq $unixTimeStamp) {
    return ""
  }

  $epochStart = Get-Date 01.01.1970
  $millisStamp = ($epochStart + ([System.TimeSpan]::frommilliseconds($unixTimeStamp))).ToLocalTime()
  return $millisStamp.ToString("MM/dd/yyyy HH:mm:ss")
}

function getTime($milliseconds) {
 $time = [System.TimeSpan]::frommilliseconds($milliseconds)
 return "$($time.Days) days, $($time.Hours) hours, $($time.Minutes) minutes, $($time.Seconds) seconds"
}

function getSubscriptionId($subscriptionName) {
  $headers = @{
    'Content-Type' = 'application/json';
    'Accept' = 'application/json';
    'Authorization' = $('Bearer '+$global:RubrikPolarisConnection.accessToken);
  }

  $endpoint = $global:RubrikPolarisConnection.PolarisURL + '/api/graphql'
  $payload = @{
      "operationName" = "O365Orgs";
      "query" = "query O365Orgs {
        o365Orgs {
          nodes {
            id
            name
            status
          }
        }
      }";
  }

  $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
  $o365Orgs = $response.data.o365Orgs.nodes

  $subscriptionIds = @()
  $o365Orgs | ForEach-Object -Process {
    if (($_.name -eq $subscriptionName) -and ($_.status -eq "ACTIVE")) {
      $subscriptionIds += $_.id
    }
  }

  if ($subscriptionIds.Count -ne 1) {
    throw "There exists zero or more than 1 subscriptions with name '$subscriptionName'." 
  }

  return $subscriptionIds[0]
}

Export-ModuleMember -Function Get-MassRecoveryProgress
