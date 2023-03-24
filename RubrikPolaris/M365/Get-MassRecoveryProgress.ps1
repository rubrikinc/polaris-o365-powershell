function Get-MassRecoveryProgress() {
    <#
    .SYNOPSIS
    
    Returns progress for mass recovery 
    
    .DESCRIPTION
    Returns the progress for the mass recovery with a given mass recovery
    instance ID

    .PARAMETER MassRecoveryInstanceId
    The instance ID of mass recovery you wish to check progress.

    .INPUTS
    None. You cannot pipe objects to Get-MassRecoveryProgress.
   
    .OUTPUTS
    System.Object. Get-MassRecoveryProgress returns status, current step, number 
    of failed, in progress, canceled, succeeded, total restore onedrive objects, etc. 
   
    .EXAMPLE
    PS> Get-MassRecoveryProgress -MassRecoveryInstanceId $massRecoveryInstanceId
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$MassRecoveryInstanceId,
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

    $payload = @{
        "operationName" = "BulkRecoveryProgress";
        "variables"     = @{
            "input" = @{
              "bulkRecoveryInstanceId" = $MassRecoveryInstanceId;
            };
          };
        "query"         = "query BulkRecoveryProgress(`$input: BulkRecoveryProgressInput!) {
            bulkRecoveryProgress(input: `$input) {
              recoveryPlanName
              startTime
              endTime
              elapsedTime
              currentStep
              failedObjects
              succeededObjects
              inProgressObjects
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

    $respProgress.startTime = getDateTime($respProgress.startTime)
    $respProgress.endTime = getDateTime($respProgress.endTime)
    $respProgress.elapsedTime = getTime($respProgress.elapsedTime)

    if ($respProgress.status -eq "Canceled") {
      $respProgress.failureReason = ""
      $cancelObjects = $respProgress.totalObjects - $respProgress.inProgressObjects - 
        $respProgress.succeededObjects - $respProgress.failedObjects
      
      $respProgress | Add-Member NoteProperty canceledObjects($cancelObjects)
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
    return "null"
  }

  $epochStart = Get-Date 01.01.1970
  $millisStamp = ($epochStart + ([System.TimeSpan]::frommilliseconds($unixTimeStamp))).ToLocalTime()
  return $millisStamp.ToString("MM/dd/yyyy HH:mm:ss")
}

function getTime($milliseconds) {
 $time = [System.TimeSpan]::frommilliseconds($milliseconds)
 return "$($time.Days) days, $($time.Hours) hours, $($time.Minutes) minutes, $($time.Seconds) seconds"
}

Export-ModuleMember -Function Get-MassRecoveryProgress
