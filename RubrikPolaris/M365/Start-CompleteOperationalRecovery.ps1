function Start-CompleteOperationalRecovery() {
    <#
    .SYNOPSIS
    Starts complete operational recovery 
    
    .DESCRIPTION
    Starts the complete operational recovery with a given initial operation recovery instance ID and M365 subscription Name.

    .PARAMETER MassRecoveryInstanceId
    The instance ID of initial operational recovery you wish to complete .

    .PARAMETER SubscriptionName
    The subscription name for the initial operational recovery you wish to complete.

    .INPUTS
    None. You cannot pipe objects to Start-CompleteOperationalRecovery().
   
    .OUTPUTS
    System.Object. The taskchainID, massRecoveryInstanceID, error, jobID and recoveryName for 
    the complete operational recovery job.  
 
    .EXAMPLE
    PS> Start-CompleteOperationalRecovery -MassRecoveryInstanceId $massRecoveryInstanceId -SubscriptionName $subscriptionName
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
      Write-Host "Error starting complete operational recovery with instance ID $MassRecoveryInstanceId. The error response is 'instance ID is not a valid UUID'."
      return
    }

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $subscriptionId = getSubscriptionId($SubscriptionName)

    Write-Host "Starting complete operational recovery with instance ID $MassRecoveryInstanceId."
    Write-Host "`n"

    $payload = @{
        "operationName" = "CompleteOperationalRecovery";
        "variables"     = @{
            "input" = @{
              "bulkRecoveryInstanceId" = $MassRecoveryInstanceId;
              "subscriptionId" = $subscriptionId;
            };
          };
        "query" = "mutation CompleteOperationalRecovery(`$input: CompleteOperationalRecoveryInput!) {
            completeOperationalRecovery(input: `$input) {
              bulkRecoveryInstanceId
              taskchainId
              recoveryName
              error
              jobId   
          }
        }";
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    
    if ($null -eq $response) {
      return 
    }

    if ($response.errors) {
      $response = $response.errors[0].message
      Write-Host "Error starting complete operational recovery with instance ID $MassRecoveryInstanceId. The error response is $($response)."
      return
    }

    $row = '' | Select-Object massRecoveryInstanceID,taskchainID, jobID, error
    $row.massRecoveryInstanceId = $response.data.completeOperationalRecovery.bulkRecoveryInstanceId
    $row.taskchainID = $response.data.completeOperationalRecovery.taskchainId
    $row.jobID = $response.data.completeOperationalRecovery.jobId
    $row.error = $response.data.completeOperationalRecovery.error

    $recoveryName = $response.data.completeOperationalRecovery.recoveryName

    Write-Host "Started complete operational recovery $recoveryName with the following details:"
    Write-Host $row
    Write-Host "`n"
    return
}
Export-ModuleMember -Function Start-CompleteOperationalRecovery

function Test-IsGuid
{
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$StringGuid
    )

   $ObjectGuid = [System.Guid]::empty
   return [System.Guid]::TryParse($StringGuid,[System.Management.Automation.PSReference]$ObjectGuid) # Returns True if successfully parsed
}
