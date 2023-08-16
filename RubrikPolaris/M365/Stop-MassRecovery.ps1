function Stop-MassRecovery() {
    <#
    .SYNOPSIS
    Stops mass recovery 
    
    .DESCRIPTION
    Stops the mass recovery with a given mass recovery instance ID

    .PARAMETER MassRecoveryInstanceId
    The instance ID of mass recovery you wish to stop.

    .PARAMETER SubscriptionName
    The subscription name for the mass recovery you wish to stop.

    .INPUTS
    None. You cannot pipe objects to Get-MassRecoveryProgress.
   
    .OUTPUTS
    System.Object. Stop-MassRecovery returns true in case it is success 
   
    .EXAMPLE
    PS> Stop-MassRecovery -MassRecoveryInstanceId $massRecoveryInstanceId -SubscriptionName $subscriptionName
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
      Write-Host "Error stopping mass recovery with instance ID $MassRecoveryInstanceId. The error response is 'instance ID is not a valid UUID'."
      return
    }

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $subscriptionId = getSubscriptionId($SubscriptionName)

    Write-Information -Message "Stopping mass recovery with instance ID $MassRecoveryInstanceId."

    $payload = @{
        "operationName" = "CancelBulkRecovery";
        "variables"     = @{
            "input" = @{
              "bulkRecoveryInstanceId" = $MassRecoveryInstanceId;
              "subscriptionId" = $subscriptionId;
            };
          };
        "query"         = "mutation CancelBulkRecovery(`$input: CancelBulkRecoveryInput!) {
          cancelBulkRecovery(input: `$input) {
            success
          }
        }";
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    
    if ($null -eq $response) {
      return 
    }

    if ($response.errors) {
      $response = $response.errors[0].message
      Write-Host "Error stopping mass recovery with instance ID $MassRecoveryInstanceId. The error response is $($response)."
      return
    }

    if ($response.data.cancelBulkRecovery.success) {
      return "Successfully canceled mass recovery for instance ID $MassRecoveryInstanceId."
    } else {
      return "Failed to cancel mass recovery for instance ID $MassRecoveryInstanceId."
    }
}
Export-ModuleMember -Function Stop-MassRecovery

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
