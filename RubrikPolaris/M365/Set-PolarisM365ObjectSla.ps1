function Set-PolarisM365ObjectSla() {
    <#
    .SYNOPSIS

    Sets the SLA Domain for a selected Office 365 object (mailbox, onedrive or subscription).

    .DESCRIPTION

    Sets the protection for an O365 mailbox, onedrive or subscription in a given Polaris account, taking
    an API token, Polaris URL, object ID, and SLA ID.

    .PARAMETER ObjectID
    The object ID(s) for an O365 user or subscription. Can be obtained using 'Get-PolarisM365Mailbox', 'Get-PolarisM365OneDrive',
    'Get-PolarisM365Mailboxes', 'Get-PolarisM365OneDrives', or 'Get-PolarisM365Subscriptions' commands. This can take an array of object IDs.

    .PARAMETER SlaID
    The SLA ID for an SLA Domain. Can be obtained through the 'Get-PolarisSLA' command. Use the string
    'UNPROTECTED' to remove any SLA from this object, or the string 'DONOTPROTECT' to explicitly not protect
    this or any child objects.

    .INPUTS

    None. You cannot pipe objects to Set-PolarisM365ObjectSla.

    .OUTPUTS

    System.String. This returns the string 'Success' if the modification was successful, or throws an
    error if the command is not successful.

    .EXAMPLE

    PS> Set-PolarisM365ObjectSla -ObjectID $my_mailbox.id -SlaID $my_sla.id
    Success

    .EXAMPLE

    PS> Set-PolarisM365ObjectSla -ObjectID $my_onedrive.id -SlaID 'DONOTPROTECT'
    Success

    .EXAMPLE

    PS> Set-PolarisM365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_subscription.id -SlaID 'UNPROTECTED'
    Success
    #>

    param(
        [Parameter(Mandatory = $True)]
        [String[]]$ObjectID,
        [Parameter(Mandatory = $True)]
        [String]$SlaID,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $payload = @{
        "operationName" = "AssignSLA";
        "variables"     = @{
            "globalSlaAssignType"  = "protectWithSlaId";
            "globalSlaOptionalFid" = $SlaID;
            "objectIds"            = $ObjectID;
        };
        "query"         = "mutation AssignSLA(`$globalSlaOptionalFid: UUID, `$globalSlaAssignType: SlaAssignTypeEnum!, `$objectIds: [UUID!]!) {
            assignSla(globalSlaOptionalFid: `$globalSlaOptionalFid, globalSlaAssignType: `$globalSlaAssignType, objectIds: `$objectIds) {
                success
            }
        }";
    }

    if ($SlaID -eq 'UNPROTECTED') {
        $payload['variables']['globalSlaOptionalFid'] = $null
        $payload['variables']['globalSlaAssignType'] = 'noAssignment'
    }

    if ($SlaID -eq 'DONOTPROTECT') {
        $payload['variables']['globalSlaOptionalFid'] = $null
        $payload['variables']['globalSlaAssignType'] = 'doNotProtect'
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    if ($response.data.assignSla.success -eq $true) {
        return 'Success'
    }
    else {
        throw 'Issue assigning SLA domain to object'
    }
}
Export-ModuleMember -Function Set-PolarisM365ObjectSla

