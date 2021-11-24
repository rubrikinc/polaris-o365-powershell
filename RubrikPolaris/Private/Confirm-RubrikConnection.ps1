function Confirm-RubrikConnection() {
    <#
    .SYNOPSIS

    Validates that the Connect-Polaris cmdlet was ran. 

    .DESCRIPTION

    Make sure the rubrikPolarisConnection global variables have been set.

    .INPUTS

    None. You cannot pipe objects to Confirm-Connect.

    .OUTPUTS

    None.
 
    .EXAMPLE

    PS> Confirm-Connect
    #>

    Write-Verbose -Message 'Validate the Rubrik token exists'
    if (-not $global:rubrikPolarisConnection.accessToken ) {
        throw 'The Rubrik API Access token was not found. Please run the `Connect-Polaris` cmdlet.'
    }

    if (-not $global:rubrikPolarisConnection.PolarisURL ) {
        throw 'The Rubrik Polaris URL was not found. Please run the `Connect-Polaris` cmdlet.'
    }
    

}



