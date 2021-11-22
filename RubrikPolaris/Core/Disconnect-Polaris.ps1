function Disconnect-Polaris() {
    <#
    .SYNOPSIS

    Disconnects from Polaris. 

    .DESCRIPTION

    Remove all information used to connect to the Polaris.

    .INPUTS

    None. You cannot pipe objects to Disconnect-Polaris.

    .OUTPUTS

    None.
 
    .EXAMPLE

    PS> Disconnect-Polaris
    #>

    Remove-Variable -Name rubrikPolarisConnection -Scope Global

    

}
Export-ModuleMember -Function Disconnect-Polaris

