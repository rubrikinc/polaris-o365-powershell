function Remove-EnterpriseApplication {
    <#
    .SYNOPSIS
    
    Removes the specified enterprise applications from Microsoft.
    
    .DESCRIPTION
    This function deletes specified enterprise applications in Microsoft Azure AD via Microsoft Graph PowerShell.
    The function prompts the user to input an RSC account name, application type (SharePoint or OneDrive), 
    and the Application IDs that need to retain.
    
    .PARAMETER AccountName
    The account name of the RSC.
    
    .PARAMETER AppType
    The application type to delete (OneDrive or SharePoint).
    
    .PARAMETER AppsToKeep
    An array of Application IDs to keep.
    
    .PARAMETER DryRun
    If set to $true, the function will only display what would be deleted without actually deleting the apps.

    .EXAMPLE
    PS> Remove-EnterpriseApplications -AccountName "MyAccountName" -AppType "OneDrive" -AppsToKeep "ID1","ID2" -DryRun $true
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$AccountName,
        [Parameter(Mandatory=$True)]
        [String]$AppType,
        [Parameter(Mandatory=$True)]
        [Array]$AppsToKeep,
        [Parameter(Mandatory=$True)]
        [Bool]$DryRun
    )

    # Check Application Type validity
    if ($AppType -ne 'SharePoint' -and $AppType -ne 'OneDrive') {
        Write-Host "Exiting script due to invalid app type."
        return
    }

    # Connect to Microsoft Graph PowerShell
    Connect-MgGraph -Scopes 'Application.ReadWrite.All'

    # Build Application Name
    $appName = 'Rubrik ' + $AppType + ' - ' + $AccountName

    # Get the list of all enterprise applications with the provided name
    $allEnterpriseApps = Get-MgServicePrincipal -Filter "DisplayName eq '$appName'"

    # Filter out the specified apps based on the App IDs to keep
    $appsToDelete = $allEnterpriseApps | Where-Object { 
        ($AppsToKeep -notcontains $_.AppId)
    }
    
    if ($DryRun -eq $true) {
        # Information about applications to delete for a dry run
        foreach ($app in $appsToDelete) {
            Write-Host "Dry run: Would delete Enterprise Application with App ID '$($app.AppId)'"
        }
    } else {
        # Actual deletion of applications
        foreach ($app in $appsToDelete) {
            Remove-MgServicePrincipal -ServicePrincipalId $app.Id
            Write-Host "Enterprise Application with App ID '$($app.AppId)' deleted successfully."
        }
    }

    # Disconnect from Microsoft Graph
    Disconnect-MgGraph
}

Export-ModuleMember -Function Remove-EnterpriseApplication
