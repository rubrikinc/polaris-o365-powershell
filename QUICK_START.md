# Rubrik Polaris Microsoft 365 Protection PowerShell Module

## Overview

This repo contains the Rubrik Polaris Microsoft 365 PowerShell Module. This can be installed through the following command which needs to be run at the root level of this repo:

```powershell
Import-Module ./RubrikPolaris/RubrikPolaris.psd1 
```

## Using the module

### Connect to Rubrik

This module supports API Service Accounts for authentication. Once created, you'll need to rename the file to `polaris-service-account.json` and move the file to the `~/.rubrik` directory. To create a connection to Rubrik, use the following command:

```powershell
Connect-Polaris

```

This token can then be used with the other commands in the module.

## Featured Commands

The following commands are available in the module:

* `Connect-Polaris`
* `Disconnect-Polaris`
* `Get-PolarisSLA`
* `Get-PolarisM365Subscriptions`
* `Get-PolarisM365Mailboxes`
* `Get-PolarisM365Mailbox`
* `Get-PolarisM365OneDrives`
* `Get-PolarisM365OneDrive`
* `Get-PolarisM365SharePoint`
* `Set-PolarisM365ObjectSla`
* `Get-PolarisM365OneDrives`
* `Get-PolarisM365OneDriveSnapshot`
* `Restore-PolarisM365OneDrive`
* `Get-PolarisJob`
* `Get-PolarisM365EnterpriseApplication`
* `New-EnterpriseApplication`
* `Start-MassRecovery`
* `Stop-MassRecovery`
* `Get-MassRecoveryProgress`
* `Start-OperationalRecovery`
* `Start-CompleteOperationalRecovery`
* `Update-EnterpriseApplicationSecret`

Each command has help which describes their usage and parameters, these can be seen using the `Get-Help <command>` command within PowerShell.

## Example workflow

```powershell
Import-Module ./RubrikPolaris/RubrikPolaris.psd1 

Connect-Polaris

# get all M365 subscriptions
$all_subs = Get-PolarisM365Subscriptions 

# get just the subscription we want
$my_sub = $all_subs | ?{$_.name -eq $sub_name}

# get our SLA domain
$my_sla = Get-PolarisSLA  -Name $sla_name

# get our mailbox user
$my_mailbox = Get-PolarisM365Mailbox  -SubscriptionId $my_sub.id -SearchString 'arif'

# get our onedrive user
$my_onedrive = Get-PolarisM365OneDrive  -SubscriptionId $my_sub.id -SearchString 'arif'

# get our SharePoint sites
$my_sharepoint_sites = Get-PolarisM365SharePoint  -SubscriptionId $my_sub.id -SearchString 'arif' - Includes 'SitesOnly'

# get our SharePoint document libraries
$my_sharepoint_libraries = Get-PolarisM365SharePoint  -SubscriptionId $my_sub.id -SearchString 'arif' - Includes 'DocumentLibrariesOnly'

# get our SharePoint sites and document libraries
$my_sharepoint_both = Get-PolarisM365SharePoint  -SubscriptionId $my_sub.id -SearchString 'arif'

# set the SLA domain for our mailbox user
Set-PolarisM365ObjectSla  -ObjectID $my_mailbox.id -SLAID $my_sla.id

# set the SLA domain for our onedrive
Set-PolarisM365ObjectSla  -ObjectID $my_onedrive.id -SLAID $my_sla.id

# or set the SLA domain of our subscription
Set-PolarisM365ObjectSla  -ObjectID $my_sub.id -SLAID $my_sla.id
```

## Bulk Assignment Use-case

We can bulk assign SLAs to M365 accounts, if for example we want to do a staggered migration. The following code can be used to do this:

```powershell
Connect-Polaris
$all_orgs = Get-PolarisM365Subscriptions 
$my_org = $all_orgs | ?{$_.name -eq $org_name}
$all_slas = Get-PolarisSLA 
$my_sla = $all_slas | ?{$_.name -eq 'Gold'}
$all_mailboxes = Get-PolarisM365Mailboxes  -SubscriptionId $my_org.id
$500_mailboxes = $all_mailboxes | select -First 500
$assign = Set-PolarisM365ObjectSla  -ObjectID $500_mailboxes.id -SlaID $my_sla.id

# Can filter using - Includes 'SitesOnly'/'DocumentLibrariesOnly'/'ListsOnly'
$all_sharepoint_objects = Get-PolarisM365SharePoint  -SubscriptionId $my_org.id
$500_sharepoint_objects = $all_sharepoint_objects | select -First 500
$assign_sharepoint = Set-PolarisM365ObjectSla  -ObjectID $500_sharepoint_objects.id -SlaID $my_sla.id
```

The key line here is this:

```powershell
$500_mailboxes = $all_mailboxes | select -First 500
```

Here we are just picking the first 500 accounts, but more specific selection could be achieved using standard PowerShell filtering on the `$all_mailboxes` array.

## Bulk removing direct SLA (Inherited SLA will be applied)

```powershell
Import-Module ./RubrikPolaris/RubrikPolaris.psd1 
Connect-Polaris
$org_name = 'my-org-name'
$all_orgs = Get-PolarisM365Subscriptions 
$my_org = $all_orgs | ?{$_.name -eq $org_name}

# Print out the org info
Write-Output $my_org

# Can filter using - Includes 'SitesOnly'/'DocumentLibrariesOnly'/'ListsOnly'
$all_sharepoint_lists = Get-PolarisM365SharePoint  -SubscriptionId $my_org.id -Includes 'ListOnly'
$500_sharepoint_lists = $all_sharepoint_lists | select -First 500
$assign_sharepoint = Set-PolarisM365ObjectSla  -ObjectID $500_sharepoint_lists.id -SlaID 'UNPROTECTED'

# Validate if the assignment was successful
Write-Output $assign_sharepoint
```
Here we are just picking the first 500 SharePoint lists, but more specific selection could be achieved using standard PowerShell filtering on the `$all_sharepoint_lists` array.
