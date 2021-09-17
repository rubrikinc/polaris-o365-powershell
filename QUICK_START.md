# Rubrik Polaris Office 365 Protection PowerShell Module

## Overview

This repo contains the Rubrik Polaris Office 365 PowerShell Module. This can be imported to a PowerShell session or script using the following command:

```powershell
. .\PolarisO365.ps1
```

## Using the module

### Getting a token

An authentication token for the Polaris instance being used will be needed to use the commands. This can be obtained by the following command:

```powershell
NAME
    Get-PolarisToken

SYNTAX
    Get-PolarisToken [-Username] <string> [-Password] <string> [-PolarisURL] <string>  [<CommonParameters>]
```

This command can be used as follows to store the token as a PowerShell variable:

```powershell
# declare our variables
$url = 'https://myinstance.my.rubrik.com'   # this is the URL for your Polaris instance
$username = 'steve@mydomain.com'            # the username for your Polaris instance
$password = 'MyPass123!'                    # the password for your Polaris instance
# store the token
$token = Get-PolarisToken -Username $username -Password $password -PolarisURL $url
```

This token can then be used with the other commands in the module.

## Featured Commands

The following commands are available in the module:

* `Get-PolarisToken`
* `Get-PolarisSLA`
* `Get-PolarisO365Subscriptions`
* `Get-PolarisO365Mailboxes`
* `Get-PolarisO365Mailbox`
* `Get-PolarisO365OneDrives`
* `Get-PolarisO365OneDrive`
* `Get-PolarisO365SharePoint`
* `Set-PolarisO365ObjectSla`

Each command has help which describes their usage and parameters, these can be seen using the `Get-Help <command>` command within PowerShell.

## Example workflow

```powershell
. .\PolarisO365.ps1 # import the module

# set variables
$url = 'https://myinstance.my.rubrik.com'
$username = 'steve@mydomain.com'
$password = 'MyPass123!'
$sub_name = 'mysubscription'
$sla_name = 'Bronze'

# get a token
$token = Get-PolarisToken -Username $username -Password $password -PolarisURL $url

# get all O365 subscriptions
$all_subs = Get-PolarisO365Subscriptions -Token $token -PolarisURL $url

# get just the subscription we want
$my_sub = $all_subs | ?{$_.name -eq $sub_name}

# get our SLA domain
$my_sla = Get-PolarisSLA -Token $token -PolarisURL $url -Name $sla_name

# get our mailbox user
$my_mailbox = Get-PolarisO365Mailbox -Token $token -PolarisURL $url -SubscriptionId $my_sub.id -SearchString 'arif'

# get our onedrive user
$my_onedrive = Get-PolarisO365OneDrive -Token $token -PolarisURL $url -SubscriptionId $my_sub.id -SearchString 'arif'

# get our SharePoint sites
$my_sharepoint_sites = Get-PolarisO365SharePoint -Token $token -PolarisURL $url -SubscriptionId $my_sub.id -SearchString 'arif' - Includes 'SitesOnly'

# get our SharePoint document libraries
$my_sharepoint_libraries = Get-PolarisO365SharePoint -Token $token -PolarisURL $url -SubscriptionId $my_sub.id -SearchString 'arif' - Includes 'DocumentLibrariesOnly'

# get our SharePoint sites and document libraries
$my_sharepoint_both = Get-PolarisO365SharePoint -Token $token -PolarisURL $url -SubscriptionId $my_sub.id -SearchString 'arif'

# set the SLA domain for our mailbox user
Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_mailbox.id -SLAID $my_sla.id

# set the SLA domain for our onedrive
Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_onedrive.id -SLAID $my_sla.id

# or set the SLA domain of our subscription
Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $my_sub.id -SLAID $my_sla.id
```

## Bulk Assignment Use-case

We can bulk assign SLAs to O365 accounts, if for example we want to do a staggered migration. The following code can be used to do this:

```powershell
$token = Get-PolarisToken -Username $username -Password $password -PolarisURL $url
$all_orgs = Get-PolarisO365Subscriptions -Token $token -PolarisURL $url
$my_org = $all_orgs | ?{$_.name -eq $org_name}
$all_slas = Get-PolarisSLA -Token $token -PolarisURL $url
$my_sla = $all_slas | ?{$_.name -eq 'Gold'}
$all_mailboxes = Get-PolarisO365Mailboxes -Token $token -PolarisURL $url -SubscriptionId $my_org.id
$500_mailboxes = $all_mailboxes | select -First 500
$assign = Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $500_mailboxes.id -SlaID $my_sla.id

# Can filter using - Includes 'SitesOnly'/'DocumentLibrariesOnly'
$all_sharepoint_objects = Get-PolarisO365SharePoint -Token $token -PolarisURL $url -SubscriptionId $my_org.id
$500_sharepoint_objects = $all_sharepoint_objects | select -First 500
$assign_sharepoint = Set-PolarisO365ObjectSla -Token $token -PolarisURL $url -ObjectID $500_sharepoint_objects.id -SlaID $my_sla.id
```

The key line here is this:

```powershell
$500_mailboxes = $all_mailboxes | select -First 500
```

Here we are just picking the first 500 accounts, but more specific selection could be achieved using standard PowerShell filtering on the `$all_mailboxes` array.
