# Installation

## Prerequisites

* PowerShell Version v6.0 or higher. Updated versions can be [downloaded directly from Microsoft](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.3)
* Microsoft Graph PowerShell Module. Run the following command to install: `Install-Module Microsoft.Graph`
* When creating SharePoint Enterprise Applications [OpenSSL v3.4.1](https://slproweb.com/products/Win32OpenSSL.html) and Windows is also required.


## Rubrik Module

Download the Rubrik PowerShell Module to your local environment and extrat the Zip file.

[![Download zip](https://user-images.githubusercontent.com/8610203/145614905-a6d64f3a-adab-4c3f-9bf9-ffa4fdf6793f.png "Download zip")](https://github.com/rubrikinc/polaris-o365-powershell/archive/refs/heads/master.zip)

1. Open PowerShell and navigate to the extracted `polaris-o365-powershell-master` folder.
2. Install the module by running `Import-Module ./RubrikPolaris/RubrikPolaris.psd1`

## Rubrik API Service Account

The PowerShell module will leverage a Rubrik API Service Account to connect the Enterprise Applications to Rubrik after they are created in Microsoft 365.

1. Log in to Rubrik.
2. Click the gear icon and select Users and Roles.
3. Click Service Accounts.
4. Click Add Service Account.
5. In Name, type the polaris-service-account. **This is case sensitive.**
6. Optional: In Description, type the description of the service account.
7. Click Next.
8. Select the Administrator role to be assigned to the service account.
9. Click Add.
10. Click Download As JSON.
11. Click Done.
12. Create a hidden folder named ".rubrik" in your home directory by running the following CLI command: `mkdir ~/.rubrik`
13. Move the downloaded polaris-service-account.json to the `~/.rubrik` directory

# Creating the Enterprise Application

Note - At most, Enterprise Applications should be created in batches of 50 to avoid Microsoft throttling slowdowns. If running the command for the first time we recommend only create one to validate functionality. 

```
$InformationPreference = "Continue"
Connect-Polaris
New-EnterpriseApplication -DataSource Exchange -Count 1
Disconnect-Polaris
```

1. Enable logging

`$InformationPreference= "Continue"`
 

2. Connect to Rubrik

`Connect-Polaris`

3. Create the Enterprise Application. After running the command you will be prompted to authenticate into Microsoft using Global administrator credentials

`New-EnterpriseApplication -DataSource Exchange -Count 1`

Valid DataSource options are Exchange, OneDrive, or SharePoint.

4. Once finished, disconnect from Polaris

`Disconnect-Polaris`


