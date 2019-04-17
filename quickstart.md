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
