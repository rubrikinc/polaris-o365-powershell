#Requires -Version 6.0 -RunAsAdministrator
function Update-EnterpriseApplicationSecret() {
    <#
    .SYNOPSIS

    Updates secret of exostomg Application in Microsoft 365 and then adds to Rubrik.

    .PARAMETER DataSource
    The type of Enterprise Application you wish to update. Valid values are: Exchange, SharePoint, OneDrive.

    .PARAMETER AppIdFile
    File name containing the id of Application, for which secret needs to be updated. Each id should be in a new line.


    .PARAMETER ExpirationInYears
    The number of years in which the Enterprise Application will be valid for before expiring. The default value is 200.

    .PARAMETER SubscriptionID
    The id of the Microsoft 365 Subscription in the Rubrik where the Enterprise Applications is authorized.

    .DESCRIPTION

    Connect to Microsoft 365 and update a secret to Azure Application. Once updated, secret is updated in Rubrik. 
    The Microsoft Graph "Application.ReadWrite.All" and "AppRoleAssignment.ReadWrite.All" permissions are required to update an Enterprise Application.
.
    
.INPUTS

    None. You cannot pipe objects to Update-EnterpriseApplicationSecret.

    .OUTPUTS

    System.Collections.ArrayList. New-EnterpriseApplication returns an array list 
    containing the AppId, Secret, and DataSource for each Microsoft 365 Enterprise Application connected
    to Rubrik.

    .EXAMPLE

    PS> Update-EnterpriseApplicationSecret -DataSource Exchange -AppIdFile RubrikPolaris/appID.txt -ExpirationInYears 1 -SubscriptionId 6db7521d-aa22-4265-a9f2-1a555c291a4f

    AppId             : fa6c4986-63fc-496e-b610-a15483d6a09d
    Secret            : __T8Q~ZfT0PKkiMfLC12ypHk1c2goR2vVITV-dxF
    DataSource        : Exchange
    CertRawDataBase64 : 
    PemRawDataBase64  : 
    #>

    param(
        [Parameter(Mandatory = $True)]
        [ValidateSet("Exchange", "SharePoint", "OneDrive")]
        [String]$DataSource,
        [Parameter(Mandatory=$true, HelpMessage="Input file containing app IDs to onboard")]
        [String]$AppIdFile,
        [Parameter(Mandatory = $true)]
        [Int]$ExpirationInYears = 200,
        [Parameter(Mandatory = $true)]
        [String]$SubscriptionId,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    if ($ExpirationInYears -le 0){
        throw "The ExpirationInYears argument must be greater than 0."
    }

    # Read app ID file
    $appIds = Get-Content -Path $AppIdFile -ErrorAction Stop

    # Validate the required 'Microsoft.Graph' module is installed
    # and provide a user friendly message when it's not.
    Write-Host  "Info: Checking for required 'Microsoft.Graph' module."
    if (Get-Module -ListAvailable -Name Microsoft.Graph)
    {
        
    }
    else
    {
        throw "The 'Microsoft.Graph' is required for this script. Run the follow command to install: Install-Module Microsoft.Graph"
    }

    $configuredSubscriptionsOnRubrik = Get-PolarisM365Subscriptions
    $numberOfSubscriptionsCOnfigured = $configuredSubscriptionsOnRubrik.Count

    Write-Information -Message "Info: Verifing a Microsoft 365 subscription has been set up on Rubrik."
    if ($numberOfSubscriptionsCOnfigured -eq 0){
        throw "A Microsoft 365 subscription must be set up before adding additional Enterprise Applications."
    } else {
        $subscriptionExist = $False
        foreach ($sub in $configuredSubscriptionsOnRubrik){
            if ($sub.subscriptionId -eq $SubscriptionId) {
                $subscriptionExist = $True
                Write-Output $subscriptionId
                break
                
            }     
        }

        if (-not $subscriptionExist){
            throw "The '$($SubscriptionId)' Subscription has not been set up on Rubrik."
        }
    }

    # SharePoint Self-Signed Certificate Creation Variables
    if ($DataSource -eq "SharePoint"){
        $sslConfigFileName = "RubrikSSLConfigTemp.txt"
        $sslConfig = "[req]
        req_extensions = v3_req
        distinguished_name = req_distinguished_name
        
        [req_distinguished_name]
        
        [v3_req]
        keyUsage=critical,digitalSignature,keyCertSign
        extendedKeyUsage=clientAuth,serverAuth"

        $privateKeySize = 2048
        $privateKeyFileName = "RubrikTempPrivateKey.pem"
        $certFileName = "RubrikTempCert.pem"
        $CertSubject = "/O=Rubrik"
        $CertExpireDays = $ExpirationInYears * 365

        $convertedCertFileName = "RubrikTempConvertedFileCert.crt"
        $openSSLVersion = openssl version
        $supportWindowsVersion = "OpenSSL 3.4.1"

        if ($IsWindows){
            if ($openSSLVersion -notmatch $supportWindowsVersion){
                throw "The SharePoint Enterprise Application creation process requires OpenSSL v3.4.1 Please download the non-light installer (https://slproweb.com/products/Win32OpenSSL.html) and try again."
            }
        } 

        if ($PSVersionTable.PSVersion.Major -lt 6){
            throw "The SharePoint Enterprise Application creation process requires PowerShell 6.0 or higher. Please upgrade and try again."
        }


        New-Item -Path . -Name  $sslConfigFileName -ItemType "file" -Value  $sslConfig | Out-Null
    } 

    Write-Information -Message "Info: Connecting to the Microsoft Graph API using the 'Application.ReadWrite.All' and 'AppRoleAssignment.ReadWrite.All' Scope."
    Connect-Graph -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All","User.Read" -ErrorAction Stop | Out-Null
    Write-Information -Message "Info: Successfully authenticated the Microsoft Graph API."


    $passwordcred = @{
        "displayName" = "New-Secret"
        "endDateTime" = Get-Date (Get-Date).ToUniversalTime().AddYears($ExpirationInYears) -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'
    }

    $o365AppType = @{
        "OneDrive" = "ONEDRIVE"
        "Exchange" = "EXCHANGE"
        "SharePoint" = "SPOINT"
    }

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }

    # Get a list of all the apps
    $endpoint = $PolarisURL + '/api/graphql'
    $payload = @{
        "operationName" = "ListO365AppsQuery";
        "query"         = "query ListO365AppsQuery(`$o365AppFilters: [AppFilter!]!) {
            listO365Apps(
                o365AppFilters: `$o365AppFilters
            ) {
                edges {
                node {
                    appId
                    subscriptionId
                    appType
                }
                }
            }
            }";
        "variables"     = @{
            "o365AppFilters" = @();
        }
    }

    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers

    $app_id_to_type = @{}
    foreach ($entApp in $response.data.listO365Apps.edges) {
        $app_id_to_type[$entApp.node.appId] = $entApp.node.appType
    }


    $enterpriceApplicationDetails = New-Object System.Collections.ArrayList
    # Process each app ID
    foreach ($appId in $appIds) {
        Write-Host -ForegroundColor Cyan "Updating Secret $appId..." -NoNewline
        $filterParams = "AppID eq '$appId'"
        $app = Get-MgApplication -Filter $filterParams

        if (-not $app_id_to_type.ContainsKey($appId)) {
            Write-Host -ForegroundColor Red "Skipping secret update for $appId, as not present in RSC"
            continue
        }

        $rsc_app_type = $app_id_to_type[$appId]
        if ($o365AppType[$DataSource] -ne $rsc_app_type) {
            Write-Host -ForegroundColor Red "Skipping secret update for $appId, as app datasource is $rsc_app_type"
            continue
        }

        $appExists = $app.count -gt 0
        if ($appExists) {
            if ($DataSource -eq "SharePoint") {
                Write-Information -Message "Info: Creating an RSA private key for the SharePoint Enterprise Application."
                if ($IsWindows){
                    # On Windows openssl genrsa does not support additional options and SHA256 is already the default.
                    openssl genrsa -traditional -out $privateKeyFileName $privateKeySize 2>$null
                } else {
                    openssl genrsa -traditional -out $privateKeyFileName $privateKeySize -sha256 -nodes 2>$null
                }
    
                Write-Information -Message "Info: Creating a x509 self-signed certificate using the private key."
                openssl req -key $privateKeyFileName -new -x509 -days $CertExpireDays -out $certFileName -sha256 -subj $CertSubject -config $sslConfigFileName -extensions v3_req
    
                # Cert Raw Data Sent to M365
                $certRawData = Get-Content "${certFileName}" -AsByteStream
                # Cert Raw Data in Base 64 sent to Rubrik
                $certRawDataBase64 = [System.Convert]::ToBase64String($certRawData)
    
                $pemRawData  = Get-Content "${privateKeyFileName}" -AsByteStream
                # Private Key in Base 64 sent to Rubrik
                $pemRawDataBase64 = [System.Convert]::ToBase64String($pemRawData)
    
            
    
                $keyCredential = New-Object Microsoft.Graph.PowerShell.Models.MicrosoftGraphKeyCredential
                $keyCredential.Type = "AsymmetricX509Cert";
                $keyCredential.Usage = "Verify";
                $keyCredential.Key = $certRawData;
            
                # Update-MgServicePrincipal -ServicePrincipalId $SPId -KeyCredentials $PrivateKeyCreds -PasswordCredentials $PasswordCreds
                Write-Information -Message "Info: Adding the certs to the Enterprise Application"
                try {
                    # Update-MgApplication  -ApplicationId $newEnterpriseApp.Id -KeyCredentials $KeyCreds
                    Update-MgApplication  -ApplicationId $app.Id -KeyCredentials $($keyCredential)
                    
                    # Update-MgServicePrincipal -ServicePrincipalId $newEnterpriseApp.Id -BodyParameter $params
                    Write-Information -Message "Info: Successfully added the certs to the Enterprise Application"
    
                }
                catch {
                    $errorMessage = $_.Exception | Out-String
    
                    Write-Host "Error adding the certification to $($app.Id) to Rubrik. The error resposne is $($errorMessage)."
                }
    
    
                Remove-Item  "$certFileName"
                Remove-Item "${privateKeyFileName}"
            }

            foreach ($password in $app.PasswordCredentials) {
                Remove-MgApplicationPassword -ApplicationId $app.ID -KeyId $password.KeyId
            }

            $addPasswordToApp = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $passwordCred  -InformationAction "SilentlyContinue"
            # Sleeping so that app password is synced on Azure side.
            Start-Sleep 30         

            Write-Information -Message "Info: Storing the completed Enterprise Application details to memory."
            $tempEntAppDetails = New-Object System.Object
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "AppId" -Value $appId
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "Secret" -Value $addPasswordToApp.SecretText
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "DataSource" -Value $DataSource
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "CertRawDataBase64" -Value $certRawDataBase64
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "PemRawDataBase64" -Value $pemRawDataBase64

            $enterpriceApplicationDetails.Add($tempEntAppDetails) | Out-Null
        } else {
            Write-Host -ForegroundColor Red "Application $appId do not exist in Azure"
        }

        if ($DataSource -ne "SharePoint"){
            $certRawData = ""
            $pemRawData = ""
            $gqlQueryArgumentType = "`$o365AppType: String!, `$o365AppClientId: String!, `$o365AppClientSecret: String!, `$o365SubscriptionId: String!, `$updateAppCredentials: Boolean"
            $gqlInput = "input: {appType: `$o365AppType, appClientId: `$o365AppClientId, appClientSecret: `$o365AppClientSecret, subscriptionId: `$o365SubscriptionId, updateAppCredentials: `$updateAppCredentials}"
        } else {
            $certRawData = $certRawDataBase64
            $pemRawData = $pemRawDataBase64
            $gqlQueryArgumentType = "`$o365AppType: String!, `$o365AppClientId: String!, `$o365AppClientSecret: String!, `$o365SubscriptionId: String!, `$o365Base64AppCertificate: String!, `$o365Base64AppPrivateKey: String!, `$updateAppCredentials: Boolean"
            $gqlInput = "input: {appType: `$o365AppType, appClientId: `$o365AppClientId, appClientSecret: `$o365AppClientSecret, subscriptionId: `$o365SubscriptionId, base64AppCertificate: `$o365Base64AppCertificate, base64AppPrivateKey: `$o365Base64AppPrivateKey, updateAppCredentials: `$updateAppCredentials}"
        }

        $payload = @{
            "operationName" = "AddCustomerO365AppMutation";
            "variables" = @{
                "o365AppType" = $o365AppType[$DataSource] ;
                "o365AppClientId" = $appId;
                "o365AppClientSecret" = $addPasswordToApp.SecretText;
                "o365SubscriptionId" = $subscriptionId;
                "o365Base64AppCertificate" =  $certRawData;
                "o365Base64AppPrivateKey" =  $pemRawData;
                "updateAppCredentials" = $true;


            };
            "query" = "mutation AddCustomerO365AppMutation($gqlQueryArgumentType) {
                insertCustomerO365App($gqlInput) {
                    success
                }
            }";
        }

        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        if ($response.data.insertCustomerO365App.success -eq $true) {
            Write-Host -ForegroundColor Green "Success."
        }
        else {
            if ($response.errors){
                $response = $response.errors[0].message
            }
            Write-Host  -ForegroundColor Red  "Error updating Application $($appId) secret. The error resposne is $($response)."
        }      
    }
    if ($null -ne $sslConfigFileName){
        Remove-Item "$sslConfigFileName"

    } 
    return $enterpriceApplicationDetails
}

Export-ModuleMember -Function Update-EnterpriseApplicationSecret
