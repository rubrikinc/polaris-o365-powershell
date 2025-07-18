#Requires -Version 6.0 -RunAsAdministrator
function New-EnterpriseApplication() {
    <#
    .SYNOPSIS

   Create a new Enterprise Application in Microsoft 365 and then adds to Rubrik.

    .PARAMETER DataSource
    The type of Enterprise Application you wish to created. Valid values are: Exchange, SharePoint and OneDrive.

    .PARAMETER ExpirationInYears
    The number of years in which the Enterprise Application will be valid for before expiring. The default value is 200.

    .PARAMETER SubscriptionName
    The name of the Microsoft 365 Subscription in the Rubrik where the Enterprise Applications shoudl be created. This is only required when multiple subscriptions have been set up.

    .DESCRIPTION

    Connect to Microsoft 365 and created a new Enterprise Application. Once created, add the Enterprise Application to Rubrik. 
    The Microsoft Graph "Application.ReadWrite.All" and "AppRoleAssignment.ReadWrite.All" permissions are required to create an Enterprise Application.
.
    
.INPUTS

    None. You cannot pipe objects to New-EnterpriseApplication.

    .OUTPUTS

    System.Collections.ArrayList. New-EnterpriseApplication returns an array list 
    containing the AppId, Secret, and DataSource for each Microsoft 365 Enterprise Application connected
    to Rubrik.

    .EXAMPLE

    PS> New-EnterpriseApplication  -DataSource "Exchange" -Count 5
    
    AppId                                Secret                                DataSource
    -----                                ------                                ----------
    f79d1f98-f0ad-4c41-82a7-a86225eq6b76 5Zm70~AL1HCcU6eCU5HHHuIBfnE5et5OE19JE SharePoint
    #>

    param(
        [Parameter(Mandatory = $True)]
        [ValidateSet("Exchange", "SharePoint", "OneDrive")]
        [String]$DataSource,
        [Parameter(Mandatory = $False)]
        [Int]$Count,
        [Int]$ExpirationInYears = 200,
        [Parameter(Mandatory = $False)]
        [String]$SubscriptionName,
        [Parameter(Mandatory = $False)]
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    if ($ExpirationInYears -le 0){
        throw "The ExpirationInYears argument must be greater than 0."
    }

    # Validate the required 'Microsoft.Graph' module is installed
    # and provide a user friendly message when it's not.
    Write-Information -Message "Info: Checking for required 'Microsoft.Graph' module."
    if (Get-Module -ListAvailable -Name Microsoft.Graph)
    {

    }
    else
    {
        throw "The 'Microsoft.Graph' is required for this script. Run the follow command to install: Install-Module Microsoft.Graph"
    }
    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }


    $endpoint = $PolarisURL + '/api/graphql'

    $configuredSubscriptionsOnRubrik = Get-PolarisM365Subscriptions
    $numberOfSubscriptionsConfigured = $configuredSubscriptionsOnRubrik.Count

    Write-Information -Message "Info: Verifing a Microsoft 365 subscription has been set up on Rubrik."
    if ($numberOfSubscriptionsConfigured -eq 0){
        throw "A Microsoft 365 subscription must be set up before adding additional Enterprise Applications."
    } elseif ($numberOfSubscriptionsConfigured -eq 1){
        $subscriptionId = $configuredSubscriptionsOnRubrik.subscriptionId
    } else {

        if ($PSBoundParameters.ContainsKey("SubscriptionName") -eq $false){
            throw "The 'SubscriptionName'parameter must be specified when multiple Microsoft 365 subscriptions are being protected."
        }

        foreach ($sub in $configuredSubscriptionsOnRubrik){
            if ($sub.name -eq $subscriptionName) {
                $subscriptionId = $sub.subscriptionId
                Write-Output $subscriptionId
                break

            }
        }

        if ($subscriptionId -eq $null){
            throw "The '$($SubscriptionName)' Subscription has not been set up on Rubrik."
        }


    }

    $o365AppType = @{
        "OneDrive" = "ONEDRIVE"
        "Exchange" = "EXCHANGE"
        "SharePoint" = "SPOINT"

    }


    if ($PSBoundParameters.ContainsKey('Count') -eq $False) {
        $Count = 1
    }

    $polarisAccountName = $PolarisURL.Replace("https://", "").Replace(".my.rubrik.com", "").Replace("http://", "")

    $applicationName = @{
        "Exchange"  = "Rubrik Exchange - " + $($polarisAccountName)
        "OneDrive" = "Rubrik OneDrive - $($polarisAccountName)"
        "SharePoint" = "Rubrik SharePoint - $($polarisAccountName)"
    }

    $secretExpiryDate = Get-Date (Get-Date).ToUniversalTime().AddYears($ExpirationInYears) -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'
    $passwordcred = @{
        "displayName" = $applicationName[$DataSource]
        "endDateTime" = $secretExpiryDate
    }

    # API Service Principal IDs
    $grapApiAppId = "00000003-0000-0000-c000-000000000000"
    $ewsApiAppId = "00000002-0000-0ff1-ce00-000000000000"
    $sharePointApiAppId = "00000003-0000-0ff1-ce00-000000000000"

    # Microsoft Teams can only be added through the  Rubrik UI due
    # to API limitations from Microsoft. Keeping this in place for future
    # proofing
    # $teamsPointApiAppId = $sharePointApiAppId

    #Static GUID identifiers which is required for New-MgApplication
    #                                Mail.ReadWrite,                        Group.ReadAll,                           Contacts.ReadWrite,                    Calendars.ReadWrite,                    User.Read.All,                          Reports.Read.All
    $exchangeGraphPermissionsGuid = 'e2a3a72e-5f79-4c64-b1b1-878b674786c9', '5b567255-7703-4780-807c-7be8301ae99b', '6918b873-d17a-4dc1-b314-35f528134491', 'ef54d2bf-783f-4e0f-bca1-3210c0444d99', 'df021288-bdef-4463-88db-98f22de89214', '230c1aed-a721-4c5d-9cb4-a90514e508ef'
    #                              User.Read.All,                          full_access_as_app,                     Mail.ReadWrite,                        Contacts.ReadWrite,                      Calendars.ReadWrite.All,                 Tasks.ReadWrite
    $exchangeEwsPermissionsGuid = 'bf24470f-10c1-436d-8d53-7b997eb473be', 'dc890d15-9560-4a4c-9b7f-a736ec74ec40', 'e2a3a72e-5f79-4c64-b1b1-878b674786c9', '6918b873-d17a-4dc1-b314-35f528134491', 'ef54d2bf-783f-4e0f-bca1-3210c0444d99', '2c6a42ca-0d4d-49ad-bc0e-21222c449a65'
    #                               Sites.Read.All,                         Files.ReadWrite.All,                    User.Read.All,                          Sites.FullControl.All
    $oneDriveGraphPermissionsGuid = '332a536c-c7ef-4017-ab91-336970924f0d', '75359482-378d-4052-8f01-80520e7db3cd', 'df021288-bdef-4463-88db-98f22de89214', 'a82116e5-55eb-4c41-a434-62fe8a61c773'
    #                               Sites.FullControl.All
    $oneDriveSpointPermissionGuid = '678536fe-1083-478a-9c59-b99265e6b0d3'
    # Same permissions as OneDrive
    $sharePointGraphPermissionGuid = $oneDriveGraphPermissionsGuid
    # Same permissions as OneDrive
    $sharePointSpointPermissionGuid = $oneDriveSpointPermissionGuid

    $enterpriseApplicationDetails = New-Object System.Collections.ArrayList
    $servicePrincipalAppRoleAssignedRetry = New-Object System.Collections.ArrayList

    $toCreateDetails = @{
        $DataSource =  $Count
    }

    Write-Information -Message "Info: Will create $Count $DataSource Enterprise Application(s)."

    $needsCert = ($DataSource -eq "SharePoint" -Or $DataSource -eq "OneDrive")
    # Self-Signed Certificate Creation Variables
    if ($needsCert){
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
                throw "The $($DataSource) Enterprise Application creation process requires OpenSSL v3.4.1 Please download the non-light installer (https://slproweb.com/products/Win32OpenSSL.html) and try again."
            }
        }

        if ($PSVersionTable.PSVersion.Major -lt 6){
            throw "The $($DataSource) Enterprise Application creation process requires PowerShell 6.0 or higher. Please upgrade and try again."
        }


        New-Item -Path . -Name  $sslConfigFileName -ItemType "file" -Value  $sslConfig | Out-Null

    }

    Write-Information -Message "Info: Connecting to the Microsoft Graph API using the 'Application.ReadWrite.All' and 'AppRoleAssignment.ReadWrite.All' Scope."
    Connect-Graph -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All","User.Read" -ErrorAction Stop | Out-Null
    Write-Information -Message "Info: Successfully authenticated the Microsoft Graph API."

    foreach ($source in $toCreateDetails.GetEnumerator()) {
        $DataSource = $source.Name
        $Count = $source.Value

        1..$Count | ForEach-Object {

            Write-Information -Message "Info: Creating a $($DataSource) Enterprise Application."

            try {

                $newEnterpriseApp = New-MgApplication -DisplayName $applicationName[$DataSource] -SignInAudience "AzureADMyOrg" -InformationAction "SilentlyContinue" -ErrorAction Stop

            }
            catch {

                $errorMessage = $_.Exception | Out-String


                if($errorMessage.Contains('Insufficient privileges to complete the operation')) {
                    throw "Microsoft returned a 'Insufficient privileges to complete the operation' error message."
                }

                while ($true) {
                    $newEnterpriseApp = New-MgApplication -DisplayName $applicationName[$DataSource] -SignInAudience "AzureADMyOrg" -InformationAction "SilentlyContinue"
                    if ($newEnterpriseApp){
                        break
                    } else {
                        Start-Sleep 5
                    }
                }

                $newEnterpriseApp = New-MgApplication -DisplayName $applicationName[$DataSource] -SignInAudience "AzureADMyOrg" -InformationAction "SilentlyContinue"


            }

            Write-Information -Message "Info: Adding a Password to the Enterprise Application."


            try {
                $addPasswordToApp = Add-MgApplicationPassword -ApplicationId $newEnterpriseApp.Id -PasswordCredential $passwordCred  -InformationAction "SilentlyContinue"
            }
            catch {

                # Wait for the Enterprise Application to be populated in the Microsoft database
                while ($true) {
                    $appStatusCheck = Get-MgApplication -Filter "AppId eq '$($newEnterpriseApp.AppId)'" -InformationAction "SilentlyContinue"
                    if ($appStatusCheck){
                        break
                    } else {
                        Start-Sleep 5
                    }
                }

                $addPasswordToApp = Add-MgApplicationPassword -ApplicationId $newEnterpriseApp.Id -PasswordCredential $passwordCred -InformationAction "SilentlyContinue"

            }

            Write-Information -Message "Info: Adding a Service Principal to the Enterprise Application."

            $newServicePrincipal = New-MgServicePrincipal -AppId $newEnterpriseApp.AppId -InformationAction "SilentlyContinue"

            Write-Information -Message "Info: Getting the Service Principal ID from Microsoft."
            $graphApiServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$($grapApiAppId)'" -InformationAction "SilentlyContinue"

            function AddApiPermissionsToEnterpriseApplication() {
                <#
                .SYNOPSIS

                Add API permissions to an existing Enterprise Application in Microsoft 365.

                .PARAMETER PermissionGuids
                The type of Enterprise Application you wish to created. Valid values are: Exchange, SharePoint and OneDrive.

                .PARAMETER ApiServicePrincipalId
                The AppId of the API you wish to add permissions to.

                #>
                param(
                    [Parameter(Mandatory = $True)]
                    [String[]]$PermissionGuids,
                    [Parameter(Mandatory = $True)]
                    [String]$ApiServicePrincipalId
                )

                foreach ( $iD in $PermissionGuids  ) {
                    try {
                        New-MgServicePrincipalAppRoleAssignedTo `
                        -ServicePrincipalId $newServicePrincipal.Id `
                        -ResourceId $ApiServicePrincipalId `
                        -PrincipalId $newServicePrincipal.Id `
                        -AppRoleId $iD -InformationAction "SilentlyContinue" | Out-Null
                    }
                    catch {
                        $tempSpDetails = New-Object System.Object
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ServicePrincipalId" -Value $newServicePrincipal.Id
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ResourceId" -Value $ApiServicePrincipalId
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "PrincipalId" -Value $newServicePrincipal.Id
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "AppRoleId" -Value $iD
                        $servicePrincipalAppRoleAssignedRetry.Add($tempSpDetails) | Out-Null
                    }
                }
            }

            if ($DataSource -eq "Exchange") {

                $ewsApiServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$($ewsApiAppId)'" -InformationAction "SilentlyContinue"

                AddApiPermissionsToEnterpriseApplication -PermissionGuids $exchangeGraphPermissionsGuid -ApiServicePrincipalId $graphApiServicePrincipal.Id

                AddApiPermissionsToEnterpriseApplication -PermissionGuids $exchangeEwsPermissionsGuid -ApiServicePrincipalId $ewsApiServicePrincipal.Id

            }
            elseif ($DataSource -eq "OneDrive") {
                Write-Information -Message "Info: Adding the required API Permissions to the Enterprise Application."

                $sharePointApiServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$($sharePointApiAppId)'" -InformationAction "SilentlyContinue"

                AddApiPermissionsToEnterpriseApplication -PermissionGuids $oneDriveGraphPermissionsGuid -ApiServicePrincipalId $graphApiServicePrincipal.Id

                AddApiPermissionsToEnterpriseApplication -PermissionGuids $oneDriveSpointPermissionGuid -ApiServicePrincipalId $sharePointApiServicePrincipal.Id
            }
            elseif ($DataSource -eq "SharePoint") {

                $sharePointApiServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$($sharePointApiAppId)'" -InformationAction "SilentlyContinue"

                AddApiPermissionsToEnterpriseApplication -PermissionGuids $sharePointGraphPermissionGuid -ApiServicePrincipalId $graphApiServicePrincipal.Id

                AddApiPermissionsToEnterpriseApplication -PermissionGuids $sharePointSpointPermissionGuid -ApiServicePrincipalId $sharePointApiServicePrincipal.Id

            }

            if ($needsCert){
                Write-Information -Message "Info: Creating an RSA private key for the $($DataSource) Enterprise Application."
                if ($IsWindows -Or $IsMacOS){
                    # On Windows/MacOS openssl genrsa does not support additional options and SHA256 is already the default.
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
                    Update-MgApplication  -ApplicationId $newEnterpriseApp.Id -KeyCredentials $($keyCredential)

                    # Update-MgServicePrincipal -ServicePrincipalId $newEnterpriseApp.Id -BodyParameter $params
                    Write-Information -Message "Info: Successfully added the certs to the Enterprise Application"

                }
                catch {
                    $errorMessage = $_.Exception | Out-String

                    Write-Host "Error adding the certification to $($newEnterpriseApp.Id) to Rubrik. The error resposne is $($errorMessage)."
                }

                Remove-Item  "$certFileName"
                Remove-Item "${privateKeyFileName}"
            }

            Write-Information -Message "Info: Storing the completed Enterprise Application details to memory."
            $tempEntAppDetails = New-Object System.Object
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "AppId" -Value $newEnterpriseApp.AppId
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "Secret" -Value $addPasswordToApp.SecretText
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "DataSource" -Value $DataSource
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "CertRawDataBase64" -Value $certRawDataBase64
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "PemRawDataBase64" -Value $pemRawDataBase64

            $enterpriseApplicationDetails.Add($tempEntAppDetails) | Out-Null
        }

    }
    if ($servicePrincipalAppRoleAssignedRetry.Count -gt 0) {
        foreach ( $retry in $servicePrincipalAppRoleAssignedRetry  ) {
            New-MgServicePrincipalAppRoleAssignedTo `
            -ServicePrincipalId $retry.ServicePrincipalId `
            -ResourceId $retry.ResourceId `
            -PrincipalId $retry.PrincipalId `
            -AppRoleId $retry.AppRoleId -InformationAction "SilentlyContinue" | Out-Null

        }
    }

    Write-Information -Message "Info: Disconnecting from the Microsoft Graph API"

    Disconnect-Graph
    $staticSleepPeriod = 60

    Write-Information -Message "Info: Waiting $($staticSleepPeriod) seconds to allow the Microsoft database to sync."
    Start-Sleep -Seconds $staticSleepPeriod
    foreach ( $app in $enterpriseApplicationDetails  ) {

        if ($needsCert) {
            $certRawData = $app.CertRawDataBase64
            $pemRawData = $app.PemRawDataBase64
            $gqlQueryArgumentType = "`$o365AppType: String!, `$o365AppClientId: String!, `$o365AppClientSecret: String!, `$o365SubscriptionId: String!, `$o365Base64AppCertificate: String!, `$o365Base64AppPrivateKey: String!, `$appSecretExpiry: DateTime"
            $gqlInput = "input: {appType: `$o365AppType, appClientId: `$o365AppClientId, appClientSecret: `$o365AppClientSecret, subscriptionId: `$o365SubscriptionId, base64AppCertificate: `$o365Base64AppCertificate, base64AppPrivateKey: `$o365Base64AppPrivateKey, appSecretExpiry: `$appSecretExpiry}"
        } else {
            $certRawData = ""
            $pemRawData = ""
            $gqlQueryArgumentType = "`$o365AppType: String!, `$o365AppClientId: String!, `$o365AppClientSecret: String!, `$o365SubscriptionId: String!, `$appSecretExpiry: DateTime"
            $gqlInput = "input: {appType: `$o365AppType, appClientId: `$o365AppClientId, appClientSecret: `$o365AppClientSecret, subscriptionId: `$o365SubscriptionId, appSecretExpiry: `$appSecretExpiry}"
        }
        $payload = @{
            "operationName" = "AddCustomerO365AppMutation";
            "variables" = @{
                "o365AppType" = $o365AppType[$app.DataSource] ;
                "o365AppClientId" = $app.AppId;
                "o365AppClientSecret" = $app.Secret;
                "o365SubscriptionId" = $subscriptionId;
                "o365Base64AppCertificate" =  $certRawData;
                "o365Base64AppPrivateKey" =  $pemRawData;
                "appSecretExpiry" = $secretExpiryDate;
            };
            "query" = "mutation AddCustomerO365AppMutation($gqlQueryArgumentType) {
                insertCustomerO365App($gqlInput) {
                    success
                }
            }";
        }

        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        if ($response.data.insertCustomerO365App.success -eq $true) {
            Write-Host "Successfully added Enterprise Application $($app.AppId) to Rubrik."
        }
        else {
            if ($response.errors){
                $response = $response.errors[0].message
            }
            Write-Host "Error adding Application $($app.AppId) to Rubrik. The error resposne is $($response)."
        }

    }


    if ($null -ne $sslConfigFileName){
        Remove-Item "$sslConfigFileName"

    } 
    return $enterpriseApplicationDetails
    
}
Export-ModuleMember -Function New-EnterpriseApplication






