function New-EnterpriseApplication() {
    <#
    .SYNOPSIS

    (In Beta) Create a new Enterprise Application and then add to Rubrik.

    .DESCRIPTION

    (In Beta) Connect to Microsoft 365 and created a new Enterprise Application. Once created, add the Enterprise Application to Rubrik. 
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
        [ValidateSet("Exchange", "SharePoint", "OneDrive", "FirstFull")]
        [String]$DataSource,
        [Parameter(Mandatory = $False)]
        [Int]$Count,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )


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

    if ($DataSource -eq "SharePoint"){
        throw "SharePoint Enterprise Applications temporarily need to be created through the Rubrik UI."
    } 

    $headers = @{
        'Content-Type'  = 'application/json';
        'Accept'        = 'application/json';
        'Authorization' = $('Bearer ' + $Token);
    }


    $endpoint = $PolarisURL + '/api/graphql'

    $payload = @{
        "operationName" = "O365OrgCountAndComplianceQuery";
        "query" = "query O365OrgCountAndComplianceQuery {
            o365Orgs {
              count
            }
          }";
    }       

    Write-Information -Message "Info: Verying a Microsoft 365 subscription has been set up on Rubrik."
    $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
    if ($response.data.o365Orgs.count -lt 1) {
        throw "A Microsoft 365 subscription must be set up before adding additional Enterprise Applications."
    }

    $o365AppType = @{
        "OneDrive" = "ONEDRIVE"
        "Exchange" = "EXCHANGE"
        "SharePoint" = "SPOINT"

    }

    Write-Information -Message "Info: Connecting to the Microsoft Graph API using the 'Application.ReadWrite.All' and 'AppRoleAssignment.ReadWrite.All' Scope."
    Connect-Graph -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All" -ErrorAction Stop | Out-Null
    Write-Information -Message "Info: Successfully authenticated the Microsoft Graph API."
    

    if ($PSBoundParameters.ContainsKey('Count') -eq $False) {
        $Count = 1
    }

    $polarisAccountName = $PolarisURL.Replace("https://", "").Replace(".my.rubrik.com", "").Replace("http://", "")

    $applicationName = @{
        "Exchange"  = "Rubrik Exchange - " + $($polarisAccountName)
        "OneDrive" = "Rubrik OneDrive - $($polarisAccountName)"
        "SharePoint" = "Rubrik SharePoint - $($polarisAccountName)"
    }


    $passwordcred = @{
        "displayName" = $applicationName[$DataSource]
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
    # Sites.Read.All, Sites.ReadWrite.All, User.Read.All
    $oneDriveGraphPermissionsGuid = '332a536c-c7ef-4017-ab91-336970924f0d', '9492366f-7969-46a4-8d15-ed1a20078fff', 'df021288-bdef-4463-88db-98f22de89214'
    # Same permissions as OneDrive + Sites.FullControl.All
    $sharePointGraphPermissionGuid = $oneDriveGraphPermissionsGuid + 'a82116e5-55eb-4c41-a434-62fe8a61c773'                                                                                                                                      
    # Sites.FullControl.All
    $sharePointSpointPermissionGuid = '678536fe-1083-478a-9c59-b99265e6b0d3'
    
    $enterpriceApplicationDetails = New-Object System.Collections.ArrayList
    $servicePrincipalAppRoleAssignedRetry = New-Object System.Collections.ArrayList

    if ($DataSource -eq "FirstFull"){
    
        $toCreateDetails = @{
            "Exchange" = 154
            "OneDrive" = 16
            # "SharePoint" = 24
        }

        Write-Information -Message "Info: Will create $($toCreateDetails.Exchange) Exchange, $($toCreateDetails.OneDrive) OneDrive, and $($toCreateDetails.SharePoint) SharePoint Enterprise Applications."

    } else {
        $toCreateDetails = @{
            $DataSource =  $Count
        }

        Write-Information -Message "Info: Will create $Count $DataSource Enterprise Application(s)."


    }

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
            if ($DataSource -eq "Exchange") {
            
                $ewsApiServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$($ewsApiAppId)'" -InformationAction "SilentlyContinue"

                foreach ( $iD in $exchangeGraphPermissionsGuid  ) {
                    try {
                        New-MgServicePrincipalAppRoleAssignedTo `
                        -ServicePrincipalId $newServicePrincipal.Id `
                        -ResourceId $graphApiServicePrincipal.Id `
                        -PrincipalId $newServicePrincipal.Id `
                        -AppRoleId $iD -ErrorAction Stop -InformationAction "SilentlyContinue"| Out-Null
                        
                    }
                    catch {
                        $tempSpDetails = New-Object System.Object
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ServicePrincipalId" -Value $newServicePrincipal.Id
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ResourceId" -Value $graphApiServicePrincipal.Id
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "PrincipalId" -Value $newServicePrincipal.Id
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "AppRoleId" -Value $iD
                        $servicePrincipalAppRoleAssignedRetry.Add($tempSpDetails) | Out-Null
                    }
                    
                }

                

                foreach ( $iD in $exchangeEwsPermissionsGuid  ) {
                    try {

                        New-MgServicePrincipalAppRoleAssignedTo `
                        -ServicePrincipalId $newServicePrincipal.Id `
                        -ResourceId $ewsApiServicePrincipal.Id `
                        -PrincipalId $newServicePrincipal.Id `
                        -AppRoleId $iD -ErrorAction Stop -InformationAction "SilentlyContinue"| Out-Null
                        
                    }
                    catch {
                        $tempSpDetails = New-Object System.Object
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ServicePrincipalId" -Value $newServicePrincipal.Id
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ResourceId" -Value $ewsApiServicePrincipal.Id
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "PrincipalId" -Value $newServicePrincipal.Id
                        $tempSpDetails | Add-Member -MemberType NoteProperty -Name "AppRoleId" -Value $iD
                        $servicePrincipalAppRoleAssignedRetry.Add($tempSpDetails) | Out-Null
                        
                    }
                    
                }


            }
            elseif ($DataSource -eq "OneDrive") {
                Write-Information -Message "Info: Adding the required API Permissions to the Enterprise Application."
                
                try {
                    foreach ( $iD in $oneDriveGraphPermissionsGuid  ) {
                        New-MgServicePrincipalAppRoleAssignedTo `
                        -ServicePrincipalId $newServicePrincipal.Id `
                        -ResourceId $graphApiServicePrincipal.Id `
                        -PrincipalId $newServicePrincipal.Id `
                        -AppRoleId $iD -ErrorAction Stop -InformationAction "SilentlyContinue"| Out-Null
        
                    }
                    
                }
                catch {
                    $tempSpDetails = New-Object System.Object
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ServicePrincipalId" -Value $newServicePrincipal.Id
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ResourceId" -Value $graphApiServicePrincipal.Id
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "PrincipalId" -Value $newServicePrincipal.Id
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "AppRoleId" -Value $iD
                    $servicePrincipalAppRoleAssignedRetry.Add($tempSpDetails) | Out-Null
                    
                }

            }
            elseif ($DataSource -eq "SharePoint") {

                $sharePointApiServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$($sharePointApiAppId)'" -InformationAction "SilentlyContinue"

                try {
                    foreach ( $iD in $sharePointGraphPermissionGuid  ) {
                        New-MgServicePrincipalAppRoleAssignedTo `
                        -ServicePrincipalId $newServicePrincipal.Id `
                        -ResourceId $graphApiServicePrincipal.Id `
                        -PrincipalId $newServicePrincipal.Id `
                        -AppRoleId $iD -InformationAction "SilentlyContinue" | Out-Null
                
                    }
                    
                }
                catch {
                    $tempSpDetails = New-Object System.Object
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ServicePrincipalId" -Value $newServicePrincipal.Id
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ResourceId" -Value $graphApiServicePrincipal.Id
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "PrincipalId" -Value $newServicePrincipal.Id
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "AppRoleId" -Value $iD
                    $servicePrincipalAppRoleAssignedRetry.Add($tempSpDetails) | Out-Null
                    
                }

                try {
                    foreach ( $iD in $sharePointSpointPermissionGuid  ) {
                        New-MgServicePrincipalAppRoleAssignedTo `
                        -ServicePrincipalId $newServicePrincipal.Id `
                        -ResourceId $sharePointApiServicePrincipal.Id `
                        -PrincipalId $newServicePrincipal.Id `
                        -AppRoleId $iD -InformationAction "SilentlyContinue" | Out-Null
                
                    }
                    
                }
                catch {
                    $tempSpDetails = New-Object System.Object
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ServicePrincipalId" -Value $newServicePrincipal.Id
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "ResourceId" -Value $sharePointApiServicePrincipal.Id
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "PrincipalId" -Value $newServicePrincipal.Id
                    $tempSpDetails | Add-Member -MemberType NoteProperty -Name "AppRoleId" -Value $iD
                    $servicePrincipalAppRoleAssignedRetry.Add($tempSpDetails) | Out-Null
                    
                }
                
            }

            Write-Information -Message "Info: Storing the completed Enterprise Application details to memory."
            $tempEntAppDetails = New-Object System.Object
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "AppId" -Value $newEnterpriseApp.AppId
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "Secret" -Value $addPasswordToApp.SecretText
            $tempEntAppDetails | Add-Member -MemberType NoteProperty -Name "DataSource" -Value $DataSource
            $enterpriceApplicationDetails.Add($tempEntAppDetails) | Out-Null

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

    try {
        Write-Information -Message "Info: Getting the Microsoft 365 Subscription name."

        $m365SubscriptionName = (Get-MgOrganization -InformationAction "SilentlyContinue").DisplayName 
    }
    catch {
        
        while ($true) {
            Start-Sleep 5
            $m365SubscriptionName = (Get-MgOrganization -InformationAction "SilentlyContinue").DisplayName 
            if ($m365SubscriptionName){
                break
            } 
        }
    }

    Write-Information -Message "Info: Disconnecting from the Microsoft Graph API"

    Disconnect-Graph
    $staticSleepPeriod = 60
    Write-Information -Message "Info: Waiting $($staticSleepPeriod) seconds to allow the Microsoft database to sync."
    Start-Sleep -Seconds $staticSleepPeriod
    foreach ( $app in $enterpriceApplicationDetails  ) {
       
        $payload = @{
            "operationName" = "AddCustomerO365AppMutation";
            "variables" = @{
                "o365AppType" = $o365AppType[$app.DataSource] ;
                "o365AppClientId" = $app.AppId;
                "o365AppClientSecret" = $app.Secret;
                "o365SubscriptionName" = $m365SubscriptionName;
            };
            "query" = "mutation AddCustomerO365AppMutation(`$o365AppType: String!, `$o365AppClientId: String!, `$o365AppClientSecret: String!, `$o365SubscriptionName: String!) {
                insertCustomerO365App(o365AppType: `$o365AppType, o365AppClientId: `$o365AppClientId, o365AppClientSecret: `$o365AppClientSecret, o365SubscriptionName: `$o365SubscriptionName) {
                    success
                }
            }";
        }
       
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        if ($response.data.insertCustomerO365App.success -eq $true) {
            Write-Host "Successfully added Enterprise Application $($app.AppId) to Rubrik."
        }
        else {
            Write-Host "Error adding Application $($app.AppId) to Rubrik. The error resposne is $($response)."
        }

    }

    return $enterpriceApplicationDetails
    
}
Export-ModuleMember -Function New-EnterpriseApplication






