<#
    .SYNOPSIS
    This scripts scans all the site collections and subsites in the tenant nd generate a metrics report for every site.
    .DESCRIPTION
    Generate a csv file which contains the information about objects of each site for e.g 
    #of lists in site
    #of libraries in site
    #of folders in a library
    #of files in a library
    #of groups for a site
    .PARAMETER User
    User ID of the SharePoint admin.
    .PARAMETER Password
    Password of the Admin user
    .PARAMETER Credentials
    Credentials of the Admin user, in case the UserId and Password are not supplied.
    .PARAMETER ALL
    Switch to generate the report on all sites
    .PARAMETER CSV
    A list of sites in case the metrics is needed for a subset of sites 
    .PARAMETER AdminURL
    Url of the tenant admin site 
    
    .OUTPUT 
    Two CSV files, One for the libraries in the site and other for the Lists in the sites
    .EXAMPLE
    PS> ./SPOMetrics.ps1 -User <UserID> -Password <Password> -AdminURL <Tenant Admin URL> -ALL 
    #>

[CmdletBinding()]param 
(
    [Parameter(Mandatory=$true, ValueFromPipeline=$false, ParameterSetName="UserPwd", HelpMessage='UserId for the Admin user for SharePoint Online.')]
	[Alias("User")]
	[string]$userid,

    [Parameter(Mandatory=$true, ValueFromPipeline=$false, ParameterSetName="UserPwd" ,HelpMessage='UserId for the Admin user for SharePoint Online.')]
	[Alias("Password")]
	[Security.SecureString]$pwd,
	
    [Parameter(Mandatory=$true, ValueFromPipeline=$false, ParameterSetName="Creds" ,HelpMessage='UserId for the Admin user for SharePoint Online.')]
	[Alias("Credentials")]
	[System.Management.Automation.PSCredential]
    $Creds,

	#[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Set -all true for all sites.')]
	#[Alias("All")]
	#[switch]$allSites = $false,


	[Parameter(Mandatory=$true, ValueFromPipeline=$false, HelpMessage='Set -all true for all sites.')]
	[Alias("AdminURL")]
	[string]$TenantAdminURL,

    [Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='CSV Files of the sites that you want to process.')]
	[Alias("csv")]
	[string]$fileCSV

)
# Plugin
Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
Import-Module SharePointPnPPowerShellOnline -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
Import-Module CredentialManager -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

Function Get-ListsFromSite() {
    write-host "In Get-ListsFromSite"
    $lists=Get-PnPList
    $lists
    return $lists
}

Function Get-GroupsFromSite($site) {
    $site
    $groups=Get-PnPGroup
    $groups

    return $groups
}

Function WriteToFile($siteUrl,$listCount)
{
    [PSCustomObject]@{
       SiteURL = $siteUrl
       ListCount = $listCount 
       } | Export-Csv metrics.csv -notype -Append 
}

Function Get-ItemcountInList($lst)
{
    $items=(Get-PnPListItem -List $lst -Fields "Title").FieldValues
    $count=$items.Length
    return $count
}

Function GetMetricsForsite($url) 
{
    $Error=""
    $blnConnect=0
    Write-host "Processing Site :$url"
    Try
    {
        Connect-PnPOnline -Url $url -Credentials $Credential
        $blnConnect=1
        $currentsite=Get-PnPSite
    }
    Catch
    {
        write-host "Count not connect"
        $blnConnect=0
    }
    #write-host $blnConnect
    if($blnConnect)
    {
        write-host "Connected"
        $listsInSite=Get-PnPList
        $AllListsInSiteInfo=@()
        $AllLibInSiteInfo=@()
        $SiteInfo=@()
        $SiteInfoHeaderList=@()
        $SiteInfoHeaderLib=@()
        $SiteInfoForLib=@()
        $SiteInfoForList=@()
        $ListCount=$listsInSite.Length
        $listsInSite | ForEach-Object { $list=$_
            $listUrl=$TenantURL+$list.RootFolder.ServerRelativeUrl
            $lstType=$list.BaseType
            if ($lstType -eq "GenericList")
            {
                Try
                {
                    write-host "Processing List : "$listUrl
                    $itemCount=$list.ItemCount
                    $ListInfo           =  [pscustomobject]@{SiteURL ="";"Group count"="";"storageUsed(inMB)"="";"List"=$listUrl;"Items" = $itemCount}
                }
                catch
                {
                    $ListInfo           =  [pscustomobject]@{SiteURL ="";"Group count"="";"storageUsed(inMB)"="";"List"=$listUrl;"Items" = "Error Occured"}
                }    
                $AllListsInSiteInfo+=$ListInfo
            }
            else
            {
                Try
                {
                    write-host "Processing Library : "$list.RootFolder.ServerRelativeUrl
                    $arr=$list.RootFolder.ServerRelativeUrl.Split("/")
                    $Rootfolder=$arr[$arr.length-1]#$list.RootFolder.ServerRelativeUrl.Split("/")[3].Tostring()                     
                    if($list.IsCatalog)
                    {
                        $Rootfolder="_catalogs/"+$Rootfolder
                    }
                    $allItems=Get-PnPFolderItem -FolderSiteRelativeUrl $Rootfolder -Recursive -ErrorAction Stop
                    if($Error[0])
                    {  
                        $LibInfo = [pscustomobject]@{SiteURL ="";"Group count"="";"Lib"=$listUrl;"Folders"="Error Occured";"Files"="Error Occured";"storageUsed(inMB)"=""}
                    }
                    else
                    {
                        #write-host $allItems 
                        if($allItems)
                        {
                            $allFolders = $allItems | Where-Object {$_.TypedObject.ToString() -eq "Microsoft.SharePoint.Client.Folder"} 
                            $allFiles   = $allItems | Where-Object {$_.TypedObject.ToString() -eq "Microsoft.SharePoint.Client.File"} 
                            if(!$allFolders)
                            {
                                $folderCount=0
                            }
                            else
                            {
                                if (!$allFolders.Length)
                                {
                                    $folderCount=1
                                }
                                else
                                {
                                    $folderCount=$allFolders.Length
                            }
                            }
                            if(!$allFiles)
                            {
                                $filesCount=0
                            }
                            else
                            {
                                if (!$allFiles.Length)
                                {
                                    $filesCount=1
                                }
                                else
                                {
                                    $filesCount=$allFiles.Length
                                }
                            }
                            $LibInfo = [pscustomobject]@{SiteURL ="";"Group count"="";"Lib"=$listUrl;"Folders"=$folderCount;"Files"=$filesCount;"storageUsed(inMB)"=""}
                        }
                        else
                        {
                            $LibInfo = [pscustomobject]@{SiteURL ="";"Group count"="";"Lib"=$listUrl;"Folders"=0;"Files"=0;"storageUsed(inMB)"=""}
                        }
                    }
                }
                catch
                {
                    Write-host "Error occured"    
                    $LibInfo = [pscustomobject]@{SiteURL ="";"Group count"="";"Lib"=$listUrl;"Folders"="Error Occured";"Files"="Error Occured";"storageUsed(inMB)"=""}
                }    
                $AllLibInSiteInfo+=$LibInfo
            }
        }
        Try
        {
            write-host "Getting Groups for site"
            $groupsInSite=Get-PnPGroup -ErrorAction Stop
            $GroupCount=$groupsInSite.Length
            write-host $GroupCount 
        }
        Catch
        {
            $ErrorMessage = $_.Exception.Message
            Write-host $ErrorMessage
            #Write-host "Error in Group count"
            $GroupCount="Error Occured"
        }
        $SiteInfoHeaderList += [pscustomobject]@{SiteURL = $url ;"Group count"=$GroupCount  ;"storageUsed(inMB)"=$sitecollectionStorage ;"List"="";"Items"=""}
        $SiteInfoHeaderLib += [pscustomobject]@{SiteURL = $url;"Group count"=$GroupCount;List="";"storageUsed(inMB)"=$sitecollectionStorage;"Lib"="";"Folders"="";"Files"=""}
    }
    else
    {
        $SiteInfoHeaderList += [pscustomobject]@{SiteURL = $url;"Group count"="Could not Connect";"storageUsed(inMB)"="Could not Connect"}
        $SiteInfoHeaderLib += [pscustomobject]@{SiteURL = $url;"Group count"="Could not Connect";"storageUsed(inMB)"="Could not Connect"}

    }
    $SiteInfoForLib+=$SiteInfoHeaderLib
    if($AllLibInSiteInfo)
    {
        $SiteInfoForLib+=$AllLibInSiteInfo
    }
    $SiteInfoForList+=$SiteInfoHeaderList
    if($AllListsInSiteInfo)
    {
        $SiteInfoForLib+=$AllLibInSiteInfo
    }


    [hashtable]$SiteInformation = @{}
    $SiteInformation.Libs=$SiteInfoForLib
    $SiteInformation.Lists=$SiteInfoForList
    return $SiteInformation
}

Function Get-Metrics($sites)
{
    $AllSitesLists = @()
    $AllSitesLibs = @()
    $Libs = New-Object -TypeName System.Collections.ArrayList
    $GroupCount=0
    $ListCount=0
    $sites | ForEach-Object { $sposite=$_
        $SiteUrl=$sposite.Url
        $sitecollectionStorage=$sposite.StorageUsageCurrent
            Try
            {
                $currentsiteURL=$SiteUrl
                write-host "______________________________________________________"
                #write-host $currentsiteURL
                $siteMetric=GetMetricsForsite($currentsiteURL)
                $AllSitesLists  +=$siteMetric.Lists
                $AllSitesLibs   +=$siteMetric.Libs
                #$Libs.Add($siteMetric.Libs) | Out-null
                Try
                {
                    write-host "Get subsites"
                    $subsites=Get-PnPSubWebs -Recurse -ErrorAction Stop
                    if($subsites)
                    {
                        write-host "Found subsites"
                    }
                    else
                    {
                        write-host "No subsites"
                    }
                    if ($subsites.Length -gt 0){
                        $subsites | ForEach-Object {$spoSubSite=$_
                            $currentsiteURL=$spoSubSite.Url
                            write-host "***********************************************"
                            #write-host $currentsiteURL
                            $siteMetric=GetMetricsForsite($currentsiteURL)
                            $AllSitesLists  +=$siteMetric.Lists
                            $AllSitesLibs   +=$siteMetric.Libs
                            $Libs.Add($siteMetric.Libs) | Out-null

                        }
                    }
                    write-host "______________________________________________________"

                }
                catch
                {
                    write-host "Error getting Subsites"
                }
            }
            Catch
            {
                $ErrorMessage = $_.Exception.Message
                $AllSites += [pscustomobject]@{SiteURL = $currentsiteURL;ListCount = "Could not Connect";"storageUsed(inMB)"="Could not Connect"}
            }
    }
    $AllSitesLibs  | Export-Csv -Path ".\SPMetricsLibs.csv" -NoTypeInformation
    $AllSitesLists | Export-Csv -Path ".\SPMetricsLists.csv" -NoTypeInformation
    write-host "Finished"

}

if (!$Creds)
{
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userid, $pwd
}
else
{
    $Credential=$Creds
}
$TenantUrl=$TenantAdminURL.Replace("-admin","")
Connect-SPOService -Url $TenantAdminURL -Credential $Credential
#if($AllSites)
#{
    $AllSiteCollections=get-sposite -Limit All  | Select Url,Title,StorageUsageCurrent| Sort-Object StorageUsageCurrent -Descending
#}
#write-host $AllSiteCollections
Get-Metrics($AllSiteCollections)




    