
function Get-PolarisM365SharePoint() {
    <#
    .SYNOPSIS
    Returns a filtered list of O365 SharePoint sites, document libraries and/or lists for a given subscription in a given Polaris account.
    
    .DESCRIPTION
    Returns a filtered list of Office 365 SharePoint sites, document libraries and/or lists from a given subscription and Polaris account, taking
    an API token, Polaris URL, subscription ID, and search string.
    
    .PARAMETER SubscriptionId
    The Polaris subscription ID for a given O365 subscription. Can be obtained with the
    'Get-PolarisM365Subscriptions' command.
    
    .PARAMETER SearchString
    Search string, used to filter site, document library or list name.
    
    .PARAMETER Includes
    It indicates if the returned object includes only SharePoint sites, document libraries, lists or a combination of them. The value can
    only be 'SitesOnly', 'DocumentLibrariesOnly', or 'ListsOnly'.
    if it's not specified, it returns all the SharePoint objects including sites, document libraries and lists by default.
    
    .INPUTS
    None. You cannot pipe objects to Get-PolarisM365SharePoint.
   
    .OUTPUTS
    System.Object. Get-PolarisM365SharePoint returns an array containing the ID, Name,
    and SLA details for the returned O365 SharePoint sites, document libraries and/or lists.
    
    .EXAMPLE
    PS> Get-PolarisM365SharePoint -SubscriptionId $my_sub.id -Includes 'SitesOnly' -SearchString 'test'
    name                   : Milan Kundera
    id                     : 12341234-1234-1234-abcd-123456789012
    type                   : O365Site
    slaAssignment          : Direct
    effectiveSlaDomainName : Gold
    #>

    param(
        [Parameter(Mandatory=$True)]
        [String]$SubscriptionId,
        [Parameter(Mandatory=$false)]
        [String]$SearchString,
        [Parameter(Mandatory=$false)]
        [ValidateSet("SitesOnly", "DocumentLibrariesOnly", "ListsOnly")]
        [String]$Includes,
        [String]$Token = $global:RubrikPolarisConnection.accessToken,
        [String]$PolarisURL = $global:RubrikPolarisConnection.PolarisURL
    )

    $headers = @{
        'Content-Type' = 'application/json';
        'Accept' = 'application/json';
        'Authorization' = $('Bearer '+$Token);
    }

    $endpoint = $PolarisURL + '/api/graphql'

    $o365Sites = $null
    $o365SharepointDrives = $null
    $o365SharepointLists = $null

    $payload = @{
        "query"         = "";
        "variables"     = @{
            "after"     = $null;
            "filter"    = @(
                @{
                    "field" = "IS_RELIC";
                    "texts" = @("false");
                },
                @{
                    "field" = "NAME_OR_EMAIL_ADDRESS";
                    "texts" = @($SearchString);
                }
            );
            "first"     = 100;
            "o365OrgId" = $SubscriptionId;
        }
    }

    $querySites = "o365Sites(after: `$after, o365OrgId: `$o365OrgId, filter: `$filter, first: `$first, sortBy: `$sortBy, sortOrder: `$sortOrder) {
        nodes {
            id
            name
            parentId
            title
            url
            hierarchyLevel
            effectiveSlaDomain {
                id
                name
            }
            objectType
            slaAssignment
        }
        pageInfo {
            endCursor
            hasNextPage
            hasPreviousPage
        }
    }"

    $queryDrives = "o365SharepointDrives(after: `$after, o365OrgId: `$o365OrgId, filter: `$filter, first: `$first, sortBy: `$sortBy, sortOrder: `$sortOrder) {
        nodes {
            id
            naturalId
            name
            parentId
            totalStorageInBytes
            usedStorageInBytes
            effectiveSlaDomain {
                id
                name
            }
            objectType
            slaAssignment
            onDemandSnapshotCount
        }
        pageInfo {
            endCursor
            hasNextPage
            hasPreviousPage
        }
    }"

    $queryLists = "o365SharepointLists(after: `$after, o365OrgId: `$o365OrgId, filter: `$filter, first: `$first, sortBy: `$sortBy, sortOrder: `$sortOrder) {
           nodes {
               id
               naturalId
               name
               parentId
               effectiveSlaDomain {
                   id
                   name
               }
               objectType
               slaAssignment
               onDemandSnapshotCount
           }
           pageInfo {
               endCursor
               hasNextPage
               hasPreviousPage
           }
        }"

    if ($Includes -eq "SitesOnly") {
        $payload.query = "query O365SharepointQuery(`$after: String, `$o365OrgId:UUID!, `$filter: [Filter!], `$first: Int!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            $($querySites)
        }"
 
        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $o365Sites = $response.data.o365Sites
    } elseif ($Includes -eq "DocumentLibrariesOnly") {
        $payload.query = "query O365SharepointQuery(`$after: String, `$o365OrgId:UUID!, `$filter: [Filter!], `$first: Int!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            $($queryDrives)
        }"

        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $o365SharepointDrives = $response.data.o365SharepointDrives
    } elseif ($Includes -eq "ListsOnly") {
        $payload.query = "query O365SharepointQuery(`$after: String, `$o365OrgId:UUID!, `$filter: [Filter!], `$first: Int!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            $($queryLists)
        }"

        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $o365SharepointLists = $response.data.o365SharepointLists
    } else {
        $payload.query = "query O365SharepointQuery(`$after: String, `$o365OrgId:UUID!, `$filter: [Filter!], `$first: Int!, `$sortBy: HierarchySortByField, `$sortOrder: HierarchySortOrder) {
            $($querySites)
            $($queryDrives)
            $($queryLists)
        }"

        $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
        $o365Sites = $response.data.o365Sites
        $o365SharepointDrives = $response.data.o365SharepointDrives
        $o365SharepointLists = $response.data.o365SharepointLists
    }

     $node_array += @()
     $sharepoint_details = @()

     if ($null -ne $o365Sites) {
        $node_array += $o365Sites.nodes
        while ($o365Sites.pageInfo.hasNextPage) {
            $payload.variables.after = $o365Sites.pageInfo.endCursor
            $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
            $o365Sites = $response.data.o365Sites
            $node_array += $o365Sites.nodes
        }
    }

    if ($null -ne $o365SharepointDrives) {
        $node_array += $o365SharepointDrives.nodes
        while ($o365SharepointDrives.pageInfo.hasNextPage) {
            $payload.variables.after = $o365SharepointDrives.pageInfo.endCursor
            $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
            $o365SharepointDrives = $response.data.o365SharepointDrives
            $node_array += $o365SharepointDrives.nodes
        }
    }

    if ($null -ne $o365SharepointLists) {
        $node_array += $o365SharepointLists.nodes
        while ($o365SharepointLists.pageInfo.hasNextPage) {
            $payload.variables.after = $o365SharepointLists.pageInfo.endCursor
            $response = Invoke-RestMethod -Method POST -Uri $endpoint -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers
            $o365SharepointLists = $response.data.o365SharepointLists
            $node_array += $o365SharepointLists.nodes
        }
    }

    foreach ($node in $node_array) {
        $row = '' | Select-Object name,id,type,slaAssignment,effectiveSlaDomainName
        $row.name = $node.name
        $row.id = $node.id
        $row.type = $node.objectType
        $row.slaAssignment = $node.slaAssignment
        $row.effectiveSlaDomainName = $node.effectiveSlaDomain.name
        $sharepoint_details += $row
    }

    return $sharepoint_details
}
Export-ModuleMember -Function Get-PolarisM365SharePoint

