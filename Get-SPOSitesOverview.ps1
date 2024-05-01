<#
.SYNOPSIS
This script retrieves an overview of SharePoint Online sites in a Microsoft 365 tenant, including details such as Group ID, Site ID, URL, Title, Storage Used, Last Activity Date, Creation Date, Owner, Owner Email, Archive Status, and Sensitivity Label.

.DESCRIPTION
The script connects to SharePoint Online and fetches details of active sites from a specific internal list used by the admin portal. It also retrieves available sensitivity labels for the tenant and applies these labels to the sites if they have any. The script aims to provide a comprehensive overview of SharePoint Online sites, useful for administrative and auditing purposes.

.PARAMETER Url
The URL of the SharePoint Online admin site to connect to.

.PARAMETER OutputFile
Path and filename of the csv file to generate the report 

.EXAMPLE
PS> .\Get-SPOSitesOverview.ps1 -Url "https://m365x761031-admin.sharepoint.com/" -OutputFile "C:\temp\sitesoverview.csv"

This example runs the script to fetch and display the SharePoint Online sites overview.

.NOTES
Make sure to have the required permissions to access the SharePoint admin site and the necessary lists.

.REQUIRED MODULES
    - PnP.PowerShell (tested with 2.4.0): This module is used to connect to SharePoint Online, fetch lists and list items, and retrieve sensitivity labels.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Url,

    [Parameter(Mandatory=$true)]
    [ValidateScript({-not (Test-Path $_)})]
    [string]$OutputFile
)

# Helper function to write some nicer outpout to the console when logging
function WriteLog() 
{
	# Logging function - Write logging to screen and log file
	param
	(
		[parameter(Mandatory = $true)]
		[System.String]
		$message
	)
	$date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$log = [string]::Format("{0} - {1}", $date, $message)
	Write-Output $log
}

# wondering if this should be part of the script or just a prereq to get the script working (e.g. just have a proper instance of PnP PowerShell setup with auth)
Connect-PnPOnline -url $url -Interactive

# Get available sensivity labels for the tenant, 
# if this fails we only have GUIDs and need to do a manual lookup later
WriteLog -Message "Trying to fetch sensitivity labels"
$labels = Get-PnPAvailableSensitivityLabel | Select-Object Id, Name

# Internal list used to show 'active' sites in the admin portal
WriteLog -Message "Trying to fetch admin list"
$list = Get-PnPList -Identity "DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS"

# Fetching all the items from the 'active sites' list
WriteLog -Message ("Fetching {0} items from list" -f $list.ItemCount.ToString())
$listItems = Get-PnPListItem -List $list -PageSize 1000

WriteLog -Message ("Enumerating through every listitem in the list")
$sites = @{}
foreach ($listItem in $listItems) 
{
    # if a site already has a label applied, let's get the name instead of returning a guid
    if (![string]::IsNullOrEmpty(($listItem["SensitivityLabel"])))
    {
        $sensitivityLabel = $labels | Where-Object { $_.Id -eq $listItem["SensitivityLabel"] }
    }
    else
    {
        $sensitivityLabel = $null
    }

    $site = New-Object PSObject -Property @{
        GroupId = $listItem["GroupId"]
        RelatedGroupId = $listItem["RelatedGroupId"]
        SiteId = $listItem["SiteId"]
        SiteUrl = $listItem["SiteUrl"]
        Title = $listItem["Title"]
        StorageUsedInMB =  ($listItem["StorageUsed"] / 1024 / 1024)
        LastActivityOn = $listItem["LastActivityOn"]
        Created = $listItem["TimeCreated"]
        Owner = $listItem["CreatedByEmail"]
        ArchiveStatus = $listItem["ArchiveStatus"]
        SensitivityLabel = $sensitivityLabel.Name
        ShouldBeArchived = ""
        ShouldHaveLabel = ""
        ShouldHaveLabelId = ""
        ErrorMessage = ""
    }

    Write-Verbose $site
    $sites.Add($site.SiteUrl, $site)
}

# Getting the owners of the Group-connected sites, this may take a while
$updatedSites = @{}

# these variables are used to keep track of the progress
$totalSites = $sites.Count
$currentProgress = 0

# for all sites we have, we're going to fetch the owner if the site is a group-connected site
foreach($siteId in $sites.Keys)
{
    $currentProgress++
    $percentComplete = ($currentProgress / $totalSites) * 100

    Write-Progress -Activity "Getting owners of sites " -Status "$currentProgress of $totalSites processed" -PercentComplete $percentComplete
    
    # get the site from the dictionary
    $site = $sites[$siteId]
    # skip if the site is not a group connected team site
    if ($site.GroupId -eq "00000000-0000-0000-0000-000000000000") 
    { 
        $updatedSites.Add($siteId, $site)
        continue; 
    }
      
    # get the group owner
    try 
    {
        $groupOwner = Get-PnPMicrosoft365GroupOwner -Identity $site.GroupId | Select-Object UserPrincipalName, Email
        # if just 1 owner is returned, it's easy
        if($groupOwner.Count -eq 1)
        {
            $site.Owner = $groupOwner.UserPrincipalName
        } 
        else 
        {
            # otherwise, if it's more than 1, we're going to enumerate and join the userprincipalnames 
            # into a single string divided by a ; character for easy copy/paste into outlook
            $owners = $groupOwner | ForEach-Object { $_.UserPrincipalName }
            $site.Owner = $owners -join "; "
        }
    }
    catch [System.Management.Automation.PSInvalidOperationException] {
        # Handle SharePoint client request exceptions
        Write-Error "Error: $($_.Exception.Message)"
        $site.ErrorMessage = $_.Exception.Message
    }
    catch {
        Write-Error "Error getting group owner for site $siteId - $site.GroupId"
        $site.ErrorMessage = $_.Exception.Message
    }
    $updatedSites.Add($siteId, $site)
}
Write-Progress -Activity "Processing Sites" -Completed

WriteLog -Message "Generating CSV file"
$updatedSites.Values | Export-csv -Path $OutputFile -NoTypeInformation
WriteLog -Message "Done"