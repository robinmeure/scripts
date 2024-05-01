<#
.SYNOPSIS
This scripts takes an input file (created by Get-SPOSitesOverview.ps1) and archives the sites marked to be archived.

.DESCRIPTION
This scripts takes an input file (created by Get-SPOSitesOverview.ps1) and archives the sites marked to be archived.

.PARAMETER Url
The URL of the SharePoint Online admin site to connect to.

.PARAMETER InputFile
Path and filename of the csv file that holds all the processed sites needed to be archived

.PARAMETER OutputFile
Path and filename of the csv file to generate the report for sites that have been archived

.EXAMPLE
PS> .\Set-SPOSiteToArchive.ps1 -Url "https://m365x761031-admin.sharepoint.com/" -InputFile "C:\temp\sitesoverview.csv" -OutputFile "C:\temp\archivedsites.csv"

.NOTES
Make sure to have the required permissions to access the SharePoint admin site and the necessary lists.

.REQUIRED MODULES
    - PnP.PowerShell (tested with 2.4.0): This module is used to connect to SharePoint Online, fetch lists and list items, and retrieve sensitivity labels.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Url,

    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_})]
    [string]$InputFile,

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


Connect-PnPOnline -url $url -Interactive

WriteLog -Message "Trying to fetch sites from csv file"
$sites = Import-Csv -Path $InputFile

WriteLog -Message ("{0} sites found in csv file" -f $sites.Count.ToString())
$sitesToBeArchived = $sites | Where-Object { $_.ShouldBeArchived -eq $true}

WriteLog -Message ("{0} sites found to be archived in csv file" -f $sitesToBeArchived.Count.ToString())
$sitesArchived = @{}

$totalSites = $sitesToBeArchived.Count
$currentProgress = 0

foreach ($site in $sitesToBeArchived) 
{
    $currentProgress++
    $percentComplete = ($currentProgress / $totalSites) * 100

    Write-Progress -Activity "Archiving" -Status "$currentProgress of $totalSites processed" -PercentComplete $percentComplete
    
    try {
        Set-PnPSiteArchiveState -Identity $site.SiteUrl -ArchiveState Archived -Force
    }
    catch {
        Write-Error "Error setting label for site $site.SiteUrl"
        $site.ErrorMessage += $_.Exception.Message
    }
   
    $sitesArchived.Add($site.SiteUrl, $site)
}
Write-Progress -Activity "Processing Sites" -Completed

WriteLog -Message "Generating CSV file"
$sitesArchived.Values | Export-csv -Path $OutputFile -NoTypeInformation
WriteLog -Message "Done"