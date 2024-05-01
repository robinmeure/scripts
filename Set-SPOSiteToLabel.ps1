<#
.SYNOPSIS
This scripts takes an input file (created by Get-SPOSitesOverview.ps1) and labels the sites with the sensitivity label specified in the input file.

.DESCRIPTION
The script connects to SharePoint Online and fetches details of active sites from a specific internal list used by the admin portal. It also retrieves available sensitivity labels for the tenant and applies these labels to the sites if they have any. The script aims to provide a comprehensive overview of SharePoint Online sites, useful for administrative and auditing purposes.

.PARAMETER InputFile
Path and filename of the csv file that holds all the processed sites needs to be labeled

.PARAMETER OutputFile
Path and filename of the csv file to generate the report for sites that have been labeled

.EXAMPLE
PS> .\Set-SPOSiteToLabel.ps1 -InputFile C:\temp\m365_input.csv -OutputFile "C:\temp\m365_labeled.csv"

This example runs the script to fetch and display the SharePoint Online sites overview.

.NOTES
Make sure that there is connection (Connect-SPOService) being made to the admin site before running this script.

.REQUIRED MODULES
    - Microsoft.Online.SharePoint.PowerShell (tested with 2.4.0): This module is used to connect to SharePoint Online, fetch lists and list items, and retrieve sensitivity labels.
#>

param(
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


WriteLog -Message "Trying to fetch sites from csv file"
$sites = Import-Csv -Path $InputFile

WriteLog -Message ("{0} sites found in csv file" -f $sites.Count.ToString())
$sitesToBeLabeled = $sites | Where-Object { $_.ShouldHaveLabel -ne ""}

WriteLog -Message ("{0} sites found to be archived in csv file" -f $sitesToBeLabeled.Count.ToString())
$totalSites = $sitesToBeLabeled.Count
$currentProgress = 0

$sitesLabeled = @{}
foreach ($site in $sitesToBeLabeled) 
{
    $currentProgress++
    $percentComplete = ($currentProgress / $totalSites) * 100

    Write-Progress -Activity "Labeling" -Status "$currentProgress of $totalSites processed" -PercentComplete $percentComplete
    
    try {
        Set-SPOSite -Identity $site.SiteUrl -SensitivityLabel $site.ShouldHaveLabelId
    }
    catch {
        Write-Error "Error setting label for site $site.SiteUrl"
        $site.ErrorMessage += $_.Exception.Message
    }

    $sitesLabeled.Add($site.SiteUrl, $site)
}
Write-Progress -Activity "Processing Sites" -Completed

WriteLog -Message "Generating CSV file"
$sitesLabeled.Values | Export-csv -Path $OutputFile -NoTypeInformation
WriteLog -Message "Done"