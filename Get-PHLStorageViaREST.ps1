# first connection is done to get all the sites
Connect-PnPOnline -Url https://tenant.sharepoint.com -Credentials $credentials
$sites = Get-PnPTenantSite

# next, this is list that will hold all the PHL information
$holdLibraryItems = New-Object System.Collections.Generic.List"[PSCustomObject]"

# looping through all the sites (can take a while)
foreach($site in $sites)
{
    # connect to each site
    Connect-PnPOnline -url $site.Url -credentials $credentials
    
    # call the REST API to retrieve the size of the library
    $output = Invoke-PnPSPRestMethod -Url "/_api/web/getFolderByServerRelativeUrl('preservationholdlibrary')?`$select=StorageMetrics&`$expand=StorageMetrics" -ErrorAction SilentlyContinue
    if ($output -ne $null)
    {
        $holdLibraryItem = New-Object PSObject -Property @{
            SiteUrl = $site.Url
            SizeInMB = ($output.StorageMetrics.TotalSize /1024 / 1024)
            }
        $holdLibraryItems.Add($holdLibraryItem)
    }
}
