Connect-PnPOnline -Url "https://tenant.sharepoint.com"

# query to fetch PHL document which are larger than 10mb
# $query = "(contentclass:STS_ListItem_1310) AND (size > 10485760)"

#query to fetch all PHL documents
$query = "(contentclass:STS_ListItem_1310)"

$selectProperties = "listid, path, size, spsiteurl, spweburl, filetype, filename, compliancetag, fileextension,path, author, lastmodifiedtime, LastModifiedTimeForRetention"
$results = Submit-PnPSearchQuery -selectproperties $selectproperties -query $query -All

# this holds all the output
$holdLibraryItems = New-Object System.Collections.Generic.List"[PSCustomObject]"

foreach($resultRow in $results.ResultRows)
{
    $holdLibraryItem = New-Object PSObject -Property @{
        FileName = $resultRow["filename"]
        FileExtension = $resultRow["fileextension"]
        Path = $resultRow["path"]
        SizeInMB =  ($resultRow["size"] / 1024 /1024 )
        Compliancetag = $resultRow["compliancetag"]
        SiteUrl =  $resultRow["spsiteurl"]
        LastModifiedDate = $resultRow["lastmodifiedtime"]
        Author = $resultRow["author"]
        LastModifiedTimeForRetention = $resultRow["LastModifiedTimeForRetention"]
    }
    
    $holdLibraryItems.Add($holdLibraryItem)
}
