$SPOTenantPath = "https://MYTENANT.sharepoint.com"
$SPOSiteName = "myMagicShare"
$connection = Connect-PnPOnline -Url "$($SPOTenantPath)/sites/$($SPOSiteName)" -Interactive -ReturnConnection:$true
if(!$connection){Write-Host "Error connecting to PnP" -ForegroundColor red;exit}

$lists = Get-PnPList -Connection $connection
$LibCount = 0
foreach ($list in $lists) {
  if (($list.BaseType -eq "GenericList") -or ($list.BaseType -eq "DocumentLibrary") ) {
    $LibCount++
  }
}

Write-Host "Number of libraries in the site collection: $($LibCount)" -ForegroundColor Green
