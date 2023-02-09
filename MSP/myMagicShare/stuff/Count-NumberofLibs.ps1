$SPOSitePath = "https://MYTENANT.sharepoint.com"
$connection = Connect-PnPOnline -Url "$($SPOSitePath)/sites/myMagicShare" -Interactive -ReturnConnection:$true
if(!$connection){Write-Host "Error connecting to PnP" -ForegroundColor red;exit}

$lists = Get-PnPList
$LibCount = 0
foreach ($list in $lists) {
  if (($list.BaseType -eq "GenericList") -or ($list.BaseType -eq "DocumentLibrary") ) {
    $LibCount++
  }
}

Write-Host "Number of libraries in the site collection: $($LibCount)" -ForegroundColor Green
