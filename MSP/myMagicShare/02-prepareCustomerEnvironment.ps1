#========================================================================
# Header
#========================================================================
#region header
#========================================================================
# Created on:		04.10.2022 13:37
# Created by:		Andreas Hähnel
# Organization:		Black Magic Cloud
# Script name:		
# Script Version: 	0.2
#========================================================================
# RequiredPermissions:
# at least "SharePoint Administrator" permissions
# 
# Modules:
# "Microsoft.Online.SharePoint.PowerShell"
# "PnP.PowerShell"
#========================================================================
# Description:
#
#========================================================================
# Useful links / infos:
#
# create site via graph api 
# https://learn.microsoft.com/en-us/sharepoint/dev/apis/site-creation-rest
#
#========================================================================
# Changelog:
# Version 0.1 04.10.2022
# - initial creation
# - creation of Admin (Config) Site
# - creation of User Site
# - creation of necessary columns
#
#========================================================================
# Prerequisites and how to use this script:
<#
1. create azure app with client secret and certificate
2. have account with at least "SharePoint Administrator" permissions
3. check global variables
4. start script
5. design mMS-Config home.aspx according your wishes
6. design mMS home.aspx according your wishes
    - welcome text
    - explanation / how to use the share
    - request share button
8. clean up mms navigation
#>
#========================================================================
# EXAMPLE (replace variables with actual values):
# download the script
# .\Prepare-EnvironmentformyMagicShare.ps1
#========================================================================
#endregion

#========================================================================
# Global Variables
#========================================================================
#region global variables

# adjust these values to your environment:
###########################################################
$mmsOwner = "admin@contoso.onmicrosoft.com"  # Owner of the myMagicShare Site(s)
$tenant = "contoso.onmicrosoft.com" # name of your tenant
$appId = "00000000-0000-0000-0000-000000000000" # AppID, you've created recently
$clientSecret = "abcdefghijklmnopqrstuvwxyz" #clientSecret of your appID
$tenantId = "00000000-0000-0000-0000-000000000000" # your tenant ID
$mmsAdminSiteTitle = "myMagicShare Admin" #the displayName of the mMS admin Site in SPO
$mmsUserSiteTitle = "myMagicShare" #the displayName of the mMS user Site in SPO
$logFilePath = "C:\_temp\" # logfile path  #$logFilePath = $PSScriptRoot
$lfn = "_LOG_prepareMyMagicShare" #no extension needed
#$durationValues = "1","3","7","14" # available share durations in days
$mmsAdminSiteQuota = 4096 # Quota for admin site in mb -> 4gb
$mmsUserSiteQuota = 1024000 #Quota for user site in mb -> 1tb
$sharingcapability = 3 # recommended value = 3 -> anonymous sharing active ### 0 -> disabled || 1 -> only existing guests || 2 -> new guests can be invited || 3 -> anonymous links can be created
$configureSPOTenantforAnyOneLinks = $true  # if SPO doesn't allow anonymous sharing for SPO, should the script enable that?
###########################################################

# DO.NOT.TOUCH anything after this line (or there will be no support)!
###########################################################
$mmsAdminSite = "myMagicShare-config"
$mmsUserSite = "myMagicShare"
$tenantURL = "https://"+ $tenant.Split(".")[0] +"-admin.sharepoint.com"
$mmsAdminSiteURL = "https://"+ $tenant.Split(".")[0] +".sharepoint.com/sites/"+ $mmsAdminSite
$mmsUserSiteURL = "https://"+ $tenant.Split(".")[0] +".sharepoint.com/sites/"+ $mmsUserSite
$moduleNames = @(
    "Microsoft.Online.SharePoint.PowerShell"
    "PnP.PowerShell"
)
switch ( $sharingcapability ) {
    0 { $sharingcapabilityToSet = 'Disabled' }
    1 { $sharingcapabilityToSet = 'ExistingExternalUserSharingOnly' }
    2 { $sharingcapabilityToSet = 'ExternalUserSharingOnly' }
    3 { $sharingcapabilityToSet = 'ExternalUserAndGuestSharing' }
}

#endregion

#========================================================================
# Functions
#========================================================================
#region functions

function Get-GraphAuthorizationToken {
    param
    (
        [string]$ResourceURL = 'https://graph.microsoft.com',
        [string][parameter(Mandatory)]
        $TenantID,
        [string][Parameter(Mandatory)]
        $ClientKey,
        [string][Parameter(Mandatory)]
        $AppID
    )
	
    #$Authority = "https://login.windows.net/$TenantID/oauth2/token"
	$Authority = "https://login.microsoftonline.com/$TenantID/oauth2/token"
	
    [Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null
    $EncodedKey = [System.Web.HttpUtility]::UrlEncode($ClientKey)
	
    $body = "grant_type=client_credentials&client_id=$AppID&client_secret=$EncodedKey&resource=$ResourceUrl"
	
    # Request a Token from the graph api
    $script:result = Invoke-RestMethod -Method Post -Uri $Authority -ContentType 'application/x-www-form-urlencoded' -Body $body
	
    $script:APIHeader = @{ 'Authorization' = "Bearer $($result.access_token)" }
}

#========================================================================
function Normalize-String {
    param(
        [Parameter(Mandatory = $true)][string]$str
    )
	
    $str = $str.ToLower()
    $str = $str.Replace(" ", "")
    $str = $str.Replace("ä", "ae")
    $str = $str.Replace("ö", "oe")
    $str = $str.Replace("ü", "ue")
    $str = $str.Replace("ß", "ss")
	
    Write-Output $str
}

#========================================================================
function write-log {
    Param (
        [parameter(Mandatory=$true)][String]$logFileName,
        [parameter(Mandatory=$false)]$note,
        [parameter(Mandatory=$false)]$mErr
    )

    $tStamp = get-date -Format HH:mm:ss

    if ($error) {
        if ($mErr -eq 1) {
            $mErrOut = $tStamp+" | ERR | "+$note
            Write-Host $mErrOut -ForegroundColor Red
            $mErrOut | Out-File -FilePath $logFilePath"\"$logFileName -Append -Force -NoClobber

            $errOut = $tStamp+" | ERR | "+$error.exception.message
            Write-Host $errOut
            $errOut | Out-File -FilePath $logFilePath"\"$logFileName -Append -Force -NoClobber

            $mErr = 0
        }
        else {
            $errOut = $tStamp+" | ERR | "+$error.exception.message
            Write-Host $errOut
            $errOut | Out-File -FilePath $logFilePath"\"$logFileName -Append -Force -NoClobber
        }
    }
    elseif ($mErr -eq 1) {
        $mErrOut = $tStamp+" | ERR | "+$note
        Write-Host $mErrOut -ForegroundColor Red
        $mErrOut | Out-File -FilePath $logFilePath"\"$logFileName -Append -Force -NoClobber

        $mErr = 0
    }
    else {
        $note = $tStamp+" | INF | "+$note
        Write-Host $note
        $note | Out-File -FilePath $logFilePath"\"$logFileName -Append -Force -NoClobber
    }

    Clear-Variable -Name "note"
    $error.clear()
}

#========================================================================
function set-execPolicy {
    $execPolicy = Get-ExecutionPolicy

    switch ( $execPolicy ) {
        'Restricted'    { $execPolicyInt = 0 }
        'AllSigned'     { $execPolicyInt = 1 }
        'Default'       { $execPolicyInt = 2 }
        'Undefined'     { $execPolicyInt = 3 }
        'RemoteSigned'  { $execPolicyInt = 4 }
        'Unrestricted'  { $execPolicyInt = 5 }
        'Bypass'        { $execPolicyInt = 6 }
    }

    $note = "Setting executionpolicy..."
    write-log -logFileName $lfn -note $note

    try {
        if ($execPolicyInt -lt 4) {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
            $note = "Executionpolicy set to <RemoteSigned> for current user."
            write-log -logFileName $lfn -note $note
        }
        else {
            $note = "Executionpolicy already more permissive than required ... nothing to do."
            write-log -logFileName $lfn -note $note
        }
    }
    catch {
        $note = "Executionpolicy could not be set!"
        write-log -logFileName $lfn -note $note
        Exit
    }
}

#========================================================================
function install-module {
    Param (
        [parameter(Mandatory=$true)]$moduleNames
    )
    
    $installedModules = Get-InstalledModule
    $iMs = @()
    $installedModules | ForEach-Object {$iMs += $_.Name}

    foreach ($moduleName in $moduleNames) {
        if ($moduleName -notin $iMs) {
            $note = "$moduleName not found - installing..."
            write-log -logFileName $lfn -note $note
            Install-Module -Name $moduleName -Force -WarningAction Ignore
            #check-module -moduleName $moduleName
        }
        else {
            $note = "$moduleName already installed"
            write-log -logFileName $lfn -note $note
            #Update-Module -Name $moduleName -Force -WarningAction Ignore
        }
		check-module -moduleName $moduleName
    }
}

#========================================================================
function check-module {
    Param (
        [parameter(Mandatory=$true)][String]$moduleName
    )

    try {
        Import-Module -Name $moduleName -WarningAction Ignore
        $note = "$moduleName successfully imported."
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "$moduleName could not be imported... exiting!"
        write-log -logFileName $lfn -note $note
        Exit
    }
}

#========================================================================
function connect-spo {
    Param (
        [parameter(Mandatory=$true)][String]$url
    )
    
    $note = "Trying to connect to SharePoint Online..."
    write-log -logFileName $lfn -note $note
    try {
        $spoCon = Connect-SPOService -Url $tenantURL
        $note = "Connection successfully established!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "No connection could be established... exiting!"
        write-log -logFileName $lfn -note $note
        Exit
    }
}

#========================================================================
function connect-pnpSite {
    Param (
        [parameter(Mandatory=$true)][String]$url
    )
    $note = "Trying to connect to SharePoint Online Site (via PNP)..."
    write-log -logFileName $lfn -note $note
    try {
        $script:pnpCon = Connect-PnPOnline -Url $url -Interactive -ReturnConnection:$true
        $note = "PNP connection successfully established!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "No PNP connection could be established... exiting!"
        write-log -logFileName $lfn -note $note
        Exit
    }
}

#========================================================================
function check-spoTenant {
    $note = "Checking tenant settings..."
    write-log -logFileName $lfn -note $note
    $tenant = Get-SPOTenant

    if (!$tenant) {
        $note = "Tenant details could not be retreived... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $tsc = $tenant.SharingCapability
    switch ( $tsc ) {
        'Disabled'                          { $tscInt = 0 }
        'ExistingExternalUserSharingOnly'   { $tscInt = 1 }
        'ExternalUserSharingOnly'           { $tscInt = 2 }
        'ExternalUserAndGuestSharing'       { $tscInt = 3 }
    }
    try {
        if ($tscInt -lt $sharingcapability) {
            $note = "Sharing capabilities are not set to accomodate myMagicShare prerequisites - correcting..."
            write-log -logFileName $lfn -note $note
            Set-SPOTenant -SharingCapability $sharingcapabilityToSet
        }
        else {
            $note = "Sharing capabilities are already set to accomodate myMagicShare prerequisites."
            write-log -logFileName $lfn -note $note
        }
    }
    catch {
        $note = "Sharing capabilities could not be set... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }
}

#========================================================================
function create-sposite {
    Param (
        [parameter(Mandatory=$true)][String]$url,
        [parameter(Mandatory=$true)][String]$owner,
        [parameter(Mandatory=$true)][String]$title,
        [parameter(Mandatory=$true)][Int]$storageQuota
    )

    $note = "Trying to create SPO Site: $title"
    write-log -logFileName $lfn -note $note
    try {
        $newSite = New-SPOSite -Url $url -Owner $owner -Title $title -Template "STS#3" -LocaleId 1033 -TimeZoneId 4 -StorageQuota $storageQuota
        $note = "Site successfully created!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Site $title could not be created... exiting!"
        write-log -logFileName $lfn -note $note
        Exit
    }
}

#========================================================================
function check-sposite {
    Param (
        [parameter(Mandatory=$true)][String]$url
    )

    $note = "Checking, if $url is created..."
    write-log -logFileName $lfn -note $note

    $siteToCheck = Get-SPOSite -Identity $url -ErrorAction SilentlyContinue

    while (!$siteToCheck) {
        $note = "... still checking ..."
        write-log -logFileName $lfn -note $note
        $siteToCheck = Get-SPOSite -Identity $url -ErrorAction SilentlyContinue
    }
    $note = "$url is created!"
    write-log -logFileName $lfn -note $note

    $error.Clear()
}

#========================================================================
function configure-sposite {
    Param (
        [parameter(Mandatory=$true)][String]$url,
        [parameter(Mandatory=$true)][String]$title
    )

    $note = "Configuring site '$title'"
    write-log -logFileName $lfn -note $note
    try{
        if ($url -eq $mmsAdminSiteURL) {
            $configSite = Set-SPOSite -Identity $url -SharingCapability "Disabled" -StorageQuotaWarningLevel ($mmsAdminSiteQuota * 0.8)

            $note = "Configuration of $title was successful!"
            write-log -logFileName $lfn -note $note
        }
        elseif ($url -eq $mmsUserSiteURL) {
            $configSite = Set-SPOSite -Identity $url -SharingCapability $sharingcapabilityToSet -DisableFlows "Disabled" -DefaultSharingLinkType "Direct" -StorageQuotaWarningLevel ($mmsUserSiteQuota * 0.8)
            
            $note = "Configuration of $title was successful!"
            write-log -logFileName $lfn -note $note
        }
        else {
            $note = "No valid URL available... check manually!"
            write-log -logFileName $lfn -note $note -mErr 1
        }
    }
    catch {
        $note = "Configuration of $title was NOT successful... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }
}

#========================================================================
function cleanup-navigation {
    Param (
        [parameter(Mandatory=$true)]$connection
    )
    
    $note = "Cleaning navigation nodes..."
    write-log -logFileName $lfn -note $note

    $navNodes = Get-PnPNavigationNode -Location QuickLaunch -Connection $connection
    foreach ($navNode in $navNodes) {
        if ($navNode.Title -ne "Home") {
            try {
                $note = "Removing $($navNode.Title)..."
                write-log -logFileName $lfn -note $note
                Remove-PnPNavigationNode -Identity $navNode.Id -Force -Connection $connection
                $note = "Success!"
                write-log -logFileName $lfn -note $note
            }
            catch {
                $note = "Error removing $($navNode.Title) from navigation!"
                write-log -logFileName $lfn -note $note
            }
        }
    }
}

#========================================================================
function add-navigation {
    Param (
        [parameter(Mandatory=$true)]$connection,
        [parameter(Mandatory=$true)][String]$title,
        [parameter(Mandatory=$true)][String]$url
    )
    
    try {
        $note = "Adding navigation note $title..."
        write-log -logFileName $lfn -note $note
        Add-PnPNavigationNode -Location QuickLaunch -Title $title -Url $url -Connection $connection
        $note = "Success!"
        write-log -logFileName $lfn -note $note
    }
    catch{
        $note = "Error adding $title to navigation!"
        write-log -logFileName $lfn -note $note
    }
}

#========================================================================
function check-spolist {
    Param (
        [parameter(Mandatory=$true)]$newList,
        [parameter(Mandatory=$true)]$connection
    )

    $listToCheck = Get-PnPList -Connection $connection -ErrorAction SilentlyContinue | Where-Object {$_.DefaultViewUrl -contains $newList.webUrl}

    $note = "Checking, if $($newList.name) is created..."
    write-log -logFileName $lfn -note $note
    while (!$listToCheck) {
        $note = "... still checking ..."
        write-log -logFileName $lfn -note $note
        $listToCheck = Get-PnPList -Connection $connection -ErrorAction SilentlyContinue | Where-Object {$_.DefaultViewUrl -contains $newList.webUrl}
    }
    $note = "$($newList.name) is created!"
    write-log -logFileName $lfn -note $note

    $error.Clear()
}

#endregion

#========================================================================
# Scriptstart
#========================================================================
#region script start
if($Error){$Error.Clear()}
clear

Write-Host "###################################################" -ForegroundColor Cyan
Write-Host "# Scriptstart" -ForegroundColor Cyan
Write-Host "###################################################" -ForegroundColor Cyan

# initialize the log file
if($Error){$error.clear()}
$note = "Starting the process to create myMagicShare!"
write-log -logFileName $lfn -note $note

# set executionpolicy to remote signed for current user to make sure the script can run
set-execPolicy
if($Error){
    $note = "Execution Policy could not be set! Please set it to remoteSigned and re-run the script if script does not run correctly!"
    write-log -logFileName $lfn -note $note -mErr 1
    $error.clear()
}

# check for sharepoint online management shell and connect
install-module -moduleNames $moduleNames
connect-spo -url $tenantURL
if($configureSPOTenantforAnyOneLinks){
    check-spoTenant # set SPO Tenant to support "anyone links"
} 

# create mMS admin site
create-sposite -url $mmsAdminSiteURL -owner $mmsOwner -title $mmsAdminSiteTitle -storageQuota $mmsAdminSiteQuota

# create mMS user site
create-sposite -url $mmsUserSiteURL -owner $mmsOwner -title $mmsUserSiteTitle -storageQuota $mmsUserSiteQuota

# configure mMS admin site
check-sposite -url $mmsAdminSiteURL
configure-sposite -url $mmsAdminSiteURL -title $mmsAdminSiteTitle

connect-pnpSite -url $mmsAdminSiteURL
$pnpAdminSiteCon = $pnpCon

cleanup-navigation -connection $pnpAdminSiteCon

# configure mMS user site
check-sposite -url $mmsUserSiteURL
configure-sposite -url $mmsUserSiteURL -title $mmsUserSiteTitle

connect-pnpSite -url $mmsUserSiteURL
$pnpUserSiteCon = $pnpCon

cleanup-navigation -connection $pnpUserSiteCon

# create & configure lists in MyMagicShare-Config
$note = "Getting Graph Access Token..."
write-log -logFileName $lfn -note $note
try {
    Get-GraphAuthorizationToken -TenantID $tenant -ClientKey $clientSecret -AppID $appId
    if ($result.access_token) {
        $note = "Token successfully retreived!"
        write-log -logFileName $lfn -note $note
    }
    else {
        $note = "Graph Token could not be retreived... exiting!"
        write-log -logFileName $lfn -note $note
        Exit
    }
}
catch {
    $note = "Error while retreiving Graph Token... exiting!"
    write-log -logFileName $lfn -note $note
    Exit
}

#region Admin Site: MyMagicShare-Config

<# mms-Config "Configuration"
- create list (Configuration) with following columns
    -Key (Title - default column)
    -Value (single line of text)
#>
$error.clear()

$note = "Creating Admin configuration list in admin site..."
write-log -logFileName $lfn -note $note
try {
    $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists"
    $bodyHashTable = @{
        "displayName" = "Configuration"
        "list" = @{
            "template" = "genericList"
        }
    }
    $newList = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
    $note = "Admin Configuration List created successfully"
    write-log -logFileName $lfn -note $note
}
catch {
    $note = "Admin Configuration List could not be created... check manually!"
    write-log -logFileName $lfn -note $note -mErr 1
}

if (!$error) {
    check-spolist -newList $newList -connection $pnpAdminSiteCon

    $note = "Creating column 'VALUE' in new list..."
    write-log -logFileName $lfn -note $note

    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            name = "Value"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $false
                appendChangesToExistingText = $false
                linesForEditing = 0
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Column successfully created"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'VALUE' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }
}

add-navigation -title $newList.name -url $newList.webUrl -connection $pnpAdminSiteCon

<# mms-Config "Log"
- create list (Log) for history with following columns
    -ShareName (Title - default column)
    -RequestTimestamp (single line of text)
    -TTL (days -> number)
    -RequesterDisplayName (single line of text)
    -RequesterEmail (single line of text)
    -Action (eg create / active / delete -> single line of text)
    -LogText (various stuff -> multiline)
#>
$error.clear()

$note = "Creating Log list on admin site..."
write-log -logFileName $lfn -note $note
try {
    $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists"
    $bodyHashTable = @{
        "displayName" = "Log"
        "list" = @{
            "template" = "genericList"
        }
    }
    $newList = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
    $note = "List successfully created!"
    write-log -logFileName $lfn -note $note
}
catch {
    $note = "Log List could not be created... check manually!"
    write-log -logFileName $lfn -note $note -mErr 1
}

if (!$error) {
    check-spolist -newList $newList -connection $pnpAdminSiteCon

    $note = "Creating RequestTimestamp column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            name = "RequestTimestamp"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $false
                appendChangesToExistingText = $false
                linesForEditing = 0
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Column successfully created!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'RequestTimestamp' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $note = "Creating TTL column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            name = "TTL"
            number = @{
                decimalPlaces = "none"
                displayAs = "number"
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Column successfully created!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'TTL' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $note = "Creating RequestorDisplayname column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            "name" = "RequesterDisplayName"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $false
                appendChangesToExistingText = $false
                linesForEditing = 0
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Successfully created new column!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'RequestorDisplayname' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $note = "Creating RequesterEmail column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            "name" = "RequesterEmail"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $false
                appendChangesToExistingText = $false
                linesForEditing = 0
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Successfully created new column!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'RequesterEmail' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $note = "Creating Action column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            "name" = "Action"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $false
                appendChangesToExistingText = $false
                linesForEditing = 0
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Successfully created new column!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'Action' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $note = "Creating LogText column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            name = "LogText"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $true
                appendChangesToExistingText = $true
                linesForEditing = 5
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Column successfully created!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'LogText' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }
}

add-navigation -title $newList.name -url $newList.webUrl -connection $pnpAdminSiteCon

<# mms-Config "ActiveShares"
- create list (ActiveShares) for current active shares with following columns
    -ShareName (Title - default column)
    -RequesterDisplayName (single line of text)
    -RequesterEmail (single line of text)
    -ShareURL (single line of text)
    -ExoirationDate (single line of text)
    -TTL (Number)
#>
$error.clear()

$note = "Creating ActiveShares list on admin site..."
write-log -logFileName $lfn -note $note
try {
    $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists"
    $bodyHashTable = @{
        "displayName" = "ActiveShares"
        "list" = @{
            "template" = "genericList"
        }
    }
    $newList = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
    $note = "List successfully created!"
    write-log -logFileName $lfn -note $note
}
catch {
    $note = "List 'ActiveShares' could not be created... check manually!"
    write-log -logFileName $lfn -note $note -mErr 1
}

if (!$error) {
    check-spolist -newList $newList -connection $pnpAdminSiteCon

    $note = "Creating RequestorDisplayname column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            "name" = "RequesterDisplayName"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $false
                appendChangesToExistingText = $false
                linesForEditing = 0
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Successfully created new column!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'RequesterDisplayName' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $note = "Creating RequesterEmail column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            "name" = "RequesterEmail"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $false
                appendChangesToExistingText = $false
                linesForEditing = 0
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Successfully created new column!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'RequesterEmail' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $note = "Creating ShareURL column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            name = "ShareURL"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $false
                appendChangesToExistingText = $false
                linesForEditing = 0
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Successfully created new column!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'ShareURL' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $note = "Creating ExpirationDate column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            name = "ExpirationDate"
            enforceUniqueValues = $false
            hidden = $false
            indexed = $false
            text = @{
                allowMultipleLines = $false
                appendChangesToExistingText = $false
                linesForEditing = 0
                maxLength = 255
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Successfully created new column!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'ExpirationDate' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    $note = "Creating TTL column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsAdminSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            name = "TTL"
            number = @{
                decimalPlaces = "none"
                displayAs = "number"
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Successfully created new column!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'TTL' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }
}

add-navigation -title $newList.name -url $newList.webUrl -connection $pnpAdminSiteCon

#endregion

#region User Site: MyMagicShare
<# mms "Queue"
- create list (Queue) for share requests with following columns
    -ShareName (Title - default column)
    -Duration (choice)
#>
$error.clear()

$note = "Creating list 'QUEUE' on user site..."
write-log -logFileName $lfn -note $note
try {
    $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsUserSite):/lists"
    $bodyHashTable = @{
        "displayName" = "Queue"
        "list" = @{
            "template" = "genericList"
        }
    }
    $newList = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
    $note = "List successfully created!"
    write-log -logFileName $lfn -note $note
}
catch {
    $note = "List 'QUEUE' could not be created... check manually!"
    write-log -logFileName $lfn -note $note -mErr 1
}

if (!$error) {
    check-spolist -newList $newList -connection $pnpUserSiteCon

    $note = "Creating Duration column in new list..."
    write-log -logFileName $lfn -note $note
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$($tenant.Split(".")[0]).sharepoint.com:/sites/$($mmsUserSite):/lists/$($newList.id)/columns"
        $bodyHashTable = @{
            name = "Duration"
            choice = @{
                allowTextEntry = $false
                choices = $durationValues
                displayAs = "dropDownMenu"
            }
        }
        $newCol = Invoke-RestMethod -Uri $uri -Method POST -Headers $script:APIHeader -Body ($bodyHashTable | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
        $note = "Successfully created new column!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Column 'Duration' could not be created... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }
}

add-navigation -title $newList.name -url $newList.webUrl -connection $pnpUserSiteCon

# add everyone to visitors & change permissions on queue list for visitors
$note = "Getting visitors group..."
write-log -logFileName $lfn -note $note
try {
    $visitors = Get-PnPGroup -Connection $pnpUserSiteCon | ? {$_.Title -ilike "*visitors*"}
    $note = "Success!"
    write-log -logFileName $lfn -note $note
}
catch {
    $note = "Visitors group could not be retreived... check manually!"
    write-log -logFileName $lfn -note $note -mErr 1
}

if ($visitors) {
    $note = "Generating <everyone except external> login name..."
    write-log -logFileName $lfn -note $note
    $loginName = "c:0-.f|rolemanager|spo-grid-all-users/"+$tenantId
    $note = "Login name = $loginName"
    write-log -logFileName $lfn -note $note

    $note = "Adding everyone except external to visitors group..."
    write-log -logFileName $lfn -note $note
    try {
        Add-PnPGroupMember -Group $visitors -LoginName $loginName -Connection $pnpUserSiteCon
        $note = "Success!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Could not add everyone except externals to visitors... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }
}

$note = "Getting <Queue> list..."
$qList = Get-PnPList -Identity $newList.name -Connection $pnpUserSiteCon
if ($qList) {
    $note = "Success!"
    write-log -logFileName $lfn -note $note

    $note = "Breaking inheritance in list..."
    write-log -logFileName $lfn -note $note
    try {
        $breakInheritance = Set-PnPList -Identity $qList.Id -BreakRoleInheritance:$true -Connection $pnpUserSiteCon
        $note = "Success!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Error... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }

    write-log -logFileName $lfn -note "Adding visitors with contribute to list..."
    try {
        $visitorsContribute = Set-PnPListPermission -Identity $qList.Id -Group $visitors -AddRole "Contribute" -Connection $pnpUserSiteCon
        $note = "Success!"
        write-log -logFileName $lfn -note $note
    }
    catch {
        $note = "Error... check manually!"
        write-log -logFileName $lfn -note $note -mErr 1
    }
}
else {
    $note = "List could not be retreived... check manually!"
    write-log -logFileName $lfn -note $note -mErr 1
}

$note = "Everything is set and done - enjoy!"
write-log -logFileName $lfn -note $note

#endregion

#endregion