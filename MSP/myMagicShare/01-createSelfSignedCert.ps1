#requires -modules PnP.PowerShell
#requires -Version 5

param(
[Parameter(Mandatory=$false)][String]$CommonName = "myMagicShareCertificate",
[Parameter(Mandatory=$false)][String]$country = "DE",
[Parameter(Mandatory=$false)][String]$state = "Bavaria",
[Parameter(Mandatory=$false)][String]$locality = "Regensburg",
[Parameter(Mandatory=$false)][String]$organization = "Black Magic Cloud",
[Parameter(Mandatory=$false)][String]$organizationUnit = "Development",
[Parameter(Mandatory=$false)][String]$OutPfx = "D:\tmp\mms\mmsCert2.pfx",
[Parameter(Mandatory=$false)][String]$OutCert = "D:\tmp\mms\mmsCert2.cer",
[Parameter(Mandatory=$false)][String]$CertPasswordFile = "D:\tmp\mms\certpass.txt"
)
function Generate-Password {
    param (
        [Parameter(Mandatory)]
        [int] $length,
        [int] $amountOfNonAlphanumeric = 1
    )
    Add-Type -AssemblyName 'System.Web'
    return [System.Web.Security.Membership]::GeneratePassword($length, $amountOfNonAlphanumeric)
}

Import-Module PnP.PowerShell
write-host "commonname: $CommonName"
$certpass = (Generate-Password -length 12 -amountOfNonAlphanumeric 0)

$certprops = @{
    CommonName = $CommonName
    Country = $country
    State = $state
    Locality = $locality
    Organization = $organization
    OrganizationUnit = $organizationUnit
    ValidYears = 99
    OutPfx = $OutPfx
    OutCert = $OutCert
    CertificatePassword = (ConvertTo-SecureString -String $certpass -AsPlainText -Force) 
}

$cert = New-PnPAzureCertificate @certprops

Write-Host "Certificate Password: $certpass"
$certpass | Out-File $CertPasswordFile
