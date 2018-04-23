<#
.SYNOPSIS
  License Office 365 users from accounts in CSV file
.DESCRIPTION
  N/A
.PARAMETER <Parameter_Name>
  N/A
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Kevin Clark
  Creation Date:  2/26/2018
  Purpose/Change: Initial script development
  
.EXAMPLE
  N/A
#>

#Connect to Office 365
Import-Module MSOnline
Import-Module ActiveDirectory
Connect-MsolService

#Build nVent custom SKUs
###### Connect to Office 365 GOES HERE ######
# Define E3 custom SKU with only critical services
$nVentE3Options = New-MsolLicenseOptions -AccountSkuId nventco:ENTERPRISEPACK -DisabledPlans BPOS_S_TODO_2,FORMS_PLAN_E3,STREAM_O365_E3,Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,PROJECTWORKMANAGEMENT,SWAY,INTUNE_O365,YAMMER_ENTERPRISE,RMS_S_ENTERPRISE
$nVentF1Options = New-MsolLicenseOptions -AccountSkuId nventco:ENTERPRISEPACK -DisabledPlans #list sub-SKUs
#
# Get all members of AD group for E3 licenses
#
$E3users = Get-ADGroupMember -Identity "NVT-AP-0365-E3"
#
# Define E3 SKU
#
$E3SKU = "nventco:ENTERPRISEPACK"
#
# Assign licenses to any unassigned group members
#
$LicCount = foreach ($E3user in $E3users)
{   ## Get Azure AD users by UPN
    $UserLic = Get-MsolUser -UserPrincipalName $E3User.UserPrincipalName
  
    ## If user has no license, assign one
    If ($UserLic.IsLicensed -eq $false)
    {
        Write-output "$($E3User.UserPrincipalName) in NVT-AP-0365-E3 group is unlicensed"
        $UserCountry = Get-ADUser $E3user.SamAccountName -Properties c | Select-Object -ExpandProperty c
        Set-MsolUser -UserPrincipalName $E3User.UserPrincipalName -UsageLocation $UserCountry
        Set-MsolUserLicense -UserPrincipalName $E3User.UserPrincipalName -AddLicenses $E3SKU -LicenseOptions $nVentE3Options
    }
}
#
# Get all members of AD group for F1 licenses
#
$F1users = Get-ADGroupMember -Identity "NVT-AP-0365-F1"
#
# Define F1 SKU
#
$F1SKU = "nventco:DESKLESSPACK"
#
# Assign licenses to any unassigned group members
#
$LicCount = foreach ($F1user in $F1users)
{   ## Get Azure AD users by UPN
    $UserLic = Get-MsolUser -UserPrincipalName $F1User.UserPrincipalName
  
    ## If user has no license, assign one
    If ($UserLic.IsLicensed -eq $false)
    {
        Write-output "$($F1User.UserPrincipalName) in NVT-AP-0365-F1 group is unlicensed"
        $UserCountry = Get-ADUser $F1user.SamAccountName -Properties c | Select-Object -ExpandProperty c
        Set-MsolUser -UserPrincipalName $F1User.UserPrincipalName -UsageLocation $UserCountry
        Set-MsolUserLicense -UserPrincipalName $F1User.UserPrincipalName -AddLicenses $F1SKU -LicenseOptions $nVentF1Options
    }
}