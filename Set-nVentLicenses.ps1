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
Connect-MsolService

#Build nVent custom SKUs
###### Connect to Office 365 GOES HERE ######
# Define E3 custom SKU with only critical services
$nVentE3Options = New-MsolLicenseOptions -AccountSkuId nventco:ENTERPRISEPACK -DisabledPlans BPOS_S_TODO_2,FORMS_PLAN_E3,STREAM_O365_E3,Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,PROJECTWORKMANAGEMENT,SWAY,INTUNE_O365,YAMMER_ENTERPRISE,RMS_S_ENTERPRISE
$nVentF1Options = New-MsolLicenseOptions -AccountSkuId nventco:ENTERPRISEPACK -DisabledPlans #list sub-SKUs
# Assign E3 SKUs
$users = import-csv .\nVent_Users.csv
foreach ($user in $users)
{
    $upn=$user.UPN
    $usagelocation=$user.country 
    If ($user.SKU = "E3"){
        Set-MsolUser -UserPrincipalName $upn -UsageLocation $usagelocation
        Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $SKU -LicenseOptions $nVentE3Options
    }
    If ($user.SKU = "F1"){
        Set-MsolUser -UserPrincipalName $upn -UsageLocation $usagelocation
        Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $SKU -LicenseOptions $nVentF1Options
    }

} 