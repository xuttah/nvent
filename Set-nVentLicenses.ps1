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

$users = import-csv .\nVent_Users.csv
foreach ($user in $users)
{
    $upn=$user.UPN
    $usagelocation=$user.country 
    $SKU=$user.SKU
    Set-MsolUser -UserPrincipalName $upn -UsageLocation $usagelocation
    Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $SKU
} 