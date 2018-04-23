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
# Constants
[string]$FileAppend = (Get-Date -Format mmddyyyy_) + (Get-Random -Maximum 9999)
$OutputFile = "C:\scripts\license_output_" + $FileAppend + ".csv"
$Username = "S0001920000000@nventco.onmicrosoft.com"
$PasswordPath = "c:\scripts\securepass.txt"
$Allresults = @()
#
# Read the password from the file and convert to SecureString
#
Write-Host "Getting password from $Passwordpath"
$SecurePassword = Get-Content $PasswordPath | ConvertTo-SecureString
#
# Build a Credential Object from the password file and the $username constant
#
$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePassword
#
# Destroy any outstanding PS Session
#
Get-PSSession | Remove-PSSession -Confirm:$false
#
# Sleep 15s to allow the sessions to tear down fully
Write-Output ("Sleeping 15 seconds for Session Tear Down")
Start-sleep -seconds 15
#
#Connect to Office 365
Import-Module MSOnline
Import-Module ActiveDirectory
Connect-MsolService
#
# Define custom SKUs with only critical services
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
        $E3thisresult = new-object PSObject
        $E3thisresult | Add-Member -MemberType NoteProperty -Name "AD Account" -Value $E3User.UserPrincipalName
        $E3thisresult | Add-Member -MemberType NoteProperty -Name "License Applied" -Value $E3SKU
        $Allresults += $E3thisresult
    }
}
Write-output "Licensed $($LicCount.Count) users for E3."
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
        $F1thisresult = new-object PSObject
        $F1thisresult | Add-Member -MemberType NoteProperty -Name "AD Account" -Value $F1User.UserPrincipalName
        $F1thisresult | Add-Member -MemberType NoteProperty -Name "License Applied" -Value $F1SKU
        $Allresults += $F1thisresult
    }
}
Write-output "Licensed $($LicCount.Count) users for F1."
Write-Output "Exporting results to $($OutputFile)."
$Allresults | Export-Csv $OutputFile -NoTypeInformation