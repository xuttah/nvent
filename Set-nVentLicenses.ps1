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
Connect-MsolService -Credential $Credential
#
# Define custom SKUs with only critical services
#
$E3SKU = "nventco:ENTERPRISEPACK"
$F1SKU = "nventco:DESKLESSPACK"
$nVentE3Options = New-MsolLicenseOptions -AccountSkuId $E3SKU -DisabledPlans BPOS_S_TODO_2,FORMS_PLAN_E3,STREAM_O365_E3,Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,PROJECTWORKMANAGEMENT,SWAY,INTUNE_O365,YAMMER_ENTERPRISE,RMS_S_ENTERPRISE,OFFICESUBSCRIPTION
$nVentE3DisabledPlans = "BPOS_S_TODO_2,FORMS_PLAN_E3,STREAM_O365_E3,Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,PROJECTWORKMANAGEMENT,SWAY,INTUNE_O365,YAMMER_ENTERPRISE,RMS_S_ENTERPRISE,OFFICESUBSCRIPTION"
$nVentF1Options = New-MsolLicenseOptions -AccountSkuId $F1SKU -DisabledPlans BPOS_S_TODO_FIRSTLINE,FORMS_PLAN_K,STREAM_O365_K,FLOW_O365_S1,POWERAPPS_O365_S1,TEAMS1,SWAY,INTUNE_O365,YAMMER_ENTERPRISE
$nVentF1DisabledPlans = "BPOS_S_TODO_FIRSTLINE,FORMS_PLAN_K,STREAM_O365_K,FLOW_O365_S1,POWERAPPS_O365_S1,TEAMS1,SWAY,INTUNE_O365,YAMMER_ENTERPRISE"
#
# Get all members of AD group for E3 licenses
#
$E3users = Get-ADGroupMember -Identity "NVT-AP-0365-E3"
#
# Assign licenses to any unassigned group members
#

$LicCount = foreach ($E3user in $E3users)
{   ## Get Azure AD users by UPN
    $thisADuser = Get-ADUser $E3user.SamAccountName -Properties userprincipalname,c | where {$_.enabled -eq $true}
    $UserLic = Get-MsolUser -UserPrincipalName $thisADuser.userprincipalname
  
    ## If user has no license, assign one
    If ($UserLic.IsLicensed -eq $false)
    {
        Write-output ("$($thisADuser.userprincipalname) in NVT-AP-0365-E3 group is unlicensed")
        Set-MsolUser -UserPrincipalName $thisADuser.userprincipalname -UsageLocation $thisADuser.c
        Set-MsolUserLicense -UserPrincipalName $thisADuser.userprincipalname -AddLicenses $E3SKU -LicenseOptions $nVentE3Options
        $E3thisresult = new-object PSObject
        $E3thisresult | Add-Member -MemberType NoteProperty -Name "AD Account" -Value $thisADuser.userprincipalname
        $E3thisresult | Add-Member -MemberType NoteProperty -Name "License Applied" -Value $E3SKU
        $E3thisresult | Add-Member -MemberType NoteProperty -Name "Disabled Plans" -Value $nVentE3DisabledPlans
        $Allresults += $E3thisresult
    }
}
Write-output ("Licensed $($LicCount.Count) users for E3.")
#
# Get all members of AD group for F1 licenses
#
$F1users = Get-ADGroupMember -Identity "NVT-AP-0365-F1"

# Assign licenses to any unassigned group members
#
$LicCount = foreach ($F1user in $F1users)
{   ## Get Azure AD users by UPN
    $thisADuser = Get-ADUser $F1user.SamAccountName -Properties userprincipalname,c | where {$_.enabled -eq $true}
    $UserLic = Get-MsolUser -UserPrincipalName $thisADuser.userprincipalname
  
    ## If user has no license, assign one
    If ($UserLic.IsLicensed -eq $false)
    {
        Write-output ("$($thisADuser.userprincipalname) in NVT-AP-0365-F1 group is unlicensed")
        Set-MsolUser -UserPrincipalName $thisADuser.userprincipalname -UsageLocation $thisADuser.c
        Set-MsolUserLicense -UserPrincipalName $thisADuser.userprincipalname -AddLicenses $F1SKU -LicenseOptions $nVentF1Options
        $F1thisresult = new-object PSObject
        $F1thisresult | Add-Member -MemberType NoteProperty -Name "AD Account" -Value $thisADuser.userprincipalname
        $F1thisresult | Add-Member -MemberType NoteProperty -Name "License Applied" -Value $F1SKU
        $F1thisresult | Add-Member -MemberType NoteProperty -Name "Disabled Plans" -Value $nVentF1DisabledPlans
        $Allresults += $F1thisresult
    }
}
Write-output ("Licensed $($LicCount.Count) users for F1.")
Write-Output ("Exporting results to $($OutputFile).")
$Allresults | Export-Csv $OutputFile -NoTypeInformation