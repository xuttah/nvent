<#
.SYNOPSIS
  Automated license management of Office 365 users from accounts in Active Directory
.DESCRIPTION
  N/A
.PARAMETER <Parameter_Name>
  N/A
.INPUTS 
  None
.OUTPUTS
  None
.NOTES
  Version:        2.0
  Author:         Kevin Clark
  Creation Date:  4/24/2018
  Purpose/Change: Implement functions, Un-license logic
  
.EXAMPLE
  N/A
#>
function Set-UnlicensedUsers {
    #
    # Get all members of AD group for E3 licenses
    #
    $E3GroupMembers = Get-ADGroupMember -Identity "NVT-AP-0365-E3"
    #
    # Assign licenses to any unassigned group members
    #

    $E3LicCount = foreach ($E3user in $E3GroupMembers)
    {   ## Get Azure AD users by UPN
        $thisADuser = Get-ADUser $E3user.SamAccountName -Properties userprincipalname,c | where {$_.enabled -eq $true}
        $UserLic = Get-MsolUser -UserPrincipalName $thisADuser.userprincipalname | %{($_.Licenses).AccountSkuID}
      
        ## If user has no license, assign one
        If ($UserLic.IsLicensed -eq $false)
        {
            Write-output ("$($thisADuser.userprincipalname) in NVT-AP-0365-E3 group is unlicensed")
            Set-MsolUser -UserPrincipalName $thisADuser.userprincipalname -UsageLocation $thisADuser.c
            Set-MsolUserLicense -UserPrincipalName $thisADuser.userprincipalname -AddLicenses $MFASKU,$E3SKU -LicenseOptions $nVentE3Options
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
    $F1GroupMembers = Get-ADGroupMember -Identity "NVT-AP-0365-F1"

    # Assign licenses to any unassigned group members
    #
    $F1LicCount = foreach ($F1user in $F1GroupMembers)
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
    $E3GroupMembers = $null
    $E3user = $null
    $UserLic = $null
}
##
function Remove-licensedUsers {
#
    # Get all members of AD group for E3 licenses
    $E3GroupMembers = Get-ADGroupMember -Identity "NVT-AP-0365-E3"
    # Get all users in Office 365 with E3 licenses
    $E3LicensedUsers = Get-MsolUser -all -Synchronized | Where-Object {($_.licenses).AccountSkuId -match "EnterprisePack"}
    #
    # Check for any synced E3-licensed user in Office 365 that isn't a member of the E3 AD Group
    #

    $LicCount = foreach ($E3user in $E3GroupMembers)
    {   ## Get Azure AD users by UPN
        $thisADuser = Get-ADUser $E3user.SamAccountName -Properties userprincipalname,c | where {$_.enabled -eq $true}
        $UserLic = Get-MsolUser -UserPrincipalName $thisADuser.userprincipalname | %{($_.Licenses).AccountSkuID}
      
        ## If user an E3 license, remove it
        If ($UserLic.IsLicensed -eq $false)
        {
            Write-output ("$($thisADuser.userprincipalname) in NVT-AP-0365-E3 group is unlicensed")
            Set-MsolUser -UserPrincipalName $thisADuser.userprincipalname -UsageLocation $thisADuser.c
            Set-MsolUserLicense -UserPrincipalName $thisADuser.userprincipalname -AddLicenses $MFASKU,$E3SKU -LicenseOptions $nVentE3Options
            $E3thisresult = new-object PSObject
            $E3thisresult | Add-Member -MemberType NoteProperty -Name "AD Account" -Value $thisADuser.userprincipalname
            $E3thisresult | Add-Member -MemberType NoteProperty -Name "License Applied" -Value $E3SKU
            $E3thisresult | Add-Member -MemberType NoteProperty -Name "Disabled Plans" -Value $nVentE3DisabledPlans
            $Allresults += $E3thisresult
        }
    }
}

#
# Constants
[string]$FileAppend = (Get-Date -Format mmddyyyy_) + (Get-Random -Maximum 9999)
$OutputFile = "C:\scripts\set-unlicensedusers_output_" + $FileAppend + ".csv"
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
$MFASKU = "nventco:MFA_STANDALONE"
$nVentE3Options = New-MsolLicenseOptions -AccountSkuId $E3SKU -DisabledPlans BPOS_S_TODO_2,FORMS_PLAN_E3,STREAM_O365_E3,Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,PROJECTWORKMANAGEMENT,SWAY,INTUNE_O365,YAMMER_ENTERPRISE,RMS_S_ENTERPRISE,OFFICESUBSCRIPTION
$nVentE3DisabledPlans = "BPOS_S_TODO_2,FORMS_PLAN_E3,STREAM_O365_E3,Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,PROJECTWORKMANAGEMENT,SWAY,INTUNE_O365,YAMMER_ENTERPRISE,RMS_S_ENTERPRISE,OFFICESUBSCRIPTION"
$nVentF1Options = New-MsolLicenseOptions -AccountSkuId $F1SKU -DisabledPlans BPOS_S_TODO_FIRSTLINE,FORMS_PLAN_K,STREAM_O365_K,FLOW_O365_S1,POWERAPPS_O365_S1,TEAMS1,SWAY,INTUNE_O365,YAMMER_ENTERPRISE
$nVentF1DisabledPlans = "BPOS_S_TODO_FIRSTLINE,FORMS_PLAN_K,STREAM_O365_K,FLOW_O365_S1,POWERAPPS_O365_S1,TEAMS1,SWAY,INTUNE_O365,YAMMER_ENTERPRISE"
##
##
## License Unlicensed Users
Set-UnlicensedUsers
#

