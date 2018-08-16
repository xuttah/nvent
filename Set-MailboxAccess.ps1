<#
.SYNOPSIS
    The script will automatically assign mailbox and recipient permissions on shared mailboxes based on groups.
.EXAMPLE
   .\SharedMailboxViaGroups.ps1 -Prefix 'SM-'
.PARAMETER Prefix
    Prefix of the groups that will manage permissions on the shared mailboxes.
#>
 
function Connect-ExchangeOnline {
  # NEEDS STORED CREDENTIAL!!!!!!
  #Creates an Exchange Online session
  $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credential -Authentication Basic -AllowRedirection
  #Import session commands
  Import-PSSession $ExchangeSession
  #Connect to MS Online
  Connect-MSOLService -credential $credential
}
function Add-SharedMailboxPermission {
  param(
    [string]$Identity,
    [string]$SharedMailboxName
  )
  try {
    Add-MailboxPermission -Identity $SharedMailboxName -User $Identity -AccessRights FullAccess -ErrorAction stop | Out-Null
    Add-RecipientPermission -Identity $SharedMailboxName -Trustee $Identity -AccessRights SendAs -Confirm:$False -ErrorAction stop | Out-Null
    Write-Output "INFO: Successfully added $Identity to $SharedMailboxName"
  } catch {
    Write-Warning "Cannot add $Identity to $SharedMailboxName`r`n$_"
  }
}
function Remove-SharedMailboxPermission {
  param(
    [string]$Identity,
    [string]$SharedMailboxName
  )
  try {
    Remove-MailboxPermission -Identity $SharedMailboxName -User $Identity -AccessRights FullAccess -Confirm:$False -ErrorAction stop -WarningAction ignore | Out-Null
    Remove-RecipientPermission -Identity $SharedMailboxName -Trustee $Identity -AccessRights SendAs -Confirm:$False -ErrorAction stop -WarningAction ignore  | Out-Null
    Write-Output "INFO: Successfully removed $Identity from $SharedMailboxName"
  } catch {
    Write-Warning "Cannot remove $Identity from $SharedMailboxName`r`n$_"
  }
}
function Sync-EXOResourceGroup {
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [string]$Prefix = '!SR '
  )
  #Get All groups to process mailboxes for
  $MasterGroups = Get-Group -ResultSize Unlimited -Identity "$Prefix*"
  foreach ($Group in $MasterGroups) {
    #Remove prefix to get the mailbox name
    $MbxName_woPrefix = $Group.Name.Replace("$Prefix",'')
	If ($MbxName_woPrefix -match " FA" ) {
		$MbxName = $MbxName_woPrefix.Replace(" FA",'')
	}
	Else {
		$MbxName = $MbxName_woPrefix.Replace(" SA",'')
	}
    $SharedMailboxName =  (Get-Mailbox -Identity $MbxName -ErrorAction ignore -WarningAction ignore).WindowsLiveID
    if ($SharedMailboxName) { 
      Write-Verbose -Message "Processing group $($Group.Name) and mailbox $SharedMailboxName"
      #Get all users with explicit permissions on the mailbox
      $SharedMailboxDelegates = Get-MailboxPermission -Identity $SharedMailboxName -ErrorAction Stop -ResultSize Unlimited | Where-Object {$_.IsInherited -eq $false -and $_.User -ne "NT AUTHORITY\SELF" -and $_.User -notlike "!SR*"} | Select-Object @{Name="User";Expression={(Get-Mailbox $_.User).UserPrincipalName}}
      #Get all group members
      $SharedMailboxMembers = Get-DistributionGroupMember -Identity $Group.Identity -ResultSize Unlimited
      #Remove users if group is empty
      if (-not($SharedMailboxMembers) -and $SharedMailboxDelegates) {
        Write-Warning "The group $Group is empty, will remove explicit permissions from $SharedMailboxName"
        foreach ($user in $SharedMailboxDelegates.User) {
          Remove-SharedMailboxPermission -Identity $user -SharedMailboxName $SharedMailboxName
        }
        #Add users if no permissions are present
      } elseif (-not($SharedMailboxDelegates)) {
        foreach ($user in $SharedMailboxMembers.WindowsLiveID) {
          Add-SharedMailboxPermission -Identity $user -SharedMailboxName $SharedMailboxName
        }
        #Process removals and adds
      } else {
        #Compare the group with the users that have actual access
        $Users = Compare-Object -ReferenceObject $SharedMailboxDelegates.User -DifferenceObject $SharedMailboxMembers.WindowsLiveID 
          
        #Add users that are members of the group but do not have access to the shared mailbox
        foreach ($user in ($users | Where-Object {$_.SideIndicator -eq "=>"})) {
          Add-SharedMailboxPermission -Identity $user.InputObject -SharedMailboxName $SharedMailboxName
        }
        #Remove users that have access to the shared mailbox but are not members of the group
        foreach ($user in ($users | Where-Object {$_.SideIndicator -eq "<="})) {
          Remove-SharedMailboxPermission -Identity $user.InputObject -SharedMailboxName $SharedMailboxName
        }
      }
    } else {
      Write-Warning "Could not find the mailbox $MbxName"
    }
  }
}
#Connect to Exchange Online
Connect-ExchangeOnline
#Start Processing groups and mailboxes
[string]$Prefix = '!SR '
Sync-EXOResourceGroup -Prefix $Prefix -Verbose