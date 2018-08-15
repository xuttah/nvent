# Set-MailboxAccess
# This script parses membership of a group (inclusive recursive membership) and
# assigns mailbox permissions to group members in Exchange Online.  This is done to 
# ensure automapping functionality is retained (not available when assigning permission to group, only members)
#
Import-Module MSOnline
Import-Module ActiveDirectory
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credential -Authentication Basic -AllowRedirection
#Import session commands
Import-PSSession $ExchangeSession
#Connect to MS Online
Connect-MSOLService -credential $credential

param([string]$Group,[string]$Mailbox)

# Helper function to get group members recursively
# This is from Steve Goodman 
# http://www.stevieg.org/2010/12/report-exchange-mailboxes-group-members-full-access/
function Get-GroupMembersRecursive 
{
    param($Group)
    [array]$Members = @()
    $Group = Get-Group $Group -ErrorAction SilentlyContinue -ResultSize Unlimited
    if (!$Group)
    {
        throw "Group not found"
        
    }
    foreach ($Member in $Group.Members)
    {
        if (Get-Group $Member -ErrorAction SilentlyContinue -ResultSize Unlimited)
        {
            $Members += Get-GroupMembersRecursive -Group $Member
        } else {
            $Members += ((get-user $Member.Name -ResultSize Unlimited).SamAccountName)
        }
    }
    $Members = $Members | Select -Unique
    return $Members
}
# List all current non inherited Full Access permissions, excluding SELF
$SAMGroup = Get-Group $Group
$SAMGroup = $SAMGroup.SamAccountName
$CurrentMailboxPermission = Get-MailboxPermission -Identity $Mailbox | where { ($_.AccessRights -like "*FullAccess*") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF")}| Select User 

# Cleanup all permissions directly added to Mailbox (not inherited), this is necessary to keep Group memberships 
# and actuall added permission the same. Note: the assumption is that Full Access permissions are only
# handed out via used security group and not manually. All other accounts and groups will have their permissions removed
foreach ($User in $CurrentMailboxPermission){
	$User = [String] $User.User
	Write-Output "Removing FullAccess permission for $User"
	Remove-MailboxPermission -Identity $Mailbox -User $User -AccessRights 'FullAccess' -InheritanceType 'All' -Confirm:$false
	}

# Listing every unique member of every group recursivly
[array]$Members = @();
$Group = Get-Group $Group -ErrorAction SilentlyContinue -ResultSize Unlimited;
if (!$Group)
{
    throw "Group not found"
}
[array]$Members = Get-GroupMembersRecursive -Group $Group

# Adding the mailbox permission Full Access to users in group and subgroups on mailbox
Write-Output "The following users have Full Access on $Mailbox :"
foreach ($Member in $Members) {
		Add-MailboxPermission -Identity $Mailbox -User $Member -AccessRights FullAccess 
}