<#
.SYNOPSIS
  Obtain UPNs of Active Directory users based on SAM accounts in CSV file
.DESCRIPTION
  N/A
.PARAMETER <Parameter_Name>
  N/A
.INPUTS
  None
.OUTPUTS
  Exports array to CSV "mbxpermissions_upn.csv"
.NOTES
  Version:        1.0
  Author:         Kevin Clark
  Creation Date:  2/8/2018
  Purpose/Change: Initial script development
  
.EXAMPLE
  N/A
#>

$users = import-csv -Path .\mailboxpermissions.csv
foreach ($user in $users) {
    $MailboxUser = Get-Aduser -Identity $user.Identity
    $PermittedUser = Get-Aduser -Identity $user.user
    If ($MailboxUser)
        { $user.MBXUPN = $MailboxUser.UserprincipalName
    }
    else {
        $user.MBXUPN = "Not Found"
    }
    If ($PermittedUser)
        { $user.PermitUPN = $PermittedUser.UserprincipalName
    }
    else {
        $user.PermitUPN = "Not Found"
    }
}
$users | Export-csv -Path mbxpermissions_upn.csv -NoTypeInformation