Get-ADGroupMember -identity "SPO_ GDPR_Members" -Recursive | Get-ADUser -Property DisplayName | Select Name,DisplayName


