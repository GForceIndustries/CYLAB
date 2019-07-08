Import-Module ActiveDirectory
. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto

# labCountryInput
# 2-letters identifying the country the lab is associated with, user prompted for input
# Example: XY | xy | xY | Xy
do { $labCountryInput = Read-Host "Please enter the 2-letter country code of the lab to remove mailboxes for" } until ($labCountryInput.length -eq 2)

# labCountryUpper
# 2-letters identifying the country the lab is associated with, upper-case, from labCountryInput
# Example: XY
$labCountryUpper = $labCountryInput.toUpper()

# labCountryLower
# 2-letters identifying the country the lab is associated with, lower-case, from labCountryInput
# Example: xy
$LabCountryLower = $labCountryInput.toLower()

# labOU
# the name of the OU to be removed from Active Directory
# Example: XYLab Users
$labOU = "${labCountryUpper}LAB Users"

# distributionGroup
# the distribution group to remove
# Example: DistributionGroupXY
$distributionGroup = "DistributionGroup${labCountryUpper}"

# remove Impersonation rights
Remove-ManagementRoleAssignment "${labCountryUpper}LABImpersonation" -Confirm:$false
Remove-ManagementScope "${labCountryUpper}LABOU" -Confirm:$false

# delete mailbox users
$n = 1
While ($n -le 5) {
$mailboxUser = "${labCountryLower}user${n}"

Remove-ADGroupMember -Identity ${distributionGroup} -Member ${mailboxUser} -Confirm:$false

Get-Mailbox ${mailboxUser} | Disable-Mailbox -Confirm:$false

Get-ADUser ${mailboxUser} | Remove-ADUser -Confirm:$false

$n += 1
}

# serviceAccountUser
# the service account username to remove
# Example: icadminxy
$serviceAccountUser = "icadmin${labCountryLower}"
Get-Mailbox ${serviceAccountUser} | Disable-Mailbox -Confirm:$false
Get-ADUser ${serviceAccountUser} | Remove-ADUser -Confirm:$false

# distributionGroup
# the distribution group to remove
# Example: DistributionGroupXY
$distributionGroup = "DistributionGroup${labCountryUpper}"
Get-DistributionGroup ${distributionGroup} | Remove-DistributionGroup -Confirm:$false

# mailboxDatabase
# the name of the mailbox database to be removed in Exchange
# Example: XYLAB
$mailboxDatabase = "${labCountryUpper}LAB"
$edbFile = Get-MailboxDatabase ${mailboxDatabase} -Status | select edbfilepath
Get-MailboxDatabase ${mailboxDatabase} | Dismount-Database -Confirm:$false
Get-MailboxDatabase ${mailboxDatabase} | Remove-MailboxDatabase -Confirm:$false

# delete mailbox database files
$edbFileString = Out-String -InputObject ${edbFile}
$edbFileString = $edbFileString.Replace("EdbFilePath","")
$edbFileString = $edbFileString.Replace("-","")
$edbFileString = $edbFileString.Trim()
$edbFileName = "\" + ${labCountryUpper} + "LAB.edb"
$edbFileDirectory = ${edbFileString}.Replace(${edbFileName},"")

Get-ChildItem -Path ${edbFileDirectory} -Recurse | Remove-Item -force -recurse
Remove-Item ${edbFileDirectory} -Force













# delete OU
Get-ADOrganizationalUnit -Identity "ou=${labOU},dc=cylab,dc=lan" | Set-ADObject -ProtectedFromAccidentalDeletion:$false -PassThru | Remove-ADOrganizationalUnit -Confirm:$false
