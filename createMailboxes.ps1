Import-Module ActiveDirectory
. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto

# labCountryInput
# 2-letters identifying the country the lab is associated with, user prompted for input
# Example: XY | xy | xY | Xy
do { $labCountryInput = Read-Host "Please enter the 2-letter country code of the lab to create mailboxes for" } until ($labCountryInput.length -eq 2)

# labEngineerInput
# name of the engineer this lab is for, user prompted for input
do { $labEngineerInput = Read-Host "Please enter the name of the engineer" } until ($labEngineerInput.length -gt 0)

# labCountryUpper
# 2-letters identifying the country the lab is associated with, upper-case, from labCountryInput
# Example: XY
$labCountryUpper = $labCountryInput.toUpper()

# labCountryLower
# 2-letters identifying the country the lab is associated with, lower-case, from labCountryInput
# Example: xy
$LabCountryLower = $labCountryInput.toLower()

# labOU
# the name of the OU to be created in Active Directory
# Example: XYLab Users
$labOU = "${labCountryUpper}LAB Users"

# Add OU
New-ADOrganizationalUnit -Name $labOU -Path "dc=cylab,dc=lan" -ProtectedFromAccidentalDeletion $true -Description $labEngineerInput

# mailboxDatabase
# the name of the mailbox database to be created in Exchange
# Example: XYLAB
$mailboxDatabase = "${labCountryUpper}LAB"

# Add Mailbox Database
New-MailboxDatabase -Server "LLAREGGUB" -Name $mailboxDatabase

# Mount Mailbox Database
Mount-Database -Identity $mailboxDatabase

# distributionGroup
# the distribution group to create
# Example: DistributionGroupXY
$distributionGroup = "DistributionGroup${labCountryUpper}"

# Add Distribution Group
New-DistributionGroup -Name $distributionGroup -OrganizationalUnit "cylab.lan/Distribution Groups" -SamAccountName $distributionGroup -Alias $distributionGroup

# serviceAccountUser
# the service account username to create
# Example: icadminxy
$serviceAccountUser = "icadmin${labCountryLower}"

# Add Service Account User
New-ADUser -Name $serviceAccountUser -AccountPassword(ConvertTo-SecureString -AsPlainText "support${labCountryLower}" -Force) -CannotChangePassword $true -ChangePasswordAtLogon $false -DisplayName $serviceAccountUser -Enabled $true -PasswordNeverExpires $true -PasswordNotRequired $false -Path "ou=${labOU},dc=cylab,dc=lan" -SamAccountName $serviceAccountUser -UserPrincipalName "${serviceAccountUser}@cylab.lan"

Enable-Mailbox -Identity "cylab.lan/${labOU}/${serviceAccountUser}" -Alias $serviceAccountUser -Database $mailboxDatabase

Get-Mailbox $serviceAccountUser | Set-Mailbox -ProhibitSendQuota 20Mb -ProhibitSendReceiveQuota 25Mb -IssueWarningQuota 15Mb -RecoverableItemsQuota 5Mb -RecoverableItemsWarningQuota 3Mb -ArchiveQuota 5Mb -ArchiveWarningQuota 3Mb -MaxSendSize 2Mb -MaxReceiveSize 2Mb

# Add Mailbox Users
$n = 1
While ($n -le 5) {
$mailboxUser = "${labCountryLower}user${n}"

New-ADUser -Name $mailboxUser -AccountPassword(ConvertTo-SecureString -AsPlainText "support${labCountryLower}" -Force) -CannotChangePassword $true -ChangePasswordAtLogon $false -DisplayName $mailboxUser -Enabled $true -PasswordNeverExpires $true -PasswordNotRequired $false -Path "ou=${labOU},dc=cylab,dc=lan" -SamAccountName $mailboxUser -UserPrincipalName "${mailboxUser}@cylab.lan"

Enable-Mailbox -Identity "cylab.lan/${labOU}/${mailboxUser}" -Alias $mailboxUser -Database $mailboxDatabase

Get-Mailbox $mailboxUser | Set-Mailbox -ProhibitSendQuota 20Mb -ProhibitSendReceiveQuota 25Mb -IssueWarningQuota 15Mb -RecoverableItemsQuota 5Mb -RecoverableItemsWarningQuota 3Mb -ArchiveQuota 5Mb -ArchiveWarningQuota 3Mb -MaxSendSize 2Mb -MaxReceiveSize 2Mb

Add-ADGroupMember -Identity $distributionGroup -Member $mailboxUser

$n += 1
}

# Impersonation Rights
New-ManagementScope -Name "${labCountryUpper}LABOU" -RecipientRoot "cylab.lan/${labCountryUpper}LAB Users" -RecipientRestrictionFilter {RecipientType -eq "UserMailbox"}
New-ManagementRoleAssignment -Name "${labCountryUpper}LABImpersonation" -Role "ApplicationImpersonation" -User "icadmin${labCountryLower}" -CustomRecipientWriteScope "${labCountryUpper}LABOU"
