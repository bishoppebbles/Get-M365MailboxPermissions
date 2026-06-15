 # Get-M365MailboxPermissions

The script queries the Exchange Online environment for a given set of cloud mailboxes and attempts to pull their `SendOnBehalf`, `FullAccess`, `DeleteItem`, `ReadPermission`, `ChangePermission`, `ChangeOwner`, `ExternalAccount`, `SendAs`, and (optionally) mailbox folder permissions.  For folder permissions it targets the `Calendar`, `Contacts`, `DeletedItems`, `Drafts`, `Inbox`, `SentItems`, and user created folders.  Your organization must assign your account the appropriate roles/permissions to run the following Exchange Online PowerShell v3 and Active Directory cmdlets:

## exchangeonlinemanagement Module
* `Get-EXOMailbox`
* `Get-EXOMailboxPermissions`
* `Get-EXORecipientPermission`
* `Get-EXOMailboxFolderStatistics`
* `Get-EXOMailboxFolderPermission`

## activedirectory Module
* `Get-ADDomain`
* `Get-ADObject`
* `Get-ADUser`
* `Get-ADGroup`
* `Get-ADGroupMember`

## Important Note
This script searches M365 mailboxes of interest based on a custom Exchange attribute.  Your environment may be different.
