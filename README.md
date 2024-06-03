# Get-M365MailboxPermissions

The script queries your M365 environment for a given set of cloud mailboxes and attempts to pull their `SendOnBehalf`, `FullAccess`, `SendAs`, and (optionally) mailbox folder permissions.  Your organization must assign your account the appropriate roles/permissions to run the following Exchange Online PowerShell v3 and Active Directory cmdlets:

* `Get-EXOMailbox`
* `Get-EXOMailboxPermissions`
* `Get-EXORecipientPermission`
* `Get-EXOMailboxFolderStatistics`
* `Get-EXOMailboxFolderPermission`
* `Get-ADObject`
