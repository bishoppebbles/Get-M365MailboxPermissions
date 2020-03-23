# Get-O365MailboxPermissions

The script queries your O365 environment for a given set of cloud mailboxes and attempts to pull their `SendOnBehalf`, `FullAccess`, and `SendAs` permissions.  Your organization must assign your account the appropriate roles/permissions to run the following Exchange Online PowerShell and Active Directory cmdlets:

* `Get-Mailbox`
* `Get-MailboxPermissions`
* `Get-RecipientPermission`
* `Get-ADUser`

## Major Note
The `$fiterString` variable searches your O365 environment for a group of mailboxes of interest.  **You must change** it accordingly to match your needs.  I didn't include it as a script option because of how I wanted the script cmdlet to work and my choice of parameter sets.
