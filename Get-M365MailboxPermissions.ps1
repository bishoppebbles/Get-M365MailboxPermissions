<#
.SYNOPSIS
    Looks for M365 mailboxes based on a location name and returns SendOnBehalf, FullAccess, and SendAs permissions for the mailboxes of interest.  Can optionally pull mailbox folder rights as well.
.DESCRIPTION
    To run this script your organization must assigned the appropriate M365 permissions/roles to execute the Get-EXOMailbox|Get-EXOMailboxPermissions|Get-EXORecipientPermissionGet-EXOMailboxFolderStatistics|Get-EXOMailboxFolderPermission and Get-AdDomain|Get-ADObject|Get-ADUser|Get-ADGroup|Get-ADGroupMember Exchange Online PowerShell and Active Directory cmdlets.  For large mailbox queries (appox. 500+) it's recommended to start a new remote session as it's possible your session will expire during the data pull.  It is also recommended to run this on a system in the domain where the majority of mailboxes of interest reside.  Otherwise a large number of queries to the Global Catalog (GC) will be performed and hinder performance.

    SendOnBehalf:
    
    Checks all mailboxes that have user accounts listed under this property.  If there are multiple it displays each account and queries the distinguished name (DN) and universal principal name (UPN) for that account.  In instances where an account has a non-unique name this query will return multiple values and display all of them.  It is up to the analyst to further determine which one(s) are accurate.
    
    
    FullAccess (open the mailbox, access its contents, but can't send mail):
    
    Looks for any non-inherited, approved (i.e., not denied) account that has full access permissions to the given mailbox that is not named 'SELF'.  Because inherited permissions are ignored this excludes the following:
        
        NT AUTHORITY\SYSTEM
        NT AUTHORITY\NETWORK SERVICE 
        Administrator
        Domain Admins
        Enterprise Admins
        Organization Management
        Exchange Servers
        Exchange Trusted Subsystem
        Managed Availability Servers
        Public Folder Management
    
    It also pulls the full distinguished name (DN) for each account as a reference (for both user and group Active Directory objects), primarily to identify accounts outside of the local organizational unit (OU).  In some instances the mailbox owner is returned as a result with having full access permissions to their own mailbox.  If this occurs the script does not include that result to provide for easier analysis.

    
    SendAs (the security principal can send messages that appear as if they are the mailbox owner (i.e., they impersonate them)):

    Gets SendAs allowed permissions on a mailbox that is not the owner (i.e., SELF).  Sometimes the permissions owner (i.e., the trustee) is the same as the mailbox owner and in these instances those results are removed.  If the permissions owner is an unresolved/orphaned SID the attempt to query its DN are bypassed.  If an Active Directory user look-up is attempted on other objects and failed this is noted in the OrganizationUnit field.

    
    Mailbox Folder Rights:
    
    Gets a mailbox's statistics to build a list of its folder paths and then pulls each folder's permissions.  Rights are excluded under the following circumstances:
        A right's user is the same as the mailbox's User Principal Name.
        If a user is 'Default.'
        If a user is assigned an access right of 'none.'
        If a user is an unresolved SID and is assigned an access right of 'Owner.'
.PARAMETER Location
    The location or organizational unit for the mailboxes of interest.
.PARAMETER Region
    The region of the mailboxes of interest.
.PARAMETER UserPrincipalName
    Specify the account that you want to use to connect.
.PARAMETER SearchBase
    The distinguished name path to use for computer object searching.
.PARAMETER Server
    The server to use for the target domain
.PARAMETER IncludeFolderRights
    Include the collection of the rights assigned to a mailbox's folders.
.PARAMETER MailboxRightsCsv
    An optional file name can be specified for the generated mailbox rights CSV output (default name: <Location>_Mailbox_Rights.csv)
.PARAMETER MailboxFolderRightsCsv
    An optional file name can be specified for the generated mailbox folder rights CSV output (default name: <Location>_Mailbox_Folder_Rights.csv)
.PARAMETER OutputTerminal
    Display the results to the PowerShell terminal instead of writing them to a CSV file.  This output is a custom object so alternatively it can be pipled to other PowerShell commands for additional processing.
.PARAMETER PermissionsType
    By default all permission types are queried except Folder Rights due to the speed.  Use this option to query a specific one: SendOnBehalfOnly, FullAccessOnly, SendAsOnly, FolderRightsOnly
.EXAMPLE
    .\Get-M365MailboxPermissions.ps1 -Location Beijing -Region Asia -UserPrincipalName bobsmith@corp.com -SearchBase 'ou=location,dc=company,dc=org' -Server company.org
    Search for mailboxes with users assigned to Beijing in the Asia region and write the CSV output to a file named 'Beijing_Mailbox_Rights.csv' in the current working directory location.
.EXAMPLE
    .\Get-M365MailboxPermissions.ps1 -Location Beijing -Region Asia -UserPrincipalName bobsmith@corp.com -SearchBase 'ou=location,dc=company,dc=org' -Server company.org -IncludeFolderRights
    Search for mailboxes with users assigned to Beijing in the Asia region and write the CSV output to a file named 'Beijing_Mailbox_Rights.csv' in the current working directory location.  It also pulls mailbox folder rights and saves it to a file named 'Beijing_Mailbox_Folder_Rights.csv (note: this is very slow).
.EXAMPLE
    .\Get-M365MailboxPermissions.ps1 -Location Beijing -Region Asia -UserPrincipalName bobsmith@corp.com -SearchBase 'ou=location,dc=company,dc=org' -Server company.org -PermissionsType FolderRightsOnly
    Search for mailboxes with users assigned to Beijing in the Asia region and ONLY write the CSV output of mailbox folder rights to a file named 'Beijing_Mailbox_Folder_Rights.csv' in the current working directory location (note: this is very slow).
.EXAMPLE
    .\Get-M365MailboxPermissions.ps1 -Location Beijing -Region Asia -UserPrincipalName bobsmith@corp.com -SearchBase 'ou=location,dc=company,dc=org' -Server company.org -MailboxRightsCsv BeijingMailboxRights.csv
    Search for mailboxes with users assigned to Beijing in the Asia region and write the CSV mailbox rights output to a file named 'BeijingMailboxRights.csv' in the current working directory location.
.EXAMPLE
    .\Get-M365MailboxPermissions.ps1 -Location Beijing -Region Asia -UserPrincipalName bobsmith@corp.com -SearchBase 'ou=location,dc=company,dc=org' -Server company.org -OutputTerminal
    Search for mailboxes with users assigned to Beijing in the Asia region and display the results in the PowerShell terminal. This output could alternatively be piped to other PowerShell commands.
.NOTES
    Version 1.20
    Last Modified: 12 June 2026
    Author: Sam Pursglove

    From Get-MailboxPermission help at https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailboxpermission?view=exchange-ps

    Identity
    - The mailbox in question
    
    User
    - The security principal (user, security group, Exchange management role group, etc.) that has permission to the mailbox
    
    AccessRights
    - The permission that the security principal has on the mailbox
      * ChangeOwner     : change the owner of the mailbox
      * ChangePermission: change the permissions on the mailbox
      * DeleteItem      : delete the mailbox
      * ExternalAccount : indicates the account isn't in the same domain
      * FullAccess      : open the mailbox, access its contents, but can't send mail
      * ReadPermission  : read the permissions on the mailbox
    
    IsInherited
    - Whether the permission is inherited (True) or directly assigned to the mailbox (False)
    - Permissions are inherited from the mailbox database and/or Active Directory
    - Typically, directly assigned permissions override inherited permissions
    
    Deny
    - Whether the permission(s) listed above is/are allowed (False) or denied (True)
    - Typically, deny permissions override allow permissions

    Default assigned user mailbox permissions
    - NT AUTHORITY\SELF
      * Directly assigned FullAccess and ReadPermission
        -- This entry gives a user permission to their own mailbox
    - Administrator, Domain Admins, Enterprise Admins and Organization Management
      * Deny FullAccess
        -- These inherited permissions prevent these users and group members from opening other users' mailboxes
      * Allow ChangeOwner, ChangePermission, DeleteItem, and ReadPermission
        -- Note that these inherited permission entries also appear to allow FullAccess, however, these users and groups do not have FullAccess to the mailbox because the inherited Deny permission entries override the inherited Allow permission entries
    - NT AUTHORITY\SYSTEM
      * Inherit FullAccess
    - NT AUTHORITY\NETWORK
      * Inherit ReadPermissionis
    - Exchange Servers
      * Inherit FullAccess and ReadPermission
    - Exchange Trusted Subsystem 
      * Inherit ChangeOwner, ChangePermission, DeleteItem, and ReadPermission
    - Managed Availability Servers
      * Inherit ReadPermission

    From Add-MailboxFolderPermission help at https://learn.microsoft.com/en-us/powershell/module/exchange/add-mailboxfolderpermission?view=exchange-ps

    Folder access roles defined: 
        Author          : CreateItems, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems
        Contributor     : CreateItems, FolderVisible
        Editor          : CreateItems, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems
        NonEditingAuthor: CreateItems, DeleteOwnedItems, FolderVisible, ReadItems
        Owner           : CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderContact, FolderOwner, FolderVisible, ReadItems
        PublishingAuthor: CreateItems, CreateSubfolders, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems
        PublishingEditor: CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems
        Reviewer        : FolderVisible, ReadItems
    Specific calendar folder roles:
        AvailabilityOnly: View only availability data
        LimitedDetails  : View availability data with subject and location
    
    Individual folder permissions defined:
        None            : The user has no access to view or interact with the folder or its contents.
        CreateItems     : The user can create items within the specified folder.
        CreateSubfolders: The user can create subfolders in the specified folder.
        DeleteAllItems  : The user can delete all items in the specified folder.
        DeleteOwnedItems: The user can only delete items that they created from the specified folder.
        EditAllItems    : The user can edit all items in the specified folder.
        EditOwnedItems  : The user can only edit items that they created in the specified folder.
        FolderContact   : The user is the contact for the specified public folder.
        FolderOwner     : The user is the owner of the specified folder. The user can view the folder, move the folder and create subfolders. The user can't read items, edit items, delete items or create items.
        FolderVisible   : The user can view the specified folder, but can't read or edit items within the specified public folder.
        ReadItems       : The user can read items within the specified folder.
#>

[CmdletBinding(DefaultParameterSetName='Csv')]
param 
(
    [Parameter(Position=0, 
               Mandatory=$true, 
               ValueFromPipeline=$false, 
               HelpMessage='Enter the mailbox search location')]
    [string]$Location,

    [Parameter(Position=1, 
               Mandatory=$true, 
               ValueFromPipeline=$false, 
               HelpMessage='Enter the mailbox region location')]
    [string]$Region,

    [Parameter(Mandatory=$true,
               ValueFromPipeline=$false,
               HelpMessage='Enter the user principal name')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory=$true, 
               ValueFromPipeline=$false, 
               HelpMessage='AD search location in DN form')]
    [string]$SearchBase,

    [Parameter(Mandatory=$true,
               ValueFromPipeline=$false,
               HelpMessage='Enter the server domain')]
    [string]$Server,

    [Parameter(Mandatory=$false,
               ValueFromPipeline=$false,
               ParameterSetName='Csv',
               HelpMessage="Switch to include the collection of a mailbox's folder access rights (default: false)")]
    [switch]$IncludeFolderRights=$false,

    [Parameter(Mandatory=$false, 
               ValueFromPipeline=$false, 
               ParameterSetName='Csv', 
               HelpMessage='Optionally, specify the name of the output CSV file (default: <Location>_Mailbox_Permissions.csv)')]
    [string]$MailboxRightsCsv='Mailbox_Rights.csv',

    [Parameter(Mandatory=$false, 
               ValueFromPipeline=$false, 
               ParameterSetName='Csv', 
               HelpMessage='Optionally, specify the name of the output CSV file (default: <Location>_Mailbox_Folder_Rights.csv)')]
    [string]$MailboxFolderRightsCsv='Mailbox_Folder_Rights.csv',

    [Parameter(Mandatory=$true,
               ValueFromPipeline=$false, 
               ParameterSetName='Terminal', 
               HelpMessage='Optionally, output the data to the PowerShell terminal (default: CSV file)')]
    [switch]$OutputTerminal,

    [Parameter(Mandatory=$false,
               ValueFromPipeline=$false,  
               HelpMessage='Query a single permission type: SendOnBehalf, FullAccess, SendAs, or FolderRights (default: all permissions types except FolderRights)')]
    [ValidateSet('SendOnBehalfOnly','FullAccessOnly','SendAsOnly','FolderRightsOnly')][string]$PermissionsType
)

Set-StrictMode -Version 3


# helper function to get the distinguished name of an object
# if there is more than one it displays each as a string
function Get-UserLocation {
    Param (
        [Parameter(Mandatory)]
        [System.Object[]]$adObject
    )

    $results = $null

    # extract the location name of the user acccount
    $adObject.msExchExtensionCustomAttribute1 | 
        ForEach-Object {
            if($_ -match 'iPostSite\|(.+)') {
                $results = $Matches[1]
            }        
        }

    if($results -eq $null) {
        $results = 'No location data'
    }

    $results
}


# helper function to get the name of an object
# if there is more than one it displays each as a string
function Get-UserPrincipalName {
    Param (
        [Parameter(Mandatory)]
        [System.Object[]]$adObject
    )

    if($adObject.Count -gt 1) {
        $results = $adObject.UserPrincipalName -join '|'
        $script:notUniqueName = $true # flag to denote if an object name is not unique within AD, all results are returned
    } else {
        $results = $adObject.UserPrincipalName
    }
    
    $results
}


# attempt to locate a GUID in AD or Exchange Online to resolve to some type of friendly name
function Resolve-GUID {
    Param (
        [Parameter(Mandatory)]
        $guid
    )

    if($found = Get-ADUser -Filter "msDS-ExternalDirectoryObjectId -eq 'User_$guid'" -Properties DisplayName,msDS-ExternalDirectoryObjectId,msExchExtensionCustomAttribute1,UserPrincipalName,ObjectClass -Server $Server) {
        [pscustomobject]@{
            Displayname                    = $found.DisplayName
            UserPrincipalName              = $found.UserPrincipalName
            msExchExtensionCustomAttribute1= $found.msExchExtensionCustomAttribute1
            ObjectClass                    = $found.ObjectClass
        }
    } else {
        try {	                    
            if($found = Get-EXORecipient -Identity $guid -Properties CustomAttribute7,ObjectClass -ErrorAction Stop) {
                [pscustomobject]@{
                    Displayname                    = $found.DisplayName
                    UserPrincipalName              = $found.PrimarySmtpAddress # convert PrimarySmtpAddress to UserPrincipalName for compatibility with the Get-UserPrincipalName function
                    msExchExtensionCustomAttribute1= $found.CustomAttribute7   # convert CustomAttribute7 to msExchExtensionCustomAttribute1 for compatibility with the Get-UserLocation function
                    ObjectClass                    = $found.ObjectClass
                }
            }        
        } catch [Microsoft.Exchange.Management.RestApiClient.RestClientException] {
            Write-Host "GUID lookup failed: $($guid.Identity) -> continuing"
        }
    }
}


# helper function to add mailbox permission results to the global array container
function Add-MailboxPermissionObject {
    Param (
        [Parameter(Position=0,Mandatory=$true)]
        $Mailbox,
        [Parameter(Position=1,Mandatory=$true)]
        $SecurityPrincipal,
        [Parameter(Position=2,Mandatory=$true)]
        $Location,
        [Parameter(Position=3,Mandatory=$true)]
        $AccessRight
    )

    $mailboxPermissions.Add(
        [PSCustomObject]@{
            Mailbox          = $Mailbox
            SecurityPrincipal= $SecurityPrincipal
            Location         = $Location
            AccessRight      = $AccessRight
        }
    ) > $null
}


# helper function to add mailbox folder permissions results to the global array container
function Add-MailboxFolderPermissionObject {
    Param (
        [Parameter(Position=0,Mandatory=$true)]
        $Mailbox,
        [Parameter(Position=1,Mandatory=$true)]
        $FolderName,
        [Parameter(Position=2,Mandatory=$true)]
        $User,
        [Parameter(Position=3,Mandatory=$true)]
        $Location,
        [Parameter(Position=4,Mandatory=$true)]
        $AccessRights
    )

    $mailboxFolderPermissions.Add(
        [PSCustomObject]@{
            Mailbox     = $Mailbox
            FolderName  = $FolderName
            User        = $User
            Location    = $Location
            AccessRights= $AccessRights -join '|'
        }
    ) > $null
}


# Returns any SendOnBehalfPermissions for the given mailbox
function Get-SendOnBehalfPermissions {
    param (
        [Parameter(Mandatory)]
        $mail
    )
   
    if ($mail.GrantSendOnBehalfTo -notlike $null) {
        
        # get the SendOnBehalf owner(s) for each mailbox that has a non-null value for this property
        $owners = Select-Object -InputObject $mail -ExpandProperty GrantSendOnBehalfTo 
    
        # obtain the distinguished name and user principal name for each SendOnBehalf owner
        foreach($owner in $owners) {
            
            # if the owner of the permissions is a GUID, attempt to resolve and obtain related details
            if($owner -match '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}') {
                
                # Don't perform an AD or Exchange lookup if the GUID object was previously discovered (stored in the global $guidLookupTable hashtable)
                if($guidLookupTable.ContainsKey($owner)) {
                    $userInfo = $guidLookupTable[$owner]
                } else {

                    # try to lookup a GUID to resolve to a friendly name
                    $userInfo = Resolve-GUID $owner

                    # save the lookup details for potential later reference
                    if($userInfo -ne $null) {
                        $guidLookupTable[$owner] = $userInfo
                    }
                }
            } else {

                # Don't perform an AD lookup if the object was previously discovered
                if($userLookupTable.ContainsKey($owner)) {
                    $userInfo = $userLookupTable[$owner]
                } else {
                    try {
                        # escape any names that use single quotes in the name
                        $escaped = $owner.Replace("'","''")

                        # object lookup in the local domain
                        $userInfo = Get-ADObject -Filter "Name -like '$($escaped)'" -Properties userPrincipalName,msExchExtensionCustomAttribute1 -SearchBase $SearchBase -Server $Server
                
                        # if an object is not located in the local domain query the Global Catalog (GC)
                        if ($userInfo -eq $null) {
                            $userInfo = Get-ADObject -Filter "Name -like '$($escaped)'" -Properties userPrincipalName,msExchExtensionCustomAttribute1 -Server ":$GCPort"
                        }

                        # if the object hasn't been located search in the local domain using a wildcard
                        if ($userInfo -eq $null) {
                            $userInfo = Get-ADObject -Filter "Name -like '$($escaped)*'" -Properties userPrincipalName,msExchExtensionCustomAttribute1 -SearchBase $SearchBase -Server $Server
                        }

                        # if an object is not located in the local domain query the Global Catalog (GC) using a wildcard
                        if ($userInfo -eq $null) {
                            $userInfo = Get-ADObject -Filter "Name -like '$($escaped)*'" -Properties userPrincipalName,msExchExtensionCustomAttribute1 -Server ":$GCPort"
                        }

                        # store the object lookup data for potential reuse
                        if ($userInfo -ne $null) {
                            $userLookupTable[$owner] = $userInfo
                        }
                    } catch [Microsoft.ActiveDirectory.Management.ADFilterParsingException] {
                        Write-Host "Name lookup failed: $($owner) (SendOnBehalf) -> continuing"
                    }
                }
            }

            # if an object was located get its Universal Principal Name and location
            if($userInfo -ne $null) {
                $location  = Get-UserLocation $userInfo

                # a group will not have a UPN (unless a GUID was resolved) so use the name instead
                if ($userInfo.ObjectClass -notcontains 'group' -or $owner -match '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}') {
                    $userUPN = Get-UserPrincipalName $userInfo
                } else {
                    $userUPN = "$owner"
                }

            } else {
                $location  = "Cannot locate the object"
                $userUPN = "$owner"
            }

            Add-MailboxPermissionObject -Mailbox $mail.UserPrincipalName -SecurityPrincipal $userUPN -Location $location -AccessRight 'SendOnBehalf'

            # reset the shared variables
            $userInfo = $null
            $location = $null
            $userName = $null
        }
    }
}


# show accounts with full access permissions to accounts other than their own
function Get-FullAccessPermissions {
    param (
        [Parameter(Mandatory)]
        $mail
    )

    $fullAccess = Get-EXOMailboxPermission -Identity $mail.UserPrincipalName | 
        Where-Object {
            $_.Deny -eq $false -and 
            $_.IsInherited -eq $false -and
            $_.AccessRights -like "*FullAccess*" -and
            $_.User -notlike 'NT AUTHORITY\SELF'
        }
    
    foreach($owner in $fullAccess) {

        # attempt to get the owner's object using it's UPN, DisplayName, or Name
        # do not show an account if it is listed with full access permissions to itself (unknown why this occurs in some instances)
        if($mail.UserPrincipalName -notlike $owner.User) {  
            
            # Don't perform an AD lookup if the object was previously discovered
            if($userLookupTable.ContainsKey($owner.user)) {
                $userInfo = $userLookupTable[$owner.user]

            } else {
                try {
                    $escaped = ($owner.User).Replace("'","''")

                    if (($userInfo = Get-ADObject -Filter "UserPrincipalName -like '$($escaped)' -or Name -like '$($escaped)' -or DisplayName -like '$($escaped)'" -Properties msExchExtensionCustomAttribute1 -SearchBase $SearchBase -Server $Server) -eq $null) {
                
                        # if an object is not located in the local domain query the Global Catalog (GC)
                        $userInfo = Get-ADObject -Filter "UserPrincipalName -like '$($escaped)' -or Name -like '$($escaped)' -or DisplayName -like '$($escaped)'" -Properties msExchExtensionCustomAttribute1 -Server ":$GCPort"
                    }
                } catch [Microsoft.ActiveDirectory.Management.ADFilterParsingException] {
                    Write-Host "Name lookup failed: $($owner.User) (FullAccess) -> continuing"
                }

                # store the object lookup data for potential reuse
                if ($userInfo -ne $null) {
                    $userLookupTable[$owner.User] = $userInfo
                }
            }

            # if an object was located get its location
            if($userInfo -ne $null) {
                $location  = Get-UserLocation $userInfo
            } else {
                $location = "Cannot find the UPN"
            }

            Add-MailboxPermissionObject -Mailbox $mail.UserPrincipalName -SecurityPrincipal $owner.User -Location $location -AccessRight 'FullAccess'
              
            # clear shared variables
            $userInfo = $null
            $location = $null
        }
    }
}


# Get the SendAs permissions for the given mailbox
function Get-SendAsPermissions {
    param (
        [Parameter(Mandatory)]
        $mail
    )

    # show accounts that have send as permissions to an account other than their own
    $recipient = Get-EXORecipientPermission -Identity $mail.UserPrincipalName | 
        Where-Object { 
            $_.Trustee -notlike "NT AUTHORITY\SELF" -and 
            $_.Trustee -notlike $mail.UserPrincipalName -and
            $_.AccessControlType -like "Allow"
        }

    foreach($owner in $recipient) {
        $trustee = $owner.Trustee

        # if the owner of the permissions is a GUID, attempt to resolve and obtain related details
        if($trustee -match '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}') {
                
            # Don't perform an AD or Exchange lookup if the GUID object was previously discovered (stored in the global $guidLookupTable hashtable)
            if($guidLookupTable.ContainsKey($trustee)) {
                $userInfo = $guidLookupTable[$trustee]

            } else {
                # try to lookup a GUID to resolve to a friendly name
                $userInfo = Resolve-GUID $trustee

                # save the lookup details to potential reference later
                if($userInfo -ne $null) {
                    $guidLookupTable[$trustee] = $userInfo
                }
            }
        # if a SID is unresolved don't attempt to look it up in Active Directory
        } elseif ($trustee -notlike "S-1-5-21-*") {

            # Don't perform an AD lookup if the object was previously discovered
            if($userLookupTable.ContainsKey($trustee)) {
                $userInfo = $userLookupTable[$trustee]
            } else {

                # try to find the trustee UserPincipalName (UPN), Name, or Display Name and if it that fails try in the Global Catalog (GC)
                try {
                    $escaped = $trustee.Replace("'","''")

                    if (($userInfo = Get-ADObject -Filter "UserPrincipalName -like '$($escaped)' -or Name -like '$($escaped)' -or DisplayName -like '$($escaped)'" -Properties msExchExtensionCustomAttribute1 -SearchBase $SearchBase -Server $Server) -eq $null) {
                
                        $userInfo = Get-ADObject -Filter "UserPrincipalName -like '$($escaped)' -or Name -like '$($escaped)' -or DisplayName -like '$($escaped)'" -Properties msExchExtensionCustomAttribute1 -Server ":$GCPort"
                    }

                    # store the object lookup data for potential reuse
                    if ($userInfo -ne $null) {
                        $userLookupTable[$trustee] = $userInfo
                    }
                } catch [Microsoft.ActiveDirectory.Management.ADFilterParsingException] {
                    Write-Host "Name lookup failed: $($trustee) (SendAs) -> continuing"
                }
            }
        } 
        
        
        if($trustee -like "S-1-5-21-*") {
            $location = "Cannot resolve the object's Security Identifier (SID)"
        } else {
            # if an object was located get its location
            if($userInfo -ne $null) {
                $location  = Get-UserLocation $userInfo

                # for GUIDs that have been resolved replace the Trustee name with the UPN
                if($trustee -match '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}') {
                    $trustee = $userInfo.UserPrincipalName
                }
            } else {
                $location = "Cannot locate the recipient trustee"
            }
        }
        
        Add-MailboxPermissionObject -Mailbox $($mail.UserPrincipalName) -SecurityPrincipal $trustee -Location $location -AccessRight 'SendAs'

        # clear shared variables
        $userInfo = $null
        $location = $null
    }
}


# Get folder permissions for the given mailbox
function Get-FolderPermissions {
    param (
        [Parameter(Mandatory)]
        $mail
    )

    # regex to help format a mailbox's folder path correctly as input to the Get-EXOMailboxFolderPermissions cmdlet
    $r = [regex]'\\'
    
    $folderPermissions = (Get-EXOMailboxFolderStatistics $mail.UserPrincipalName).Identity | 
        ForEach-Object {$r.Replace($_, ':\', 1)} | 
        Get-EXOMailboxFolderPermission -ErrorAction SilentlyContinue |
        Where-Object {
            -not (($_.AccessRights -eq 'None') `
                  -or ($_.User -like 'Default') `
                  -or ($_.User -like $mail.UserPrincipalName) `
                  -or ($_.User -like 'NT:S-1-5-21-*' -and $_.AccessRights -eq 'Owner')
            )
        }

    foreach($folder in $folderPermissions) {
        $userInfo = $null
        
        # Don't attempt to lookup unresolved SIDs
        if($folder.User.DisplayName -notlike 'NT:S-1-5-21-*') {
        
            # Don't perform an AD lookup if the object was previously discovered
            if($userLookupTableFolder.ContainsKey($folder.User.DisplayName)) {
                
                $userInfo = $userLookupTableFolder[$folder.User.DisplayName]
            } else {
            
                # escape any names that use single quotes in the name
                $escaped = ($folder.User.DisplayName).Replace("'","''")
                $userInfo = Get-ADObject -Filter "DisplayName -eq '$escaped'" -Properties DisplayName,msExchExtensionCustomAttribute1,extensionAttribute7 -Server $Server

                # store the object lookup data for potential reuse
                if ($userInfo -ne $null) {
                    $userLookupTableFolder[$folder.User.DisplayName] = $userInfo
                }
            }

            if($userInfo -ne $null) {

                # extract the location name of the object
                if($userInfo.ObjectClass -eq 'group') {
                    $obj = [PSCustomObject]@{
                                DisplayName                    = $userInfo.DisplayName
                                DistinguishedName              = $userInfo.DistinguishedName
                                msExchExtensionCustomAttribute1= $userInfo.extensionAttribute7 # convert extensionAttribute7 to msExchExtensionCustomAttribute1 for compatibility with the Get-UserLocation function
                                Name                           = $userInfo.Name
                                ObjectClass                    = $userInfo.ObjectClass
                    }

                    $location = Get-UserLocation $obj
                } else {
                    $location = Get-UserLocation $userInfo
                }                    
            } else {
                    # The object could not be located in the directory
                $location = 'AD lookup failed'
            } 
    
        } else {
            # Return no location data for an orphaned SID
            $location = ''
        }

        Add-MailboxFolderPermissionObject -Mailbox $mail.UserPrincipalName -FolderName $($folder.FolderName) -User $($folder.User.DisplayName) -Location $location -AccessRight $($folder.AccessRights)
    }
}

# Confirm the required PowerShell modules are available or installed.
if (Get-Module ActiveDirectory) {
    Write-Host 'The Active Directy PowerShell module is installed.'
} else {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host 'The Active Directory PowerShell module was imported.'
    } catch [System.IO.FileNotFoundException] {
        Write-Host 'The Active Directory PowerShell module is unavailable. Exiting.'
        Return
    }
}

if (Get-Module exchangeonlinemanagement) {
    Write-Host 'The Exchange Online PowerShell module is installed.'
} else {
    try {
        Import-Module exchangeonlinemanagement -ErrorAction Stop
        Write-Host 'The Exchange Online PowerShell module was imported.'
    } catch {
        Write-Host 'The Exchange Online PowerShell module is unavailable. Exiting.'
        Return
    }
}

$counter                 = 1                                        # global progress counter variable
$GCPort                  = 3268                                     # global catalog server port number
$notUniqueName           = $false                                   # flag to warn if an AD user name search returns non-unique results
$mailboxes               = New-Object System.Collections.ArrayList  # global array to hold all mailbox rights output data
$mailboxPermissions      = New-Object System.Collections.ArrayList  # global array to hold all mailbox rights output data
$mailboxFolderPermissions= New-Object System.Collections.ArrayList  # global array to hold all folder permissions output data
$workstationDnsRoot      = (Get-ADDomain).DNSRoot
$connect                 = $false
$userLookupTable         = @{}                                      # dictionary to save AD user object data to reduce redundant lookups
$guidLookupTable         = @{}                                      # dictionary to save GUIDs found in AD or Exchange to reduce redundant lookups
$userLookupTableFolder   = @{}                                      # dictionary to save AD user object data to reduce mailbox folder redundant lookups

# Connection to Exchange Online unless a v3 session is already established
if(($conn = Get-ConnectionInformation)) {
    foreach($c in $conn) {
        if ($c.Name -like "ExchangeOnline_3") {
            if ($c.State -like "Connected" -and $c.TokenStatus -like "Active") {
                Write-Host 'An Exchange Online PowerShell session is already extablished.'
            } else {
                $connect = $true
            }
        }
    }
} else {
    $connect = $true
}

# connect to exchange online if no connection exists
if($connect) {
    $params = @{
        UserPrincipalName = $UserPrincipalName
    }
                
    $exVer = (Get-Module exchangeonlinemanagement).Version

    # if the Exchange module version is 3.7.2 or higher it seems you need to disable the 
    # Web Account Manager (DisableWAM) to avoid an authentication issue acquiring a token.
    if ([int]$exVer.Major -gt 3) {
        $params['DisableWAM'] = $true
    } elseif ([int]$exVer.Major -eq 3 -and [int]$exVer.Minor -gt 7) {
        $params['DisableWAM'] = $true
    } elseif ([int]$exVer.Major -eq 3 -and [int]$exVer.Minor -eq 7 -and [int]$exVer.Build -ge 2) {
        $params['DisableWAM'] = $true
    }

    Connect-ExchangeOnline @params
}


Write-Host "Please Wait: Searching for $($Location) Mailboxes"

# get M365 mailbox accounts that have not migrated domains
$mail1 = Get-EXOMailbox -Filter "CustomAttribute7 -like '*$($Location)*'" -Properties GrantSendOnBehalfTo,IsMailboxEnabled -ResultSize Unlimited |
    Where-Object {$_.IsMailboxEnabled -eq 'True'}

# get M365 mailbox accounts that have migrated domains
$users = Get-ADGroup -Filter "Name -like `"*Users_$($Region)_$($Location)`"" -SearchBase "ou=groups,$SearchBase" -Server $Server |
    Get-ADGroupMember |
    Get-ADUser  -Server $Server

$mail2 = $users | 
    Where-Object {$mail1.UserPrincipalName -notcontains $_.UserPrincipalName} |  # ensure there are no duplicate Exchange mailbox lookups
    Get-EXOMailbox -Properties GrantSendOnBehalfTo,IsMailboxEnabled -ResultSize Unlimited -ErrorAction SilentlyContinue |
    Where-Object {$_.IsMailboxEnabled -eq 'True'}


# consolidate all mailbox results into a single container object
foreach($mail in $mail1) {
    $mailboxes.Add($mail) | Out-Null
}

foreach($mail in $mail2) {
    $mailboxes.Add($mail) | Out-Null
}


# main script loop
foreach($mailbox in $mailboxes) {

    $activity        = "Get-M365MailboxPermissions for $($Location) ($($counter) of $($mailboxes.Count) mailboxes)"
    $currentStatus   = "Getting mailbox permissions for $($mailbox.DisplayName)"
    $percentComplete = [int](($counter/$mailboxes.Count)*100)
    
    if ($PermissionsType -like 'SendOnBehalfOnly') {
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendOnBehalf Permission"
        Get-SendOnBehalfPermissions $mailbox
    
    } elseif ($PermissionsType -like 'FullAccessOnly') {
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "FullAccess Permission"
        Get-FullAccessPermissions $mailbox
    
    } elseif ($PermissionsType -like 'SendAsOnly') {
        
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendAs Permission"
        Get-SendAsPermissions $mailbox
    
    } elseif ($PermissionsType -like 'FolderRightsOnly') {
        
        $currentStatus   = "Getting mailbox folder permissions for $($mailbox.DisplayName)"
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "Folder Rights"
        Get-FolderPermissions $mailbox
    
    } else {
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendOnBehalf Permission"
        Get-SendOnBehalfPermissions $mailbox
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "FullAccess Permission"
        Get-FullAccessPermissions $mailbox
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendAs Permission"
        Get-SendAsPermissions $mailbox

        if ($IncludeFolderRights) {
            $currentStatus   = "Getting mailbox folder permissions for $($mailbox.DisplayName)"
            Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "Folder Rights"
            Get-FolderPermissions $mailbox
        }
    }

    $counter++
}

if ($notUniqueName) {
    Write-Host "WARNING: An Active Directory user lookup returned non-unique results and will display all possible User Principle Names (UPN)" -ForegroundColor Red
}

# output data to a CSV unless the -OutputTerminal switch is used
if ($IncludeFolderRights -or ($PermissionsType -like 'FolderRightsOnly')) {
    $mailboxFolderPermissions | Export-Csv -Path "$($Location)_$($MailboxFolderRightsCsv)" -NoTypeInformation
}

# output data to a CSV unless the -OutputTerminal switch is used
if ($OutputTerminal) {
    $mailboxPermissions
} elseif ($PermissionsType -notlike 'FolderRightsOnly') {
    $mailboxPermissions | Export-Csv -Path "$($Location)_$($MailboxRightsCsv)" -NoTypeInformation
}

Write-Host 'Disconnect Exchange Online.'
Disconnect-ExchangeOnline -Confirm:$false