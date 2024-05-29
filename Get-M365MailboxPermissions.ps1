﻿<#
.SYNOPSIS
    Looks for M365 mailboxes based on a location name and returns SendOnBehalf, FullAccess, and SendAs permissions for the mailboxes of interest. 
.DESCRIPTION
    To run this script your organization must assigned the appropriate M365 permissions/roles to execute the Get-EXOMailbox, Get-EXOMailboxPermissions, Get-EXORecipientPermission, and Get-ADObject Exchange Online PowerShell and Active Directory cmdlets.  For large mailbox queries (appox. 500+) it's recommended to start a new remote session as it's possible your session will expire during the data pull.  It is also recommended to run this on a system in the domain where the majority of mailboxes of interest reside.  Otherwise a large number of queries to the Global Catalog (GC) will be performed and hinder performance.

    
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
.PARAMETER Location
    The location or organizational unit for the mailboxes of interest.
.PARAMETER UserPrincipleName
    Specify the account that you want to use to connect.
.PARAMETER MailboxRightsCsv
    An optional file name can be specified for the generated mailbox rights CSV output (default name: <Location>_Mailbox_Rights.csv)
.PARAMETER MailboxFolderRightsCsv
    An optional file name can be specified for the generated mailbox folder rights CSV output (default name: <Location>_Mailbox_Folder_Rights.csv)
.PARAMETER OutputTerminal
    Display the results to the PowerShell terminal instead of writing them to a CSV file.  This output is a custom object so alternatively it can be pipled to other PowerShell commands for additional processing.
.PARAMETER PermissionsType
    By default all permission types are queried.  Use this option to query a specific one: SendOnBehalfOnly, FullAccessOnly, SendAsOnly
.EXAMPLE
    .\Get-M365MailboxPermissions.ps1 -Location Beijing -UserPrincipleName bobsmith@corp.com

    Search for mailboxes with users assigned to Beijing and write the CSV output to a file named 'Beijing_Mailbox_Rights.csv' in the current working directory location.
.EXAMPLE
    .\Get-M365MailboxPermissions.ps1 -Location Beijing -UserPrincipleName bobsmith@corp.com -CsvFileName BeijingMailboxRights.csv

    Search for mailboxes with users assigned to Beijing and write the CSV output to a file named 'BeijingMailboxRights.csv' in the current working directory location.
.EXAMPLE
    .\Get-M365MailboxPermissions.ps1 -Location Beijing -UserPrincipleName bobsmith@corp.com -OutputTerminal

    Search for mailboxes with users assigned to Beijing and display the results in the PowerShell terminal. This output could alternatively be piped to other PowerShell commands.
.NOTES
    Version 1.02 - Last Modified 29 MAY 2024
    Author: Sam Pursglove


    From Get-MailboxPermission help at https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailboxpermission?view=exchange-ps

    Identity
    - The mailbox in question
    
    User
    - The security principal (user, security group, Exchange management role group, etc.) that has permission to the mailbox
    
    AccessRights
    - The permission that the security principal has on the mailbox
      * ChangeOwner: change the owner of the mailbox
      * ChangePermission: change the permissions on the mailbox
      * DeleteItem: delete the mailbox
      * ExternalAccount: indicates the account isn't in the same domain
      * FullAccess: open the mailbox, access its contents, but can't send mail
      * ReadPermission: read the permissions on the mailbox
    
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
#>

[CmdletBinding(DefaultParameterSetName='Csv')]
param 
(
    [Parameter(Position=0, 
               Mandatory=$true, 
               ValueFromPipeline=$false, 
               HelpMessage='Enter the mailbox search location')]
    [string]$Location,

    [Parameter(Mandatory=$true,
               ValueFromPipeline=$false,
               HelpMessage='Enter the user principle name')]
    [string]$UserPrincipleName,

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
               HelpMessage='Query a single permission type: SendOnBehalf, FullAccess, or SendAs (default: all permissions types)')]
    [ValidateSet('SendOnBehalfOnly','FullAccessOnly','SendAsOnly')][string]$PermissionsType
)

Set-StrictMode -Version 3


# helper function to get the distinguished name of an object
# if there is more than one it displays each as a string
function Get-DistinguishedName {
    Param (
        [Parameter(Mandatory)]
        [System.Object[]]$adObject
    )

    # capture of the DN starting from the first OU= paramater
    $results = $adObject.DistinguishedName | 
        ForEach-Object {
            $_ -match ",OU=[a-zA-z ]+,DC=.+$" | Out-Null
            $tempStr = $Matches.Values
            $tempStr.Substring(1, $tempStr.Length - 1)
        }

    if (($results | Measure-Object).Count -gt 1) {
        $results = $results -join '|'
        $script:notUniqueName = $true # flag to denote if an object name is not unique within AD, all results are returned
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


# helper function to add mailbox permission results to the global array container
function Add-MailboxPermissionObject {
    Param (
        [Parameter(Position=0,Mandatory=$true)]
        $Mailbox,
        [Parameter(Position=1,Mandatory=$true)]
        $SecurityPrincipal,
        [Parameter(Position=2,Mandatory=$true)]
        $OrganizationalUnit,
        [Parameter(Position=3,Mandatory=$true)]
        $AccessRight
    )

    $mailboxPermissions.Add(
        [PSCustomObject]@{
            'Mailbox'           = $Mailbox;
            'SecurityPrincipal' = $SecurityPrincipal;
            'OrganizationalUnit'= $OrganizationalUnit;
            'AccessRight'       = $AccessRight;
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
        $AccessRight
    )

    $mailboxFolderPermissions.Add(
        [PSCustomObject]@{
            'Mailbox'    = $Mailbox;
            'FolderName' = $FolderName;
            'User       '= $User;
            'AccessRight'= $AccessRight;
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
            try {
                # object lookup in the local domain
                $userInfo = Get-ADObject -Filter "Name -like '$($owner)'" -Properties userPrincipalName
                
                # if an object is not located in the local domain query the Global Catalog (GC)
                if ($userInfo -eq $null) {
                    $userInfo = Get-ADObject -Filter "Name -like '$($owner)'" -Properties userPrincipalName -Server "$($globalCatalogServer):$GCPort"
                }

                # if the object hasn't been located search in the local domain using a wildcard
                if ($userInfo -eq $null) {
                    $userInfo = Get-ADObject -Filter "Name -like '$($owner)*'" -Properties userPrincipalName
                }

                # if an object is not located in the local domain query the Global Catalog (GC) using a wildcard
                if ($userInfo -eq $null) {
                    $userInfo = Get-ADObject -Filter "Name -like '$($owner)*'" -Properties userPrincipalName -Server "$($globalCatalogServer):$GCPort"
                }
            } catch [Microsoft.ActiveDirectory.Management.ADFilterParsingException] {
                Write-Host "Distinguished Name (DN) lookup parsing error: $($owner) (SendOnBehalf) -> continuing"
            }

            # if an object was located get its Universal Principal Name and Distinguished Name
            if($userInfo -ne $null) {
                $userDN  = Get-DistinguishedName $userInfo

                # a group will not have a UPN so use the name instead
                if ($userInfo.ObjectClass -notlike 'group') {
                    $userUPN = Get-UserPrincipalName $userInfo
                } else {
                    $userUPN = "$owner"
                }

            } else {
                $userDN  = "Cannot locate the object's User Principal Name (UPN) and Distinguished Name (DN)"
                $userUPN = "$owner"
            }

            Add-MailboxPermissionObject -Mailbox $mail.UserPrincipalName -SecurityPrincipal $userUPN -OrganizationalUnit $userDN -AccessRight 'SendOnBehalf'

            # reset the shared variables
            $userInfo = $null
            $userDN   = $null
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

        # attempt to get the owner's distinguished name (DN) using it's UPN or Name
        # do not show an account if it is listed with full access permissions to itself (unknown why this occurs in some instances)
        if($mail.UserPrincipalName -notlike $owner.User) {  
            
            try {
                if (($userInfo = Get-ADObject -Filter "UserPrincipalName -like '$($owner.User)' -or Name -like '$($owner.User)'") -eq $null) {
                
                    # if an object is not located in the local domain query the Global Catalog (GC)
                    $userInfo = Get-ADObject -Filter "UserPrincipalName -like '$($owner.User)' -or Name -like '$($owner.User)'" -Server "$($globalCatalogServer):$GCPort"
                }
            } catch [Microsoft.ActiveDirectory.Management.ADFilterParsingException] {
                Write-Host "Distinguished Name (DN) lookup parsing error: $($owner.User) (FullAccess) -> continuing"
            }

            # if an object was located get its Distinguished Name
            if($userInfo -ne $null) {
                $userDN  = Get-DistinguishedName $userInfo
            } else {
                $userDN = "Cannot locate the object's Universal Principal Name (UPN)"
            }

            Add-MailboxPermissionObject -Mailbox $mail.UserPrincipalName -SecurityPrincipal $owner.User -OrganizationalUnit $userDN -AccessRight 'FullAccess'
              
            # clear shared variables
            $userInfo = $null
            $userDN   = $null
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
  
        # try to find the trustee Distinguished Name (DN)
        # if a SID is unresolved don't attempt to look it up in Active Directory
        if ($owner.Trustee -notlike "S-1-5-21-*") {

            try {
                # try to find the AD object by UPN or Name
                if (($userInfo = Get-ADObject -Filter "UserPrincipalName -like '$($owner.Trustee)' -or Name -like '$($owner.Trustee)'") -eq $null) {
                
                    # try to find the AD object by UPN in the Global Catalog (GC)
                    $userInfo = Get-ADObject -Filter "UserPrincipalName -like '$($owner.Trustee)' -or Name -like '$($owner.Trustee)'" -Server "$($globalCatalogServer):$GCPort"
                }
            } catch [Microsoft.ActiveDirectory.Management.ADFilterParsingException] {
                Write-Host "Distinguished Name (DN) lookup parsing error: $($owner.Trustee) (SendAs) -> continuing"
            }

            # if an object was located get its Distinguished Name
            if($userInfo -ne $null) {
                $userDN  = Get-DistinguishedName $userInfo
            } else {
                $userDN = "Cannot locate the object's Universal Principal Name (UPN)"
            }
        } else {
            $userDN = "Cannot resolve the object's Security Identifier (SID)"
        }
        
        Add-MailboxPermissionObject -Mailbox $mail.UserPrincipalName -SecurityPrincipal $owner.Trustee -OrganizationalUnit $userDN -AccessRight 'SendAs'

        # clear shared variables
        $userInfo = $null
        $userDN   = $null
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
        Where-Object { -not (`
            (($_.User -like 'Default' -or $_.User -like 'Anonymous') `
                    -and $_.AccessRights -like 'None') `
               -or ($_.User -like 'NT:S-1-5-21-*' -and $_.AccessRights -like 'Owner') `
               -or ($_.User -like $mail.UserPrincipalName)
        )}

    foreach($folder in $folderPermissions) {
        Add-MailboxFolderPermissionObject -Mailbox $mail.UserPrincipalName -FolderName $($folder.FolderName) -User $($folder.User.DisplayName) -AccessRight $folder.AccessRights
    }
}

# Confirm the required PowerShell modules are available or installed.
if (Get-Module ActiveDirectory) {
    Write-Output 'The Active Directy PowerShell module is installed.'
} else {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Output 'The Active Directy PowerShell module was imported.'
    } catch [System.IO.FileNotFoundException] {
        Write-Output 'The Active Directory PowerShell module is unavailable.  Exiting.'
        Return
    }
}

if (Get-Module exchangeonlinemanagement) {
    Write-Output 'The Exchange Online PowerShell module is installed.'
} else {
    try {
        Import-Module exchangeonlinemanagement -ErrorAction Stop
        Write-Output 'The Exchange Online PowerShell module was imported.'
    } catch {
        Write-Output 'The Exchange Online PowerShell module is unavailable.  Exiting.'
        Return
    }
}

$counter                 = 1                                        # global progress counter variable
$GCPort                  = 3268                                     # global catalog server port number
$notUniqueName           = $false                                   # flag to warn if an AD user name search returns non-unique results
$mailboxPermissions      = New-Object System.Collections.ArrayList  # global array to hold all mailbox rights output data
$mailboxFolderPermissions= New-Object System.Collections.ArrayList  # global array to hold all folder permissions output data
$globalCatalogServer     = Get-ADDomainController -Discover -Service GlobalCatalog # GC server for AD object lookups outside domain of interest

Connect-ExchangeOnline -UserPrincipalName $UserPrincipleName

Write-Progress -Activity "Get-O365MailboxPermissions" -Status "Please Wait: Searching for $($Location) Mailboxes"
$mailboxes = Get-EXOMailbox -Filter "CustomAttribute7 -like '*$($Location)*'" -Properties 'GrantSendOnBehalfTo','IsMailboxEnabled' -ResultSize Unlimited |
    Where-Object {$_.IsMailboxEnabled -eq 'True'}


# main script loop
foreach($mailbox in $mailboxes) {

    $activity        = "Get-M365MailboxPermissions for $($Location) ($($counter) of $($mailboxes.Length) mailboxes)"
    $currentStatus   = "Getting permissions for $($mailbox.DisplayName)"
    $percentComplete = [int](($counter/$mailboxes.Length)*100)
    
    if ($PermissionsType -like 'SendOnBehalfOnly') {
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendOnBehalf Permission"
        Get-SendOnBehalfPermissions $mailbox
    
    } elseif ($PermissionsType -like 'FullAccessOnly') {
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "FullAccess Permission"
        Get-FullAccessPermissions $mailbox
    
    } elseif ($PermissionsType -like 'SendAsOnly') {
        
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendAs Permission"
        Get-SendAsPermissions $mailbox
    
    } else {
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendOnBehalf Permission"
        Get-SendOnBehalfPermissions $mailbox
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "FullAccess Permission"
        Get-FullAccessPermissions $mailbox
    
        Write-Progress -Activity $activity -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendAs Permission"
        Get-SendAsPermissions $mailbox
    }

    $counter++
}

if ($notUniqueName) {
    Write-Host "WARNING: An Active Directory user lookup returned non-unique results and will display all possible User Principle Names (UPN)" -ForegroundColor Red
}

# output data to a CSV unless the -OutputTerminal switch is used
if ($OutputTerminal) {
    $mailboxPermissions
} else {
    $mailboxPermissions | Export-Csv -Path "$($Location)_$($Mailbox_Rights.Csv)" -NoTypeInformation
}

# output data to a CSV unless the -OutputTerminal switch is used
if ($OutputTerminal) {
    $mailboxFolderPermissions
} else {
    $mailboxFolderPermissions | Export-Csv -Path "$($Location)_$($Mailbox_Folder_Rights.Csv)" -NoTypeInformation
}