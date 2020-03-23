﻿<#
.SYNOPSIS
    Looks for O365 mailboxes based on a location name and returns SendOnBehalf, FullAccess, and SendAs permissions for the mailboxes of interest. 

.DESCRIPTION
    To run this script your organization must assigned the appropriate O365 permissions/roles to execute the Get-Mailbox, Get-MailboxPermissions, Get-RecipientPermission, and Get-ADUser Exchange Online PowerShell and Active Directory cmdlets.

    SendOnBehalf:
    Checks all mailboxes that have user accounts listed under this property.  If there are multiple it displays each account and queries the distinguished name (DN) and universal principal name (UPN) for that account.  In instances where an account has a non-unique name this query will return multiple values and display all of them.  It is up to the analyst to further determine which one(s) are accurate.
    
    FullAccess:
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
    It also pulls the full distinguised name (DN) for each account as a reference, primarily to identify accounts outside of the local organizational unit (OU).  In some instances the mailbox owner is returned as a result with having full access permissions to their own mailbox.  If this occurs the script does not include that result to provide for easier analysis.

    SendAs:
    Gets SendAs allowed permissions on a mailbox that is not the owner (i.e., SELF).  Sometimes the permissions owner (i.e., the trustee) is the same as the mailbox owner and in these instances those results are removed.  If the permissions owner is an unresolved/orphaned SID the attempt to query its DN are bypassed.  If an Active Directory user look-up is attempted on other objects and failed this is noted in the OrganizationUnit field.

.PARAMETER Location
    The location or organizational unit for the mailboxes of interest.
    
.PARAMETER CsvFileName
    An optional file name can be specified for the generated CSV output (default name: MailboxPermissions.csv)

.PARAMETER OutputTerminal
    Display the results to the PowerShell terminal instead of writing them to a CSV file.  This output is a custom object so alternatively it can be pipled to other PowerShell commands.
    
.EXAMPLE
    .\Get-O365MailboxPermissions.ps1 -Location Beijing

    Search for mailboxes with users assigned to Beijing and write the CSV output to a file named 'MailboxPermissions.csv' in the current working directory location.
    
.EXAMPLE
    .\Get-O365MailboxPermissions.ps1 -Location Beijing -CsvFileName BeijingMailboxPermissions.csv

    Search for mailboxes with users assigned to Beijing and write the CSV output to a file named 'BeijingMailboxPermissions.csv' in the current working directory location.

.EXAMPLE
    .\Get-O365MailboxPermissions.ps1 -Location Beijing -OutputTerminal

    Search for mailboxes with users assigned to Beijing and display the results in the PowerShell terminal. This output could alternatively be piped to other PowerShell commands.

.NOTES
    Version 0.3 - Last Modified 23 MAR 2020
    Author: Sam Pursglove
#>

[CmdletBinding(DefaultParameterSetName='Csv')]
param 
(
    [Parameter(Position=0, 
               Mandatory=$true, 
               ValueFromPipeline=$false, 
               HelpMessage='Enter the mailbox search location')]
    [string]$Location,

    [Parameter(Mandatory=$false, 
               ValueFromPipeline=$false, 
               ParameterSetName='Csv', 
               HelpMessage='Optionally, specify the name of the output CSV file (default: MailboxPermissions.csv)')]
    [string]$CsvFileName='MailboxPermissions.csv',

    [Parameter(Mandatory=$false,
               ValueFromPipeline=$false, 
               ParameterSetName='Terminal', 
               HelpMessage='Optionally, output the data to the PowerShell terminal (default: CSV file)')]
    [switch]$OutputTerminal
)

Set-StrictMode -Version 3


# helper function to get the distinguised name of an object
# if there is more than one it displays each as a string
function Get-DistinguishedName {
    Param (
        [Parameter(Mandatory)]
        [System.Object[]]$adUser
    )

    if($adUser.Count -gt 1) {
        $results = $adUser.DistinguishedName | ForEach-Object {$_.Split(",")[-4..-1]}
        $results = $results -join ','
        $script:notUniqueName = $true # flag to denote if an object name is not unique within AD, all results are returned
    } else {
        $results = $adUser.DistinguishedName.Split(",")[-4..-1] -join ','
    }
    
    $results
}


# helper function to get the name of an object
# if there is more than one it displays each as a string
function Get-UserPrincipalName {
    Param (
        [Parameter(Mandatory)]
        [System.Object[]]$adUser
    )

    if($adUser.Count -gt 1) {
        $results = $adUser.UserPrincipalName -join ','
        $script:notUniqueName = $true # flag to denote if an object name is not unique within AD, all results are returned
    } else {
        $results = -join $adUser.UserPrincipalName
    }
    
    $results
}


# helper function to add mailbox permission results to the global array container
function Add-MailboxPermissionObject {
    Param (
        [Parameter(Position=0,Mandatory=$true)]
        $Mailbox,
        [Parameter(Position=1,Mandatory=$true)]
        $Owner,
        [Parameter(Position=2,Mandatory=$true)]
        $OrganizationalUnit,
        [Parameter(Position=3,Mandatory=$true)]
        $Permission
    )

    $mailboxPermissions.Add(
        [PSCustomObject] @{
            'Mailbox'           = $Mailbox;
            'Owner'             = $Owner;
            'OrganizationalUnit'= $OrganizationalUnit;
            'Permission'        = $Permission;
        }
    ) > $null
}


# Returns any SendOnBehalfPermissions for the given mailbox
function Get-SendOnBehalfPermissions {
    param (
        [Parameter(Mandatory)]
        $mail
    )
   
    if ($mailbox.GrantSendOnBehalfTo -notlike $null) {
        
        # get the SendOnBehalfTo owner(s) for each mailbox that has a non-null value for this property
        $owners = Select-Object -InputObject $mailbox -ExpandProperty GrantSendOnBehalfTo 
    
        # obtain the distinguisted name for each SendOnBehalf owner
        foreach($owner in $owners) {
            try {
                $userInfo = Get-ADUser -Filter "Name -like '$($owner)'"
            
            # if an account is not located in the local domain query the whole Global Catalog (GC)
            } catch [System.Management.Automation.RuntimeException] {
                $userInfo = Get-ADUser -Filter "Name -like '$($owner)'" -Server "$($globalCatalogServer):$GCPort"
            }

            # if an account was located get its Universal Principal Name and Distinguised Name
            if($userInfo -eq $null) {
                $userDN   = "The object's User Principal Name (UPN) and Distinguised Name (DN) cannot be located"
                $userUPN = "$owner"
            } else {
                $userDN   = Get-DistinguishedName $userInfo
                $userUPN = Get-UserPrincipalName $userInfo
            }

            Add-MailboxPermissionObject -Mailbox $mailbox.UserPrincipalName -Owner $userUPN -OrganizationalUnit $userDN -Permission 'SendOnBehalf'

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

    $fullAccess = Get-MailboxPermission -Identity $mail.UserPrincipalName | 
        Where-Object {
            $_.Deny -eq $false -and 
            $_.IsInherited -eq $false -and
            $_.AccessRights -like "*FullAccess*" -and
            $_.User -notlike 'NT AUTHORITY\SELF'
        }
    
    foreach($owner in $fullAccess) {

        # do not show an account if it is listed with full access permissions to itself 
        # (unknown why this occurs in some instances)
        if($mail.UserPrincipalName -notlike $owner.User) {  
            
            try {
                $userInfo = Get-ADUser -Filter "UserPrincipalName -like '$($owner.User)'"
            
            # if an account is not located in the local domain query the whole Global Catalog (GC)
            } catch [System.Management.Automation.RuntimeException] {
                $userInfo = Get-ADUser -Filter "UserPrincipalName -like '$($owner.User)'" -Server "$($globalCatalogServer):$GCPort"
            }

            # if an account was located get it's distinguised name
            if($userInfo -eq $null) {
                $userDN = "The object's Universal Principal Name (UPN) cannot be located"
            } else {
                $userDN = Get-DistinguishedName $userInfo
            }

            Add-MailboxPermissionObject -Mailbox $mail.UserPrincipalName -Owner $owner.User -OrganizationalUnit $userDN -Permission 'FullAccess'
              
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
    $recipient = Get-RecipientPermission -Identity $mail.UserPrincipalName | 
        Where-Object {
            $_.Trustee -notlike "NT AUTHORITY\SELF" -and 
            $_.AccessControlType -like "Allow" -and
            $_.Trustee -notlike (Get-ADUser -Filter "name -like '$($_.Identity)'").UserPrincipalName
        }

    foreach($owner in $recipient) {
  
        # if a SID is unresolved don't attempt to look it up in Active Directory
        if ($owner.Trustee -notlike "S-1-5-21-*") {
            try {
                $userInfo = Get-ADUser -Filter "UserPrincipalName -like '$($owner.Trustee)'"
            
            # if an account is not located in the local domain query the whole Global Catalog (GC)
            } catch [System.Management.Automation.RuntimeException] {
                $userInfo = Get-ADUser -Filter "UserPrincipalName -like '$($owner.Trustee)'" -Server "$($globalCatalogServer):$GCPort"
            }

            # if an account was located get it's distinguised name
            if($userInfo -eq $null) {
                $userDN = "The object's Universal Principal Name (UPN) cannot be located"
            } else {
                $userDN = Get-DistinguishedName $userInfo
            }
        } else {
            $userDN = "The object's Security Identifier (SID) cannot be resolved"
        }

        Add-MailboxPermissionObject -Mailbox $mail.UserPrincipalName -Owner $owner.Trustee -OrganizationalUnit $userDN -Permission 'SendAs'

        # clear shared variables
        $userInfo = $null
        $userDN   = $null
    }
}


$counter            = 1                                        # global progress counter variable
$GCPort             = 3268                                     # global catalog server port number
$notUniqueName      = $false                                   # flag to warn if an AD user name search returns non-unique results
$mailboxPermissions = New-Object System.Collections.ArrayList  # global array to hold all output data
$globalCatalogServer= Get-ADDomainController -Discover -Service GlobalCatalog # GC server for AD object lookups outside domain of interest
$filterString       = "OrganizationalUnit -like '*$($Location)*' -and IsMailboxEnabled -eq 'True'" # filter to find mailboxes of interest


Write-Progress -Activity "Get-O365MailboxPermissions" -Status "Please Wait: Searching for $($Location) Mailboxes"
$mailboxes = Get-Mailbox -Filter $filterString -ResultSize Unlimited

# main script loop
foreach($mailbox in $mailboxes) {

    $percentComplete = [int](($counter/$mailboxes.Length)*100)
    $currentStatus   = "Getting Permissions for $($mailbox.DisplayName) ($($counter) of $($mailboxes.Length))"
    
    Write-Progress -Activity "Get-O365MailboxPermissions" -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendOnBehalf Permission"
    Get-SendOnBehalfPermissions $mailbox
    
    Write-Progress -Activity "Get-O365MailboxPermissions" -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "FullAccess Permission"
    Get-FullAccessPermissions $mailbox
    
    Write-Progress -Activity "Get-O365MailboxPermissions" -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "SendAs Permission"
    Get-SendAsPermissions $mailbox

    $counter++
}


if ($notUniqueName) {
    Write-Host "WARNING: An Active Directory user lookup returned non-unique results and will display all possible User Principle Names (UPN)" -ForegroundColor Red
}

# output data to a CSV unless the -OutputTerminal switch is used
if ($OutputTerminal) {
    $mailboxPermissions
} else {
    $mailboxPermissions | Export-Csv -Path $CsvFileName -NoTypeInformation
}