#Requires -Version 7.0

<#
.SYNOPSIS
    Central phonebook for Microsoft Teams: synchronises a SharePoint list
    to the personal Outlook contacts of a group of users.

.DESCRIPTION
    Central synchronisation between:

        Entra ID group (e.g. everyone with Teams Phone)
                |
        real Exchange UserMailboxes
                |
        scope group
        (technical Exchange Application RBAC scope)
                |
        central phonebook
        (Microsoft List / SharePoint)
                |
        personal Exchange contacts
                |
        Teams People / Teams Calls

    What it does:
    - Reads the target group from Entra ID via Microsoft Graph.
    - Filters out Teams resource accounts and other objects without a UserMailbox.
    - Keeps the technical Exchange RBAC scope group in sync.
    - Reads the central phonebook from SharePoint.
    - Creates missing contacts.
    - Updates existing centrally managed contacts.
    - Deletes centrally managed contacts that were removed from SharePoint.
    - Never touches personal contacts, except when merging copies of list
      items (see AdoptPersonalDuplicates).
    - Cleans up contacts for users who leave the target group.
    - Only after that removes a former target user from the RBAC scope group.
    - Contains several safety stops against unintended mass deletes.

    Requirements for a local test:
    First connect to Microsoft Graph app-only.
    Then run Connect-ExchangeOnline.

    The Graph app has:
      - Sites.Selected
      - GroupMember.ReadBasic.All
      - User.ReadBasic.All

    Contacts.ReadWrite is NOT granted tenant-wide in Entra ID.
    That access goes through Exchange Application RBAC.

    Azure Automation:
    Automation account with a system-assigned managed identity.
    Runbook parameters: ManagedIdentity = $true, DryRun = $true/$false,
    plus the environment parameters below.

    The managed identity has:
      - Graph: Sites.Selected (read on the phonebook site),
               GroupMember.ReadBasic.All, User.ReadBasic.All
      - Exchange.ManageAsApp
      - Exchange: Application Contacts.ReadWrite within a management
                  scope on the scope group
      - An Exchange role group with:
          View-Only Recipients +
          a custom role (only Add/Remove-DistributionGroupMember)
          with a write scope that matches only the scope group

    See README.md and docs/SETUP.md.

.PARAMETER DryRun
    Safe default: read and report only. Use $false to make changes.

.PARAMETER ManagedIdentity
    $true = connect using the managed identity (Azure Automation).
    $false = expects an existing Graph app-only + Exchange Online session (local).

.PARAMETER FunctionsOnly
    Loads only config + functions and stops. Intended for dot-sourcing
    during local tests.

.PARAMETER Organization
    Exchange Online organization (tenant) used for the managed identity
    connection, e.g. contoso.onmicrosoft.com.

.PARAMETER TeamsPhoneGroupId
    Object ID of the Entra ID group that determines who gets the central
    phonebook (typically everyone with Teams Phone).

.PARAMETER ScopeGroup
    Name or address of the technical mail-enabled security group that is
    used as the Exchange Application RBAC scope. Managed by this script.

.PARAMETER SharePointHost
    SharePoint host name, e.g. contoso.sharepoint.com.

.PARAMETER SharePointPath
    Server-relative path of the site that holds the list, e.g. /sites/Intranet.

.PARAMETER ListName
    Display name of the Microsoft List that is the phonebook source.

.PARAMETER NameColumnName
    Display name of the list column that holds the contact name.

.PARAMETER PhoneColumnName
    Display name of the list column that holds the phone number.

.PARAMETER ExtensionName
    Name of the open extension that marks contacts created by this script.
    Use your own reverse-domain name and never change it after go-live.

.PARAMETER ExtensionSourceSystem
    Free text stored in the extension (sourceSystem) for identification.

.PARAMETER MinimumTargetUsers
    Fewer valid target users than this stops the whole run.

.PARAMETER MinimumSourceContacts
    Fewer valid list items than this stops the whole run. Default 1;
    set it to a value somewhat below the normal size of your list.

.PARAMETER MaxDeletesPerMailbox
    More regular DELETEs in one mailbox than this skips that mailbox.

.PARAMETER MaxFormerUsersPerRun
    More users leaving the target group in one run than this stops the run.

.PARAMETER MaxFormerUserCleanupDeletes
    More managed contacts at a former user than this stops that cleanup.

.PARAMETER AdoptPersonalDuplicates
    Take over (ADOPT) personal copies of source contacts.

.PARAMETER DeleteExtraPersonalDuplicates
    Delete (DUPDELETE) extra personal copies with exactly our spelling.

.PARAMETER MaxPersonalDuplicateDeletesPerMailbox
    More DUPDELETEs in one mailbox than this skips that mailbox.

.PARAMETER MaxLookalikeRatio
    Fail-closed ratio for contacts without our extension that match the source.

.EXAMPLE
    .\Sync-CentralPhonebook.ps1 -Organization contoso.onmicrosoft.com `
        -TeamsPhoneGroupId 00000000-0000-0000-0000-000000000000 `
        -ScopeGroup "Phonebook Sync Users" -SharePointHost contoso.sharepoint.com `
        -SharePointPath /sites/Intranet -ListName "Central Phonebook"

    Dry run (default): reads and reports, changes nothing.

.EXAMPLE
    .\Sync-CentralPhonebook.ps1 <same parameters> -DryRun $false

    Live run.

.EXAMPLE
    . .\Sync-CentralPhonebook.ps1 <same parameters> -FunctionsOnly

    Load only the functions (no sync, no changes), for example to test
    Get-ManagedContacts on its own.

.NOTES
    Built by Adyta (Lambert van der Plas) - https://adyta.nl
    License: MIT
#>

param(
    # Safe default: read and report only.
    [bool]$DryRun = $true,

    # $true = connect using the managed identity (Azure Automation).
    # $false = expects an existing Graph app-only + EXO session (local).
    [bool]$ManagedIdentity = $false,

    # Loads only config + functions and stops.
    # Intended for dot-sourcing during local tests.
    [switch]$FunctionsOnly,

    # --------------------------------------------------------
    # Environment
    # --------------------------------------------------------

    # Tenant / Exchange organization for the managed identity connection,
    # e.g. contoso.onmicrosoft.com.
    [Parameter(Mandatory)]
    [string]$Organization,

    # Entra ID group that functionally determines who uses Teams Phone
    # and should therefore get the central phonebook.
    [Parameter(Mandatory)]
    [string]$TeamsPhoneGroupId,

    # Technical mail-enabled security group that is used as the
    # Exchange Application RBAC scope.
    [Parameter(Mandatory)]
    [string]$ScopeGroup,

    # SharePoint / Microsoft List, e.g. contoso.sharepoint.com and /sites/Intranet.
    [Parameter(Mandatory)]
    [string]$SharePointHost,

    [Parameter(Mandatory)]
    [string]$SharePointPath,

    [Parameter(Mandatory)]
    [string]$ListName,

    # Display names of the list columns for name and phone number.
    [string]$NameColumnName = "Name",

    [string]$PhoneColumnName = "Phone",

    # Marker that identifies only our own contacts.
    # Use your own reverse-domain name; never change it after go-live.
    [string]$ExtensionName = "com.contoso.centralphonebook",

    [string]$ExtensionSourceSystem = "Central Phonebook",

    # --------------------------------------------------------
    # Safety limits
    # --------------------------------------------------------

    # If Graph suddenly returns hardly any target users:
    # STOP, so that not everyone is treated as a former user.
    [int]$MinimumTargetUsers = 2,

    # An unexpectedly small source probably means an error.
    # Set this somewhat below the normal size of your list
    # (e.g. 80 for a list of about 100 contacts).
    [int]$MinimumSourceContacts = 1,

    # Maximum number of regular contact deletes per mailbox per run.
    # For example when contacts disappear from the central list.
    [int]$MaxDeletesPerMailbox = 25,

    # Maximum number of users that may leave the target group in
    # one run. More probably means a source/group error.
    [int]$MaxFormerUsersPerRun = 3,

    # A former target user normally gets about the size of the list
    # cleaned up. This is only an extreme sanity check.
    [int]$MaxFormerUserCleanupDeletes = 250,

    # MERGE of personal copies of source contacts (same name+number,
    # case ignored), only if there is no managed contact yet for that
    # source item:
    # - preference for the copy with a different spelling (entered by the user);
    # - that one gets our extension + our naming (ADOPT) and is kept;
    # - other copies with EXACTLY our spelling are deleted
    #   (DUPDELETE). Copies with a different spelling are always kept.
    [bool]$AdoptPersonalDuplicates = $true,

    [bool]$DeleteExtraPersonalDuplicates = $true,

    # Maximum number of DUPDELETEs per mailbox per run.
    [int]$MaxPersonalDuplicateDeletesPerMailbox = 25,

    # Fail-closed check: if this share (or more) of the source is already
    # present as a contact WITHOUT our extension in the mailbox, extension
    # detection is probably failing. Merge would then treat our own contacts
    # as duplicates. The mailbox is skipped. (Example: a user with 10 of
    # 200 source contacts saved personally = 5%, well below the limit.)
    [double]$MaxLookalikeRatio = 0.5
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"


# ============================================================
# CONFIG
# ============================================================

# DryRun, the environment values and the safety limits are script
# parameters, see the param block.

# Tenant / Exchange organization for the managed identity connection.
$ExchangeOrganization = $Organization

# Graph returns an open extension on Outlook resources (contacts,
# messages, events) with a fully qualified id:
#   Microsoft.OutlookServices.OpenTypeExtension.<extensionName>
# This id is used in $expand=extensions($filter=id eq '...').
# $expand=extensions without a filter is NOT allowed on Outlook contacts
# (ErrorGraphExtensionExpandRequiresFilter).
$ExtensionId = "Microsoft.OutlookServices.OpenTypeExtension.$ExtensionName"


# ============================================================
# LOGGING
# ============================================================

# Azure Automation does not capture the Information stream (Write-Host).
# In Automation (-ManagedIdentity $true) all output therefore goes to
# Verbose (the runbook has logVerbose enabled) and WARN/ERROR also to Warning.
# Locally: plain Write-Host with colour.
$AutomationLogging = [bool]$ManagedIdentity

# logVerbose sets $VerbosePreference to Continue in Automation; Graph/EXO
# would then log thousands of their own lines. We only want to see our
# own Write-Verbose -Verbose lines.
if ($AutomationLogging) {
    $VerbosePreference = "SilentlyContinue"
}

function Write-Console {
    param(
        [Parameter(Position = 0)]
        [string]$Object = "",

        [System.ConsoleColor]$ForegroundColor
    )

    if ($AutomationLogging) {

        if (-not [string]::IsNullOrEmpty($Object)) {
            Write-Verbose -Message $Object -Verbose
        }

        return
    }

    if ($PSBoundParameters.ContainsKey("ForegroundColor")) {
        Write-Host $Object -ForegroundColor $ForegroundColor
    }
    else {
        Write-Host $Object
    }
}


function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if ($AutomationLogging -and $Level -in @("WARN", "ERROR")) {
        Write-Warning "[$Timestamp] [$Level] $Message"
    }

    switch ($Level) {
        "INFO" {
            Write-Console "[$Timestamp] [INFO]    $Message"
        }

        "WARN" {
            Write-Console "[$Timestamp] [WARN]    $Message" -ForegroundColor Yellow
        }

        "ERROR" {
            Write-Console "[$Timestamp] [ERROR]   $Message" -ForegroundColor Red
        }

        "SUCCESS" {
            Write-Console "[$Timestamp] [SUCCESS] $Message" -ForegroundColor Green
        }
    }
}


# ============================================================
# GRAPH HELPERS
# ============================================================

function Get-GraphPaged {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $Results = @()

    do {
        $Response = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Uri

        $HasValue = $false
        $Value    = $null
        $NextLink = $null

        # Invoke-MgGraphRequest often returns a Hashtable/Dictionary.
        if ($Response -is [System.Collections.IDictionary]) {

            if ($Response.Contains("value")) {
                $HasValue = $true
                $Value = $Response["value"]
            }

            if ($Response.Contains("@odata.nextLink")) {
                $NextLink = [string]$Response["@odata.nextLink"]
            }
        }
        else {
            # Fallback for regular PowerShell objects.
            $ValueProperty = $Response.PSObject.Properties["value"]

            if ($null -ne $ValueProperty) {
                $HasValue = $true
                $Value = $ValueProperty.Value
            }

            $NextLinkProperty = $Response.PSObject.Properties["@odata.nextLink"]

            if ($null -ne $NextLinkProperty) {
                $NextLink = [string]$NextLinkProperty.Value
            }
        }

        if ($HasValue) {
            $Results += @($Value)
        }
        else {
            $Results += $Response
        }

        $Uri = $NextLink
    }
    while (-not [string]::IsNullOrWhiteSpace($Uri))

    return @($Results)
}


function Get-PropertyValue {
    # Reads a property StrictMode-safe from a Hashtable/IDictionary
    # (Invoke-MgGraphRequest) or a regular PowerShell/Exchange object.
    # Missing property = $null, never an exception.
    param(
        $InputObject,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {

        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }

        return $null
    }

    $Property = $InputObject.PSObject.Properties[$Name]

    if ($null -ne $Property) {
        return $Property.Value
    }

    return $null
}


function Get-GraphErrorText {
    # Invoke-MgGraphRequest puts the real Graph error code (e.g.
    # ErrorGraphExtensionExpandRequiresFilter / ErrorAccessDenied) in
    # ErrorDetails; Exception.Message only contains "BadRequest".
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $Text = [string]$ErrorRecord.Exception.Message

    if (
        $null -ne $ErrorRecord.ErrorDetails -and
        -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.ErrorDetails.Message)
    ) {
        $Text += " | " + [string]$ErrorRecord.ErrorDetails.Message
    }

    return $Text
}


function Get-CentralExtension {
    param(
        [Parameter(Mandatory)]
        $Contact
    )

    foreach ($Extension in @(Get-PropertyValue -InputObject $Contact -Name "extensions")) {

        if ($null -eq $Extension) {
            continue
        }

        $Id   = [string](Get-PropertyValue -InputObject $Extension -Name "id")
        $Name = [string](Get-PropertyValue -InputObject $Extension -Name "extensionName")

        # Exact comparison (case-insensitive). No EndsWith/Contains,
        # so that only our own extension counts as 'managed'.
        if (
            $Id   -eq $ExtensionId -or
            $Id   -eq $ExtensionName -or
            $Name -eq $ExtensionName
        ) {
            return $Extension
        }
    }

    return $null
}


function Get-MailboxContacts {
    # Reads ALL contacts in the default contacts folder of the mailbox,
    # with only our own open extension expanded.
    # Returns per contact: Contact, Extension, SourceId, IsManaged.
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    $UserEncoded = [uri]::EscapeDataString($UserPrincipalName.Trim())

    # OData: string literal between single quotes, double any single quote.
    $ExtensionIdLiteral = $ExtensionId.Replace("'", "''")

    $ExpandFilter = [uri]::EscapeDataString(
        "id eq '$ExtensionIdLiteral'"
    )

    $Uri =
        "https://graph.microsoft.com/v1.0/users/$UserEncoded/contacts" +
        "?`$select=id,displayName,businessPhones" +
        "&`$expand=extensions(`$filter=$ExpandFilter)" +
        "&`$top=100"

    $Contacts = @(Get-GraphPaged -Uri $Uri)

    $Result = @()

    foreach ($Contact in $Contacts) {

        if ($null -eq $Contact) {
            continue
        }

        $Extension = Get-CentralExtension -Contact $Contact

        $SourceId = ""

        if ($null -ne $Extension) {
            $SourceId = ([string](Get-PropertyValue -InputObject $Extension -Name "sourceId")).Trim()
        }

        $Result += [PSCustomObject]@{
            Contact   = $Contact
            Extension = $Extension
            SourceId  = $SourceId
            IsManaged = ($null -ne $Extension)
        }
    }

    return @($Result)
}


function Get-ManagedContacts {
    # Only contacts with our extension ($ExtensionName).
    # UserPrincipalName may also be an Entra ID object id.
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    return @(
        Get-MailboxContacts -UserPrincipalName $UserPrincipalName |
        Where-Object { $_.IsManaged }
    )
}


function New-CentralContact {
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        $Source
    )

    $UserEncoded = [uri]::EscapeDataString($UserPrincipalName)

    $Body = @{
        displayName = $Source.Name

        businessPhones = @(
            $Source.Phone
        )

        extensions = @(
            @{
                "@odata.type" = "microsoft.graph.openTypeExtension"
                extensionName = $ExtensionName
                sourceId      = $Source.SourceId
                sourceSystem  = $ExtensionSourceSystem
            }
        )
    } | ConvertTo-Json -Depth 10

    Invoke-MgGraphRequest `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserEncoded/contacts" `
        -Body $Body `
        -ContentType "application/json" |
        Out-Null
}


function Update-CentralContact {
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$ContactId,

        [Parameter(Mandatory)]
        $Source
    )

    $UserEncoded    = [uri]::EscapeDataString($UserPrincipalName)
    $ContactEncoded = [uri]::EscapeDataString($ContactId)

    $Body = @{
        displayName = $Source.Name

        businessPhones = @(
            $Source.Phone
        )
    } | ConvertTo-Json -Depth 5

    Invoke-MgGraphRequest `
        -Method PATCH `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserEncoded/contacts/$ContactEncoded" `
        -Body $Body `
        -ContentType "application/json" |
        Out-Null
}


function Get-FirstPhone {
    param(
        $Contact
    )

    $Phones = @(Get-PropertyValue -InputObject $Contact -Name "businessPhones")

    if ($Phones.Count -gt 0 -and $null -ne $Phones[0]) {
        return ([string]$Phones[0]).Trim()
    }

    return ""
}


function Get-ContactKey {
    # Name + number, case-insensitive. Only for merging personal
    # copies; managed contacts are always matched via sourceId.
    param(
        $Name,
        $Phone
    )

    return ("{0}|{1}" -f ([string]$Name).Trim(), ([string]$Phone).Trim()).ToLowerInvariant()
}


function Add-CentralExtension {
    # Turns an existing (personal) contact into a managed contact.
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$ContactId,

        [Parameter(Mandatory)]
        $Source
    )

    $UserEncoded    = [uri]::EscapeDataString($UserPrincipalName)
    $ContactEncoded = [uri]::EscapeDataString($ContactId)

    $Body = @{
        "@odata.type" = "microsoft.graph.openTypeExtension"
        extensionName = $ExtensionName
        sourceId      = $Source.SourceId
        sourceSystem  = $ExtensionSourceSystem
    } | ConvertTo-Json -Depth 5

    Invoke-MgGraphRequest `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserEncoded/contacts/$ContactEncoded/extensions" `
        -Body $Body `
        -ContentType "application/json" |
        Out-Null
}


function Remove-CentralContact {
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$ContactId
    )

    $UserEncoded    = [uri]::EscapeDataString($UserPrincipalName)
    $ContactEncoded = [uri]::EscapeDataString($ContactId)

    Invoke-MgGraphRequest `
        -Method DELETE `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserEncoded/contacts/$ContactEncoded" |
        Out-Null
}


# ============================================================
# MAILBOX SYNC
# ============================================================

function Sync-UserContacts {
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$DisplayName,

        [Parameter(Mandatory)]
        [array]$SourceContacts,

        [Parameter(Mandatory)]
        [hashtable]$SourceById,

        # SourceIds of SharePoint items that are temporarily invalid
        # (empty name/number). Their contacts are NOT deleted.
        [hashtable]$ProtectedSourceIds = @{},

        [Parameter(Mandatory)]
        [bool]$DryRun
    )

    $UPN = $UserPrincipalName.Trim()

    $Result = [ordered]@{
        User      = $UPN
        Adopt     = 0
        Create    = 0
        Update    = 0
        Delete    = 0
        DupDelete = 0
        Unchanged = 0
        Status    = "UNKNOWN"
    }

    Write-Log "Contact sync: $DisplayName <$UPN>" "INFO"

    try {
        $AllContacts = @(
            Get-MailboxContacts `
                -UserPrincipalName $UPN
        )

        $Managed   = @($AllContacts | Where-Object { $_.IsManaged })
        $Unmanaged = @($AllContacts | Where-Object { -not $_.IsManaged })
    }
    catch {

        $Message = Get-GraphErrorText -ErrorRecord $_

        if ($Message -match "403|Forbidden|AccessDenied") {

            Write-Log "$UPN does not have effective Graph mailbox access yet. Probably RBAC propagation." "WARN"

            $Result.Status = "RBAC pending / 403"

            return [PSCustomObject]$Result
        }

        Write-Log "Reading contacts failed for $UPN : $Message" "ERROR"

        $Result.Status = "ERROR reading contacts"

        return [PSCustomObject]$Result
    }


    # --------------------------------------------------------
    # DUPLICATE SOURCE-ID CHECK
    # --------------------------------------------------------

    $Duplicates = @(
        $Managed |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace(
                [string]$_.SourceId
            )
        } |
        Group-Object SourceId |
        Where-Object Count -gt 1
    )

    if ($Duplicates.Count -gt 0) {

        Write-Log "$UPN contains duplicate centrally managed SourceIds. Mailbox is skipped." "ERROR"

        $Result.Status = "ERROR duplicate SourceId"

        return [PSCustomObject]$Result
    }


    # --------------------------------------------------------
    # MANAGED CONTACTS WITHOUT SOURCE-ID
    # --------------------------------------------------------

    # Our extension without a usable sourceId is inconsistent.
    # Such contacts are never changed/deleted automatically.
    $ManagedWithoutSourceId = @(
        $Managed |
        Where-Object {
            [string]::IsNullOrWhiteSpace([string]$_.SourceId)
        }
    )

    if ($ManagedWithoutSourceId.Count -gt 0) {

        Write-Log "$UPN has $($ManagedWithoutSourceId.Count) managed contact(s) without sourceId. These are ignored." "WARN"
    }


    # --------------------------------------------------------
    # PERSONAL COPIES OF SOURCE CONTACTS (MERGE CANDIDATES)
    # --------------------------------------------------------

    # Key = name + first business number (case-insensitive,
    # otherwise exact). Only used to take over personal copies of
    # source contacts, never to identify existing managed contacts
    # (that is always and only sourceId).
    $SourceKeys = @{}

    foreach ($Source in $SourceContacts) {
        $SourceKeys[(Get-ContactKey -Name $Source.Name -Phone $Source.Phone)] = $true
    }

    $UnmanagedByKey = @{}
    $Lookalikes     = 0

    foreach ($Item in $Unmanaged) {

        $Key = Get-ContactKey `
            -Name  (Get-PropertyValue -InputObject $Item.Contact -Name "displayName") `
            -Phone (Get-FirstPhone -Contact $Item.Contact)

        if (-not $SourceKeys.ContainsKey($Key)) {
            continue
        }

        $Lookalikes++

        if (-not $UnmanagedByKey.ContainsKey($Key)) {
            $UnmanagedByKey[$Key] = [System.Collections.Generic.List[object]]::new()
        }

        $UnmanagedByKey[$Key].Add($Item.Contact)
    }


    # --------------------------------------------------------
    # SAFETY: OUR CONTACTS WITHOUT A RECOGNISED EXTENSION?
    # --------------------------------------------------------

    # If extension detection fails (e.g. a wrong extension id in the
    # $expand filter), almost the whole source appears to be present as
    # personal copies in the mailbox. Merge would then delete our own
    # contacts as 'duplicates'. A user with a few personally saved
    # colleagues/suppliers stays well below this limit.
    # Ratio on UNIQUE source items (a user may have saved the same
    # contact several times).
    $LookalikeRatio = $UnmanagedByKey.Count / [double]$SourceContacts.Count

    if ($LookalikeRatio -ge $MaxLookalikeRatio) {

        Write-Log (
            "$UPN has $Lookalikes contacts ($($UnmanagedByKey.Count) unique source items) WITHOUT our extension with the same name+number as the source " +
            "($([math]::Round($LookalikeRatio * 100))% of the source, limit $([math]::Round($MaxLookalikeRatio * 100))%). " +
            "Extension detection may be failing. Mailbox skipped."
        ) "ERROR"

        $Result.Status = "SAFETY STOP lookalikes ($Lookalikes)"

        return [PSCustomObject]$Result
    }


    # --------------------------------------------------------
    # INDEX EXISTING CONTACTS
    # --------------------------------------------------------

    $ExistingBySourceId = @{}

    foreach ($ManagedContact in $Managed) {

        if (
            -not [string]::IsNullOrWhiteSpace(
                [string]$ManagedContact.SourceId
            )
        ) {
            $ExistingBySourceId[
                [string]$ManagedContact.SourceId
            ] = $ManagedContact.Contact
        }
    }


    # --------------------------------------------------------
    # PLAN: ADOPT / CREATE / UPDATE BEFORE WE CHANGE ANYTHING
    # --------------------------------------------------------

    $ToAdopt  = @()
    $ToCreate = @()
    $ToUpdate = @()
    $UsedKeys = @{}

    foreach ($Source in $SourceContacts) {

        $Existing = $ExistingBySourceId[
            [string]$Source.SourceId
        ]

        if ($null -ne $Existing) {

            $ExistingName  = ([string](Get-PropertyValue -InputObject $Existing -Name "displayName")).Trim()
            $ExistingPhone = Get-FirstPhone -Contact $Existing

            # -cne: a change in capitalisation only is also a name change.
            if (
                ($ExistingName -cne [string]$Source.Name) -or
                ($ExistingPhone -ne [string]$Source.Phone)
            ) {
                $ToUpdate += [PSCustomObject]@{
                    Source    = $Source
                    ContactId = [string](Get-PropertyValue -InputObject $Existing -Name "id")
                }
            }
            else {
                $Result.Unchanged++
            }

            continue
        }


        # No managed contact yet: take over a personal copy?
        $Key = Get-ContactKey -Name $Source.Name -Phone $Source.Phone

        if (
            $AdoptPersonalDuplicates -and
            $UnmanagedByKey.ContainsKey($Key) -and
            $UnmanagedByKey[$Key].Count -gt 0
        ) {

            $Pool = $UnmanagedByKey[$Key]
            $Pick = $null

            # Prefer a DIFFERENT spelling (capitalisation): that one was
            # entered by the user. That contact is kept and gets our
            # naming.
            foreach ($Candidate in $Pool) {

                $CandidateName = ([string](Get-PropertyValue -InputObject $Candidate -Name "displayName")).Trim()

                if ($CandidateName -cne [string]$Source.Name) {
                    $Pick = $Candidate
                    break
                }
            }

            if ($null -eq $Pick) {
                $Pick = $Pool[0]
            }

            [void]$Pool.Remove($Pick)
            $UsedKeys[$Key] = [string]$Source.Name

            $PickName = ([string](Get-PropertyValue -InputObject $Pick -Name "displayName")).Trim()

            $ToAdopt += [PSCustomObject]@{
                Source      = $Source
                ContactId   = [string](Get-PropertyValue -InputObject $Pick -Name "id")
                CurrentName = $PickName
                NeedsUpdate = (
                    ($PickName -cne [string]$Source.Name) -or
                    ((Get-FirstPhone -Contact $Pick) -ne [string]$Source.Phone)
                )
            }

            continue
        }

        $ToCreate += $Source
    }


    # Remaining copies of a contact we are taking over. Only copies
    # with EXACTLY our spelling are deleted; a different spelling was
    # entered by the user and is always kept.
    $PersonalDuplicatesToDelete = @()

    if ($AdoptPersonalDuplicates -and $DeleteExtraPersonalDuplicates) {

        foreach ($Key in @($UsedKeys.Keys)) {

            foreach ($Contact in $UnmanagedByKey[$Key]) {

                $ContactName = ([string](Get-PropertyValue -InputObject $Contact -Name "displayName")).Trim()

                if ($ContactName -ceq $UsedKeys[$Key]) {
                    $PersonalDuplicatesToDelete += $Contact
                }
            }
        }
    }


    # --------------------------------------------------------
    # DETERMINE DELETE SET BEFORE WE CHANGE ANYTHING
    # --------------------------------------------------------

    $ContactsToDelete = @()

    foreach ($ManagedContact in $Managed) {

        $SourceId = [string]$ManagedContact.SourceId

        if ([string]::IsNullOrWhiteSpace($SourceId)) {
            continue
        }

        # Temporarily invalid SharePoint item: keep the contact.
        if ($ProtectedSourceIds.ContainsKey($SourceId)) {
            continue
        }

        if (-not $SourceById.ContainsKey($SourceId)) {
            $ContactsToDelete += $ManagedContact.Contact
        }
    }


    if ($ContactsToDelete.Count -gt $MaxDeletesPerMailbox) {

        Write-Log "$UPN would delete $($ContactsToDelete.Count) contacts. Maximum is $MaxDeletesPerMailbox. Mailbox skipped entirely." "ERROR"

        $Result.Status = "DELETE SAFETY STOP"

        return [PSCustomObject]$Result
    }

    if ($PersonalDuplicatesToDelete.Count -gt $MaxPersonalDuplicateDeletesPerMailbox) {

        Write-Log "$UPN would delete $($PersonalDuplicatesToDelete.Count) personal duplicates. Maximum is $MaxPersonalDuplicateDeletesPerMailbox. Mailbox skipped entirely." "ERROR"

        $Result.Status = "DUPLICATE DELETE SAFETY STOP"

        return [PSCustomObject]$Result
    }


    # --------------------------------------------------------
    # EXECUTE
    # --------------------------------------------------------

    try {

        # 1. ADOPT: personal copy becomes a managed contact.
        foreach ($Adopt in $ToAdopt) {

            $Result.Adopt++

            $Suffix = if ($Adopt.NeedsUpdate) { " (name/number will be aligned)" } else { "" }

            Write-Log "[ADOPT][$UPN] '$($Adopt.CurrentName)' -> sourceId $($Adopt.Source.SourceId) $($Adopt.Source.Name) - $($Adopt.Source.Phone)$Suffix"

            if (-not $DryRun) {

                Add-CentralExtension `
                    -UserPrincipalName $UPN `
                    -ContactId $Adopt.ContactId `
                    -Source $Adopt.Source

                if ($Adopt.NeedsUpdate) {

                    Update-CentralContact `
                        -UserPrincipalName $UPN `
                        -ContactId $Adopt.ContactId `
                        -Source $Adopt.Source
                }
            }
        }


        # 2. CREATE
        foreach ($Source in $ToCreate) {

            $Result.Create++

            Write-Log "[CREATE][$UPN] $($Source.Name) - $($Source.Phone)"

            if (-not $DryRun) {

                New-CentralContact `
                    -UserPrincipalName $UPN `
                    -Source $Source
            }
        }


        # 3. UPDATE
        foreach ($Update in $ToUpdate) {

            $Result.Update++

            Write-Log "[UPDATE][$UPN] $($Update.Source.Name) - $($Update.Source.Phone)"

            if (-not $DryRun) {

                Update-CentralContact `
                    -UserPrincipalName $UPN `
                    -ContactId $Update.ContactId `
                    -Source $Update.Source
            }
        }


        # 4. DELETE managed contacts that are no longer in SharePoint
        foreach ($Contact in $ContactsToDelete) {

            $Result.Delete++

            Write-Log "[DELETE][$UPN] $(Get-PropertyValue -InputObject $Contact -Name 'displayName')"

            if (-not $DryRun) {

                Remove-CentralContact `
                    -UserPrincipalName $UPN `
                    -ContactId ([string](Get-PropertyValue -InputObject $Contact -Name "id"))
            }
        }


        # 5. DUPLICATE DELETE: only AFTER successful adopts (on an error
        #    above we never get here).
        foreach ($Contact in $PersonalDuplicatesToDelete) {

            $Result.DupDelete++

            Write-Log "[DUPDELETE][$UPN] personal copy '$(Get-PropertyValue -InputObject $Contact -Name 'displayName')' - $(Get-FirstPhone -Contact $Contact)"

            if (-not $DryRun) {

                Remove-CentralContact `
                    -UserPrincipalName $UPN `
                    -ContactId ([string](Get-PropertyValue -InputObject $Contact -Name "id"))
            }
        }


        $Result.Status = "OK"
    }
    catch {

        Write-Log "Synchronisation of $UPN partially failed: $(Get-GraphErrorText -ErrorRecord $_)" "ERROR"

        $Result.Status = "ERROR during sync"
    }


    return [PSCustomObject]$Result
}


if ($FunctionsOnly) {

    Write-Log "FunctionsOnly: config + functions loaded, no sync performed." "INFO"

    return
}


# ============================================================
# 0. CHECK CONNECTIONS
# ============================================================

Write-Console ""
Write-Console "============================================" -ForegroundColor Cyan
Write-Console "          CENTRAL PHONEBOOK SYNC" -ForegroundColor Cyan
Write-Console "============================================" -ForegroundColor Cyan
Write-Console ""

Write-Log "Checking connections. DryRun = $DryRun, ManagedIdentity = $ManagedIdentity."


if ($ManagedIdentity) {

    Write-Log "Connecting to Microsoft Graph using the managed identity."

    Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop

    Write-Log "Connecting to Exchange Online using the managed identity ($ExchangeOrganization)."

    Connect-ExchangeOnline `
        -ManagedIdentity `
        -Organization $ExchangeOrganization `
        -ShowBanner:$false `
        -ErrorAction Stop
}


$MgContext = Get-MgContext

if (-not $MgContext) {
    throw "No Microsoft Graph session present."
}

# App-only / identity auth only. Get-MgContext documents AuthType
# as Delegated or AppOnly; a managed identity may be named differently
# depending on the module version. Therefore only Delegated is refused.
if ([string]$MgContext.AuthType -eq "Delegated") {
    throw "Graph is connected Delegated. Required: app-only or managed identity."
}


$RequiredGraphPermissions = @(
    "GroupMember.ReadBasic.All",
    "User.ReadBasic.All",
    "Sites.Selected"
)

$ContextScopes = @($MgContext.Scopes | Where-Object { $_ })

if ($ContextScopes.Count -gt 0) {

    foreach ($Permission in $RequiredGraphPermissions) {

        if ($Permission -notin $ContextScopes) {
            throw "Graph token is missing required permission: $Permission"
        }
    }
}
else {

    # With a managed identity the module does not always fill Scopes.
    # Missing permissions then lead to a Graph error further on (fail closed).
    Write-Log "Graph context contains no scope list; permission check skipped." "WARN"
}


Write-Log "Graph connected ($($MgContext.AuthType)) with app $($MgContext.ClientId)." "SUCCESS"


# Get-ConnectionInformation requires no Exchange RBAC role
# (Get-OrganizationConfig does; the managed identity deliberately lacks it).
$ExoConnections = @(
    Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { [string]$_.State -eq "Connected" }
)

if ($ExoConnections.Count -eq 0) {
    throw "No working Exchange Online session present."
}

Write-Log "Exchange Online session active ($($ExoConnections[0].Organization))." "SUCCESS"


# ============================================================
# 1. READ ENTRA ID TARGET GROUP
# ============================================================

Write-Console ""
Write-Console "=== 1. ENTRA ID TARGET GROUP ===" -ForegroundColor Cyan


$GroupMembers = @(
    Get-MgGroupMemberAsUser `
        -GroupId $TeamsPhoneGroupId `
        -All `
        -Property Id,DisplayName,UserPrincipalName,Mail
)


Write-Log "$($GroupMembers.Count) user objects found in target group $TeamsPhoneGroupId."


# ============================================================
# 2. ONLY REAL EXCHANGE USERMAILBOXES
# ============================================================

Write-Console ""
Write-Console "=== 2. VALID MAILBOX USERS ===" -ForegroundColor Cyan


$ValidMembers = @()
$SkippedMembers = @()


foreach ($Member in $GroupMembers) {

    $UPN = [string]$Member.UserPrincipalName

    if ([string]::IsNullOrWhiteSpace($UPN)) {

        $SkippedMembers += [PSCustomObject]@{
            DisplayName       = $Member.DisplayName
            UserPrincipalName = $null
            Reason            = "No UPN"
        }

        continue
    }


    # FAIL CLOSED: with -Filter, "no mailbox" returns an empty
    # result, not an exception. Every exception is therefore a real
    # error (throttling, expired session, ...) and aborts the run.
    # Treating every error as "no mailbox" would let a temporary
    # Exchange error turn a valid user into a former user
    # (= contacts deleted + removed from the RBAC scope).
    $UPNFilterValue = $UPN.Replace("'", "''")

    try {

        $Mailboxes = @(
            Get-EXOMailbox `
                -Filter "UserPrincipalName -eq '$UPNFilterValue'" `
                -PropertySets Minimum `
                -ResultSize 2 `
                -ErrorAction Stop
        )
    }
    catch {

        throw "Safety stop: mailbox check for $UPN failed. Run aborted. $($_.Exception.Message)"
    }


    if ($Mailboxes.Count -eq 0) {

        $SkippedMembers += [PSCustomObject]@{
            DisplayName       = $Member.DisplayName
            UserPrincipalName = $UPN
            Reason            = "No Exchange mailbox"
        }

        continue
    }

    if ($Mailboxes.Count -gt 1) {
        throw "Safety stop: multiple mailboxes found for UPN $UPN."
    }


    $Mailbox = $Mailboxes[0]

    if ([string]$Mailbox.RecipientTypeDetails -ne "UserMailbox") {

        $SkippedMembers += [PSCustomObject]@{
            DisplayName       = $Member.DisplayName
            UserPrincipalName = $UPN
            Reason            = "RecipientTypeDetails = $($Mailbox.RecipientTypeDetails)"
        }

        continue
    }


    # Stable key = Entra ID object id. Exchange knows it as
    # ExternalDirectoryObjectId. Must be exactly equal to the Graph id.
    $MemberId  = ([string]$Member.Id).Trim().ToLowerInvariant()
    $MailboxId = ([string](Get-PropertyValue -InputObject $Mailbox -Name "ExternalDirectoryObjectId")).Trim().ToLowerInvariant()

    if ([string]::IsNullOrWhiteSpace($MemberId) -or $MemberId -ne $MailboxId) {
        throw "Safety stop: Entra ID ($MemberId) and Exchange ExternalDirectoryObjectId ($MailboxId) do not match for $UPN."
    }


    $ValidMembers += [PSCustomObject]@{
        Id                 = $MemberId
        DisplayName        = [string]$Member.DisplayName
        UserPrincipalName  = $UPN.Trim()
        Mail               = [string]$Member.Mail
        PrimarySmtpAddress = [string]$Mailbox.PrimarySmtpAddress
    }
}


Write-Log "Valid UserMailboxes: $($ValidMembers.Count)."
Write-Log "Skipped objects: $($SkippedMembers.Count)."


if ($ValidMembers.Count -lt $MinimumTargetUsers) {

    throw (
        "Safety stop: only {0} valid target users found. Minimum is {1}." -f
        $ValidMembers.Count,
        $MinimumTargetUsers
    )
}


$ValidMembers |
    Sort-Object DisplayName |
    Select-Object DisplayName,UserPrincipalName,Mail |
    Format-Table -AutoSize


if ($SkippedMembers.Count -gt 0) {

    Write-Console ""
    Write-Console "Skipped objects:" -ForegroundColor DarkGray

    $SkippedMembers |
        Sort-Object DisplayName |
        Format-Table DisplayName,UserPrincipalName,Reason -AutoSize
}


# ============================================================
# 3. READ AND VALIDATE SHAREPOINT MASTER SOURCE
# ============================================================

Write-Console ""
Write-Console "=== 3. CENTRAL PHONEBOOK ===" -ForegroundColor Cyan


$Site = Invoke-MgGraphRequest `
    -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/sites/$SharePointHost`:$SharePointPath"


$Lists = @(
    Get-GraphPaged `
        -Uri "https://graph.microsoft.com/v1.0/sites/$($Site.id)/lists"
)


$List = $Lists |
    Where-Object {
        $_.displayName -eq $ListName
    } |
    Select-Object -First 1


if (-not $List) {
    throw "Microsoft List '$ListName' not found."
}


$Columns = @(
    Get-GraphPaged `
        -Uri "https://graph.microsoft.com/v1.0/sites/$($Site.id)/lists/$($List.id)/columns"
)


$NameColumn = $Columns |
    Where-Object {
        $_.displayName -eq $NameColumnName
    } |
    Select-Object -First 1


$PhoneColumn = $Columns |
    Where-Object {
        $_.displayName -eq $PhoneColumnName
    } |
    Select-Object -First 1


if (-not $NameColumn) {
    throw "Column '$NameColumnName' not found."
}

if (-not $PhoneColumn) {
    throw "Column '$PhoneColumnName' not found."
}


$NameField  = [string]$NameColumn.name
$PhoneField = [string]$PhoneColumn.name


$ListItemsUri = (
    "https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items?`$expand=fields&`$top=200" `
    -f $Site.id, $List.id
)

$ListItems = @(
    Get-GraphPaged -Uri $ListItemsUri
)


$SourceContacts = @()
$InvalidSourceItems = @()


foreach ($Item in $ListItems) {

    $Name = (
        [string]$Item.fields.$NameField
    ).Trim()

    $Phone = (
        [string]$Item.fields.$PhoneField
    ).Trim()


    if (
        [string]::IsNullOrWhiteSpace($Name) -or
        [string]::IsNullOrWhiteSpace($Phone)
    ) {

        $InvalidSourceItems += [PSCustomObject]@{
            ID     = $Item.id
            Name   = $Name
            Number = $Phone
        }

        continue
    }


    $SourceContacts += [PSCustomObject]@{
        SourceId = [string]$Item.id
        Name     = $Name
        Phone    = $Phone
    }
}


if ($SourceContacts.Count -lt $MinimumSourceContacts) {

    throw (
        "Safety stop: central phonebook contains only {0} valid contacts. Minimum is {1}." -f
        $SourceContacts.Count,
        $MinimumSourceContacts
    )
}


Write-Log "Central source contains $($SourceContacts.Count) valid contacts." "SUCCESS"


if ($InvalidSourceItems.Count -gt 0) {

    Write-Log "$($InvalidSourceItems.Count) invalid SharePoint items are skipped." "WARN"

    $InvalidSourceItems |
        Format-Table ID,Name,Number -AutoSize
}


$SourceById = @{}

foreach ($Source in $SourceContacts) {

    $SourceById[
        [string]$Source.SourceId
    ] = $Source
}


# Items that are (temporarily) empty do not count as "deleted from
# SharePoint". Existing contacts for them are kept.
$ProtectedSourceIds = @{}

foreach ($Invalid in $InvalidSourceItems) {

    $ProtectedSourceIds[
        [string]$Invalid.ID
    ] = $true
}


# ============================================================
# 4. CURRENT EXCHANGE RBAC SCOPE GROUP
# ============================================================

Write-Console ""
Write-Console "=== 4. EXCHANGE RBAC SCOPE GROUP ===" -ForegroundColor Cyan


# Comparison on Entra ID object id (ExternalDirectoryObjectId), not on
# UPN vs PrimarySmtpAddress. Those two can differ; a valid user would
# then count as "new" AND as "former user" on every run.
function Get-ScopeMemberInfo {

    $Members = @(
        Get-DistributionGroupMember `
            -Identity $ScopeGroup `
            -ResultSize Unlimited `
            -ErrorAction Stop
    )

    foreach ($ScopeMember in $Members) {

        [PSCustomObject]@{
            Key                  = ([string](Get-PropertyValue -InputObject $ScopeMember -Name "ExternalDirectoryObjectId")).Trim().ToLowerInvariant()
            DisplayName          = [string](Get-PropertyValue -InputObject $ScopeMember -Name "DisplayName")
            PrimarySmtpAddress   = ([string](Get-PropertyValue -InputObject $ScopeMember -Name "PrimarySmtpAddress")).Trim()
            RecipientTypeDetails = [string](Get-PropertyValue -InputObject $ScopeMember -Name "RecipientTypeDetails")
        }
    }
}


$CurrentScopeMembers = @(Get-ScopeMemberInfo)


$ScopeMembersWithoutKey = @(
    $CurrentScopeMembers |
    Where-Object { [string]::IsNullOrWhiteSpace($_.Key) }
)

if ($ScopeMembersWithoutKey.Count -gt 0) {

    # Fail closed: without a stable key never clean up as a former user.
    Write-Log "$($ScopeMembersWithoutKey.Count) scope member(s) without ExternalDirectoryObjectId. These are not touched." "WARN"

    $ScopeMembersWithoutKey |
        Format-Table DisplayName,PrimarySmtpAddress,RecipientTypeDetails -AutoSize
}


$TargetKeys = @(
    $ValidMembers |
    ForEach-Object { $_.Id }
)


$CurrentScopeKeys = @(
    $CurrentScopeMembers |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_.Key) } |
    ForEach-Object { $_.Key }
)


$ScopeToAdd = @(
    $ValidMembers |
    Where-Object { $_.Id -notin $CurrentScopeKeys }
)


$FormerUsers = @(
    $CurrentScopeMembers |
    Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.Key) -and
        $_.Key -notin $TargetKeys
    }
)


Write-Log "New scope members needed: $($ScopeToAdd.Count)."
Write-Log "Former target users found: $($FormerUsers.Count)."


if ($FormerUsers.Count -gt $MaxFormerUsersPerRun) {

    throw (
        "Safety stop: {0} users would leave the target group. Maximum per run is {1}." -f
        $FormerUsers.Count,
        $MaxFormerUsersPerRun
    )
}


if ($ScopeToAdd.Count -gt 0) {

    Write-Console ""
    Write-Console "To be added to RBAC scope:" -ForegroundColor Yellow

    $ScopeToAdd |
        Select-Object DisplayName,UserPrincipalName |
        Format-Table -AutoSize
}


if ($FormerUsers.Count -gt 0) {

    Write-Console ""
    Write-Console "Left the target group:" -ForegroundColor Yellow

    $FormerUsers |
        Select-Object DisplayName,PrimarySmtpAddress |
        Format-Table -AutoSize
}


# ============================================================
# 5. ADD NEW USERS TO RBAC SCOPE
# ============================================================

Write-Console ""
Write-Console "=== 5. UPDATE RBAC SCOPE ===" -ForegroundColor Cyan


$ScopeAddErrors = 0

foreach ($User in $ScopeToAdd) {

    Write-Log (
        "[SCOPE ADD] {0} <{1}>" -f
        $User.DisplayName,
        $User.UserPrincipalName
    ) "INFO"


    if (-not $DryRun) {

        # Caught per user: a failed add does not block the whole
        # run. The user is then not in the re-read scope and gets
        # status RBAC PROPAGATION PENDING in step 6.
        try {

            # Bypass: the managed identity is not ManagedBy of the group.
            # Allowed via a custom role (only Add/Remove-DistributionGroupMember)
            # whose write scope matches only the scope group.
            Add-DistributionGroupMember `
                -Identity $ScopeGroup `
                -Member $User.PrimarySmtpAddress `
                -BypassSecurityGroupManagerCheck `
                -ErrorAction Stop
        }
        catch {

            $ScopeAddErrors++

            Write-Log "[SCOPE ADD] failed for $($User.UserPrincipalName): $($_.Exception.Message)" "ERROR"
        }
    }
}


# In live mode re-read so that our local state is correct.
if (
    -not $DryRun -and
    $ScopeToAdd.Count -gt 0
) {

    $CurrentScopeMembers = @(Get-ScopeMemberInfo)

    $CurrentScopeKeys = @(
        $CurrentScopeMembers |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.Key) } |
        ForEach-Object { $_.Key }
    )
}


# ============================================================
# 6. CONTACT SYNC FOR ACTIVE TARGET USERS
# ============================================================

Write-Console ""
Write-Console "=== 6. CONTACT SYNC ACTIVE USERS ===" -ForegroundColor Cyan


$Results = @()


foreach ($User in $ValidMembers) {

    $UPN = (
        [string]$User.UserPrincipalName
    ).Trim()

    # During a dry run a new scope user is of course not actually
    # added. We cannot read that mailbox via Application RBAC yet.
    if ($User.Id -notin $CurrentScopeKeys) {

        $Status = if ($DryRun) {
            "SCOPE ADD REQUIRED"
        }
        else {
            "RBAC PROPAGATION PENDING"
        }


        Write-Log "$UPN : $Status" "WARN"


        $Results += [PSCustomObject]@{
            User      = $UPN
            Adopt     = 0
            Create    = 0
            Update    = 0
            Delete    = 0
            DupDelete = 0
            Unchanged = 0
            Status    = $Status
        }

        continue
    }


    $SyncResult = Sync-UserContacts `
        -UserPrincipalName $UPN `
        -DisplayName $User.DisplayName `
        -SourceContacts $SourceContacts `
        -SourceById $SourceById `
        -ProtectedSourceIds $ProtectedSourceIds `
        -DryRun $DryRun


    $Results += $SyncResult
}


# ============================================================
# 7. CLEAN UP FORMER TARGET USERS
# ============================================================

Write-Console ""
Write-Console "=== 7. CLEAN UP FORMER TARGET USERS ===" -ForegroundColor Cyan


$FormerUserResults = @()


foreach ($Former in $FormerUsers) {

    # Display by mail address, Graph calls by Entra ID object id:
    # /users/{PrimarySmtpAddress} does not work if SMTP differs from UPN.
    $UPN      = $Former.PrimarySmtpAddress
    $GraphId  = $Former.Key


    Write-Log "Cleaning up former target user: $UPN ($GraphId)" "WARN"


    $CleanupStatus = "UNKNOWN"
    $ManagedCount  = 0


    try {

        $Managed = @(
            Get-ManagedContacts `
                -UserPrincipalName $GraphId
        )


        $ManagedCount = $Managed.Count


        Write-Log "$UPN has $ManagedCount centrally managed contacts."


        if ($ManagedCount -gt $MaxFormerUserCleanupDeletes) {

            Write-Log (
                "$UPN cleanup stopped: $ManagedCount contacts found; " +
                "maximum is $MaxFormerUserCleanupDeletes."
            ) "ERROR"


            $CleanupStatus = "CLEANUP SAFETY STOP"


            $FormerUserResults += [PSCustomObject]@{
                User            = $UPN
                ManagedContacts = $ManagedCount
                ScopeRemoved    = $false
                Status          = $CleanupStatus
            }

            continue
        }


        if (-not $DryRun) {

            foreach ($ManagedContact in $Managed) {

                Remove-CentralContact `
                    -UserPrincipalName $GraphId `
                    -ContactId ([string](Get-PropertyValue -InputObject $ManagedContact.Contact -Name "id"))
            }

            # Check: only remove from RBAC scope if really nothing is left.
            $Remaining = @(
                Get-ManagedContacts `
                    -UserPrincipalName $GraphId
            )

            if ($Remaining.Count -gt 0) {
                throw "After cleanup $($Remaining.Count) managed contact(s) still present."
            }
        }


        # Only remove from RBAC if cleanup was successful.
        if (-not $DryRun) {

            $RemoveMember = if ([string]::IsNullOrWhiteSpace($Former.PrimarySmtpAddress)) {
                $GraphId
            }
            else {
                $Former.PrimarySmtpAddress
            }

            Remove-DistributionGroupMember `
                -Identity $ScopeGroup `
                -Member $RemoveMember `
                -BypassSecurityGroupManagerCheck `
                -Confirm:$false `
                -ErrorAction Stop
        }


        $CleanupStatus = if ($DryRun) {
            "WOULD CLEANUP + REMOVE SCOPE"
        }
        else {
            "CLEANED + REMOVED SCOPE"
        }


        Write-Log "$UPN : $CleanupStatus" "SUCCESS"


        $FormerUserResults += [PSCustomObject]@{
            User            = $UPN
            ManagedContacts = $ManagedCount
            ScopeRemoved    = (-not $DryRun)
            Status          = $CleanupStatus
        }
    }
    catch {

        Write-Log (
            "Cleanup failed for $UPN. " +
            "User stays in RBAC scope. " +
            (Get-GraphErrorText -ErrorRecord $_)
        ) "ERROR"


        $FormerUserResults += [PSCustomObject]@{
            User            = $UPN
            ManagedContacts = $ManagedCount
            ScopeRemoved    = $false
            Status          = "ERROR CLEANUP"
        }
    }
}


# ============================================================
# 8. RESULTS
# ============================================================

Write-Console ""
Write-Console "============================================" -ForegroundColor Cyan
Write-Console "                SYNC RESULT" -ForegroundColor Cyan
Write-Console "============================================" -ForegroundColor Cyan
Write-Console ""


if ($Results.Count -gt 0) {

    $Results |
        Sort-Object User |
        Format-Table `
            User,
            Adopt,
            Create,
            Update,
            Delete,
            DupDelete,
            Unchanged,
            Status `
            -AutoSize
}


if ($FormerUserResults.Count -gt 0) {

    Write-Console ""
    Write-Console "Former target users:" -ForegroundColor Yellow

    $FormerUserResults |
        Format-Table `
            User,
            ManagedContacts,
            ScopeRemoved,
            Status `
            -AutoSize
}


# ============================================================
# 9. TOTALS
# ============================================================

$TotalCreate = (
    $Results |
    Measure-Object Create -Sum
).Sum

$TotalUpdate = (
    $Results |
    Measure-Object Update -Sum
).Sum

$TotalDelete = (
    $Results |
    Measure-Object Delete -Sum
).Sum

$TotalUnchanged = (
    $Results |
    Measure-Object Unchanged -Sum
).Sum

$TotalAdopt = (
    $Results |
    Measure-Object Adopt -Sum
).Sum

$TotalDupDelete = (
    $Results |
    Measure-Object DupDelete -Sum
).Sum


Write-Console ""
Write-Console "============================================" -ForegroundColor Cyan
Write-Console "                  TOTALS" -ForegroundColor Cyan
Write-Console "============================================" -ForegroundColor Cyan

Write-Console "Source contacts        : $($SourceContacts.Count)"
Write-Console "Entra ID user objects  : $($GroupMembers.Count)"
Write-Console "Valid UserMailboxes    : $($ValidMembers.Count)"
Write-Console "Skipped objects        : $($SkippedMembers.Count)"
Write-Console "New scope members      : $($ScopeToAdd.Count)"
Write-Console "Former target users    : $($FormerUsers.Count)"
Write-Console ""
Write-Console "Contact ADOPT          : $TotalAdopt"
Write-Console "Contact CREATE         : $TotalCreate"
Write-Console "Contact UPDATE         : $TotalUpdate"
Write-Console "Contact DELETE         : $TotalDelete"
Write-Console "Personal DUPDELETE     : $TotalDupDelete"
Write-Console "Contact unchanged      : $TotalUnchanged"
Write-Console ""
Write-Console "DryRun                 : $DryRun"
Write-Console "============================================"


# ============================================================
# 10. FINAL STATUS
# ============================================================

# Statuses that are not an error: OK, and the transient ones (new user /
# RBAC propagation) that the next run picks up. Everything else needs attention
# and fails the job, so that it is visible in Azure Automation.
$TransientStatuses = @(
    "OK",
    "SCOPE ADD REQUIRED",
    "RBAC PROPAGATION PENDING",
    "RBAC pending / 403"
)

$FailedMailboxes = @(
    $Results |
    Where-Object { [string]$_.Status -notin $TransientStatuses }
)

$FailedFormerUsers = @(
    $FormerUserResults |
    Where-Object { [string]$_.Status -like "ERROR*" -or [string]$_.Status -like "*SAFETY STOP*" }
)

Write-Console ""

if (
    $FailedMailboxes.Count -gt 0 -or
    $FailedFormerUsers.Count -gt 0 -or
    $ScopeAddErrors -gt 0
) {

    $Problems = @(
        $FailedMailboxes   | ForEach-Object { "$($_.User)=$($_.Status)" }
        $FailedFormerUsers | ForEach-Object { "$($_.User)=$($_.Status)" }
        if ($ScopeAddErrors -gt 0) { "SCOPE ADD errors=$ScopeAddErrors" }
    )

    Write-Log ("Sync completed WITH ERRORS: " + ($Problems -join "; ")) "ERROR"

    throw "Central phonebook sync: $($Problems.Count) problem(s). See log."
}

if ($DryRun) {

    Write-Log "DRY RUN completed. No changes were made." "SUCCESS"
}
else {

    Write-Log "Production synchronisation completed." "SUCCESS"
}
