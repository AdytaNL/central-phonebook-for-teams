# Central phonebook for Microsoft Teams

**Shared contacts for Microsoft Teams.** A PowerShell runbook that mirrors a central
SharePoint list (Microsoft List) of names and phone numbers into the **personal Outlook
contacts** of a group of users.
Because Teams shows a user's Outlook contacts under **People** and in **Calls**, every
Teams Phone user gets the same central phonebook, maintained in one place.

Background, design choices and pitfalls: [One phone book for everyone: under the hood](https://adyta.nl/en/engineering/one-phone-book-for-everyone/under-the-hood/).

## Requirements

- Microsoft 365 with Exchange Online mailboxes for the target users (Teams shows contacts from the default contacts folder of the mailbox).
- An Azure subscription for Azure Automation (PowerShell 7.4 runtime; modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, ExchangeOnlineManagement).
- Admin rights to grant Microsoft Graph application permissions, a Sites.Selected site grant and Exchange Application RBAC.
- A SharePoint site with a list of two text columns (name and phone number).

## Quick start

1. Create the list, the target group and the mail-enabled scope group ([docs/SETUP.md](docs/SETUP.md), step 1).
2. Create the Automation account and grant the permissions ([docs/SETUP.md](docs/SETUP.md), steps 2-8).
3. Import the runbook, run it as a **dry run** (the default), check the output, then schedule it with `DryRun = $false` ([docs/SETUP.md](docs/SETUP.md), steps 9-10).

## Why

Microsoft Teams has no central, shared contact list for external numbers
(suppliers, partners, on-call numbers, branch offices, and so on):

- The organisation's address book (GAL) only contains directory objects. Org contacts
  or a shared mailbox's contacts do not show up as contacts in Teams People or Calls.
- What Teams *does* show are each user's own Outlook contacts.

So this script takes a SharePoint list as the single source and keeps a copy of it in
every target user's mailbox. Editors maintain the list; the script does the rest,
every hour.

## Architecture

Two flows meet in the runbook: **what** is synced (the data) and **to whom** (the users).

```
WHAT                                         TO WHOM
SharePoint list (source, editors)            Entra ID target group (e.g. Teams Phone users)
        |                                            |  only RecipientTypeDetails = UserMailbox
        |                                            v
        |                                    Mail-enabled scope group (Exchange RBAC scope,
        |                                    kept in sync by the script)
        |                                            |
        +-------------> Azure Automation runbook <---+
                        (hourly, managed identity)
                                  |
                                  v
              Personal contacts in each target mailbox
              (only contacts carrying our extension are managed)
                                  |
                                  v
                   Outlook / Teams People / Teams Calls
```

Per run the script:

1. Reads the members of the target Entra ID group and keeps only real **UserMailboxes**.
   Teams resource accounts and other objects without a user mailbox drop out.
2. Reads the SharePoint list (a name column and a phone number column).
3. Mirrors the Exchange scope group to the target group (adds new users first).
4. Synchronises the contacts for every target user (see below).
5. Cleans up the contacts of users who left the target group, and only after that
   removes them from the scope group.

## How managed contacts are recognised

Every contact the script creates gets an Outlook **open extension**:

| Field | Value |
|---|---|
| extensionName | the `-ExtensionName` parameter, e.g. `com.contoso.centralphonebook` |
| id (as Graph returns it) | `Microsoft.OutlookServices.OpenTypeExtension.<extensionName>` |
| sourceId | ID of the SharePoint list item |
| sourceSystem | the `-ExtensionSourceSystem` parameter |

The link between SharePoint and the mailbox is **only the sourceId**, never the name.
Names and numbers can therefore change safely. Contacts without this extension are the
user's own contacts and are not managed (except for the merge case below).

Contacts are read with
`$expand=extensions($filter=id eq 'Microsoft.OutlookServices.OpenTypeExtension.<extensionName>')`.
Graph does not allow `$expand=extensions` on Outlook contacts without such a filter
(`ErrorGraphExtensionExpandRequiresFilter`).

Pick your own reverse-domain extension name and never change it after go-live: contacts
with the old name would no longer be recognised.

## Actions per mailbox

| Action | When |
|---|---|
| **CREATE** | The list item has no managed contact yet and no personal copy. |
| **UPDATE** | A managed contact exists, but the name (case-sensitive) or number differs. |
| **DELETE** | A managed contact whose sourceId is no longer in the list. |
| **ADOPT** | The user already has a contact with the same name + number (case ignored) and there is no managed contact yet. That contact gets our extension and our spelling. A copy with different capitalisation is preferred, because the user typed it. The user's extra fields (email, notes) are kept; the business phone number follows the list from then on. |
| **DUPDELETE** | After an ADOPT: remaining copies with **exactly** the list spelling are deleted. Copies with a different spelling are always kept. |

Execution order: ADOPT, CREATE, UPDATE, DELETE, DUPDELETE. On an error the mailbox stops
immediately, so duplicates are never deleted after a failed adopt.

A list item with an empty name or number is skipped, but its existing contact is **not**
deleted.

Note: an adopted contact is a managed contact from then on. If the item disappears from
the list, the contact is deleted at that user, including the user's own extra fields.

Merging is controlled by `AdoptPersonalDuplicates` (ADOPT) and `DeleteExtraPersonalDuplicates`
(DUPDELETE), both `$true` by default. Set them to `$false` to leave personal copies alone.

## Safety stops (fail closed)

| Parameter | Default | Effect |
|---|---|---|
| `DryRun` | `$true` | Read and report only. |
| `MinimumTargetUsers` | 2 | Fewer valid target users: the whole run stops. |
| `MinimumSourceContacts` | 1 | Fewer valid list items: the whole run stops. **Set this to a value somewhat below the normal size of your list** (for example 80 for a list of about 100), so that an empty or truncated source never leads to mass deletes. |
| `MaxDeletesPerMailbox` | 25 | More DELETEs in one mailbox: that mailbox is skipped. |
| `MaxPersonalDuplicateDeletesPerMailbox` | 25 | More DUPDELETEs in one mailbox: that mailbox is skipped. |
| `MaxFormerUsersPerRun` | 3 | More users leaving the target group in one run: the whole run stops. |
| `MaxFormerUserCleanupDeletes` | 250 | More managed contacts at a former user: cleanup stops, the user stays in scope. |
| `MaxLookalikeRatio` | 0.5 | If 50% or more of the source items already exist in a mailbox as contacts *without* our extension, extension detection is probably failing. The mailbox is skipped (otherwise merge would treat our own contacts as duplicates). |

Also note:

- An Exchange error during the mailbox check (e.g. throttling) aborts the whole run, so a
  temporary error never turns a valid user into a former user.
- Target group and scope group are compared on Entra ID object ID
  (`ExternalDirectoryObjectId`), not on UPN or email address.
- Duplicate sourceIds in a mailbox: that mailbox is skipped.
- A former user is removed from the scope group only after all managed contacts are gone
  and a re-read confirms nothing is left. Otherwise the RBAC access would disappear before
  cleanup and cleanup would no longer be possible.
- If a mailbox or former user ends with an error status, or a scope change fails, the
  script ends with an exception. The Automation job is then **Failed**, which you can alert
  on. Transient statuses (RBAC propagation of a new user) do not count as an error.

## Permission model (least privilege)

The script runs as the **system-assigned managed identity** of a dedicated Azure
Automation account. There is no tenant-wide `Contacts.ReadWrite` and no Exchange
administrator role.

| Where | Permission | Purpose |
|---|---|---|
| Microsoft Graph (application) | `Sites.Selected` + `read` on one site | Read the phonebook list, nothing else in SharePoint |
| Microsoft Graph (application) | `GroupMember.ReadBasic.All` | Read members of the target group |
| Microsoft Graph (application) | `User.ReadBasic.All` | Name / UPN of the members |
| Office 365 Exchange Online (application) | `Exchange.ManageAsApp` | Connect to Exchange Online unattended (grants no Exchange rights by itself) |
| Exchange Application RBAC | `Application Contacts.ReadWrite` with a management scope `MemberOfGroup` = scope group | Read/write contacts, only in mailboxes of scope group members |
| Exchange role group | `View-Only Recipients` | `Get-EXOMailbox`, `Get-DistributionGroupMember` |
| Exchange role group | Custom role derived from *Security Group Creation and Membership* with only `Add-/Remove-DistributionGroupMember`, with a write scope that matches only the scope group | Manage the members of the scope group and nothing else (with `-BypassSecurityGroupManagerCheck`) |

Why a separate scope group? Exchange Application RBAC scopes mailboxes with a management
scope. `MemberOfGroup` only counts **direct** members of a mail-enabled group, so the script
keeps a flat technical group in sync with your (possibly dynamic) target group. Only **direct**
user members of the target group are read, so do not rely on nested groups there.

Use a dedicated Automation account: the managed identity's permissions apply to every
runbook in that account.

## Installation

See [docs/SETUP.md](docs/SETUP.md) for the step-by-step setup with PowerShell examples:
Automation account, modules, Graph permissions, the SharePoint site grant, Exchange
Application RBAC, the custom role, the runbook, the schedule and an alert.

### SharePoint list

Create a Microsoft List with two text columns, for example **Name** and **Phone**
(the column display names are parameters). Store phone numbers as text so a leading zero
or `+` is kept. Give only the phonebook editors edit rights on the list.

## Parameters

| Parameter | Required | Example / default |
|---|---|---|
| `Organization` | yes | `contoso.onmicrosoft.com` |
| `TeamsPhoneGroupId` | yes | `00000000-0000-0000-0000-000000000000` |
| `ScopeGroup` | yes | `Phonebook Sync Users` |
| `SharePointHost` | yes | `contoso.sharepoint.com` |
| `SharePointPath` | yes | `/sites/Intranet` |
| `ListName` | yes | `Central Phonebook` |
| `NameColumnName` | no | `Name` |
| `PhoneColumnName` | no | `Phone` |
| `ExtensionName` | no | `com.contoso.centralphonebook` (use your own) |
| `ExtensionSourceSystem` | no | `Central Phonebook` |
| `DryRun` | no | `$true` |
| `ManagedIdentity` | no | `$false` (use `$true` in Azure Automation) |
| `FunctionsOnly` | no | switch |
| Safety limits | no | see the table above |

## Testing locally

Connect yourself first: Microsoft Graph **app-only** (delegated sessions are refused) with
an identity that has the same Graph permissions, and `Connect-ExchangeOnline`. The
identity that reads and writes contacts also needs an Exchange Application RBAC assignment
(`Application Contacts.ReadWrite` within the same management scope).

In practice that identity is an app registration with a certificate, set up like the
managed identity in [docs/SETUP.md](docs/SETUP.md):

- the same app roles (step 4) and its own read grant on the site (step 5);
- an Exchange service principal for it (step 6, with the app's application ID and the
  object ID of its Enterprise application);
- the same role assignment in the same management scope (step 7), and membership of the
  role group from step 8 if you also run the scope group changes locally.

Then connect with `Connect-MgGraph -ClientId <app-id> -TenantId <tenant> -CertificateThumbprint <thumbprint>`
and `Connect-ExchangeOnline -AppId <app-id> -Organization <tenant> -CertificateThumbprint <thumbprint>`.

```powershell
$p = @{
    Organization      = "contoso.onmicrosoft.com"
    TeamsPhoneGroupId = "00000000-0000-0000-0000-000000000000"
    ScopeGroup        = "Phonebook Sync Users"
    SharePointHost    = "contoso.sharepoint.com"
    SharePointPath    = "/sites/Intranet"
    ListName          = "Central Phonebook"
}

# Dry run (default): reads and reports, changes nothing
.\Sync-CentralPhonebook.ps1 @p

# Live
.\Sync-CentralPhonebook.ps1 @p -DryRun $false

# Load only the functions, e.g. to inspect one mailbox
. .\Sync-CentralPhonebook.ps1 @p -FunctionsOnly
@(Get-ManagedContacts -UserPrincipalName "<user UPN>").Count
```

The mandatory parameters are required with `-FunctionsOnly` too. Pass booleans as
`$true`/`$false` (e.g. `-DryRun $false`); the string `false` gives a binding error.

Dot-sourcing also sets `Set-StrictMode -Version Latest` and `$ErrorActionPreference = 'Stop'`
in your own session.

Always start with a dry run and read the planned CREATE / UPDATE / DELETE / ADOPT /
DUPDELETE lines before running live.

## Azure Automation

- Runtime: PowerShell 7.4 runtime environment with `Microsoft.Graph.Authentication`,
  `Microsoft.Graph.Groups` and `ExchangeOnlineManagement`.
- Runbook: the content of `Sync-CentralPhonebook.ps1`, with **verbose logging enabled**.
- Schedule: for example every hour, with `ManagedIdentity = true`, `DryRun = false` and the
  environment parameters.
- Alert: a metric alert on `TotalJob` for this runbook with status Failed, Suspended or
  Stopped.

Azure Automation does not capture `Write-Host`. With `-ManagedIdentity $true` the script
therefore writes its log lines to the Verbose stream (and warnings/errors also to the
Warning stream). Module noise from Graph/EXO is suppressed. In the job: **Output** shows the
result table per mailbox, **All logs** shows the log lines.

### Statuses per mailbox

| Status | Meaning |
|---|---|
| `OK` | Synchronised |
| `SCOPE ADD REQUIRED` | Dry run: the user would be added to the scope group |
| `RBAC PROPAGATION PENDING` / `RBAC pending / 403` | Just added; Exchange RBAC not active yet. The next run picks the user up. |
| `ERROR reading contacts` | Graph error while reading (the Graph error code is in the log) |
| `ERROR duplicate SourceId` | Two managed contacts with the same sourceId; remove one by hand |
| `SAFETY STOP lookalikes (n)` | Many contacts without our extension that match the source; check extension detection |
| `DELETE SAFETY STOP` | Too many deletes; check the list |
| `DUPLICATE DELETE SAFETY STOP` | Too many personal duplicates; check the mailbox |
| `ERROR during sync` | Error halfway; see the log |

Former users: `WOULD CLEANUP + REMOVE SCOPE` (dry run), `CLEANED + REMOVED SCOPE`,
`CLEANUP SAFETY STOP`, `ERROR CLEANUP` (stays in scope, retried next run).

## Limitations and caveats

- **Keep resource accounts and service accounts out of the target group where you can.**
  The script skips everything that is not a `UserMailbox`, but a cleaner target group is
  better. If you use a dynamic group, add for example `(user.accountEnabled -eq true)` to
  the rule so disabled accounts drop out.
- **A managed identity cannot be the owner (ManagedBy) of a group.** That is why the
  script uses `-BypassSecurityGroupManagerCheck`, allowed only through the narrow custom
  role whose write scope matches only the scope group.
- **Do not manage the scope group by hand** and do not nest groups in it: Application RBAC
  only counts direct members.
- **Leavers:** remove a user from the target group *before* removing their licence or
  mailbox. If the mailbox can no longer be reached, cleanup ends as `ERROR CLEANUP`, the user
  stays in the scope group and the job fails until it is sorted out.
- **New users** are synced one run later: Exchange RBAC needs time to propagate after they
  are added to the scope group.
- **Write-Host is not logged in Azure Automation**; enable verbose logging on the runbook.
- Deleting a list item removes the contact for **all** users. Restore an item from the
  site's recycle bin: it keeps its ID and the contact comes back on the next run.
- Do not rename the list columns without updating the parameters: a missing column stops
  the sync.
- Only the default contacts folder is read and written; only the first business phone
  number is used.

## Uninstall / roll back

The safety stops (`MinimumTargetUsers`, `MaxFormerUsersPerRun`) also block a mass cleanup,
on purpose. To remove the phonebook: empty the target group in steps, or temporarily lower
`MinimumTargetUsers` and raise `MaxFormerUsersPerRun` (and `MaxFormerUserCleanupDeletes` if
needed), and let a run clean up the contacts (dry run first). Then remove the schedule, the
runbook, the role assignments, the role group, the custom role, the management scopes, the
Exchange service principal, the scope group, the `Sites.Selected` grant, the alert and the
Automation account. The commands are in
[docs/SETUP.md](docs/SETUP.md#uninstall--roll-back).

## License

MIT, see [LICENSE](LICENSE).

## Author

Built by Adyta (Lambert van der Plas) – https://adyta.nl
