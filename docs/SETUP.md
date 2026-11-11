# Setup

> **The commands below are a reference, not a turnkey installer.** Walk through them in a
> test tenant first, read every command before you run it, and replace every
> `<placeholder>` and every `contoso` value with your own. Run steps one at a time and check
> the result before continuing.

Overview:

1. Prepare the target group, the scope group and the SharePoint list
2. Create the Automation account with a system-assigned managed identity
3. Runtime environment and modules
4. Microsoft Graph app roles for the managed identity
5. `Sites.Selected`: read on one site
6. Exchange: service principal for the managed identity
7. Exchange: management scope and `Application Contacts.ReadWrite`
8. Exchange: narrow custom role for the scope group members
9. Runbook: import, publish, test, schedule
10. Alert on failed jobs

To remove everything again, see [Uninstall / roll back](#uninstall--roll-back).

Placeholders used:

| Placeholder | Meaning |
|---|---|
| `<tenant>` | e.g. `contoso.onmicrosoft.com` |
| `<subscription-id>` | Azure subscription ID |
| `<resource-group>` | e.g. `rg-phonebook` |
| `<automation-account>` | e.g. `aa-central-phonebook` |
| `<location>` | e.g. `westeurope` |
| `<mi-object-id>` | Object ID of the managed identity's service principal (Enterprise application) |
| `<mi-app-id>` | Application (client) ID of the managed identity |
| `<target-group-id>` | Object ID of the Entra ID target group |
| `<scope-group-smtp>` | Primary SMTP address of the scope group |

---

## 1. Groups and list

**Target group** (Entra ID): the users who should get the phonebook, typically everyone
with Teams Phone. Can be assigned or dynamic. Only **direct** user members are read, so do not use nested groups. For a dynamic group, add a condition that
keeps disabled accounts out, and make sure Teams resource accounts do not match, for example:

```
(<your membership condition>) -and (user.accountEnabled -eq true)
```

**Scope group** (Exchange Online): a mail-enabled security group that the script manages.
No nested groups.

```powershell
Connect-ExchangeOnline -Organization <tenant>

New-DistributionGroup -Name "Phonebook Sync Users" -Type Security
Get-DistributionGroup "Phonebook Sync Users" |
    Format-List DisplayName, PrimarySmtpAddress, DistinguishedName, ExternalDirectoryObjectId
```

Note the `PrimarySmtpAddress` (`<scope-group-smtp>`) and the `DistinguishedName`.

Tip: the scope group is a technical group, so hide it from the address list:

```powershell
Set-DistributionGroup "Phonebook Sync Users" -HiddenFromAddressListsEnabled $true
```

**SharePoint list**: a Microsoft List (e.g. *Central Phonebook* on `/sites/Intranet`) with
two single-line text columns, e.g. **Name** and **Phone**. Give the phonebook editors
edit rights on the list only.

---

## 2. Automation account with managed identity

Use a dedicated Automation account: a managed identity's permissions apply to all runbooks
in that account.

```powershell
Connect-AzAccount -Tenant <tenant>
Set-AzContext -Subscription <subscription-id>

New-AzResourceGroup -Name <resource-group> -Location <location>

New-AzAutomationAccount `
    -ResourceGroupName <resource-group> `
    -Name <automation-account> `
    -Location <location> `
    -Plan Basic `
    -AssignSystemIdentity

# Object ID of the managed identity (service principal)
(Get-AzAutomationAccount -ResourceGroupName <resource-group> -Name <automation-account>).Identity.PrincipalId
```

Find the application (client) ID in Entra ID > Enterprise applications (filter on
*Managed Identities*), or:

```powershell
Connect-MgGraph -Scopes "Application.Read.All"
Get-MgServicePrincipal -ServicePrincipalId <mi-object-id> | Format-List DisplayName, AppId, Id
```

---

## 3. Runtime environment and modules

Create a PowerShell **7.4** runtime environment and add the modules. Tested with
`Microsoft.Graph.Authentication` 2.40.0, `Microsoft.Graph.Groups` 2.40.0 and
`ExchangeOnlineManagement` 3.10.1. You can do this in the portal (Automation account >
Runtime environments) or via the REST API:

```powershell
$aa  = "/subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.Automation/automationAccounts/<automation-account>"
$api = "api-version=2024-10-23"

Invoke-AzRestMethod -Method PUT -Path "$aa/runtimeEnvironments/re-phonebook?$api" -Payload (@{
    location   = "<location>"
    properties = @{ runtime = @{ language = "PowerShell"; version = "7.4" }; defaultPackages = @{} }
} | ConvertTo-Json -Depth 5)

# Add the modules one at a time; Microsoft.Graph.Groups depends on
# Microsoft.Graph.Authentication, so wait until each one shows Succeeded.
$packages = [ordered]@{
    "Microsoft.Graph.Authentication" = "2.40.0"
    "Microsoft.Graph.Groups"         = "2.40.0"
    "ExchangeOnlineManagement"       = "3.10.1"
}

foreach ($name in $packages.Keys) {
    Invoke-AzRestMethod -Method PUT -Path "$aa/runtimeEnvironments/re-phonebook/packages/$name`?$api" -Payload (@{
        properties = @{ contentLink = @{ uri = "https://www.powershellgallery.com/api/v2/package/$name/$($packages[$name])" } }
    } | ConvertTo-Json -Depth 5)

    # Check the state before adding the next one:
    # (Invoke-AzRestMethod -Method GET -Path "$aa/runtimeEnvironments/re-phonebook/packages/$name`?$api").Content
}
```

---

## 4. Microsoft Graph app roles

Grant application permissions to the managed identity: on Microsoft Graph
`Sites.Selected`, `GroupMember.ReadBasic.All`, `User.ReadBasic.All`, and on Office 365
Exchange Online `Exchange.ManageAsApp`. Managed identities have no "API permissions"
blade, so this is done with PowerShell (requires an admin who can grant app roles).

```powershell
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All", "Application.Read.All"

$miObjectId = "<mi-object-id>"

$grants = @(
    # Well-known application IDs: Microsoft Graph and Office 365 Exchange Online
    @{ AppId = "00000003-0000-0000-c000-000000000000"; Role = "Sites.Selected" }
    @{ AppId = "00000003-0000-0000-c000-000000000000"; Role = "GroupMember.ReadBasic.All" }
    @{ AppId = "00000003-0000-0000-c000-000000000000"; Role = "User.ReadBasic.All" }
    @{ AppId = "00000002-0000-0ff1-ce00-000000000000"; Role = "Exchange.ManageAsApp" }
)

foreach ($g in $grants) {

    $resourceSp = @(Get-MgServicePrincipal -Filter "appId eq '$($g.AppId)'")

    if ($resourceSp.Count -ne 1) {
        throw "Expected exactly one service principal for appId $($g.AppId), found $($resourceSp.Count)."
    }

    $appRole = $resourceSp[0].AppRoles | Where-Object { $_.Value -eq $g.Role -and $_.AllowedMemberTypes -contains "Application" }

    if (-not $appRole) {
        throw "App role $($g.Role) not found on $($resourceSp[0].DisplayName)."
    }

    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $miObjectId `
        -PrincipalId $miObjectId `
        -ResourceId $resourceSp[0].Id `
        -AppRoleId $appRole.Id
}

# Verify
Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $miObjectId |
    Format-Table ResourceDisplayName, AppRoleId
```

`Exchange.ManageAsApp` only allows connecting to Exchange Online unattended. It does not
grant any Exchange permissions by itself; those follow in steps 7 and 8.

---

## 5. Sites.Selected: read on one site

`Sites.Selected` grants nothing until a site is explicitly granted. Grant **read** on the
site that holds the list. The admin session needs the delegated permission
`Sites.FullControl.All` for this call (this needs admin consent for Microsoft Graph PowerShell),
and the signed-in admin must be a SharePoint Administrator (or higher) and a member of the site.

```powershell
Connect-MgGraph -Scopes "Sites.FullControl.All"

$site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Intranet"

$body = @{
    roles               = @("read")
    grantedToIdentities = @(
        @{ application = @{ id = "<mi-app-id>"; displayName = "<automation-account>" } }
    )
} | ConvertTo-Json -Depth 5

Invoke-MgGraphRequest -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions" `
    -Body $body -ContentType "application/json"

# Verify
(Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions").value
```

Note: `grantedToIdentities` takes the **application (client) ID**, not the object ID.

**Revoke the consent afterwards.** `Sites.FullControl.All` is only needed for this one call.
If you consented for the whole organisation, revoke it in the Microsoft Entra admin center:
Enterprise apps > *Microsoft Graph Command Line Tools* > Permissions > Admin consent >
`...` next to the permission > Revoke permission. Consent for your own account only (the
User consent tab) cannot be revoked in the portal; remove the scope from the grant with
PowerShell instead:

```powershell
Connect-MgGraph -Scopes "DelegatedPermissionGrant.ReadWrite.All", "Application.Read.All"

# Microsoft Graph Command Line Tools (the client behind Connect-MgGraph)
$cli = Get-MgServicePrincipal -Filter "appId eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'"

Get-MgOauth2PermissionGrant -Filter "clientId eq '$($cli.Id)'" -All |
    Where-Object { ($_.Scope -split ' ') -contains 'Sites.FullControl.All' } |
    Format-List Id, ConsentType, PrincipalId, Scope

# For the grant you created in this step: keep the other scopes, drop Sites.FullControl.All
$g     = Get-MgOauth2PermissionGrant -OAuth2PermissionGrantId "<grant-id>"
$scope = (($g.Scope -split ' ') | Where-Object { $_ -and $_ -ne 'Sites.FullControl.All' }) -join ' '
Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $g.Id -BodyParameter @{ scope = $scope }
```

If `Sites.FullControl.All` was the only scope in that grant, remove the grant instead with
`Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $g.Id`.

---

## 6. Exchange: service principal

Register the managed identity in Exchange Online so that Exchange RBAC can be assigned to
it.

```powershell
Connect-ExchangeOnline -Organization <tenant>

New-ServicePrincipal `
    -AppId <mi-app-id> `
    -ObjectId <mi-object-id> `
    -DisplayName "<automation-account>"

$sp = Get-ServicePrincipal | Where-Object AppId -eq "<mi-app-id>"
$sp | Format-List DisplayName, AppId, ObjectId, Identity
```

`-ObjectId` is the object ID of the service principal (Enterprise application), not of an
app registration.

---

## 7. Exchange: management scope and Application Contacts.ReadWrite

The management scope limits contact access to the mailboxes of the scope group's members
(direct members only).

```powershell
$groupDn = (Get-DistributionGroup "Phonebook Sync Users").DistinguishedName

New-ManagementScope `
    -Name "Phonebook Sync Scope" `
    -RecipientRestrictionFilter "MemberOfGroup -eq '$groupDn'"

New-ManagementRoleAssignment `
    -Name "Central Phonebook Contacts RW" `
    -App $sp.Identity `
    -Role "Application Contacts.ReadWrite" `
    -CustomResourceScope "Phonebook Sync Scope"

# Verify
Get-ManagementRoleAssignment -RoleAssignee $sp.Identity |
    Format-List Name, Role, CustomResourceScope
```

You can test the effect with `Test-ServicePrincipalAuthorization -Identity $sp.Identity -Resource <user UPN>`
once the user is in the scope group (RBAC changes can take a while to propagate).

Do **not** also grant `Contacts.ReadWrite` in Entra ID: Entra ID grants and Exchange
Application RBAC add up, so that would give access to every mailbox regardless of the scope.

---

## 8. Exchange: narrow custom role for the scope group members

The script needs to add and remove members of the scope group, read recipients, and
nothing else. A managed identity cannot be the owner (ManagedBy) of a group, so it uses
`-BypassSecurityGroupManagerCheck`. That parameter is available in the built-in role
*Security Group Creation and Membership*, which is far too broad. So create a custom role
derived from it with only `Add-DistributionGroupMember` and `Remove-DistributionGroupMember`,
and limit it with a write scope that matches only the scope group.

```powershell
$roleName   = "Phonebook Scope Group Members"
$writeScope = "Phonebook Scope Group Only"
$roleGroup  = "Phonebook Sync Management"
$keep       = @("Add-DistributionGroupMember", "Remove-DistributionGroupMember")

# 8a. Write scope: must match exactly one recipient, the scope group
New-ManagementScope `
    -Name $writeScope `
    -RecipientRestrictionFilter "PrimarySmtpAddress -eq '<scope-group-smtp>'"

Get-Recipient -RecipientPreviewFilter (Get-ManagementScope $writeScope).RecipientFilter -ResultSize Unlimited |
    Format-Table DisplayName, PrimarySmtpAddress, RecipientTypeDetails
# Expect exactly one row: the scope group (MailUniversalSecurityGroup). Stop if not.

# 8b. Custom role with only Add/Remove-DistributionGroupMember
New-ManagementRole -Name $roleName -Parent "Security Group Creation and Membership"

Get-ManagementRoleEntry "$roleName\*" |
    Where-Object { $_.Name -notin $keep } |
    ForEach-Object { Remove-ManagementRoleEntry -Identity "$roleName\$($_.Name)" -Confirm:$false }

Get-ManagementRoleEntry "$roleName\*" | Format-Table Name, Parameters
# Expect only the two cmdlets; check that BypassSecurityGroupManagerCheck is in Parameters.

# 8c. Role group: custom role with the write scope + View-Only Recipients
New-RoleGroup `
    -Name $roleGroup `
    -Roles $roleName `
    -CustomRecipientWriteScope $writeScope `
    -Description "Central phonebook sync: manage members of the scope group only, read recipients."

New-ManagementRoleAssignment -SecurityGroup $roleGroup -Role "View-Only Recipients"

Add-RoleGroupMember -Identity $roleGroup -Member $sp.Identity

# Verify
Get-ManagementRoleAssignment -RoleAssignee $roleGroup |
    Format-Table Role, CustomRecipientWriteScope
```

`View-Only Recipients` covers `Get-EXOMailbox` and `Get-DistributionGroupMember`. The script
checks the Exchange connection with `Get-ConnectionInformation`, which needs no role.

---

## 9. Runbook: import, publish, test, schedule

Working in a new session? Re-run the `$aa` and `$api` lines from step 3 first.

Create the runbook on the runtime environment from step 3, upload the script, enable
verbose logging (Automation does not capture `Write-Host`) and publish:

```powershell
$rb = "$aa/runbooks/Sync-CentralPhonebook"

Invoke-AzRestMethod -Method PUT -Path "$rb`?$api" -Payload (@{
    location   = "<location>"
    properties = @{
        runbookType        = "PowerShell"
        runtimeEnvironment = "re-phonebook"
        logVerbose         = $true
        logProgress        = $false
        description        = "Central phonebook: SharePoint list to personal contacts. Default DryRun."
    }
} | ConvertTo-Json -Depth 5)

```

Upload the script as draft content (content type `text/powershell`) and publish. With the
Azure CLI, signed in with `az login` to the same subscription (`az rest` is used here because
it can send the script file unchanged as the request body):

```powershell
$rbUrl = "https://management.azure.com$rb"

az rest --method put  --url "$rbUrl/draft/content?$api" --headers "Content-Type=text/powershell" --body "@Sync-CentralPhonebook.ps1"
az rest --method post --url "$rbUrl/publish?$api"
```

Or in the portal: Runbook > Edit > paste the script > Save > Publish.

**Test run (dry run)** from the portal (Runbook > Start) or:

```powershell
$params = @{
    ManagedIdentity   = $true
    DryRun            = $true
    Organization      = "<tenant>"
    TeamsPhoneGroupId = "<target-group-id>"
    ScopeGroup        = "Phonebook Sync Users"
    SharePointHost    = "contoso.sharepoint.com"
    SharePointPath    = "/sites/Intranet"
    ListName          = "Central Phonebook"
    # MinimumSourceContacts = <somewhat below your list size>
}

Start-AzAutomationRunbook `
    -ResourceGroupName <resource-group> `
    -AutomationAccountName <automation-account> `
    -Name Sync-CentralPhonebook `
    -Parameters $params
```

In the first dry run every user shows `SCOPE ADD REQUIRED`. Check the planned actions in
the job's **All logs**. Then run live once (`DryRun = $false`): users are added to the scope
group and show `RBAC PROPAGATION PENDING`. After RBAC has propagated (typically 30 minutes to
two hours) the next run creates the contacts.

**Schedule** (e.g. hourly):

```powershell
New-AzAutomationSchedule `
    -ResourceGroupName <resource-group> `
    -AutomationAccountName <automation-account> `
    -Name "Hourly - Central Phonebook" `
    -StartTime (Get-Date).AddMinutes(10) `
    -HourInterval 1 `
    -TimeZone "Europe/Amsterdam"

$params.DryRun = $false

Register-AzAutomationScheduledRunbook `
    -ResourceGroupName <resource-group> `
    -AutomationAccountName <automation-account> `
    -RunbookName Sync-CentralPhonebook `
    -ScheduleName "Hourly - Central Phonebook" `
    -Parameters $params
```

Use your own time zone. After a script change, upload and publish again, then start a dry
run before the next scheduled run.

---

## 10. Alert on failed jobs

The script throws at the end when a mailbox or former user has an error status, so the job
becomes **Failed**. Alert on that with a metric alert on `TotalJob`:

```powershell
$ag = Invoke-AzRestMethod -Method PUT `
    -Path "/subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.Insights/actionGroups/ag-phonebook?api-version=2023-01-01" `
    -Payload (@{
        location   = "Global"
        properties = @{
            groupShortName = "Phonebook"
            enabled        = $true
            emailReceivers = @(@{ name = "admins"; emailAddress = "<alert-mailbox>"; useCommonAlertSchema = $true })
        }
    } | ConvertTo-Json -Depth 5)

$agId = ($ag.Content | ConvertFrom-Json).id

Invoke-AzRestMethod -Method PUT `
    -Path "/subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.Insights/metricAlerts/central-phonebook-sync-failed?api-version=2018-03-01" `
    -Payload (@{
        location   = "global"
        properties = @{
            description         = "Sync-CentralPhonebook job failed, stopped or suspended."
            severity            = 2
            enabled             = $true
            scopes              = @($aa)
            evaluationFrequency = "PT15M"
            windowSize          = "PT1H"
            autoMitigate        = $true
            targetResourceType  = "Microsoft.Automation/automationAccounts"
            criteria            = @{
                "odata.type" = "Microsoft.Azure.Monitor.SingleResourceMultipleMetricCriteria"
                allOf        = @(@{
                    criterionType   = "StaticThresholdCriterion"
                    name            = "FailedJobs"
                    metricNamespace = "Microsoft.Automation/automationAccounts"
                    metricName      = "TotalJob"
                    operator        = "GreaterThan"
                    threshold       = 0
                    timeAggregation = "Total"
                    dimensions      = @(
                        @{ name = "RunbookName"; operator = "Include"; values = @("Sync-CentralPhonebook") }
                        @{ name = "Status";  operator = "Include"; values = @("Failed", "Suspended", "Stopped") }
                    )
                })
            }
            actions = @(@{ actionGroupId = $agId })
        }
    } | ConvertTo-Json -Depth 10)
```

---

## Uninstall / roll back

The safety stops also block a mass cleanup: with fewer than `MinimumTargetUsers` (default 2)
valid target users the whole run stops, and with more than `MaxFormerUsersPerRun` (default 3)
users leaving in one run as well. Each former user also has a limit of
`MaxFormerUserCleanupDeletes` (default 250) managed contacts. So remove the contacts first,
then the configuration.

1. **Remove the contacts.** Either empty the target group in steps of at most three users
   and let a run clean up after each step (for the last one or two users you still need
   `MinimumTargetUsers = 0`), or temporarily set `MinimumTargetUsers = 0` and
   `MaxFormerUsersPerRun` to the number of users (and `MaxFormerUserCleanupDeletes` above
   your list size if needed), empty the target group and start one run. Start it as a dry
   run first and check that every user shows `WOULD CLEANUP + REMOVE SCOPE`; then run it
   live and check for `CLEANED + REMOVED SCOPE`. Adopted contacts are managed contacts, so
   they are removed too.
2. **Remove the schedule and the runbook**, so nothing runs against a half-removed setup:

   ```powershell
   Remove-AzAutomationSchedule -ResourceGroupName <resource-group> -AutomationAccountName <automation-account> -Name "Hourly - Central Phonebook" -Force
   Remove-AzAutomationRunbook  -ResourceGroupName <resource-group> -AutomationAccountName <automation-account> -Name Sync-CentralPhonebook -Force
   ```

3. **Remove the Exchange configuration.** Role assignments go before the scopes and the
   custom role; removing the role group also removes its role assignments.

   ```powershell
   Connect-ExchangeOnline -Organization <tenant>

   Remove-ManagementRoleAssignment "Central Phonebook Contacts RW" -Confirm:$false
   Remove-RoleGroup "Phonebook Sync Management" -Confirm:$false
   Remove-ManagementRole "Phonebook Scope Group Members" -Confirm:$false
   Remove-ManagementScope "Phonebook Sync Scope" -Confirm:$false
   Remove-ManagementScope "Phonebook Scope Group Only" -Confirm:$false
   Remove-ServicePrincipal -Identity "<mi-app-id>" -Confirm:$false
   Remove-DistributionGroup "Phonebook Sync Users" -Confirm:$false
   ```

4. **Remove the `Sites.Selected` grant** on the site
   (`DELETE /sites/{site-id}/permissions/{permission-id}`; find the permission ID with the
   verify call from step 5). This call does not support delegated permissions: run it from
   an app-only session with the application permission `Sites.FullControl.All`.
5. **Delete the Automation account.** This also deletes its system-assigned managed
   identity, and with it the Graph app role assignments from step 4.

   ```powershell
   Remove-AzAutomationAccount -ResourceGroupName <resource-group> -Name <automation-account> -Force
   ```

   Remove the alert rule and action group from step 10 as well, or the whole resource group
   if it holds nothing else.

---

## Checklist

- [ ] Target group contains only real users (no resource accounts, no disabled accounts)
- [ ] Scope group exists, is mail-enabled security, has no nested groups
- [ ] Managed identity has the four app roles and read on exactly one site
- [ ] No tenant-wide `Contacts.ReadWrite` anywhere
- [ ] `Application Contacts.ReadWrite` assigned with the `MemberOfGroup` scope
- [ ] Write scope matches exactly one recipient; custom role has exactly two cmdlets
- [ ] `MinimumSourceContacts` set to a value somewhat below your list size
- [ ] `ExtensionName` set to your own reverse-domain name
- [ ] Dry run checked before the first live run; alert tested
