# AzureADRecon: Azure Active Directory Recon [![Follow ADRecon on Twitter](https://img.shields.io/twitter/follow/ad_recon.svg?style=social&label=Follow%20%40ad_recon)](https://twitter.com/intent/user?screen_name=ad_recon "Follow ADRecon on Twitter")

AzureADRecon is a tool which extracts and combines various artefacts (as highlighted below) out of an Azure AD environment with a valid credential. The information can be presented in a specially formatted Microsoft Excel report that includes summary views with metrics to facilitate analysis and provide a holistic picture of the current state of the target environment.

The tool is useful to various classes of security professionals like auditors, DFIR, students, administrators, etc. It can also be an invaluable post-exploitation tool for a penetration tester.

The tool uses the Microsoft Graph PowerShell SDK (`Microsoft.Graph`) and supports modern authentication mechanisms including direct Access Tokens and Azure AD Seamless Single Sign-On (Kerberos ticket for `HTTP/autologon.microsoftazuread-sso.com`).

The following information is gathered by the tool:

* Tenant
* Domain
* Licenses
* Users
* ServicePrincipals
* DirectoryRoles
* DirectoryRoleMembers
* Groups
* GroupMembers
* Devices
* ConditionalAccessPolicies
* Applications and application/service-principal owners
* CredentialExpiry (metadata only)
* ConsentGrants and granted ApplicationPermissions
* PIMEligibility
* MFARegistration, SecurityDefaults, and AuthenticationMethodsPolicy

## Getting Started

These instructions will get you a copy of the tool up and running on your local machine.

### Prerequisites

* PowerShell:
  * Windows PowerShell 5.1 or PowerShell Core (pwsh 7.x+) across Windows, Linux, and macOS
  * Microsoft Graph Module:
    * `Install-Module -Name Microsoft.Graph`
    * Scopes required: `AuditLog.Read.All`, `User.Read.All`, `UserAuthenticationMethod.Read.All`, `Policy.Read.All`, `Directory.Read.All`; `RoleEligibilitySchedule.Read.Directory` is additionally requested when collecting PIM eligibility
* Optional Dependencies:
  * **Microsoft Excel** (COM automation on Windows, for generating Excel report)
  * **AADInternals** or **MSAL.PS** (optional alternative providers for Seamless SSO / IWA token acquisition)

### User collection permissions

`Authentication_RequestFromUnsupportedUserRole` means Graph denied the operation based on the signed-in user's role. Having an access token or consenting to scopes alone does not grant that role.

The user collector requests `signInActivity`. This property requires `AuditLog.Read.All` and Microsoft Entra ID P1/P2; delegated sign-in reporting also requires a supported Entra role, such as Reports Reader. See [list users](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0) and [sign-in report permissions](https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0). If access is denied, AzureADRecon retries without `signInActivity`. Sign-in timestamps, logon ages, never-logged-in flags, and dormancy are blank (unknown) when activity data is unavailable.

Reading another user's authentication methods requires `UserAuthenticationMethod.Read.All` and, for delegated access, a supported role such as Global Reader, Authentication Administrator, or Privileged Authentication Administrator. See [authentication method permissions](https://learn.microsoft.com/en-us/graph/api/authentication-list-methods?view=graph-rest-1.0). Denied method access leaves method details blank. Registration and capability come independently from the [registration report](https://learn.microsoft.com/en-us/graph/api/authenticationmethodsroot-list-userregistrationdetails?view=graph-rest-1.0), which requires `AuditLog.Read.All` and a supported delegated role. Missing report entries (including disabled users) remain unknown. `MFAStatus` is `Registered`, `NotRegistered`, or `Unknown`; it never claims MFA is enforced. `MFA Registered` and `MFA Capable` are separate nullable booleans.

Dormancy uses `lastSuccessfulSignInDateTime`, covering successful interactive or non-interactive access. Failed attempts cannot make a dormant account appear active. Attempt timestamps are preserved separately, including the legacy attempt-age columns. Missing successful history leaves dormancy and the legacy `Never Logged In` flag unknown. The legacy `Never Non-Interactive Logged In` flag remains unknown because attempt history cannot establish successful access. See [sign-in activity semantics](https://learn.microsoft.com/en-us/graph/api/resources/signinactivity?view=graph-rest-1.0).

Missing dates, account state, and license lists remain blank. Known `False` and zero values are retained.

After an administrator grants or activates the required role and consents to permissions, sign in again and acquire a fresh Graph token to collect the restricted fields. The script does not change tenant permissions or roles.

Offline user collection regression checks: `pwsh -NoProfile -File ./tests/UserCollection.Tests.ps1`.

### Installing

If you have git installed, you can start by cloning the [repository](https://github.com/adrecon/AzureADRecon/):

```
git clone https://github.com/adrecon/AzureADRecon.git
```

Otherwise, you can [download a zip archive of the latest release](https://github.com/adrecon/AzureADRecon/archive/master.zip). The intent is to always keep the master branch in a working state.

## Authentication Methods

AzureADRecon supports multiple authentication methods:

1. **Direct Access Token (`-AccessToken`)**: Authenticate directly using a pre-acquired Microsoft Graph Bearer token (string or SecureString). Compatible with tokens from Azure CLI (`az account get-access-token --resource-type ms-graph`), `Get-AzAccessToken`, AADInternals, PRT, or browser sessions.
2. **Azure AD Seamless SSO (`-SeamlessSSO`)**:
   * **Windows (Domain-Joined)**: Uses the current session via SSPI / Negotiate, or imports a supplied KRB-CRED (`.kirbi`) file / Base64 ticket into that session using native Windows APIs. Requests a service ticket for `HTTP/autologon.microsoftazuread-sso.com`, receives a SAML assertion, and exchanges it for a Graph token.
   * **Linux / Cross-Platform**: Uses an existing Kerberos ccache file or Base64-encoded ccache along with `-Domain` and `-DC` (Domain Controller). Automatically configures GSSAPI / `KRB5_CONFIG` and utilizes `curl --negotiate` to acquire the service ticket and exchange it for a Graph access token.
3. **Application & Certificate Authentication (`MSGraph-Credentials.csv`)**: Client ID + Secret or Certificate-based authentication.
4. **Interactive**: With no token or application credentials, use Graph browser sign-in. `-Credential` accepts an **application client ID and client secret**, not a user/password pair; user credentials are explicitly rejected. Report attribution comes from the actual Graph context.

## Usage

### Examples

#### 1. Direct Access Token Authentication
```powershell
# Using an Access Token obtained from Azure CLI or Az PowerShell
$token = Get-AzAccessToken -ResourceTypeName MSGraph
.\AzureADRecon.ps1 -AccessToken $token

# Or pass a raw token string
.\AzureADRecon.ps1 -AccessToken "eyJ0eXAiOiJKV1QiLCJhbGciOi..."
```

Token response objects with a `Token`, `AccessToken`, or `access_token` property are accepted. Pass the **full response** to preserve expiry metadata. Use a token issued for **Microsoft Graph**. Startup validation makes a real Graph organization request; this requires `Organization.Read.All` or `Directory.Read.All`, plus the permissions needed by the selected collectors.

#### Token validity and renewal

1. **Startup:** read issuer expiry metadata (`ExpiresOn`, `ExpiresOnUtc`, `expires_on`, or `expires_in`), replace an expired/nearly expired token when a provider is available, and verify acceptance with Graph. An opaque token is supported. JWT claims are not decoded or trusted for validation. A future/not-yet-valid token is rejected by Graph; check the machine's UTC clock when lifetime validation fails.
2. **Before reads:** renew at the configured window, default **300 seconds before expiry** (`-TokenRefreshWindowSeconds`, range 60–900). Graph checks acceptance of every replacement, and `/organization` plus `/me` must identify the same tenant and delegated user before collection continues. Without a provider, a known expiry stops collection at this window.
3. **Unexpected 401:** reacquire once through the provider and replay the entire read, discarding partial output from the failed attempt. A second authentication failure stops collection. SDK `-All` commands can perform multiple internal pages; if expiry occurs inside one, the complete command is replayed once. Renewal is checked at script request boundaries, without a background timer interrupting running requests.
4. **Failure and audit:** rejected replacements, changed identities, stale expiry, or an unavailable refresh source stop collection and preserve completed collector checkpoints. Renewal/failure events go into the manifest without token contents. Successful renewal does not mark a collector incomplete.

Unix `expires_on` values are converted from an explicitly UTC epoch, independently of the machine's local time zone. Earlier code used a local-time epoch, which could reject a fresh SSO token as expired on hosts west of UTC, or delay renewal on hosts east of UTC. A local metadata rejection now identifies `Operation=TokenMetadata`, `Reason=ExpiryWithinRenewalWindow`, and safe `ExpiresOnUtc`/`ClientUtc` timestamps; it has no HTTP status because Graph was not contacted for that lease. Graph verification failures identify their connection/organization/user check and a fixed reason label (such as `InvalidTokenLifetime`, `UnreadableToken`, or `WrongAudience`) without copying response bodies. Cleanup treats an absent/already-disconnected Graph session as complete.

**SDK-managed interactive/certificate/application authentication** retains SDK/MSAL automatic renewal; the script does not inspect its private token cache or reconnect on every request. **Seamless SSO** installs its own provider and reacquires through the Kerberos/SAML exchange when due, using the original domain/DC/ticket settings. Forced renewal bypasses cached AADInternals/MSAL/Az shortcuts. The Kerberos ticket, DC connectivity, tenant policy, and consent must still allow issuance; the script does not extend a TGT. If a legacy SSO source returns no expiry metadata, it warns and reacquires every renewal-window interval, with one 401 recovery attempt; renewal before an unknown expiry cannot be guaranteed.

**Supplied tokens** need a way to obtain replacements. Use `-AccessTokenProvider` with a scriptblock returning exactly one response containing a token and expiry metadata. It receives an object with `Reason` (`Startup`, `Proactive`, or `Unauthorized`), `ForceRefresh`, `ExpiresOnUtc`, `TenantId`, and `UserId`. Use `ForceRefresh` to bypass the source's cache when supported. A response must be valid beyond the renewal window; returning a cached near-expiry token fails instead of looping. Write status output to verbose/information streams. Provider exceptions are suppressed in favor of a credential-safe error. Provider callbacks and tokens are not saved to disk.

For an existing Az.Accounts delegated session, load/sign in to Az before starting Graph:

```powershell
Connect-AzAccount -Tenant 'your-tenant-id'
$provider = {
    param($request)
    Get-AzAccessToken -ResourceTypeName MSGraph -TenantId 'your-tenant-id' -ErrorAction Stop
}
.\AzureADRecon.ps1 -TenantID 'your-tenant-id' -AccessTokenProvider $provider -OutputType CSV
```

Custom providers can return `@{ Token = $secureToken; ExpiresOn = $expiryUtc }` or a full MSAL/OAuth response. Preserve an existing response's absolute expiry when reusing it: `expires_in` is relative to each acquisition, so do not repeatedly return a cached response with its original relative lifetime. For raw strings/SecureStrings, `-AccessTokenExpiresOn '2026-10-24T12:00:00Z'` supplies missing expiry metadata, but a provider is still needed for renewal. A bare token without expiry can be checked by Graph at startup, but scheduled renewal is unavailable. Raw app-only tokens cannot be bound to a user through `/me`; use SDK-managed certificate/application authentication for those runs. Delegated provider/SSO runs require `/me` access and use the verified user ID for resume identity when the SDK context omits an account.

Microsoft documents [issuer response metadata and opaque tokens](https://learn.microsoft.com/en-us/entra/identity-platform/access-tokens), [Graph authentication](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands), and [Get-AzAccessToken](https://learn.microsoft.com/en-us/powershell/module/az.accounts/get-azaccesstoken).

If you see `InvalidAuthenticationToken` / "The token could not be read", obtain a fresh Graph token and pass the token value directly. Do not join status messages or command output into the token string. Seamless SSO status messages are displayed separately from its returned token.

#### 2. Azure AD Seamless SSO (Windows Domain-Joined Host)
```powershell
# Authenticates silently using the logged-on user's Kerberos ticket
.\AzureADRecon.ps1 -SeamlessSSO
```

With a Base64-encoded KRB-CRED TGT on Windows:

```powershell
.\AzureADRecon.ps1 -SeamlessSSO -TGT $base64Tgt -TenantID "<tenant-id>"

# Or supply a binary KRB-CRED file
.\AzureADRecon.ps1 -SeamlessSSO -TGT "C:\Tickets\ticket.kirbi" -TenantID "<tenant-id>"
```

This path uses an embedded C# adapter for the [Windows LSA API](https://learn.microsoft.com/en-us/windows/win32/api/ntsecapi/nf-ntsecapi-lsacallauthenticationpackage); no ticket conversion utility or extra PowerShell module is required. Base64 input stays in memory. Explicit ticket requests skip module-based and cached Az logins and use the supplied ticket through the operating system.

Import modifies the **current Windows logon session's ticket cache**, which can be shared by other processes. Tickets remain cached after the script exits; the script does not purge existing tickets. Import is limited to the current session. Domain controller connectivity and tenant Seamless SSO configuration are still required. `-DC` configures the non-Windows Kerberos path; Windows uses its domain configuration. Graph consent, user roles, and Conditional Access still apply.

#### 3. Seamless SSO from Linux with a ccache
```bash
# Using a .ccache ticket file
pwsh ./AzureADRecon.ps1 -SeamlessSSO \
    -TGT /tmp/krb5cc_1000 \
    -Domain corp.example.com \
    -DC 10.0.0.1 \
    -TenantID 8f3c7...

```

KRB-CRED (`.kirbi`) import is Windows-only. On Linux/macOS, supply a ccache. The tool does not convert ticket formats.

#### 4. Seamless SSO with a Base64-encoded ccache
```bash
# Pass the Base64-encoded contents of an existing ccache file
pwsh ./AzureADRecon.ps1 -SeamlessSSO \
    -TGT "<Base64-encoded ccache contents>" \
    -Domain corp.example.com \
    -DC 10.0.0.1 \
    -TenantID 8f3c7...
```

#### 5. Interactive or application credential authentication
```powershell
# Interactive user authentication
.\AzureADRecon.ps1 -TenantID "<tenant-id>"

# Application client-secret authentication; enter the client secret at the prompt
$credential = Get-Credential -UserName "00000000-0000-0000-0000-000000000001"
.\AzureADRecon.ps1 -TenantID "<tenant-id>" -Credential $credential
```

Certificate settings in `MSGraph-Credentials.csv` use app-only authentication without delegated `-Scopes`. `CertificateName` maps to Graph's `-CertificateSubjectName`. `Certificate` refers to a certificate thumbprint in `Cert:\LocalMachine\My`; the private key must be accessible to the process.

#### 6. Configuration via `MSGraph-Credentials.csv`
Populate relevant parameters in `MSGraph-Credentials.csv`:
```csv
Category,Value,Comment
TenantID,your-tenant-id,Tenant ID
AccessToken,,Direct Bearer/Access Token
SeamlessSSO,true,Set to true to use Seamless SSO
Domain,corp.example.com,Active Directory domain name
DC,10.0.0.1,Domain Controller IP or hostname
TGT,/path/to/ticket.ccache,Ticket file or Base64 contents: KRB-CRED on Windows; ccache on Linux/macOS
ClientID,,Required for ClientSecret / Certificate
ClientSecret,,Required for ClientSecret authentication
```
And execute:
```powershell
.\AzureADRecon.ps1 -Method MSGraph
```

#### 7. Generate Excel Report from Previous Output
```powershell
.\AzureADRecon.ps1 -GenExcel C:\AzureADRecon-Report-<timestamp>
```

When you run AzureADRecon, an `AzureADRecon-Report-<timestamp>` folder will be created containing the raw CSV/JSON files and `AzureADRecon-Report.xlsx`.

---

### Parameters

```
-Method <String>
    Which method to use; MSGraph (default)

-Credential <PSCredential>
    Application client ID (GUID) as username and client secret as password. User/password credentials are not supported.

-AccessToken <String|SecureString|TokenResponse>
    Microsoft Graph Access Token to authenticate directly without prompt or client credentials.

-AccessTokenProvider <ScriptBlock>
    Acquire initial/replacement Graph tokens with expiry metadata. Receives a renewal request object.

-AccessTokenExpiresOn <DateTimeOffset|String|UnixSeconds>
    Explicit expiry for a supplied token without response metadata; does not provide renewal credentials.

-TokenRefreshWindowSeconds <Int>
    Renewal window before expiry, from 60 to 900 seconds. Default 300.

-SeamlessSSO
    Switch to authenticate via Azure AD Seamless Single Sign-On (Kerberos ticket for HTTP/autologon.microsoftazuread-sso.com).

-TGT <String>
    Ticket file path or Base64 contents. Windows supports KRB-CRED (`.kirbi`) import into the current logon session. Linux/macOS supports ccache.

-Domain <String>
    Active Directory domain name (FQDN) for Kerberos/SeamlessSSO authentication.

-DC <String>
    Domain Controller IP address or hostname for Kerberos authentication.

-TenantID <String>
    The Azure Tenant ID to connect to when you have multiple tenants or need realm routing.

-OutputDir <String>
    Output directory for reports, CollectionManifest.json, and checkpoints. Defaults to a new timestamped directory.

-Collect <String>
    Which modules to run (Comma separated; e.g Tenant,Domain. Default all).
    Valid values include: Tenant, Domain, Licenses, Users, ServicePrincipals, DirectoryRoles, DirectoryRoleMembers, Groups, GroupMembers, Devices, ConditionalAccessPolicies.

-OutputType <String>
    Output formats, e.g CSV,JSON,STDOUT,Excel. Default: CSV and Excel. Excel adds CSV and requires Windows desktop Excel.
    Valid values include: STDOUT, CSV, XML, JSON, HTML, Excel, All (excludes STDOUT).

-DormantTimeSpan <Int>
    Timespan for Dormant accounts in days. (Default 30 days)

-PassMaxAge <Int>
    User password-age threshold in days. (Default 90 days)

-PageSize <Int>
    Graph page size, 1–999 (default 200); signInActivity queries are capped at 500.

-Threads <Int>
    The number of threads to use during processing objects. (Default 10)

-ClientID <String>
    Optional application client ID for interactive or configured app authentication.

-Resume
    Resume an interrupted run. Requires -OutputDir with the prior manifest and checkpoints.

-Logo <String>
    Which Logo to use in the Excel report? Valid values: AzureADRecon (default), Payatu.

-Log <Switch>
    Create AzureADRecon transcript console log.
```

### Bugs, Issues and Feature Requests

Please report all bugs, issues and feature requests in the [issue tracker](https://github.com/adrecon/AzureADRecon/issues). Or let me (@prashant3535) know directly.

### Contributing

Pull requests are always welcome.

### Mad props

Thanks for the awesome work by @_wald0, @CptJesus, @harmj0y, @mattifestation, @PyroTek3, @darkoperator, @ITsecurityAU Team, @CTXIS Team and others.

### License

AzureADRecon is a tool which gathers information about the Azure Active Directory and generates a report which can provide a holistic picture of the current state of the target environment.

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Affero General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License along with this program. If not, see http://www.gnu.org/licenses/.

This program borrows and uses code from many sources. All attempts are made to credit the original author. If you find that your code is used without proper credit, please shoot an insult to @prashant3535, Thanks.

## Authentication regression checks

Run the offline checks without Microsoft Graph modules or tenant credentials:

```powershell
pwsh -NoProfile -File ./tests/Authentication.Tests.ps1
```

Offline Kerberos import and routing checks (native calls mocked):

```powershell
pwsh -NoProfile -File ./tests/KerberosTickets.Tests.ps1
```

Live Windows ticket import and tenant sign-in must be verified on a domain-connected Windows host; the offline tests do not establish that a particular ticket or tenant will accept authentication.

## Partial collection, manifests, and recovery

Each run writes `CollectionManifest.json` with its verified tenant, authenticated identity when available, Graph context scopes, collector status, row counts, separate data/export status, exports, and availability events. States include `Success`, `Empty`, `Partial`, `PermissionDenied`, `ExportFailed`, `Failed`, and `NotStarted`. Warnings and export failures are therefore visible after execution without requiring a transcript. Tokens, client secrets, and ticket bytes are not included in the manifest.

Completed collectors receive hashed CLIXML checkpoints in `.checkpoints`. To continue a failed run after fixing a connection or refreshing credentials:

```powershell
.\AzureADRecon.ps1 -TenantID "<tenant-id>" -OutputDir "C:\Reports\run1" -Collect Default -Resume
```

Resume verifies the script hash, actual tenant, context account/client ID, selected collectors, settings, and checkpoint file hashes. It skips complete data collection (including when only its export previously failed), retries denied/partial/failed data collection, and regenerates selected exports. A missing, corrupt, or unreadable checkpoint is rejected and its dataset is recollected; unverified data is never imported. Failed resume authentication or identity checks preserve the previous usable manifest. An unidentified raw-token context cannot resume safely; use a delegated provider/SSO session with a verified `/me` identity or a context that identifies the account or application. Changed code/settings or a different identity require a fresh output directory. This is **collector-level recovery**, not mid-page resume; a failed collector starts again. The original snapshot date remains the basis for age calculations. Collector timestamps expose when data was fetched; a resumed run is not an atomic tenant snapshot. Earlier manifests are retained in `.manifests`.

Checkpoints contain the same sensitive directory inventory as the reports; store and delete them according to your normal report-handling practices. Running without `-Resume` will not overwrite an existing run manifest. Graph reads retry transient failures with bounded waits alongside SDK retry behavior. Supplied-token providers perform proactive renewal and one bounded authentication recovery attempt; unavailable or rejected replacements stop collection. Long server retry delays leave the collector resumable. Authentication disconnect, working-directory restoration, transcript stop, and Excel COM cleanup run in `finally` paths.

Denied membership queries produce `CollectionStatus=Unknown` rows with known parent IDs and blank member fields. These are availability markers, not actual members. Limited-information members retain their IDs. Microsoft documents a [Graph v1.0 group-members limitation](https://learn.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0): service principals may be omitted by that endpoint. Pagination does not remove that limitation.

CSV uses native platform paths and UTF-8; complex field values are serialized as JSON inside CSV cells. JSON/XML exports preserve nested values, and export errors are recorded in the manifest. Excel is Windows-only; unsupported hosts still produce CSV and mark Excel unavailable. The workbook includes MFA registration/capability/unknown summaries, not claims of enforcement. Percentages use observed field values rather than treating unknowns as negative results.

### Error handling and run outcomes

Ordinary request, parsing, and collector failures do not abort independent work. Query boundaries treat nonterminating errors as failures and discard incomplete request output. Permission denials and failed requests remain unavailable rather than becoming a successful empty result. Failed collectors have no known row count; the run finishes as `Partial`, with diagnostic categories, known error codes, exception types, HTTP status/request ID when available, and script line numbers in the manifest. Raw exception messages, request headers, bearer tokens, and target objects are excluded from persisted diagnostics and CLI failure summaries.

Malformed parser/user records retain an ID when available and `CollectionStatus=Unknown`; healthy rows continue. A deleted user or failed authentication-method request leaves method details unknown. After five consecutive non-permission/non-404 method lookup failures following bounded retries, remaining method lookups are skipped for that Users collector. Basic user inventory and independent registration/capability data continue. `Authentication Methods Collection Status` distinguishes a known successful method query from unknown details. Permission denials still allow other users to be queried. Resume recollects incomplete Users data.

Each report format is attempted independently. CSV/XML/JSON/HTML are written to temporary files and moved into place after a successful write; a failed write preserves a previous report. CSV includes the union of row columns so an unknown marker row cannot hide later fields. A checkpoint write failure still allows exports, but its dataset must be recollected on resume. Attribution, HTML index, and Excel failures do not discard other reports. Cleanup failures are recorded separately and do not replace the primary failure. Optional transcript startup failure does not prevent collection.

Invalid setup, an unwritable initial manifest, invalid resume identity/settings, or unrecoverable authentication stop the run cleanly. Continuing those cases could produce unusable or mixed-identity results. An initial manifest must be writable; later manifest failures warn and allow exports to continue, but reliable resume/status may then be unavailable. Cancellation remains a stop request. Process termination, PowerShell syntax/parameter-binding errors before execution, and catastrophic host/resource failures cannot be made recoverable by the script.

CLI invocation prints concise warnings instead of an unhandled exception dump. Exit codes are **0** for success, **2** for partial results (including reporting/logging limitations), and **1** for a fatal run failure. Inspect `CollectionManifest.json` before treating a report as complete. Calling `Invoke-AzureADRecon` directly still throws fatal failures so PowerShell callers can catch them; ordinary collector failures continue. Dot-sourcing does not call `exit`; `$script:AADRExitCode` exposes its outcome. Filesystem failures can prevent the final status from being persisted, so also retain the console warning/exit code.

## Inventory and report interpretation

`-Collect Default` includes all collectors. Individual names include `Tenant`, `Domain`, `Licenses`, `Users`, `ServicePrincipals`, `DirectoryRoles`, `DirectoryRoleMembers`, `Groups`, `GroupMembers`, `Devices`, `ConditionalAccessPolicies`, `Applications`, `ApplicationOwners`, `ServicePrincipalOwners`, `CredentialExpiry`, `ConsentGrants`, `ApplicationPermissions`, `PIMEligibility`, `MFARegistration`, `SecurityDefaults`, and `AuthenticationMethodsPolicy`.

Application registrations include requested permissions. Service-principal exposed OAuth scopes are definitions, **not grants**; the legacy `OAuth2Permissions (Count)` now reflects known exposed scopes and stays blank if unavailable. `ConsentGrants` exports actual delegated consent grants. `ApplicationPermissions` exports granted app-role assignments, preserving resource/role IDs and resolving role names where available. `CredentialExpiry` exports key IDs, dates, and expiration state, never secret values or private keys. `Licenses` includes prepaid-unit categories and service-plan details. `PIMEligibility` includes principal/role/scope IDs and schedule dates; eligibility is not an active role assignment.

`MFA Enforcement` is explicitly `NotEvaluated`. Registration does not prove enforcement, and no single user-wide boolean can describe every resource/sign-in context. Conditional Access, security-defaults, and authentication-method policy snapshots provide evidence for a contextual assessment. The tool does not perform that policy simulation or modify any policy. Existing roles, consent, licenses, and Conditional Access still govern what can be read. Unsupported or denied operations are reflected in the manifest.

## Validation

Run every offline suite in isolated PowerShell processes:

```powershell
pwsh -NoProfile -File ./tests/Run-Tests.ps1
```

The GitHub Actions regression workflow runs PowerShell Core on Windows, Linux, and macOS, plus Windows PowerShell 5.1. Tests mock Graph and Windows LSA; they do not require tenant credentials. They cover authentication parameter binding, identity attribution, successful-sign-in dormancy, nullable MFA facts, inventory endpoints, retry behavior, manifests/checkpoint recovery, exporters, parser versioning, and native ticket cleanup/status handling.

For an explicitly enabled live check on a domain-connected Windows host with Microsoft Graph and desktop Excel installed:

```powershell
$env:AADR_LIVE_TESTS = '1'
.\tests\LiveWindows.Smoke.ps1 -TenantID "<tenant-id>" -OutputDirectory "C:\Reports\live-check" -VerifyExcel
# Add -SeamlessSSO to test Windows SSO. Optional AADR_LIVE_TGT supplies a Base64 KRB-CRED.
```

Live tenant access and desktop Excel rendering cannot be established by the offline tests and must be verified in the target environment. Imported KRB-CRED remains in the current Windows logon session cache; these tests do not purge it.

### Re-running in the same PowerShell session

The parser is compiled once per C# source version. Repeated runs reuse the existing type; changed source uses a new namespace and cannot silently reuse an old parser. The Windows Kerberos adapter also guards repeated initialization.
