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

## Getting Started

These instructions will get you a copy of the tool up and running on your local machine.

### Prerequisites

* PowerShell:
  * Windows PowerShell 5.1 or PowerShell Core (pwsh 7.x+) across Windows, Linux, and macOS
  * Microsoft Graph Module:
    * `Install-Module -Name Microsoft.Graph`
    * Scopes required: `AuditLog.Read.All`, `User.Read.All`, `UserAuthenticationMethod.Read.All`, `Policy.Read.All`, `Directory.Read.All`
* Optional Dependencies:
  * **Microsoft Excel** (COM automation on Windows, for generating Excel report)
  * **Impacket** (`impacket-ticketConverter` / `ticketConverter.py` on Linux, for automatic `.kirbi` to `.ccache` ticket conversion)
  * **AADInternals** or **MSAL.PS** (optional alternative providers for Seamless SSO / IWA token acquisition)

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
   * **Windows (Domain-Joined)**: Automatically queries the Kerberos ticket for `HTTP/autologon.microsoftazuread-sso.com` via SSPI / Negotiate, receives a SAML assertion, and exchanges it for a Graph token.
   * **Linux / Cross-Platform**: Uses a TGT (`.ccache`, `.kirbi`, or Base64 /nowrap string) along with `-Domain` and `-DC` (Domain Controller). Automatically configures GSSAPI / `KRB5_CONFIG` and utilizes `curl --negotiate` to acquire the service ticket and exchange it for a Graph access token.
3. **Application & Certificate Authentication (`MSGraph-Credentials.csv`)**: Client ID + Secret or Certificate-based authentication.
4. **Interactive / PSCredential (`-Credential`)**: Prompt for user credentials or supply a `[PSCredential]` object.

## Usage

### Examples

#### 1. Direct Access Token Authentication
```powershell
# Using an Access Token obtained from Azure CLI or Az PowerShell
$token = (Get-AzAccessToken -ResourceTypeName MSGraph).Token
.\AzureADRecon.ps1 -AccessToken $token

# Or pass a raw token string
.\AzureADRecon.ps1 -AccessToken "eyJ0eXAiOiJKV1QiLCJhbGciOi..."
```

#### 2. Azure AD Seamless SSO (Windows Domain-Joined Host)
```powershell
# Authenticates silently using the logged-on user's Kerberos ticket
.\AzureADRecon.ps1 -SeamlessSSO
```

#### 3. Seamless SSO from Linux with TGT (.ccache or .kirbi)
```bash
# Using a .ccache ticket file
pwsh ./AzureADRecon.ps1 -SeamlessSSO \
    -TGT /tmp/krb5cc_1000 \
    -Domain corp.example.com \
    -DC 10.0.0.1 \
    -TenantID 8f3c7...

# Using a .kirbi ticket file (automatically converted if impacket-ticketConverter is in PATH)
pwsh ./AzureADRecon.ps1 -SeamlessSSO \
    -TGT /path/to/ticket.kirbi \
    -Domain corp.example.com \
    -DC 10.0.0.1 \
    -TenantID 8f3c7...
```

#### 4. Seamless SSO with Base64 (nowrap) TGT String
```bash
# Directly passing a Base64-encoded (nowrap) ticket string from Rubeus or klist
pwsh ./AzureADRecon.ps1 -SeamlessSSO \
    -TGT "doIFvjCCBbqgAwIBBaEDAgEWoi..." \
    -Domain corp.example.com \
    -DC 10.0.0.1 \
    -TenantID 8f3c7...
```

#### 5. User Credentials (PSCredential)
```powershell
# Interactive prompt
.\AzureADRecon.ps1

# Or with PSCredential object
$username = "username@fqdn"
$passwd = ConvertTo-SecureString "PlainTextPassword" -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential ($username, $passwd)
.\AzureADRecon.ps1 -Credential $creds
```

#### 6. Configuration via `MSGraph-Credentials.csv`
Populate relevant parameters in `MSGraph-Credentials.csv`:
```csv
Category,Value,Comment
TenantID,your-tenant-id,Tenant ID
AccessToken,,Direct Bearer/Access Token
SeamlessSSO,true,Set to true to use Seamless SSO
Domain,corp.example.com,Active Directory domain name
DC,10.0.0.1,Domain Controller IP or hostname
TGT,/path/to/ticket.ccache,Ticket path or Base64 string
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
    Domain or Azure AD user credentials.

-AccessToken <String|SecureString>
    Microsoft Graph Access Token to authenticate directly without prompt or client credentials.

-SeamlessSSO
    Switch to authenticate via Azure AD Seamless Single Sign-On (Kerberos ticket for HTTP/autologon.microsoftazuread-sso.com).

-TGT <String>
    Path to Kerberos ticket (.ccache or .kirbi) or Base64-encoded (nowrap) ticket string for Seamless SSO.

-Domain <String>
    Active Directory domain name (FQDN) for Kerberos/SeamlessSSO authentication.

-DC <String>
    Domain Controller IP address or hostname for Kerberos authentication.

-TenantID <String>
    The Azure Tenant ID to connect to when you have multiple tenants or need realm routing.

-OutputDir <String>
    Path for AzureADRecon output folder to save the CSV/XML/JSON/HTML files and the AzureADRecon-Report.xlsx. (Default pwd)

-Collect <String>
    Which modules to run (Comma separated; e.g Tenant,Domain. Default all).
    Valid values include: Tenant, Domain, Licenses, Users, ServicePrincipals, DirectoryRoles, DirectoryRoleMembers, Groups, GroupMembers, Devices, ConditionalAccessPolicies.

-OutputType <String>
    Output Type; Comma separated; e.g CSV,STDOUT,Excel (Default STDOUT with -Collect parameter, else CSV and Excel).
    Valid values include: STDOUT, CSV, XML, JSON, HTML, Excel, All (excludes STDOUT).

-DormantTimeSpan <Int>
    Timespan for Dormant accounts in days. (Default 30 days)

-PassMaxAge <Int>
    Maximum machine account password age in days. (Default 90 days)

-PageSize <Int>
    The PageSize to set for Graph queries. (Default 200)

-Threads <Int>
    The number of threads to use during processing objects. (Default 10)

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
