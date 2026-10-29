<#

.SYNOPSIS

    AzureADRecon is a tool which gathers information about the Azure Active Directory and generates a report which can provide a holistic picture of the current state of the target environment.

.DESCRIPTION

    AzureADRecon is a tool which extracts and combines various artefacts (as highlighted below) out of an Azure AD environment with a valid credential. The information can be presented in a specially formatted Microsoft Excel report that includes summary views with metrics to facilitate analysis and provide a holistic picture of the current state of the target environment.
    The tool is useful to various classes of security professionals like auditors, DFIR, students, administrators, etc. It can also be an invaluable post-exploitation tool for a penetration tester.
    The tool requires AzureAD PowerShell Module to be installed.

    The following information is gathered by the tool:
    * Tenant
    * Domain
    * Users
    * Licenses
    * ServicePrincipals
    * DirectoryRoles
    * DirectoryRoleMembers
    * Groups
    * GroupMembers
    * Devices
    * ConditionalAccessPolicies

    Author     : Prashant Mahajan

.NOTES

    The following commands can be used to turn off ExecutionPolicy: (Requires Admin Privs)

    PS > $ExecPolicy = Get-ExecutionPolicy
    PS > Set-ExecutionPolicy bypass
    PS > .\AzureADRecon.ps1
    PS > Set-ExecutionPolicy $ExecPolicy

    OR

    Start the PowerShell as follows:
    powershell.exe -ep bypass

    OR

    Already have a PowerShell open ?
    PS > $Env:PSExecutionPolicyPreference = 'Bypass'

    OR

    powershell.exe -nologo -executionpolicy bypass -noprofile -file AzureADRecon.ps1

.PARAMETER Method
	Which method to use; MSGraph (default).

.PARAMETER Credential
	Domain Credentials.

.PARAMETER AccessToken
    Microsoft Graph Access Token (String or SecureString) to authenticate directly.

.PARAMETER AccessTokenProvider
    Callback accepting a renewal request and returning one Graph token response with expiry metadata.

.PARAMETER AccessTokenExpiresOn
    UTC expiry date or Unix seconds for a supplied token without issuer expiry metadata.

.PARAMETER TokenRefreshWindowSeconds
    Renew supplied tokens this many seconds before expiry (60-900; default 300).

.PARAMETER SeamlessSSO
    Switch to authenticate via Azure AD Seamless Single Sign-On (Kerberos ticket for HTTP/autologon.microsoftazuread-sso.com).

.PARAMETER TGT
    Kerberos ticket file or Base64 contents: KRB-CRED (.kirbi) on Windows; ccache on Linux/macOS.

.PARAMETER Domain
    Active Directory domain name (FQDN) for Kerberos/SeamlessSSO authentication.

.PARAMETER DC
    Domain Controller IP address or hostname for Kerberos authentication.

.PARAMETER GenExcel
	Path for AzureADRecon output folder containing the CSV files to generate the AzureADRecon-Report.xlsx. Use it to generate the AzureADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run AzureADRecon.

.PARAMETER TenantID
    The Azure TenantID to connect to when you have multiple tenants.

.PARAMETER OutputDir
	Path for AzureADRecon output folder to save the files and the AzureADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist)

.PARAMETER Collect
    Which modules to run; Comma separated; e.g Tenant,Domain (Default all)
    Valid values include: Tenant, Domain, Licenses, Users, ServicePrincipals, DirectoryRoles, DirectoryRoleMembers, Groups, GroupMembers, Devices, ConditionalAccessPolicies.

.PARAMETER OutputType
    Output Type; Comma seperated; e.g STDOUT,CSV,XML,JSON,HTML,Excel (Default STDOUT with -Collect parameter, else CSV and Excel).
    Valid values include: STDOUT, CSV, XML, JSON, HTML, Excel, All (excludes STDOUT).

.PARAMETER DormantTimeSpan
    Timespan for Dormant accounts. (Default 30 days)

.PARAMETER PassMaxAge
    Timespan for Password age. (Default 90 days)

.PARAMETER PageSize
    The PageSize to set for the MSGraph query. (Default 200)

.PARAMETER Threads
    The number of threads to use during processing objects. (Default 10)

.PARAMETER Log
    Create AzureADRecon Log using Start-Transcript

.PARAMETER Logo
    Which Logo to use in the excel file? (Default AzureADRecon)
    Values include AzureADRecon, Payatu.

.EXAMPLE

    .\AzureADRecon.ps1

    Or

    $username = "00000000-0000-0000-0000-000000000001" # Application client ID
    $passwd = ConvertTo-SecureString "ApplicationClientSecret" -AsPlainText -Force
    $creds = New-Object System.Management.Automation.PSCredential ($username, $passwd)
	.\AzureADRecon.ps1 -Credential $creds -TenantID "<tenant-id>"

.EXAMPLE

	.\AzureADRecon.ps1 -GenExcel C:\AzureADRecon-Report-<timestamp>
    [*] AzureADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Generating AzureADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\AzureADRecon-Report-<timestamp>\AzureADRecon-Report.xlsx

.LINK

    https://github.com/adrecon/AzureADRecon
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false, HelpMessage = "Which method to use; MSGraph (default).")]
    [ValidateSet('MSGraph')]
    [string] $Method = 'MSGraph',

    [Parameter(Mandatory = $false, HelpMessage = "Which Tenant ID to connect to, picks up the default tenant ID if nothing specified")]
    [string] $TenantID,

    [Parameter(Mandatory = $false)]
    [string] $ClientID,

    [Parameter(Mandatory = $false)]
    [switch] $Resume,

    [Parameter(Mandatory = $false, HelpMessage = "Application client ID (GUID) and client secret as PSCredential; user/password sign-in is not supported.")]
    [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $false, HelpMessage = "Microsoft Graph Access Token (String or SecureString).")]
    $AccessToken,

    [Parameter(Mandatory = $false)]
    [scriptblock] $AccessTokenProvider,

    [Parameter(Mandatory = $false)]
    $AccessTokenExpiresOn,

    [Parameter(Mandatory = $false)]
    [ValidateRange(60, 900)] [int] $TokenRefreshWindowSeconds = 300,

    [Parameter(Mandatory = $false, HelpMessage = "Authenticate using Azure AD Seamless SSO (Kerberos ticket for HTTP/autologon.microsoftazuread-sso.com).")]
    [switch] $SeamlessSSO,

    [Parameter(Mandatory = $false, HelpMessage = "Active Directory domain name (FQDN) for Kerberos/SeamlessSSO authentication.")]
    [string] $Domain,

    [Parameter(Mandatory = $false, HelpMessage = "Domain Controller IP address or hostname for Kerberos authentication.")]
    [string] $DC,

    [Parameter(Mandatory = $false, HelpMessage = "Kerberos ticket file or Base64 contents: KRB-CRED (.kirbi) on Windows; ccache on Linux/macOS.")]
    [string] $TGT,

    [Parameter(Mandatory = $false, HelpMessage = "Path for AzureADRecon output folder containing the CSV files to generate the AzureADRecon-Report.xlsx. Use it to generate the AzureADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run AzureADRecon.")]
    [string] $GenExcel,

    [Parameter(Mandatory = $false, HelpMessage = "Path for AzureADRecon output folder to save the CSV/XML/JSON/HTML files and the AzureADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist)")]
    [string] $OutputDir,

    [Parameter(Mandatory = $false, HelpMessage = "Which modules to run; Comma separated; e.g Tenant,Domain (Default all) Valid values include: Tenant, Domain, Licenses, Users, ServicePrincipals, DirectoryRoles, DirectoryRoleMembers, Groups, GroupMembers, Devices")]
    [ValidateSet('Tenant', 'Domain', 'Licenses', 'Users', 'ServicePrincipals', 'DirectoryRoles', 'DirectoryRoleMembers', 'Groups', 'GroupMembers', 'Devices', 'ConditionalAccessPolicies', 'Applications', 'ApplicationOwners', 'ServicePrincipalOwners', 'CredentialExpiry', 'ConsentGrants', 'ApplicationPermissions', 'PIMEligibility', 'MFARegistration', 'SecurityDefaults', 'AuthenticationMethodsPolicy', 'Default')]
    [array] $Collect = 'Default',

    [Parameter(Mandatory = $false, HelpMessage = "Output type; Comma seperated; e.g STDOUT,CSV,XML,JSON,HTML,Excel (Default STDOUT with -Collect parameter, else CSV and Excel)")]
    [ValidateSet('STDOUT', 'CSV', 'XML', 'JSON', 'EXCEL', 'HTML', 'All', 'Default')]
    [array] $OutputType = 'Default',

    [Parameter(Mandatory = $false, HelpMessage = "Timespan for Dormant accounts. Default 30 days")]
    [ValidateRange(1,1000)]
    [int] $DormantTimeSpan = 30,

    [Parameter(Mandatory = $false, HelpMessage = "Timespan for Password age. Default 90 days")]
    [ValidateRange(1,1000)]
    [int] $PassMaxAge = 90,

    [Parameter(Mandatory = $false, HelpMessage = "The PageSize to set for the MSGraph query. Default 200")]
    [ValidateRange(1,999)]
    [int] $PageSize = 200,

    [Parameter(Mandatory = $false, HelpMessage = "The number of threads to use during processing of objects. Default 10")]
    [ValidateRange(1,100)]
    [int] $Threads = 10,

    [Parameter(Mandatory = $false, HelpMessage = "Create AzureADRecon Log using Start-Transcript")]
    [switch] $Log,

    [Parameter(Mandatory = $false, HelpMessage = "Which Logo to use in the excel file? Default AzureADRecon")]
    [ValidateSet('AzureADRecon', 'Payatu')]
    [string] $Logo = "AzureADRecon"
)

$script:AADRScriptPath = $PSCommandPath

$MSGraphSource = @"
// Thanks Dennis Albuquerque for the C# multithreading code
using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading;

namespace AADRecon
{
    public static class AzureADClass
    {
        //Values taken from https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference
        private static Dictionary<string, string> AzureSkuIDDictionary = new Dictionary<string, string>()
        {
            {"0c266dff-15dd-4b49-8397-2bb16070ed52", "Audio Conferencing"},
            {"2b9c8e7c-319c-43a2-a2a0-48c5c6161de7", "Azure Active Directory Basic"},
            {"078d2b04-f1bd-4111-bbd4-b4b1b354cef4", "Azure Active Directory Premium P1"},
            {"84a661c4-e949-4bd2-a560-ed7766fcaf2b", "Azure Active Directory Premium P2"},
            {"c52ea49f-fe5d-4e95-93ba-1de91d380f89", "Azure Information Protection Plan 1"},
            {"ea126fc5-a19e-42e2-a731-da9d437bffcf", "Dynamics 365 Customer Engagement Plan Enterprise Edition"},
            {"749742bf-0d37-4158-a120-33567104deeb", "Dynamics 365 For Customer Service Enterprise Edition"},
            {"cc13a803-544e-4464-b4e4-6d6169a138fa", "Dynamics 365 For Financials Business Edition"},
            {"8edc2cf8-6438-4fa9-b6e3-aa1660c640cc", "Dynamics 365 For Sales And Customer Service Enterprise Edition"},
            {"1e1a282c-9c54-43a2-9310-98ef728faace", "Dynamics 365 For Sales Enterprise Edition"},
            {"8e7a3d30-d97d-43ab-837c-d7701cef83dc", "Dynamics 365 For Team Members Enterprise Edition"},
            {"ccba3cfe-71ef-423a-bd87-b6df3dce59a9", "Dynamics 365 Unf Ops Plan Ent Edition"},
            {"efccb6f7-5641-4e0e-bd10-b4976e1bf68e", "Enterprise Mobility + Security E3"},
            {"b05e124f-c7cc-45a0-a6aa-8cf78c946968", "Enterprise Mobility + Security E5"},
            {"4b9405b0-7788-4568-add1-99614e613b69", "Exchange Online (Plan 1)"},
            {"19ec0d23-8335-4cbd-94ac-6050e30712fa", "Exchange Online (Plan 2)"},
            {"ee02fd1b-340e-4a4b-b355-4a514e4c8943", "Exchange Online Archiving For Exchange Online"},
            {"90b5e015-709a-4b8b-b08e-3200f994494c", "Exchange Online Archiving For Exchange Server"},
            {"7fc0182e-d107-4556-8329-7caaa511197b", "Exchange Online Essentials"},
            {"e8f81a67-bd96-4074-b108-cf193eb9433b", "Exchange Online Essentials"},
            {"80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82", "Exchange Online Kiosk"},
            {"cb0a98a8-11bc-494c-83d9-c1b1ac65327e", "Exchange Online Pop"},
            {"061f9ace-7d42-4136-88ac-31dc755f143f", "Intune"},
            {"b17653a4-2443-4e8c-a550-18249dda78bb", "Microsoft 365 A1"},
            {"4b590615-0888-425a-a965-b3bf7789848d", "Microsoft 365 A3 For Faculty"},
            {"7cfd9a2b-e110-4c39-bf20-c6a3f36a3121", "Microsoft 365 A3 For Students"},
            {"e97c048c-37a4-45fb-ab50-922fbf07a370", "Microsoft 365 A5 For Faculty"},
            {"46c119d4-0379-4a9d-85e4-97c66d3f909e", "Microsoft 365 A5 For Students"},
            {"cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46", "Microsoft 365 Business"},
            {"05e9a617-0261-4cee-bb44-138d3ef5d965", "Microsoft 365 E3"},
            {"06ebc4ee-1bb5-47dd-8120-11324bc54e06", "Microsoft 365 E5"},
            {"d61d61cc-f992-433f-a577-5bd016037eeb", "Microsoft 365 E3_Usgov_Dod"},
            {"ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658", "Microsoft 365 E3_Usgov_Gcchigh"},
            {"184efa21-98c3-4e5d-95ab-d07053a96e67", "Microsoft 365 E5 Compliance"},
            {"26124093-3d78-432b-b5dc-48bf992543d5", "Microsoft 365 E5 Security"},
            {"44ac31e7-2999-4304-ad94-c948886741d4", "Microsoft 365 E5 Security For Ems E5"},
            {"66b55226-6b4f-492c-910c-a3b7a3c9d993", "Microsoft 365 F1"},
            {"111046dd-295b-4d6d-9724-d52ac90bd1f2", "Microsoft Defender Advanced Threat Protection"},
            {"906af65a-2970-46d5-9b58-4e9aa50f0657", "Microsoft Dynamics Crm Online Basic"},
            {"d17b27af-3f49-4822-99f9-56a661538792", "Microsoft Dynamics Crm Online"},
            {"ba9a34de-4489-469d-879c-0f0f145321cd", "Ms Imagine Academy"},
            {"a4585165-0533-458a-97e3-c400570268c4", "Office 365 A5 For Faculty"},
            {"ee656612-49fa-43e5-b67e-cb1fdf7699df", "Office 365 A5 For Students"},
            {"1b1b1f7a-8355-43b6-829f-336cfccb744c", "Office 365 Advanced Compliance"},
            {"4ef96642-f096-40de-a3e9-d83fb2f90211", "Office 365 Advanced Threat Protection (Plan 1)"},
            {"cdd28e44-67e3-425e-be4c-737fab2899d3", "Office 365 Business"},
            {"b214fe43-f5a3-4703-beeb-fa97188220fc", "Office 365 Business"},
            {"3b555118-da6a-4418-894f-7df1e2096870", "Office 365 Business Essentials"},
            {"dab7782a-93b1-4074-8bb1-0e61318bea0b", "Office 365 Business Essentials"},
            {"f245ecc8-75af-4f8e-b61f-27d8114de5f3", "Office 365 Business Premium"},
            {"ac5cef5d-921b-4f97-9ef3-c99076e5470f", "Office 365 Business Premium"},
            {"18181a46-0d4e-45cd-891e-60aabd171b4e", "Office 365 E1"},
            {"6634e0ce-1a9f-428c-a498-f84ec7b8aa2e", "Office 365 E2"},
            {"6fd2c87f-b296-42f0-b197-1e91e994b900", "Office 365 E3"},
            {"189a915c-fe4f-4ffa-bde4-85b9628d07a0", "Office 365 E3 Developer"},
            {"b107e5a3-3e60-4c0d-a184-a7e4395eb44c", "Office 365 E3_Usgov_Dod"},
            {"aea38a85-9bd5-4981-aa00-616b411205bf", "Office 365 E3_Usgov_Gcchigh"},
            {"1392051d-0cb9-4b7a-88d5-621fee5e8711", "Office 365 E4"},
            {"c7df2760-2c81-4ef7-b578-5b5392b571df", "Office 365 E5"},
            {"26d45bd9-adf1-46cd-a9e1-51e9a5524128", "Office 365 E5 Without Audio Conferencing"},
            {"4b585984-651b-448a-9e53-3b10f069cf7f", "Office 365 F1"},
            {"04a7fb0d-32e0-4241-b4f5-3f7618cd1162", "Office 365 Midsize Business"},
            {"c2273bd0-dff7-4215-9ef5-2c7bcfb06425", "Office 365 Proplus"},
            {"bd09678e-b83c-4d3f-aaba-3dad4abd128b", "Office 365 Small Business"},
            {"fc14ec4a-4169-49a4-a51e-2c852931814b", "Office 365 Small Business Premium"},
            {"e6778190-713e-4e4f-9119-8b8238de25df", "Onedrive For Business (Plan 1)"},
            {"ed01faf2-1d88-4947-ae91-45ca18703a96", "Onedrive For Business (Plan 2)"},
            {"b30411f5-fea1-4a59-9ad9-3db7c7ead579", "Power Apps Per User Plan"},
            {"45bc2c81-6072-436a-9b0b-3b12eefbc402", "Power Bi For Office 365 Add-On"},
            {"f8a1db68-be16-40ed-86d5-cb42ce701560", "Power Bi Pro"},
            {"a10d5e58-74da-4312-95c8-76be4e5b75a0", "Project For Office 365"},
            {"776df282-9fc0-4862-99e2-70e561b9909e", "Project Online Essentials"},
            {"09015f9f-377f-4538-bbb5-f75ceb09358a", "Project Online Premium"},
            {"2db84718-652c-47a7-860c-f10d8abbdae3", "Project Online Premium Without Project Client"},
            {"53818b1b-4a27-454b-8896-0dba576410e6", "Project Online Professional"},
            {"f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c", "Project Online With Project For Office 365"},
            {"1fc08a02-8b3d-43b9-831e-f76859e04e1a", "Sharepoint Online (Plan 1)"},
            {"a9732ec9-17d9-494c-a51c-d6b45b384dcb", "Sharepoint Online (Plan 2)"},
            {"e43b5b99-8dfb-405f-9987-dc307f34bcbd", "Skype For Business Cloud Pbx"},
            {"b8b749f8-a4ef-4887-9539-c95b1eaa5db7", "Skype For Business Online (Plan 1)"},
            {"d42c793f-6c78-4f43-92ca-e8f6a02b035f", "Skype For Business Online (Plan 2)"},
            {"d3b4fe1f-9992-4930-8acb-ca6ec609365e", "Skype For Business Pstn Domestic And International Calling"},
            {"0dab259f-bf13-4952-b7f8-7db8f131b28d", "Skype For Business Pstn Domestic Calling"},
            {"54a152dc-90de-4996-93d2-bc47e670fc06", "Skype For Business Pstn Domestic Calling (120 Minutes)"},
            {"4b244418-9658-4451-a2b8-b5e2b364e9bd", "Visio Online Plan 1"},
            {"c5928f49-12ba-48f7-ada3-0d743a3601d5", "Visio Online Plan 2"},
            {"cb10e6cd-9da4-4992-867b-67546b1db821", "Windows 10 Enterprise E3"},
            {"488ba24a-39a9-4473-8ee5-19291e71b002", "Windows 10 Enterprise E5"}
        };

        // Add missing SkuIDs to the dictionary
        private static void UpdateSkuIDDictionary(Object[] AdLicenses)
        {
            foreach (PSObject AdLicense in AdLicenses)
            {
                if (!AzureADClass.AzureSkuIDDictionary.ContainsKey(Convert.ToString(GetValue(AdLicense, "SkuId"))))
                {
                    AzureADClass.AzureSkuIDDictionary.Add(Convert.ToString(GetValue(AdLicense, "SkuId")),Convert.ToString(GetValue(AdLicense, "SkuPartNumber")));
                }
            }
        }

		private static readonly Dictionary<string, string> Replacements = new Dictionary<string, string>()
        {
            //{System.Environment.NewLine, ""},
            //{",", ";"},
            {"\"", "'"}
        };

        private static object GetValue(PSObject record, string name)
        {
            if (record == null) return null;
            var member = record.Members[name];
            return member == null ? null : member.Value;
        }

        private static PSObject FailedRecord(object source, Exception error)
        {
            PSObject row = new PSObject();
            object id = null;
            object displayName = null;
            try { id = GetValue(source as PSObject, "Id"); } catch { }
            try { displayName = GetValue(source as PSObject, "DisplayName"); } catch { }
            row.Members.Add(new PSNoteProperty("Id", id));
            row.Members.Add(new PSNoteProperty("DisplayName", displayName));
            row.Members.Add(new PSNoteProperty("CollectionStatus", "Unknown"));
            row.Members.Add(new PSNoteProperty("CollectionError", error.GetType().FullName));
            return row;
        }

        private static object GetDictionaryValue(object source, string key)
        {
            var dictionary = source as System.Collections.IDictionary;
            return dictionary != null && dictionary.Contains(key) ? dictionary[key] : null;
        }

        private static string JoinValues(object source)
        {
            if (source == null) return null;
            var values = source as System.Collections.IEnumerable;
            if (values == null || source is string) return CleanString(source);
            var result = new List<string>();
            foreach (var value in values)
                if (value != null) result.Add(CleanString(value));
            return string.Join(",", result.ToArray());
        }

        public static string CleanString(Object StringtoClean)
        {
            if (StringtoClean == null) return null;
            // Remove extra spaces and new lines
            string CleanedString = string.Join(" ", ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
            foreach (string Replacement in Replacements.Keys)
            {
                CleanedString = CleanedString.Replace(Replacement, Replacements[Replacement]);
            }
            return CleanedString;
        }

        public static int ObjectCount(Object[] ADRObject)
        {
            return ADRObject == null ? 0 : ADRObject.Length;
        }

        public static Object[] TenantParser(Object[] AdTenant, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdTenant, numOfThreads, "Tenant");
            return ADRObj;
        }

        public static Object[] DomainParser(Object[] AdDomain, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDomain, numOfThreads, "Domain");
            return ADRObj;
        }

        public static Object[] LicenseParser(Object[] AdLicenses, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdLicenses, numOfThreads, "Licenses");
            return ADRObj;
        }

        public static Object[] UserParser(Object[] AdUsers, Object[] AdLicenses, int numOfThreads)
        {
            AzureADClass.UpdateSkuIDDictionary(AdLicenses);
            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "Users");
            return ADRObj;
        }

        public static Object[] ServicePrincipalParser(Object[] AdServicePrincipals, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdServicePrincipals, numOfThreads, "ServicePrincipals");
            return ADRObj;
        }

        public static Object[] DirectoryRoleParser(Object[] AdDirectoryRoles, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDirectoryRoles, numOfThreads, "DirectoryRoles");
            return ADRObj;
        }

        public static Object[] GroupParser(Object[] AdGroups, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, "Groups");
            return ADRObj;
        }

        public static Object[] DeviceParser(Object[] AdDevices, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDevices, numOfThreads, "Devices");
            return ADRObj;
        }

        static Object[] runProcessor(Object[] arrayToProcess, int numOfThreads, string processorType)
        {
            if (arrayToProcess == null || arrayToProcess.Length == 0) return new Object[0];
            if (numOfThreads < 1) throw new ArgumentOutOfRangeException("numOfThreads");
            int totalRecords = arrayToProcess.Length;
            IRecordProcessor recordProcessor = recordProcessorFactory(processorType);
            IResultsHandler resultsHandler = new SimpleResultsHandler ();
            int numberOfRecordsPerThread = totalRecords / numOfThreads;
            int remainders = totalRecords % numOfThreads;

            Thread[] threads = new Thread[numOfThreads];
            for (int i = 0; i < numOfThreads; i++)
            {
                int numberOfRecordsToProcess = numberOfRecordsPerThread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Array.Copy(arrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
                ProcessorThread processorThread = new ProcessorThread(i, recordProcessor, resultsHandler, sliceToProcess);
                threads[i] = new Thread(processorThread.processThreadRecords);
                threads[i].Start();
            }
            foreach (Thread t in threads)
            {
                t.Join();
            }

            return resultsHandler.finalise();
        }

        static IRecordProcessor recordProcessorFactory(string name)
        {
            switch (name)
            {
                case "Tenant":
                    return new TenantRecordProcessor();
                case "Domain":
                    return new DomainRecordProcessor();
                case "Licenses":
                    return new LicenseRecordProcessor();
                case "Users":
                    return new UserRecordProcessor();
                case "ServicePrincipals":
                    return new ServicePrincipalRecordProcessor();
                case "DirectoryRoles":
                    return new DirectoryRoleRecordProcessor();
                case "Groups":
                    return new GroupRecordProcessor();
                case "Devices":
                    return new DeviceRecordProcessor();
            }
            throw new ArgumentException("Invalid processor type " + name);
        }

        class ProcessorThread
        {
            readonly int id;
            readonly IRecordProcessor recordProcessor;
            readonly IResultsHandler resultsHandler;
            readonly Object[] objectsToBeProcessed;

            public ProcessorThread(int id, IRecordProcessor recordProcessor, IResultsHandler resultsHandler, Object[] objectsToBeProcessed)
            {
                this.recordProcessor = recordProcessor;
                this.id = id;
                this.resultsHandler = resultsHandler;
                this.objectsToBeProcessed = objectsToBeProcessed;
            }

            public void processThreadRecords()
            {
                for (int i = 0; i < objectsToBeProcessed.Length; i++)
                {
                    try
                    {
                        Object[] result = recordProcessor.processRecord(objectsToBeProcessed[i]);
                        resultsHandler.processResults(result); //this is a thread safe operation
                    }
                    catch (Exception error)
                    {
                        resultsHandler.processResults(new PSObject[] { FailedRecord(objectsToBeProcessed[i], error) });
                    }
                }
            }
        }

        //The interface and implmentation class used to process a record (this implemmentation just returns a log type string)

        interface IRecordProcessor
        {
            PSObject[] processRecord(Object record);
        }

        class TenantRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADTenant = (PSObject) record;

                    List<PSObject> AzureADTenantObjList = new List<PSObject>();
                    PSObject AzureADTenantObj = new PSObject();
                    int icount = 0;

                    //List<Microsoft.Azure.PowerShell.Cmdlets.Resources.MSGraph.Models.ApiV10.IMicrosoftGraphVerifiedDomain> VerifiedDomainsList = new List<Microsoft.Azure.PowerShell.Cmdlets.Resources.MSGraph.Models.ApiV10.IMicrosoftGraphVerifiedDomain>();
                    List<string> TechnicalNotificationMailsList = new List<string>();
                    //List<Microsoft.Open.AzureAD.Model.AssignedPlan> AssignedPlanList = new List<Microsoft.Open.AzureAD.Model.AssignedPlan>();
                    //int count = 0;
                    string TechnicalNotificationMails = null;

                    Object[] ObjValues = new Object[]{
                        "DisplayName", CleanString(GetValue(AzureADTenant, "DisplayName"))
                    };

                    for (icount = 0; icount < ObjValues.Length; icount++)
                    {
                        AzureADTenantObj = new PSObject();
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                        AzureADTenantObjList.Add( AzureADTenantObj );
                        icount++;
                    }

                    //if (((List<Microsoft.Open.AzureAD.Model.VerifiedDomain>) GetValue(AzureADTenant, "VerifiedDomains")).Count != 0)
                    //{
                    //    VerifiedDomainsList = (List<Microsoft.Open.AzureAD.Model.VerifiedDomain>) GetValue(AzureADTenant, "VerifiedDomains");
                    //    count = 0;
                    //    foreach (Microsoft.Open.AzureAD.Model.VerifiedDomain value in VerifiedDomainsList)
                    //    {
                    //        ObjValues = new Object[]{
                    //            "VerifiedDomain(" + count + ") - Name", value.Name,
                    //            "VerifiedDomain(" + count + ") - Type", value.Type,
                    //            "VerifiedDomain(" + count + ") - Capabilities", value.Capabilities,
                    //            "VerifiedDomain(" + count + ") - _Default", value._Default,
                    //            "VerifiedDomain(" + count + ") - Initial", value.Initial,
                    //            "VerifiedDomain(" + count + ") - Id", value.Id
                    //        };
                    //        for (icount = 0; icount < ObjValues.Length; icount++)
                    //        {
                    //            AzureADTenantObj = new PSObject();
                    //            AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                    //            AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                    //            AzureADTenantObjList.Add( AzureADTenantObj );
                    //            icount++;
                    //        }
                    //        count++;
                    //    }
                    //}

                    TechnicalNotificationMails = JoinValues(GetValue(AzureADTenant, "TechnicalNotificationMails"));

                    ObjValues = new Object[]{
                        "DirSyncEnabled", GetValue(AzureADTenant, "DirSyncEnabled"),
                        "CompanyLastDirSyncTime", GetValue(AzureADTenant, "CompanyLastDirSyncTime"),
                        "Street", GetValue(AzureADTenant, "Street"),
                        "City", GetValue(AzureADTenant, "City"),
                        "PostalCode", GetValue(AzureADTenant, "PostalCode"),
                        "State", GetValue(AzureADTenant, "State"),
                        "Country", GetValue(AzureADTenant, "Country"),
                        "CountryLetterCode", GetValue(AzureADTenant, "CountryLetterCode"),
                        "TechnicalNotificationMails", TechnicalNotificationMails,
                        "TelephoneNumber", GetValue(AzureADTenant, "TelephoneNumber")
                    };
                    for (icount = 0; icount < ObjValues.Length; icount++)
                    {
                        AzureADTenantObj = new PSObject();
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                        AzureADTenantObjList.Add( AzureADTenantObj );
                        icount++;
                    }

                    //if (((List<Microsoft.Open.AzureAD.Model.AssignedPlan>) GetValue(AzureADTenant, "AssignedPlans")).Count != 0)
                    //{
                    //    AssignedPlanList = (List<Microsoft.Open.AzureAD.Model.AssignedPlan>) GetValue(AzureADTenant, "AssignedPlans");
                    //    count = 0;
                    //    foreach (Microsoft.Open.AzureAD.Model.AssignedPlan value in AssignedPlanList)
                    //    {
                    //        ObjValues = new Object[]{
                    //            "AssignedPlan(" + count + ") - AssignedTimestamp", value.AssignedTimestamp,
                    //            "AssignedPlan(" + count + ") - CapabilityStatus", value.CapabilityStatus,
                    //            "AssignedPlan(" + count + ") - Service", value.Service,
                    //            "AssignedPlan(" + count + ") - ServicePlanId", value.ServicePlanId
                    //        };
                    //        for (icount = 0; icount < ObjValues.Length; icount++)
                    //        {
                    //            AzureADTenantObj = new PSObject();
                    //            AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                    //            AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                    //            AzureADTenantObjList.Add( AzureADTenantObj );
                    //            icount++;
                    //        }
                    //        count++;
                    //    }
                    //}

                    ObjValues = new Object[]{
                        "ObjectType", GetValue(AzureADTenant, "ObjectType"),
                        "ObjectId", GetValue(AzureADTenant, "ObjectId")
                    };
                    for (icount = 0; icount < ObjValues.Length; icount++)
                    {
                        AzureADTenantObj = new PSObject();
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                        AzureADTenantObjList.Add( AzureADTenantObj );
                        icount++;
                    }

                    return AzureADTenantObjList.ToArray();
                }
                catch (Exception e)
                {
                    return new PSObject[] { FailedRecord(record, e) };
                }
            }
        }

        class DomainRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADDomain = (PSObject) record;

                    //List<string> SupportedServicesList = new List<string>();
                    string SupportedServices = null;

                    PSObject AzureADDomainObj = new PSObject();
                    AzureADDomainObj.Members.Add(new PSNoteProperty("Id", CleanString(GetValue(AzureADDomain, "Id"))));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("AuthenticationType", GetValue(AzureADDomain, "authenticationType")));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("AvailabilityStatus", GetValue(AzureADDomain, "availabilityStatus")));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isAdminManaged", GetValue(AzureADDomain, "isAdminManaged")));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isDefault", GetValue(AzureADDomain, "isDefault")));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isInitial", GetValue(AzureADDomain, "isInitial")));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isRoot", GetValue(AzureADDomain, "isRoot")));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isVerified", GetValue(AzureADDomain, "isVerified")));

                    SupportedServices = JoinValues(GetValue(AzureADDomain, "SupportedServices"));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("SupportedServices", SupportedServices));

                    //AzureADDomainObj.Members.Add(new PSNoteProperty("SupportedServices", GetValue(AzureADDomain, "SupportedServices")));

                    //AzureADDomainObj.Members.Add(new PSNoteProperty("ForceDeleteState", GetValue(AzureADDomain, "ForceDeleteState")));
                    //AzureADDomainObj.Members.Add(new PSNoteProperty("State", GetValue(AzureADDomain, "State")));
                    return new PSObject[] { AzureADDomainObj };
                }
                catch (Exception e)
                {
                    return new PSObject[] { FailedRecord(record, e) };
                }
            }
        }

        class LicenseRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADLicense = (PSObject) record;

                    //string ServicePlanName = null;
                    //List<Microsoft.Open.AzureAD.Model.ServicePlanInfo> ServicePlansList = new List<Microsoft.Open.AzureAD.Model.ServicePlanInfo>();

                    PSObject AzureADLicenseObj = new PSObject();

                    AzureADLicenseObj.Members.Add(new PSNoteProperty("SkuPartNumber", GetValue(AzureADLicense, "SkuPartNumber")));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("SkuId", GetValue(AzureADLicense, "SkuId")));

                    AzureADLicenseObj.Members.Add(new PSNoteProperty("ConsumedUnits", GetValue(AzureADLicense, "ConsumedUnits")));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("AppliesTo", GetValue(AzureADLicense, "AppliesTo")));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("CapabilityStatus", GetValue(AzureADLicense, "CapabilityStatus")));

                    //if (((List<Microsoft.Open.AzureAD.Model.ServicePlanInfo>) GetValue(AzureADLicense, "ServicePlans")).Count != 0)
                    //{
                    //    ServicePlansList = (List<Microsoft.Open.AzureAD.Model.ServicePlanInfo>) GetValue(AzureADLicense, "ServicePlans");
                    //    foreach (Microsoft.Open.AzureAD.Model.ServicePlanInfo value in ServicePlansList)
                    //    {
                    //        ServicePlanName = ServicePlanName + "," + Convert.ToString(value.ServicePlanName);
                    //    }
                    //    ServicePlanName = ServicePlanName.TrimStart(',');
                    //}

                    //AzureADLicenseObj.Members.Add(new PSNoteProperty("PrepaidUnits-Enabled", (((Microsoft.Open.AzureAD.Model.LicenseUnitsDetail) GetValue(AzureADLicense, "PrepaidUnits")).Enabled)));
                    //AzureADLicenseObj.Members.Add(new PSNoteProperty("PrepaidUnits-Suspended", (((Microsoft.Open.AzureAD.Model.LicenseUnitsDetail) GetValue(AzureADLicense, "PrepaidUnits")).Suspended)));
                    //AzureADLicenseObj.Members.Add(new PSNoteProperty("PrepaidUnits-Warning", (((Microsoft.Open.AzureAD.Model.LicenseUnitsDetail) GetValue(AzureADLicense, "PrepaidUnits")).Warning)));
                    //AzureADLicenseObj.Members.Add(new PSNoteProperty("ServicePlans-Name", ServicePlanName));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("ObjectId", GetValue(AzureADLicense, "Id") ?? GetValue(AzureADLicense, "ObjectId")));

                    return new PSObject[] { AzureADLicenseObj };
                }
                catch (Exception e)
                {
                    return new PSObject[] { FailedRecord(record, e) };
                }
            }
        }

        class UserRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureAdUser = (PSObject) record;

                    //List<Microsoft.Open.AzureAD.Model.AssignedLicense> AssignedLicensesList = new List<Microsoft.Open.AzureAD.Model.AssignedLicense>();
                    //string AssignedLicenses = null;
                    //string AssignedLicensesName = null;

                    PSObject AzureADUserObj = new PSObject();
                    AzureADUserObj.Members.Add(new PSNoteProperty("UserPrincipalName", CleanString(GetValue(AzureAdUser, "UserPrincipalName"))));
                    AzureADUserObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(GetValue(AzureAdUser, "DisplayName"))));
                    AzureADUserObj.Members.Add(new PSNoteProperty("Enabled", GetValue(AzureAdUser, "AccountEnabled")));
                    AzureADUserObj.Members.Add(new PSNoteProperty("UserType", GetValue(AzureAdUser, "UserType")));
                    AzureADUserObj.Members.Add(new PSNoteProperty("JobTitle", GetValue(AzureAdUser, "JobTitle")));
                    AzureADUserObj.Members.Add(new PSNoteProperty("Department", GetValue(AzureAdUser, "Department")));
                    AzureADUserObj.Members.Add(new PSNoteProperty("Mobile", GetValue(AzureAdUser, "Mobile")));
                    AzureADUserObj.Members.Add(new PSNoteProperty("TelephoneNumber", GetValue(AzureAdUser, "TelephoneNumber")));
                    AzureADUserObj.Members.Add(new PSNoteProperty("DirSyncEnabled", GetValue(AzureAdUser, "DirSyncEnabled")));
                    AzureADUserObj.Members.Add(new PSNoteProperty("OnPremisesSecurityIdentifier",  AzureAdUser.Members["OnPremisesSecurityIdentifier"] != null ? GetValue(AzureAdUser, "OnPremisesSecurityIdentifier") : null));
                    AzureADUserObj.Members.Add(new PSNoteProperty("OnPremisesDistinguishedName", CleanString(GetDictionaryValue(GetValue(AzureAdUser, "ExtensionProperty"), "onPremisesDistinguishedName"))));
                    AzureADUserObj.Members.Add(new PSNoteProperty("LastDirSyncTime", GetValue(AzureAdUser, "LastDirSyncTime")));
                    AzureADUserObj.Members.Add(new PSNoteProperty("whenCreated", GetValue(AzureAdUser, "CreatedDateTime") ?? GetDictionaryValue(GetValue(AzureAdUser, "ExtensionProperty"), "createdDateTime")));

                    //if (((List<Microsoft.Open.AzureAD.Model.AssignedLicense>) GetValue(AzureAdUser, "AssignedLicenses")).Count != 0)
                    //{
                    //    AssignedLicensesList = (List<Microsoft.Open.AzureAD.Model.AssignedLicense>) GetValue(AzureAdUser, "AssignedLicenses");
                    //    foreach (Microsoft.Open.AzureAD.Model.AssignedLicense value in AssignedLicensesList)
                    //    {
                    //        AssignedLicenses = AssignedLicenses + "," + Convert.ToString(value.SkuId);
                    //        AssignedLicensesName = AssignedLicensesName + "," + (AzureADClass.AzureSkuIDDictionary.ContainsKey(Convert.ToString(value.SkuId)) ? AzureADClass.AzureSkuIDDictionary[Convert.ToString(value.SkuId)] : Convert.ToString(value.SkuId));
                    //    }
                    //    AssignedLicenses = AssignedLicenses.TrimStart(',');
                    //    AssignedLicensesName = AssignedLicensesName.TrimStart(',');
                    //}

                    //AzureADUserObj.Members.Add(new PSNoteProperty("Assigned Licenses (Count)", ((List<Microsoft.Open.AzureAD.Model.AssignedLicense>) GetValue(AzureAdUser, "AssignedLicenses")).Count));
                    //AzureADUserObj.Members.Add(new PSNoteProperty("AssignedLicensesName", AssignedLicensesName));
                    //AzureADUserObj.Members.Add(new PSNoteProperty("AssignedLicenses", AssignedLicenses));

                    return new PSObject[] { AzureADUserObj };
                }
                catch (Exception e)
                {
                    return new PSObject[] { FailedRecord(record, e) };
                }
            }
        }

        class ServicePrincipalRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureAdServicePrincipal = (PSObject) record;

                    List<string> AADRList = new List<string>();
                    string AADRString = null;
                    //string AADRString2 = null;
                    //List<Microsoft.Open.AzureAD.Model.AppRole> AppRolesList = new List<Microsoft.Open.AzureAD.Model.AppRole>();
                    //List<Microsoft.Open.AzureAD.Model.OAuth2Permission> OAuth2PermissionList = new List<Microsoft.Open.AzureAD.Model.OAuth2Permission>();

                    PSObject AzureAdServicePrincipalObj = new PSObject();
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(GetValue(AzureAdServicePrincipal, "DisplayName"))));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("Enabled", GetValue(AzureAdServicePrincipal, "AccountEnabled")));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("PublisherName", GetValue(AzureAdServicePrincipal, "PublisherName")));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("AppId", GetValue(AzureAdServicePrincipal, "AppId")));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("ObjectId", GetValue(AzureAdServicePrincipal, "Id") ?? GetValue(AzureAdServicePrincipal, "ObjectId")));

                    AADRList = new List<string>();
                    AADRString = null;
                    AADRString = JoinValues(GetValue(AzureAdServicePrincipal, "ServicePrincipalNames"));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("ServicePrincipalNames", AADRString));

                    AADRList = new List<string>();
                    AADRString = null;
                    AADRString = JoinValues(GetValue(AzureAdServicePrincipal, "ReplyUrls"));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("ReplyUrls", AADRString));

                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("LogoutUrl", GetValue(AzureAdServicePrincipal, "LogoutUrl")));

                    object scopeSource = GetValue(AzureAdServicePrincipal, "Oauth2PermissionScopes") ?? GetValue(AzureAdServicePrincipal, "Oauth2Permissions");
                    object scopeCount = null;
                    string scopeText = null;
                    var scopes = scopeSource as System.Collections.IEnumerable;
                    if (scopes != null)
                    {
                        var descriptions = new List<string>();
                        int count = 0;
                        foreach (object scope in scopes)
                        {
                            var item = PSObject.AsPSObject(scope);
                            descriptions.Add(Convert.ToString(GetValue(item, "Id")) + ":" + Convert.ToString(GetValue(item, "Value")));
                            count++;
                        }
                        scopeCount = count;
                        scopeText = string.Join(",", descriptions.ToArray());
                    }
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("Exposed OAuth2 Scopes (Count)", scopeCount));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("Exposed OAuth2 Scopes", scopeText));
                    // Backward-compatible count describes exposed scopes, never permission grants.
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("OAuth2Permissions (Count)", scopeCount));

                    return new PSObject[] { AzureAdServicePrincipalObj };
                }
                catch (Exception e)
                {
                    return new PSObject[] { FailedRecord(record, e) };
                }
            }
        }

        class DirectoryRoleRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADDirectoryRole = (PSObject) record;

                    PSObject AzureADDirectoryRoleObj = new PSObject();
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(GetValue(AzureADDirectoryRole, "DisplayName"))));
                    //AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("RoleDisabled", GetValue(AzureADDirectoryRole, "RoleDisabled")));
                    //AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("IsSystem", GetValue(AzureADDirectoryRole, "IsSystem")));
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("Description", CleanString(GetValue(AzureADDirectoryRole, "Description"))));
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("Id", GetValue(AzureADDirectoryRole, "Id")));
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("RoleTemplateId", GetValue(AzureADDirectoryRole, "RoleTemplateId")));
                    return new PSObject[] { AzureADDirectoryRoleObj };
                }
                catch (Exception e)
                {
                    return new PSObject[] { FailedRecord(record, e) };
                }
            }
        }

        class GroupRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADGroup = (PSObject) record;

                    PSObject AzureADGroupObj = new PSObject();
                    AzureADGroupObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(GetValue(AzureADGroup, "DisplayName"))));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("OnPremisesSyncEnabled", GetValue(AzureADGroup, "OnPremisesSyncEnabled")));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("OnPremisesLastSyncDateTime", GetValue(AzureADGroup, "OnPremisesLastSyncDateTime")));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("OnPremisesSecurityIdentifier", GetValue(AzureADGroup, "OnPremisesSecurityIdentifier")));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("SecurityEnabled", GetValue(AzureADGroup, "SecurityEnabled")));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("Description", CleanString(GetValue(AzureADGroup, "Description"))));
                    return new PSObject[] { AzureADGroupObj };
                }
                catch (Exception e)
                {
                    return new PSObject[] { FailedRecord(record, e) };
                }
            }
        }

        class DeviceRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADDevice = (PSObject) record;

                    PSObject AzureADDeviceObj = new PSObject();
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(GetValue(AzureADDevice, "DisplayName"))));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("AccountEnabled", GetValue(AzureADDevice, "AccountEnabled")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("OnPremisesSyncEnabled", GetValue(AzureADDevice, "OnPremisesSyncEnabled")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("OnPremisesLastSyncDateTime", GetValue(AzureADDevice, "OnPremisesLastSyncDateTime")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("OperatingSystem", GetValue(AzureADDevice, "OperatingSystem")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("OperatingSystemVersion", GetValue(AzureADDevice, "OperatingSystemVersion")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("ApproximateLastSignInDateTime", GetValue(AzureADDevice, "ApproximateLastSignInDateTime")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("TrustType", GetValue(AzureADDevice, "TrustType")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("ProfileType", GetValue(AzureADDevice, "ProfileType")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("IsManaged", GetValue(AzureADDevice, "IsManaged")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("DeviceOwnership", GetValue(AzureADDevice, "DeviceOwnership")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("EnrollmentProfileName", GetValue(AzureADDevice, "EnrollmentProfileName")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("IsCompliant", GetValue(AzureADDevice, "IsCompliant")));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("DeviceId", GetValue(AzureADDevice, "DeviceId")));
                    return new PSObject[] { AzureADDeviceObj };
                }
                catch (Exception e)
                {
                    return new PSObject[] { FailedRecord(record, e) };
                }
            }
        }

        //The interface and implmentation class used to handle the results (this implementation just writes the strings to a file)

        interface IResultsHandler
        {
            void processResults(Object[] t);

            Object[] finalise();
        }

        class SimpleResultsHandler : IResultsHandler
        {
            private Object lockObj = new Object();
            private List<Object> processed = new List<Object>();

            public SimpleResultsHandler()
            {
            }

            public void processResults(Object[] results)
            {
                lock (lockObj)
                {
                    if (results.Length != 0)
                    {
                        for (var i = 0; i < results.Length; i++)
                        {
                            processed.Add((PSObject)results[i]);
                        }
                    }
                }
            }

            public Object[] finalise()
            {
                return processed.ToArray();
            }
        }
	}
}
"@

Function Initialize-AADRParserType
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $Source)

    # CLR types cannot be replaced or unloaded from the session. Reuse this source
    # version, but give changed source a different namespace to avoid stale parsers.
    $hasher = [System.Security.Cryptography.SHA256]::Create()
    Try
    {
        $hash = [BitConverter]::ToString($hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Source))).Replace('-', '')
    }
    Finally { $hasher.Dispose() }

    $namespace = "AADRecon.V$hash"
    $typeName = "$namespace.AzureADClass"
    $parserType = $typeName -as [type]
    If ($parserType) { return $parserType }

    If ($Source -notmatch '(?m)^namespace AADRecon\s*$')
    {
        throw 'Parser source does not declare the expected AADRecon namespace.'
    }
    $versionedSource = $Source -replace '(?m)^namespace AADRecon\s*$', "namespace $namespace"
    # Default framework references work on both Windows PowerShell 5.1 and Core.
    # Compiling once avoids the former overlapping Core/CLR compilation branches.
    $DebugPreference = 'SilentlyContinue'
    Add-Type -TypeDefinition $versionedSource -ErrorAction Stop | Out-Null
    $parserType = $typeName -as [type]
    If (-Not $parserType) { throw 'Parser compilation did not load the expected class.' }
    return $parserType
}

Function Join-AADRPath
{
    param([Parameter(Mandatory = $true)] [object[]] $Parts)
    $path = [string]$Parts[0]
    For ($i = 1; $i -lt $Parts.Count; $i++) { $path = Join-Path -Path $path -ChildPath ([string]$Parts[$i]) }
    $path
}

Function ConvertTo-AADRDateTime
{
    param($Value)
    If ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    Try
    {
        If ($Value -is [DateTimeOffset]) { return $Value.UtcDateTime }
        If ($Value -is [DateTime]) { return $Value.ToUniversalTime() }
        return ([DateTimeOffset]::Parse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture)).UtcDateTime
    }
    Catch
    {
        Add-AADREvent -Operation DateConversion -Status Unavailable -Detail 'A date field could not be parsed; it remains unknown.'
        return $null
    }
}

Function Get-AADRProperty
{
    param($Object, [string] $Name)
    If ($null -eq $Object) { return $null }
    If ($Object -is [System.Collections.IDictionary]) { return ,($Object[$Name]) }
    $p = $Object.PSObject.Properties[$Name]
    If ($p) { return ,($p.Value) }
    return $null
}

Function Test-AADRGraphAuthenticationError
{
    param([Parameter(Mandatory = $true)] $Record)
    If ($Record.FullyQualifiedErrorId -match 'AADRTokenVerificationFailed' -and $Record.Exception.Data['AADRSafeDiagnostic'])
    { return $Record.Exception.Data['AADRSafeDiagnostic'].Category -eq 'AuthenticationFailed' }
    $status = Get-AADRProperty $Record.Exception 'ResponseStatusCode'
    If ($null -eq $status) { $status = Get-AADRProperty $Record.Exception 'StatusCode' }
    If ($null -eq $status) { $status = Get-AADRProperty (Get-AADRProperty $Record.Exception 'Response') 'StatusCode' }
    $details = @($Record.FullyQualifiedErrorId, $Record.Exception.Message, $Record.ErrorDetails.Message) -join ' '
    return ([string]$status -in @('401', 'Unauthorized') -or
        $details -match 'AADRGraphAuthenticationFailed|AADRTokenRenewalFailed|InvalidAuthenticationToken|ExpiredAuthenticationToken|Invalid token lifetime|AuthenticationFailed|token.*expired')
}

Function Stop-AADRGraphAuthentication
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] $Record)
    If ($Record.FullyQualifiedErrorId -match 'AADRGraphAuthenticationFailed|AADRTokenRenewalFailed|AADRTokenVerificationFailed') { $PSCmdlet.ThrowTerminatingError($Record) }
    # Do not include request headers, bearer tokens, or raw SDK exception details.
    $message = 'Microsoft Graph rejected the authentication token (401). An invalid token lifetime can mean the token has expired or is not yet valid; check the system UTC clock. Automatic recovery was unavailable or the replacement was also rejected. Acquire a fresh Microsoft Graph token, configure AccessTokenProvider, use SeamlessSSO with a valid Kerberos session, or use SDK-managed interactive/certificate/application authentication. Completed collector checkpoints are retained. Resume requires the same identified Graph account/application, script, and settings; the failed collector restarts.'
    $exception = New-Object System.Security.Authentication.AuthenticationException -ArgumentList $message
    $exception.Data['AADRSafeDiagnostic'] = Get-AADRErrorDiagnostic $Record
    $errorRecord = New-Object System.Management.Automation.ErrorRecord -ArgumentList $exception, 'AADRGraphAuthenticationFailed', ([System.Management.Automation.ErrorCategory]::AuthenticationError), $null
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}

Function Invoke-AADRReadRequest
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [scriptblock] $Query, [int] $MaxRetries = 3)
    $authenticationRetried = $false
    For ($attempt = 0; ; $attempt++)
    {
        Update-AADRTokenSession
        Try { $result = @(& $Query); return $result }
        Catch
        {
            $record = $_
            If (Test-AADRGraphAuthenticationError -Record $record)
            {
                If (-Not $authenticationRetried -and $record.FullyQualifiedErrorId -notmatch 'AADRGraphAuthenticationFailed|AADRTokenRenewalFailed' -and
                    $script:AADRTokenSession -and $script:AADRTokenSession.Ready -and $script:AADRTokenSession.Provider)
                {
                    $authenticationRetried = $true
                    Update-AADRTokenSession -Force
                    continue
                }
                Stop-AADRGraphAuthentication -Record $record
            }
            $status = Get-AADRProperty $record.Exception 'ResponseStatusCode'
            If ($null -eq $status) { $status = Get-AADRProperty $record.Exception 'StatusCode' }
            If ($null -eq $status) { $status = Get-AADRProperty (Get-AADRProperty $record.Exception 'Response') 'StatusCode' }
            $details = @($record.FullyQualifiedErrorId, $record.Exception.Message, $record.ErrorDetails.Message) -join ' '
            $transient = ([string]$status -in @('429','500','502','503','504','TooManyRequests','ServiceUnavailable','GatewayTimeout')) -or
                ($details -match 'TooManyRequests|ServiceUnavailable|GatewayTimeout|InternalServerError')
            If ($record.Exception.Data['AADRRetryExhausted']) { throw }
            If (-Not $transient) { throw }
            If ($attempt -ge $MaxRetries) { $record.Exception.Data['AADRRetryExhausted'] = $true; throw }
            $delay = [Math]::Min(30, [Math]::Pow(2, $attempt + 1))
            $headers = Get-AADRProperty (Get-AADRProperty $record.Exception 'Response') 'Headers'
            If ($null -eq $headers) { $headers = Get-AADRProperty $record.Exception 'ResponseHeaders' }
            $retryAfter = Get-AADRProperty $headers 'Retry-After'
            $seconds = 0
            If ($retryAfter)
            {
                $text = [string](@($retryAfter)[0])
                $retryDate = [DateTimeOffset]::MinValue
                If (-Not [int]::TryParse($text, [ref]$seconds) -and
                    [DateTimeOffset]::TryParse($text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$retryDate))
                { $seconds = [int][Math]::Max(0, [Math]::Ceiling(($retryDate.UtcDateTime - [DateTime]::UtcNow).TotalSeconds)) }
            }
            Else
            {
                $typedRetry = Get-AADRProperty $headers 'RetryAfter'
                If ($typedRetry.Delta) { $seconds = [int][Math]::Ceiling($typedRetry.Delta.TotalSeconds) }
                ElseIf ($typedRetry.Date) { $seconds = [int][Math]::Max(0, [Math]::Ceiling(($typedRetry.Date.UtcDateTime - [DateTime]::UtcNow).TotalSeconds)) }
            }
            # Do not retry early when the server asks for more than the bounded wait.
            If ($seconds -gt 60) { $record.Exception.Data['AADRRetryExhausted'] = $true; throw }
            $delay = [Math]::Max($delay, $seconds)
            Write-Verbose "Transient Graph failure; retry $($attempt + 1)/$MaxRetries in $delay seconds."
            Start-Sleep -Seconds $delay
        }
    }
}

Function Get-AADRGraphObject
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $Path)
    $uri = 'https://graph.microsoft.com/v1.0/' + $Path.TrimStart('/')
    Invoke-AADRReadRequest -Query { Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop }
}

Function Get-AADRGraphPages
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $Path)
    $uri = If ($Path.StartsWith('https://')) { $Path } Else { 'https://graph.microsoft.com/v1.0/' + $Path.TrimStart('/') }
    $seen = @{}
    $rows = New-Object System.Collections.Generic.List[object]
    While ($uri)
    {
        $address = [uri]$uri
        If ($address.Scheme -ne 'https' -or $address.Host -ne 'graph.microsoft.com' -or $address.Port -ne 443 -or $address.UserInfo)
        { throw 'Unexpected Graph pagination host.' }
        If ($seen.ContainsKey($uri)) { throw 'Graph returned a repeated pagination link.' }
        $seen[$uri] = $true
        $response = Invoke-AADRReadRequest -Query { Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop }
        $value = Get-AADRProperty $response 'value'
        If ($null -eq $value -and -Not ($response.PSObject.Properties['value'] -or ($response -is [System.Collections.IDictionary] -and $response.Contains('value'))))
        { throw 'Graph collection response did not contain a value array.' }
        ForEach ($row in $value) { If ($null -ne $row) { $rows.Add($row) } }
        $uri = [string](Get-AADRProperty $response '@odata.nextLink')
    }
    $rows.ToArray()
}

Function Add-AADREvent
{
    param([string] $Operation, [string] $Status, [string] $Detail, $Diagnostic)
    If ($script:AADRRun)
    {
        $event = [pscustomobject]@{TimeUtc=[DateTime]::UtcNow.ToString('o'); Collector=$script:AADRCurrentCollector; Operation=$Operation; Status=$Status; Detail=$Detail}
        If ($null -ne $Diagnostic) { $event | Add-Member NoteProperty Diagnostic $Diagnostic }
        $script:AADRRun.Events.Add($event)
    }
}

Function Test-AADRControlFlowError
{
    param($Record)
    $exception = $Record.Exception
    For ($i=0; $null -ne $exception -and $i -lt 10; $i++)
    {
        If ($exception -is [System.Management.Automation.PipelineStoppedException] -or
            $exception -is [System.OperationCanceledException] -or $exception -is [System.OutOfMemoryException]) { return $true }
        $exception = $exception.InnerException
    }
    return $false
}

Function Get-AADRErrorDiagnostic
{
    param([Parameter(Mandatory = $true)] $Record)
    If ($Record.FullyQualifiedErrorId -match 'AADRGraphAuthenticationFailed|AADRTokenRenewalFailed|AADRTokenVerificationFailed' -and $Record.Exception.Data['AADRSafeDiagnostic'])
    { return $Record.Exception.Data['AADRSafeDiagnostic'] }
    $status = Get-AADRProperty $Record.Exception 'ResponseStatusCode'
    If ($null -eq $status) { $status = Get-AADRProperty $Record.Exception 'StatusCode' }
    If ($null -eq $status) { $status = Get-AADRProperty (Get-AADRProperty $Record.Exception 'Response') 'StatusCode' }
    $httpStatus = 0
    Try { $httpStatus = [int]$status } Catch { $httpStatus = 0 }
    $code = ([string]$Record.FullyQualifiedErrorId -split ',')[0]
    # A PowerShell string throw uses the message itself as FullyQualifiedErrorId.
    # Only known codes may be exposed; even a token-shaped string must stay private.
    If ($code -notin @('AADRGraphAuthenticationFailed','AADRTokenRenewalFailed','InvalidAuthenticationToken','ExpiredAuthenticationToken','AuthenticationFailed','Authentication_RequestFromUnsupportedUserRole','Authorization_RequestDenied','Authorization_IdentityNotFound','AccessDenied','Forbidden','MissingPermission','NonPremiumTenant','Request_ResourceNotFound','ResourceNotFound','Request_BadRequest','BadRequest','TooManyRequests','ServiceUnavailable','GatewayTimeout','InternalServerError','FileOpenFailure','PathNotFound','UnauthorizedAccess','DirectoryNotFound','ParameterArgumentValidationError','ParameterBindingException'))
    { $code = $Record.Exception.GetType().Name }
    $reason = 'Unspecified'
    $details = @($Record.FullyQualifiedErrorId, $Record.Exception.Message, $Record.ErrorDetails.Message) -join ' '
    # Match known descriptions but emit only fixed labels, never the original text.
    If ($details -match 'Invalid token lifetime') { $reason='InvalidTokenLifetime' }
    ElseIf ($details -match 'token.*not yet valid|token.*not.*valid.*before|IDX10222') { $reason='TokenNotYetValid' }
    ElseIf ($details -match 'token.*expired|IDX10223') { $reason='TokenExpired' }
    ElseIf ($details -match 'token could not be read|Unable to decode|IDX12709|IDX14100') { $reason='UnreadableToken' }
    ElseIf ($details -match 'Invalid audience|audience.*validation.*failed|IDX10214') { $reason='WrongAudience' }
    ElseIf ($details -match 'No application to sign out from|No active (Graph )?(session|context)|not signed in') { $reason='NoActiveSession' }
    ElseIf ($details -match 'insufficient privileges|Authorization_RequestDenied|AccessDenied|Forbidden') { $reason='InsufficientPermissions' }
    $category = 'UnexpectedError'
    $guidance = 'Inspect the failed operation and retry that dataset with -Resume after correcting the cause.'
    If (Test-AADRControlFlowError $Record) { $category='Interrupted'; $guidance='Execution was cancelled or could not continue safely.' }
    ElseIf (Test-AADRGraphAuthenticationError $Record) { $category='AuthenticationFailed'; $guidance='Renew the authentication source and resume with the same identity.' }
    ElseIf (Test-AADRGraphPermissionError $Record) { $category='PermissionDenied'; $guidance='Check Graph permissions, delegated role, and required tenant licensing.' }
    ElseIf (Test-AADRGraphNotFoundError $Record) { $category='NotFound'; $guidance='The referenced object or resource is unavailable; retain its known identity and retry if needed.' }
    ElseIf ($httpStatus -eq 429 -or $code -match 'TooManyRequests') { $category='Throttled'; $guidance='Graph throttling persisted after bounded retries. Retry the affected dataset later.' }
    ElseIf ($httpStatus -in @(500,502,503,504) -or $code -match 'ServiceUnavailable|GatewayTimeout|InternalServerError') { $category='ServiceUnavailable'; $guidance='Graph remained unavailable after bounded retries. Retry the affected dataset later.' }
    ElseIf ($Record.Exception -is [System.IO.IOException] -or $Record.Exception -is [UnauthorizedAccessException] -or $Record.CategoryInfo.Category -in @('WriteError','PermissionDenied','ResourceUnavailable')) { $category='InputOutput'; $guidance='Check the file path, available disk space, permissions, and file locks.' }
    ElseIf ($Record.Exception -is [System.Net.WebException] -or $Record.Exception -is [TimeoutException] -or $code -match 'HttpRequestException|Timeout|ConnectionError') { $category='NetworkError'; $guidance='Check network connectivity and retry the affected dataset.' }
    ElseIf ($Record.CategoryInfo.Category -in @('InvalidArgument','InvalidData','ParserError') -or $Record.Exception -is [FormatException]) { $category='InvalidData'; $guidance='Check input values and the affected record. Unavailable values must remain unknown.' }
    # Never persist raw exception messages, target objects, source lines, URLs, or headers.
    If ($Record.Exception.Message -match 'initial manifest could not be saved|output directory already contains a run') { $guidance='Choose a new writable output directory, or use -Resume with the matching existing run.' }
    ElseIf ($Record.Exception.Message -match 'Resume .*differ|Resume requires|Resume manifest') { $guidance='Resume requires the same verified identity, script, settings, and readable manifest. Use a new directory when these differ.' }
    ElseIf ($Record.Exception.Message -match 'Install Microsoft.Graph') { $guidance='Install the Microsoft.Graph PowerShell module before running collection.' }
    ElseIf ($Record.Exception.Message -match 'TenantID .*required|ClientID and TenantID .*required|-Credential requires') { $guidance='Check tenant ID, application client ID, and the chosen authentication parameters.' }
    If ($reason -in @('InvalidTokenLifetime','TokenNotYetValid','TokenExpired')) { $guidance='Compare the machine UTC clock and token issuer expiry metadata. A new acquisition was attempted when supported; a repeated lifetime rejection requires correcting the token source or clock, not more retries.' }
    ElseIf ($reason -eq 'UnreadableToken') { $guidance='The resource could not read the supplied token. Confirm the token source and Microsoft.Graph.Authentication module version; reissuing the same format may not help.' }
    ElseIf ($reason -eq 'WrongAudience') { $guidance='Obtain a token for Microsoft Graph; tokens for Azure Resource Manager or Azure AD Graph cannot be used here.' }
    $requestId = $null
    $headers = Get-AADRProperty $Record.Exception 'ResponseHeaders'
    If ($null -eq $headers) { $headers = Get-AADRProperty (Get-AADRProperty $Record.Exception 'Response') 'Headers' }
    $headerValue = Get-AADRProperty $headers 'request-id'
    $parsedId = [guid]::Empty
    If ($headerValue -and [guid]::TryParse([string](@($headerValue)[0]), [ref]$parsedId)) { $requestId=$parsedId.ToString() }
    [pscustomobject]@{Category=$category; Code=$code; Reason=$reason; Operation=$null; ExceptionType=$Record.Exception.GetType().FullName; HttpStatus=$(If ($httpStatus -gt 0) {$httpStatus} Else {$null}); RequestId=$requestId; ScriptLine=$Record.InvocationInfo.ScriptLineNumber; Guidance=$guidance}
}

Function Stop-AADRTokenVerification
{
    [CmdletBinding()]
    param($Record, [ValidateSet('GraphConnection','OrganizationProbe','IdentityOrganization','IdentityUser','TokenMetadata')] [string] $Operation)
    $diagnostic = Get-AADRErrorDiagnostic $Record
    $diagnostic.Operation = $Operation
    Add-AADREvent -Operation $Operation -Status $diagnostic.Category -Detail 'Startup/renewal verification failed. Token endpoint success does not establish Graph acceptance.' -Diagnostic $diagnostic
    $exception = New-Object System.Security.Authentication.AuthenticationException -ArgumentList "Token verification failed at $Operation ($($diagnostic.Reason)). $($diagnostic.Guidance)"
    $exception.Data['AADRSafeDiagnostic'] = $diagnostic
    $record = New-Object System.Management.Automation.ErrorRecord -ArgumentList $exception, 'AADRTokenVerificationFailed', ([System.Management.Automation.ErrorCategory]::AuthenticationError), $null
    $PSCmdlet.ThrowTerminatingError($record)
}

Function Disconnect-AADRGraphSession
{
    [CmdletBinding()]
    param()
    Try
    {
        # Failed verification may already have disconnected. No session is a
        # successful cleanup, not another failure to show after the primary one.
        If ((Get-Command Get-MgContext -ErrorAction SilentlyContinue) -and -Not (Get-MgContext -ErrorAction Stop)) { return }
        Disconnect-MgGraph -ErrorAction Stop | Out-Null
    }
    Catch
    {
        $diagnostic = Get-AADRErrorDiagnostic $_
        If ($diagnostic.Reason -eq 'NoActiveSession') { return }
        $script:AADRCleanupFailed=$true
        Add-AADREvent -Operation Cleanup -Status CleanupFailed -Detail 'Graph disconnect failed.' -Diagnostic $diagnostic
        Write-Warning "[Cleanup] Graph disconnect failed ($($diagnostic.Code))." -WarningAction Continue
    }
}

Function Invoke-AADRReportStep
{
    [CmdletBinding()]
    param([string] $Name, [scriptblock] $Query)
    $ErrorActionPreference = 'Stop'
    Try { & $Query }
    Catch
    {
        If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }
        $diagnostic = Get-AADRErrorDiagnostic $_
        Add-AADREvent -Operation $Name -Status ExportFailed -Detail 'Report operation failed; other outputs will still be attempted.' -Diagnostic $diagnostic
        Write-Warning "[$Name] Report operation failed ($($diagnostic.Category)). Other outputs will continue." -WarningAction Continue
    }
}

Function Write-AADRManifest
{
    [CmdletBinding()]
    param([switch] $Required)
    If (-Not $script:AADRRun) { return }
    $path = Join-Path $script:AADRRun.OutputDirectory 'CollectionManifest.json'
    $temp = $path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    Try
    {
        If (Test-Path -LiteralPath $path -PathType Container) { throw [System.IO.IOException]::new('The manifest path is a directory.') }
        $script:AADRRun | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $temp -Encoding UTF8 -ErrorAction Stop
        Move-Item -LiteralPath $temp -Destination $path -Force -ErrorAction Stop
    }
    Catch
    {
        If (Test-AADRControlFlowError $_) { throw }
        If (-Not $script:AADRManifestFailureRecorded)
        {
            $script:AADRManifestFailureRecorded = $true
            Add-AADREvent -Operation Manifest -Status ManifestFailed -Detail 'Manifest could not be saved. Resume may be unavailable; report exports will still be attempted.' -Diagnostic (Get-AADRErrorDiagnostic $_)
            Write-Warning '[Manifest] Could not save collection status. Check the output path, disk space, permissions, and file locks. Resume may be unavailable.' -WarningAction Continue
        }
        If ($Required) { throw 'The initial manifest could not be saved. Choose a writable output directory before collecting.' }
    }
    Finally { If (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue } }
}

Function Get-AADRCollectionRegistry
{
    [ordered]@{
        Tenant = { Get-AADRTenant -Method MSGraph -Threads $Threads }
        Domain = { Get-AADRDomain -Method MSGraph -Threads $Threads }
        Licenses = { Get-AADRLicense -Method MSGraph -Threads $Threads }
        Users = { Get-AADRUser -Method MSGraph -date $date -DormantTimeSpan $DormantTimeSpan -PassMaxAge $PassMaxAge -PageSize $PageSize -Threads $Threads }
        ServicePrincipals = { Get-AADRServicePrincipal -Method MSGraph -date $date -Threads $Threads }
        DirectoryRoles = { Get-AADRDirectoryRole -Method MSGraph -PageSize $PageSize -Threads $Threads }
        DirectoryRoleMembers = { Get-AADRDirectoryRoleMember -Method MSGraph -PageSize $PageSize -Threads $Threads }
        Groups = { Get-AADRGroup -Method MSGraph -date $date -PageSize $PageSize -Threads $Threads }
        GroupMembers = { Get-AADRGroupMember -Method MSGraph -PageSize $PageSize -Threads $Threads }
        Devices = { Get-AADRDevice -Method MSGraph -Threads $Threads }
        ConditionalAccessPolicies = { Get-AADRConditionalAccessPolicy -Method MSGraph }
        Applications = { Get-AADRApplicationInventory }
        ApplicationOwners = { Get-AADRObjectOwners -Kind applications }
        ServicePrincipalOwners = { Get-AADRObjectOwners -Kind servicePrincipals }
        CredentialExpiry = { Get-AADRCredentialExpiry -date $date }
        ConsentGrants = { Get-AADRConsentGrants }
        ApplicationPermissions = { Get-AADRApplicationPermissions }
        PIMEligibility = { Invoke-AADRGraphCollection -Name PIMEligibility -Query { Get-AADRGraphPages 'roleManagement/directory/roleEligibilityScheduleInstances' } }
        MFARegistration = { Get-AADRMFARegistration | Select-Object id,userPrincipalName,userDisplayName,isMfaRegistered,isMfaCapable,isPasswordlessCapable,isSsprRegistered,isSsprEnabled,isSsprCapable,lastUpdatedDateTime,@{Name='MethodsRegistered';Expression={$_.methodsRegistered -join ', '}} }
        SecurityDefaults = { Invoke-AADRGraphCollection -Name SecurityDefaults -Query { Get-AADRGraphObject 'policies/identitySecurityDefaultsEnforcementPolicy' } }
        AuthenticationMethodsPolicy = { Invoke-AADRGraphCollection -Name AuthenticationMethodsPolicy -Query { Get-AADRGraphObject 'policies/authenticationMethodsPolicy' } }
    }
}

Function Invoke-AADRQuery
{
    [CmdletBinding()]
    param([scriptblock] $Query)
    $ErrorActionPreference = 'Stop'
    & $Query
}

Function Invoke-AADRCollectorRun
{
    [CmdletBinding()]
    param([string] $Name, [scriptblock] $Query, [string] $OutputDirectory, [string[]] $Formats, [switch] $Resume)
    $ErrorActionPreference = 'Stop'
    $entry = [pscustomobject]@{Name=$Name; Status='Running'; DataStatus='Running'; RowCount=$null; StartedUtc=[DateTime]::UtcNow.ToString('o'); FinishedUtc=$null; Checkpoint=$null; CheckpointHash=$null; Resumed=$false; Exports=@(); Error=$null; Diagnostic=$null}
    $script:AADRRun.Collectors.Add($entry)
    $script:AADRCurrentCollector = $Name
    $eventStart = $script:AADRRun.Events.Count
    Try
    {
        Write-AADRManifest
        $prior = @($script:AADRPreviousCollectors | Where-Object { $_.Name -eq $Name -and ($_.DataStatus -in @('Success','Empty') -or $_.Status -in @('Success','Empty')) } | Select-Object -Last 1)
        $rows = @()
        If ($Resume -and $prior.Count -eq 1 -and $prior[0].Checkpoint -and $prior[0].CheckpointHash)
        {
            Try
            {
                $cache = Join-AADRPath @($OutputDirectory, '.checkpoints', ($Name + '.clixml'))
                If (-Not (Test-Path -LiteralPath $cache) -or (Get-FileHash -LiteralPath $cache -Algorithm SHA256 -ErrorAction Stop).Hash -ne $prior[0].CheckpointHash)
                { throw 'Checkpoint integrity verification failed.' }
                $cachedRows = Import-Clixml -LiteralPath $cache -ErrorAction Stop
                $rows = @($cachedRows | Where-Object { $null -ne $_ })
                $entry.Resumed = $true
            }
            Catch
            {
                If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }
                Add-AADREvent -Operation Checkpoint -Status Recovered -Detail 'Checkpoint was missing, corrupt, or unreadable. It was rejected; this dataset will be recollected.' -Diagnostic (Get-AADRErrorDiagnostic $_)
                Write-Warning "[$Name] Checkpoint could not be verified. Recollecting this dataset." -WarningAction Continue
            }
        }
        If (-Not $entry.Resumed)
        {
            $warnings = @()
            $rows = @(Invoke-AADRQuery -Query $Query -WarningVariable +warnings | Where-Object { $null -ne $_ })
            ForEach ($warning in $warnings) { Add-AADREvent -Operation $Name -Status Warning -Detail 'Collector emitted a warning; consult the console or transcript.' }
        }
        $entry.RowCount = $rows.Count
        $events = @($script:AADRRun.Events | Select-Object -Skip $eventStart | Where-Object Status -notin @('TokenRenewed','Recovered'))
        $entry.Status = If (@($events | Where-Object Status -eq 'PermissionDenied').Count -gt 0 -and $rows.Count -eq 0 -and @($events | Where-Object Status -eq 'Unavailable').Count -eq 0) { 'PermissionDenied' }
            ElseIf ($events.Count -gt 0 -or @($rows | Where-Object { $_.CollectionStatus -eq 'Unknown' }).Count -gt 0) { 'Partial' }
            ElseIf ($rows.Count -eq 0) { 'Empty' } Else { 'Success' }
        $entry.DataStatus = $entry.Status
        $unknownRows = @($rows | Where-Object { $_.CollectionStatus -eq 'Unknown' }).Count
        If ($unknownRows -gt 0) { Add-AADREvent -Operation $Name -Status Unavailable -Detail "$unknownRows record(s) retain an unknown collection status." }
        If ($entry.DataStatus -in @('Success','Empty'))
        {
            $temp = $null
            Try
            {
                $cacheDir = Join-Path $OutputDirectory '.checkpoints'
                $null = New-Item -ItemType Directory -Path $cacheDir -Force -ErrorAction Stop
                $cache = Join-Path $cacheDir ($Name + '.clixml')
                $temp = $cache + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
                Export-Clixml -InputObject $rows -LiteralPath $temp -Depth 20 -ErrorAction Stop
                Move-Item -LiteralPath $temp -Destination $cache -Force -ErrorAction Stop
                $entry.CheckpointHash = (Get-FileHash -LiteralPath $cache -Algorithm SHA256 -ErrorAction Stop).Hash
                $entry.Checkpoint = Join-AADRPath @('.checkpoints', ($Name + '.clixml'))
            }
            Catch
            {
                If (Test-AADRControlFlowError $_) { throw }
                $entry.Checkpoint = $null; $entry.CheckpointHash = $null; $entry.Status = 'Partial'
                Add-AADREvent -Operation Checkpoint -Status CheckpointFailed -Detail 'Collected data could not be checkpointed; exports will still be attempted. This dataset must be recollected on resume.' -Diagnostic (Get-AADRErrorDiagnostic $_)
                Write-Warning "[$Name] Checkpoint could not be saved. Report exports will still be attempted." -WarningAction Continue
            }
            Finally { If ($temp -and (Test-Path -LiteralPath $temp)) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue } }
        }
        If ($rows.Count -gt 0)
        {
            ForEach ($format in $Formats | Where-Object { $_ -ne 'EXCEL' })
            {
                $exportStart = $script:AADRRun.Events.Count
                Try
                {
                    Export-ADR -ADRObj $rows -AADROutputDir $OutputDirectory -OutputType @($format) -ADRModuleName $Name -ErrorAction Stop
                }
                Catch
                {
                    If (Test-AADRControlFlowError $_) { throw }
                    Add-AADREvent -Operation ($Name + ' ' + $format) -Status ExportFailed -Detail 'Export failed; remaining output formats will still be attempted.' -Diagnostic (Get-AADRErrorDiagnostic $_)
                    Write-Warning "[$Name] $format export failed. Remaining formats will continue." -WarningAction Continue
                }
                $failed = @($script:AADRRun.Events | Select-Object -Skip $exportStart | Where-Object Status -eq 'ExportFailed').Count -gt 0
                $entry.Exports += [pscustomobject]@{Format=$format; Status=$(If ($failed) {'Failed'} Else {'Success'}); Path=$(If ($format -eq 'STDOUT') { $null } Else { Join-AADRPath @(($format + '-Files'), ($Name + '.' + $format.ToLowerInvariant())) })}
                If ($failed) { $entry.Status = 'ExportFailed' }
            }
        }
        Else
        {
            ForEach ($format in $Formats | Where-Object { $_ -ne 'EXCEL' })
            { $entry.Exports += [pscustomobject]@{Format=$format; Status=$(If ($entry.DataStatus -eq 'Empty') { 'NotWrittenEmpty' } Else { 'Unavailable' }); Path=$null} }
        }
    }
    Catch
    {
        $diagnostic = Get-AADRErrorDiagnostic $_
        $entry.Status = If ($diagnostic.Category -eq 'PermissionDenied') { 'PermissionDenied' } Else { 'Failed' }
        If ($entry.DataStatus -eq 'Running') { $entry.DataStatus = $entry.Status }
        $entry.Error = $diagnostic.Category
        $entry.Diagnostic = $diagnostic
        Add-AADREvent -Operation $Name -Status $entry.Status -Detail 'Dataset did not complete; it must not be treated as empty or complete.' -Diagnostic $diagnostic
        If (Test-AADRControlFlowError $_) { throw }
        If (Test-AADRGraphAuthenticationError $_)
        {
            Add-AADREvent -Operation $Name -Status AuthenticationFailed -Detail 'Valid authentication could not be recovered. Collection stopped; completed checkpoints are retained.'
            Stop-AADRGraphAuthentication -Record $_
        }
        Write-Warning "[$Name] Collection failed ($($diagnostic.Category)). Remaining collectors will continue. $($diagnostic.Guidance)" -WarningAction Continue
    }
    Finally
    {
        $entry.FinishedUtc = [DateTime]::UtcNow.ToString('o')
        Try { Write-AADRManifest }
        Finally { $script:AADRCurrentCollector = $null }
    }
}

Function Get-AADRApplicationInventory
{
    [CmdletBinding()]
    param()
    $apps = @(Invoke-AADRGraphCollection -Name Applications -Query { Get-AADRGraphPages 'applications?$select=id,appId,displayName,createdDateTime,signInAudience,publisherDomain,requiredResourceAccess' })
    ForEach ($app in $apps)
    {
        [pscustomobject][ordered]@{
            Id=$app.id; AppId=$app.appId; DisplayName=$app.displayName; CreatedDateTime=$app.createdDateTime
            SignInAudience=$app.signInAudience; PublisherDomain=$app.publisherDomain
            RequestedResourceAccess=$(If ($null -ne $app.requiredResourceAccess) { ConvertTo-Json -InputObject $app.requiredResourceAccess -Depth 15 -Compress } Else { $null })
        }
    }
}

Function Get-AADRObjectOwners
{
    [CmdletBinding()]
    param([ValidateSet('applications','servicePrincipals')] [string] $Kind)
    $parents = @(Invoke-AADRGraphCollection -Name "$Kind owners parent list" -Query { Get-AADRGraphPages ($Kind + '?$select=id,appId,displayName') })
    ForEach ($parent in $parents)
    {
        $state = @{}
        $owners = @(Invoke-AADRGraphCollection -Name "$Kind owners" -State $state -Query { Get-AADRGraphPages ($Kind + '/' + [uri]::EscapeDataString($parent.id) + '/owners') })
        If (-Not $state.Available)
        {
            [pscustomobject]@{ObjectId=$parent.id; ObjectName=$parent.displayName; ObjectType=$Kind; OwnerId=$null; OwnerName=$null; OwnerPrincipalName=$null; OwnerType=$null; CollectionStatus='Unknown'}
        }
        ForEach ($owner in $owners)
        {
            [pscustomobject]@{ObjectId=$parent.id; ObjectName=$parent.displayName; ObjectType=$Kind; OwnerId=$owner.id; OwnerName=$owner.displayName; OwnerPrincipalName=$owner.userPrincipalName; OwnerType=(Get-AADRProperty $owner '@odata.type'); CollectionStatus='Collected'}
        }
    }
}

Function Get-AADRCredentialExpiry
{
    [CmdletBinding()]
    param([DateTime] $date = (Get-Date))
    ForEach ($kind in @('applications','servicePrincipals'))
    {
        $objects = @(Invoke-AADRGraphCollection -Name "$kind credential metadata" -Query { Get-AADRGraphPages ($kind + '?$select=id,appId,displayName,passwordCredentials,keyCredentials') })
        ForEach ($object in $objects)
        {
            ForEach ($property in @('passwordCredentials','keyCredentials'))
            {
                $credentials = Get-AADRProperty $object $property
                If ($null -eq $credentials)
                {
                    Add-AADREvent -Operation CredentialExpiry -Status Unavailable -Detail 'Credential metadata was omitted for an object.'
                    [pscustomobject]@{ObjectType=$kind; ObjectId=$object.id; AppId=$object.appId; ObjectName=$object.displayName; CredentialType=$property; KeyId=$null; DisplayName=$null; StartDateTime=$null; EndDateTime=$null; DaysToExpiry=$null; Expired=$null; CollectionStatus='Unknown'}
                }
                ForEach ($credential in $credentials)
                {
                    $end = If ($credential.endDateTime) { [DateTimeOffset]$credential.endDateTime } Else { $null }
                    [pscustomobject]@{
                        ObjectType=$kind; ObjectId=$object.id; AppId=$object.appId; ObjectName=$object.displayName
                        CredentialType=$property; KeyId=$credential.keyId; DisplayName=$credential.displayName
                        StartDateTime=$credential.startDateTime; EndDateTime=$credential.endDateTime
                        DaysToExpiry=$(If ($null -ne $end) { [Math]::Floor(($end.UtcDateTime - $date.ToUniversalTime()).TotalDays) } Else { $null })
                        Expired=$(If ($null -ne $end) { $end.UtcDateTime -lt $date.ToUniversalTime() } Else { $null })
                        CollectionStatus='Collected'
                    }
                }
            }
        }
    }
}

Function Get-AADRConsentGrants
{
    [CmdletBinding()]
    param()
    Invoke-AADRGraphCollection -Name ConsentGrants -Query { Get-AADRGraphPages 'oauth2PermissionGrants' } |
        Select-Object id, clientId, resourceId, principalId, consentType, scope
}

Function Get-AADRApplicationPermissions
{
    [CmdletBinding()]
    param()
    $principals = @(Invoke-AADRGraphCollection -Name 'ApplicationPermissions service principals' -Query { Get-AADRGraphPages 'servicePrincipals?$select=id,appId,displayName,appRoles' })
    $resourceMap = @{}
    ForEach ($resource in $principals) { $resourceMap[[string]$resource.id] = $resource }
    ForEach ($principal in $principals)
    {
        $state = @{}
        $assignments = @(Invoke-AADRGraphCollection -Name 'ApplicationPermissions assignments' -State $state -Query { Get-AADRGraphPages ('servicePrincipals/' + [uri]::EscapeDataString($principal.id) + '/appRoleAssignments') })
        If (-Not $state.Available)
        {
            [pscustomobject]@{ClientId=$principal.id; ClientAppId=$principal.appId; ClientName=$principal.displayName; AssignmentId=$null; ResourceId=$null; ResourceName=$null; AppRoleId=$null; PermissionValue=$null; PermissionDisplayName=$null; CollectionStatus='Unknown'}
        }
        ForEach ($assignment in $assignments)
        {
            # Assignment IDs describe granted application permissions, not exposed API scopes.
            $resource = $resourceMap[[string]$assignment.resourceId]
            $role = @($resource.appRoles | Where-Object { $_.id -eq $assignment.appRoleId } | Select-Object -First 1)
            $permissionValue = If ($role.Count -gt 0) { $role[0].value } Else { $null }
            $permissionName = If ($role.Count -gt 0) { $role[0].displayName } Else { $null }
            [pscustomobject]@{ClientId=$principal.id; ClientAppId=$principal.appId; ClientName=$principal.displayName; AssignmentId=$assignment.id; ResourceId=$assignment.resourceId; ResourceName=$assignment.resourceDisplayName; AppRoleId=$assignment.appRoleId; PermissionValue=$permissionValue; PermissionDisplayName=$permissionName; CollectionStatus='Collected'}
        }
    }
}

Function Get-AADRMFARegistration
{
    [CmdletBinding()]
    param()
    If (-Not $script:AADRMFARegistrationLoaded)
    {
        $state = @{}
        $script:AADRMFARegistrationRows = @(Invoke-AADRGraphCollection -Name MFARegistration -State $state -Query { Get-AADRGraphPages 'reports/authenticationMethods/userRegistrationDetails' })
        $script:AADRMFARegistrationAvailable = $state.Available
        $script:AADRMFARegistrationFailureStatus = $state.FailureStatus
        $script:AADRMFARegistrationLoaded = $true
    }
    If (-Not $script:AADRMFARegistrationAvailable)
    {
        $failureStatus = If ($script:AADRMFARegistrationFailureStatus) { $script:AADRMFARegistrationFailureStatus } Else { 'Unavailable' }
        Add-AADREvent -Operation MFARegistration -Status $failureStatus -Detail 'Registration/capability report is unavailable; registration remains unknown.'
    }
    $script:AADRMFARegistrationRows
}

Function Get-AADRGraphConnectionParameters
{
    [CmdletBinding()]
    param([hashtable] $Config, [System.Management.Automation.PSCredential] $Credential, [string[]] $Scopes)
    $connection = @{NoWelcome=$true; ContextScope='Process'; ErrorAction='Stop'}
    If ($Config.TenantID) { $connection.TenantId = $Config.TenantID }
    If ($Credential -and $Credential -ne [System.Management.Automation.PSCredential]::Empty)
    {
        $appId = [guid]::Empty
        If (-Not [guid]::TryParse($Credential.UserName, [ref]$appId))
        { throw '-Credential requires an application client ID (GUID) as username and client secret as password. User/password authentication is not supported by this parameter. Use interactive authentication for user sign-in.' }
        If (-Not $Config.TenantID) { throw 'TenantID is required for application credential authentication.' }
        $connection.ClientSecretCredential = $Credential
    }
    ElseIf ($Config.ClientSecret)
    {
        If (-Not $Config.ClientID -or -Not $Config.TenantID) { throw 'ClientID and TenantID are required with ClientSecret.' }
        $connection.ClientSecretCredential = [System.Management.Automation.PSCredential]::new($Config.ClientID, (ConvertTo-SecureString $Config.ClientSecret -AsPlainText -Force))
    }
    ElseIf ($Config.CertificateThumbprint -or $Config.CertificateName -or $Config.Certificate)
    {
        If (-Not $Config.ClientID -or -Not $Config.TenantID) { throw 'ClientID and TenantID are required for certificate authentication.' }
        $connection.ClientId = $Config.ClientID
        If ($Config.CertificateThumbprint) { $connection.CertificateThumbprint = $Config.CertificateThumbprint }
        ElseIf ($Config.CertificateName) { $connection.CertificateSubjectName = $Config.CertificateName }
        Else { $connection.Certificate = Get-Item -LiteralPath ('Cert:\LocalMachine\My\' + $Config.Certificate) -ErrorAction Stop }
    }
    Else
    {
        $connection.Scopes = $Scopes
        If ($Config.ClientID) { $connection.ClientId = $Config.ClientID }
    }
    $connection
}

Function Get-DateDiff
{
<#
.SYNOPSIS
    Get difference between two dates.

.DESCRIPTION
    Returns the difference between two dates.

.PARAMETER Date1
    [DateTime]
    Date

.PARAMETER Date2
    [DateTime]
    Date

.OUTPUTS
    [System.ValueType.TimeSpan]
    Returns the difference between the two dates.
#>
    param (
        [Parameter(Mandatory = $true)]
        [DateTime] $Date1,

        [Parameter(Mandatory = $true)]
        [DateTime] $Date2
    )

    If ($Date2 -gt $Date1)
    {
        $DDiff = $Date2 - $Date1
    }
    Else
    {
        $DDiff = $Date1 - $Date2
    }
    Return $DDiff
}

Function Invoke-AADRAtomicFileWrite
{
    [CmdletBinding()]
    param([string] $Path, [scriptblock] $Writer)
    $ErrorActionPreference = 'Stop'
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $temp = $fullPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    Try
    {
        If (Test-Path -LiteralPath $fullPath -PathType Container) { throw [System.IO.IOException]::new('The report file path is a directory.') }
        $null = & $Writer $temp
        Move-Item -LiteralPath $temp -Destination $fullPath -Force -ErrorAction Stop
    }
    Finally { If (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue } }
}

Function Export-ADRCSV
{
<#
.SYNOPSIS
    Exports Object to a CSV file.

.DESCRIPTION
    Exports Object to a CSV file using Export-CSV.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the CSV File.

.OUTPUTS
    CSV file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        Invoke-AADRAtomicFileWrite -Path $ADFileName -Writer {
            param($targetFile)
            $columns = New-Object 'System.Collections.Generic.List[string]'
            $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            $flatRows = @($ADRObj | ForEach-Object {
                $flat = [ordered]@{}
                ForEach ($property in $_.PSObject.Properties)
                {
                    $value = $property.Value
                    If ($null -ne $value -and $value -isnot [string] -and $value -isnot [ValueType])
                    { $value = ConvertTo-Json -InputObject $value -Depth 20 -Compress -WarningAction Stop -ErrorAction Stop }
                    $flat[$property.Name] = $value
                    If ($seen.Add($property.Name)) { $columns.Add($property.Name) }
                }
                [pscustomobject]$flat
            })
            # Unknown marker rows can have fewer fields than collected rows. Preserve
            # the union of columns regardless of which row happens to be first.
            $flatRows | Select-Object -Property $columns.ToArray() | Export-Csv -LiteralPath $targetFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        }
    }
    Catch
    {
        If (Test-AADRControlFlowError $_) { throw }
        Add-AADREvent -Operation 'Export-ADRCSV' -Status ExportFailed -Detail 'Output file could not be written.' -Diagnostic (Get-AADRErrorDiagnostic $_)
        Write-Warning "[Export-ADRCSV] Failed to export $($ADFileName)." -WarningAction Continue

    }
}

Function Export-ADRXML
{
<#
.SYNOPSIS
    Exports Object to a XML file.

.DESCRIPTION
    Exports Object to a XML file using Export-Clixml.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the XML File.

.OUTPUTS
    XML file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        Invoke-AADRAtomicFileWrite -Path $ADFileName -Writer {
            param($targetFile)
            (ConvertTo-Xml -NoTypeInformation -InputObject $ADRObj -Depth 20 -ErrorAction Stop).Save($targetFile)
        }
    }
    Catch
    {
        If (Test-AADRControlFlowError $_) { throw }
        Add-AADREvent -Operation 'Export-ADRXML' -Status ExportFailed -Detail 'Output file could not be written.' -Diagnostic (Get-AADRErrorDiagnostic $_)
        Write-Warning "[Export-ADRXML] Failed to export $($ADFileName)." -WarningAction Continue

    }
}

Function Export-ADRJSON
{
<#
.SYNOPSIS
    Exports Object to a JSON file.

.DESCRIPTION
    Exports Object to a JSON file using ConvertTo-Json.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the JSON File.

.OUTPUTS
    JSON file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        Invoke-AADRAtomicFileWrite -Path $ADFileName -Writer {
            param($targetFile)
            ConvertTo-JSON -InputObject $ADRObj -Depth 20 -WarningAction Stop -ErrorAction Stop | Out-File -LiteralPath $targetFile -Encoding UTF8 -ErrorAction Stop
        }
    }
    Catch
    {
        If (Test-AADRControlFlowError $_) { throw }
        Add-AADREvent -Operation 'Export-ADRJSON' -Status ExportFailed -Detail 'Output file could not be written.' -Diagnostic (Get-AADRErrorDiagnostic $_)
        Write-Warning "[Export-ADRJSON] Failed to export $($ADFileName)." -WarningAction Continue

    }
}

Function Export-ADRHTML
{
<#
.SYNOPSIS
    Exports Object to a HTML file.

.DESCRIPTION
    Exports Object to a HTML file using ConvertTo-Html.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the HTML File.

.OUTPUTS
    HTML file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName,

        [Parameter(Mandatory = $false)]
        [String] $AADROutputDir = $null
    )

$Header = @"
<style type="text/css">
th {
	color:white;
	background-color:blue;
	position: sticky;
	top: 0px;
}
td, th {
	border:0px solid black;
	border-collapse:collapse;
	white-space:pre;
}
tr:nth-child(2n+1) {
    background-color: #dddddd;
}
tr:hover td {
    background-color: #c1d5f8;
}
table, tr, td, th {
	padding: 0px;
	margin: 0px;
	white-space:pre;
}
table {
	margin-left:1px;
}
</style>
"@
    Try
    {
        Invoke-AADRAtomicFileWrite -Path $ADFileName -Writer {
            param($targetFile)
            If ([System.IO.Path]::GetFileName($ADFileName) -ieq 'Index.html')
            {
                $HTMLPath  = (Join-AADRPath @($AADROutputDir, 'HTML-Files'))
                $HTMLPath = (Resolve-Path -LiteralPath $HTMLPath -ErrorAction Stop).Path
                $HTMLFiles = Get-ChildItem -LiteralPath $HTMLPath -Name -ErrorAction Stop
                $HTML = $HTMLFiles | ConvertTo-HTML -Title "AzureADRecon" -Property @{Label="Table of Contents";Expression={"<a href='$($_)'>$($_)</a>"}} -Head $Header

                [System.Net.WebUtility]::HtmlDecode($HTML) | Out-File -LiteralPath $targetFile -Encoding UTF8 -ErrorAction Stop
            }
            Else
            {
                If ($ADRObj -is [array])
                {
                    $ADRObj | Select-Object * | ConvertTo-HTML -As Table -Head $Header | Out-File -LiteralPath $targetFile -Encoding UTF8 -ErrorAction Stop
                }
                Else
                {
                    ConvertTo-HTML -InputObject $ADRObj -As Table -Head $Header | Out-File -LiteralPath $targetFile -Encoding UTF8 -ErrorAction Stop
                }
            }
        }
    }
    Catch
    {
        If (Test-AADRControlFlowError $_) { throw }
        Add-AADREvent -Operation 'Export-ADRHTML' -Status ExportFailed -Detail 'Output file could not be written.' -Diagnostic (Get-AADRErrorDiagnostic $_)
        Write-Warning "[Export-ADRHTML] Failed to export $($ADFileName)." -WarningAction Continue

    }
}

Function Export-ADR
{
<#
.SYNOPSIS
    Helper function for all output types supported.

.DESCRIPTION
    Helper function for all output types supported.

.PARAMETER ADObjectDN
    [PSObject]
    ADRObj

.PARAMETER ADROutputDir
    [String]
    Path for AzureADRecon output folder.

.PARAMETER OutputType
    [array]
    Output Type.

.PARAMETER ADRModuleName
    [String]
    Module Name.

.OUTPUTS
    STDOUT, CSV, XML, JSON and/or HTML file, etc.
#>
    param(
        [Parameter(Mandatory = $true)]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [String] $AADROutputDir,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $true)]
        [String] $ADRModuleName
    )

    Switch ($OutputType)
    {
        'STDOUT'
        {
            If ($ADRModuleName -ne "AboutAzureADRecon")
            {
                If ($ADRObj -is [array])
                {
                    # Fix for InvalidOperationException: The object of type "Microsoft.PowerShell.Commands.Internal.Format.FormatStartData" is not valid or not in the correct sequence.
                    $ADRObj | Out-String -Stream
                }
                Else
                {
                    # Fix for InvalidOperationException: The object of type "Microsoft.PowerShell.Commands.Internal.Format.FormatStartData" is not valid or not in the correct sequence.
                    $ADRObj | Format-List | Out-String -Stream
                }
            }
        }
        'CSV'
        {
            $ADFileName  = (Join-AADRPath @($AADROutputDir, 'CSV-Files', ($ADRModuleName + '.csv')))
            Export-ADRCSV -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'XML'
        {
            $ADFileName  = (Join-AADRPath @($AADROutputDir, 'XML-Files', ($ADRModuleName + '.xml')))
            Export-ADRXML -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'JSON'
        {
            $ADFileName  = (Join-AADRPath @($AADROutputDir, 'JSON-Files', ($ADRModuleName + '.json')))
            Export-ADRJSON -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'HTML'
        {
            $ADFileName  = (Join-AADRPath @($AADROutputDir, 'HTML-Files', ($ADRModuleName + '.html')))
            Export-ADRHTML -ADRObj $ADRObj -ADFileName $ADFileName -AADROutputDir $AADROutputDir
        }
    }
}

Function Get-ADRExcelComObj
{
<#
.SYNOPSIS
    Creates a ComObject to interact with Microsoft Excel.

.DESCRIPTION
    Creates a ComObject to interact with Microsoft Excel if installed, else warning is raised.

.OUTPUTS
    [System.__ComObject] and [System.MarshalByRefObject]
    Creates global variables $excel and $workbook.
#>

    #Check if Excel is installed.
    Try
    {
        # Suppress verbose output
        $SaveVerbosePreference = $script:VerbosePreference
        $script:VerbosePreference = 'SilentlyContinue'
        $global:excel = New-Object -ComObject excel.application
        If ($SaveVerbosePreference)
        {
            $script:VerbosePreference = $SaveVerbosePreference
            Remove-Variable SaveVerbosePreference -ErrorAction SilentlyContinue
        }
    }
    Catch
    {
        If ($SaveVerbosePreference)
        {
            $script:VerbosePreference = $SaveVerbosePreference
            Remove-Variable SaveVerbosePreference -ErrorAction SilentlyContinue
        }
        Write-Warning "[Get-ADRExcelComObj] Excel does not appear to be installed. Skipping generation of AzureADRecon-Report.xlsx. Use the -GenExcel parameter to generate the AzureADRecon-Report.xslx on a host with Microsoft Excel installed."
        Write-Verbose "[EXCEPTION] operation unavailable"
        Return $null
    }
    $excel.Visible = $true
    $excel.Interactive = $false
    $global:workbook = $excel.Workbooks.Add()
    If ($workbook.Worksheets.Count -eq 3)
    {
        $workbook.WorkSheets.Item(3).Delete()
        $workbook.WorkSheets.Item(2).Delete()
    }
}

Function Get-ADRExcelComObjRelease
{
<#
.SYNOPSIS
    Releases the ComObject created to interact with Microsoft Excel.

.DESCRIPTION
    Releases the ComObject created to interact with Microsoft Excel.

.PARAMETER ComObjtoRelease
    ComObjtoRelease

.PARAMETER Final
    Final
#>
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        $ComObjtoRelease,

        [Parameter(Mandatory = $false)]
        [bool] $Final = $false
    )
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.releasecomobject(v=vs.110).aspx
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.finalreleasecomobject(v=vs.110).aspx
    If ($null -eq $ComObjtoRelease -or -Not [System.Runtime.InteropServices.Marshal]::IsComObject($ComObjtoRelease)) { return }
    Try
    {
        If ($Final) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObjtoRelease) | Out-Null }
        Else { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObjtoRelease) | Out-Null }
    }
    Catch
    {
        If (Test-AADRControlFlowError $_) { throw }
        $script:AADRExcelFailed = $true
        Add-AADREvent -Operation 'Excel cleanup' -Status CleanupFailed -Detail 'A COM reference could not be released.' -Diagnostic (Get-AADRErrorDiagnostic $_)
        Write-Warning '[Excel cleanup] A COM reference could not be released.' -WarningAction Continue
    }
}

Function Get-ADRExcelWorkbook
{
<#
.SYNOPSIS
    Adds a WorkSheet to the Workbook.

.DESCRIPTION
    Adds a WorkSheet to the Workbook using the $workboook global variable and assigns it a name.

.PARAMETER name
    [string]
    Name of the WorkSheet.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $name
    )

    $workbook.Worksheets.Add() | Out-Null
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = $name

    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet -ErrorAction SilentlyContinue
}

Function Get-ADRExcelImport
{
<#
.SYNOPSIS
    Helper to import CSV to the current WorkSheet.

.DESCRIPTION
    Helper to import CSV to the current WorkSheet. Supports two methods.

.PARAMETER ADFileName
    [string]
    Filename of the CSV file to import.

.PARAMETER method
    [int]
    Method to use for the import.

.PARAMETER row
    [int]
    Row.

.PARAMETER column
    [int]
    Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ADFileName,

        [Parameter(Mandatory = $false)]
        [int] $Method = 1,

        [Parameter(Mandatory = $false)]
        [int] $row = 1,

        [Parameter(Mandatory = $false)]
        [int] $column = 1
    )

    $excel.ScreenUpdating = $false
    If ($Method -eq 1)
    {
        If (Test-Path $ADFileName)
        {
            $worksheet = $workbook.Worksheets.Item(1)
            $TxtConnector = ("TEXT;" + $ADFileName)
            $CellRef = $worksheet.Range("A1")
            #Build, use and remove the text file connector
            $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)

            #65001: Unicode (UTF-8)
            $worksheet.QueryTables.item($Connector.name).TextFilePlatform = 65001
            $worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
            $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
            $worksheet.QueryTables.item($Connector.name).Refresh() | Out-Null
            $worksheet.QueryTables.item($Connector.name).delete()

            Get-ADRExcelComObjRelease -ComObjtoRelease $CellRef
            Remove-Variable CellRef -ErrorAction SilentlyContinue
            Get-ADRExcelComObjRelease -ComObjtoRelease $Connector
            Remove-Variable Connector -ErrorAction SilentlyContinue

            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        }
        Remove-Variable ADFileName -ErrorAction SilentlyContinue
    }
    Elseif ($Method -eq 2)
    {
        $worksheet = $workbook.Worksheets.Item(1)
        If (Test-Path $ADFileName)
        {
            $ADTemp = Import-Csv -Path $ADFileName
            $ADTemp | ForEach-Object {
                Foreach ($prop in $_.PSObject.Properties)
                {
                    $worksheet.Cells.Item($row, $column) = $prop.Name
                    $worksheet.Cells.Item($row, $column + 1) = $prop.Value
                    $row++
                }
            }
            Remove-Variable ADTemp -ErrorAction SilentlyContinue
            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
        }
        Else
        {
            $worksheet.Cells.Item($row, $column) = "Error!"
        }
        Remove-Variable ADFileName -ErrorAction SilentlyContinue
    }
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet -ErrorAction SilentlyContinue
}

# Thanks Anant Shrivastava for the suggestion of using Pivot Tables for generation of the Stats sheets.
Function Get-ADRExcelPivotTable
{
<#
.SYNOPSIS
    Helper to add Pivot Table to the current WorkSheet.

.DESCRIPTION
    Helper to add Pivot Table to the current WorkSheet.

.PARAMETER SrcSheetName
    [string]
    Source Sheet Name.

.PARAMETER PivotTableName
    [string]
    Pivot Table Name.

.PARAMETER PivotRows
    [array]
    Row names from Source Sheet.

.PARAMETER PivotColumns
    [array]
    Column names from Source Sheet.

.PARAMETER PivotFilters
    [array]
    Row/Column names from Source Sheet to use as filters.

.PARAMETER PivotValues
    [array]
    Row/Column names from Source Sheet to use for Values.

.PARAMETER PivotPercentage
    [array]
    Row/Column names from Source Sheet to use for Percentage.

.PARAMETER PivotLocation
    [array]
    Location of the Pivot Table in Row/Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $SrcSheetName,

        [Parameter(Mandatory = $true)]
        [string] $PivotTableName,

        [Parameter(Mandatory = $false)]
        [array] $PivotRows,

        [Parameter(Mandatory = $false)]
        [array] $PivotColumns,

        [Parameter(Mandatory = $false)]
        [array] $PivotFilters,

        [Parameter(Mandatory = $false)]
        [array] $PivotValues,

        [Parameter(Mandatory = $false)]
        [array] $PivotPercentage,

        [Parameter(Mandatory = $false)]
        [string] $PivotLocation = "R1C1"
    )

    $excel.ScreenUpdating = $false
    $SrcWorksheet = $workbook.Sheets.Item($SrcSheetName)
    $workbook.ShowPivotTableFieldList = $false

    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottablesourcetype-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottableversionlist-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfieldorientation-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/constants-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfiltertype-enumeration-excel

    # xlDatabase = 1 # this just means local sheet data
    # xlPivotTableVersion12 = 3 # Excel 2007
    $PivotFailed = $false
    Try
    {
        $PivotCaches = $workbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, $SrcWorksheet.UsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
    }
    Catch
    {
        $PivotFailed = $true
        Write-Verbose "[PivotCaches().Create] Failed"
        Write-Verbose "[EXCEPTION] operation unavailable"
    }
    If ( $PivotFailed -eq $true )
    {
        $rows = $SrcWorksheet.UsedRange.Rows.Count
        If ($SrcSheetName -eq "Users")
        {
            $PivotCols = "A1:C"
        }
        $UsedRange = $SrcWorksheet.Range($PivotCols+$rows)
        $PivotCaches = $workbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, $UsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
        Remove-Variable rows -ErrorAction SilentlyContinue
	    Remove-Variable PivotCols -ErrorAction SilentlyContinue
        Remove-Variable UsedRange -ErrorAction SilentlyContinue
    }
    Remove-Variable PivotFailed -ErrorAction SilentlyContinue
    $PivotTable = $PivotCaches.CreatePivotTable($PivotLocation,$PivotTableName)
    # $workbook.ShowPivotTableFieldList = $true

    If ($PivotRows)
    {
        ForEach ($Row in $PivotRows)
        {
            $PivotField = $PivotTable.PivotFields($Row)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
        }
    }

    If ($PivotColumns)
    {
        ForEach ($Col in $PivotColumns)
        {
            $PivotField = $PivotTable.PivotFields($Col)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
        }
    }

    If ($PivotFilters)
    {
        ForEach ($Fil in $PivotFilters)
        {
            $PivotField = $PivotTable.PivotFields($Fil)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
        }
    }

    If ($PivotValues)
    {
        ForEach ($Val in $PivotValues)
        {
            $PivotField = $PivotTable.PivotFields($Val)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
        }
    }

    If ($PivotPercentage)
    {
        ForEach ($Val in $PivotPercentage)
        {
            $PivotField = $PivotTable.PivotFields($Val)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
            $PivotField.Calculation = [Microsoft.Office.Interop.Excel.XlPivotFieldCalculation]::xlPercentOfTotal
            $PivotTable.ShowValuesRow = $false
        }
    }

    # $PivotFields.Caption = ""
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotField
    Remove-Variable PivotField -ErrorAction SilentlyContinue
    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotTable
    Remove-Variable PivotTable -ErrorAction SilentlyContinue
    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotCaches
    Remove-Variable PivotCaches -ErrorAction SilentlyContinue
    Get-ADRExcelComObjRelease -ComObjtoRelease $SrcWorksheet
    Remove-Variable SrcWorksheet -ErrorAction SilentlyContinue
}

Function Get-ADRExcelAttributeStats
{
<#
.SYNOPSIS
    Helper to add Attribute Stats to the current WorkSheet.

.DESCRIPTION
    Helper to add Attribute Stats to the current WorkSheet.

.PARAMETER SrcSheetName
    [string]
    Source Sheet Name.

.PARAMETER Title1
    [string]
    Title1.

.PARAMETER PivotTableName
    [string]
    PivotTableName.

.PARAMETER PivotRows
    [string]
    PivotRows.

.PARAMETER PivotValues
    [string]
    PivotValues.

.PARAMETER PivotPercentage
    [string]
    PivotPercentage.

.PARAMETER Title2
    [string]
    Title2.

.PARAMETER ObjAttributes
    [OrderedDictionary]
    Attributes.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $SrcSheetName,

        [Parameter(Mandatory = $true)]
        [string] $Title1,

        [Parameter(Mandatory = $true)]
        [string] $PivotTableName,

        [Parameter(Mandatory = $true)]
        [string] $PivotRows,

        [Parameter(Mandatory = $true)]
        [string] $PivotValues,

        [Parameter(Mandatory = $true)]
        [string] $PivotPercentage,

        [Parameter(Mandatory = $true)]
        [string] $Title2,

        [Parameter(Mandatory = $true)]
        [System.Object] $ObjAttributes
    )

    $excel.ScreenUpdating = $false
    $worksheet = $workbook.Worksheets.Item(1)
    $SrcWorksheet = $workbook.Sheets.Item($SrcSheetName)

    $row = 1
    $column = 1
    $worksheet.Cells.Item($row, $column) = $Title1
    $worksheet.Cells.Item($row,$column).Style = "Heading 2"
    $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
    $MergeCells = $worksheet.Range("A1:C1")
    $MergeCells.Select() | Out-Null
    $MergeCells.MergeCells = $true
    Remove-Variable MergeCells -ErrorAction SilentlyContinue

    Get-ADRExcelPivotTable -SrcSheetName $SrcSheetName -PivotTableName $PivotTableName -PivotRows @($PivotRows) -PivotValues @($PivotValues) -PivotPercentage @($PivotPercentage) -PivotLocation "R2C1"
    $excel.ScreenUpdating = $false

    $row = 2
    "Type","Count","Percentage" | ForEach-Object {
        $worksheet.Cells.Item($row, $column) = $_
        $worksheet.Cells.Item($row, $column).Font.Bold = $True
        $column++
    }

    $row = 3
    $column = 1
    For($row = 3; $row -le 6; $row++)
    {
        $temptext = [string] $worksheet.Cells.Item($row, $column).Text
        switch ($temptext.ToUpper())
        {
            "TRUE" { $worksheet.Cells.Item($row, $column) = "Enabled" }
            "FALSE" { $worksheet.Cells.Item($row, $column) = "Disabled" }
            "(BLANK)" { $worksheet.Cells.Item($row, $column) = "Unknown" }
            "GRAND TOTAL" { $worksheet.Cells.Item($row, $column) = "Total" }
        }
    }

    If ($ObjAttributes)
    {
        $row = 1
        $column = 6
        $worksheet.Cells.Item($row, $column) = $Title2
        $worksheet.Cells.Item($row,$column).Style = "Heading 2"
        $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
        $MergeCells = $worksheet.Range("F1:L1")
        $MergeCells.Select() | Out-Null
        $MergeCells.MergeCells = $true
        Remove-Variable MergeCells -ErrorAction SilentlyContinue

        $row++
        "Category","Enabled Count","Enabled Percentage","Disabled Count","Disabled Percentage","Total Count","Total Percentage" | ForEach-Object {
            $worksheet.Cells.Item($row, $column) = $_
            $worksheet.Cells.Item($row, $column).Font.Bold = $True
            $column++
        }
        $ExcelColumn = ($SrcWorksheet.Columns.Find("Enabled"))
        If (-Not $ExcelColumn)
        {
            Write-Warning '[Get-ADRExcelAttributeStats] Enabled column is unavailable; account-state statistics were skipped.'
            return
        }
        $EnabledColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"
        $column = 6
        $i = 2

        $ObjAttributes.keys | ForEach-Object {
            $attribute = If ($_ -eq 'MFAUnknown') { 'MFAStatus' } Else { $_ }
            $ExcelColumn = ($SrcWorksheet.Columns.Find($attribute))
            If (-Not $ExcelColumn)
            {
                Write-Warning "[Get-ADRExcelAttributeStats] Column $attribute is unavailable; its statistics were skipped."
                return
            }
            $ColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"
            $row++
            $i++
            If ($_ -eq "UserType")
            {
                $worksheet.Cells.Item($row, $column) = "Guest"
            }
            ElseIf ($_ -eq "MFAStat")
            {
                $worksheet.Cells.Item($row, $column) = "MFA (Not registered)"
            }
            ElseIf ($_ -eq "MFAStatu")
            {
                $worksheet.Cells.Item($row, $column) = "MFA (Registered)"
            }
            ElseIf ($_ -eq "MFAUnknown")
            {
                $worksheet.Cells.Item($row, $column) = 'MFA (Unknown)'
            }
            ElseIf ($_ -eq "MFA Capable")
            {
                $worksheet.Cells.Item($row, $column) = "MFA (Capable)"
            }
            Else
            {
                $worksheet.Cells.Item($row, $column).Formula = "='" + $SrcWorksheet.Name + "'!" + $ExcelColumn.Address($false, $false)
            }

            If ($_ -eq "PasswordPolicies")
            {
                # Remove count of "None"
                $worksheet.Cells.Item($row, $column+1).Formula = "=COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"TRUE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')' + "-COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"TRUE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + '"None"' + ')'
            }
            Else
            {
                $worksheet.Cells.Item($row, $column+1).Formula = "=COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"TRUE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            }
            If ($_ -eq "PasswordPolicies")
            {
                # Remove count of "None"
                $worksheet.Cells.Item($row, $column+3).Formula = "=COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"FALSE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')' + "-COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"FALSE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + '"None"' + ')'
            }
            Else
            {
                $worksheet.Cells.Item($row, $column+3).Formula = "=COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"FALSE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            }
            If ($_ -eq "AssignedLicenses")
            {
                # Remove count of FieldName
                $worksheet.Cells.Item($row, $column+5).Formula = "=COUNTIF('" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')-1'
            }
            ElseIf ($_ -eq "PasswordPolicies")
            {
                # Remove count of "None" and FieldName
                $worksheet.Cells.Item($row, $column+5).Formula = "=COUNTIF('" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')' + "-COUNTIF('" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + '"None"' + ')-1'
            }
            Else
            {
                $worksheet.Cells.Item($row, $column+5).Formula = "=COUNTIF('" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            }
            # Show blank counts where no relevant field values were collected.
            $fieldRange = "'" + $SrcWorksheet.Name + "'!" + $ColAddress
            $enabledRange = "'" + $SrcWorksheet.Name + "'!" + $EnabledColAddress
            ForEach ($offset in @(1, 3, 5))
            {
                $formula = [string]$worksheet.Cells.Item($row, $column+$offset).Formula
                $unavailable = If ($offset -eq 1) { 'COUNTIFS(' + $enabledRange + ',"TRUE",' + $fieldRange + ',"<>")=0' }
                    ElseIf ($offset -eq 3) { 'COUNTIFS(' + $enabledRange + ',"FALSE",' + $fieldRange + ',"<>")=0' }
                    Else { 'COUNTIF(' + $fieldRange + ',"<>")<=1' }
                $worksheet.Cells.Item($row, $column+$offset).Formula = '=IF(' + $unavailable + ',"",' + $formula.TrimStart('=') + ')'
            }
            $worksheet.Cells.Item($row, $column+2).Formula = '=IF(G' + $i + '="","",IFERROR(G' + $i + '/COUNTIFS(' + $enabledRange + ',"TRUE",' + $fieldRange + ',"<>"),""))'
            $worksheet.Cells.Item($row, $column+4).Formula = '=IF(I' + $i + '="","",IFERROR(I' + $i + '/COUNTIFS(' + $enabledRange + ',"FALSE",' + $fieldRange + ',"<>"),""))'
            $worksheet.Cells.Item($row, $column+6).Formula = '=IF(K' + $i + '="","",IFERROR(K' + $i + '/(COUNTIF(' + $fieldRange + ',"<>")-1),""))'

        }

        # http://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/
        "H", "J" , "L" | ForEach-Object {
            $rng = $_ + $($row - $ObjAttributes.Count + 1) + ":" + $_ + $($row)
            $worksheet.Range($rng).NumberFormat = "0.00%"
        }
    }
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $SrcWorksheet
    Remove-Variable SrcWorksheet -ErrorAction SilentlyContinue
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet -ErrorAction SilentlyContinue
}

Function Get-ADRExcelChart
{
<#
.SYNOPSIS
    Helper to add charts to the current WorkSheet.

.DESCRIPTION
    Helper to add charts to the current WorkSheet.

.PARAMETER ChartType
    [int]
    Chart Type.

.PARAMETER ChartLayout
    [int]
    Chart Layout.

.PARAMETER ChartTitle
    [string]
    Title of the Chart.

.PARAMETER RangetoCover
    WorkSheet Range to be covered by the Chart.

.PARAMETER ChartData
    Data for the Chart.

.PARAMETER StartRow
    Start row to calculate data for the Chart.

.PARAMETER StartColumn
    Start column to calculate data for the Chart.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ChartType,

        [Parameter(Mandatory = $true)]
        [int] $ChartLayout,

        [Parameter(Mandatory = $true)]
        [string] $ChartTitle,

        [Parameter(Mandatory = $true)]
        $RangetoCover,

        [Parameter(Mandatory = $false)]
        $ChartData = $null,

        [Parameter(Mandatory = $false)]
        $StartRow = $null,

        [Parameter(Mandatory = $false)]
        $StartColumn = $null
    )

    $excel.ScreenUpdating = $false
    $excel.DisplayAlerts = $false
    $worksheet = $workbook.Worksheets.Item(1)
    $chart = $worksheet.Shapes.AddChart().Chart
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlcharttype-enumeration-excel
    $chart.chartType = [int]([Microsoft.Office.Interop.Excel.XLChartType]::$ChartType)
    $chart.ApplyLayout($ChartLayout)
    If ($null -eq $ChartData)
    {
        If ($null -eq $StartRow)
        {
            $start = $worksheet.Range("A1")
        }
        Else
        {
            $start = $worksheet.Range($StartRow)
        }
        # get the last cell
        $X = $worksheet.Range($start,$start.End([Microsoft.Office.Interop.Excel.XLDirection]::xlDown))
        If ($null -eq $StartColumn)
        {
            $start = $worksheet.Range("B1")
        }
        Else
        {
            $start = $worksheet.Range($StartColumn)
        }
        # get the last cell
        $Y = $worksheet.Range($start,$start.End([Microsoft.Office.Interop.Excel.XLDirection]::xlDown))
        $ChartData = $worksheet.Range($X,$Y)

        Get-ADRExcelComObjRelease -ComObjtoRelease $X
        Remove-Variable X -ErrorAction SilentlyContinue
        Get-ADRExcelComObjRelease -ComObjtoRelease $Y
        Remove-Variable Y -ErrorAction SilentlyContinue
        Get-ADRExcelComObjRelease -ComObjtoRelease $start
        Remove-Variable start -ErrorAction SilentlyContinue
    }
    $chart.SetSourceData($ChartData)
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.chartclass.plotby?redirectedfrom=MSDN&view=excel-pia#Microsoft_Office_Interop_Excel_ChartClass_PlotBy
    $chart.PlotBy = [Microsoft.Office.Interop.Excel.XlRowCol]::xlColumns
    $chart.seriesCollection(1).Select() | Out-Null
    $chart.SeriesCollection(1).ApplyDataLabels() | out-Null
    # modify the chart title
    $chart.HasTitle = $True
    $chart.ChartTitle.Text = $ChartTitle
    # Reposition the Chart
    $temp = $worksheet.Range($RangetoCover)
    # $chart.parent.placement = 3
    $chart.parent.top = $temp.Top
    $chart.parent.left = $temp.Left
    $chart.parent.width = $temp.Width
    # $chart.Legend.Delete()
    $excel.ScreenUpdating = $true
    $excel.DisplayAlerts = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $chart
    Remove-Variable chart -ErrorAction SilentlyContinue
    Get-ADRExcelComObjRelease -ComObjtoRelease $ChartData
    Remove-Variable ChartData -ErrorAction SilentlyContinue
    Get-ADRExcelComObjRelease -ComObjtoRelease $temp
    Remove-Variable temp -ErrorAction SilentlyContinue
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet -ErrorAction SilentlyContinue
}

Function Get-ADRExcelSort
{
<#
.SYNOPSIS
    Sorts a WorkSheet in the active Workbook.

.DESCRIPTION
    Sorts a WorkSheet in the active Workbook.

.PARAMETER ColumnName
    [string]
    Name of the Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ColumnName
    )

    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Activate();

    $ExcelColumn = ($worksheet.Columns.Find($ColumnName))
    If ($ExcelColumn)
    {
        If ($ExcelColumn.Text -ne $ColumnName)
        {
            $BeginAddress = $ExcelColumn.Address(0,0,1,1)
            $End = $False
            Do {
                #Write-Verbose "[Get-ADRExcelSort] $($ExcelColumn.Text) selected instead of $($ColumnName) in the $($worksheet.Name) worksheet."
                $ExcelColumn = ($worksheet.Columns.FindNext($ExcelColumn))
                $Address = $ExcelColumn.Address(0,0,1,1)
                If ( ($Address -eq $BeginAddress) -or ($ExcelColumn.Text -eq $ColumnName) )
                {
                    $End = $True
                }
            } Until ($End -eq $True)
        }
        If ($ExcelColumn.Text -eq $ColumnName)
        {
            # Sort by Column
            $workSheet.ListObjects.Item(1).Sort.SortFields.Clear()
            $workSheet.ListObjects.Item(1).Sort.SortFields.Add($ExcelColumn) | Out-Null
            $worksheet.ListObjects.Item(1).Sort.Apply()
        }
        Else
        {
            Write-Verbose "[Get-ADRExcelSort] $($ColumnName) not found in the $($worksheet.Name) worksheet."
        }
    }
    Else
    {
        Write-Verbose "[Get-ADRExcelSort] $($ColumnName) not found in the $($worksheet.Name) worksheet."
    }
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet -ErrorAction SilentlyContinue
}

Function Export-ADRExcel
{
    [CmdletBinding()]
    param([string] $ExcelPath, [string] $Logo = 'AzureADRecon')
    $script:AADRExcelFailed = $false
    If ($env:OS -ne 'Windows_NT')
    {
        $script:AADRExcelFailed = $true
        Add-AADREvent -Operation Excel -Status ExportFailed -Detail 'Desktop Excel requires Windows.'
        Write-Warning 'Desktop Excel generation requires Windows. CSV reports remain available.'
        return
    }
    $ErrorActionPreference = 'Stop'
    $global:excel = $null
    $global:workbook = $null
    Try
    {
        $warnings = @()
        Export-AADRExcelCore -ExcelPath $ExcelPath -Logo $Logo -WarningVariable +warnings
        If ($warnings.Count -gt 0) { $script:AADRExcelFailed=$true; Add-AADREvent -Operation Excel -Status ExportFailed -Detail 'Excel generation emitted a warning; verify the workbook before use.' }
    }
    Catch
    {
        If (Test-AADRControlFlowError $_) { throw }
        $script:AADRExcelFailed = $true
        Add-AADREvent -Operation Excel -Status ExportFailed -Detail 'Desktop Excel generation failed.' -Diagnostic (Get-AADRErrorDiagnostic $_)
        Write-Warning '[Export-ADRExcel] Workbook generation failed. CSV reports remain available; check Excel installation and workbook access.' -WarningAction Continue
    }
    Finally
    {
        If ($global:excel)
        {
            Try { $global:excel.Quit() }
            Catch { $script:AADRExcelFailed=$true; Add-AADREvent -Operation 'Excel cleanup' -Status CleanupFailed -Detail 'Excel could not be closed cleanly.' -Diagnostic (Get-AADRErrorDiagnostic $_) }
        }
        ForEach ($name in @('workbook','excel'))
        {
            $object = Get-Variable -Name $name -Scope Global -ValueOnly -ErrorAction SilentlyContinue
            Get-ADRExcelComObjRelease -ComObjtoRelease $object -Final $true
            Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
        }
    }
}

Function Export-AADRExcelCore
{
<#
.SYNOPSIS
    Automates the generation of the AzureADRecon report.

.DESCRIPTION
    Automates the generation of the AzureADRecon report. If specific files exist, they are imported into the AzureADRecon report.

.PARAMETER ExcelPath
    [string]
    Path for AzureADRecon output folder containing the CSV files to generate the AzureADRecon-Report.xlsx

.PARAMETER Logo
    [string]
    Which Logo to use in the excel file? (Default AzureADRecon)

.OUTPUTS
    Creates the AzureADRecon-Report.xlsx report in the folder.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $ExcelPath,

        [Parameter(Mandatory = $false)]
        [string] $Logo = "AzureADRecon"
    )

    If ($PSVersionTable.PSEdition -eq "Core")
    {
        If ($env:OS -eq "Windows_NT")
        {
            $returndir = Get-Location
            Try
            {
                Set-Location 'C:\Windows\assembly\' -ErrorAction Stop
                $refFolder = Get-ChildItem -Recurse -Filter Microsoft.Office.Interop.Excel.dll -ErrorAction Stop | Select-Object -First 1 -ExpandProperty DirectoryName
                If ($refFolder)
                {
                    Set-Location -LiteralPath $refFolder -ErrorAction Stop
                    Add-Type -AssemblyName "Microsoft.Office.Interop.Excel" -ErrorAction Stop
                }
            }
            Finally { Set-Location -LiteralPath $returndir -ErrorAction SilentlyContinue }
            Remove-Variable returndir -ErrorAction SilentlyContinue
            Remove-Variable refFolder -ErrorAction SilentlyContinue
        }
    }

    $ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path
    $ReportPath = (Join-AADRPath @($ExcelPath, 'CSV-Files'))
    If (!(Test-Path $ReportPath))
    {
        Write-Warning "[Export-ADRExcel] Could not locate the CSV-Files directory ... Exiting"
        Write-Verbose "[EXCEPTION] operation unavailable"
        Return $null
    }
    Get-ADRExcelComObj
    If ($excel)
    {
        Write-Output "[*] Generating AzureADRecon-Report.xlsx"

        $ADFileName = (Join-AADRPath @($ReportPath, 'AboutAzureADRecon.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelImport -ADFileName $ADFileName

            $Method = ((Import-CSV -Path $ADFileName) | Where-Object Category -eq 'MSGraph' | Select-Object -First 1).Category

            $workbook.Worksheets.Item(1).Name = "About AzureADRecon"
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(3,2) , "https://github.com/adrecon/AzureADRecon", "" , "", "github.com/adrecon/AzureADRecon") | Out-Null
            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null

            Remove-Variable ADFileName -ErrorAction SilentlyContinue
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'Tenant.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Tenant"
            Get-ADRExcelImport -ADFileName $ADFileName
            $TenantObj = Import-CSV -Path $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue
            $TenantName = -join($TenantObj[0].Value,"-")
            Remove-Variable TenantObj -ErrorAction SilentlyContinue
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'Domain.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Domain"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'ConditionalAccessPolicies.csv'))
        If (Test-Path $ADFileName) {
            Get-ADRExcelWorkbook -Name "ConditionalAccessPolicies"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'Licenses.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Licenses"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'Devices.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Devices"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'GroupMembers.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "GroupMembers"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue

            Get-ADRExcelSort -ColumnName "GroupName"
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'Groups.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Groups"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue

            Get-ADRExcelSort -ColumnName "DisplayName"
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'DirectoryRoleMembers.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "DirectoryRoleMembers"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue

            Get-ADRExcelSort -ColumnName "DirectoryRole"
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'DirectoryRoles.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "DirectoryRoles"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'ServicePrincipals.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "ServicePrincipals"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue
        }

        $ADFileName = (Join-AADRPath @($ReportPath, 'Users.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Users"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName -ErrorAction SilentlyContinue

            Get-ADRExcelSort -ColumnName "UserPrincipalName"

            $worksheet = $workbook.Worksheets.Item(1)

            # Freeze First Row and Column
            $worksheet.Select()
            $worksheet.Application.ActiveWindow.splitcolumn = 1
            $worksheet.Application.ActiveWindow.splitrow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true

            $worksheet.Cells.Item(1,3).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,3).font.ColorIndex = 2
            # Set Filter to Enabled Accounts only
            $worksheet.UsedRange.Select() | Out-Null
            $excel.Selection.AutoFilter(3,$true) | Out-Null
            $worksheet.Cells.Item(1,1).Select() | Out-Null
            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet -ErrorAction SilentlyContinue
        }

        # User Stats
        $ADFileName = (Join-AADRPath @($ReportPath, 'Users.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "User Stats"
            Remove-Variable ADFileName -ErrorAction SilentlyContinue

            $ObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            $ObjAttributes.Add("UserType", '"Guest"')

            If ($Method -eq "MSGraph")
            {
                $ObjAttributes.Add("Password Age (>", '"TRUE"')
                $ObjAttributes.Add("Never Logged In", '"TRUE"')
                $ObjAttributes.Add("Dormant (>", '"TRUE"')
                $ObjAttributes.Add("MFAStat", '"NotRegistered"')
                $ObjAttributes.Add("MFAStatu", '"Registered"')
                $ObjAttributes.Add("MFA Capable", '"TRUE"')
                $ObjAttributes.Add("MFAUnknown", '"Unknown"')
            }

            Get-ADRExcelAttributeStats -SrcSheetName "Users" -Title1 "User Accounts in AzureAD" -PivotTableName "User Accounts Status" -PivotRows "Enabled" -PivotValues "UserPrincipalName" -PivotPercentage "UserPrincipalName"  -Title2 "Status of User Accounts in AzureAD" -ObjAttributes $ObjAttributes

            $totalCell = $workbook.Worksheets.Item(1).Columns.Item(1).Find('Total')
            $summaryEndRow = If ($totalCell) { [Math]::Max(3, $totalCell.Row - 1) } Else { 4 }
            Get-ADRExcelChart -ChartType "xlPie" -ChartLayout 3 -ChartTitle "User Accounts in AzureAD" -RangetoCover "A12:D24" -ChartData $workbook.Worksheets.Item(1).Range("A3:A$summaryEndRow,B3:B$summaryEndRow")
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(11,1) , "" , "Users!A1", "", "Raw Data") | Out-Null

            Get-ADRExcelChart -ChartType "xlBarClustered" -ChartLayout 1 -ChartTitle "Status of User Accounts in AzureAD" -RangetoCover "F12:L24" -ChartData $workbook.Worksheets.Item(1).Range("F2:F10,G2:G10")
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(11,6) , "" , "Users!A1", "", "Raw Data") | Out-Null

            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
        }

        # Create Table of Contents
        Get-ADRExcelWorkbook -Name "Table of Contents"
        $worksheet = $workbook.Worksheets.Item(1)

        $excel.ScreenUpdating = $false
        # Image format and properties
        # $path = "C:\AzureADRecon_Logo.jpg"
        # $base64aadrecon = [convert]::ToBase64String((Get-Content $path -Encoding byte))

        If ($Logo -eq "Payatu")
        {
            $base64aadrecon = "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"
        }
        Else
        {
		    $base64aadrecon = "/9j/4AAQSkZJRgABAQAAkACQAAD/4QCeRXhpZgAATU0AKgAAAAgABQESAAMAAAABAAEAAAEaAAUAAAABAAAASgEbAAUAAAABAAAAUgEoAAMAAAABAAIAAIdpAAQAAAABAAAAWgAAAAAAAACQAAAAAQAAAJAAAAABAAOShgAHAAAAEgAAAISgAgAEAAAAAQAAAaagAwAEAAAAAQAAAD4AAAAAQVNDSUkAAABTY3JlZW5zaG90/+EJIWh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8APD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS40LjAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIi8+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgPD94cGFja2V0IGVuZD0idyI/PgD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZjwCyBOmACZjs+EJ+/+IPrElDQ19QUk9GSUxFAAEBAAAPnGFwcGwCEAAAbW50clJHQiBYWVogB+QAAQAHABYAIwAjYWNzcEFQUEwAAAAAQVBQTAAAAAAAAAAAAAAAAAAAAAAAAPbWAAEAAAAA0y1hcHBsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARZGVzYwAAAVAAAABiZHNjbQAAAbQAAASCY3BydAAABjgAAAAjd3RwdAAABlwAAAAUclhZWgAABnAAAAAUZ1hZWgAABoQAAAAUYlhZWgAABpgAAAAUclRSQwAABqwAAAgMYWFyZwAADrgAAAAgdmNndAAADtgAAAAwbmRpbgAADwgAAAA+Y2hhZAAAD0gAAAAsbW1vZAAAD3QAAAAoYlRSQwAABqwAAAgMZ1RSQwAABqwAAAgMYWFiZwAADrgAAAAgYWFnZwAADrgAAAAgZGVzYwAAAAAAAAAIRGlzcGxheQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAG1sdWMAAAAAAAAAJgAAAAxockhSAAAAFAAAAdhrb0tSAAAADAAAAexuYk5PAAAAEgAAAfhpZAAAAAAAEgAAAgpodUhVAAAAFAAAAhxjc0NaAAAAFgAAAjBkYURLAAAAHAAAAkZubE5MAAAAFgAAAmJmaUZJAAAAEAAAAnhpdElUAAAAFAAAAohlc0VTAAAAEgAAApxyb1JPAAAAEgAAApxmckNBAAAAFgAAAq5hcgAAAAAAFAAAAsR1a1VBAAAAHAAAAthoZUlMAAAAFgAAAvR6aFRXAAAACgAAAwp2aVZOAAAADgAAAxRza1NLAAAAFgAAAyJ6aENOAAAACgAAAwpydVJVAAAAJAAAAzhlbkdCAAAAFAAAA1xmckZSAAAAFgAAA3BtcwAAAAAAEgAAA4ZoaUlOAAAAEgAAA5h0aFRIAAAADAAAA6pjYUVTAAAAGAAAA7ZlbkFVAAAAFAAAA1xlc1hMAAAAEgAAApxkZURFAAAAEAAAA85lblVTAAAAEgAAA95wdEJSAAAAGAAAA/BwbFBMAAAAEgAABAhlbEdSAAAAIgAABBpzdlNFAAAAEAAABDx0clRSAAAAFAAABExwdFBUAAAAFgAABGBqYUpQAAAADAAABHYATABDAEQAIAB1ACAAYgBvAGoAac7st+wAIABMAEMARABGAGEAcgBnAGUALQBMAEMARABMAEMARAAgAFcAYQByAG4AYQBTAHoA7QBuAGUAcwAgAEwAQwBEAEIAYQByAGUAdgBuAP0AIABMAEMARABMAEMARAAtAGYAYQByAHYAZQBzAGsA5gByAG0ASwBsAGUAdQByAGUAbgAtAEwAQwBEAFYA5AByAGkALQBMAEMARABMAEMARAAgAGMAbwBsAG8AcgBpAEwAQwBEACAAYwBvAGwAbwByAEEAQwBMACAAYwBvAHUAbABlAHUAciAPAEwAQwBEACAGRQZEBkgGRgYpBBoEPgQ7BEwEPgRABD4EMgQ4BDkAIABMAEMARCAPAEwAQwBEACAF5gXRBeIF1QXgBdlfaYJyAEwAQwBEAEwAQwBEACAATQDgAHUARgBhAHIAZQBiAG4A/QAgAEwAQwBEBCYEMgQ1BEIEPQQ+BDkAIAQWBBoALQQ0BDgEQQQ/BDsENQQ5AEMAbwBsAG8AdQByACAATABDAEQATABDAEQAIABjAG8AdQBsAGUAdQByAFcAYQByAG4AYQAgAEwAQwBECTAJAgkXCUAJKAAgAEwAQwBEAEwAQwBEACAOKg41AEwAQwBEACAAZQBuACAAYwBvAGwAbwByAEYAYQByAGIALQBMAEMARABDAG8AbABvAHIAIABMAEMARABMAEMARAAgAEMAbwBsAG8AcgBpAGQAbwBLAG8AbABvAHIAIABMAEMARAOIA7MDxwPBA8kDvAO3ACADvwO4A8wDvQO3ACAATABDAEQARgDkAHIAZwAtAEwAQwBEAFIAZQBuAGsAbABpACAATABDAEQATABDAEQAIABhACAAQwBvAHIAZQBzMKsw6TD8AEwAQwBEAAB0ZXh0AAAAAENvcHlyaWdodCBBcHBsZSBJbmMuLCAyMDIwAABYWVogAAAAAAAA8xYAAQAAAAEWylhZWiAAAAAAAACC9AAAPWT///+8WFlaIAAAAAAAAEwkAAC0hQAACuZYWVogAAAAAAAAJ74AAA4XAADIi2N1cnYAAAAAAAAEAAAAAAUACgAPABQAGQAeACMAKAAtADIANgA7AEAARQBKAE8AVABZAF4AYwBoAG0AcgB3AHwAgQCGAIsAkACVAJoAnwCjAKgArQCyALcAvADBAMYAywDQANUA2wDgAOUA6wDwAPYA+wEBAQcBDQETARkBHwElASsBMgE4AT4BRQFMAVIBWQFgAWcBbgF1AXwBgwGLAZIBmgGhAakBsQG5AcEByQHRAdkB4QHpAfIB+gIDAgwCFAIdAiYC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"
        }

        $bytes = [System.Convert]::FromBase64String($base64aadrecon)
        Remove-Variable base64aadrecon -ErrorAction SilentlyContinue

        $CompanyLogo = (Join-AADRPath @($ReportPath, 'AzureADRecon_Logo.jpg'))
		$p = New-Object IO.MemoryStream($bytes, 0, $bytes.length)
		$p.Write($bytes, 0, $bytes.length)
        Add-Type -AssemblyName System.Drawing
		$picture = [System.Drawing.Image]::FromStream($p, $true)
		$picture.Save($CompanyLogo)

        Remove-Variable bytes -ErrorAction SilentlyContinue
        Remove-Variable p -ErrorAction SilentlyContinue
        Remove-Variable picture -ErrorAction SilentlyContinue

        $LinkToFile = $false
        $SaveWithDocument = $true
        $Left = 0
        $Top = 0
        $Width = 180
        $Height = 50

        # Add image to the Sheet
        $worksheet.Shapes.AddPicture($CompanyLogo, $LinkToFile, $SaveWithDocument, $Left, $Top, $Width, $Height) | Out-Null

        Remove-Variable LinkToFile -ErrorAction SilentlyContinue
        Remove-Variable SaveWithDocument -ErrorAction SilentlyContinue
        Remove-Variable Left -ErrorAction SilentlyContinue
        Remove-Variable Top -ErrorAction SilentlyContinue
        Remove-Variable Width -ErrorAction SilentlyContinue
        Remove-Variable Height -ErrorAction SilentlyContinue

        If (Test-Path -Path $CompanyLogo)
        {
            Remove-Item $CompanyLogo
        }
        Remove-Variable CompanyLogo -ErrorAction SilentlyContinue

        $row = 5
        $column = 1
        $worksheet.Cells.Item($row,$column)= "Table of Contents"
        $worksheet.Cells.Item($row,$column).Style = "Heading 2"
        $row++

        For($i=2; $i -le $workbook.Worksheets.Count; $i++)
        {
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,$column) , "" , "'$($workbook.Worksheets.Item($i).Name)'!A1", "", $workbook.Worksheets.Item($i).Name) | Out-Null
            $row++
        }

        $row++
		$workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,1) , "https://github.com/adrecon/AzureADRecon", "" , "", "github.com/adrecon/AzureADRecon") | Out-Null

        $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null

        $excel.Windows.Item(1).Displaygridlines = $false
        $excel.ScreenUpdating = $true
        $ADStatFileName = Join-Path $ExcelPath 'AzureADRecon-Report.xlsx'
        Try
        {
            # Disable prompt if file exists
            $excel.DisplayAlerts = $False
            $workbook.SaveAs($ADStatFileName)
            Write-Output "[+] Excelsheet Saved to: $ADStatFileName"
        }
        Catch
        {
            throw [System.IO.IOException]::new('Excel workbook could not be saved. Check workbook access and disk space.')
        }
        $excel.Quit()
        Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet -Final $true
        Remove-Variable worksheet -ErrorAction SilentlyContinue
        Get-ADRExcelComObjRelease -ComObjtoRelease $workbook -Final $true
        Remove-Variable -Name workbook -Scope Global -ErrorAction SilentlyContinue
        Get-ADRExcelComObjRelease -ComObjtoRelease $excel -Final $true
        Remove-Variable -Name excel -Scope Global -ErrorAction SilentlyContinue
        Remove-Variable Method -ErrorAction SilentlyContinue
    }
}

Function Test-AADRGraphNotFoundError
{
    param([Parameter(Mandatory = $true)] [System.Management.Automation.ErrorRecord] $Record)
    $details = @($Record.FullyQualifiedErrorId, $Record.Exception.Message, $Record.ErrorDetails.Message) -join ' '
    If ($details -match 'Request_ResourceNotFound|ResourceNotFound|ErrorItemNotFound') { return $true }
    ForEach ($status in @($Record.Exception.ResponseStatusCode, $Record.Exception.StatusCode, $Record.Exception.Response.StatusCode))
    {
        If ([string]$status -in @('404', 'NotFound')) { return $true }
    }
    return $false
}

Function Invoke-AADRGraphCollection
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $Name,
          [Parameter(Mandatory = $true)] [scriptblock] $Query,
          [hashtable] $State,
          [switch] $AllowNotFound)

    If ($null -ne $State) { $State.Available = $false; $State.FailureStatus = $null; $State.Diagnostic = $null }
    Try
    {
        $items = @(Invoke-AADRReadRequest -Query $Query)
        If ($null -ne $State) { $State.Available = $true }
        $items | Where-Object { $null -ne $_ }
    }
    Catch
    {
        If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }
        $diagnostic = Get-AADRErrorDiagnostic $_
        If ($AllowNotFound -and (Test-AADRGraphNotFoundError -Record $_))
        {
            If ($null -ne $State) { $State.FailureStatus='Unavailable'; $State.Diagnostic=$diagnostic }
            Add-AADREvent -Operation $Name -Status Unavailable -Detail 'Referenced object was not found.'
            Write-Warning "[$Name] Referenced object was not found. Its known ID is retained and unavailable details remain blank."
            return
        }
        If (Test-AADRGraphPermissionError -Record $_)
        {
            If ($null -ne $State) { $State.FailureStatus='PermissionDenied'; $State.Diagnostic=$diagnostic }
            Add-AADREvent -Operation $Name -Status PermissionDenied -Detail 'Graph denied access to this dataset.' -Diagnostic $diagnostic
            Write-Warning "[$Name] Graph denied access. This data is unavailable, not an empty result. Check Graph permissions, delegated role, and licensing. Other collections will continue." -WarningAction Continue
        }
        Else
        {
            If ($null -ne $State) { $State.FailureStatus='Unavailable'; $State.Diagnostic=$diagnostic }
            Add-AADREvent -Operation $Name -Status Unavailable -Detail 'Dataset request failed. Partial pages were discarded; data remains unknown. Independent requests will continue.' -Diagnostic $diagnostic
            Write-Warning "[$Name] Dataset unavailable ($($diagnostic.Category)). Independent requests will continue. $($diagnostic.Guidance)" -WarningAction Continue
        }
    }
}

Function Get-AADRTenant
{
<#
.SYNOPSIS
    Returns information of the current AzureAD Tenant.

.DESCRIPTION
    Returns information of the current AzureAD Tenant.

.PARAMETER Method
    [string]
    Which method to use; MSGraph (default).

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    $AADRTenantObj = $null
    If ($Method -eq 'MSGraph')
    {
        $AADRTenant = @(Invoke-AADRGraphCollection -Name 'Get-AADRTenant' -Query { Get-MgOrganization -All -ErrorAction Stop })
        If ($AADRTenant)
        {
            $AADRTenantObj = @()
            $AADRTenant | ForEach-Object {

                $ObjValues = @("DisplayName", $_.DisplayName, "Creation Date", $(If (-Not [string]::IsNullOrWhiteSpace([string]$_.CreatedDateTime)) { [datetime]$_.CreatedDateTime } Else { $null }))

                For ($i = 0; $i -lt $($ObjValues.Count); $i++)
                {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$i]
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i + 1]
                    $i++
                    $AADRTenantObj += $Obj
                }

                For ($i = 0; $i -lt $_.VerifiedDomains.Count; $i++)
                {
                    $ObjValues = @("VerifiedDomain($i) - Name", $_.VerifiedDomains[$i].Name, "VerifiedDomain($i) - Type", $_.VerifiedDomains[$i].Type, "VerifiedDomain($i) - Capabilities", $_.VerifiedDomains[$i].Capabilities, "VerifiedDomain($i) - _Default", $_.VerifiedDomains[$i].IsDefault, "VerifiedDomain($i) - Initial", $_.VerifiedDomains[$i].IsInitial)

                    For ($icnt = 0; $icnt -lt $($ObjValues.Count); $icnt++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$icnt]
                        $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$icnt + 1]
                        $icnt++
                        $AADRTenantObj += $Obj
                    }
                }

                $TechnicalNotificationMails = If ($null -ne $_.TechnicalNotificationMails) { $_.TechnicalNotificationMails -join ',' } Else { $null }

                $ObjValues = @("OnPremisesSyncEnabled", $_.OnPremisesSyncEnabled, "OnPremisesLastSyncDateTime", $_.OnPremisesLastSyncDateTime, "Street", $_.Street, "City", $_.City, "PostalCode", $_.PostalCode, "State", $_.State, "Country", $_.Country, "CountryLetterCode", $_.CountryLetterCode, "TelephoneNumber", $_.TelephoneNumber, "TechnicalNotificationMails", $TechnicalNotificationMails)

                For ($i = 0; $i -lt $($ObjValues.Count); $i++)
                {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$i]
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i + 1]
                    $i++
                    $AADRTenantObj += $Obj
                }

                For ($i = 0; $i -lt $_.AssignedPlans.Count; $i++)
                {

                    $ObjValues = @("AssignedPlan($i) - AssignedTimestamp", $(If (-Not [string]::IsNullOrWhiteSpace([string]$_.AssignedPlans[$i].AssignedDateTime)) { [datetime]$_.AssignedPlans[$i].AssignedDateTime } Else { $null }), "AssignedPlan($i) - CapabilityStatus", $_.AssignedPlans[$i].CapabilityStatus, "AssignedPlan($i) - Service", $_.AssignedPlans[$i].Service, "AssignedPlan($i) - ServicePlanId", $_.AssignedPlans[$i].ServicePlanId)

                    For ($icnt = 0; $icnt -lt $($ObjValues.Count); $icnt++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$icnt]
                        $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$icnt + 1]
                        $icnt++
                        $AADRTenantObj += $Obj
                    }
                }
            }
            Remove-Variable AADRTenant -ErrorAction SilentlyContinue
        }
    }

    If ($AADRTenantObj)
    {
        Return $AADRTenantObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRDomain
{
<#
.SYNOPSIS
    Returns information of the current (or specified) AzureAD.

.DESCRIPTION
    Returns information of the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; MSGraph (default).

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    $AADRDomainObj = $null
    If ($Method -eq 'MSGraph')
    {
        $AADRDomain = @(Invoke-AADRGraphCollection -Name 'Get-AADRDomain' -Query { Get-MgDomain -All -ErrorAction Stop })
        If ($AADRDomain)
        {
            $AADRDomainObj = $script:AADRParserType::DomainParser($AADRDomain, $Threads)
            Remove-Variable AADRDomain -ErrorAction SilentlyContinue
        }
    }

    If ($AADRDomainObj)
    {
        Return $AADRDomainObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRLicense
{
<#
.SYNOPSIS
    Returns information of the current (or specified) AzureAD.

.DESCRIPTION
    Returns information of the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    $records = @(Invoke-AADRGraphCollection -Name Licenses -Query { Get-AADRGraphPages 'subscribedSkus' })
    ForEach ($record in $records)
    {
        [pscustomobject][ordered]@{
            ObjectId=$record.id; SkuId=$record.skuId; SkuPartNumber=$record.skuPartNumber
            ConsumedUnits=$record.consumedUnits; CapabilityStatus=$record.capabilityStatus; AppliesTo=$record.appliesTo
            'PrepaidUnits-Enabled'=$record.prepaidUnits.enabled; 'PrepaidUnits-Suspended'=$record.prepaidUnits.suspended
            'PrepaidUnits-Warning'=$record.prepaidUnits.warning; 'PrepaidUnits-LockedOut'=$record.prepaidUnits.lockedOut
            'ServicePlans (Count)'=$(If ($null -ne $record.servicePlans) { @($record.servicePlans).Count } Else { $null })
            ServicePlans=$(If ($null -ne $record.servicePlans) { ConvertTo-Json -InputObject $record.servicePlans -Depth 10 -Compress } Else { $null })
        }
    }
}

Function Test-AADRGraphPermissionError
{
    param([Parameter(Mandatory = $true)] [System.Management.Automation.ErrorRecord] $Record)

    # Graph SDK versions expose service codes through different error surfaces.
    $details = @($Record.FullyQualifiedErrorId, $Record.Exception.Message, $Record.ErrorDetails.Message) -join ' '
    If ($details -match 'Authentication_RequestFromUnsupportedUserRole|Authorization_RequestDenied|AccessDenied|Forbidden|Authentication_MSGraphPermissionMissing|Authentication_RequestFromNonPremiumTenant')
    {
        return $true
    }
    ForEach ($status in @($Record.Exception.ResponseStatusCode, $Record.Exception.StatusCode,
        $Record.Exception.Response.StatusCode))
    {
        If ([string]$status -in @('403', 'Forbidden')) { return $true }
    }
    return $false
}

Function Get-AADRUser
{
<#
.SYNOPSIS
    Returns all users in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all users in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; MSGraph (default).

.PARAMETER date
    [DateTime]
    Date when AzureADRecon was executed.

.PARAMETER DormantTimeSpan
    [int]
    Timespan for Dormant accounts. Default 30 days.

.PARAMETER PassMaxAge
    [int]
    Timespan for Password age. Default 90 days.

.PARAMETER PageSize
    [int]
    The PageSize to set for the MSGraph query. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $true)]
        [int] $DormantTimeSpan = 30,

        [Parameter(Mandatory = $true)]
        [int] $PassMaxAge = 90,

        [Parameter(Mandatory = $true)]
        [int] $PageSize = 200,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    $date = $date.ToUniversalTime()
    If ($Method -eq 'MSGraph')
    {
        $userProperties = @('accountEnabled', 'assignedLicenses', 'createdDateTime', 'department', 'displayName',
            'jobTitle', 'mobilePhone', 'id', 'businessPhones', 'userPrincipalName', 'userType', 'lastPasswordChangeDateTime')
        $signInActivityAvailable = $true
        $mfaWarningShown = $false
        $mfaUnavailableWarningShown = $false
        $mfaConsecutiveFailures = 0
        $mfaCircuitOpen = $false
        $userRecordFailures = 0
        $UserObj = @()
        Try
        {
            $AzureADUsers = @(Invoke-AADRReadRequest -Query { Get-MgUser -All -Property ($userProperties + 'signInActivity') -PageSize ([Math]::Min($PageSize, 500)) -ErrorAction Stop })
        }
        Catch
        {
            If (-Not (Test-AADRGraphPermissionError -Record $_)) { throw }
            $signInActivityAvailable = $false
            Add-AADREvent -Operation 'Users signInActivity' -Status PermissionDenied -Detail 'Sign-in activity was denied; basic user collection retried.'
            Write-Warning '[Get-AADRUser] Graph denied signInActivity access. Retrying basic user collection. Sign-in and dormancy fields will be blank (unknown). Sign-in activity requires AuditLog.Read.All, Entra ID P1/P2, and a supported role for delegated access (for example Reports Reader).'
            Try
            {
                # A fresh request discards any partial pages from the failed query.
                $AzureADUsers = @(Invoke-AADRReadRequest -Query { Get-MgUser -All -Property $userProperties -PageSize $PageSize -ErrorAction Stop })
            }
            Catch
            {
                If (-Not (Test-AADRGraphPermissionError -Record $_)) { throw }
                Add-AADREvent -Operation Users -Status PermissionDenied -Detail 'Basic user access was denied.'
                Write-Warning '[Get-AADRUser] Graph also denied basic user access. Users were not collected. Check User.Read.All or Directory.Read.All and the signed-in account permissions.'
                return $null
            }
        }
        $registrationMap = @{}
        ForEach ($registration in @(Get-AADRMFARegistration)) { If ($registration.id) { $registrationMap[[string]$registration.id] = $registration } }
        If($AzureADUsers)
        {
            $ProcessedUserCount = 0
            Write-Verbose "[*] Total Users: $($AzureADUsers.Count)"
            $UserObj = @()
            $AzureADUsers | ForEach-Object {
                $ProcessedUserCount++
                $currentUser = $_
                Try
                {
                # Graph can omit dates or return null; keep unavailable report values unknown.
                $createdDateTime = ConvertTo-AADRDateTime $_.createdDateTime
                $lastPasswordChange = ConvertTo-AADRDateTime $_.lastPasswordChangeDateTime
                $lastSignIn = If ($signInActivityAvailable) { ConvertTo-AADRDateTime $_.signInActivity.lastSignInDateTime } Else { $null }
                $lastNonInteractiveSignIn = If ($signInActivityAvailable) { ConvertTo-AADRDateTime $_.signInActivity.lastNonInteractiveSignInDateTime } Else { $null }
                $lastSuccess = If ($signInActivityAvailable) { ConvertTo-AADRDateTime $_.signInActivity.lastSuccessfulSignInDateTime } Else { $null }
                $successAge = If ($null -ne $lastSuccess) { [Math]::Max(0, [Math]::Floor(($Date.ToUniversalTime() - $lastSuccess.ToUniversalTime()).TotalDays)) } Else { $null }
                $passwordAge = If ($null -ne $lastPasswordChange) { [Math]::Max(0, [Math]::Floor(($Date - $lastPasswordChange).TotalDays)) } Else { $null }

                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name 'Id' -Value $_.Id
                $Obj | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $(If (-Not [string]::IsNullOrWhiteSpace([string]$_.UserPrincipalName)) { [string]$_.UserPrincipalName } Else { $null })
                $Obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $(If (-Not [string]::IsNullOrWhiteSpace([string]$_.DisplayName)) { [string]$_.DisplayName } Else { $null })
                $Obj | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $(If (-Not [string]::IsNullOrWhiteSpace([string]$_.accountEnabled)) { [bool]$_.accountEnabled } Else { $null })
                $Obj | Add-Member -MemberType NoteProperty -Name "UserType" -Value $(If (-Not [string]::IsNullOrWhiteSpace([string]$_.userType)) { [string]$_.userType } Else { $null })
                $Obj | Add-Member -MemberType NoteProperty -Name "JobTitle" -Value $(If (-Not [string]::IsNullOrWhiteSpace([string]$_.JobTitle)) { [string]$_.JobTitle } Else { $null })
                $Obj | Add-Member -MemberType NoteProperty -Name "Department" -Value $(If (-Not [string]::IsNullOrWhiteSpace([string]$_.Department)) { [string]$_.Department } Else { $null })
                $Obj | Add-Member -MemberType NoteProperty -Name "Mobile" -Value $(If (-Not [string]::IsNullOrWhiteSpace([string]$_.MobilePhone)) { [string]$_.MobilePhone } Else { $null })

                # Missing sign-in history does not establish that a user has never signed in.
                $NeverLoggedIn = If ($null -ne $lastSuccess) { $False } Else { $null }
                $DaysSinceLastLogon = If ($null -ne $lastSignIn) { [Math]::Max(0, [Math]::Floor(($Date - $lastSignIn).TotalDays)) } Else { $null }

                $Obj | Add-Member -MemberType NoteProperty -Name "Never Logged In" -Value $NeverLoggedIn
                $Obj | Add-Member -MemberType NoteProperty -Name "Logon Age (days)" -Value $DaysSinceLastLogon

                $NeverNonInteractiveLoggedIn = $null # Attempt history cannot establish successful non-interactive access.
                $DaysSinceLastNonInteractiveLogon = If ($null -ne $lastNonInteractiveSignIn) { [Math]::Max(0, [Math]::Floor(($Date - $lastNonInteractiveSignIn).TotalDays)) } Else { $null }

                $Obj | Add-Member -MemberType NoteProperty -Name $("Dormant (> " + $DormantTimeSpan + " days)") -Value $(If ($null -eq $successAge) { $null } Else { $successAge -gt $DormantTimeSpan })
                $Obj | Add-Member -MemberType NoteProperty -Name "Never Non-Interactive Logged In" -Value $NeverNonInteractiveLoggedIn
                $Obj | Add-Member -MemberType NoteProperty -Name 'Non-Interactive Attempt Recorded' -Value $(If ($null -ne $lastNonInteractiveSignIn) { $true } Else { $null })
                $Obj | Add-Member -MemberType NoteProperty -Name "Non-Interactive Logon Age (days)" -Value $DaysSinceLastNonInteractiveLogon
                $Obj | Add-Member -MemberType NoteProperty -Name "Password Age (days)" -Value $passwordAge
                $Obj | Add-Member -MemberType NoteProperty -Name $("Password Age (>" + $PassMaxAge + " days)") -Value $(If ($null -ne $passwordAge) { $passwordAge -gt $PassMaxAge } Else { $null })
                $Obj | Add-Member -MemberType NoteProperty -Name 'LastSuccessfulSignInDateTime' -Value $lastSuccess
                $Obj | Add-Member -MemberType NoteProperty -Name 'Successful Logon Age (days)' -Value $successAge
                $Obj | Add-Member -MemberType NoteProperty -Name 'Last Interactive Attempt' -Value $lastSignIn
                $Obj | Add-Member -MemberType NoteProperty -Name 'Last Non-Interactive Attempt' -Value $lastNonInteractiveSignIn
                $Obj | Add-Member -MemberType NoteProperty -Name "LastSignInDateTime" -Value $lastSignIn
                $Obj | Add-Member -MemberType NoteProperty -Name "LastNonInteractiveSignInDateTime" -Value $lastNonInteractiveSignIn
                $Obj | Add-Member -MemberType NoteProperty -Name "lastPasswordChangeDateTime" -Value $lastPasswordChange

                # GetMFAStatusReport.ps1 https://o365reports.com/2022/04/27/get-mfa-status-of-office-365-users-using-microsoft-graph-powershell

                $Is3rdPartyAuthenticatorUsed = "False"
                $MFAPhone = $null
                $MicrosoftAuthenticatorDevice = $null
                Write-Progress -Activity "`n Processed users count: $ProcessedUserCount Currently processing: $($_.DisplayName)`n" -ErrorAction SilentlyContinue
                $mfaAvailable = $true
                Try
                {
                    If ($mfaCircuitOpen -or [string]::IsNullOrWhiteSpace([string]$_.Id))
                    {
                        $mfaAvailable = $false
                        $MFAData = @()
                    }
                    Else
                    {
                        # Capture the pipeline user before entering a nested request scope.
                        $mfaUserId = $_.Id
                        [array]$MFAData = @(Invoke-AADRReadRequest -Query { Get-MgUserAuthenticationMethod -UserId $mfaUserId -ErrorAction Stop })
                        $mfaConsecutiveFailures = 0
                    }
                }
                Catch
                {
                    If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }
                    $mfaAvailable = $false
                    $MFAData = @()
                    If (Test-AADRGraphPermissionError -Record $_)
                    {
                        $mfaConsecutiveFailures = 0
                        Add-AADREvent -Operation 'Users authentication methods' -Status PermissionDenied -Detail 'Authentication methods for a user were denied.' -Diagnostic (Get-AADRErrorDiagnostic $_)
                        If (-Not $mfaWarningShown)
                        {
                            Write-Warning '[Get-AADRUser] Graph denied authentication method access. Affected method details remain unknown. Registration/capability are independent. Check UserAuthenticationMethod.Read.All and the delegated role. Other users will still be queried.' -WarningAction Continue
                            $mfaWarningShown = $true
                        }
                    }
                    Else
                    {
                        Add-AADREvent -Operation 'Users authentication methods' -Status Unavailable -Detail 'Authentication methods for a user could not be collected; details remain unknown.' -Diagnostic (Get-AADRErrorDiagnostic $_)
                        If (-Not (Test-AADRGraphNotFoundError $_)) { $mfaConsecutiveFailures++ } Else { $mfaConsecutiveFailures = 0 }
                        If (-Not $mfaUnavailableWarningShown)
                        {
                            Write-Warning '[Get-AADRUser] Authentication method requests failed after retry or a user was removed during collection. Affected details remain unknown; basic user collection continues.' -WarningAction Continue
                            $mfaUnavailableWarningShown = $true
                        }
                        If ($mfaConsecutiveFailures -ge 5)
                        {
                            $mfaCircuitOpen = $true
                            Add-AADREvent -Operation 'Users authentication methods' -Status Unavailable -Detail 'Five consecutive method request failures after retries. Remaining method lookups were skipped; basic users and registration/capability remain independent.'
                            Write-Warning '[Get-AADRUser] Repeated authentication method failures. Skipping remaining method lookups for this collector; their details will remain unknown.' -WarningAction Continue
                        }
                    }
                }
                $AuthenticationMethod = @()
                $AdditionalDetails = @()

                foreach ($MFA in $MFAData)
                {
                    $AuthMethod = $null
                    $AuthMethodDetails = $null
                    If ($null -eq $MFA -or $null -eq $MFA.AdditionalProperties)
                    {
                        $mfaAvailable = $false
                        continue
                    }
                    Switch ($MFA.AdditionalProperties["@odata.type"])
                    {
                        "#microsoft.graph.passwordAuthenticationMethod" {
                            $AuthMethod = 'PasswordAuthentication'
                            $AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
                        }
                        "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                            # Microsoft Authenticator App
                            $AuthMethod = 'AuthenticatorApp'
                            $AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
                            $MicrosoftAuthenticatorDevice = $MFA.AdditionalProperties["displayName"]
                        }
                        "#microsoft.graph.phoneAuthenticationMethod" {
                            # Phone authentication
                            $AuthMethod = 'PhoneAuthentication'
                            $AuthMethodDetails = $MFA.AdditionalProperties["phoneType", "phoneNumber"] -join ' '
                            $MFAPhone = $MFA.AdditionalProperties["phoneNumber"]
                        }
                        "#microsoft.graph.fido2AuthenticationMethod" {
                            # FIDO2 key
                            $AuthMethod = 'Fido2'
                            $AuthMethodDetails = $MFA.AdditionalProperties["model"]
                        }
                        "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                            # Windows Hello
                            $AuthMethod = 'WindowsHelloForBusiness'
                            $AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
                        }
                        "#microsoft.graph.emailAuthenticationMethod" {
                            # Email Authentication
                            $AuthMethod = 'EmailAuthentication'
                            $AuthMethodDetails = $MFA.AdditionalProperties["emailAddress"]
                        }
                        "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                            # Temporary Access pass
                            $AuthMethod = 'TemporaryAccessPass'
                            $AuthMethodDetails = 'Access pass lifetime (minutes): ' + $MFA.AdditionalProperties["lifetimeInMinutes"]
                        }
                        "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" {
                            # Passwordless
                            $AuthMethod = 'PasswordlessMSAuthenticator'
                            $AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
                        }
                        "#microsoft.graph.softwareOathAuthenticationMethod" {
                            $AuthMethod = 'SoftwareOath'
                            $Is3rdPartyAuthenticatorUsed = "True"
                        }
                        default { $mfaAvailable = $false }
                    }
                    $AuthenticationMethod += $AuthMethod
                    If ($AuthMethodDetails -ne $null)
                    {
                        $AdditionalDetails += "$AuthMethod : $AuthMethodDetails"
                    }
                }

                #To remove duplicate authentication methods
                $AuthenticationMethod =$AuthenticationMethod | Sort-Object | Get-Unique
                $AuthenticationMethods= $AuthenticationMethod  -join ","
                $AdditionalDetail=$AdditionalDetails -join ", "
                # Registration/capability are report facts; method presence is not enforcement.
                $registration = $registrationMap[[string]$_.Id]
                $mfaRegistered = Get-AADRProperty $registration 'isMfaRegistered'
                $mfaCapable = Get-AADRProperty $registration 'isMfaCapable'
                $MFAStatus = If ($null -eq $mfaRegistered) { 'Unknown' } ElseIf ($mfaRegistered) { 'Registered' } Else { 'NotRegistered' }
                $Obj | Add-Member -MemberType NoteProperty -Name 'MFA Registered' -Value $mfaRegistered
                $Obj | Add-Member -MemberType NoteProperty -Name 'MFA Capable' -Value $mfaCapable
                $Obj | Add-Member -MemberType NoteProperty -Name 'MFA Enforcement' -Value 'NotEvaluated'
                $Obj | Add-Member -MemberType NoteProperty -Name 'MFA Enforcement Evidence' -Value 'Registration does not prove enforcement; evaluate Conditional Access and security defaults for the resource/sign-in context.'

                If (-Not $mfaAvailable)
                {
                    $AuthenticationMethods = $null
                    $AdditionalDetail = $null
                    $MFAPhone = $null
                    $MicrosoftAuthenticatorDevice = $null
                    $Is3rdPartyAuthenticatorUsed = $null
                }

                If (-Not $mfaAvailable) { Add-AADREvent -Operation 'Users authentication methods' -Status Unavailable -Detail 'User authentication method details are incomplete.' }
                $Obj | Add-Member -MemberType NoteProperty -Name "MFAStatus" -Value $MFAStatus
                $Obj | Add-Member -MemberType NoteProperty -Name "AuthenticationMethods" -Value $AuthenticationMethods
                $Obj | Add-Member -MemberType NoteProperty -Name "MFA Phone" -Value $MFAPhone
                $Obj | Add-Member -MemberType NoteProperty -Name "Microsoft Authenticator Configured Device" -Value $MicrosoftAuthenticatorDevice
                $Obj | Add-Member -MemberType NoteProperty -Name "Is 3rd-Party Authenticator Used" -Value $Is3rdPartyAuthenticatorUsed
                $Obj | Add-Member -MemberType NoteProperty -Name "Additional Details" -Value $AdditionalDetail
                $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value $createdDateTime
                $Obj | Add-Member -MemberType NoteProperty -Name "Assigned Licenses (Count)" -Value $(If ($null -ne $_.assignedLicenses) { @($_.assignedLicenses).Count } Else { $null })
                $Obj | Add-Member -MemberType NoteProperty -Name 'Authentication Methods Collection Status' -Value $(If ($mfaAvailable) {'Collected'} Else {'Unknown'})
                $UserObj += $Obj
                }
                Catch
                {
                    If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }
                    $userRecordFailures++
                    $recordDiagnostic = Get-AADRErrorDiagnostic $_
                    $rowId = $null
                    Try { $rowId = Get-AADRProperty $currentUser 'Id' } Catch { }
                    $UserObj += [pscustomobject]@{Id=$rowId; CollectionStatus='Unknown'; CollectionError=$recordDiagnostic.Category; MFAStatus='Unknown'; 'Authentication Methods Collection Status'='Unknown'}
                }
            }
            If ($userRecordFailures -gt 0)
            {
                Add-AADREvent -Operation 'Users records' -Status Unavailable -Detail "$userRecordFailures user record(s) could not be processed; known IDs and unknown status markers are retained."
                Write-Warning "[Get-AADRUser] $userRecordFailures user record(s) could not be processed. Unknown rows retain their known IDs; other users were processed." -WarningAction Continue
            }
            Write-Progress -Activity "Processed users" -Completed -ErrorAction SilentlyContinue
            Remove-Variable AzureADUsers -ErrorAction SilentlyContinue
        }
    }

    If ($UserObj)
    {
        Return $UserObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRServicePrincipal
{
<#
.SYNOPSIS
    Returns all service principals in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all service principals in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use.

.PARAMETER date
    [DateTime]
    Date when AzureADRecon was executed.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    $ServicePrincipalObj = $null
    If ($Method -eq 'MSGraph')
    {
        $records = @(Invoke-AADRGraphCollection -Name 'Get-AADRServicePrincipal' -Query { Get-MgServicePrincipal -All -Property id,appId,displayName,accountEnabled,publisherName,servicePrincipalNames,replyUrls,logoutUrl,oauth2PermissionScopes -ErrorAction Stop })
        If ($records.Count -gt 0) { $ServicePrincipalObj = $script:AADRParserType::ServicePrincipalParser($records, $Threads) }
    }

    If ($ServicePrincipalObj)
    {
        Return $ServicePrincipalObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRDirectoryRole
{
<#
.SYNOPSIS
    Returns all directory roles in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all directory roles in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; MSGraph (default).

.PARAMETER PageSize
    [int]
    The PageSize to set for the MSGraph query. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [int] $PageSize = 200,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    $AADRDirectoryRolesObj = $null
    If ($Method -eq 'MSGraph')
    {
        $AADRDirectoryRoles = @(Invoke-AADRGraphCollection -Name 'Get-AADRDirectoryRole' -Query { Get-MgDirectoryRole -All -ErrorAction Stop })
        If ($AADRDirectoryRoles)
        {
            Write-Verbose "[*] Total DirectoryRoles: $($AADRDirectoryRoles.Count)"
            $AADRDirectoryRolesObj = $script:AADRParserType::DirectoryRoleParser($AADRDirectoryRoles, $Threads)
            Remove-Variable AADRDirectoryRoles -ErrorAction SilentlyContinue
        }
    }

    If ($AADRDirectoryRolesObj)
    {
        Return $AADRDirectoryRolesObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRMemberProperty
{
    param([Parameter(Mandatory = $true)] $Member, [Parameter(Mandatory = $true)] [string] $Name)
    $property = $Member.PSObject.Properties[$Name]
    If ($property -and $null -ne $property.Value) { return $property.Value }
    If ($Member.AdditionalProperties) { return $Member.AdditionalProperties[$Name] }
    return $null
}

Function Get-AADRDirectoryRoleMember
{
<#
.SYNOPSIS
    Returns all directory role membership in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all directory role membership in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; MSGraph (default).

.PARAMETER PageSize
    [int]
    The PageSize to set for the MSGraph query. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [int] $PageSize = 200,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -ne 'MSGraph') { return $null }
    $parents = @(Invoke-AADRGraphCollection -Name 'Get-AADRDirectoryRoleMember' -Query { Get-MgDirectoryRole -All -ErrorAction Stop })
    Try
    {
        ForEach ($parent in $parents)
        {
            $parentName = $script:AADRParserType::CleanString($parent.DisplayName)
            $state = @{Available = $false}
            $members = @()
            If (-Not [string]::IsNullOrWhiteSpace([string]$parent.Id))
            {
                $members = @(Invoke-AADRGraphCollection -Name 'Get-AADRDirectoryRoleMember membership' -State $state -Query { Get-MgDirectoryRoleMember -DirectoryRoleId $parent.Id -All -ErrorAction Stop })
            }
            Else { Write-Warning '[Get-AADRDirectoryRoleMember] Parent ID is unavailable. Membership remains unknown.' }
            If (-Not $state.Available)
            {
                # An explicit status row distinguishes denied membership from a known empty group/role.
                [pscustomobject][ordered]@{
                    DirectoryRole = $parentName; MemberName = $null; MemberUserPrincipalName = $null
                    DirectoryRoleId = $parent.Id; MemberId = $null; Type = $null; CollectionStatus = 'Unknown'
                }
                continue
            }
            ForEach ($member in $members)
            {
                $name = $script:AADRParserType::CleanString((Get-AADRMemberProperty -Member $member -Name 'displayName'))
                $upn = $script:AADRParserType::CleanString((Get-AADRMemberProperty -Member $member -Name 'userPrincipalName'))
                $rawType = Get-AADRMemberProperty -Member $member -Name '@odata.type'
                $type = switch ($rawType)
                {
                    '#microsoft.graph.user' { 'User' }
                    '#microsoft.graph.group' { 'Group' }
                    '#microsoft.graph.servicePrincipal' { 'ServicePrincipal' }
                    '#microsoft.graph.device' { 'Device' }
                    default { $rawType }
                }
                    # Limited-information members retain their IDs even if details are denied.
                    If ($member.Id -and ($type -eq 'User' -or $type -eq 'Group') -and
                        ([string]::IsNullOrWhiteSpace($name) -or ($type -eq 'User' -and [string]::IsNullOrWhiteSpace($upn))))
                    {
                        $detailState = @{}
                        $details = If ($type -eq 'User')
                        {
                            Invoke-AADRGraphCollection -Name 'Get-AADRDirectoryRoleMember user details' -AllowNotFound -State $detailState -Query { Get-MgUser -UserId $member.Id -ErrorAction Stop }
                        }
                        Else
                        {
                            Invoke-AADRGraphCollection -Name 'Get-AADRDirectoryRoleMember group details' -AllowNotFound -State $detailState -Query { Get-MgGroup -GroupId $member.Id -ErrorAction Stop }
                        }
                        If ($details)
                        {
                            If ([string]::IsNullOrWhiteSpace($name)) { $name = $script:AADRParserType::CleanString($details.DisplayName) }
                            If ($type -eq 'User' -and [string]::IsNullOrWhiteSpace($upn)) { $upn = $script:AADRParserType::CleanString($details.UserPrincipalName) }
                        }
                    }
                [pscustomobject][ordered]@{
                    DirectoryRole = $parentName; MemberName = $name; MemberUserPrincipalName = $upn
                    DirectoryRoleId = $parent.Id; MemberId = $member.Id; Type = $type; CollectionStatus = 'Collected'
                }
            }
            Write-Progress -Activity 'Processing DirectoryRole membership' -Status $(If ([string]::IsNullOrWhiteSpace($parentName)) {'Unnamed object'} Else {$parentName}) -ErrorAction SilentlyContinue
        }
    }
    Finally { Write-Progress -Activity 'Processing DirectoryRole membership' -Completed }
}

Function Get-AADRGroup
{
<#
.SYNOPSIS
    Returns all groups in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all groups in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; MSGraph (default).

.PARAMETER PageSize
    [int]
    The PageSize to set for the MSGraph query. Default 200.

.PARAMETER date
    [DateTime]
    Date when AzureADRecon was executed.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $true)]
        [int] $PageSize = 200,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    $AADRGroupsObj = $null
    If ($Method -eq 'MSGraph')
    {
        $AADRGroups = @(Invoke-AADRGraphCollection -Name 'Get-AADRGroup' -Query { Get-MgGroup -All -PageSize $PageSize -ErrorAction Stop })
        If($AADRGroups)
        {
            Write-Verbose "[*] Total Groups: $($script:AADRParserType::ObjectCount($AADRGroups))"
            $AADRGroupsObj = $script:AADRParserType::GroupParser($AADRGroups, $Threads)
            Remove-Variable AADRGroups -ErrorAction SilentlyContinue
        }
    }

    If ($AADRGroupsObj)
    {
        Return $AADRGroupsObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRGroupMember
{
<#
.SYNOPSIS
    Returns all group membership in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all group membership in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; MSGraph (default).

.PARAMETER PageSize
    [int]
    The PageSize to set for the MSGraph query. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [int] $PageSize = 200,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -ne 'MSGraph') { return $null }
    $parents = @(Invoke-AADRGraphCollection -Name 'Get-AADRGroupMember' -Query { Get-MgGroup -All -ErrorAction Stop -PageSize $PageSize })
    Try
    {
        ForEach ($parent in $parents)
        {
            $parentName = $script:AADRParserType::CleanString($parent.DisplayName)
            $state = @{Available = $false}
            $members = @()
            If (-Not [string]::IsNullOrWhiteSpace([string]$parent.Id))
            {
                $members = @(Invoke-AADRGraphCollection -Name 'Get-AADRGroupMember membership' -State $state -Query { Get-MgGroupMember -GroupId $parent.Id -All -ErrorAction Stop })
            }
            Else { Write-Warning '[Get-AADRGroupMember] Parent ID is unavailable. Membership remains unknown.' }
            If (-Not $state.Available)
            {
                # An explicit status row distinguishes denied membership from a known empty group/role.
                [pscustomobject][ordered]@{
                    GroupName = $parentName; MemberName = $null; MemberUserPrincipalName = $null
                    GroupId = $parent.Id; MemberId = $null; Type = $null; CollectionStatus = 'Unknown'
                }
                continue
            }
            ForEach ($member in $members)
            {
                $name = $script:AADRParserType::CleanString((Get-AADRMemberProperty -Member $member -Name 'displayName'))
                $upn = $script:AADRParserType::CleanString((Get-AADRMemberProperty -Member $member -Name 'userPrincipalName'))
                $rawType = Get-AADRMemberProperty -Member $member -Name '@odata.type'
                $type = switch ($rawType)
                {
                    '#microsoft.graph.user' { 'User' }
                    '#microsoft.graph.group' { 'Group' }
                    '#microsoft.graph.servicePrincipal' { 'ServicePrincipal' }
                    '#microsoft.graph.device' { 'Device' }
                    default { $rawType }
                }
                [pscustomobject][ordered]@{
                    GroupName = $parentName; MemberName = $name; MemberUserPrincipalName = $upn
                    GroupId = $parent.Id; MemberId = $member.Id; Type = $type; CollectionStatus = 'Collected'
                }
            }
            Write-Progress -Activity 'Processing Group membership' -Status $(If ([string]::IsNullOrWhiteSpace($parentName)) {'Unnamed object'} Else {$parentName}) -ErrorAction SilentlyContinue
        }
    }
    Finally { Write-Progress -Activity 'Processing Group membership' -Completed }
}

Function Get-AADRDevice
{
<#
.SYNOPSIS
    Returns all devices in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all devices in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    $AADRDevicesObj = $null
    If ($Method -eq 'MSGraph')
    {
        $AADRDevices = @(Invoke-AADRGraphCollection -Name 'Get-AADRDevice' -Query { Get-MgDevice -All -ErrorAction Stop })
        If ($AADRDevices)
        {
            Write-Verbose "[*] Total Devices: $($script:AADRParserType::ObjectCount($AADRDevices))"
            $AADRDevicesObj = $script:AADRParserType::DeviceParser($AADRDevices, $Threads)
            <#
            $AADRDevicesObj = @()
            $AADRDevices | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $([string] $_.DisplayName)
                $Obj | Add-Member -MemberType NoteProperty -Name "AccountEnabled" -Value $([string] $_.AccountEnabled)
                $Obj | Add-Member -MemberType NoteProperty -Name "OnPremisesSyncEnabled" -Value $([string] $_.OnPremisesSyncEnabled)
                $Obj | Add-Member -MemberType NoteProperty -Name "OnPremisesLastSyncDateTime" -Value $([string] $_.OnPremisesLastSyncDateTime)
                $Obj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $([string] $_.OperatingSystem)
                $Obj | Add-Member -MemberType NoteProperty -Name "OperatingSystemVersion" -Value $([string] $_.OperatingSystemVersion)
                $Obj | Add-Member -MemberType NoteProperty -Name "ApproximateLastSignInDateTime" -Value $([string] $_.ApproximateLastSignInDateTime)
                $Obj | Add-Member -MemberType NoteProperty -Name "TrustType" -Value $([string] $_.TrustType)
                $Obj | Add-Member -MemberType NoteProperty -Name "ProfileType" -Value $([string] $_.ProfileType)
                $Obj | Add-Member -MemberType NoteProperty -Name "IsManaged" -Value $([string] $_.IsManaged)
                $Obj | Add-Member -MemberType NoteProperty -Name "DeviceOwnership" -Value $([string] $_.DeviceOwnership)
                $Obj | Add-Member -MemberType NoteProperty -Name "EnrollmentProfileName" -Value $([string] $_.EnrollmentProfileName)
                $Obj | Add-Member -MemberType NoteProperty -Name "IsCompliant" -Value $([string] $_.IsCompliant)
                $Obj | Add-Member -MemberType NoteProperty -Name "DeviceId" -Value $([string] $_.DeviceId)
                $AADRDevicesObj += $Obj
            }
            #>
            Remove-Variable AADRDevices -ErrorAction SilentlyContinue
        }
    }

    If ($AADRDevicesObj)
    {
        Return $AADRDevicesObj
    }
    Else
    {
        Return $null
    }
}

Function Convert-ADRGuidArray {
    <#
.SYNOPSIS
    Translates an array of object GUIDs to "DisplayName (GUID)" using a lookup table.

.DESCRIPTION
    Used by Get-AADRConditionalAccessPolicy to make the Include/Exclude fields human
    readable. Tokens not present in the map - keywords such as All, None, Office365,
    GuestsOrExternalUsers, AllTrusted, or GUIDs that could not be resolved - are kept
    as-is. The result is returned as a comma-separated string.

.PARAMETER Values
    The array of tokens (GUIDs and/or keywords) to translate.

.PARAMETER Map
    [hashtable]
    GUID -> DisplayName lookup table.

.OUTPUTS
    [string]
#>
    param (
        [Parameter(Mandatory = $false)] $Values,
        [Parameter(Mandatory = $false)] [hashtable] $Map
    )
    If (-not $Values) { return '' }
    ($Values | ForEach-Object {
        $tok = [string] $_
        If ($Map -and $Map.ContainsKey($tok) -and -Not [string]::IsNullOrWhiteSpace([string]$Map[$tok])) { "$($Map[$tok]) ($tok)" } Else { $tok }
    }) -join ', '
}

Function Get-AADRConditionalAccessPolicy {
    <#
.SYNOPSIS
    Returns all conditional access policies in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all conditional access policies in the current (or specified) AzureAD.
    Referenced object GUIDs (roles, groups, users, applications and named locations)
    are resolved to "DisplayName (GUID)" where the required scopes are granted
    (Policy.Read.All, User.Read.All, Directory.Read.All). GUIDs that cannot be
    resolved are left unchanged.

.PARAMETER Method
    [string]
    Which method to use; MSGraph (default).

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method
    )

    $AADRCAPoliciesObj = $null
    If ($Method -eq 'MSGraph') {
        $AADRCAPolicies = @(Invoke-AADRGraphCollection -Name ConditionalAccessPolicies -Query { Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop })

        If ($AADRCAPolicies) {
            Write-Verbose "[*] Total Conditional Access Policies: $($AADRCAPolicies.Count)"

            # ---- Build GUID -> DisplayName lookup tables (best-effort; unresolved GUIDs are kept as-is) ----
            $RoleMap = @{}; $LocationMap = @{}; $GroupMap = @{}; $UserMap = @{}; $AppMap = @{}

            # Directory role templates cover the roles referenced by CA policies - single call.
            Try { Invoke-AADRReadRequest -Query { Get-MgDirectoryRoleTemplate -All -ErrorAction Stop } | ForEach-Object { $RoleMap[[string] $_.Id] = $_.DisplayName } }
            Catch { If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }; Add-AADREvent -Operation 'Conditional Access reference labels' -Status Unavailable -Detail 'A reference label could not be resolved; IDs are retained.' -Diagnostic (Get-AADRErrorDiagnostic $_); Write-Warning "[Get-AADRConditionalAccessPolicy] Could not enumerate directory role templates (Directory.Read.All?). $((Get-AADRErrorDiagnostic $_).Category)" }

            # Named locations - single call.
            Try { Invoke-AADRReadRequest -Query { Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop } | ForEach-Object { $LocationMap[[string] $_.Id] = $_.DisplayName } }
            Catch { If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }; Add-AADREvent -Operation 'Conditional Access reference labels' -Status Unavailable -Detail 'A reference label could not be resolved; IDs are retained.' -Diagnostic (Get-AADRErrorDiagnostic $_); Write-Warning "[Get-AADRConditionalAccessPolicy] Could not enumerate named locations. $((Get-AADRErrorDiagnostic $_).Category)" }

            # Well-known first-party Microsoft application IDs (resolved without a Graph call).
            $WellKnownApps = @{
                '00000002-0000-0ff1-ce00-000000000000' = 'Office 365 Exchange Online'
                '00000003-0000-0ff1-ce00-000000000000' = 'Office 365 SharePoint Online'
                '00000003-0000-0000-c000-000000000000' = 'Microsoft Graph'
                '00000004-0000-0ff1-ce00-000000000000' = 'Skype for Business Online'
                '00000006-0000-0ff1-ce00-000000000000' = 'Office 365 Portal'
                '00000012-0000-0000-c000-000000000000' = 'Microsoft Rights Management Services'
                '797f4846-ba00-4fd7-ba43-dac1f8f63013' = 'Windows Azure Service Management API'
                'c5393580-f805-4401-95e8-94b7a6ef2fc2' = 'Office 365 Management APIs'
            }

            # Collect distinct referenced GUIDs so groups/users/apps are resolved with the fewest calls.
            $guidRegex = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
            $GroupIds = New-Object System.Collections.Generic.HashSet[string]
            $UserIds = New-Object System.Collections.Generic.HashSet[string]
            $AppIds = New-Object System.Collections.Generic.HashSet[string]
            $AADRCAPolicies | ForEach-Object {
                $c = $_.Conditions
                @($c.Users.IncludeGroups) + @($c.Users.ExcludeGroups) | Where-Object { $_ -match $guidRegex } | ForEach-Object { [void] $GroupIds.Add([string] $_) }
                @($c.Users.IncludeUsers) + @($c.Users.ExcludeUsers)  | Where-Object { $_ -match $guidRegex } | ForEach-Object { [void] $UserIds.Add([string] $_) }
                @($c.Applications.IncludeApplications) + @($c.Applications.ExcludeApplications) | Where-Object { $_ -match $guidRegex } | ForEach-Object { [void] $AppIds.Add([string] $_) }
            }

            ForEach ($gid in $GroupIds) {
                Try { $GroupMap[$gid] = (Invoke-AADRReadRequest -Query { Get-MgGroup -GroupId $gid -ErrorAction Stop }).DisplayName }
                Catch { If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }; Add-AADREvent -Operation 'Conditional Access reference labels' -Status Unavailable -Detail 'A reference label could not be resolved; IDs are retained.' -Diagnostic (Get-AADRErrorDiagnostic $_); Write-Warning "[Get-AADRConditionalAccessPolicy] Could not resolve group $gid. $((Get-AADRErrorDiagnostic $_).Category)" }
            }
            ForEach ($uid in $UserIds) {
                Try { $u = Invoke-AADRReadRequest -Query { Get-MgUser -UserId $uid -ErrorAction Stop }; $UserMap[$uid] = $(If ($u.UserPrincipalName) { $u.UserPrincipalName } Else { $u.DisplayName }) }
                Catch { If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }; Add-AADREvent -Operation 'Conditional Access reference labels' -Status Unavailable -Detail 'A reference label could not be resolved; IDs are retained.' -Diagnostic (Get-AADRErrorDiagnostic $_); Write-Warning "[Get-AADRConditionalAccessPolicy] Could not resolve user $uid. $((Get-AADRErrorDiagnostic $_).Category)" }
            }
            ForEach ($aid in $AppIds) {
                If ($WellKnownApps.ContainsKey($aid)) { $AppMap[$aid] = $WellKnownApps[$aid]; continue }
                Try { $sp = @(Invoke-AADRReadRequest -Query { Get-MgServicePrincipal -Filter "appId eq '$aid'" -ErrorAction Stop }) | Select-Object -First 1; If ($sp) { $AppMap[$aid] = $sp.DisplayName } }
                Catch { If ((Test-AADRControlFlowError $_) -or (Test-AADRGraphAuthenticationError $_)) { throw }; Add-AADREvent -Operation 'Conditional Access reference labels' -Status Unavailable -Detail 'A reference label could not be resolved; IDs are retained.' -Diagnostic (Get-AADRErrorDiagnostic $_); Write-Warning "[Get-AADRConditionalAccessPolicy] Could not resolve application $aid. $((Get-AADRErrorDiagnostic $_).Category)" }
            }

            $AADRCAPoliciesObj = @()
            $AADRCAPolicies | ForEach-Object {
                $Conditions = $_.Conditions
                $Apps = $Conditions.Applications
                $Usr = $Conditions.Users
                $Grant = $_.GrantControls
                $Session = $_.SessionControls

                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $([string] $_.DisplayName)
                $Obj | Add-Member -MemberType NoteProperty -Name "State" -Value $([string] $_.State)
                $Obj | Add-Member -MemberType NoteProperty -Name "IncludeUsers" -Value $(Convert-ADRGuidArray -Values $Usr.IncludeUsers -Map $UserMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "ExcludeUsers" -Value $(Convert-ADRGuidArray -Values $Usr.ExcludeUsers -Map $UserMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "IncludeGroups" -Value $(Convert-ADRGuidArray -Values $Usr.IncludeGroups -Map $GroupMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "ExcludeGroups" -Value $(Convert-ADRGuidArray -Values $Usr.ExcludeGroups -Map $GroupMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "IncludeRoles" -Value $(Convert-ADRGuidArray -Values $Usr.IncludeRoles -Map $RoleMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "ExcludeRoles" -Value $(Convert-ADRGuidArray -Values $Usr.ExcludeRoles -Map $RoleMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "IncludeApplications" -Value $(Convert-ADRGuidArray -Values $Apps.IncludeApplications -Map $AppMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "ExcludeApplications" -Value $(Convert-ADRGuidArray -Values $Apps.ExcludeApplications -Map $AppMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "UserActions" -Value $([string] ($Apps.IncludeUserActions -join ', '))
                $Obj | Add-Member -MemberType NoteProperty -Name "ClientAppTypes" -Value $([string] ($Conditions.ClientAppTypes -join ', '))
                $Obj | Add-Member -MemberType NoteProperty -Name "IncludePlatforms" -Value $([string] ($Conditions.Platforms.IncludePlatforms -join ', '))
                $Obj | Add-Member -MemberType NoteProperty -Name "ExcludePlatforms" -Value $([string] ($Conditions.Platforms.ExcludePlatforms -join ', '))
                $Obj | Add-Member -MemberType NoteProperty -Name "IncludeLocations" -Value $(Convert-ADRGuidArray -Values $Conditions.Locations.IncludeLocations -Map $LocationMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "ExcludeLocations" -Value $(Convert-ADRGuidArray -Values $Conditions.Locations.ExcludeLocations -Map $LocationMap)
                $Obj | Add-Member -MemberType NoteProperty -Name "GrantControlsOperator" -Value $([string] $Grant.Operator)
                $Obj | Add-Member -MemberType NoteProperty -Name "GrantControls" -Value $([string] ($Grant.BuiltInControls -join ', '))
                $Obj | Add-Member -MemberType NoteProperty -Name "SignInFrequency" -Value $([string] $Session.SignInFrequency.Value)
                $Obj | Add-Member -MemberType NoteProperty -Name "CreatedDateTime" -Value $([string] $_.CreatedDateTime)
                $Obj | Add-Member -MemberType NoteProperty -Name "ModifiedDateTime" -Value $([string] $_.ModifiedDateTime)
                $Obj | Add-Member -MemberType NoteProperty -Name "Id" -Value $([string] $_.Id)
                $AADRCAPoliciesObj += $Obj
            }
            Remove-Variable AADRCAPolicies -ErrorAction SilentlyContinue
        }
    }

    If ($AADRCAPoliciesObj) {
        Return $AADRCAPoliciesObj
    }
    Else {
        Return $null
    }
}

Function Remove-EmptyAADROutputDir
{
    param([string] $AADROutputDir, [array] $OutputType)
    ForEach ($format in @('CSV','XML','JSON','HTML'))
    {
        $directory = Join-Path $AADROutputDir ($format + '-Files')
        If ((Test-Path -LiteralPath $directory) -and @(Get-ChildItem -LiteralPath $directory -Force).Count -eq 0)
        { Remove-Item -LiteralPath $directory -Force }
    }
    If ((Test-Path -LiteralPath $AADROutputDir) -and @(Get-ChildItem -LiteralPath $AADROutputDir -Force).Count -eq 0)
    { Remove-Item -LiteralPath $AADROutputDir -Force }
}

Function ConvertTo-AADRGraphAccessToken
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyString()]
        $Token
    )

    # Accept provider responses without stringifying objects or mixed pipeline output.
    If ($null -eq $Token -or $Token -is [array])
    {
        throw "AccessToken must contain exactly one token, not an array of tokens or status messages."
    }
    If ($Token -isnot [string] -and $Token -isnot [System.Security.SecureString])
    {
        $tokenProperty = $null
        ForEach ($name in @('access_token', 'AccessToken', 'Token'))
        {
            If ($Token -is [System.Collections.IDictionary] -and $Token.Contains($name))
            {
                $tokenProperty = $Token[$name]
                break
            }
            If ($Token.PSObject.Properties[$name])
            {
                $tokenProperty = $Token.PSObject.Properties[$name].Value
                break
            }
        }
        $Token = $tokenProperty
    }

    If ($Token -is [System.Security.SecureString])
    {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token)
        Try
        {
            $rawToken = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        }
        Finally
        {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
    ElseIf ($Token -is [string])
    {
        $rawToken = $Token
    }
    Else
    {
        throw "AccessToken must be a string, SecureString, or a provider response with a Token, AccessToken, or access_token property."
    }

    # Normalize common copy/paste wrappers; never merge status text into a token.
    $rawToken = $rawToken.Trim().Trim([char[]]@([char]34, [char]39)).Trim()
    $rawToken = $rawToken -replace '^(?i)Bearer\s+', ''
    $rawToken = $rawToken.Trim().Trim([char[]]@([char]34, [char]39)).Trim()
    If ([string]::IsNullOrWhiteSpace($rawToken) -or $rawToken -match '[\s\p{Cc}]' -or
        $rawToken -in @('System.Security.SecureString', 'Bearer'))
    {
        throw "AccessToken is empty or contains invalid text. Supply only the Microsoft Graph access token, without status messages or serialized objects."
    }

    # Graph owns token validation: its access tokens are not necessarily three-part JWTs.
    ConvertTo-SecureString -String $rawToken -AsPlainText -Force
}

Function ConvertTo-AADRTokenLease
{
    param([Parameter(Mandatory = $true)] $Token, $ExpiresOn)
    $expiry = $ExpiresOn
    If ($null -eq $expiry)
    {
        ForEach ($name in @('ExpiresOn', 'ExpiresOnUtc', 'expires_on'))
        {
            $expiry = Get-AADRProperty $Token $name
            If ($null -ne $expiry) { break }
        }
    }
    $expiresUtc = $null
    If ($null -ne $expiry)
    {
        Try
        {
            # OAuth v1 expires_on is Unix seconds; Az/MSAL ExpiresOn is a date.
            If ($expiry -is [DateTimeOffset]) { $expiresUtc = $expiry.UtcDateTime }
            ElseIf ($expiry -is [DateTime]) { $expiresUtc = $expiry.ToUniversalTime() }
            ElseIf ([string]$expiry -match '^\d{9,12}$') { $expiresUtc = [DateTime]::SpecifyKind([datetime]'1970-01-01', [DateTimeKind]::Utc).AddSeconds([double]$expiry) }
            Else { $expiresUtc = ([DateTimeOffset]::Parse([string]$expiry, [Globalization.CultureInfo]::InvariantCulture)).UtcDateTime }
        }
        Catch { throw 'Token expiry metadata is invalid. Supply ExpiresOn as an ISO 8601 UTC date or Unix seconds.' }
    }
    Else
    {
        $lifetime = Get-AADRProperty $Token 'expires_in'
        If ($null -ne $lifetime)
        {
            $seconds = 0.0
            If (-Not [double]::TryParse([string]$lifetime, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$seconds) -or $seconds -le 0 -or $seconds -gt 31536000)
            { throw 'Token expires_in metadata is invalid.' }
            $expiresUtc = [DateTime]::UtcNow.AddSeconds($seconds)
        }
    }
    # Expiry is a scheduling hint from the issuer response, never proof of validity.
    # Graph tokens can be opaque; do not require or decode JWT claims.
    [pscustomobject]@{Token=(ConvertTo-AADRGraphAccessToken $Token); ExpiresOnUtc=$expiresUtc}
}

Function Get-AADRTokenIdentity
{
    # Direct calls deliberately avoid the renewal wrapper (and recursion).
    Try { $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id' -OutputType PSObject -ErrorAction Stop }
    Catch { Stop-AADRTokenVerification -Record $_ -Operation IdentityOrganization }
    $organizations = @(Get-AADRProperty $org 'value')
    If ($organizations.Count -ne 1 -or -Not $organizations[0].id) { throw 'Could not verify token tenant.' }
    Try { $me = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id' -OutputType PSObject -ErrorAction Stop }
    Catch { Stop-AADRTokenVerification -Record $_ -Operation IdentityUser }
    If (-Not $me.id) { throw 'Could not verify the delegated token user.' }
    [pscustomobject]@{TenantId=[string]$organizations[0].id; UserId=[string]$me.id}
}

Function Invoke-AADRTokenProvider
{
    [CmdletBinding()]
    param([string] $Reason)
    $ErrorActionPreference = 'Stop'
    $session = $script:AADRTokenSession
    $forceRefresh = $Reason -ne 'Startup' -or ($null -ne $session.ExpiresOnUtc -and $session.ExpiresOnUtc -le [DateTime]::UtcNow.AddSeconds($session.WindowSeconds))
    $request = [pscustomobject]@{Reason=$Reason; ForceRefresh=$forceRefresh; ExpiresOnUtc=$session.ExpiresOnUtc; TenantId=$session.Identity.TenantId; UserId=$session.Identity.UserId}
    Try { $responses = @(& $session.Provider $request) }
    Catch { throw 'Token provider failed. Reauthenticate its source session or renew the Kerberos ticket, then resume the run. Provider exception details are suppressed to protect credentials.' }
    If ($responses.Count -ne 1 -or $null -eq $responses[0]) { throw 'Token provider must return exactly one token response; write status messages to the verbose/information stream.' }
    $lease = ConvertTo-AADRTokenLease $responses[0]
    If ($null -eq $lease.ExpiresOnUtc -and $session.Mode -eq 'Provider')
    { throw 'AccessTokenProvider must return expiry metadata (ExpiresOn, expires_on, or expires_in) with the token.' }
    If ($null -ne $lease.ExpiresOnUtc -and $lease.ExpiresOnUtc -le [DateTime]::UtcNow.AddSeconds($session.WindowSeconds))
    {
        If ($Reason -eq 'Startup') { return Invoke-AADRTokenProvider -Reason Proactive }
        $exception = New-Object System.Security.Authentication.AuthenticationException -ArgumentList 'Token provider returned an expired or nearly expired token. It must acquire a fresh token beyond the renewal window.'
        $record = New-Object System.Management.Automation.ErrorRecord -ArgumentList $exception, 'AADRTokenVerificationFailed', ([System.Management.Automation.ErrorCategory]::AuthenticationError), $null
        $diagnostic = Get-AADRErrorDiagnostic $record
        $diagnostic.Category = 'AuthenticationFailed'
        $diagnostic.Reason = 'ExpiryWithinRenewalWindow'
        $diagnostic.Operation = 'TokenMetadata'
        $diagnostic | Add-Member NoteProperty ExpiresOnUtc $lease.ExpiresOnUtc.ToString('o')
        $diagnostic | Add-Member NoteProperty ClientUtc ([DateTime]::UtcNow.ToString('o'))
        $diagnostic.Guidance = 'Issuer expiry metadata falls within the renewal window. Check UTC conversion, system time, and the token source. No Graph request was made for this rejected lease.'
        $exception.Data['AADRSafeDiagnostic'] = $diagnostic
        Add-AADREvent -Operation TokenMetadata -Status AuthenticationFailed -Detail 'Token expiry metadata was rejected locally before connecting to Graph.' -Diagnostic $diagnostic
        $PSCmdlet.ThrowTerminatingError($record)
    }
    return $lease
}

Function Initialize-AADRTokenSession
{
    [CmdletBinding()]
    param($Token, [scriptblock] $Provider, $ExpiresOn, [int] $WindowSeconds = 300, [string] $Mode = 'Provider', [hashtable] $ProviderParameters)
    $script:AADRTokenSession = [pscustomobject]@{Provider=$Provider; ProviderParameters=$ProviderParameters; Mode=$Mode; ExpiresOnUtc=$null; WindowSeconds=$WindowSeconds; Identity=$null; Busy=$false; Ready=$false; NextRenewalUtc=$null}
    $session = $script:AADRTokenSession
    $lease = $null
    If ($null -ne $Token) { $lease = ConvertTo-AADRTokenLease -Token $Token -ExpiresOn $ExpiresOn; $session.ExpiresOnUtc = $lease.ExpiresOnUtc }
    If ($null -eq $lease -or ($Provider -and $Mode -eq 'Provider' -and $null -eq $lease.ExpiresOnUtc) -or
        ($null -ne $lease.ExpiresOnUtc -and $lease.ExpiresOnUtc -le [DateTime]::UtcNow.AddSeconds($WindowSeconds)))
    {
        If (-Not $Provider) { throw 'The supplied token is expired or within the renewal window. Provide AccessTokenProvider or acquire a fresh token before starting.' }
        $lease = Invoke-AADRTokenProvider -Reason Startup
    }
    Try { Connect-AADRGraphAccessToken -Token $lease.Token }
    Catch
    {
        If (-Not $Provider -or -Not (Test-AADRGraphAuthenticationError $_)) { throw }
        # An opaque token can be expired/revoked even when metadata looks current.
        $lease = Invoke-AADRTokenProvider -Reason Unauthorized
        Connect-AADRGraphAccessToken -Token $lease.Token
    }
    If ($Provider) { $session.Identity = Get-AADRTokenIdentity }
    $session.ExpiresOnUtc = $lease.ExpiresOnUtc
    $session.NextRenewalUtc = [DateTime]::UtcNow.AddSeconds($WindowSeconds)
    $session.Ready = $true
    If ($null -eq $lease.ExpiresOnUtc)
    {
        If ($Provider) { Write-Warning 'Token source returned no expiry metadata. Graph verified the token; Seamless SSO will be reacquired periodically, but renewal before expiry cannot be guaranteed.' }
        Else { Write-Warning 'Token expiry is unknown. Graph verified it at startup, but proactive renewal requires a provider and expiry metadata. Pass the full token response or AccessTokenExpiresOn.' }
    }
    ElseIf (-Not $Provider) { Write-Warning 'No AccessTokenProvider was supplied. Collection will stop before the known token expiry; a token string alone cannot be refreshed.' }
}

Function Update-AADRTokenSession
{
    [CmdletBinding()]
    param([switch] $Force)
    $session = $script:AADRTokenSession
    If ($null -eq $session -or -Not $session.Ready -or $session.Busy) { return }
    $due = If ($null -ne $session.ExpiresOnUtc) { [DateTime]::UtcNow -ge $session.ExpiresOnUtc.AddSeconds(-$session.WindowSeconds) }
        Else { $null -ne $session.Provider -and [DateTime]::UtcNow -ge $session.NextRenewalUtc }
    If (-Not $Force -and -Not $due) { return }
    $session.Busy = $true
    Try
    {
        If (-Not $session.Provider)
        { throw 'Token renewal is required, but no AccessTokenProvider is configured. Acquire a fresh token or use SDK-managed authentication. Completed collector checkpoints are retained.' }
        $reason = If ($Force) { 'Unauthorized' } Else { 'Proactive' }
        $lease = Invoke-AADRTokenProvider -Reason $reason
        Connect-AADRGraphAccessToken -Token $lease.Token
        $identity = Get-AADRTokenIdentity
        If ($identity.TenantId -ne $session.Identity.TenantId -or $identity.UserId -ne $session.Identity.UserId)
        { throw 'Renewed token identifies a different tenant or user. Collection stopped to prevent mixing identities.' }
        $session.ExpiresOnUtc = $lease.ExpiresOnUtc
        $session.NextRenewalUtc = [DateTime]::UtcNow.AddSeconds($session.WindowSeconds)
        Add-AADREvent -Operation Authentication -Status TokenRenewed -Detail 'Replacement token accepted by Graph; tenant and delegated user matched.'
        If ($script:AADRRun -and $script:AADRRun.Status -eq 'Running') { Write-AADRManifest }
    }
    Catch
    {
        $renewalError = $_
        $session.Ready = $false
        Disconnect-AADRGraphSession
        Add-AADREvent -Operation Authentication -Status AuthenticationFailed -Detail 'Token renewal failed; collection stopped and completed collector checkpoints retained.'
        $exception = New-Object System.Security.Authentication.AuthenticationException -ArgumentList $renewalError.Exception.Message
        $diagnostic = Get-AADRErrorDiagnostic $renewalError
        $diagnostic.Category = 'AuthenticationFailed'
        $diagnostic.Guidance = 'Reauthenticate the token source and check provider expiry metadata or Kerberos availability before resuming.'
        $exception.Data['AADRSafeDiagnostic'] = $diagnostic
        $record = New-Object System.Management.Automation.ErrorRecord -ArgumentList $exception, 'AADRTokenRenewalFailed', ([System.Management.Automation.ErrorCategory]::AuthenticationError), $null
        $PSCmdlet.ThrowTerminatingError($record)
    }
    Finally { $session.Busy = $false }
}

Function Connect-AADRGraphAccessToken
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] $Token)

    $secureToken = ConvertTo-AADRGraphAccessToken -Token $Token
    $operation = 'GraphConnection'
    Try
    {
        Connect-MgGraph -AccessToken $secureToken -NoWelcome -ErrorAction Stop | Out-Null
        # A Graph context alone does not prove that the supplied token is accepted.
        # Organization access is also required by tenant collection.
        $operation = 'OrganizationProbe'
        Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id' -ErrorAction Stop | Out-Null
    }
    Catch
    {
        $verificationError = $_
        Disconnect-AADRGraphSession
        Stop-AADRTokenVerification -Record $verificationError -Operation $operation
    }
}

Function Initialize-AADRWindowsKerberosInterop
{
    If ('AzureADRecon.Native.KerberosTicketCache' -as [type]) { return }

    # Keep the native API adapter separate so failure and cleanup paths can be tested offline.
    # ABI: microsoft/win32metadata generation/WinSDK/RecompiledIdlHeaders/um/NTSecAPI.h
    # Calls: https://learn.microsoft.com/windows/win32/api/ntsecapi/nf-ntsecapi-lsacallauthenticationpackage
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace AzureADRecon.Native
{
    [StructLayout(LayoutKind.Sequential)]
    public struct LsaString
    {
        public ushort Length;
        public ushort MaximumLength;
        public IntPtr Buffer;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct Luid { public uint LowPart; public int HighPart; }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct KerbCryptoKey32 { public int KeyType; public uint Length; public uint Offset; }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct KerbSubmitTicketRequest
    {
        public int MessageType;
        public Luid LogonId;
        public uint Flags;
        public KerbCryptoKey32 Key;
        public uint KerbCredSize;
        public uint KerbCredOffset;
    }

    public interface ILsaApi
    {
        int Connect(out IntPtr handle);
        int Lookup(IntPtr handle, ref LsaString name, out uint package);
        int Call(IntPtr handle, uint package, IntPtr request, uint length,
            out IntPtr response, out uint responseLength, out int protocolStatus);
        int FreeResponse(IntPtr response);
        int Close(IntPtr handle);
        uint ToWinError(int status);
    }

    internal sealed class WindowsLsaApi : ILsaApi
    {
        [DllImport("secur32.dll")] private static extern int LsaConnectUntrusted(out IntPtr handle);
        [DllImport("secur32.dll")] private static extern int LsaLookupAuthenticationPackage(IntPtr handle, ref LsaString name, out uint package);
        [DllImport("secur32.dll")] private static extern int LsaCallAuthenticationPackage(IntPtr handle, uint package,
            IntPtr request, uint length, out IntPtr response, out uint responseLength, out int protocolStatus);
        [DllImport("secur32.dll")] private static extern int LsaFreeReturnBuffer(IntPtr buffer);
        [DllImport("secur32.dll")] private static extern int LsaDeregisterLogonProcess(IntPtr handle);
        [DllImport("advapi32.dll")] private static extern uint LsaNtStatusToWinError(int status);
        public int Connect(out IntPtr handle) { return LsaConnectUntrusted(out handle); }
        public int Lookup(IntPtr handle, ref LsaString name, out uint package) { return LsaLookupAuthenticationPackage(handle, ref name, out package); }
        public int Call(IntPtr handle, uint package, IntPtr request, uint length,
            out IntPtr response, out uint responseLength, out int protocolStatus)
        { return LsaCallAuthenticationPackage(handle, package, request, length, out response, out responseLength, out protocolStatus); }
        public int FreeResponse(IntPtr response) { return LsaFreeReturnBuffer(response); }
        public int Close(IntPtr handle) { return LsaDeregisterLogonProcess(handle); }
        public uint ToWinError(int status) { return LsaNtStatusToWinError(status); }
    }

    public static class KerberosTicketCache
    {
        // Only validate the outer DER envelope here; Windows validates ticket contents.
        public static void Validate(byte[] ticket)
        {
            if (ticket == null || ticket.Length < 4 || ticket[0] != 0x76)
                throw new ArgumentException("Expected a DER-encoded KRB-CRED (.kirbi) ticket.");
            int offset = 2;
            long length = ticket[1];
            if ((length & 0x80) != 0)
            {
                int count = (int)(length & 0x7f);
                if (count == 0 || count > 4 || ticket.Length < 2 + count)
                    throw new ArgumentException("Invalid KRB-CRED length.");
                length = 0;
                for (int i = 0; i < count; i++) length = (length << 8) | ticket[offset++];
            }
            if (length < 2 || length != ticket.Length - offset || ticket[offset] != 0x30)
                throw new ArgumentException("KRB-CRED is truncated or has an invalid envelope.");
        }

        public static void Import(byte[] ticket) { Import(ticket, new WindowsLsaApi()); }

        private static void Check(int status, string operation, ILsaApi api)
        {
            if (status < 0)
                throw new Win32Exception(unchecked((int)api.ToWinError(status)),
                    operation + " failed (NTSTATUS 0x" + unchecked((uint)status).ToString("X8") + ").");
        }

        public static void Import(byte[] ticket, ILsaApi api)
        {
            Validate(ticket);
            if (api == null) throw new ArgumentNullException("api");
            IntPtr handle = IntPtr.Zero;
            IntPtr response = IntPtr.Zero;
            IntPtr buffer = IntPtr.Zero;
            LsaString name = new LsaString();
            int bufferSize = 0;
            try
            {
                Check(api.Connect(out handle), "LsaConnectUntrusted", api);
                name.Length = 8;
                name.MaximumLength = 9;
                name.Buffer = Marshal.StringToHGlobalAnsi("Kerberos");
                uint package;
                Check(api.Lookup(handle, ref name, out package), "LsaLookupAuthenticationPackage", api);

                int headerSize = Marshal.SizeOf(typeof(KerbSubmitTicketRequest));
                bufferSize = checked(headerSize + ticket.Length);
                buffer = Marshal.AllocHGlobal(bufferSize);
                KerbSubmitTicketRequest request = new KerbSubmitTicketRequest();
                request.MessageType = 21; // KerbSubmitTicketMessage
                // LogonId zero targets the caller's logon session; no arbitrary LUID support.
                request.KerbCredOffset = (uint)headerSize;
                request.KerbCredSize = (uint)ticket.Length;
                Marshal.StructureToPtr(request, buffer, false);
                Marshal.Copy(ticket, 0, IntPtr.Add(buffer, headerSize), ticket.Length);
                uint responseLength;
                int protocolStatus;
                Check(api.Call(handle, package, buffer, (uint)bufferSize,
                    out response, out responseLength, out protocolStatus), "LsaCallAuthenticationPackage", api);
                Check(protocolStatus, "Kerberos ticket submission", api);
            }
            finally
            {
                if (buffer != IntPtr.Zero)
                {
                    for (int i = 0; i < bufferSize; i++) Marshal.WriteByte(buffer, i, 0);
                    Marshal.FreeHGlobal(buffer);
                }
                if (response != IntPtr.Zero) api.FreeResponse(response);
                if (name.Buffer != IntPtr.Zero) Marshal.FreeHGlobal(name.Buffer);
                if (handle != IntPtr.Zero) api.Close(handle);
            }
        }
    }
}
'@ -ErrorAction Stop
}

Function Import-AADRWindowsKerberosTicket
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [byte[]] $TicketBytes)

    If ($env:OS -ne 'Windows_NT')
    {
        throw 'KRB-CRED (.kirbi or Base64) import requires Windows. On Linux/macOS, supply a ccache.'
    }
    Initialize-AADRWindowsKerberosInterop
    [AzureADRecon.Native.KerberosTicketCache]::Import($TicketBytes)
    Write-Host '[Invoke-AzureADRecon] Imported KRB-CRED into the current Windows logon session. The ticket remains in the session cache after this script exits.'
}

Function Invoke-AADRWindowsSSORequest
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $Endpoint,
          [Parameter(Mandatory = $true)] [string] $Body)

    $request = [System.Net.HttpWebRequest]::Create($Endpoint)
    $request.Timeout = 60000
    $request.ReadWriteTimeout = 60000
    $request.Method = 'POST'
    $request.ContentType = 'application/soap+xml; charset=utf-8'
    $request.UseDefaultCredentials = $true
    $request.AllowAutoRedirect = $false
    $request.ConnectionGroupName = [Guid]::NewGuid().ToString('N')
    $requestStream = $null
    $response = $null
    $reader = $null
    Try
    {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
        $request.ContentLength = $bytes.Length
        $requestStream = $request.GetRequestStream()
        $requestStream.Write($bytes, 0, $bytes.Length)
        $requestStream.Dispose()
        $requestStream = $null
        $response = $request.GetResponse()
        $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
        $reader.ReadToEnd()
    }
    Finally
    {
        ForEach ($resource in @($reader, $response, $requestStream))
        {
            If ($resource)
            {
                Try { $resource.Dispose() }
                Catch { Add-AADREvent -Operation 'SSO cleanup' -Status CleanupFailed -Detail 'An HTTP resource could not be disposed.' -Diagnostic (Get-AADRErrorDiagnostic $_) }
            }
        }
        Try { $null = $request.ServicePoint.CloseConnectionGroup($request.ConnectionGroupName) }
        Catch { Add-AADREvent -Operation 'SSO cleanup' -Status CleanupFailed -Detail 'The HTTP connection group could not be closed.' -Diagnostic (Get-AADRErrorDiagnostic $_) }
    }
}

Function Get-AADRSeamlessSSOToken
{
<#
.SYNOPSIS
    Acquires an AccessToken for Microsoft Graph using Azure AD Seamless SSO (Kerberos HTTP/autologon.microsoftazuread-sso.com).

.DESCRIPTION
    Leverages domain-joined Windows Kerberos ticket authentication or Kerberos credential caches (ccache)
    against the Azure AD Seamless SSO endpoint (HTTP/autologon.microsoftazuread-sso.com).
    Supports Windows KRB-CRED import, existing ccache files, custom DC and Domain configuration, AADInternals, MSAL.PS, curl --negotiate,
    or native WS-Trust / SAML bearer token exchange.

.PARAMETER TenantID
    Tenant ID or domain name.

.PARAMETER Domain
    Active Directory domain name (FQDN).

.PARAMETER DC
    Domain Controller IP or hostname.

.PARAMETER TGT
    Kerberos ticket file or Base64 contents. Windows imports KRB-CRED (.kirbi) into the current logon session; Linux/macOS uses ccache.
#>
    param(
        [Parameter(Mandatory = $false)]
        [string] $TenantID,

        [Parameter(Mandatory = $false)]
        [string] $Domain,

        [Parameter(Mandatory = $false)]
        [string] $DC,

        [Parameter(Mandatory = $false)]
        [string] $TGT,

        [switch] $IncludeMetadata,

        [switch] $ForceRefresh
    )

    Write-Host "[Invoke-AzureADRecon] Attempting Seamless SSO (Kerberos HTTP/autologon.microsoftazuread-sso.com)..."

    $useWindowsTicket = $false
    $previousKrb5CCName = $env:KRB5CCNAME
    $previousKrb5Config = $env:KRB5_CONFIG
    $tempFilesToClean = [System.Collections.Generic.List[string]]::new()

    Try
    {
        # Import KRB-CRED on Windows or use an existing ccache without conversion.
        If (-Not [string]::IsNullOrEmpty($TGT))
        {
            $resolvedTGT = $null
            $ticketBytes = $null
            Try
            {
                If (Test-Path -LiteralPath $TGT -PathType Leaf -ErrorAction SilentlyContinue)
                {
                    $resolvedTGT = (Resolve-Path -LiteralPath $TGT).Path
                    $ticketBytes = [System.IO.File]::ReadAllBytes($resolvedTGT)
                }
                Else
                {
                    Try
                    {
                        $ticketBytes = [System.Convert]::FromBase64String(($TGT.Trim() -replace '\s+', ''))
                    }
                    Catch
                    {
                        throw 'TGT must be an existing ticket file or Base64-encoded ticket contents.'
                    }
                }

                If (($resolvedTGT -and $resolvedTGT.EndsWith('.kirbi', [System.StringComparison]::OrdinalIgnoreCase)) -or
                    ($ticketBytes.Length -gt 0 -and $ticketBytes[0] -eq 0x76))
                {
                    Import-AADRWindowsKerberosTicket -TicketBytes $ticketBytes
                    $useWindowsTicket = $true
                    # Force SSPI to use the Windows session rather than an inherited ccache.
                    $env:KRB5CCNAME = $null
                }
                Else
                {
                    If ($env:OS -eq 'Windows_NT')
                    {
                        throw 'ccache input requires Linux/macOS. On Windows, supply a KRB-CRED (.kirbi) file or Base64 KRB-CRED ticket.'
                    }
                    If ($ticketBytes.Length -lt 2 -or $ticketBytes[0] -ne 0x05 -or $ticketBytes[1] -notin @(1, 2, 3, 4))
                    {
                        throw 'TGT is neither a KRB-CRED ticket nor a supported ccache (invalid format header).'
                    }
                    If (-Not $resolvedTGT)
                    {
                        $resolvedTGT = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ticket_$([System.Guid]::NewGuid().ToString('N')).ccache")
                        $tempFilesToClean.Add($resolvedTGT)
                        [System.IO.File]::WriteAllBytes($resolvedTGT, $ticketBytes)
                    }
                    $env:KRB5CCNAME = $resolvedTGT
                    Write-Host "[Invoke-AzureADRecon] Set KRB5CCNAME to $resolvedTGT"
                }
            }
            Finally
            {
                If ($ticketBytes) { [Array]::Clear($ticketBytes, 0, $ticketBytes.Length) }
            }
        }

        # --- Handle Domain and DC configuration on Linux ---
        If ($env:OS -ne "Windows_NT" -and -Not [string]::IsNullOrEmpty($Domain) -and -Not [string]::IsNullOrEmpty($DC))
        {
            $realmUpper = $Domain.ToUpper()
            $tempKrb5 = [System.IO.Path]::GetTempFileName()
            $tempFilesToClean.Add($tempKrb5)
            $krb5ConfigContent = @"
[libdefaults]
    default_realm = $realmUpper
    dns_lookup_realm = false
    dns_lookup_kdc = false

[realms]
    $realmUpper = {
        kdc = $DC
        admin_server = $DC
    }
"@
            [System.IO.File]::WriteAllText($tempKrb5, $krb5ConfigContent)
            $env:KRB5_CONFIG = $tempKrb5
            Write-Host "[Invoke-AzureADRecon] Configured Kerberos realm $realmUpper with KDC $DC"
        }

        # Strategy 1: AADInternals module if available
        If (-Not $ForceRefresh -and [string]::IsNullOrEmpty($TGT) -and (Get-Module -ListAvailable -Name AADInternals))
        {
            Try
            {
                Write-Verbose "[Invoke-AzureADRecon] Using AADInternals for Seamless SSO token acquisition..."
                Import-Module AADInternals -ErrorAction Stop
                $tokenParam = @{}
                If (-Not [string]::IsNullOrEmpty($TenantID)) { $tokenParam['Tenant'] = $TenantID }
                If (-Not [string]::IsNullOrEmpty($Domain)) { $tokenParam['Domain'] = $Domain }
                $tokenObj = Get-AADIntAccessTokenForMSGraph @tokenParam -ErrorAction Stop
                If ($tokenObj)
                {
                    If ($IncludeMetadata) { return $tokenObj }
                    If ($tokenObj -is [string]) { return $tokenObj }
                    If ($tokenObj.access_token) { return $tokenObj.access_token }
                }
            }
            Catch
            {
                Write-Verbose "[Invoke-AzureADRecon] AADInternals token acquisition failed: operation unavailable"
            }
        }

        # Strategy 2: MSAL.PS module if available
        If (-Not $ForceRefresh -and [string]::IsNullOrEmpty($TGT) -and (Get-Module -ListAvailable -Name MSAL.PS))
        {
            Try
            {
                Write-Verbose "[Invoke-AzureADRecon] Using MSAL.PS Integrated Windows Authentication..."
                Import-Module MSAL.PS -ErrorAction Stop
                $msalParams = @{
                    ClientId              = "14d82eec-204b-4c2f-b7e8-296a70dab67e" # Microsoft Graph PowerShell
                    Scopes                = @("https://graph.microsoft.com/.default")
                    IntegratedWindowsAuth = $true
                }
                If (-Not [string]::IsNullOrEmpty($TenantID)) { $msalParams['TenantId'] = $TenantID }
                $msalRes = Get-MsalToken @msalParams -ErrorAction Stop
                If ($msalRes.AccessToken)
                {
                    If ($IncludeMetadata) { return $msalRes }
                    return $msalRes.AccessToken
                }
            }
            Catch
            {
                Write-Verbose "[Invoke-AzureADRecon] MSAL.PS IWA acquisition failed: operation unavailable"
            }
        }

        # Strategy 3: Az.Accounts (ONLY if already loaded in session to prevent assembly version conflict with Microsoft.Graph)
        If (-Not $ForceRefresh -and [string]::IsNullOrEmpty($TGT) -and (Get-Module -Name Az.Accounts))
        {
            Try
            {
                Write-Verbose "[Invoke-AzureADRecon] Checking active Az session for MSGraph access token..."
                $azToken = Get-AzAccessToken -ResourceTypeName MSGraph -ErrorAction Stop
                If ($azToken.Token)
                {
                    If ($IncludeMetadata) { return $azToken }
                    return $azToken.Token
                }
            }
            Catch
            {
                Write-Verbose "[Invoke-AzureADRecon] Az.Accounts token acquisition failed: operation unavailable"
            }
        }

        # Strategy 4: Linux curl --negotiate or Windows WS-Trust request
        $realm = If (-Not [string]::IsNullOrEmpty($TenantID)) { $TenantID } ElseIf (-Not [string]::IsNullOrEmpty($Domain)) { $Domain } Else { "common" }
        $endpoint = "https://autologon.microsoftazuread-sso.com/$realm/winauth/trust/2005/windowstransport"

        $rstXml = @"
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing">
  <s:Header>
    <a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>
    <a:To s:mustUnderstand="1">$endpoint</a:To>
  </s:Header>
  <s:Body>
    <t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust">
      <wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy">
        <a:EndpointReference>
          <a:Address>urn:federation:MicrosoftOnline</a:Address>
        </a:EndpointReference>
      </wsp:AppliesTo>
      <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>
      <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>
      <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>
    </t:RequestSecurityToken>
  </s:Body>
</s:Envelope>
"@

        $responseXml = $null

        # If on Linux or non-Windows, or curl is available and KRB5CCNAME is set, use curl with --negotiate
        $curlCmd = Get-Command "curl" -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        If (-Not $useWindowsTicket -and $curlCmd -and ($env:OS -ne "Windows_NT" -or -Not [string]::IsNullOrEmpty($env:KRB5CCNAME)))
        {
            Try
            {
                Write-Verbose "[Invoke-AzureADRecon] Executing curl with --negotiate using Kerberos cache..."
                $tempReq = [System.IO.Path]::GetTempFileName()
                $tempFilesToClean.Add($tempReq)
                [System.IO.File]::WriteAllText($tempReq, $rstXml)

                $curlOutput = (& $curlCmd.Source -s --connect-timeout 15 --max-time 60 --negotiate -u : -X POST "$endpoint" -H "Content-Type: application/soap+xml; charset=utf-8" --data-binary "@$tempReq") -join "`n"
                If ($LASTEXITCODE -eq 0 -and -Not [string]::IsNullOrEmpty($curlOutput) -and $curlOutput.Contains("Assertion"))
                {
                    $responseXml = $curlOutput
                }
            }
            Catch
            {
                Write-Verbose "[Invoke-AzureADRecon] curl --negotiate exchange error: operation unavailable"
            }
        }

        # Fallback to .NET HttpWebRequest with UseDefaultCredentials (for Windows SSPI or .NET GSSAPI)
        If ([string]::IsNullOrEmpty($responseXml))
        {
            Try
            {
                Write-Verbose "[Invoke-AzureADRecon] Performing .NET HttpWebRequest Kerberos exchange..."
                $responseXml = Invoke-AADRWindowsSSORequest -Endpoint $endpoint -Body $rstXml
            }
            Catch
            {
                Write-Verbose "[Invoke-AzureADRecon] .NET HttpWebRequest exchange error: operation unavailable"
            }
        }

        # Parse assertion from response
        If (-Not [string]::IsNullOrEmpty($responseXml))
        {
            $xmlDoc = [xml]$responseXml
            $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
            $nsMgr.AddNamespace("saml", "urn:oasis:names:tc:SAML:1.0:assertion")
            $assertionNode = $xmlDoc.SelectSingleNode("//saml:Assertion", $nsMgr)
            If (-Not $assertionNode)
            {
                # Try finding any assertion node regardless of namespace prefix
                $assertionNode = $xmlDoc.SelectSingleNode("//*[local-name()='Assertion']")
            }

            If ($assertionNode)
            {
                $assertionXml = $assertionNode.OuterXml
                $assertionBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($assertionXml))

                # Primary attempt: OAuth v1.0 token endpoint with resource
                Try
                {
                    $tokenBody = @{
                        grant_type = "urn:ietf:params:oauth:grant-type:saml1_1-bearer"
                        assertion  = $assertionBase64
                        resource   = "https://graph.microsoft.com"
                        client_id  = "14d82eec-204b-4c2f-b7e8-296a70dab67e" # Microsoft Graph PowerShell
                    }
                    $tokenRes = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$realm/oauth2/token" -Method POST -Body $tokenBody -ErrorAction Stop
                    If ($tokenRes.access_token)
                    {
                        Write-Host "[Invoke-AzureADRecon] Successfully acquired token via Seamless SSO Kerberos exchange."
                        If ($IncludeMetadata) { return $tokenRes }
                        return ([string]$tokenRes.access_token).Trim()
                    }
                }
                Catch
                {
                    Write-Verbose "[Invoke-AzureADRecon] OAuth v1 token exchange failed (operation unavailable), attempting v2.0 endpoint..."
                    # Secondary attempt: OAuth v2.0 token endpoint with scope
                    Try
                    {
                        $tokenBodyV2 = @{
                            grant_type = "urn:ietf:params:oauth:grant-type:saml1_1-bearer"
                            assertion  = $assertionBase64
                            scope      = "https://graph.microsoft.com/.default openid profile offline_access"
                            client_id  = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
                        }
                        $tokenResV2 = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$realm/oauth2/v2.0/token" -Method POST -Body $tokenBodyV2 -ErrorAction Stop
                        If ($tokenResV2.access_token)
                        {
                            Write-Host "[Invoke-AzureADRecon] Successfully acquired token via Seamless SSO Kerberos exchange (v2.0)."
                            If ($IncludeMetadata) { return $tokenResV2 }
                            return ([string]$tokenResV2.access_token).Trim()
                        }
                    }
                    Catch
                    {
                        Write-Verbose "[Invoke-AzureADRecon] OAuth v2.0 token exchange also failed: operation unavailable"
                    }
                }
            }
        }
    }
    Finally
    {
        $env:KRB5CCNAME = $previousKrb5CCName
        $env:KRB5_CONFIG = $previousKrb5Config
        ForEach ($f in $tempFilesToClean)
        {
            If (Test-Path $f) { Remove-Item -Path $f -Force -ErrorAction SilentlyContinue }
        }
    }

    Write-Warning "[Invoke-AzureADRecon] Seamless SSO / Kerberos ticket authentication failed."
    Write-Warning "[Invoke-AzureADRecon] Ensure the TGT/ccache ticket is valid, line-of-sight to the DC (port 88) exists, and SPN HTTP/autologon.microsoftazuread-sso.com is accessible."
    Write-Warning "[Invoke-AzureADRecon] Alternatively, acquire an access token manually and pass it using -AccessToken."
    return $null
}

Function Get-AADRAbout
{
<#
.SYNOPSIS
    Returns information about AzureADRecon.

.DESCRIPTION
    Returns information about AzureADRecon.

.PARAMETER Method
    [string]
    Which method to use.

.PARAMETER date
    [DateTime]
    Date

.PARAMETER AADReconVersion
    [string]
    AADRecon Version.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER RanonComputer
    [string]
    Details of the Computer running AzureADRecon.

.PARAMETER TotalTime
    [string]
    TotalTime.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $true)]
        [string] $AzureADReconVersion,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [string] $RanonComputer,

        [Parameter(Mandatory = $true)]
        [string] $TotalTime
    )

    $AboutAzureADRecon = @()

    $Version = $Method

    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    $Username = If ($mgContext.Account) { $mgContext.Account } ElseIf ($mgContext.AppName) { $mgContext.AppName }
        ElseIf ($mgContext.ClientId) { $mgContext.ClientId } Else { 'Unknown (Graph context did not identify the caller)' }

    $ObjValues = @("Date", $($date), "AzureADRecon", "https://github.com/adrecon/AzureADRecon", $Version, $($AzureADReconVersion), "Ran as user", $Username, "Ran on computer", $RanonComputer, "Execution Time (mins)", $($TotalTime))

    For ($i = 0; $i -lt $($ObjValues.Count); $i++)
    {
        $Obj = New-Object PSObject
        $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$i]
        $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i+1]
        $i++
        $AboutAzureADRecon += $Obj
    }
    Return $AboutAzureADRecon
}

Function Invoke-AzureADRecon
{
    [CmdletBinding()]
    param(
        [string] $GenExcel,
        [ValidateSet('MSGraph')] [string] $Method = 'MSGraph',
        [array] $Collect = @('Default'),
        [System.Management.Automation.PSCredential] $Credential = [System.Management.Automation.PSCredential]::Empty,
        $AccessToken, [scriptblock] $AccessTokenProvider, $AccessTokenExpiresOn,
        [ValidateRange(60,900)] [int] $TokenRefreshWindowSeconds = 300,
        [switch] $SeamlessSSO, [string] $TenantID, [string] $ClientID,
        [string] $Domain, [string] $DC, [string] $TGT,
        [array] $OutputType = @('Default'), [string] $AADROutputDir,
        [int] $DormantTimeSpan = 30, [int] $PassMaxAge = 90, [ValidateRange(1,999)] [int] $PageSize = 200, [ValidateRange(1,100)] [int] $Threads = 10,
        [string] $Logo = 'AzureADRecon', [switch] $Resume
    )

    $ErrorActionPreference = 'Stop'
    $returndir = Get-Location
    $connectionAttempted = $false
    $manifestWritable = $false
    $script:AADRRun = $null
    $script:AADRLastFailure = $null
    $script:AADRCleanupFailed = $false
    $script:AADRManifestFailureRecorded = $false
    $script:AADRManifestWritable = $false
    $script:AADRExcelFailed = $false
    $script:AADRTokenSession = $null
    $script:AADRCurrentCollector = $null
    $script:AADRPreviousCollectors = @()
    $script:AADRMFARegistrationLoaded = $false
    $script:AADRMFARegistrationRows = @()
    $ssoParameters = $null
    $stage = 'Output setup'
    Try
    {
        If ($GenExcel) { $stage='Excel generation'; Export-ADRExcel -ExcelPath $GenExcel -Logo $Logo; return }
        If ($Resume -and -Not $AADROutputDir) { throw '-Resume requires -OutputDir pointing to a previous run.' }
        $formats = @($OutputType | ForEach-Object { $_.ToUpperInvariant() })
        If ($formats -contains 'ALL') { $formats = @('CSV','XML','JSON','HTML','EXCEL') }
        If ($formats -contains 'DEFAULT') { $formats = @('CSV','EXCEL') }
        If ($formats -contains 'EXCEL' -and $formats -notcontains 'CSV') { $formats += 'CSV' }
        $formats = @($formats | Select-Object -Unique)
        If (-Not $AADROutputDir) { $AADROutputDir = Join-Path $returndir ('AzureADRecon-Report-' + (Get-Date -Format 'yyyyMMddHHmmssfff')) }
        $null = New-Item -ItemType Directory -Path $AADROutputDir -Force -ErrorAction Stop
        $AADROutputDir = (Resolve-Path -LiteralPath $AADROutputDir).Path
        ForEach ($format in $formats | Where-Object { $_ -in @('CSV','XML','JSON','HTML') })
        { $null = New-Item -ItemType Directory -Path (Join-Path $AADROutputDir ($format + '-Files')) -Force -ErrorAction Stop }

        $stage = 'Collection configuration'
        $registry = Get-AADRCollectionRegistry
        $selected = If ($Collect -contains 'Default') { @($registry.Keys) } Else { @($Collect | Select-Object -Unique) }
        ForEach ($name in $selected) { If (-Not $registry.Contains($name)) { throw "Unknown collector: $name" } }
        $selected = @($selected | ForEach-Object { $requested = $_; $registry.Keys | Where-Object { $_ -ieq $requested } | Select-Object -First 1 })
        $schemaVersion = 2
        $script:AADRRun = [pscustomobject][ordered]@{
            SchemaVersion=$schemaVersion; RunId=[guid]::NewGuid().ToString(); StartedUtc=[DateTime]::UtcNow.ToString('o'); FinishedUtc=$null
            Status='Starting'; TenantId=$null; Account=$null; ClientId=$null; AuthType=$null; ContextScopes=@()
            OutputDirectory=$AADROutputDir; OutputFormats=$formats; SelectedCollectors=$selected; Settings=@{DormantTimeSpan=$DormantTimeSpan; PassMaxAge=$PassMaxAge; PageSize=$PageSize}
            ScriptHash=(Get-FileHash -LiteralPath $script:AADRScriptPath -Algorithm SHA256).Hash
            SnapshotUtc=[DateTime]::UtcNow.ToString('o'); Collectors=(New-Object System.Collections.Generic.List[object]); Events=(New-Object System.Collections.Generic.List[object])
        }
        $priorManifestPath = Join-Path $AADROutputDir 'CollectionManifest.json'
        $prior = If ($Resume -and (Test-Path -LiteralPath $priorManifestPath)) { Get-Content -LiteralPath $priorManifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } Else { $null }
        If ($Resume -and -Not $prior) { throw 'Resume manifest was not found.' }
        If (-Not $Resume -and (Test-Path -LiteralPath $priorManifestPath))
        { throw 'This output directory already contains a run. Use -Resume or choose a new output directory.' }

        # Preserve previous run history before writing the new manifest, even if this run fails.
        If (Test-Path -LiteralPath $priorManifestPath)
        {
            $history = Join-Path $AADROutputDir '.manifests'
            $null = New-Item -ItemType Directory -Path $history -Force
            Copy-Item -LiteralPath $priorManifestPath -Destination (Join-Path $history ($script:AADRRun.RunId + '.json')) -ErrorAction Stop
        }
        # A failed resume authentication/identity check must leave the last usable manifest intact.
        $manifestWritable = -Not $Resume
        $script:AADRManifestWritable = $manifestWritable
        If ($manifestWritable) { Write-AADRManifest -Required }
        $config = @{}
        $configPath = Join-Path (Split-Path -Parent $script:AADRScriptPath) 'MSGraph-Credentials.csv'
        If (Test-Path -LiteralPath $configPath)
        {
            ForEach ($row in Import-Csv -LiteralPath $configPath -ErrorAction Stop)
            { If (-Not [string]::IsNullOrWhiteSpace($row.Value)) { $config[$row.Category] = $row.Value } }
        }
        ForEach ($key in @('TenantID','ClientID','Domain','DC','TGT'))
        {
            $explicit = Get-Variable -Name $key -ValueOnly
            If ($explicit) { $config[$key] = $explicit }
        }
        If (-Not $AccessToken -and $config.AccessToken) { $AccessToken = $config.AccessToken }
        If (-Not $SeamlessSSO -and $config.SeamlessSSO -in @('true','1')) { $SeamlessSSO = $true }
        $stage = 'Authentication'
        If (-Not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) { throw 'Install Microsoft.Graph before running AzureADRecon.' }
        $script:AADRParserType = Initialize-AADRParserType -Source $MSGraphSource
        $connectionAttempted = $true
        If ($SeamlessSSO)
        {
            If ($AccessToken) { throw 'Choose either AccessToken or SeamlessSSO.' }
            If ($AccessTokenProvider) { throw 'Choose either AccessTokenProvider or SeamlessSSO (which supplies its own provider).' }
            $ssoParameters = @{TenantID=$config.TenantID; Domain=$config.Domain; DC=$config.DC; TGT=$config.TGT; IncludeMetadata=$true}
            $AccessTokenProvider = {
                param($request)
                $parameters = $script:AADRTokenSession.ProviderParameters
                Get-AADRSeamlessSSOToken @parameters -ForceRefresh:$request.ForceRefresh
            }
        }
        If ($AccessToken -or $AccessTokenProvider)
        {
            $mode = If ($SeamlessSSO) { 'SeamlessSSO' } Else { 'Provider' }
            Initialize-AADRTokenSession -Token $AccessToken -Provider $AccessTokenProvider -ExpiresOn $AccessTokenExpiresOn -WindowSeconds $TokenRefreshWindowSeconds -Mode $mode -ProviderParameters $ssoParameters
        }
        Else
        {
            $scopes = @('AuditLog.Read.All','User.Read.All','UserAuthenticationMethod.Read.All','Policy.Read.All','Directory.Read.All')
            If ($selected -contains 'PIMEligibility') { $scopes += 'RoleEligibilitySchedule.Read.Directory' }
            $parameters = Get-AADRGraphConnectionParameters -Config $config -Credential $Credential -Scopes $scopes
            Connect-MgGraph @parameters | Out-Null
        }
        # Bind every run/checkpoint to the actual tenant returned by Graph, not a CLI alias.
        $organization = @(Get-AADRGraphPages 'organization?$select=id')
        If ($organization.Count -ne 1 -or -Not $organization[0].id) { throw 'Could not verify the authenticated tenant.' }
        $context = Get-MgContext -ErrorAction Stop
        $script:AADRRun.TenantId = [string]$organization[0].id
        $accountIdentity = $context.Account
        If (-Not $accountIdentity -and $script:AADRTokenSession -and $script:AADRTokenSession.Identity) { $accountIdentity = $script:AADRTokenSession.Identity.UserId }
        $script:AADRRun.Account = $accountIdentity
        $script:AADRRun.ClientId = $context.ClientId
        $script:AADRRun.AuthType = $context.AuthType
        $script:AADRRun.ContextScopes = @($context.Scopes | Where-Object { $null -ne $_ })
        $requestedTenantGuid = [guid]::Empty
        If ($config.TenantID -and [guid]::TryParse($config.TenantID, [ref]$requestedTenantGuid) -and [string]$requestedTenantGuid -ne $script:AADRRun.TenantId)
        { throw 'The authenticated token belongs to a different tenant than the requested TenantID.' }
        If ($Resume)
        {
            $stage = 'Resume validation'
            If (-Not $accountIdentity -and -Not $context.ClientId) { throw 'Resume requires a verified account or application. Use a delegated token provider, interactive authentication, or application credentials.' }
            $selectionChanged = (@($prior.SelectedCollectors | Sort-Object -Unique) -join ',') -ne (@($selected | Sort-Object -Unique) -join ',')
            If ($selectionChanged -or $prior.SchemaVersion -ne $schemaVersion -or $prior.TenantId -ne $script:AADRRun.TenantId -or
                $prior.ScriptHash -ne $script:AADRRun.ScriptHash -or $prior.ClientId -ne $context.ClientId -or $prior.Account -ne $accountIdentity -or
                $prior.Settings.DormantTimeSpan -ne $DormantTimeSpan -or $prior.Settings.PassMaxAge -ne $PassMaxAge -or $prior.Settings.PageSize -ne $PageSize)
            { throw 'Resume identity, script version, or collection settings differ from the saved run. Start a new output directory.' }
            $script:AADRPreviousCollectors = @($prior.Collectors)
            $script:AADRRun.SnapshotUtc = $prior.SnapshotUtc
        }
        $date = [datetime]$script:AADRRun.SnapshotUtc
        $script:AADRRun.Status = 'Running'
        $manifestWritable = $true
        $script:AADRManifestWritable = $true
        Write-AADRManifest -Required
        $stage = 'Collection'
        ForEach ($name in $selected)
        {
            Write-Host "[-] $name"
            Invoke-AADRCollectorRun -Name $name -Query $registry[$name] -OutputDirectory $AADROutputDir -Formats $formats -Resume:$Resume
        }
        $totalTime = '{0:N2}' -f (([DateTime]::UtcNow - [datetime]$script:AADRRun.StartedUtc).TotalMinutes)
        $stage = 'Report generation'
        $about = @(Invoke-AADRReportStep -Name 'Run attribution' -Query { Get-AADRAbout -Method MSGraph -date $date -AzureADReconVersion '2.0' -RanonComputer ([Environment]::MachineName) -TotalTime $totalTime })
        If ($about.Count -gt 0)
        {
            ForEach ($format in $formats | Where-Object { $_ -ne 'EXCEL' })
            { Invoke-AADRReportStep -Name ('About ' + $format) -Query { Export-ADR -ADRObj $about -AADROutputDir $AADROutputDir -OutputType @($format) -ADRModuleName AboutAzureADRecon } }
        }
        If ($formats -contains 'HTML') { Invoke-AADRReportStep -Name 'HTML index' -Query { Export-ADRHTML -ADRObj ([pscustomobject]@{}) -ADFileName (Join-AADRPath @($AADROutputDir,'HTML-Files','Index.html')) -AADROutputDir $AADROutputDir } }
        If ($formats -contains 'EXCEL')
        {
            Invoke-AADRReportStep -Name Excel -Query { Export-ADRExcel -ExcelPath $AADROutputDir -Logo $Logo }
        }
        $script:AADRRun.Status = If (@($script:AADRRun.Collectors | Where-Object Status -notin @('Success','Empty')).Count -gt 0 -or @($script:AADRRun.Events | Where-Object Status -in @('ExportFailed','ManifestFailed','CheckpointFailed','CleanupFailed')).Count -gt 0) { 'Partial' } Else { 'Success' }
        Write-Host "Output directory: $AADROutputDir"
    }
    Catch
    {
        $diagnostic = Get-AADRErrorDiagnostic $_
        $script:AADRLastFailure = [pscustomobject]@{Stage=$stage; Diagnostic=$diagnostic}
        If ($script:AADRRun) { $script:AADRRun.Status = 'Failed'; Add-AADREvent -Operation Run -Status Failed -Detail ("Run stopped during $stage. Completed collector checkpoints are retained.") -Diagnostic $diagnostic }
        If (Test-AADRGraphAuthenticationError -Record $_) { Stop-AADRGraphAuthentication -Record $_ }
        throw
    }
    Finally
    {
        If ($connectionAttempted) { Disconnect-AADRGraphSession }
        Try { Set-Location -LiteralPath $returndir -ErrorAction Stop }
        Catch { $script:AADRCleanupFailed=$true; Add-AADREvent -Operation Cleanup -Status CleanupFailed -Detail 'Could not restore the working directory.' -Diagnostic (Get-AADRErrorDiagnostic $_); Write-Warning '[Cleanup] Could not restore the working directory.' -WarningAction Continue }
        Try
        {
        If ($script:AADRRun -and $manifestWritable)
        {
            ForEach ($name in $selected)
            {
                If (@($script:AADRRun.Collectors | Where-Object Name -eq $name).Count -eq 0)
                { $script:AADRRun.Collectors.Add([pscustomobject]@{Name=$name;Status='NotStarted';DataStatus='NotStarted';RowCount=$null;StartedUtc=$null;FinishedUtc=$null;Checkpoint=$null;CheckpointHash=$null;Resumed=$false;Exports=@();Error=$null}) }
            }
            $script:AADRRun.FinishedUtc = [DateTime]::UtcNow.ToString('o')
            If ($script:AADRRun.Status -eq 'Success' -and @($script:AADRRun.Events | Where-Object Status -in @('CleanupFailed','ManifestFailed')).Count -gt 0) { $script:AADRRun.Status = 'Partial' }
            Write-AADRManifest
            # The final write itself can be the first I/O failure.
            If ($script:AADRRun.Status -eq 'Success' -and $script:AADRManifestFailureRecorded) { $script:AADRRun.Status = 'Partial' }
        }
        }
        Finally
        {
            $script:AADRCurrentCollector = $null
            $script:AADRTokenSession = $null
            Write-Progress -Activity 'AzureADRecon' -Completed -ErrorAction SilentlyContinue
        }
    }
}

$transcribing = $false
$transcriptStartFailure = $null
$script:AADRExitCode = 0
Try
{
    If ($Log)
    {
        Try
        {
            Start-Transcript -Path (Join-Path (Get-Location) 'AzureADRecon-Console-Log.txt') -ErrorAction Stop | Out-Null
            $transcribing = $true
        }
        Catch
        {
            If (Test-AADRControlFlowError $_) { throw }
            $transcriptStartFailure = Get-AADRErrorDiagnostic $_
            Write-Warning '[Transcript] Console logging could not start. Collection will continue; check the log path and file permissions.' -WarningAction Continue
        }
    }
    Invoke-AzureADRecon -GenExcel $GenExcel -Method $Method -Collect $Collect -Credential $Credential -AccessToken $AccessToken -AccessTokenProvider $AccessTokenProvider -AccessTokenExpiresOn $AccessTokenExpiresOn -TokenRefreshWindowSeconds $TokenRefreshWindowSeconds -SeamlessSSO:$SeamlessSSO -TenantID $TenantID -ClientID $ClientID -Domain $Domain -DC $DC -TGT $TGT -Resume:$Resume -OutputType $OutputType -AADROutputDir $OutputDir -DormantTimeSpan $DormantTimeSpan -PassMaxAge $PassMaxAge -PageSize $PageSize -Threads $Threads -Logo $Logo
}
Catch
{
    If (Test-AADRControlFlowError $_) { throw }
    $script:AADRExitCode = 1
    $failure = $script:AADRLastFailure
    $diagnostic = If ($failure) { $failure.Diagnostic } Else { Get-AADRErrorDiagnostic $_ }
    $stage = If ($failure) { $failure.Stage } Else { 'Startup' }
    $check = If ($diagnostic.Operation) { $diagnostic.Operation } Else { $stage }
    $status = If ($diagnostic.HttpStatus) { 'HTTP ' + $diagnostic.HttpStatus } Else { 'no HTTP status' }
    Write-Warning "[AzureADRecon] Stopped during $stage at $check ($($diagnostic.Category), $($diagnostic.Code), $($diagnostic.Reason), $status). $($diagnostic.Guidance) Check CollectionManifest.json when available. Completed report files are retained." -WarningAction Continue
}
Finally
{
    If ($transcriptStartFailure)
    { Add-AADREvent -Operation Transcript -Status LoggingFailed -Detail 'Console transcript could not start; collection proceeded without it.' -Diagnostic $transcriptStartFailure }
    If ($transcribing)
    {
        Try { Stop-Transcript -ErrorAction Stop | Out-Null }
        Catch
        {
            $script:AADRCleanupFailed = $true
            Add-AADREvent -Operation Transcript -Status CleanupFailed -Detail 'Console transcript could not be stopped.' -Diagnostic (Get-AADRErrorDiagnostic $_)
            Write-Warning '[Transcript] Console logging could not be stopped cleanly.' -WarningAction Continue
        }
    }
    If ($script:AADRRun -and $script:AADRManifestWritable)
    {
        If ($script:AADRRun.Status -eq 'Success' -and @($script:AADRRun.Events | Where-Object Status -in @('LoggingFailed','CleanupFailed','ManifestFailed')).Count -gt 0) { $script:AADRRun.Status = 'Partial' }
        Write-AADRManifest
        If ($script:AADRRun.Status -eq 'Success' -and $script:AADRManifestFailureRecorded) { $script:AADRRun.Status = 'Partial' }
    }
    If ($script:AADRExitCode -eq 0 -and (($script:AADRRun -and $script:AADRRun.Status -eq 'Partial') -or $script:AADRExcelFailed -or $script:AADRCleanupFailed -or $transcriptStartFailure)) { $script:AADRExitCode = 2 }
}
# Keep interactive dot-sourcing usable; CLI and & invocation expose a meaningful exit code.
If ($MyInvocation.InvocationName -ne '.') { exit $script:AADRExitCode }
