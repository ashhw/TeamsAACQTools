# TeamsAACQTools PowerShell Module

Streamline the process of building Auto Attendants, Call Queues, and Resource Accounts in Microsoft Teams with the TeamsAACQTools PowerShell Module. This module is designed to alleviate the complexities of deploying these components, particularly during Skype for Business migrations when existing data is available. The module facilitates the conversion of the "RGSConfig.zip" file from Skype for Business into an Excel workbook that can be seamlessly used to configure these elements in Teams.

## Table of Contents
- [Pre-Requirements](#pre-requirements)
- [Module Installation](#module-installation)
- [New Features/Fixes](#new-featuresfixes)
- [Usage Instructions](#usage-instructions)
  - [Export the RGS Configuration](#export-the-rgs-configuration)
  - [Build the Excel Workbook](#build-the-excel-workbook)
  - [Build Call Queues in Teams](#build-call-queues-in-teams)
  - [Build Auto Attendants in Teams](#build-auto-attendants-in-teams)
- [Functionality](#functionality)
  - [Import-NasAACQData Function](#import-nasaacqdata-function)
  - [Import-NasCQ Function](#import-nascq-function)
  - [Import-NasAA Function](#import-nasaa-function)
- [Parameters](#parameters)

## Pre-Requirements
Ensure the following prerequisites are met before using the TeamsAACQTools PowerShell Module:
1. Install the ImportExcel module: Run as Administrator in PowerShell: `Install-Module ImportExcel`
2. Install the MicrosoftTeams Module: Run as Administrator in PowerShell: `Install-Module MicrosoftTeams`
3. Download ffmpeg: Choose your OS and download from [here](https://www.ffmpeg.org/download.html). Copy it to a chosen location, e.g., "C:\ffmpeg".

**Note:** You can also use the `-InstallModules` switch for an interactive installation of required modules. Avoid using this in automation scenarios.

## Module Installation
1. Download the module from the GitHub repository.
2. Copy the TeamsAACQTools folder into your PowerShell module repository on your local workstation. To find your local module repository, type `Get-Module -ListAvailable | Select-Object -ExpandProperty ModuleBase` in a PowerShell window.
3. Open PowerShell and run: `Import-Module TeamsAACQTools`.

## New Features/Fixes
**Call Queues 24/08/23**
- Fixed issue with agents not being added to call queues.

**Auto Attendants 01/08/22**
- Added support for greeting audio files and non-business hours audio files.
- Enhanced action target logic to detect endpoint type.
- Introduced the required `sharedvoicemail` parameter (MSFT change).
- Resolved greetings error related to object type.
- Updated `identity` to `identities` for resource account (MSFT change).
- Expanded Excel export with additional fields.
- Implemented logging with specified log folder.

## Export the RGS Configuration

To export the Response Group configuration from Skype for Business, use the following command in the Skype for Business Management Shell:

```powershell
Export-CsRgsConfiguration -Source "ApplicationServer:my-sfb-fe01.somedomain.com" -FileName "C:\Exports\Rgs.zip"
```

Please adjust the command parameters according to your specific Skype environment.

## Build the Excel Workbook

Once you have exported the RGSConfig.zip (ensure it's unzipped to your desired location) from Skype for Business, you can construct the Excel workbook using this example command:

```powershell
Import-NasAACQData -rootFolder "C:\AACQ\RgsImportExport" -ffmpeglocation "C:\ffmpeg\bin\ffmpeg.exe" -TenantDomain "mytenant.onmicrosoft.com" -CQRAPrefix "ra_cq_" -AARAPrefix "ra_aa_" -cqReplacementSuffix "CQ" -aaReplacementSuffix "AA"  -logFolder "C:\path\to\log" -Verbose
```

**Important Notes**:
- Agent Alert Time will be rounded up to the next multiple of 15 (Teams only allows multiples of 15).
- If OverflowThreshold isn't specified but OverflowAction is, the default OverflowThreshold will be set to 1.
- Agents will be listed as a SIP address without the "sip:" prefix, separated by commas.
- If Language isn't specified, the default Language will be en-GB.

## Build Call Queues in Teams

After creating the AACQDataImport.xlsx file and verifying its accuracy, you can execute the following command to create Call Queues in Teams, along with associated Resource Accounts and Resource Account Associations:

```powershell
Import-NasCQ -CQData "C:\AACQ\RgsImportExport\AACQDataImport-TestLab.xlsx" -logFolder "C:\path\to\log" -Verbose
```

By default, this command prompts you to back up the Call Queues before proceeding. You can skip this prompt by using the `-NoBackup` switch.

## Build Auto Attendants in Teams

Once the Call Queues are established in Teams, you can use the Auto Attendant function to create Auto Attendants and associate them with the Call Queues (if needed), along with Resource Accounts and Resource Account Associations:

```powershell
Import-NasAA -AAData "C:\AACQ\RgsImportExport\AACQDataImport-TestLab.xlsx" -logFolder "C:\path\to\log" -Verbose
```

By default, this command prompts you to back up the Auto Attendants before proceeding. You can bypass this prompt using the `-NoBackup` switch.

**Additional Information**:
- Business Hours will be rounded to the nearest multiple of 15. If this exceeds 00:00, the Auto Attendant is assumed to be open 24 hours.
- If Business Hours aren't specified for every day, it's assumed that the Auto Attendant is open 24/7.
- If Business Hours aren't specified for particular days, the Auto Attendant is assumed to be open every day except for those with no specified times.

## Functionality ##

## Import-NasAACQData Function

This function allows you to import AACQ (Auto Attendant and Call Queue) data. It facilitates the conversion and configuration of data from Skype for Business Response Group export into a format suitable for Teams.

### Parameters

**-rootFolder**  
Path to the RgsImportExport folder from the Skype for Business RGS export, e.g., "C:\Folder\SomeFolder\RgsImportExport".

**-Interactive**  
Use this parameter when there's a duplicate Music on Hold entry in the Response Group configuration. It prompts you to choose the correct Music on Hold or LanguageID entry.

**-CQRAPrefix (OPTIONAL)**  
Prefix for the call queue resource account, e.g., racq_.

**-AARAPrefix (OPTIONAL)**  
Prefix for the Auto Attendant resource account, e.g., raaa_.

**-cqReplacementSuffix (OPTIONAL)**  
Suffix for the Call Queue name, e.g., "Queue".

**-aaReplacementSuffix (OPTIONAL)**  
Suffix for the Auto Attendant name, e.g., "Attendant".

**-TenantDomain**  
Tenant domain used to build the userprincipalname for Resource Accounts, e.g., "@testlab.onmicrosoft.com".

**-ffmpeglocation**  
Path to ffmpeg.exe for audio file conversion to .mp3.

**-SkipAudio [OPTIONAL]**  
Use this to skip audio conversion, for manual or new audio file usage.

**-logFolder [MANDATORY]**  
Location to output the log transcript. Mandatory for reviewing build process details.

## Import-NasCQ Function

This function is for building Teams Call Queues along with associated Resource Accounts and Account Associations.

### Parameters

**-CQData**  
Excel workbook specifying Call Queue data and configuration.

**-InstallModules**  
Prompt to install required modules if needed.

**-NoCreateCQ**  
Exclude call queue creation.

**-NoRA**  
Exclude resource account creation.

**-NoBackup**  
Exclude backup process. Useful for new setups without existing Call Queues.

**-logFolder [MANDATORY]**  
Location to output the log transcript. Mandatory for reviewing build process details.

## Import-NasAA Function

This function is for building Teams Auto Attendants along with associated Resource Accounts and Account Associations.

### Parameters

**-AAData**  
Excel workbook specifying Auto Attendant data and configuration.

**-InstallModules**  
Prompt to install required modules if needed.

**-NoCreateAA**  
Exclude Auto Attendant creation.

**-NoRA**  
Exclude resource account creation.

**-NoBackup**  
Exclude backup process. Useful for new setups without existing Auto Attendants.

**-logFolder [MANDATORY]**  
Location to output the log transcript. Mandatory for reviewing build process details.
