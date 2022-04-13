# NasAACQTools

Building out several Auto Attendants, Call Queues and Resource Accounts in Microsoft Teams is painstakingly slow, especially if the deployment is a Skype for Business migration and you already have the data required. This module will help build out the Auto Attendants, Call Queues and Resource accounts from an Excel workbook.

The module can be used to take the "RGSConfig.zip" file from Skype for Business and convert it into an Excel workbook that can be used to build the configuration in Teams.

NOTE: Only supports Powershell 7.0 and above.
## Pre-Requirements
1. Install the ImportExcel module - Run as Administrator in PowerShell: Install-Module ImportExcel
2. Install the MicrosoftTeams Module - Run as Administrator in PowerShell: Install-Module MicrosoftTeams
3. Download ffmpeg - Choose your OS and download from here: https://www.ffmpeg.org/download.html
    a. Copy this to your chosen location, example "C:\ffmpeg"

NOTE: You can also specify the -InstallModules switch to provide an interactive approach to install the required modules. Do not use if you are using as part of an automation job.
## Module Installation
1. Download the module from the github repository.
2. Copy the TeamsAACQTools folder into your PowerShell module repository on your local workstation.
3. Open PowerShell and run "Import-Module TeamsAACQTools".

**Export the RGS configuration**
To export the Response Group configuration from Skype for Business, run the following command on the Skype for Business Management Shell:
Example: "Export-CsRgsConfiguration -Source "ApplicationServer:my-sfb-fe01.somedomain.com" -FileName "C:\Exports\Rgs.zip""
Note: Amend the command to suit the Skype environment.

**Build the Excel workbook**
Once you have the RGSConfig.zip (Unzip this to your chosen location) exported from Skype for Business, you can build the Excel workbook using the following example command:
Import-NasAACQData -rootFolder "C:\AACQ\RgsImportExport" -ffmpeglocation "C:\ffmpeg\bin\ffmpeg.exe" -TenantDomain "mytenant.onmicrosoft.com" -CQRAPrefix "ra_cq_" -AARAPrefix "ra_aa_" -cqReplacementSuffix "CQ" -aaReplacementSuffix "AA" -Verbose

**Build the Call Queues in Teams**
Once you have built the AACQDataImport.xlsx file and verified that the data is correct, you can run it against Teams to build the Call Queues, along with Resource Accounts and Resource Account Association:
Import-NasCQ -CQData "C:\AACQ\RgsImportExport\AACQDataImport-TestLab.xlsx"

**Build the Auto Attendants in Teams**
Once you have built the Call Queues in Teams, you can run the Auto Attendant function to build the Auto Attendants and associate with the Call Queues (If reuqired), along with Resource Accounts and Resource Account Association:
Import-NasAA -AAData "C:\AACQ\RgsImportExport\AACQDataImport-TestLab.xlsx"

## Import-NasAACQData
Build Excel workbook by specifying the Resource Account UPN prefix
```powershell
Import-NasAACQData [[-rootFolder] <string>] [[-ffmpeglocation] <string>] [[-TenantDomain] <string>] [[-AARAPrefix] <string>] [[-CQRAPrefix] <string>]
Import-NasAACQData -rootFolder <"Location of RgsImportExport folder"> -ffmpeglocation <"Location of ffmpeg.exe"> -TenantDomain <"Tenant Domain"> -CQRAPrefix racq_ -AARAPrefix raaa_
```
### Parameters
**-rootFolder**  
Specify the path to the RgsImportExport folder from the Skype for Business RGS export.
e.g. "C:\Folder\SomeFolder\RgsImportExport"

**-Interactive**  
If there is a duplicate Music on Hold entry in the Skype for Business Response Group configuration, i.e. Two workflows target the same call queue (therefore using call queue Music on Hold). Use -Interactive parameter to be prompted with an interface to choose the correct Music on Hold to use. This also applies to the LanguageID.

**-CQRAPrefix (OPTIONAL)**  
Specify the prefix of the call queue resource account. i.e. racq_

**-AARAPrefix (OPTIONAL)**  
Specify the prefix of the call queue resource account. i.e. raaa_

**-cqReplacementSuffix (OPTIONAL)**  
Specify the suffix for the Call Queue name. i.e. "Queue"

**-aaReplacementSuffix (OPTIONAL)**  
Specify the suffix for the Auto Attendant name. i.e. "Attendant"

**-TenantDomain**  
Specify the tenant domain, this is used to build the userprincipalname for the Resource Accounts. i.e. "@testlab.onmicrosoft.com"

**-ffmpeglocation**  
Specify the ffmpeg.exe path for the audio file conversion to .mp3.

**-SkipAudio [OPTIONAL]**  
Specify this to skip the audio conversion as you would like to do this manually or will be using new audio files.

## Import-NasCQ
Create the resource accounts, build the Call Queues and associate the resource accounts with the Call Queues.
(Note: By default, you will be prompted to backup any existing Call Queues, to skip this step, use the -NoBackup switch.)
```powershell
Import-NasCQ [-CQData] <String> [-InstallModules] [-NoCreateCQ] [-NoRA]
Import-NasCQ -CQData <"Path to Excel Workbook">
```
### Parameters
**-CQData**  
Must be specified as the Excel workbook used to build the call queues and resource accounts.

**-InstallModules**  
Specify if you wish to be prompted to install required modules

**-NoCreateCQ**  
Specify if you wish to exclude call queue creation

**-NoRA**  
Specify if you wish to exclude resource account creation.

**-NoBackup**  
Specify if you wish to exclude the backup process. (Useful for Greenfield sites without any existing Call Queues)

## Import-NasAA
Create the resource accounts, build the Auto Attendants and associate the resource accounts with the Auto Attendants.
(Note: By default, you will be prompted to backup any existing Auto Attendants, to skip this step, use the -NoBackup switch.)
```powershell
Import-NasAA [-AAData] <String> [-InstallModules] [-NoCreateAA] [-NoRA]
Import-NasCQ -AAData <"Path to Excel Workbook">
```
### Parameters
**-AAData**  
Must be specified as the Excel workbook used to build the Auto Attendants and resource accounts.

**-InstallModules**  
Specify if you wish to be prompted to install required modules

**-NoCreateAA**  
Specify if you wish to exclude Auto Attendant creation

**-NoRA**  
Specify if you wish to exclude resource account creation.

**-NoBackup**  
Specify if you wish to exclude the backup process. (Useful for Greenfield sites without any existing Auto Attendants)

