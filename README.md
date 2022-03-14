# NasAACQTools

Building out several Auto Attendants, Call Queues and Resource Accounts in Microsoft Teams is painstakingly slow, especially if the deployment is a Skype for Business migration and you already have the data required. This module will help build out the Auto Attendants, Call Queues and Resource accounts from an Excel workbook.

The module can be used to take the "RGSConfig.zip" file from Skype for Business and convert it into an Excel workbook that can be used to build the configuration in Teams.

## Syntax
### Import-NasAACQData
Build Excel workbook by specifying the Resource Account UPN prefix
```powershell
Import-NasAACQData [[-rootFolder] <string>] [[-ffmpeglocation] <string>] [[-TenantDomain] <string>] [[-AARAPrefix] <string>] [[-CQRAPrefix] <string>]
Import-NasAACQData -rootFolder <"Location of Excel Workbook"> -ffmpeglocation <"Location of ffmpeg.exe"> -TenantDomain <"Tenant Domain"> -CQRAPrefix racq_ -AARAPrefix raaa_
```
#### Parameters
-Interactive
If there is a duplicate Music on Hold entry in the Skype for Business Response Group configuration, i.e. Two workflows target the same call queue (therefore using call queue Music on Hold). Use -Interactive parameter to be prompted with an interface to choose the correct Music on Hold to use.

-CQRAPrefix
Specify the prefix of the call queue resource account. i.e. racq_

-AARAPrefix
Specify the prefix of the call queue resource account. i.e. raaa_

-TenantDomain
Specify the tenant domain, this is used to build the userprincipalname for the Resource Accounts. i.e. "@testlab.onmicrosoft.com"

-ffmpeglocation
Specify the ffmpeg.exe path for the audio file conversion to .mp3.

-SkipAudio
Specify this to skip the audio conversion as you would like to do this manually or will be using new audio files.

### Import-NasCQ
Create the resource accounts, build the call queues and associate the resource accounts
```powershell
Import-NasCQ [-CQData] <String> [-InstallModules] [-NoCreateCQ] [-NoRA]
Import-NasCQ -CQData <"Path to Excel Workbook">
```
#### Parameters
-CQData
Must be specified as the Excel workbook used to build the call queues and resource accounts.

-InstallModules
Specify if you wish to be prompted to install required modules

-NoCreateCQ
Specify if you wish to exclude call queue creation

-NoRA
Specify if you wish to exclude resource account creation.

