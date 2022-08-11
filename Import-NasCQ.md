---
external help file: TeamsAACQTools-help.xml
Module Name: TeamsAACQTools
online version:
schema: 2.0.0
---

# Import-NasCQ

## SYNOPSIS
Used to build the Call Queues in Teams from the Excel workbook.

## SYNTAX

```
Import-NasCQ [[-TenantDomain] <String>] [-CQData] <String> [-InstallModules] [-NoCreateCQ] [-NoRA] [-NoBackup]
 [<CommonParameters>]
```

## DESCRIPTION
Once you have built the AACQDataImport.xlsx file and verified that the data is correct, you can run it against Teams to build the Call Queues, along with Resource Accounts and Resource Account Association.

## EXAMPLES

### EXAMPLE 1
```
Import-NasCQ -CQData "C:\AACQ\RgsImportExport\AACQDataImport-TestLab.xlsx"
```

## PARAMETERS

### -TenantDomain
Tenant domain in the following format: \<"YourTenant".onmicrosoft.com\>

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CQData
Call Queue data import

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -InstallModules
Give the user the option to install the required modules

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCreateCQ
Create the call queues with no prompts

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRA
{{ Fill NoRA Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoBackup
{{ Fill NoBackup Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
