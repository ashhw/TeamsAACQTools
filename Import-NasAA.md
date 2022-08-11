---
external help file: TeamsAACQTools-help.xml
Module Name: TeamsAACQTools
online version:
schema: 2.0.0
---

# Import-NasAA

## SYNOPSIS
Used to build the Auto Attendants in Teams by specifying the Excel workbook.

## SYNTAX

```
Import-NasAA [-AAData] <String> [-InstallModules] [-NoCreateAA] [-NoRA] [-Wireframe] [-NoBackup]
 [[-logFolder] <String>] [<CommonParameters>]
```

## DESCRIPTION
Once you have built the Call Queues in Teams, you can run the Auto Attendant function to build the Auto Attendants and associate with the Call Queues (If reuqired), along with Resource Accounts and Resource Account Association.

## EXAMPLES

### EXAMPLE 1
```
Import-NasAA -AAData "C:\AACQ\RgsImportExport\AACQDataImport-TestLab.xlsx"
```

## PARAMETERS

### -AAData
Auto Attendant data import

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
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

### -NoCreateAA
Don't createthe Auto Attendants

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
Don't create the resource accounts

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

### -Wireframe
Creates the basics for the auto attendant, RA, linking the call queue etc

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
No backup, used for a greenfield environment and automation

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

### -logFolder
Root folder used for log file purposes

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
