---
external help file: TeamsAACQTools-help.xml
Module Name: TeamsAACQTools
online version:
schema: 2.0.0
---

# Import-NasAACQData

## SYNOPSIS
Used to Build the Excel workbook from the RGS Configuration from Skype for Business.

## SYNTAX

```
Import-NasAACQData [[-rootFolder] <String>] [-interactive] [[-ffmpeglocation] <String>]
 [[-aaRAPrefix] <String>] [[-cqRAPrefix] <String>] [[-tenantDomain] <String>] [-skipAudio]
 [[-cqNamePrefix] <String>] [[-cqReplacementSuffix] <String>] [[-aaNamePrefix] <String>]
 [[-aaReplacementSuffix] <String>] [<CommonParameters>]
```

## DESCRIPTION
Once you have the RGSConfig.zip (Unzip this to your chosen location) exported from Skype for Business, you can build the Excel workbook using the example command.

## EXAMPLES

### EXAMPLE 1
```
Import-NasAACQData -rootFolder "C:\AACQ\RgsImportExport" -ffmpeglocation "C:\ffmpeg\bin\ffmpeg.exe" -TenantDomain "mytenant.onmicrosoft.com" -CQRAPrefix "ra_cq_" -AARAPrefix "ra_aa_" -cqReplacementSuffix "CQ" -aaReplacementSuffix "AA" -Verbose
```

## PARAMETERS

### -rootFolder
{{ Fill rootFolder Description }}

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

### -interactive
{{ Fill interactive Description }}

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

### -ffmpeglocation
{{ Fill ffmpeglocation Description }}

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

### -aaRAPrefix
{{ Fill aaRAPrefix Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -cqRAPrefix
{{ Fill cqRAPrefix Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -tenantDomain
{{ Fill tenantDomain Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -skipAudio
{{ Fill skipAudio Description }}

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

### -cqNamePrefix
{{ Fill cqNamePrefix Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -cqReplacementSuffix
{{ Fill cqReplacementSuffix Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -aaNamePrefix
{{ Fill aaNamePrefix Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -aaReplacementSuffix
{{ Fill aaReplacementSuffix Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
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
