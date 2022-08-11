---
external help file: TeamsAACQTools-help.xml
Module Name: TeamsAACQTools
online version:
schema: 2.0.0
---

# Remove-StringSpecialCharacter

## SYNOPSIS
This function will remove the special character from a string.

## SYNTAX

```
Remove-StringSpecialCharacter [[-String] <String[]>] [[-SpecialCharacterToKeep] <String[]>]
 [<CommonParameters>]
```

## DESCRIPTION
This function will remove the special character from a string.
I'm using Unicode Regular Expressions with the following categories
\p{L} : any kind of letter from any language.
\p{Nd} : a digit zero through nine in any script except ideographic
http://www.regular-expressions.info/unicode.html
http://unicode.org/reports/tr18/

## EXAMPLES

### EXAMPLE 1
```
Remove-StringSpecialCharacter -String "^&*@wow*(&(*&@"
wow
```

### EXAMPLE 2
```
Remove-StringSpecialCharacter -String "wow#@!`~)(\|?/}{-_=+*"
wow
```

### EXAMPLE 3
```
Remove-StringSpecialCharacter -String "wow#@!`~)(\|?/}{-_=+*" -SpecialCharacterToKeep "*","_","-"
wow-_*
```

## PARAMETERS

### -String
Specifies the String on which the special character will be removed

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: Text

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -SpecialCharacterToKeep
Specifies the special character to keep in the output

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: Keep

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
Francois-Xavier Cat
@lazywinadmin
lazywinadmin.com
github.com/lazywinadmin

## RELATED LINKS
