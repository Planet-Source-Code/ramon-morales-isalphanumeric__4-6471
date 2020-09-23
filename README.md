<div align="center">

## isAlphaNumeric


</div>

### Description

Function to determine if the passed parameter is AlphaNumeric (the string contains only A-Z, a-z or 1-0). Heavily commented. Bug Fixed.
 
### More Info
 
The string to be evaluated.

See comments.

Returns True is the string is AlphaNumeric and False if it is not.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ramon Morales](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ramon-morales.md)
**Level**          |Beginner
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Validation/ Processing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/validation-processing__4-16.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ramon-morales-isalphanumeric__4-6471/archive/master.zip)

### API Declarations

Public Domain


### Source Code

```
	Function IsAlphaNumeric(sText)
	'***************************************************************************
	'Checks to see if sText is made up of only Alphabetic characters (A-Z) or
	'Numbers. If it has any other characters, IsAlphaNumeric will be False.
	'***************************************************************************
		Dim nLen, nLoop, sTemp,	sSingleCharacter
		Dim bAlphaStatus
	'***************************************************************************
	'Default value
	'***************************************************************************
		bAlphaStatus = True
	'***************************************************************************
	'Gets length of the sText variable.
	'***************************************************************************
		sTemp = Trim(sText)
		nLen = Len(sTemp)
	'***************************************************************************
	'If the length of sText is 0, then it is not AlphaNumeric and
	'IsAlphaNumeric = False.
	'***************************************************************************
		If nLen = 0 then
			bAlphaStatus = False
		End If	'If nLen = 0 then
		If nLen > 0 then
	'***************************************************************************
	'Convert sText to uppercase to make comparisons easier.
	'***************************************************************************
			sTemp = Ucase(sTemp)
	'***************************************************************************
	'Will loop nLen times. It will check each of the characters of sText against
	'the comparison string (which is A-Z and 1-0). It will check it one
	'character at a time (beginning with the farthest left character).If the
	'Instr command shows a 0 (meaning it could not find a match), that character
	'was not AlphaNumeric.
	'***************************************************************************
			For nLoop =1 to nLen
				sSingleCharacter = Mid(sTemp,nLoop,1)
				If Instr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", sSingleCharacter)= 0 then
					bAlphaStatus = False
					Exit For
				End If
			Next
	'***************************************************************************
	'If sText managed to get through the above filters without changing
	'IsAlphaNumeric to False, then IsAlphaNumeric is True.
	'***************************************************************************
			If bAlphaStatus <> False then
				bAlphaStatus = True
			End If
		End If	'If nLen > 0 then
		IsAlphaNumeric = bAlphaStatus
	End Function
```

