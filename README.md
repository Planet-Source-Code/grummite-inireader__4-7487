<div align="center">

## INIreader


</div>

### Description

This is a VBS script that will look for a item in a INI file. Comes in handy when you need to read INI files created by other applications. It's Read-only, no writing to INI files done here.
 
### More Info
 
Changes need made in the

'code for the INI filename

'and INI item parameters.

'Look for the Code between

'the lines of ''''''s.

It returns a message box

'with the string being

'search for.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Grummite](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/grummite.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VbScript \(browser/client side\)

**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/grummite-inireader__4-7487/archive/master.zip)





### Source Code

```
'Comments/Questions:
' Email me at VBS@grummite.com
'''''''''''''''''''''''''''''''''''''''
Option Explicit
Const ForReading = 1
Dim TheString	'The String we are looking for
Dim g_ShellObj	'Object used for sending text to a message box
''''''''''''''''''''''''''''''''''''''
'change INI File here
Const Filespec="\\SERVER\C$\FILENAME.INI"
''''''''''''''''''''''''''''''''''''''
Set g_ShellObj = CreateObject("Wscript.Shell")
'Starting Main function
''''''''''''''''''''''''''''''''''''''
	'Proper use is: ReadFromINI(INI file, Item in brackets, Item we are looking for)
	TheString=ReadFromINI(Filespec,"PutBracketItemHere","PutItemBeingLookedForHere")
''''''''''''''''''''''''''''''''''''''
	'This shows what has been found
	WScript.Echo Now() & " --> Ended **" & TheString & "**"
Function ReadFromINI(INIfile,BracketItem,TheItem)
Dim fsoIN, Fin					'Objects for Reading.
Dim FoundBracket, FoundTheItem		'Keeps tracks of what we have found so far.
Dim CurrStr						'Last string that was read from the INI file.
Dim I 						'Integer used for stepping through CurrStr.
Dim StringFound					'String we are looking for.
Dim C							'Current character while stepping through CurrStr
'Initialize variables
FoundBracket=False
FoundTheItem=False
CurrSTr=""
StringFound=""
	'Create an object and open file for reading.
	Set fsoIN = CreateObject("Scripting.FileSystemObject")
	Set Fin = FsoIN.OpenTextFile(INIfile, ForReading)
	'Stepping through file line by line to find what we are looking for.
	Do While Fin.AtEndOfStream <> True
		CurrStr=Fin.readline
		If left(CurrStr,1)="[" Then	'Looking for an item in brackets
			If ucase(mid(CurrSTr,2,len(BracketItem)))=ucase(BracketItem) Then
				FoundBracket=True
			Else
				FoundBracket=False
			End If
		Else
			'Once we are within the right section we start searching for
			'the correct item we are looking for.
			If FoundBracket Then
				'Compare each item to the item we are looking for.
				If ucase(left(CurrSTr,len(TheItem)))=ucase(TheItem) Then
					'We found the item! We must find where the equal sign
					'is so we don't include it in our result.
					I = len(TheItem)+1
					Do While I<len(CurrStr)
						C = MID(CurrStr,I,1)
						If C<>" " And C<>"=" Then
							'This is not the right item but similar name.
							'example: We're looking for "TheGreatThing" while
							'we found "TheGreatThingy". (Notice the "y")
							i=Len(CurrStr)+10
						Else
							If C="=" Then
								'We found the equal sign, we can now create our
								'String!
								StringFound=Right(CurrStr,Len(CurrStr)-I)
								I=Len(CurrStr)
								FoundTheItem=True
							Else
								'Just a space, we got to keep stepping through
								'the string until we find that equal sign.
								I=I+1
							End If
						End If
					Loop
				End If
			End If
		End If
	Loop
	'Close the file and clear the object.
	Fin.close
	Set fsoIN=Nothing
	'Can't forget to Set the function's variable
	ReadFromINI=TRIM(StringFound)
End Function
'Have a nice day!
```

