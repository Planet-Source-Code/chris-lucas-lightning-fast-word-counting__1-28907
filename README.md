<div align="center">

## Lightning Fast Word Counting


</div>

### Description

A small ultra-fast function used to count the number of words in a string.
 
### More Info
 
Text -the string in which to count words

The number of words counted is returned as the value of the function itself (as a long).

As this function makes use of CopyMemory it should be allowed to run until finished. Stopping the project in the IDE could result in VB crashing (as with ALL API calls).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris\_Lucas ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-lucas.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-lucas-lightning-fast-word-counting__1-28907/archive/master.zip)

### API Declarations

```
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
  Destination As Any, _
  Source As Any, _
  ByVal Length As Long)
```


### Source Code

```
' © Christopher Lucas 2001
' You may freely use and distribute this code
' in all your applications. Recognition is
' appreciated though.
Public Function WordCount(Text As String) As Long
  Dim dest() As Byte
  Dim i As Long
  If LenB(Text) Then
    ' Move the string's byte array into dest()
    ReDim dest(LenB(Text))
    CopyMemory dest(0), ByVal StrPtr(Text), LenB(Text) - 1
    ' Now loop through the array and count the words
    For i = 0 To UBound(dest) Step 2
      If dest(i) > 32 Then
         Do Until dest(i) < 33
          i = i + 2
         Loop
         WordCount = WordCount + 1
      End If
    Next i
    Erase dest
  Else
    WordCount = 0
  End If
End Function
```

