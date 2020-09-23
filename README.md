<div align="center">

## Advanced Base Converter


</div>

### Description

Updated code! This function converts numbers into any base, and any base into decimal! Works frek-in' awesome!
 
### More Info
 
Not all bases will work; there are not enough characters to be able to do every base. An error will occur when this happens.

' The two functions given are reversible. For example:

ConvertToBase(45, 16) = "2D"

ConvertFromBase("2D", 16) = 45

Will not (yet) deal with negative numbers, and rounds all numbers that are decimal.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon Feucht](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-feucht.md)
**Level**          |Advanced
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-feucht-advanced-base-converter__1-2757/archive/master.zip)





### Source Code

```
Function ConvertToBase(DecNumber As Double, NewBase As Integer) As String
Dim ModBase As Double
 Do
 ModBase = CDbl(DecNumber - (Int(DecNumber / NewBase)) * NewBase)
 DecNumber = Int(DecNumber / NewBase)
 If ModBase > 9 Then ModBase = ModBase + 7
 ConvertToBase = Chr(ModBase + 48) & ConvertToBase
 Loop Until DecNumber = 0
End Function
Function ConvertFromBase(BaseNumber As String, OldBase As Integer) As Double
Dim i As Integer, LetterVal As Integer
 On Error Resume Next
 For i = 1 To Len(BaseNumber)
 LetterVal = Asc(Mid(BaseNumber, Len(BaseNumber) - i + 1, 1)) - 48
 If LetterVal > 9 Then LetterVal = LetterVal - 7
 If LetterVal > OldBase Then GoTo InvalidNumber
 ConvertFromBase = ConvertFromBase + (OldBase ^ (i - 1)) * LetterVal
 Next i
InvalidNumber:
End Function
```

