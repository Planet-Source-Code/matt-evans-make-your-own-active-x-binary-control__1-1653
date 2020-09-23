<div align="center">

## \* Make your own Active X Binary Control\!\!\! \*


</div>

### Description

* This is the code to put in the General place of you UserControl. By Doing this you can make a Active-X control that Makes Text to Binary, Binary to Text, and to see if a string is Binary! *
 
### More Info
 
Text to change

none!!! easy for beginners!

Changed text


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Evans](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-evans.md)
**Level**          |Unknown
**User Rating**    |3.7 (44 globes from 12 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-evans-make-your-own-active-x-binary-control__1-1653/archive/master.zip)





### Source Code

```
'enjoy! ;D
'put this in a module, we don't want the user to
'see this lil function, he has no need too
Public Function ChrAscii(Char As String) As Long
 Dim GetAscii&
 For GetAscii& = 0 To 255
  If Mid(Char$, 1, 1) = Chr(GetAscii) Then
   ChrAscii = GetAscii
  Exit Function
  End If
 Next GetAscii&
End Function
'Double Click on the user control, and in the General Declarations
'Put this... these are the subs the you will use
Public Function TextToBinary(StringT As String) As String
Dim Ascii, FinalBinary$, GetNum&
FinalBinary$ = ""
For GetNum& = 1 To Len(StringT$)
 Ascii = ChrAscii(Mid(StringT$, GetNum, 1))
' 128
 If Ascii >= 128 Then
   FinalBinary$ = FinalBinary$ & "1"
  Ascii = Ascii - 128
 Else
  FinalBinary$ = FinalBinary$ & "0"
 End If
 ' 64
 If Ascii >= 64 Then
  FinalBinary$ = FinalBinary$ & "1"
  Ascii = Ascii - 64
 Else
  FinalBinary$ = FinalBinary$ & "0"
 End If
 ' 32
 If Ascii >= 32 Then
  FinalBinary$ = FinalBinary$ & "1"
  Ascii = Ascii - 32
 Else
  FinalBinary$ = FinalBinary$ & "0"
 End If
 ' 16
 If Ascii >= 16 Then
  FinalBinary$ = FinalBinary$ & "1"
  Ascii = Ascii - 16
 Else
  FinalBinary$ = FinalBinary$ & "0"
 End If
 ' 8
 If Ascii >= 8 Then
  FinalBinary$ = FinalBinary$ & "1"
  Ascii = Ascii - 8
 Else
  FinalBinary$ = FinalBinary$ & "0"
 End If
 ' 4
 If Ascii >= 4 Then
  FinalBinary$ = FinalBinary$ & "1"
  Ascii = Ascii - 4
 Else
  FinalBinary$ = FinalBinary$ & "0"
 End If
 ' 2
  If Ascii >= 2 Then
   FinalBinary$ = FinalBinary$ & "1"
   Ascii = Ascii - 2
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If
 ' 1
  If Ascii >= 1 Then
   FinalBinary$ = FinalBinary$ & "1"
   Ascii = Ascii - 1
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If
  If Mid(StringT$, GetNum + 1, 1) = Chr(32) Then
    FinalBinary$ = FinalBinary$ '& " "
  Else
    FinalBinary$ = FinalBinary$ '& Chr(32)
  End If
 Next GetNum&
 TextToBinary$ = FinalBinary$
End Function
Public Function BinaryToText(BinaryString As String) As String
Dim GetBinary&, Num$, Binary&, FinalString$, NewString$
NextChr:
For GetBinary& = 1 To 8
 Num$ = Mid(BinaryString$, GetBinary&, 1)
 Select Case Num$
  Case "1"
    If GetBinary = 1 Then
       Binary = Binary + 128
      ElseIf GetBinary = 2 Then
       Binary = Binary + 64
      ElseIf GetBinary = 3 Then
       Binary = Binary + 32
      ElseIf GetBinary = 4 Then
       Binary = Binary + 16
      ElseIf GetBinary = 5 Then
        Binary = Binary + 8
      ElseIf GetBinary = 6 Then
        Binary = Binary + 4
      ElseIf GetBinary = 7 Then
        Binary = Binary + 2
      ElseIf GetBinary = 8 Then
        Binary = Binary + 1
    End If
  End Select
 Next GetBinary&
FinalString$ = FinalString$ & Chr(Binary)
NewString$ = Mid(BinaryString$, 9)
 If NewString$ = "" Then
  BinaryToText$ = FinalString$
 Else
  BinaryString$ = NewString$
  Binary = 0
  GoTo NextChr
 End If
End Function
Public Function IsBinary(StringB As String) As Boolean
Dim XX$, GetLet&
For GetLet& = 1 To Len(StringB$)
 XX$ = Mid(StringB$, GetLet&, 1)
 If XX$ <> "0" Or XX$ <> "1" Then
  If XX$ = "0" Or XX$ = "1" Then GoTo GetNext
  IsBinary = False
  Exit Function
 Else
  '''
 End If
GetNext:
Next GetLet&
IsBinary = True
 End Function
```

