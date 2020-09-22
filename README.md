<div align="center">

## BASIC PRODUCTION


</div>

### Description

Dos Basic like POKE and PEEK functions for VB6
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BASIC PRODUCTION](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/basic-production.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/basic-production-basic-production__1-54863/archive/master.zip)

### API Declarations

```
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Integer)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub GetMem8 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Currency)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Currency)
```


### Source Code

```
Sub POKE(ByVal Address As Variant, ByVal Value As Variant, Optional ByVal HowMuchBits As Byte = 32)
 Select Case HowMuchBits
 Case 8
  PutMem1 Address, Value
 Case 16
  PutMem2 Address, Value
 Case 32
  PutMem4 Address, Value
 Case 64
   PutMem8 Address, Value
 Case Else
  MsgBox "Invalid value length" & vbCr & vbCr & "Must be one from: 8/16/32/64" & vbCr & vbCr & vbTab & "8 - Byte (unsigned)" & vbCr & vbTab & "16 - Word/Integer" & vbCr & vbTab & "32 - Dword/Long" & vbCr & vbTab & "64 - Qword/Currency"
 End Select
End Sub
Function PEEK(ByVal Address As Long, Optional ByVal HowMuchBits As Byte = 32) As Variant
 Dim Value As Variant
 Select Case HowMuchBits
 Case 8
  GetMem1 Address, Value
 Case 16
  GetMem2 Address, Value
 Case 32
  GetMem4 Address, Value
 Case 64
   GetMem8 Address, Value
 Case Else
  MsgBox "Invalid value length" & vbCr & vbCr & "Must be one from: 8/16/32/64" & vbCr & vbCr & vbTab & "8 - Byte (unsigned)" & vbCr & vbTab & "16 - Word/Integer" & vbCr & vbTab & "32 - Dword/Long" & vbCr & vbTab & "64 - Qword/Currency"
  Exit Function
 End Select
 PEEK = Value
End Function
Private Sub Form_Load()
 Dim Var_Byte As Byte, Var_Int As Integer, Var_Lng As Long, Var_Curr As Currency
 Var_Byte = 123: Var_Int = 1234: Var_Lng = 123456: Var_Curr = CDec(5234567890#)
 Dim strMsg As String
 strMsg = "Get value of variables by address with PEEK:" & vbCr
 strMsg = strMsg & "BYTE: " & PEEK(VarPtr(Var_Byte), 8) & vbCr
 strMsg = strMsg & "INTEGER: " & PEEK(VarPtr(Var_Int), 16) & vbCr
 strMsg = strMsg & "LONG: " & PEEK(VarPtr(Var_Lng)) & vbCr
 strMsg = strMsg & "CURRENCY: " & PEEK(VarPtr(Var_Curr), 64)
 MsgBox strMsg
 POKE VarPtr(Var_Byte), 210, 8
 POKE VarPtr(Var_Int), 4321, 16
 POKE VarPtr(Var_Lng), 654321
 POKE VarPtr(Var_Curr), CDec(9999999999#), 64
 strMsg = "Values of variables was changed with POKE:" & vbCr
 strMsg = strMsg & "BYTE: " & Var_Byte & vbCr
 strMsg = strMsg & "INTEGER: " & Var_Int & vbCr
 strMsg = strMsg & "LONG: " & Var_Lng & vbCr
 strMsg = strMsg & "CURRENCY: " & Var_Curr & vbCr
 MsgBox strMsg
End Sub
```

