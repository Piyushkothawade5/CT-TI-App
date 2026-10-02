VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoDNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rst As Recordset

Public Function AutoGen() As String
Dim Flag As String
Flag = "0"
Dim str1(25) As String
str1(0) = "A"
str1(1) = "B"
str1(2) = "C"
str1(3) = "D"
str1(4) = "E"
str1(5) = "F"
str1(6) = "I"
str1(7) = "J"
str1(8) = "K"
str1(9) = "L"
str1(10) = "M"
str1(11) = "N"
str1(12) = "O"
str1(13) = "P"
str1(14) = "Q"
str1(15) = "R"
str1(16) = "S"
str1(17) = "T"
str1(18) = "U"
str1(19) = "V"
str1(20) = "W"
str1(21) = "X"
str1(22) = "Y"
str1(23) = "Z"

Dim charNo As Integer
Dim DegNo As String

'Set rst = New Recordset
'rst.Open "select * from gen where charecter = 'DEGCH'", con, adOpenDynamic, adLockOptimistic
'If rst.EOF = False And rst.EOF = False Then
'    charNo = rst!code
'Else
'    rst.AddNew
'    rst!charecter = "DEGCH"
'    rst!code = 1
'    charNo = a
'    rst.Update
'End If
'Set rst = Nothing

charNo = 1

Set rst = New Recordset
rst.Open "select * from gen where charecter = 'DEGNO'", con, adOpenDynamic, adLockOptimistic
If rst.EOF = False And rst.EOF = False Then
    If rst!code < 99999 Then
        rst!code = rst!code + 1
        DegNo = "CT" & Format(rst!code, "0000000")
        AutoGen = DegNo
    Else
        Flag = "1"
        rst!code = 1
        charNo = charNo + 1
        DegNo = "CT" & Format(rst!code, "0000000")
        rst.Update
        
        AutoGen = DegNo
    End If
    rst.Update
    
Else
    rst.AddNew
    rst!charecter = "DEGNO"
    rst!code = 1
    DegNo = "CT" & Format(rst!code, "0000000")
    rst.Update
    AutoGen = DegNo
End If

    If Flag = "1" Then
    
        Set rst = New Recordset
        rst.Open "select * from gen where charecter = 'DEGCH'", con, adOpenDynamic, adLockOptimistic
        If rst.EOF = False And rst.EOF = False Then
            rst!code = charNo
        End If
        Set rst = Nothing
        
        Set rst = New Recordset
        rst.Open "select * from gen where charecter = 'DEGNO'", con, adOpenDynamic, adLockOptimistic
        If rst.EOF = False And rst.EOF = False Then
            rst!code = 1
        End If
        Set rst = Nothing
        
        Flag = "0"
        
    End If
Set rst = Nothing
End Function
