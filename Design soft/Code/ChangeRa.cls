VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChangeRa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rst As Recordset

Public Function prcInterChangeRatio(RatioIn As String) As String
    
Dim value As String * 1
Dim str1 As String

    Set rst = New Recordset
    rst.Open "delete from ratiocheck", con, adOpenDynamic, adLockOptimistic
    Set rst = Nothing

For i = 1 To Len(RatioIn) + 1
    value = Mid(RatioIn, i, i)
    If value = "-" Or i > Len(RatioIn) Then
        Set rst = New Recordset
        rst.Open "select * from ratiocheck", con, adOpenDynamic, adLockOptimistic
        rst.AddNew
        rst.Fields(0) = str1
        rst.Update
        Set rst = Nothing
        str1 = ""
    Else
        str1 = str1 & value
    End If
Next

str1 = ""

Set rst = New Recordset
rst.Open "select * from ratiocheck order by ratio", con, adOpenDynamic, adLockOptimistic
If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    Do Until rst.EOF
        str1 = str1 & "-" & rst.Fields(0)
        rst.MoveNext
    Loop

End If
Set rst = Nothing
prcInterChangeRatio = Mid(str1, 2)
End Function

Public Function prcInterChangeRaredCon(RatioIn As String) As String

If Len(RatioIn) = 0 Then
    prcInterChangeRaredCon = ""
    Exit Function
End If

Dim value As String * 1
Dim str1 As String

    Set rst = New Recordset
    rst.Open "delete from ratiocheck", con, adOpenDynamic, adLockOptimistic
    Set rst = Nothing

For i = 1 To Len(RatioIn) + 1
    value = Mid(RatioIn, i, i)
    If value = "-" Or i > Len(RatioIn) Then
        Set rst = New Recordset
        rst.Open "select * from ratiocheck", con, adOpenDynamic, adLockOptimistic
        rst.AddNew
        rst.Fields(0) = str1
        rst.Update
        Set rst = Nothing
        str1 = ""
    Else
        str1 = str1 & value
    End If
Next

str1 = ""

Set rst = New Recordset
rst.Open "select * from ratiocheck order by ratio desc", con, adOpenDynamic, adLockOptimistic
If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    str1 = rst.Fields(0)

End If
Set rst = Nothing
prcInterChangeRaredCon = CDbl(str1)
End Function
