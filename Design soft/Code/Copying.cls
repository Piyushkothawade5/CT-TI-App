VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Copying"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Sub CopyGrid(G1ColNo As Integer)

On Error GoTo err

''+++++++++++++++++++++++++++++++
''copying g1 to g2

'Seting col 1 for Grid G2
CTfrm.G2.Cols = 1
Count = 0
'Colume Count
Dim count1 As Integer
For i = 1 To CTfrm.G1.Cols - 1
    If CInt(CTfrm.G1.TextMatrix(1, i)) <> G1ColNo Then
        count1 = count1 + 1
    End If
Next

'Adding no of cols as per count
''din king as Integer
CTfrm.G2.Cols = count1 + 1
CTfrm.G2.Rows = 110
king = 0
For i = 1 To CTfrm.G1.Cols - 1
    If CInt(CTfrm.G1.TextMatrix(1, i)) <> G1ColNo Then
        king = king + 1
        For j = 0 To 109
            CTfrm.G2.TextMatrix(j, king) = CTfrm.G1.TextMatrix(j, i)
        Next
    End If
Next

''ended copying
''++++++++++++++++++++++++++++++++++++

''++++++++++++++++++++++++++++++++++++
''g1 is set to 1 so that
''g2 to be loaded without the deleted core
    CTfrm.G1.Cols = 1
    CTfrm.G1.Cols = CTfrm.G2.Cols
''end
''+++++++++++++++++++++++++++++++++++++

''++++++++++++++++++++++++++++++++++++
''Proper data without delete core details
    For i = 1 To CTfrm.G2.Cols - 1
        For j = 0 To 109
            CTfrm.G1.TextMatrix(j, i) = CTfrm.G2.TextMatrix(j, i)
        Next
    Next
''end
''+++++++++++++++++++++++++++++++++++++

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If



End Sub
