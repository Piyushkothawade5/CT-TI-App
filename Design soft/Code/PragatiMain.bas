Attribute VB_Name = "PragatiMain"
Public con As ADODB.Connection
Public con1 As ADODB.Connection
Public com As Command
Public loggedas As String
Dim hundreds, thousands, lakhs, crores As Integer
Dim words(100) As String
Public InCalcu As String
Dim place As String
Public DCode As String
Public NonStad As String
Public MemoStc As String
Public MemoTime As String
Public CancelCT As String
Public rssave As ADODB.Recordset
Public rsgen As ADODB.Recordset
Public PIDS As Double

Public Sub Main()

prcConnect
'COSTMASTERfrm.Show

'ChangeMod
MDIDesign.Show
'Mouldfrm.Show
End Sub

Private Sub ChangeMod()
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select * from primarycond11", con, adOpenDynamic, adLockOptimistic
    If rs.EOF = False And rs.BOF = False Then
        rs.MoveFirst
        Do Until rs.EOF
            rs!mod = Trim(rs!mod)
            rs.Update
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
End Sub

Public Sub prcLastCal()

CTfrm.wounded.Rows = 1
CTfrm.Mould.Rows = 1
CTfrm.Casting.Rows = 1
CTfrm.Former.Rows = 1
CTfrm.txtPriMeanL = ""
CTfrm.txtPriLength = ""
CTfrm.txtPriTotLength = ""
CTfrm.txtpricuWt = ""
CTfrm.txtPRiRes20 = ""
CTfrm.txtPRiRes75 = ""
CTfrm.txtWattPri = ""
CTfrm.txtCoreWT = ""
CTfrm.txtSecCuWt = ""
CTfrm.txtPrimaryCopper = ""
CTfrm.txtResinWt = ""

End Sub

'Public Sub prcConnect()
'On Error GoTo trap
'Dim s As String
's = "c:\order processing\pragathi.mdb"
'
'Set con = New adodb.Connection
'
'With con
'
'    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= '" & s & "' ;Persist Security Info=False"
'    .Open
'
'End With
'
'Exit Sub
'trap:
'If Err.Number = -2147467259 Then
'    MsgBox "Unable to establish database connection." & vbCrLf & "Details : Your database server may be offline." & vbCrLf & "Turn your database server online and try again." & vbCrLf & "Or database may be missing.", vbCritical
'    End
'End If

Public Sub prcConnect()
On Error GoTo trap


Set con = New Connection

With con

    '.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\On Going Project\gemini\ENHANCE\Enhance\Pragathi.mdb;Persist Security Info=False" '"Provider=SQLOLEDB.1;Persist Security Info=False;User ID=prag;Initial Catalog=pragati;Data Source=server"
    .ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=ENHANCE"
    .Open

End With


trap:
If err.Number = -2147467259 Then
    MsgBox "Unable To Establish Database Connection." & vbCrLf & " Your Database Server May Be Offline." & vbCrLf & " Turn Your Database Server Online And Try Again.", vbCritical
    End
End If

End Sub

'Public Sub prcConnect1()
'On Error GoTo trap
'
'
'Set con1 = New Connection
'
'With con1
'
'    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\Pragathi.mdb;Persist Security Info=False"
'    .Open
'
'End With
'
'
'trap:
'If Err.Number = -2147467259 Then
'    MsgBox "Unable To Establish Database Connection." & vbCrLf & " Your Database Server May Be Offline." & vbCrLf & " Turn Your Database Server Online And Try Again.", vbCritical
'    End
'End If
'
'End Sub
'
'

''





Public Function convert(value As String) As String
Class_Initialize
Length = Len(value)
Do Until Length <= 0
place = ""
    If Length = 9 Then
        If Val(Left(Trim(value), 2)) = 1 Then
            place = words(Val(Left(Trim(value), 2))) & " Crore"
        ElseIf Val(Left(Trim(value), 2)) > 1 Then
            place = words(Val(Left(Trim(value), 2))) & " Crores"
        End If
        Length = 7
        value = Mid(Trim(value), 3)
    ElseIf Length = 8 Then
        If Val(Left(Trim(value), 1)) = 1 Then
            place = words(Val(Left(Trim(value), 1))) & " Crore"
        ElseIf Val(Left(Trim(value), 1)) > 1 Then
            place = words(Val(Left(Trim(value), 1))) & " Crores"
        End If
        Length = 7
        value = Mid(Trim(value), 2)
    ElseIf Length = 7 Then
        If Val(Left(Trim(value), 2)) = 1 Then
            place = words(Val(Left(Trim(value), 2))) & " Lakh"
        ElseIf Val(Left(Trim(value), 2)) > 1 Then
            place = words(Val(Left(Trim(value), 2))) & " Lakhs"
        End If
        Length = 5
        value = Mid(Trim(value), 3)
    ElseIf Length = 6 Then
        If Val(Left(Trim(value), 1)) = 1 Then
            place = words(Val(Left(Trim(value), 1))) & " Lakh"
        ElseIf Val(Left(Trim(value), 1)) > 1 Then
            place = words(Val(Left(Trim(value), 1))) & " Lakhs"
        End If
        Length = 5
        value = Mid(Trim(value), 2)
    ElseIf Length = 5 Then
        If Val(Left(Trim(value), 2)) = 1 Then
            place = words(Val(Left(Trim(value), 2))) & " Thousand"
        ElseIf Val(Left(Trim(value), 2)) > 1 Then
            place = words(Val(Left(Trim(value), 2))) & " Thousand"
        End If
        Length = 3
        value = Mid(Trim(value), 3)
    ElseIf Length = 4 Then
        If Val(Left(Trim(value), 1)) = 1 Then
            place = words(Val(Left(Trim(value), 1))) & " Thousand"
        ElseIf Val(Left(Trim(value), 1)) > 1 Then
            place = words(Val(Left(Trim(value), 1))) & " Thousand"
        End If
        Length = 3
        value = Mid(Trim(value), 2)
    ElseIf Length = 3 Then
        If Val(Left(Trim(value), 1)) >= 1 Then
            place = words(Val(Left(Trim(value), 1))) & " Hundred"
        End If
        Length = 1
        value = Mid(Trim(value), 2)
    Else
        If Val(Left(Trim(value), 2)) >= 1 Then
            place = "and " & words(Val(Left(Trim(value), 2)))
        End If
        Length = 0
    End If
    convert = convert & place & " "
Loop
End Function

Public Sub Class_Initialize()
words(0) = ""
words(1) = "One"
words(2) = "Two"
words(3) = "Three"
words(4) = "Four"
words(5) = "Five"
words(6) = "Six"
words(7) = "Seven"
words(8) = "Eight"
words(9) = "Nine"
words(10) = "Ten"
words(11) = "Eleven"
words(12) = "Twelve"
words(13) = "Thirteen"
words(14) = "Fourteen"
words(15) = "Fifteen"
words(16) = "Sixteen"
words(17) = "Seventeen"
words(18) = "Eighteen"
words(19) = "Nineteen"
words(20) = "Twenty"
words(21) = "Twenty One"
words(22) = "Twenty Two"
words(23) = "Twenty Three"
words(24) = "Twenty Four"
words(25) = "Twenty Five"
words(26) = "Twenty Six"
words(27) = "Twenty Seven"
words(28) = "Twenty Eight"
words(29) = "Twenty Nine"
words(30) = "Thirty"
words(31) = "Thirty One"
words(32) = "Thirty Two"
words(33) = "Thirty Three"
words(34) = "Thirty Four"
words(35) = "Thirty Five"
words(36) = "Thirty Six"
words(37) = "Thirty Seven"
words(38) = "Thirty Eight"
words(39) = "Thirty Nine"
words(40) = "Forty"
words(41) = "Forty One"
words(42) = "Forty Two"
words(43) = "Forty Three"
words(44) = "Forty Four"
words(45) = "Forty Five"
words(46) = "Forty Six"
words(47) = "Forty Seven"
words(48) = "Forty Eight"
words(49) = "Forty Nine"
words(50) = "Fifty"
words(51) = "Fifty One"
words(52) = "Fifty Two"
words(53) = "Fifty Three"
words(54) = "Fifty Four"
words(55) = "Fifty Five"
words(56) = "Fifty Six"
words(57) = "Fifty Seven"
words(58) = "Fifty Eight"
words(59) = "Fifty Nine"
words(60) = "Sixty"
words(61) = "Sixty One"
words(62) = "Sixty Two"
words(63) = "Sixty Three"
words(64) = "Sixty Four"
words(65) = "Sixty Five"
words(66) = "Sixty Six"
words(67) = "Sixty Seven"
words(68) = "Sixty Eight"
words(69) = "Sixty Nine"
words(70) = "Seventy"
words(71) = "Seventy One"
words(72) = "Seventy Two"
words(73) = "Seventy Three"
words(74) = "Seventy Four"
words(75) = "Seventy Five"
words(76) = "Seventy Six"
words(77) = "Seventy Seven"
words(78) = "Seventy Eight"
words(79) = "Seventy Nine"
words(80) = "Eighty"
words(81) = "Eighty One"
words(82) = "Eighty Two"
words(83) = "Eighty Three"
words(84) = "Eighty Four"
words(85) = "Eighty Five"
words(86) = "Eighty Six"
words(87) = "Eighty Seven"
words(88) = "Eighty Eight"
words(89) = "Eighty Nine"
words(90) = "Ninety"
words(91) = "Ninety One"
words(92) = "Ninety Two"
words(93) = "Ninety Three"
words(94) = "Ninety Four"
words(95) = "Ninety Five"
words(96) = "Ninety Six"
words(97) = "Ninety Seven"
words(98) = "Ninety Eight"
words(99) = "Ninety Nine"
End Sub

Public Function RoundUp(no As Double) As Double
        
Dim str1 As String
Dim value As String
If InStr(1, no, ".") > 0 Then
    value = Mid(no, InStr(1, no, ".") + 1)
    str1 = Mid(no, 1, InStr(1, no, ".") - 1)
    If value > 1 And value < 50 Then
        RoundUp = str1 & "." & 5
    ElseIf value >= 50 Then
        RoundUp = CDbl(str1) + 1
    End If
Else
    RoundUp = no
End If
End Function
Public Sub rsave()
Set rssave = New ADODB.Recordset
With rssave
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .ActiveConnection = con
End With
End Sub
Public Sub rgen()
Set rsgen = New ADODB.Recordset
With rsgen
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .ActiveConnection = con
End With
End Sub

