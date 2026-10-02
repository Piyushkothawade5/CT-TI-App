VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form STDesignSearchfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Design"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbCancel 
      Caption         =   "Can&cel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2235
      TabIndex        =   4
      ToolTipText     =   "Press to exit the print form"
      Top             =   3870
      Width           =   1395
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Press to view Preview"
      Top             =   3885
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Matched Des. No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   3435
      End
      Begin MSForms.ListBox lstDno 
         Height          =   3195
         Left            =   90
         TabIndex        =   2
         Top             =   540
         Width           =   3435
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "6059;5636"
         ColumnCount     =   3
         cColumnInfo     =   3
         MatchEntry      =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1763;1763;1762"
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1380
      Top             =   3945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   1830
      Top             =   3900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "STDesignSearchfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rst As Recordset

Private Sub cmbCancel_Click()
Unload Me
Set PrintDesfrm = Nothing
    MDIDesign.mnnPrintDesign.Enabled = True
    MDIDesign.Toolbar1.Buttons(3).Enabled = True
End Sub

Private Sub cmdOK_Click()

On Error GoTo err

If Len(lstDno) = 0 Then
    MsgBox "Please select a design No", vbInformation
    Exit Sub
End If

MousePointer = vbHourglass

Dim Rep As Report
Set Rep = New Report
Rep.ReportDesign (lstDno)

Set rst = New Recordset
rst.Open "select equip from designspec where des_code = '" & lstDno & "'", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
        CrystalReport1.DiscardSavedData = True
    If InStr(1, UCase(rst!Equip), "WOUND") > 0 Then
        CrystalReport1.ReportFileName = "\\server\c\reports\design\ctdesign5.rpt"
    Else
        CrystalReport1.ReportFileName = "\\server\c\reports\design\ctdesignwindow.rpt"
    End If
    CrystalReport2.ReportFileName = "\\server\c\reports\design\costing.rpt"
    CrystalReport1.ReportTitle = "HT Current Transformer Design No : " & lstDno
    CrystalReport1.Action = 1
    CrystalReport2.DiscardSavedData = True
    CrystalReport2.ReportTitle = lstDno
    CrystalReport2.Action = 1
End If
Set rst = Nothing

MousePointer = vbDefault

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If

End Sub

Private Sub Form_Activate()

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

End Sub

Private Sub Form_GotFocus()
    MDIDesign.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub Form_Load()
    
'FillDno
MDIDesign.StatusBar1.Panels(2).Text = "Search Design Results"

End Sub


'Private Sub FillDno()
'
'lstDno.Clear
'Set rst = New Recordset
'rst.Open "select distinct des_code from designspec order by des_code desc", con, adOpenStatic, adLockReadOnly
'If rst.EOF = False And rst.BOF = False Then
'    rst.MoveFirst
'    Do Until rst.EOF
'        lstDno.AddItem rst.Fields(0)
'        rst.MoveNext
'    Loop
'End If
'Set rst = Nothing
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set PrintDesfrm = Nothing
'    MDIDesign.mnnPrintDesign.Enabled = True
'    MDIDesign.Toolbar1.Buttons(3).Enabled = True
    MDIDesign.StatusBar1.Panels(2).Text = ""

End Sub



Private Sub lstDno_Click()
cmdOK_Click
End Sub

Private Sub lstDno_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Select Design No."
End Sub

Private Sub txtSearch_Change()

On Error GoTo err

For i = 0 To lstDno.ListCount - 1
    If UCase(Mid(lstDno.List(i), 1, Len(txtSearch))) = UCase(txtSearch) Then
        
        lstDno.ListIndex = i
        lstDno.TopIndex = i
        Exit Sub
        
    End If

Next
lstDno.Text = ""
For Each Control In PrintDesfrm
    
    If TypeOf Control Is TextBox And Control.Name <> "txtSearch" Then
    
        Control.Text = ""
    
    End If

Next

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If

'txtTo.Text = ""
'txtItem_code = ""
End Sub

Private Sub txtSearch_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Enter design No."
End Sub


