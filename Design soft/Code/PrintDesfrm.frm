VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PrintDesfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Design"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "PrintDesfrm.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   3705
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1395
      Top             =   5445
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Design"
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
      Left            =   2250
      TabIndex        =   3
      ToolTipText     =   "Press to exit the print form"
      Top             =   5370
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
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Press to view Preview"
      Top             =   5370
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Des. No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   45
      TabIndex        =   4
      Top             =   45
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
         TabIndex        =   0
         Top             =   240
         Width           =   3435
      End
      Begin VB.ListBox lstDno 
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
         Height          =   4650
         Left            =   90
         TabIndex        =   1
         Top             =   540
         Width           =   3435
      End
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   1845
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Costing"
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
Attribute VB_Name = "PrintDesfrm"
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

Dim Rep As Report
Set Rep = New Report
Rep.ReportDesign (lstDno)


Set rst = New Recordset
rst.Open "select equip from designspec where des_code = '" & lstDno & "'", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
        CrystalReport1.DiscardSavedData = True
    If InStr(1, UCase(rst!Equip), "WOUND") > 0 Then
    '**************
        'CrystalReport1.ReportFileName = "c:\ctdesign5.rpt"
        
        CrystalReport1.ReportFileName = App.Path & "\REPORTs\ctdesign5.rpt"
    Else
        CrystalReport1.ReportFileName = App.Path & "\REPORTs\ctdesignwindow.rpt"
    End If
    'CrystalReport2.ReportFileName = "c:\Costing--F.rpt"
    CrystalReport2.ReportFileName = App.Path & "\REPORTs\Costing.rpt"
    
    CrystalReport1.ReportTitle = "HT Current Transformer Design No : " & lstDno
    CrystalReport1.Connect = "dsn=enhance;uid=sa;pwd=shilpraj;database=gem0506"
    
    CrystalReport1.Action = 1
    CrystalReport2.DiscardSavedData = True
    CrystalReport2.Connect = "dsn=enhance;uid=sa;pwd=shilpraj;database=gem0506"
    CrystalReport2.ReportTitle = lstDno
    CrystalReport2.Action = 1
End If
Set rst = Nothing

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If

End Sub

Private Sub Form_Activate()

Me.Top = (MDIDesign.ScaleHeight - Me.Height) / 2
Me.Left = (MDIDesign.ScaleWidth - Me.Width) / 2

End Sub

Private Sub Form_GotFocus()
    MDIDesign.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub Form_Load()
    
FillDno
MDIDesign.StatusBar1.Panels(2).Text = "Print Design"

End Sub


Private Sub FillDno()

lstDno.Clear
Set rst = New Recordset
rst.Open "select distinct des_code from designspec order by des_code desc", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
    rst.MoveFirst
    Do Until rst.EOF
        lstDno.AddItem rst.Fields(0)
        rst.MoveNext
    Loop
End If
Set rst = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set PrintDesfrm = Nothing
    MDIDesign.mnnPrintDesign.Enabled = True
    MDIDesign.Toolbar1.Buttons(3).Enabled = True
    MDIDesign.StatusBar1.Panels(2).Text = ""

End Sub

Private Sub lstDno_DblClick()
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
