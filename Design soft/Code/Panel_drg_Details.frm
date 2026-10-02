VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Panel_drg_Details 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Panel Details"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   6060
      TabIndex        =   14
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
      Height          =   495
      Left            =   4935
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   4785
      TabIndex        =   12
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add/Modify"
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame framDrg 
      Caption         =   "Drg Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1515
      Left            =   4440
      TabIndex        =   8
      Top             =   600
      Width           =   2715
      Begin MSForms.ComboBox cmbDrgType 
         Height          =   315
         Left            =   1140
         TabIndex        =   17
         Top             =   960
         Width           =   1515
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbDrg 
         Height          =   315
         Left            =   1140
         TabIndex        =   16
         Top             =   420
         Width           =   1515
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Drg. Type :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Drg. :"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   420
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "&>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   7
      Top             =   3780
      Width           =   675
   End
   Begin VB.CommandButton cmdIn 
      Caption         =   "<&<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   2400
      Width           =   675
   End
   Begin VB.Frame framPanel 
      Caption         =   "Panel Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1755
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   4395
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   675
         Left            =   1140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1020
         Width           =   3135
      End
      Begin MSForms.ComboBox cmbCus 
         Height          =   315
         Left            =   1140
         TabIndex        =   19
         Top             =   240
         Width           =   3135
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5530;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Customer  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Details :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   510
         TabIndex        =   5
         Top             =   1020
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Panel Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   960
      End
      Begin MSForms.ComboBox cmbPanelName 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   660
         Width           =   3135
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5530;556"
         ColumnCount     =   3
         cColumnInfo     =   3
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;3527;0"
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1875
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   3307
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "Panel_drg_Details.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Panel Details...."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   15
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Panel_drg_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gmodi As Boolean

Private Sub cmbCus_Click()
fillpanel
End Sub

Private Sub cmbDrg_Click()
fillDrgType
End Sub

Private Sub cmbPanelName_Click()
On Error GoTo err
Dim i As Integer
txtDetails = ""
txtDetails = Trim(cmbPanelName.List(cmbPanelName.ListIndex, 2))
Grid.Rows = 1
rgen
rsgen.Open "Select * from PANEL_DRG_FILL WHERE panel_Code=" & cmbPanelName.List(cmbPanelName.ListIndex, 0) & " and CUSTOMER='" & Trim(cmbCus) & "'"
If rsgen.EOF = False And rsgen.BOF = False Then
    txtDetails = Trim(rsgen!panel_details)
    For i = 1 To rsgen.RecordCount
        Grid.Rows = Grid.Rows + 1
        Grid.Row = Grid.Rows - 1
        Grid.TextMatrix(Grid.Row, 0) = Grid.Row
        Grid.TextMatrix(Grid.Row, 1) = Trim(rsgen!panel_CODE) & vbNullString
        Grid.TextMatrix(Grid.Row, 2) = Trim(rsgen!Drg) & vbNullString
        Grid.TextMatrix(Grid.Row, 3) = Trim(rsgen!DRG_TYPE) & vbNullString
        rsgen.MoveNext
    Next
End If
    
If Grid.Rows > 1 Then
    Grid.Enabled = True
    cmdOut.Enabled = True
Else
    cmdOut.Enabled = False
    Grid.Enabled = False
End If

Set rsgen = Nothing
Exit Sub
err:
    MsgBox err.Description
End Sub
    
Private Sub cmdadd_Click()
    prcunlock
    prcclear
End Sub

Private Sub cmdcancel_Click()
prclock
prcclear
End Sub
Private Sub fillpanel()
On Error GoTo err
Dim ctr As Integer
    rgen
    rsgen.Open "Select * from PANEL_MASTER where CUSTOMER='" & cmbCus & "' order by Panel_Name"
    cmbPanelName.Clear
    While rsgen.EOF = False
        cmbPanelName.AddItem Trim(rsgen!panel_CODE)
        cmbPanelName.List(ctr, 1) = Trim(rsgen!panel_NAME)
        cmbPanelName.List(ctr, 2) = Trim(rsgen!panel_details)
        ctr = ctr + 1
        rsgen.MoveNext
    Wend
    Set rsgen = Nothing
Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmdin_Click()
On Error GoTo err
If Trim(cmbPanelName.ListIndex) = -1 Then
    MsgBox "Select Panel !!!", vbInformation
    Exit Sub
End If
If Trim(cmbDrg.ListIndex) = -1 Then
    MsgBox "Enter Drg Value !!!", vbInformation
    Exit Sub
End If
If Trim(cmbDrgType.ListIndex) = -1 Then
    MsgBox "Enter Drg Type Valuew !!!", vbInformation
    Exit Sub
End If
Dim i As Integer
For i = 1 To Grid.Rows - 1
    If Trim(Grid.TextMatrix(i, 2)) = Trim(cmbDrg) And Trim(Grid.TextMatrix(i, 3)) = Trim(cmbDrgType) Then
        MsgBox "Drg and Drg Type already exist inside the grid !!!", vbExclamation
        Exit Sub
    End If
Next
If gmodi <> True Then
    Grid.Rows = Grid.Rows + 1
    Grid.Row = Grid.Rows - 1
End If
    Grid.RowHeight(Grid.Row) = 350
    Grid.TextMatrix(Grid.Row, 0) = Grid.Row
    Grid.TextMatrix(Grid.Row, 1) = cmbPanelName.List(cmbPanelName.ListIndex, 0) & vbNullString
    Grid.TextMatrix(Grid.Row, 2) = Trim(cmbDrg)
    Grid.TextMatrix(Grid.Row, 3) = Trim(cmbDrgType) & vbNullString
    Grid.Enabled = True
    cmbDrg = ""
    cmbDrgType = ""
    cmdOut.Enabled = True
    gmodi = False
    If Grid.Rows > 1 Then
        framPanel.Enabled = False
    Else
        framPanel.Enabled = True
    End If
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub




Private Sub cmdout_Click()
Dim Ans As Integer
Dim ctr As Integer
    
    If Grid.Rows = 1 Then
        cmdOut.Enabled = False
    End If
    If gmodi = True Then
        MsgBox "Can not Remove While Modifying", vbInformation
        Exit Sub
    End If
    If Grid.Rows = Null Then Exit Sub
        Ans = MsgBox("Are you sure, you want to remove " & UCase(Grid.TextMatrix(Grid.RowSel, 1)) & "  ?", vbQuestion + vbYesNo)
    If Ans = vbNo Then Exit Sub
    
    If Grid.Rows = 2 Then
        Grid.Rows = 1
        cmdOut.Enabled = False
        framPanel.Enabled = True
    Else
        Grid.RemoveItem (Grid.RowSel)
        framPanel.Enabled = False
    End If
    
    ctr = 1
    For ctr = 1 To Grid.Rows - 1
        Grid.TextMatrix(ctr, 0) = ctr
    Next

End Sub

Private Sub cmdSave_Click()
On Error GoTo err
Dim i As Integer
con.Execute "Delete PANEL_DRG_DETAILS where PANEL_CODE=" & cmbPanelName.List(cmbPanelName.ListIndex, 0) & ""
'rsave
'rssave.Open "Select * from PANEL_DRG_DETAILS where PANEL_CODE='" & cmbPanelName.List(cmbPanelName.ListIndex, 0) & "'"
'If rssave.EOF = False And rssave.BOF = False Then
    For i = 1 To Grid.Rows - 1
        rsave
        rssave.Open "Select * from PANEL_DRG_DETAILS"
        rssave.AddNew
        rssave!panel_CODE = Trim(Grid.TextMatrix(i, 1))
        rssave!Drg = Trim(Grid.TextMatrix(i, 2))
        rssave!DRG_TYPE = Trim(Grid.TextMatrix(i, 3))
        rssave.Update
        
        rssave.MoveNext
    Next
    Set rssave = Nothing
    MsgBox "Save Sucessfully !!!", vbInformation
    prclock
    prcclear

Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub gridhead()
Grid.Cols = 4
Grid.Rows = 1
Grid.TextMatrix(0, 0) = "Sr.No"
Grid.ColWidth(0) = 500
Grid.TextMatrix(0, 1) = "PanelCode"
Grid.ColWidth(1) = 0
Grid.TextMatrix(0, 2) = "Drg"
Grid.ColWidth(2) = 1500
Grid.TextMatrix(0, 3) = "Drg Type"
Grid.ColWidth(3) = 1500

End Sub
Private Sub Form_Load()
prclock
gridhead
prcclear

fillDrg
FillCus
End Sub
Private Sub prclock()
    framPanel.Enabled = False
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
End Sub
Private Sub prcunlock()
    framPanel.Enabled = True
    cmdAdd.Enabled = False
    cmdModify.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
End Sub
Private Sub prcclear()
Grid.Rows = 1
cmbDrg = ""
cmbDrgType = ""
txtDetails = ""
cmbCus = ""
cmbPanelName = ""
End Sub

Private Sub Grid_DblClick()
If Grid.Rows = 1 Then Exit Sub
gmodi = True
Dim i As Integer
For i = 1 To cmbDrg.ListCount - 1
    If cmbDrg.List(i, 0) = Trim(Grid.TextMatrix(Grid.Row, 2)) Then
        cmbDrg.ListIndex = i
        Exit For
    End If
Next


For i = 1 To cmbDrgType.ListCount - 1
    If cmbDrgType.List(i, 0) = Trim(Grid.TextMatrix(Grid.Row, 3)) Then
        cmbDrgType.ListIndex = i
        Exit For
    End If
Next

cmbDrg = Trim(Grid.TextMatrix(Grid.Row, 2))
cmbDrgType = Trim(Grid.TextMatrix(Grid.Row, 3))

For i = 1 To Grid.Cols - 1
    Grid.TextMatrix(Grid.Row, i) = "Modify..."
Next
Grid.Enabled = False

End Sub

Private Sub fillDrg()
On Error GoTo err
    rgen
    rsgen.Open "Select distinct Drg from drg order by Drg"
    cmbDrg.Clear
    While rsgen.EOF = False
        cmbDrg.AddItem Trim(rsgen!Drg)
        rsgen.MoveNext
    Wend
    Set rsgen = Nothing
Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub fillDrgType()
On Error GoTo err
    cmbDrgType.Clear
    rgen
    rsgen.Open "Select distinct DRGTYPE,Drg from drg where drg='" & cmbDrg & "' order by Drgtype"
    While rsgen.EOF = False
        If IsNull(rsgen!drgtype) = False Then
            cmbDrgType.AddItem Trim(rsgen!drgtype)
        End If
        rsgen.MoveNext
    Wend
    Set rsgen = Nothing
Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub FillCus()
cmbCus.AddItem "SIEMENS"
cmbCus.AddItem "ABB"
cmbCus.AddItem "OTHERS"

End Sub

