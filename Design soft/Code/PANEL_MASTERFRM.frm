VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form PANEL_MASTERFRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Panel Master"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PANEL_MASTERFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   6720
      TabIndex        =   4
      Top             =   2610
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   420
      Left            =   5580
      TabIndex        =   3
      Top             =   2610
      Width           =   1140
   End
   Begin VB.CommandButton cmdmodify 
      Caption         =   "&Modify"
      Height          =   420
      Left            =   4440
      TabIndex        =   6
      Top             =   2610
      Width           =   1140
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add"
      Height          =   420
      Left            =   3300
      TabIndex        =   0
      Top             =   2610
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Caption         =   "Panel Details:"
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
      Height          =   1770
      Left            =   3285
      TabIndex        =   8
      Top             =   720
      Width           =   4605
      Begin VB.TextBox txtcode 
         Height          =   240
         Left            =   45
         TabIndex        =   11
         Text            =   " "
         Top             =   1440
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtdetails 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   1170
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "PANEL_MASTERFRM.frx":0442
         Top             =   990
         Width           =   3270
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1170
         TabIndex        =   1
         Text            =   " "
         Top             =   615
         Width           =   3270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Customer :"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin MSForms.ComboBox cmbCustomer 
         Height          =   330
         Left            =   1170
         TabIndex        =   13
         Top             =   240
         Width           =   3270
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5768;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Details :"
         Height          =   240
         Left            =   45
         TabIndex        =   10
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Panel :"
         Height          =   240
         Left            =   45
         TabIndex        =   9
         Top             =   660
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Panel:"
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
      Height          =   3075
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3300
      Begin MSForms.ListBox lstpanel 
         Height          =   2805
         Left            =   90
         TabIndex        =   5
         Top             =   210
         Width           =   3120
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "5503;5212"
         ColumnCount     =   3
         cColumnInfo     =   3
         MatchEntry      =   0
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;4233;0"
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Panel Master...."
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
      Left            =   5220
      TabIndex        =   12
      Top             =   135
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4200
      Picture         =   "PANEL_MASTERFRM.frx":0444
      Top             =   -15
      Width           =   480
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3375
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   4515
   End
End
Attribute VB_Name = "PANEL_MASTERFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdadd_Click()
On Error GoTo trap
    txtname = ""
    txtDetails = ""
    txtcode = ""
    prcunlock
    cmbCustomer.SetFocus
Exit Sub
trap:
    MsgBox err.Description
End Sub

Private Sub cmdcancel_Click()
On Error GoTo trap
    prcclear
    prclock
Exit Sub
trap:
    MsgBox err.Description
End Sub

Private Sub cmdmodify_Click()
On Error GoTo trap
    If lstpanel.ListIndex = -1 Then
        MsgBox "Select Panel from the list!!!", vbExclamation
        lstpanel.SetFocus
        Exit Sub
    End If
    prcunlock
Exit Sub
trap:
    MsgBox err.Description
End Sub

Private Sub cmdSave_Click()
On Error GoTo trap

    If Trim(cmbCustomer.ListIndex) = -1 Then
        MsgBox "Select Customer!!!", vbExclamation
        cmbCustomer.SetFocus
        Exit Sub
    End If
    If Trim(txtname.Text) = "" Then
        MsgBox "Enter the Name of the Panel!!!", vbExclamation
        txtname.SetFocus
        Exit Sub
    End If
    If Trim(txtDetails.Text) = "" Then
        MsgBox "Enter the Details of the Panel!!!", vbExclamation
        txtDetails.SetFocus
        Exit Sub
    End If
    If Trim(txtcode.Text) = "" Then
        txtcode.Text = gen()
    End If
    MousePointer = vbHourglass
    con.BeginTrans
    rsave
    rssave.Open "select * from panel_Master where panel_Code=" & (txtcode.Text) & ""
    If rssave.EOF = True And rssave.BOF = True Then
        rssave.AddNew
        rssave!panel_CODE = txtcode.Text
    End If
    rssave!panel_NAME = Trim(txtname.Text) & vbNullString
    rssave!panel_details = Trim(txtDetails.Text) & vbNullString
    rssave!CUSTOMER = Trim(cmbCustomer)
    rssave.Update
    Set rssave = Nothing
    con.CommitTrans
    MsgBox "Saved Successfully !!!", vbInformation
    prcclear
    prclock
    MousePointer = vbDefault
Exit Sub
trap:
    MsgBox err.Description
    con.RollbackTrans
End Sub

Private Sub Form_Load()
On Error GoTo trap
    prcclear
    prclock
    FillCus
Exit Sub
trap:
    MsgBox err.Description
End Sub

Private Sub lstpanel_Click()
On Error GoTo trap
    If lstpanel.ListIndex = -1 Then Exit Sub
    txtcode = Trim(lstpanel.List(lstpanel.ListIndex, 0)) & vbNullString
    txtname = Trim(lstpanel.List(lstpanel.ListIndex, 1)) & vbNullString
    txtDetails = Trim(lstpanel.List(lstpanel.ListIndex, 2)) & vbNullString
    cmbCustomer = Trim(lstpanel.List(lstpanel.ListIndex, 3)) & vbNullString
Exit Sub
trap:
    MsgBox err.Description
End Sub

Private Sub prcclear()
On Error GoTo trap
    txtcode = ""
    txtname = ""
    txtDetails = ""
    cmbCustomer = ""
    fillpanel
Exit Sub
trap:
    MsgBox err.Description
End Sub

Private Sub prclock()
On Error GoTo trap
    Frame1.Enabled = True
    Frame2.Enabled = False
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
Exit Sub
trap:
    MsgBox err.Description
End Sub

Private Sub prcunlock()
On Error GoTo trap
    Frame1.Enabled = False
    Frame2.Enabled = True
    cmdAdd.Enabled = False
    cmdModify.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
Exit Sub
trap:
    MsgBox err.Description
End Sub

Private Function gen() As Integer
On Error GoTo trap
    rgen
    rsgen.Open "select max(panel_CODE) from panel_Master"
    If IsNull(rsgen.Fields(0)) = True Then
        gen = 1
    Else
        gen = rsgen.Fields(0) + 1
    End If
    Set rsgen = Nothing
Exit Function
trap:
    MsgBox err.Description
End Function

Private Sub fillpanel()
On Error GoTo trap
Dim i As Integer
    i = 0
    lstpanel.Clear
    rgen
    rsgen.Open "Select panel_code,panel_name,PANEL_DETAILS,CUSTOMER from panel_master order by panel_name "
    While rsgen.EOF = False
        lstpanel.AddItem Trim(rsgen!panel_CODE) & vbNullString
        lstpanel.List(i, 1) = Trim(rsgen!panel_NAME) & vbNullString
        lstpanel.List(i, 2) = Trim(rsgen!panel_details) & vbNullString
        lstpanel.List(i, 3) = Trim(rsgen!CUSTOMER) & vbNullString
        i = i + 1
        rsgen.MoveNext
    Wend
    Set rsgen = Nothing
Exit Sub
trap:
    MsgBox err.Description
End Sub

Private Sub FillCus()
cmbCustomer.AddItem "SIEMENS"
cmbCustomer.AddItem "ABB"
cmbCustomer.AddItem "OTHERS"

End Sub
