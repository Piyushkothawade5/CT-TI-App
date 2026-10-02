VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form SecCondfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secondary Conductor Details"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "SecCondfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbswg 
      Height          =   315
      Left            =   1395
      TabIndex        =   0
      Top             =   495
      Width           =   2070
   End
   Begin VB.TextBox txtt_odn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2205
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1905
      Width           =   1140
   End
   Begin VB.TextBox txtt_swg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2205
      MaxLength       =   20
      TabIndex        =   2
      Top             =   945
      Width           =   1140
   End
   Begin VB.TextBox txtt_cs 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2205
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1425
      Width           =   1140
   End
   Begin VB.TextBox txtt_ods 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2205
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2385
      Width           =   1140
   End
   Begin VB.TextBox txtt_part 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2205
      MaxLength       =   20
      TabIndex        =   5
      Top             =   2850
      Width           =   1140
   End
   Begin VB.Shape Shape2 
      Height          =   2355
      Left            =   90
      Top             =   855
      Width           =   3345
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Secondary Wire Gause"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   180
      TabIndex        =   17
      Top             =   90
      Width           =   4785
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   330
      Left            =   45
      Top             =   45
      Width           =   5055
   End
   Begin MSForms.CommandButton cmddelete 
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   1575
      Width           =   1515
      Caption         =   "Delete"
      PicturePosition =   327683
      Size            =   "2672;873"
      Accelerator     =   68
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdmodify 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   1035
      Width           =   1515
      Caption         =   "Modify    "
      PicturePosition =   327683
      Size            =   "2672;873"
      Accelerator     =   79
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdcancel 
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   2655
      Width           =   1515
      Caption         =   "Cancel   "
      PicturePosition =   327683
      Size            =   "2672;873"
      Accelerator     =   67
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdsave 
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   2115
      Width           =   1515
      Caption         =   "Save      "
      PicturePosition =   327683
      Size            =   "2672;873"
      Accelerator     =   83
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdadd 
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   480
      Width           =   1515
      Caption         =   "Add        "
      PicturePosition =   327683
      Size            =   "2672;873"
      Accelerator     =   65
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "S. W. G. :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   16
      Top             =   495
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Diameter(Bare) :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   885
      TabIndex        =   15
      Top             =   1965
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SWG :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   705
      TabIndex        =   14
      Top             =   1005
      Width           =   1425
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C/S Area :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   1485
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Diameter(Insulated) :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   12
      Top             =   2445
      Width           =   1920
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   975
      TabIndex        =   11
      Top             =   2910
      Width           =   1155
   End
End
Attribute VB_Name = "SecCondfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbswg_Click()
    If Len(cmbswg.Text) = 0 Then
        Exit Sub
    End If
    
    prcclear
    
    Dim str As String
    Dim rst As New ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select * from secconductor where t_swg = " & Trim(cmbswg) & "", con, adOpenDynamic, adLockOptimistic
    If rst.EOF = False And rst.BOF = False Then
        For Each obj In Me
            If TypeOf obj Is TextBox Then
                amol = ""
                amol = Mid(obj.Name, 4, Len(obj.Name))
                If IsNull(rst.Fields(amol)) = False Then
'                    MsgBox rst.Fields(amol)
                    obj.Text = rst.Fields(amol)
                End If
            End If
        Next
    End If
    Set rst = Nothing
End Sub

Private Sub cmbswg_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub cmdadd_Click()
    prcenable
    cmdAdd.Enabled = False
    cmdModify.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
    cmdModify.Tag = ""
    prcclear
    cmbswg.Text = ""
    cmbswg.Locked = True
End Sub
Private Sub prcenable()
    For Each obj In Me
        If TypeOf obj Is TextBox Then
            obj.Locked = False
        End If
    Next
End Sub

Private Sub prcdisble()
    For Each obj In Me
        If TypeOf obj Is TextBox Then
            obj.Locked = True
        End If
    Next
End Sub

Private Sub cmdcancel_Click()
    cmdModify.Tag = ""
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    prcdisble
    prcclear
    cmbswg.Locked = False
    txtt_swg.Locked = False
End Sub

Private Sub cmddelete_Click()

On Error GoTo err
    
    Dim rst As New ADODB.Recordset
    If txtt_swg = "" Then
        MsgBox "Select SWG To Delete", vbInformation
        Exit Sub
    End If
    Dim joel As Integer
    joel = MsgBox("Are you sure you want to delete this record", vbQuestion + vbYesNo)
    
    If joel = vbNo Then Exit Sub
    
    Set rst = New ADODB.Recordset
    rst.Open "delete from secconductor where t_swg = '" & Trim(txtt_swg) & "' ", con, adOpenKeyset, adLockReadOnly
    Set rst = Nothing
    prcclear
    fillswg
    
err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If
End Sub

Private Sub cmdmodify_Click()
    If cmbswg.Text = "" Then
        MsgBox "Select Swg", vbInformation
        Exit Sub
    End If
    
    cmdModify.Tag = "MOD"
    cmdAdd.Enabled = False
    cmdModify.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
    prcenable
    txtt_swg.Locked = True
End Sub

Private Sub cmdSave_Click()

On Error GoTo err

    Dim rst As New ADODB.Recordset
    For Each obj In Me
        If TypeOf obj Is TextBox Then
            If Len(Trim(obj.Text)) = 0 Then
                MsgBox "Enter Proper Details", vbInformation
                obj.SetFocus
                Exit Sub
            End If
        End If
    Next
    
    If txtt_swg = "" Then
        MsgBox "Enter SWG", vbInformation
        Exit Sub
    End If
    
'    Set rst = New ADODB.Recordset
'    rst.Open "select * from secconductor where t_swg = '" & Trim(txtt_swg) & "' ", con, adOpenKeyset, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        MsgBox "this SWG already exists", vbInformation
'        Set rst = Nothing
'        Exit Sub
'    End If
    
    
    If cmdModify.Tag = "MOD" Then
        Set rst = New ADODB.Recordset
        rst.Open "delete from secconductor where t_swg = '" & Trim(cmbswg) & "' ", con, adOpenKeyset, adLockReadOnly
        Set rst = Nothing
        cmdModify.Tag = ""
    End If

    Set rst = New ADODB.Recordset
    rst.Open "select * from secConductor", con, adOpenDynamic, adLockOptimistic
    rst.AddNew
    For Each obj In Me
        If TypeOf obj Is TextBox Then
                amol = ""
                If Len(obj.Text) <> 0 Then
                    amol = Mid(obj.Name, 4, Len(obj.Name))
                    rst.Fields(amol) = obj.Text
                End If
        End If
    Next
    rst.Update
    MsgBox "Record Saved", vbInformation
    
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    prcclear
    prcdisble
    fillswg
    cmbswg.Locked = False
    txtt_swg.Locked = False
    
err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If
End Sub


Private Sub fillswg()
    Dim rst As New ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select distinct t_swg from secconductor", con, adOpenKeyset, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        rst.MoveFirst
        cmbswg.Clear
        Do Until rst.EOF
            If Not IsNull(rst.Fields(0)) Then
                cmbswg.AddItem rst.Fields(0)
            End If
            rst.MoveNext
        Loop
    End If
    Set rst = Nothing
End Sub

Private Sub Form_Activate()
    Me.Left = (MDIDesign.ScaleWidth - Me.ScaleWidth) / 2
    Me.Top = (MDIDesign.ScaleHeight - Me.ScaleHeight) / 2
End Sub

Private Sub Form_GotFocus()
    MDIDesign.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub Form_Load()
    prcdisble
    fillswg
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    MDIDesign.StatusBar1.Panels(2).Text = "Secondary Details"
End Sub
Private Sub prcclear()
    For Each obj In Me
        If TypeOf obj Is TextBox Then
            obj.Text = ""
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIDesign.mnuSecCon.Enabled = True
    MDIDesign.Toolbar1.Buttons(5).ButtonMenus(4).Enabled = True
    MDIDesign.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub txtt_cs_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtt_odn_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtt_ods_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtt_part_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtt_swg_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub
