VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CoreSizefrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Common Core Dimension"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "CoreSizefrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtnsv 
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
      Left            =   675
      TabIndex        =   13
      Top             =   855
      Width           =   1500
   End
   Begin VB.Frame frmdim 
      Caption         =   "Dimension :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   2250
      TabIndex        =   8
      Top             =   45
      Width           =   5055
      Begin VB.TextBox txtWindowL 
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
         Left            =   2295
         TabIndex        =   17
         Top             =   720
         Width           =   825
      End
      Begin VB.TextBox txtWindowW 
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
         Left            =   3870
         TabIndex        =   16
         Top             =   720
         Width           =   825
      End
      Begin VB.TextBox txtOAllW 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3870
         TabIndex        =   10
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox txtOAllL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2295
         TabIndex        =   9
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   3465
         TabIndex        =   19
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "WINDOW (OD X ID) :"
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
         Left            =   225
         TabIndex        =   18
         Top             =   720
         Width           =   2595
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   3465
         TabIndex        =   15
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "OVER ALL (OD X ID) :"
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
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Width           =   2010
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4110
      Left            =   45
      TabIndex        =   6
      Top             =   1665
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   7250
      _Version        =   393216
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbnsv 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   675
      TabIndex        =   0
      Top             =   135
      Width           =   1530
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Common Core Dimension"
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
      Height          =   285
      Left            =   405
      TabIndex        =   20
      Top             =   1305
      Width           =   6405
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   45
      Top             =   1260
      Width           =   7260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "NSV :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   14
      Top             =   855
      Width           =   795
   End
   Begin MSForms.CommandButton cmdmod 
      Height          =   450
      Left            =   1710
      TabIndex        =   12
      ToolTipText     =   "Modify Commom Core Size "
      Top             =   5850
      Width           =   1335
      Caption         =   "Modify"
      PicturePosition =   327683
      Size            =   "2355;794"
      Accelerator     =   77
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdcancel 
      Height          =   450
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Press Cancel Commom Core Size "
      Top             =   5850
      Width           =   1335
      Caption         =   "Cancel"
      PicturePosition =   327683
      Size            =   "2355;794"
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
      Height          =   450
      Left            =   3375
      TabIndex        =   4
      ToolTipText     =   "Press Save Commom Core Size "
      Top             =   5850
      Width           =   1335
      Caption         =   "Save"
      PicturePosition =   327683
      Size            =   "2355;794"
      Accelerator     =   115
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdadd 
      Height          =   450
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Press to new Commom Core Size "
      Top             =   5850
      Width           =   1335
      Caption         =   "Add"
      PicturePosition =   327683
      Size            =   "2355;794"
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
      Caption         =   "NSV :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -180
      TabIndex        =   5
      Top             =   135
      Width           =   795
   End
   Begin MSForms.CommandButton cmdin 
      Height          =   495
      Left            =   6390
      TabIndex        =   2
      ToolTipText     =   "Press to pull in the data"
      Top             =   1665
      Width           =   930
      Caption         =   "IN << "
      PicturePosition =   327683
      Size            =   "1640;873"
      Accelerator     =   83
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmddelete 
      Height          =   495
      Left            =   6390
      TabIndex        =   3
      ToolTipText     =   "Press to remove in the data"
      Top             =   5265
      Width           =   930
      Caption         =   "Out >>"
      PicturePosition =   327683
      Size            =   "1640;873"
      Accelerator     =   68
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "CoreSizefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim row As Integer

Private Sub cmbnsv_Click()
    filldata
    cmdCancel.Enabled = True
End Sub

Private Sub cmbnsv_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Select Nominal System Voltage"

End Sub

Private Sub cmdadd_Click()
    setgrid
    cmbNSV.Text = ""
    cmbNSV.Enabled = False
    cmdAdd.Enabled = False
    cmdSave.Enabled = True
    frmdim.Enabled = True
    txtnsv.SetFocus
    cmdCancel.Enabled = True
    cmdAdd.Tag = "ADD"
End Sub

Private Sub cmdcancel_Click()
    setgrid
    fillnsv
    cmdSave.Enabled = False
    cmdAdd.Enabled = True
    txtWindowL.Text = ""
    txtWindowW.Text = ""
    txtOAllW.Text = ""
    txtOAllL.Text = ""
    txtnsv.Text = ""
    cmbNSV.Enabled = True
    cmddelete.Enabled = False
    cmdCancel.Enabled = False
    cmdmod.Enabled = False
    cmdmod.Tag = ""
    Grid.Tag = ""
    cmdAdd.Tag = ""
    cmbNSV.SetFocus
End Sub

Private Sub cmddelete_Click()
    If Grid.row <= 0 Then
        
        MsgBox "There is no rows to be Deleted", vbInformation
        cmdSave.Enabled = False
        Exit Sub
        
    End If
        
    If Grid.Rows >= 1 Then
    
            If Grid.Rows = 2 Then
                Grid.Rows = 1
                Grid.Cols = 6
                cmdSave.Enabled = False
            Else
                Grid.RemoveItem (Grid.RowSel)
                cmdSave.Enabled = True
                Exit Sub
            End If

    Else
        MsgBox "Please Select The Row You Want To Delete", vbInformation
        Grid.Enabled = True
    End If
        
  
    
cmdSave.Enabled = True
End Sub

Private Sub cmdin_Click()
    If txtWindowL.Text = "" Or txtWindowW.Text = "" Or txtOAllW.Text = "" Or txtOAllL.Text = "" Then
        MsgBox "Enter Proper Core Details", vbInformation
        Exit Sub
    End If
    
    Dim a As Integer
    If Grid.Tag <> "MOD" Then
        Grid.Rows = Grid.Rows + 1
        a = Grid.Rows - 1
    Else
        a = row
    End If
    Grid.TextMatrix(a, 0) = Grid.row
    Grid.TextMatrix(a, 1) = txtnsv
    Grid.TextMatrix(a, 2) = txtWindowL
    Grid.TextMatrix(a, 3) = txtWindowW
    Grid.TextMatrix(a, 4) = txtOAllL
    Grid.TextMatrix(a, 5) = txtOAllW
    row = 0
    txtnsv.SetFocus
    txtnsv.Text = ""
    txtWindowL.Text = ""
    txtWindowW.Text = ""
    txtOAllW.Text = ""
    txtOAllL.Text = ""
    Grid.Tag = ""
    If Grid.Rows > 0 Then cmdSave.Enabled = True
    
End Sub
Private Sub cmdMod_Click()
    cmdmod.Tag = "MOD"
    cmdmod.Enabled = False
    cmdSave.Enabled = True
    cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdAdd.Tag = ""
End Sub

Private Sub cmdSave_Click()
    If Grid.Rows = 1 Then
        MsgBox "No Records To Save", vbInformation
        Exit Sub
    End If
    Dim rst As New ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "delete from coresize where nsv = '" & cmbNSV & "'", con, adOpenDynamic, adLockOptimistic
    Set rst = Nothing
    
    For i = 1 To Grid.Rows - 1
        Set rst = New ADODB.Recordset
        rst.Open "select * from coresize", con, adOpenDynamic, adLockOptimistic
        rst.AddNew
        rst!NSV = Grid.TextMatrix(i, 1)
        rst!windowl = Grid.TextMatrix(i, 2)
        rst!windoww = Grid.TextMatrix(i, 3)
        rst!oallL = Grid.TextMatrix(i, 4)
        rst!oallw = Grid.TextMatrix(i, 5)
        rst.Update
        Set rst = Nothing
    Next
    MsgBox "Record Saved", vbInformation
    setgrid
    fillnsv
    cmbNSV.Enabled = True
    cmdSave.Enabled = False
    cmdAdd.Enabled = True
    cmdmod.Enabled = False
    txtWindowL.Text = ""
    txtWindowW.Text = ""
    txtOAllW.Text = ""
    txtOAllL.Text = ""
    txtnsv.Text = ""
    frmdim.Enabled = False
    cmdmod.Tag = ""
    Grid.Tag = ""
    cmdAdd.Tag = ""
End Sub

Private Sub Form_Activate()
    Me.Left = (MDIDesign.ScaleWidth - Me.ScaleWidth) / 2
    Me.Top = (MDIDesign.ScaleHeight - Me.ScaleHeight) / 2
End Sub

Private Sub Form_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub Form_Load()
    setgrid
    fillnsv
    cmdSave.Enabled = False
    cmddelete.Enabled = False
    frmdim.Enabled = False
    cmdCancel.Enabled = False
    cmdmod.Enabled = False
End Sub

Private Sub fillnsv()
    Dim rst As New ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select distinct nsv from coresize order by nsv", con, adOpenKeyset, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        rst.MoveFirst
        cmbNSV.Clear
        Do Until rst.EOF
            If Not IsNull(rst.Fields(0)) Then
                cmbNSV.AddItem rst.Fields(0)
            End If
            rst.MoveNext
        Loop
    End If
    Set rst = Nothing
End Sub


Private Sub filldata()
    Dim rst As New ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select * from coresize where nsv = " & cmbNSV & "", con, adOpenKeyset, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        rst.MoveFirst
        Grid.Rows = 1
        For i = 1 To rst.RecordCount
            Grid.Rows = Grid.Rows + 1
            Grid.TextMatrix(i, 0) = i
            Grid.TextMatrix(i, 1) = rst!NSV
            Grid.TextMatrix(i, 2) = rst!windowl
            Grid.TextMatrix(i, 3) = rst!windoww
            Grid.TextMatrix(i, 4) = rst!oallL
            Grid.TextMatrix(i, 5) = rst!oallw
            rst.MoveNext
        Next
        cmddelete.Enabled = True
        cmdmod.Enabled = True
    End If
End Sub


Public Sub setgrid()
    Grid.Rows = 1
    Grid.Cols = 6
    Grid.TextMatrix(0, 0) = "Sr"
    Grid.ColWidth(0) = 400
    Grid.TextMatrix(0, 1) = "NSV"
    Grid.TextMatrix(0, 2) = "WindowL"
    Grid.TextMatrix(0, 3) = "WindowW"
    Grid.TextMatrix(0, 4) = "OAllL"
    Grid.TextMatrix(0, 5) = "OAllW"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdAdd.Tag = "ADD" Or cmdmod.Tag = "MOD" Then
        joel = MsgBox("Are you sure you want to exit ", vbQuestion + vbYesNo)
        If joel = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIDesign.Toolbar1.Buttons(7).Enabled = True
    MDIDesign.mnuComCore.Enabled = True
    MDIDesign.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub Grid_DblClick()
    If cmdmod.Tag = "MOD" Then
        If Grid.row = 0 Then Exit Sub
        txtnsv = Grid.TextMatrix(Grid.RowSel, 1)
        txtWindowL = Grid.TextMatrix(Grid.RowSel, 2)
        txtWindowW = Grid.TextMatrix(Grid.RowSel, 3)
        txtOAllL = Grid.TextMatrix(Grid.RowSel, 4)
        txtOAllW = Grid.TextMatrix(Grid.RowSel, 5)
        Grid.Tag = "MOD"
        row = Grid.RowSel
        cmdSave.Enabled = True
        frmdim.Enabled = True
        For i = 1 To Grid.Cols - 1
            Grid.TextMatrix(Grid.RowSel, i) = "Modifying..."
        Next
    End If
End Sub



Private Sub txtnsv_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Enter Nominal System Voltage"
End Sub

Private Sub txtOAllL_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Core Size"

End Sub

Private Sub txtOAllL_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtOAllW_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Core Size"
End Sub

Private Sub txtOAllW_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtWindowL_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Core Size"
End Sub

Private Sub txtWindowL_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtWindowW_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Core Size"
End Sub

Private Sub txtWindowW_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub
