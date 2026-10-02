VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form CommandDetailfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCI 
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
      Height          =   330
      Left            =   2115
      TabIndex        =   5
      Top             =   150
      Width           =   2985
   End
   Begin VB.TextBox txtII 
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
      Height          =   330
      Left            =   2100
      TabIndex        =   4
      Top             =   555
      Width           =   2985
   End
   Begin VB.TextBox txtDPN 
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
      Height          =   330
      Left            =   2100
      TabIndex        =   3
      Top             =   960
      Width           =   2985
   End
   Begin VB.TextBox txtDPSC 
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
      Height          =   330
      Left            =   2115
      TabIndex        =   2
      Top             =   1365
      Width           =   2985
   End
   Begin VB.TextBox txtSF 
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
      Height          =   330
      Left            =   2115
      TabIndex        =   1
      Top             =   1770
      Width           =   2985
   End
   Begin VB.TextBox Text8 
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
      Height          =   330
      Left            =   2115
      TabIndex        =   0
      Top             =   2175
      Width           =   2985
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   465
      Left            =   975
      TabIndex        =   13
      Top             =   3525
      Width           =   1740
      Size            =   "3069;820"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   465
      Left            =   975
      TabIndex        =   12
      Top             =   2925
      Width           =   1740
      Size            =   "3069;820"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current I :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -1395
      TabIndex        =   11
      Top             =   150
      Width           =   3435
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "II :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -1575
      TabIndex        =   10
      Top             =   555
      Width           =   3615
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Delta P N  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -1530
      TabIndex        =   9
      Top             =   960
      Width           =   3570
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Delta P S. C. "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   -1575
      TabIndex        =   8
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stacking Factor :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   -1440
      TabIndex        =   7
      Top             =   1770
      Width           =   3480
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BUILD UP T Cm :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -1575
      TabIndex        =   6
      Top             =   2175
      Width           =   3615
   End
End
Attribute VB_Name = "CommandDetailfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As Recordset

Private Sub cmdadd_Click()

PrcEnabled

End Sub

Private Sub cmdSave_Click()

Set rst = New Recordset
rst.Open "select * from commandb", con, adOpenDynamic, adLockOptimistic

If rst.EOF = False And rst.BOF = False Then

    rst!ci = Trim(txtCI.Text)
    rst!ii = Trim(txtII.Text)
    rst!delPn = Trim(txtDPN.Text)
    rst!delPsc = Trim(txtDPSC.Text)
    rst!sf = Trim(txtSF.Text)
    
    rst.Update

End If

rst.Close
Set rst = Nothing

PrcDisabled

End Sub

Private Sub Form_Load()

On Error Resume Next

prcConnect

PrcDisabled

Set rst = New Recordset
rst.Open "select * from commandb", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then

    txtCI.Text = rst!ci
    txtII.Text = rst!ii
    txtDPN.Text = rst!delPn
    txtDPSC.Text = rst!delPsc
    txtSF.Text = rst!sf
    
Else

    MsgBox "Command Transformer Details not set", vbCritical

End If


End Sub

Public Sub PrcEnabled()

For Each Control In CommandDetailfrm

    If TypeOf Control Is TextBox Then
    
        Control.Enabled = True
            
    End If

Next

End Sub

Public Sub PrcDisabled()

For Each Control In CommandDetailfrm

    If TypeOf Control Is TextBox Then
    
        Control.Enabled = False
            
    End If

Next

End Sub

