VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form COSTMASTERfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cost Master"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10605
   Begin VB.TextBox txtDetails 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   945
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   795
      Width           =   6435
   End
   Begin VB.TextBox txtQty 
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
      Height          =   315
      Left            =   8685
      TabIndex        =   3
      Top             =   915
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -60
      TabIndex        =   12
      Top             =   5790
      Width           =   10635
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   5
      Top             =   5250
      Width           =   645
   End
   Begin VB.CommandButton cmdIn 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   4
      Top             =   1500
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Details :"
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
      Height          =   4305
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9855
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4005
         Left            =   60
         TabIndex        =   9
         Top             =   225
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7064
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
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7425
      TabIndex        =   20
      Top             =   465
      Width           =   195
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   9930
      TabIndex        =   19
      Top             =   990
      Width           =   195
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7455
      TabIndex        =   18
      Top             =   120
      Width           =   195
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   7740
      TabIndex        =   17
      Top             =   60
      Width           =   2985
   End
   Begin MSForms.ComboBox cmbProduct 
      Height          =   315
      Left            =   945
      TabIndex        =   16
      Top             =   60
      Width           =   6480
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "11430;556"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
      Object.Width           =   "3527;705"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product :"
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
      Left            =   240
      TabIndex        =   15
      Top             =   75
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Details :"
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
      Left            =   -30
      TabIndex        =   14
      Top             =   810
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8085
      TabIndex        =   13
      Top             =   915
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   -150
      X2              =   10620
      Y1              =   1350
      Y2              =   1350
   End
   Begin MSForms.CommandButton cmdadd 
      Height          =   450
      Left            =   780
      TabIndex        =   11
      ToolTipText     =   "Press to new Commom Core Size "
      Top             =   5250
      Visible         =   0   'False
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
   Begin MSForms.CommandButton cmdsave 
      Height          =   450
      Left            =   3930
      TabIndex        =   7
      ToolTipText     =   "Press Save Commom Core Size "
      Top             =   5955
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
   Begin MSForms.CommandButton cmdcancel 
      Height          =   450
      Left            =   9180
      TabIndex        =   8
      ToolTipText     =   "Press Cancel Commom Core Size "
      Top             =   5940
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
   Begin MSForms.CommandButton cmdmod 
      Height          =   450
      Left            =   60
      TabIndex        =   6
      ToolTipText     =   "Modify Commom Core Size "
      Top             =   5940
      Width           =   1335
      Caption         =   "Edit"
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
   Begin MSForms.ComboBox cmbItem 
      Height          =   315
      Left            =   930
      TabIndex        =   1
      Top             =   420
      Width           =   6435
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "11351;556"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1763;5291"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   330
      TabIndex        =   10
      Top             =   420
      Width           =   555
   End
End
Attribute VB_Name = "COSTMASTERfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As Recordset
Dim MODI As Boolean

Private Sub cmbItem_Click()
    If cmbItem.ListIndex = -1 Then Exit Sub
    txtDetails = cmbItem.List(cmbItem.ListIndex, 1)
End Sub

Private Function TakeItemData(item_code As String) As String
    Dim rs As Recordset
    Set rs = New Recordset
    rs.Open "select * from itemmast where item_code = '" & item_code & "'", con, adOpenStatic, adLockReadOnly
    If rs.EOF = False And rs.BOF = False Then
        TakeItemData = rs!item_desc1 & rs!item_desc2 & rs!item_desc3 & rs!item_desc4 & vbNullString
    End If
    Set rs = Nothing
End Function

Private Function FindDuplicate() As Boolean
    
    For i = 1 To Grid.Rows - 1
        If cmbItem.List(cmbItem.ListIndex, 0) = Grid.TextMatrix(i, 1) Then
            MsgBox "This item already exits", vbInformation
            FindDuplicate = True
            Exit Function
        End If
    Next
    FindDuplicate = False
End Function

Private Sub cmbType_Click()
    Grid.Rows = 1
    Set rst = New Recordset
    rst.Open "select * from COSTMASTER_VIEW where name ='" & cmbType.List(cmbType.ListIndex, 1) & "' order by sr", con, adOpenDynamic, adLockOptimistic
    If rst.EOF = False And rst.BOF = False Then
        rst.MoveFirst
        Do Until rst.EOF
            Grid.Rows = Grid.Rows + 1
            Grid.row = Grid.Rows - 1
            Grid.RowHeight(Grid.row) = 300
            Grid.TextMatrix(Grid.row, 0) = rst!sr & vbNullString
            Grid.TextMatrix(Grid.row, 1) = rst!item_code & vbNullString
            Grid.TextMatrix(Grid.row, 2) = rst!item_desc1 & vbNullString
            Grid.TextMatrix(Grid.row, 3) = rst!qty & vbNullString
            Grid.TextMatrix(Grid.row, 4) = rst!Name & vbNullString
            rst.MoveNext
        Loop
    End If
    Set rst = Nothing
End Sub

Private Sub cmbType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub FillProductDetails()
Set rst = New ADODB.Recordset
rst.Open "select * from V_COSTMASTER_VIEW where pro_code= '" & cmbProduct & "' order by sr", con, adOpenStatic, adLockReadOnly
If rst.BOF = False And rst.EOF = False Then
    rst.MoveFirst
    Do Until rst.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.row = Grid.Rows - 1
        Grid.RowHeight(Grid.row) = 500
        Grid.TextMatrix(Grid.row, 0) = rst!sr
        Grid.TextMatrix(Grid.row, 1) = rst!item_code & vbNullString
        Grid.TextMatrix(Grid.row, 2) = rst!PRODUCT_DESCRIPTION
        Grid.TextMatrix(Grid.row, 3) = rst!qty
        rst.MoveNext
'        Grid.TextMatrix(Grid.row, 4) = rst!Name & vbNullString
    Loop
End If
End Sub

Private Sub cmbProduct_Click()
Grid.Rows = 1
Label4.Caption = ""
If cmbProduct.ListIndex >= 0 Then
    Label4.Caption = cmbProduct.List(cmbProduct.ListIndex, 1)
    FillProductDetails
End If

End Sub

Private Sub cmdcancel_Click()
    prcclear
    prclock
    MODI = False
    Grid.Tag = ""
End Sub

Private Sub cmdin_Click()
'    If IsNumeric(txtQty.Text) = False Then
'        MsgBox "Please Enter only Numeric value", vbInformation
'    End If
    If cmbItem.ListIndex = -1 Then
        MsgBox "Please select item code!!!", vbInformation
        cmbItem.SetFocus
        Exit Sub
    End If
    If Trim(txtDetails) = "" Then
        MsgBox "Please select item code!!!", vbInformation
        cmbItem.SetFocus
        Exit Sub
    End If
    If Val(txtQty) = 0 Then
        MsgBox "Please enter item quantity!!!", vbInformation
        txtQty.SetFocus
        Exit Sub
    End If
    If Val(txtQty) > 10000 Then
        MsgBox "Please enter item quantity less than 10000!!!", vbInformation
        txtQty.SetFocus
        Exit Sub
    End If

'    If Val(txtQty) = 0 Then
'        MsgBox "Please enter item quantity!!!", vbInformation
'        txtQty = 0
''        Exit Sub
'    End If
    
    If Grid.Tag <> "MOD" Then
        If FindDuplicate Then Exit Sub
    End If
    
    If Grid.Tag <> "MOD" Then
        Grid.Rows = Grid.Rows + 1
        Grid.row = Grid.Rows - 1
        Grid.RowHeight(Grid.row) = 500
    End If
    
    Grid.TextMatrix(Grid.row, 0) = Grid.row
    Grid.TextMatrix(Grid.row, 1) = cmbItem.List(cmbItem.ListIndex, 0)
    Grid.TextMatrix(Grid.row, 2) = Trim(txtDetails)
    If Trim(txtQty.Text) <> "" Then
        Grid.TextMatrix(Grid.row, 3) = Trim(txtQty)
    Else
        Grid.TextMatrix(Grid.row, 3) = 0
    End If
    'Grid.TextMatrix(Grid.row, 4) = cmbType.List(cmbType.ListIndex, 1)
    
    cmbItem.Text = ""
    cmbItem.SetFocus
    txtDetails.Text = ""
    Grid.Tag = ""
    
End Sub

Private Sub gridhead()
    
    Grid.Rows = 1
    Grid.Cols = 4
    Grid.TextMatrix(0, 0) = "Sr"
    Grid.ColWidth(0) = 500
    Grid.TextMatrix(0, 1) = "ItemCode"
    Grid.ColWidth(1) = 1000
    Grid.TextMatrix(0, 2) = "Description"
    Grid.ColWidth(2) = 5000
    Grid.TextMatrix(0, 3) = "Qty"
    Grid.ColWidth(3) = 900
'    Grid.TextMatrix(0, 4) = "Type"
'    Grid.ColWidth(4) = 800

    
End Sub

Private Sub FillItem()
    Dim ctr As Integer
    'product code
    Set rst = New Recordset
    rst.Open "select PRODUCT_CODE,PRODUCT_DESCRIPTION from mst_part where product_code not like '1201%'  order by product_code", con, adOpenStatic, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        rst.MoveFirst
        Do Until rst.EOF
            cmbItem.AddItem rst.Fields(0) & vbNullString
            cmbItem.List(ctr, 1) = rst.Fields(1) '& rst.Fields(2) & rst.Fields(3) & rst.Fields(4) & vbNullString
            ctr = ctr + 1
            rst.MoveNext
        Loop
    End If
    Set rst = Nothing

End Sub

Private Sub FillProduct()
    
    Dim ctr As Integer
    cmbProduct.Clear
    Set rst = New ADODB.Recordset
    rst.Open "select product_code,product_description from MST_PART where product_code like '1201%'", con, adOpenStatic, adLockReadOnly
    If rst.BOF = False And rst.EOF = False Then
        rst.MoveFirst
        Do Until rst.EOF
            cmbProduct.AddItem rst.Fields(0)
            cmbProduct.List(ctr, 1) = rst.Fields(1)
            ctr = ctr + 1
            rst.MoveNext
        Loop
    End If
    Set rst = Nothing
    
    
    
End Sub

Private Sub prcunlock()
    
    cmdAdd.Enabled = False
    cmdmod.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    Frame1.Enabled = True
    cmdIn.Enabled = True
    cmdOut.Enabled = True
'    txtDetails.Enabled = True
    cmbItem.Enabled = True
    txtQty.Enabled = True
'    cmbType.Enabled = False

End Sub

Private Sub prclock()

    cmdAdd.Enabled = True
    cmdmod.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    Frame1.Enabled = False
    cmdIn.Enabled = False
    cmdOut.Enabled = False
    txtDetails.Enabled = False
    cmbItem.Enabled = True
    txtQty.Enabled = False
'    cmbType.Enabled = True

End Sub

Private Sub prcclear()
    cmbItem.ListIndex = -1
    cmbProduct.ListIndex = -1
    Grid.Rows = 1
    txtDetails = ""
    txtQty = "0.000"
    Label4.Caption = ""
'        FillItem
'    cmbType.Clear
'    cmbType.AddItem "Window/Bar"
'    cmbType.List(0, 1) = "1"
'    cmbType.AddItem "Wound"
'    cmbType.List(1, 1) = "2"
End Sub

Private Sub cmdMod_Click()
If cmbProduct.ListIndex = -1 Then
    MsgBox "Select a product!!!", vbInformation
    cmbProduct.SetFocus
    Exit Sub
End If
If Grid.Rows = 1 Then
    MODI = False
Else
    MODI = True
End If
prcunlock
End Sub

'Private Sub cmdMod_Click()
''    prcClear
'    If cmbType.ListIndex = -1 Then
'        MsgBox "Please select the CT Type!!!", vbExclamation
'        Exit Sub
'    End If
'    prcunlock
'    MODI = True
'    cmbItem.SetFocus
'
'End Sub

Private Sub cmdout_Click()
    If Grid.Rows = 1 Then Exit Sub
    Dim Ans As Integer
    Dim i As Integer
    
'    If Grid.RowSel = False Then
'        MsgBox "Plz select ", vbInformation
'    End If
'
    Ans = MsgBox("Are you sure you want to delete this item?", vbQuestion + vbYesNo)
    If Ans = vbNo Then Exit Sub
    
    
    
    If Grid.Rows = 2 Then
        Grid.Rows = 1
    Else
        Grid.RemoveItem (Grid.RowSel)
    End If
    
    For i = 1 To Grid.Rows - 1
        Grid.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    If Grid.Rows = 1 Then
        MsgBox "No data Added!!!", vbExclamation
        Exit Sub
    End If
    MousePointer = vbHourglass
'    con.BeginTrans
    If MODI = True Then
        Set rst = New Recordset
        rst.Open "delete from new_costmaster where product_code = '" & cmbProduct.List(cmbProduct.ListIndex, 0) & "'", con, adOpenDynamic, adLockOptimistic
        Set rst = Nothing
    End If
    
    Set rst = New Recordset
    rst.Open "select * from new_costmaster", con, adOpenDynamic, adLockOptimistic
    For i = 1 To Grid.Rows - 1
        rst.AddNew
        rst!sr = i
        rst!item_code = Grid.TextMatrix(i, 1)
        rst!qty = Grid.TextMatrix(i, 3)
        rst!product_code = cmbProduct.List(cmbProduct.ListIndex, 0)
        rst.Update
    Next
    
    MsgBox "Data saved...", vbInformation
    prcclear
    prclock
    MODI = False
'    con.CommitTrans
    Grid.Tag = ""
err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description, vbCritical
'        con.RollbackTrans
    End If
    MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
'    Me.Left = (MDIDesign.ScaleWidth - Me.ScaleWidth) / 2
'    Me.Top = (MDIDesign.ScaleHeight - Me.ScaleHeight) / 2
End Sub

Private Sub Form_Load()
    gridhead
    FillItem
    FillProduct
    prcclear
    prclock
    Grid.Tag = ""
End Sub

Private Sub Grid_Click()
    
    cmbItem.Text = Grid.TextMatrix(Grid.row, 1)
    txtQty = Grid.TextMatrix(Grid.row, 3)
    Grid.Tag = "MOD"
    txtQty.SetFocus
    
End Sub

Private Sub txtQty_GotFocus()
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If Not Chr(KeyAscii) Like "[0-9.]" And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtQty_LostFocus()
    If IsNumeric(txtQty) = False Then
        MsgBox "Enter numeric value!!!", vbInformation
        txtQty = 0
    End If
    If Len(Trim(txtQty)) > 0 Then
        txtQty = Format(txtQty, "0.000")
    End If
End Sub

