VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CTfrm 
   Caption         =   "H.T.CURRENT TRANSFORMER Designing"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11835
   ForeColor       =   &H8000000F&
   Icon            =   "CTfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   11835
   Begin TabDlg.SSTab SSTab1 
      Height          =   7170
      Left            =   45
      TabIndex        =   81
      Top             =   45
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   12647
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "C.T. Specifications [&1]"
      TabPicture(0)   =   "CTfrm.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdIN"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdRemCore"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label65"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label16"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label32"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label37"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Shape5"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label75"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label76"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdCancel"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CommandButton1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDesNo"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtRatio(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Frame5"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Frame2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtRateConti"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtRatio(4)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtRatioR2(4)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtRatio(3)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtRatioR2(3)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtRatioR2(2)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtRatio(1)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtRatioR2(1)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtRatioR2(0)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtRatio(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmbFreq"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmbTime"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmbSTC"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmbIL"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmbHSV"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtnoofcores"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmbEquipType"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmbNSV"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtRatioR2Dum(0)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtRatioDum(1)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtRatioR2Dum(1)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtRatioDum(2)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtRatioR2Dum(2)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtRatioDum(3)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtRatioR2Dum(3)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtRatioDum(4)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtRatioR2Dum(4)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtRateContiDum"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtRatioDum(0)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).ControlCount=   56
      TabCaption(1)   =   "Selection of Core(s)[&2]"
      TabPicture(1)   =   "CTfrm.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame17"
      Tab(1).Control(1)=   "Frame11"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(3)=   "cmbStyle"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Frame10"
      Tab(1).Control(6)=   "Label35"
      Tab(1).Control(7)=   "Shape2"
      Tab(1).Control(8)=   "cmdICore"
      Tab(1).Control(9)=   "cmdRemoCore"
      Tab(1).Control(10)=   "Label30"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Primary Details[&3]"
      TabPicture(2)   =   "CTfrm.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape6"
      Tab(2).Control(1)=   "Shape1"
      Tab(2).Control(2)=   "Label38"
      Tab(2).Control(3)=   "Label39"
      Tab(2).Control(4)=   "Label40"
      Tab(2).Control(5)=   "Label42"
      Tab(2).Control(6)=   "Label43"
      Tab(2).Control(7)=   "Label44"
      Tab(2).Control(8)=   "Label45"
      Tab(2).Control(9)=   "Label46"
      Tab(2).Control(10)=   "Label47"
      Tab(2).Control(11)=   "Label48"
      Tab(2).Control(12)=   "Label49"
      Tab(2).Control(13)=   "Label50"
      Tab(2).Control(14)=   "Label51"
      Tab(2).Control(15)=   "Label41"
      Tab(2).Control(16)=   "Shape3"
      Tab(2).Control(17)=   "Label31"
      Tab(2).Control(18)=   "cmbConStyle"
      Tab(2).Control(19)=   "txtAmper"
      Tab(2).Control(20)=   "txtNominal"
      Tab(2).Control(21)=   "txtPriCros"
      Tab(2).Control(22)=   "txtDelStc"
      Tab(2).Control(23)=   "txtTop"
      Tab(2).Control(24)=   "txtSide"
      Tab(2).Control(25)=   "txtTotalCon"
      Tab(2).Control(26)=   "txtCondPer"
      Tab(2).Control(27)=   "txtWidth"
      Tab(2).Control(28)=   "txtminCs"
      Tab(2).Control(29)=   "txtIDyn"
      Tab(2).Control(30)=   "txtIDATurns"
      Tab(2).Control(31)=   "txtDelR"
      Tab(2).Control(32)=   "txtTurns"
      Tab(2).Control(33)=   "lstCond"
      Tab(2).ControlCount=   34
      TabCaption(3)   =   "Secondary Details[&4]"
      TabPicture(3)   =   "CTfrm.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Dimensions[&5]"
      TabPicture(4)   =   "CTfrm.frx":037A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label59"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "WindowSize"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label64"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label91"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label90"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label87"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label86"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label85"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label60"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Label92"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Frame21"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "CrystalReport2"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Frame15"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "wounded"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "FrameWound"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "txtdrg"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "cmbType"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "Frame14"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "Frame16"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "cmdsave"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "Frame22"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "CrystalReport1"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "cmbCal"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "cmdS1"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "cmdS2"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "cmdS3"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "cmdS4"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "txtPRiRes75"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "txtWattPri"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "txtPRiRes20"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "CMBINSULATION"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).ControlCount=   31
      Begin VB.ListBox lstCond 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4905
         Left            =   -67260
         TabIndex        =   270
         Top             =   1650
         Width           =   3015
      End
      Begin VB.ComboBox txtTurns 
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
         Left            =   -71070
         TabIndex        =   269
         Top             =   1410
         Width           =   1125
      End
      Begin VB.ComboBox CMBINSULATION 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "CTfrm.frx":0396
         Left            =   9630
         List            =   "CTfrm.frx":03A6
         TabIndex        =   88
         Top             =   5400
         Width           =   1500
      End
      Begin VB.TextBox txtPRiRes20 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10125
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   2430
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtWattPri 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10125
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   3150
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtPRiRes75 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10125
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   2790
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdS4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10350
         TabIndex        =   92
         ToolTipText     =   "Go to Secondary Details"
         Top             =   6075
         Width           =   1050
      End
      Begin VB.CommandButton cmdS3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9180
         TabIndex        =   91
         ToolTipText     =   "Go to Primary Details"
         Top             =   6075
         Width           =   1050
      End
      Begin VB.CommandButton cmdS2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7965
         TabIndex        =   90
         ToolTipText     =   "Goto Core Size"
         Top             =   6075
         Width           =   1050
      End
      Begin VB.CommandButton cmdS1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6795
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Go to C.T. Specifications  "
         Top             =   6075
         Width           =   1050
      End
      Begin VB.CommandButton cmbCal 
         Caption         =   "Cal."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5715
         TabIndex        =   87
         ToolTipText     =   "Generates specification  "
         Top             =   2385
         Width           =   825
      End
      Begin VB.TextBox txtRatioDum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -67935
         TabIndex        =   10
         Top             =   900
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtRateContiDum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70500
         MaxLength       =   5
         TabIndex        =   8
         Top             =   2115
         Width           =   1185
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7410
         Top             =   6555
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         Connect         =   "odbc-enhance"
         UserName        =   "sa"
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Frame Frame22 
         Caption         =   "Weight :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   6975
         TabIndex        =   249
         Top             =   3465
         Visible         =   0   'False
         Width           =   4290
         Begin VB.TextBox txtResinWt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   1350
            Width           =   960
         End
         Begin VB.TextBox txtPrimaryCopper 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   990
            Width           =   960
         End
         Begin VB.TextBox txtSecCuWt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   630
            Width           =   960
         End
         Begin VB.TextBox txtCoreWT 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   101
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RESIN WEIGHT (kg) :"
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
            Left            =   405
            TabIndex        =   253
            Top             =   1350
            Width           =   2715
         End
         Begin VB.Label Label82 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRIMARY CU WEIGHT (kg) :"
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
            Left            =   135
            TabIndex        =   252
            Top             =   990
            Width           =   2985
         End
         Begin VB.Label Label83 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SECONDARY  CU WEIGHT (kg) :"
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
            Left            =   180
            TabIndex        =   251
            Top             =   630
            Width           =   2940
         End
         Begin VB.Label Label84 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CORE WEIGTH  (kg) :"
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
            Left            =   0
            TabIndex        =   250
            Top             =   270
            Width           =   3120
         End
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   8460
         TabIndex        =   93
         Top             =   6525
         Width           =   1545
      End
      Begin VB.TextBox txtRatioR2Dum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   -64785
         TabIndex        =   19
         Top             =   2340
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtRatioDum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   -67935
         TabIndex        =   18
         Top             =   2340
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtRatioR2Dum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -64785
         TabIndex        =   17
         Top             =   1980
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtRatioDum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -67935
         TabIndex        =   16
         Top             =   1980
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtRatioR2Dum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -64785
         TabIndex        =   15
         Top             =   1620
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtRatioDum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -67935
         TabIndex        =   14
         Top             =   1620
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtRatioR2Dum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -64785
         TabIndex        =   13
         Top             =   1260
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtRatioDum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -67935
         TabIndex        =   12
         Top             =   1260
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtRatioR2Dum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -64785
         TabIndex        =   11
         Top             =   900
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.ComboBox cmbNSV 
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
         Left            =   -74010
         TabIndex        =   2
         Top             =   1305
         Width           =   1140
      End
      Begin VB.Frame Frame16 
         Caption         =   "Window Clearances :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   135
         TabIndex        =   231
         Top             =   5805
         Visible         =   0   'False
         Width           =   4425
         Begin VB.Label WOR1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RESIN CLEARANCE WIDTH WISE (mm) :"
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
            Left            =   -45
            TabIndex        =   265
            Top             =   225
            Width           =   3660
         End
         Begin VB.Label WOR1A 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3780
            TabIndex        =   234
            Top             =   225
            Width           =   510
         End
         Begin VB.Label WOR2A 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3780
            TabIndex        =   233
            Top             =   540
            Width           =   510
         End
         Begin VB.Label WOR2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RESIN CLEARENCE HT. WISE (mm) :"
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
            Left            =   225
            TabIndex        =   232
            Top             =   540
            Width           =   3390
         End
      End
      Begin VB.Frame Frame14 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   135
         TabIndex        =   227
         Top             =   4365
         Width           =   6450
         Begin MSFlexGridLib.MSFlexGrid Casting 
            Height          =   1095
            Left            =   45
            TabIndex        =   106
            Top             =   225
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   1931
            _Version        =   393216
            ForeColor       =   0
            ForeColorFixed  =   192
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
      Begin VB.ComboBox cmbType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "CTfrm.frx":03B8
         Left            =   4590
         List            =   "CTfrm.frx":03C8
         TabIndex        =   86
         Top             =   2430
         Width           =   1050
      End
      Begin VB.TextBox txtdrg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2205
         TabIndex        =   85
         Text            =   "PR-"
         Top             =   2430
         Width           =   1545
      End
      Begin VB.Frame FrameWound 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   135
         TabIndex        =   225
         Top             =   5445
         Width           =   6450
         Begin MSFlexGridLib.MSFlexGrid Mould 
            Height          =   1275
            Left            =   45
            TabIndex        =   107
            Top             =   225
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   2249
            _Version        =   393216
            ForeColor       =   0
            ForeColorFixed  =   192
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
      Begin MSFlexGridLib.MSFlexGrid wounded 
         Height          =   1725
         Left            =   225
         TabIndex        =   224
         Top             =   630
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   3043
         _Version        =   393216
         ForeColor       =   0
         ForeColorFixed  =   192
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
      Begin VB.TextBox txtDelR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   5130
         Width           =   1095
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   6450
         Left            =   -74910
         TabIndex        =   165
         Top             =   360
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   11377
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Core Details"
         TabPicture(0)   =   "CTfrm.frx":03DA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label52"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Shape4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label88"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtTWh"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtCoreType"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame3"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Frame8"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmbSelctCore"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Frame13"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmCal"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "All Core Details"
         TabPicture(1)   =   "CTfrm.frx":03F6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "G1"
         Tab(1).ControlCount=   1
         Begin VB.CommandButton cmCal 
            Caption         =   "&Calculate"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7830
            TabIndex        =   80
            ToolTipText     =   "Calculate core details"
            Top             =   4815
            Width           =   2850
         End
         Begin VB.Frame Frame13 
            Caption         =   "Data Entry :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3930
            Left            =   270
            TabIndex        =   213
            Top             =   2295
            Width           =   6540
            Begin VB.TextBox txtCompen 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   4815
               TabIndex        =   84
               Top             =   3510
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.TextBox txtCompen 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   4050
               TabIndex        =   83
               Top             =   3510
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.TextBox txtCompen 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   3285
               TabIndex        =   82
               Top             =   3510
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.TextBox txtCompen 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   2520
               TabIndex        =   116
               Top             =   3510
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.TextBox txtCompen 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1755
               TabIndex        =   79
               Top             =   3510
               Width           =   690
            End
            Begin VB.TextBox txtGEdit 
               BackColor       =   &H00FEDFC7&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2835
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   214
               Text            =   "CTfrm.frx":0412
               Top             =   2340
               Visible         =   0   'False
               Width           =   2265
            End
            Begin VB.ComboBox cmbMetal 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   315
               ItemData        =   "CTfrm.frx":0418
               Left            =   3780
               List            =   "CTfrm.frx":0428
               TabIndex        =   74
               Top             =   585
               Width           =   2310
            End
            Begin VB.CheckBox ChK2 
               Caption         =   "DO YOU WANT TO SET DIFF. CONDUCTOR"
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
               Left            =   1755
               TabIndex        =   77
               ToolTipText     =   "Click for Different C/S"
               Top             =   1260
               Width           =   4470
            End
            Begin VB.TextBox txtSecTap 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3780
               Locked          =   -1  'True
               TabIndex        =   72
               Top             =   270
               Width           =   2265
            End
            Begin VB.TextBox txtSecTurn 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1755
               Locked          =   -1  'True
               TabIndex        =   71
               Top             =   270
               Width           =   960
            End
            Begin VB.TextBox txtHt 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   1755
               TabIndex        =   73
               Text            =   "15"
               Top             =   585
               Width           =   555
            End
            Begin VB.TextBox txtNo 
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
               Left            =   2340
               MaxLength       =   1
               TabIndex        =   76
               Top             =   900
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtSwg 
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
               Left            =   1755
               TabIndex        =   75
               Top             =   900
               Visible         =   0   'False
               Width           =   555
            End
            Begin MSFlexGridLib.MSFlexGrid SWG 
               Height          =   1860
               Left            =   1755
               TabIndex        =   78
               Top             =   1575
               Visible         =   0   'False
               Width           =   3750
               _ExtentX        =   6615
               _ExtentY        =   3281
               _Version        =   393216
               BorderStyle     =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "% Compensation :"
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
               Left            =   90
               TabIndex        =   240
               Top             =   3465
               Width           =   1635
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               Caption         =   "Tap at :"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3060
               TabIndex        =   219
               Top             =   225
               Width           =   690
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Sec. Turns :"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   -225
               TabIndex        =   218
               Top             =   225
               Width           =   1950
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Height. (mm) :"
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
               Left            =   -225
               TabIndex        =   217
               Top             =   540
               Width           =   1950
            End
            Begin VB.Label Label67 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Conductor :"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   90
               TabIndex        =   216
               Top             =   855
               Width           =   1635
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "Core Material :"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2340
               TabIndex        =   215
               Top             =   540
               Width           =   1410
            End
         End
         Begin VB.TextBox cmbSelctCore 
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
            Left            =   1575
            MaxLength       =   1
            TabIndex        =   70
            Top             =   405
            Width           =   600
         End
         Begin VB.Frame Frame8 
            Caption         =   "Specifications :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   270
            TabIndex        =   194
            Top             =   720
            Width           =   6540
            Begin VB.Frame Frame20 
               Caption         =   "Special Protection :"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   135
               TabIndex        =   203
               Top             =   675
               Visible         =   0   'False
               Width           =   6270
               Begin VB.TextBox txtV1 
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
                  Height          =   285
                  Left            =   5670
                  TabIndex        =   221
                  Top             =   270
                  Width           =   510
               End
               Begin VB.TextBox txtIE1 
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
                  Height          =   285
                  Left            =   4455
                  TabIndex        =   206
                  Top             =   270
                  Width           =   645
               End
               Begin VB.TextBox txtRCT1 
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
                  Height          =   285
                  Left            =   2745
                  TabIndex        =   205
                  Top             =   270
                  Width           =   690
               End
               Begin VB.TextBox txtVK1 
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
                  Height          =   285
                  Left            =   630
                  TabIndex        =   204
                  Top             =   270
                  Width           =   600
               End
               Begin VB.Label Label74 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ie<=           mA"
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
                  Left            =   4005
                  TabIndex        =   210
                  Top             =   270
                  Width           =   1545
               End
               Begin VB.Label Label73 
                  BackStyle       =   0  'Transparent
                  Caption         =   "RCT(75 C)<=  "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1485
                  TabIndex        =   209
                  Top             =   270
                  Width           =   2535
                  WordWrap        =   -1  'True
               End
               Begin VB.Label Label72 
                  BackStyle       =   0  'Transparent
                  Caption         =   "VK>= "
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
                  Left            =   45
                  TabIndex        =   208
                  Top             =   270
                  Width           =   1545
               End
               Begin VB.Label Label71 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "@"
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
                  Left            =   5400
                  TabIndex        =   207
                  Top             =   270
                  Width           =   240
               End
            End
            Begin VB.TextBox txtCoreRatio 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1305
               Locked          =   -1  'True
               TabIndex        =   195
               Top             =   360
               Width           =   3840
            End
            Begin VB.Frame Frame19 
               Caption         =   "Frame6"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   135
               TabIndex        =   196
               Top             =   675
               Visible         =   0   'False
               Width           =   4515
               Begin VB.TextBox txtClass1 
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
                  Height          =   285
                  Left            =   2070
                  TabIndex        =   199
                  Top             =   270
                  Width           =   735
               End
               Begin VB.TextBox txtISF1 
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
                  Height          =   285
                  Left            =   3690
                  TabIndex        =   198
                  Top             =   270
                  Width           =   735
               End
               Begin VB.TextBox txtVA1 
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
                  Height          =   285
                  Left            =   480
                  TabIndex        =   197
                  Top             =   270
                  Width           =   645
               End
               Begin VB.Label Label70 
                  Caption         =   "ISF <="
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
                  Left            =   3015
                  TabIndex        =   202
                  Top             =   270
                  Width           =   825
               End
               Begin VB.Label Label69 
                  Caption         =   "VA "
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
                  Left            =   90
                  TabIndex        =   201
                  Top             =   270
                  Width           =   420
               End
               Begin VB.Label Label68 
                  Caption         =   "Class "
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
                  Left            =   1395
                  TabIndex        =   200
                  Top             =   270
                  Width           =   645
               End
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               Caption         =   "Core Ratio :"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   180
               TabIndex        =   211
               Top             =   315
               Width           =   1095
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Core Details :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2940
            Left            =   6975
            TabIndex        =   169
            Top             =   720
            Width           =   4290
            Begin VB.TextBox txtMean 
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
               Height          =   285
               Left            =   3285
               TabIndex        =   183
               Top             =   2115
               Width           =   780
            End
            Begin VB.TextBox txtEArea 
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
               Height          =   285
               Left            =   3285
               TabIndex        =   182
               Top             =   1755
               Width           =   780
            End
            Begin VB.TextBox txtArea 
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
               Height          =   285
               Left            =   3285
               TabIndex        =   181
               Top             =   1395
               Width           =   780
            End
            Begin VB.TextBox txtBUP 
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
               Height          =   285
               Left            =   3285
               TabIndex        =   180
               Top             =   1035
               Width           =   780
            End
            Begin VB.TextBox txtWt 
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
               Height          =   285
               Left            =   3285
               TabIndex        =   179
               Top             =   2475
               Width           =   780
            End
            Begin VB.Frame Frame12 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   540
               TabIndex        =   170
               Top             =   270
               Width           =   3615
               Begin VB.TextBox txtWW1 
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
                  Height          =   285
                  Left            =   2745
                  TabIndex        =   174
                  Top             =   405
                  Width           =   780
               End
               Begin VB.TextBox txtWL1 
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
                  Height          =   285
                  Left            =   1755
                  TabIndex        =   173
                  Top             =   405
                  Width           =   780
               End
               Begin VB.TextBox txtOW1 
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
                  Height          =   285
                  Left            =   2745
                  TabIndex        =   172
                  Top             =   45
                  Width           =   780
               End
               Begin VB.TextBox txtOL1 
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
                  Height          =   285
                  Left            =   1755
                  TabIndex        =   171
                  Top             =   45
                  Width           =   780
               End
               Begin VB.Label Label57 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   2565
                  TabIndex        =   178
                  Top             =   450
                  Width           =   150
               End
               Begin VB.Label Label56 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "WD (mm X mm) :"
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
                  Left            =   135
                  TabIndex        =   177
                  Top             =   405
                  Width           =   1545
               End
               Begin VB.Label Label55 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   2565
                  TabIndex        =   176
                  Top             =   45
                  Width           =   150
               End
               Begin VB.Label Label54 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "OD (mm X mm) :"
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
                  Left            =   135
                  TabIndex        =   175
                  Top             =   0
                  Width           =   1515
               End
            End
            Begin VB.Frame Frame18 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   2205
               TabIndex        =   184
               Top             =   270
               Visible         =   0   'False
               Width           =   2040
               Begin VB.TextBox txtID1 
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
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   186
                  Top             =   405
                  Width           =   780
               End
               Begin VB.TextBox txtOD1 
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
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   185
                  Top             =   45
                  Width           =   780
               End
               Begin VB.Label Label61 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "OD (mm) :"
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
                  Left            =   75
                  TabIndex        =   188
                  Top             =   45
                  Width           =   975
               End
               Begin VB.Label Label62 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ID (mm) :"
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
                  Left            =   135
                  TabIndex        =   187
                  Top             =   405
                  Width           =   915
               End
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "MEAN MAG PATH. LENGTH (Cms) :"
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
               Left            =   90
               TabIndex        =   193
               Top             =   2070
               Width           =   3165
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "EFFECTIVE AREA Aeff (Cm Sq) :"
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
               Left            =   270
               TabIndex        =   192
               Top             =   1710
               Width           =   2985
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "AREA A. (Cm Sq) :"
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
               Left            =   1170
               TabIndex        =   191
               Top             =   1395
               Width           =   2085
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "BUILD UP T (mm) :"
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
               Left            =   1440
               TabIndex        =   190
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "WEIGHT (Kg) :"
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
               Left            =   1890
               TabIndex        =   189
               Top             =   2430
               Width           =   1365
            End
         End
         Begin MSFlexGridLib.MSFlexGrid G1 
            Height          =   5865
            Left            =   -74865
            TabIndex        =   168
            Top             =   495
            Width           =   11265
            _ExtentX        =   19870
            _ExtentY        =   10345
            _Version        =   393216
            Rows            =   57
            AllowUserResizing=   3
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
         Begin VB.TextBox txtCoreType 
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
            Left            =   6885
            TabIndex        =   167
            Top             =   5850
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.TextBox txtTWh 
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
            Left            =   6885
            TabIndex        =   166
            Text            =   " "
            Top             =   5535
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label88 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Press to View Core Results..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   7650
            TabIndex        =   256
            Top             =   4455
            Width           =   3210
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderWidth     =   2
            Height          =   1275
            Left            =   7650
            Top             =   4275
            Width           =   3255
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Select Core :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   212
            Top             =   360
            Width           =   1410
         End
      End
      Begin VB.ComboBox cmbEquipType 
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
         ItemData        =   "CTfrm.frx":044E
         Left            =   -73560
         List            =   "CTfrm.frx":0450
         TabIndex        =   0
         Top             =   855
         Width           =   2400
      End
      Begin VB.TextBox txtnoofcores 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69885
         MaxLength       =   1
         TabIndex        =   1
         Top             =   900
         Width           =   555
      End
      Begin VB.ComboBox cmbHSV 
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
         Left            =   -74010
         TabIndex        =   5
         Top             =   1710
         Width           =   1140
      End
      Begin VB.ComboBox cmbIL 
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
         Left            =   -72075
         TabIndex        =   6
         Top             =   1710
         Width           =   2760
      End
      Begin VB.ComboBox cmbSTC 
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
         Left            =   -72075
         TabIndex        =   3
         Top             =   1305
         Width           =   1185
      End
      Begin VB.ComboBox cmbTime 
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
         ItemData        =   "CTfrm.frx":0452
         Left            =   -69960
         List            =   "CTfrm.frx":0462
         TabIndex        =   4
         Top             =   1305
         Width           =   645
      End
      Begin VB.ComboBox cmbFreq 
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
         ItemData        =   "CTfrm.frx":0474
         Left            =   -74010
         List            =   "CTfrm.frx":047E
         TabIndex        =   7
         Top             =   2115
         Width           =   825
      End
      Begin VB.TextBox txtRatio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -67935
         TabIndex        =   37
         Top             =   900
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtRatioR2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -64965
         TabIndex        =   38
         Top             =   900
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtRatioR2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -64965
         TabIndex        =   40
         Top             =   1260
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtRatio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -67935
         TabIndex        =   39
         Top             =   1260
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtRatioR2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -64965
         TabIndex        =   41
         Top             =   1620
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtRatioR2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -64965
         TabIndex        =   43
         Top             =   1980
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtRatio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -67935
         TabIndex        =   42
         Top             =   1980
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtRatioR2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   -64965
         TabIndex        =   45
         Top             =   2340
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtRatio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   -67935
         TabIndex        =   44
         Top             =   2340
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtRateConti 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70500
         MaxLength       =   5
         TabIndex        =   9
         Top             =   2115
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtIDATurns 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         TabIndex        =   69
         Top             =   6480
         Width           =   1095
      End
      Begin VB.TextBox txtIDyn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         TabIndex        =   68
         Top             =   6030
         Width           =   1095
      End
      Begin VB.TextBox txtminCs 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         TabIndex        =   59
         Top             =   1890
         Width           =   1095
      End
      Begin VB.TextBox txtWidth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   4635
         Width           =   1095
      End
      Begin VB.TextBox txtCondPer 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   3255
         Width           =   3345
      End
      Begin VB.TextBox txtTotalCon 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   2805
         Width           =   3345
      End
      Begin VB.TextBox txtSide 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtTop 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   4185
         Width           =   1095
      End
      Begin VB.TextBox txtDelStc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   5580
         Width           =   1095
      End
      Begin VB.TextBox txtPriCros 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71085
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Frame Frame17 
         Caption         =   "Core Sizes :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74235
         TabIndex        =   139
         Top             =   3015
         Width           =   5640
         Begin MSFlexGridLib.MSFlexGrid GCore 
            Height          =   3615
            Left            =   90
            TabIndex        =   140
            Top             =   270
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   6376
            _Version        =   393216
            ForeColor       =   0
            ForeColorFixed  =   192
            SelectionMode   =   1
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
      Begin VB.Frame Frame11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74235
         TabIndex        =   130
         Top             =   1845
         Visible         =   0   'False
         Width           =   2445
         Begin VB.TextBox txtID 
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
            Left            =   1215
            TabIndex        =   51
            Top             =   675
            Width           =   1005
         End
         Begin VB.TextBox txtOD 
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
            Left            =   1215
            TabIndex        =   50
            Top             =   315
            Width           =   1005
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OD (mm) :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   132
            Top             =   315
            Width           =   975
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID (mm) :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   131
            Top             =   675
            Width           =   915
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Standard Dimensions for Rectangluar Core :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6450
         Left            =   -68475
         TabIndex        =   129
         Top             =   540
         Width           =   5145
         Begin MSFlexGridLib.MSFlexGrid Grid3 
            Height          =   6090
            Left            =   90
            TabIndex        =   58
            Top             =   270
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   10742
            _Version        =   393216
            ForeColor       =   8388608
            ForeColorFixed  =   192
            SelectionMode   =   1
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
      Begin VB.ComboBox cmbStyle 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "CTfrm.frx":048A
         Left            =   -71355
         List            =   "CTfrm.frx":04CD
         TabIndex        =   46
         Top             =   675
         Width           =   2715
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   128
         Top             =   1080
         Width           =   6270
         Begin VB.ComboBox cmbCore 
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
            Height          =   315
            Left            =   4770
            TabIndex        =   49
            Top             =   270
            Width           =   1365
         End
         Begin VB.OptionButton OptCom 
            Caption         =   "Common Core Size"
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
            Index           =   0
            Left            =   90
            TabIndex        =   47
            Top             =   270
            Width           =   2220
         End
         Begin VB.OptionButton OptCom 
            Caption         =   "Different Core Size"
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
            Index           =   1
            Left            =   2430
            TabIndex        =   48
            Top             =   270
            Width           =   2220
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Core Details :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   -74865
         TabIndex        =   112
         Top             =   2745
         Width           =   11535
         Begin VB.TextBox txtinv 
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
            Left            =   2565
            TabIndex        =   21
            Top             =   315
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Frame Frame4 
            Caption         =   "Accuracy :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   4365
            TabIndex        =   117
            Top             =   180
            Width           =   7080
            Begin VB.ComboBox IvAcc 
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
               Height          =   315
               Left            =   5265
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   225
               Width           =   1725
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               Caption         =   "Ratio : "
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
               Left            =   4365
               TabIndex        =   127
               Top             =   225
               Width           =   780
            End
            Begin MSForms.OptionButton OptInd 
               Height          =   330
               Left            =   2160
               TabIndex        =   25
               Top             =   225
               Width           =   1950
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               DisplayStyle    =   5
               Size            =   "3440;582"
               Value           =   "0"
               Caption         =   "Different Accuracy"
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.OptionButton OptComm 
               Height          =   330
               Left            =   135
               TabIndex        =   24
               Top             =   225
               Width           =   1950
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               DisplayStyle    =   5
               Size            =   "3440;582"
               Value           =   "0"
               Caption         =   "Common Accuracy"
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.ComboBox cmbnoofcores 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   315
            Width           =   1140
         End
         Begin VB.TextBox txtRatio2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1215
            Width           =   2940
         End
         Begin VB.ComboBox cmbCType 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "CTfrm.frx":050E
            Left            =   1305
            List            =   "CTfrm.frx":051B
            TabIndex        =   22
            Top             =   765
            Width           =   2940
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frame6"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4365
            TabIndex        =   118
            Top             =   855
            Visible         =   0   'False
            Width           =   6900
            Begin VB.ComboBox cmbClass 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "CTfrm.frx":0549
               Left            =   3172
               List            =   "CTfrm.frx":054B
               TabIndex        =   28
               Top             =   270
               Width           =   1275
            End
            Begin VB.ComboBox cmdVA 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "CTfrm.frx":054D
               Left            =   855
               List            =   "CTfrm.frx":0578
               TabIndex        =   27
               Top             =   270
               Width           =   1320
            End
            Begin VB.ComboBox cmbISF 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "CTfrm.frx":05B1
               Left            =   5445
               List            =   "CTfrm.frx":05B3
               TabIndex        =   29
               Top             =   270
               Width           =   1275
            End
            Begin VB.Label Label13 
               Caption         =   "Class : "
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
               Left            =   2520
               TabIndex        =   121
               Top             =   270
               Width           =   645
            End
            Begin VB.Label label100 
               Caption         =   "VA : "
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
               Left            =   450
               TabIndex        =   120
               Top             =   270
               Width           =   420
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "ISF :"
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
               Left            =   4590
               TabIndex        =   119
               Top             =   270
               Width           =   825
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Special Protection :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   4365
            TabIndex        =   122
            Top             =   855
            Width           =   7080
            Begin VB.ComboBox txtIe2 
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
               ItemData        =   "CTfrm.frx":05B5
               Left            =   6210
               List            =   "CTfrm.frx":05C2
               TabIndex        =   33
               Top             =   270
               Width           =   780
            End
            Begin VB.TextBox txtVK 
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
               TabIndex        =   30
               Top             =   270
               Width           =   600
            End
            Begin VB.TextBox txtRCT 
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
               Left            =   2970
               TabIndex        =   31
               Top             =   270
               Width           =   600
            End
            Begin VB.TextBox txtIe 
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
               Left            =   4860
               TabIndex        =   32
               Top             =   315
               Width           =   600
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "@"
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
               Left            =   5895
               TabIndex        =   126
               Top             =   270
               Width           =   240
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "VK>= "
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
               Left            =   90
               TabIndex        =   125
               Top             =   270
               Width           =   1725
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "  RCT(75 C)<=   "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1620
               TabIndex        =   124
               Top             =   270
               Width           =   2670
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ie<=           mA"
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
               Left            =   3915
               TabIndex        =   123
               Top             =   270
               Width           =   1860
            End
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Select Core : "
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
            Left            =   90
            TabIndex        =   115
            Top             =   270
            Width           =   1230
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ratio : "
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
            Left            =   630
            TabIndex        =   114
            Top             =   1215
            Width           =   690
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Core Type : "
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
            Left            =   90
            TabIndex        =   113
            Top             =   720
            Width           =   1230
         End
      End
      Begin VB.TextBox txtNominal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67170
         TabIndex        =   110
         Top             =   3465
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.TextBox txtAmper 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67170
         TabIndex        =   109
         Top             =   3825
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Frame Frame5 
         Caption         =   "Core Details :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   -74865
         TabIndex        =   108
         Top             =   4455
         Width           =   10815
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   2265
            Left            =   90
            TabIndex        =   36
            Top             =   225
            Width           =   10635
            _ExtentX        =   18759
            _ExtentY        =   3995
            _Version        =   393216
            ForeColorFixed  =   192
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
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74235
         TabIndex        =   133
         Top             =   1845
         Width           =   5640
         Begin VB.TextBox txtOL 
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
            Left            =   3285
            TabIndex        =   52
            Top             =   315
            Width           =   825
         End
         Begin VB.TextBox txtOW 
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
            Left            =   4590
            TabIndex        =   53
            Top             =   315
            Width           =   870
         End
         Begin VB.TextBox txtWL 
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
            Left            =   3285
            TabIndex        =   54
            Top             =   675
            Width           =   825
         End
         Begin VB.TextBox txtWW 
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
            Left            =   4590
            TabIndex        =   55
            Top             =   675
            Width           =   870
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Over All Dimension (mm X mm) :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   225
            TabIndex        =   137
            Top             =   270
            Width           =   3015
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4275
            TabIndex        =   136
            Top             =   360
            Width           =   150
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Window Dimension (mm X mm) :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   135
            Top             =   630
            Width           =   3075
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4275
            TabIndex        =   134
            Top             =   720
            Width           =   150
         End
      End
      Begin VB.Frame Frame15 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   135
         TabIndex        =   228
         Top             =   2925
         Width           =   6450
         Begin MSFlexGridLib.MSFlexGrid Former 
            Height          =   1095
            Left            =   45
            TabIndex        =   105
            Top             =   225
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   1931
            _Version        =   393216
            ForeColor       =   0
            ForeColorFixed  =   192
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
      Begin VB.TextBox txtRatio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -67935
         TabIndex        =   255
         Top             =   1620
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtDesNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73470
         TabIndex        =   254
         Top             =   6570
         Visible         =   0   'False
         Width           =   2940
      End
      Begin Crystal.CrystalReport CrystalReport2 
         Left            =   7875
         Top             =   6570
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "\\Server\c\Reports\Design\Costing--F.rpt"
         WindowTitle     =   "Costing"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         Connect         =   "odbc-enhance"
         UserName        =   "sa"
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Frame Frame21 
         Caption         =   "Primary Winding :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   6975
         TabIndex        =   243
         Top             =   495
         Visible         =   0   'False
         Width           =   4290
         Begin VB.TextBox txtpricuWt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   1440
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox txtPriTotLength 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   1080
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox txtPriLength 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   95
            Top             =   720
            Width           =   960
         End
         Begin VB.TextBox txtPriMeanL 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   94
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label80 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRI. CU. WEIGHT (kg) :"
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
            Left            =   405
            TabIndex        =   247
            Top             =   1440
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label Label79 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL LENGTH (mm) :"
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
            Left            =   630
            TabIndex        =   246
            Top             =   1080
            Visible         =   0   'False
            Width           =   2490
         End
         Begin VB.Label Label78 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "LENGTH (mm) :"
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
            Left            =   180
            TabIndex        =   245
            Top             =   720
            Width           =   2940
         End
         Begin VB.Label Label77 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRI. WDG. MEAN LENGTH (mm) :"
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
            Left            =   0
            TabIndex        =   244
            Top             =   360
            Width           =   3120
         End
      End
      Begin MSForms.ComboBox cmbConStyle 
         Height          =   330
         Left            =   -71085
         TabIndex        =   268
         Top             =   945
         Width           =   3990
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "7038;582"
         ColumnCount     =   4
         cColumnInfo     =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;0;0;7055"
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   465
         Left            =   -64035
         TabIndex        =   267
         ToolTipText     =   "Cancel Design"
         Top             =   6150
         Width           =   765
         Caption         =   "Search"
         Size            =   "1349;820"
         Accelerator     =   99
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label92 
         Height          =   285
         Left            =   6795
         TabIndex        =   266
         Top             =   5760
         Width           =   4605
         ForeColor       =   8388608
         BackColor       =   16761024
         Caption         =   "Go Back To Stage"
         Size            =   "8123;503"
         BorderColor     =   16711680
         SpecialEffect   =   6
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SECONDARY INSULATION :"
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
         Left            =   6840
         TabIndex        =   264
         Top             =   5400
         Width           =   2715
      End
      Begin VB.Label Label85 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PRI. RESISTANCE 20 Deg C (m  "
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
         Left            =   6930
         TabIndex        =   263
         Top             =   2450
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label Label86 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "WATT LOSS :"
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
         Left            =   7335
         TabIndex        =   262
         Top             =   3150
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label Label87 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PRI. RESISTANCE 75 Deg C (m"
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
         Left            =   6840
         TabIndex        =   261
         Top             =   2760
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label Label90 
         Caption         =   "W) : "
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9675
         TabIndex        =   260
         Top             =   2430
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label91 
         Caption         =   "W) : "
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9720
         TabIndex        =   259
         Top             =   2745
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSForms.CommandButton cmdCancel 
         Height          =   465
         Left            =   -64035
         TabIndex        =   257
         ToolTipText     =   "Cancel Design"
         Top             =   6645
         Width           =   780
         Caption         =   "Can."
         Size            =   "1376;820"
         Accelerator     =   99
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label76 
         Height          =   285
         Left            =   -73110
         TabIndex        =   242
         Top             =   2160
         Width           =   330
         Caption         =   "Hz"
         Size            =   "582;503"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label75 
         Caption         =   "Specification :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74730
         TabIndex        =   241
         Top             =   450
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         Height          =   2130
         Left            =   -74910
         Top             =   585
         Width           =   11490
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rated Cont. Pri. Current : "
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
         Left            =   -72750
         TabIndex        =   239
         Top             =   2115
         Width           =   2220
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Freq : "
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
         Left            =   -74775
         TabIndex        =   238
         Top             =   2115
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.L. : "
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
         Left            =   -72570
         TabIndex        =   237
         Top             =   1710
         Width           =   510
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.O. Cores : "
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
         Left            =   -71490
         TabIndex        =   236
         Top             =   900
         Width           =   1545
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         Caption         =   "H.S.V. : "
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
         Left            =   -74775
         TabIndex        =   235
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Secondary Wonded Dimensions :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   230
         Top             =   405
         Width           =   2850
      End
      Begin VB.Label WindowSize 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   229
         Top             =   2700
         Width           =   5595
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MODEL NO / DRG NO :                              TYPE :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   226
         Top             =   2430
         Width           =   4335
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Selection"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -67320
         TabIndex        =   223
         Top             =   495
         Width           =   3885
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   510
         Left            =   -66975
         Shape           =   4  'Rounded Rectangle
         Top             =   450
         Width           =   3135
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Core Selection"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -74775
         TabIndex        =   222
         Top             =   630
         Width           =   3300
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   510
         Left            =   -74865
         Shape           =   4  'Rounded Rectangle
         Top             =   540
         Width           =   3435
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Current Desity (A/Sq mm) :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74730
         TabIndex        =   220
         Top             =   5130
         Width           =   3570
      End
      Begin VB.Label Label1 
         Caption         =   "Equip. Type : "
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
         Left            =   -74730
         TabIndex        =   164
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "N.S.V. :"
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
         Left            =   -74775
         TabIndex        =   163
         Top             =   1305
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "S.T.C. : "
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
         Left            =   -72795
         TabIndex        =   162
         Top             =   1305
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Time(S) : "
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
         Left            =   -70905
         TabIndex        =   161
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Core 1 Ratio :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   -69285
         TabIndex        =   160
         Top             =   900
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Core 1 Ratio :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   -69285
         TabIndex        =   159
         Top             =   1260
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Core 1 Ratio :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   -69285
         TabIndex        =   158
         Top             =   1620
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Core 1 Ratio :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   -69285
         TabIndex        =   157
         Top             =   1980
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Core 1 Ratio :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   -69285
         TabIndex        =   156
         Top             =   2340
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label14 
         Caption         =   "Primary Current (A)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -67935
         TabIndex        =   155
         Top             =   675
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label Label17 
         Caption         =   "Sec. Current (A)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -65190
         TabIndex        =   154
         Top             =   675
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I Dynamic Ampere turns :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74730
         TabIndex        =   153
         Top             =   6480
         Width           =   3570
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I Dynamic Ampere (kA) :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74730
         TabIndex        =   152
         Top             =   6030
         Width           =   3570
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Required C/S (Sq.mm) :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74730
         TabIndex        =   151
         Top             =   1890
         Width           =   3570
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width (mm) :"
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
         Left            =   -72315
         TabIndex        =   150
         Top             =   4635
         Width           =   1155
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top Winding Ht.(mm) :"
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
         Left            =   -73125
         TabIndex        =   149
         Top             =   4170
         Width           =   1965
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Side/Inside Wdg Ht. (mm) :"
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
         Left            =   -74220
         TabIndex        =   148
         Top             =   3720
         Width           =   3060
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Caption         =   "Cross Section"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -67215
         TabIndex        =   147
         Top             =   1410
         Width           =   2940
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Conductors :"
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
         Left            =   -74730
         TabIndex        =   146
         Top             =   2805
         Width           =   3570
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conductor Per Coil :"
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
         Left            =   -72915
         TabIndex        =   145
         Top             =   3255
         Width           =   1755
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "STC Current Desity (1sec) (A/Sq mm) :"
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
         Left            =   -74730
         TabIndex        =   144
         Top             =   5580
         Width           =   3570
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Turns (Nos) :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74730
         TabIndex        =   143
         Top             =   1440
         Width           =   3570
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Primary C/S (Sq. mm) :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74730
         TabIndex        =   142
         Top             =   2340
         Width           =   3570
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Conductor Style :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74730
         TabIndex        =   141
         Top             =   945
         Width           =   3570
      End
      Begin MSForms.CommandButton cmdICore 
         Height          =   465
         Left            =   -74910
         TabIndex        =   56
         ToolTipText     =   "Press To pull core size"
         Top             =   3105
         Width           =   600
         Caption         =   ">>"
         Size            =   "1058;820"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRemoCore 
         Height          =   465
         Left            =   -74910
         TabIndex        =   57
         ToolTipText     =   "Delete Core size"
         Top             =   6525
         Width           =   600
         Caption         =   "<<"
         Size            =   "1058;820"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label30 
         Caption         =   "Core Shape : "
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
         Left            =   -71355
         TabIndex        =   138
         Top             =   450
         Width           =   1500
      End
      Begin MSForms.CommandButton cmdRemCore 
         Height          =   465
         Left            =   -64035
         TabIndex        =   35
         ToolTipText     =   "Delete Core specification"
         Top             =   5655
         Width           =   780
         Caption         =   ">> Del"
         Size            =   "1376;820"
         Accelerator     =   79
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdIN 
         Height          =   465
         Left            =   -64035
         TabIndex        =   34
         ToolTipText     =   "Press to pull core specificaton"
         Top             =   4545
         Width           =   780
         Caption         =   "<< In"
         Size            =   "1376;820"
         Accelerator     =   73
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Shape Shape1 
         Height          =   5115
         Left            =   -67350
         Top             =   1515
         Width           =   3210
      End
      Begin VB.Shape Shape6 
         Height          =   6270
         Left            =   -74775
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   11175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid TestGrid 
      Height          =   825
      Left            =   270
      TabIndex        =   111
      Top             =   6390
      Visible         =   0   'False
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   1455
      _Version        =   393216
      SelectionMode   =   1
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
   Begin MSFlexGridLib.MSFlexGrid G2 
      Height          =   5865
      Left            =   135
      TabIndex        =   248
      Top             =   1260
      Visible         =   0   'False
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   10345
      _Version        =   393216
      Rows            =   57
      AllowUserResizing=   3
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
   Begin MSFlexGridLib.MSFlexGrid GCross 
      Height          =   1725
      Left            =   2745
      TabIndex        =   258
      Top             =   5490
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3043
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
Attribute VB_Name = "CTfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As Recordset
Dim ListIndexofCondutor As Integer
Dim Flag1 As Boolean
Dim NoofSec As String
Dim colu As Integer
Dim STCMemory As String
Dim STCTimeMemory As String
Dim MODTEST As Integer
'compensation WOR1.Vis WindowSize
Dim compensation As Double

'MiniCrossSec
Dim MinCross As Double

'Former setting
Dim LastSet As String * 1

Dim OALLWin, OALLLen As Double
Dim OALLWin1, OALLLen1 As Double
Dim NoofPrim As Integer
Dim NoofSeco As Integer

Dim coiltype As String * 1
Dim CoilName As String

Dim PrimPoin As Integer
Dim SecPoin As Integer

Dim sinD, cosD As Double 'ErrorRead

Dim PriTurns As String
Dim GNo As Double
Dim BMatQuatar As Double
Dim BMatHalf As Double
Dim LowDen As Double

''Primary Rated continuse current
Dim PriRateContinus  As String

Private Sub SearchForStyle()

If Len(txtTurns) > 0 Then
    If CDbl(txtTurns) > 1 Then
        If cmbConStyle.Text = "Double Coil - Bare Conductor" Then
            MsgBox "This coil type is not applicable for 1 turns", vbCritical
            cmbConStyle.ListIndex = -1
            cmbConStyle.SetFocus
            txtPriCros = ""
            txtTop = ""
            txtSide = ""
            txtCondPer = ""
            txtTotalCon = ""
            lstCond.Clear
            Exit Sub
        ElseIf cmbConStyle.Text = "Single Coil - Bare Conductor" Then
            MsgBox "This coil type is not applicable for 1 turns", vbCritical
            cmbConStyle.ListIndex = -1
            cmbConStyle.SetFocus
            txtPriCros = ""
            txtTop = ""
            txtSide = ""
            txtCondPer = ""
            txtTotalCon = ""
            lstCond.Clear
            Exit Sub
        Else
            PrcCalculate
        End If
    ElseIf CDbl(txtTurns) = 1 Then
        If cmbConStyle.Text = "Double Coil - Normal Wdg Ht." Then
            MsgBox "This coil type is not applicable for more than 1 turns", vbCritical
            cmbConStyle.ListIndex = -1
            cmbConStyle.SetFocus
            txtPriCros = ""
            txtTop = ""
            txtSide = ""
            txtCondPer = ""
            txtTotalCon = ""
            lstCond.Clear
            Exit Sub
        ElseIf cmbConStyle.Text = "Double Coil - Equall Wdg Ht." Then
            MsgBox "This coil type is not applicable for more than 1 turns", vbCritical
            cmbConStyle.ListIndex = -1
            cmbConStyle.SetFocus
            txtPriCros = ""
            txtTop = ""
            txtSide = ""
            txtCondPer = ""
            txtTotalCon = ""
            lstCond.Clear
            Exit Sub
        ElseIf cmbConStyle.Text = "Single Coil" Then
            MsgBox "This coil type is not applicable for more than 1 turns", vbCritical
            cmbConStyle.ListIndex = -1
            cmbConStyle.SetFocus
            txtPriCros = ""
            txtTop = ""
            txtSide = ""
            txtCondPer = ""
            txtTotalCon = ""
            lstCond.Clear
            Exit Sub
        ElseIf cmbConStyle.Text = "Series Coil" Then
            MsgBox "This coil type is not applicable for more than 1 turns", vbCritical
            cmbConStyle.ListIndex = -1
            cmbConStyle.SetFocus
            txtPriCros = ""
            txtTop = ""
            txtSide = ""
            txtCondPer = ""
            txtTotalCon = ""
            lstCond.Clear
            Exit Sub
'        Else
'            PrcCalculate
        End If

    End If
    
End If

End Sub

Private Sub ChK2_Click()
Dim value As String * 1
Dim str1 As String

Dim RateCs As Double

If ChK2.value = 1 Then
    txtSwg.Visible = False
    txtNo.Visible = False
    Label67.Visible = True
    SWG.Visible = True
    DiaType = "D"
    For i = 1 To SWG.Rows - 1
        txtMinCsSec = prcCalSecC(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1), Val(SWG.TextMatrix(i, 0)), cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
        If Len(Trim(txtRateConti)) > 0 Then
            'if rate continuous  present compare with rate cont. with  stc
            RateCs = prcCalRate(txtRateConti, Val(SWG.TextMatrix(i, 0)), txtRatioR2(CInt(cmbSelctCore - 1)))
            CompValue = RateCs
        Else
            'else stc and rate current
            txtSecCS = prcCalNorSecC(txtRatioR2(CInt(cmbSelctCore - 1)))
            CompValue = RateCs
        End If
'        Dim CompValue As Double
        If CDbl(CompValue) >= CDbl(txtMinCsSec) Then
            Set rst = New ADODB.Recordset
            rst.Open "select * from secconductor where t_cs >= " & CompValue & " ", con, adOpenStatic, adLockReadOnly
            If rst.EOF = False And rst.BOF = False Then
                SWG.TextMatrix(i, 1) = rst!t_swg
                SWG.TextMatrix(i, 2) = "S"
            End If
            Set rst = Nothing
        Else
            Set rst = New ADODB.Recordset
            rst.Open "select * from secconductor where t_cs >= " & txtMinCsSec & " ", con, adOpenStatic, adLockReadOnly
            If rst.EOF = False And rst.BOF = False Then
                SWG.TextMatrix(i, 1) = rst!t_swg
                SWG.TextMatrix(i, 2) = "S"
            End If
            Set rst = Nothing
        End If
    Next
Else
    txtSwg.Visible = True
    txtNo.Visible = True
    DiaType = "S"
    SWG.Visible = False
End If

Exit Sub
















Chkoption_Click
If ChK2.value = 0 Then 'With same dia
    cmdIn.Tag = ""
    DiaType = "D"
'    If cmbStyle = "Rectangluar" Then
'        'For Assending of Rectangular with tap
'        If Grid.Tag = "A" Then
'            CalLayersTapS
'        End If
'    Else
'        'For Assending of Round with tap
'        If Grid.Tag = "A" Then
'            CalLayersTapSRound
'        End If
'    End If
Else 'With different dia
    cmdIn.Tag = ""
    DiaType = "S"
'    If cmbStyle = "Rectangluar" Then
'        'For Assending of Rectangular with tap
'        If Grid.Tag = "A" Then
'            CalLayersTapDiff
'        End If
'    Else
'        'For Assending of round with tap
'        If Grid.Tag = "A" Then
'            CalLayersTapDiffRound
'        End If
'    End If
End If

Exit Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
If GCross.Rows > 1 Then
    txtTWh = 1.5
    Dim TLay As Byte
    TLay = 0
    For K = 1 To GCross.Rows - 1
        txtTWh = txtTWh + (GCross.TextMatrix(K, 1) * GCross.TextMatrix(K, 6))
        TLay = TLay + GCross.TextMatrix(K, 6)
    Next
    TLay = TLay - 1
    txtTWh = txtTWh + (TLay * 0.25) + 1.5 + 3
End If

'Wounded Dimensions
Dim TWH As Double
'txtTWh = 12.553
'txtTWh = 11.86
txtTWh = Round(txtTWh, 2)
MsgBox Mid(txtTWh, InStr(1, txtTWh, ".") + 1)
If ChK2.value = 1 Then
    
    If InStr(1, txtTWh, ".") > 0 Then
        If Mid(txtTWh, InStr(1, txtTWh, ".") + 1) >= 0 And Mid(txtTWh, InStr(1, txtTWh, ".") + 1) < 25 Then
            TWH = Mid(txtTWh, 1, InStr(1, txtTWh, ".") - 1) & ".0"
        ElseIf Mid(txtTWh, InStr(1, txtTWh, ".") + 1) >= 25 And Mid(txtTWh, InStr(1, txtTWh, ".") + 1) < 50 Then
            TWH = Mid(txtTWh, 1, InStr(1, txtTWh, ".") - 1) & ".5"
        ElseIf Mid(txtTWh, InStr(1, txtTWh, ".") + 1) >= 50 And Mid(txtTWh, InStr(1, txtTWh, ".") + 1) < 75 Then
            TWH = Mid(txtTWh, 1, InStr(1, txtTWh, ".") - 1) & ".5"
        ElseIf Mid(txtTWh, InStr(1, txtTWh, ".") + 1) >= 75 Then
            TWH = Mid(txtTWh, 1, InStr(1, txtTWh, ".") - 1)
            TWH = TWH + 1
        End If
    End If
    
Else

    If InStr(1, txtWHt, ".") > 0 Then
        If Mid(txtWHt, InStr(1, txtWHt, ".") + 1) >= 0 And Mid(txtWHt, InStr(1, txtWHt, ".") + 1) < 25 Then
            TWH = Mid(txtWHt, 1, InStr(1, txtWHt, ".") - 1) & ".0"
        ElseIf Mid(txtWHt, InStr(1, txtWHt, ".") + 1) >= 25 And Mid(txtWHt, InStr(1, txtWHt, ".") + 1) < 50 Then
            TWH = Mid(txtWHt, 1, InStr(1, txtWHt, ".") - 1) & ".5"
        ElseIf Mid(txtWHt, InStr(1, txtWHt, ".") + 1) >= 50 And Mid(txtWHt, InStr(1, txtWHt, ".") + 1) < 75 Then
            TWH = Mid(txtWHt, 1, InStr(1, txtWHt, ".") - 1) & ".5"
        ElseIf Mid(txtWHt, InStr(1, txtWHt, ".") + 1) >= 75 Then
            TWH = Mid(txtWHt, 1, InStr(1, txtWHt, ".") - 1)
            TWH = TWH + 1
        End If
    End If
End If

If cmbStyle = "Rectangluar" Then
    txtOW2 = txtOW1 + 2 * (TWH)
    txtOL2 = txtOL1 + 2 * (TWH)
    
    txtWL2 = txtWL1 - 2 * (TWH)
    txtWW2 = txtWW1 - 2 * (TWH)
Else
    
    txtOD2 = txtOD1 + 2 * (TWH)
    txtID1 = txtID1 + 2 * (TWH)
    
End If

txtHt = txtHt + 2 * (TWH)

End Sub

Private Sub prcFillInd()

'On Error GoTo err

    Dim int1 As Double
    Dim value As String * 1
    Dim str1 As String
    Dim Gnumber As Double
    Dim K As Byte
Individualfrm.txtCoppWt = 0
Individualfrm.txtRatio = txtRatio(cmbSelctCore - 1) & "/" & txtRatioR2(cmbSelctCore - 1)
ARatio = txtRatio(cmbSelctCore - 1)
Individualfrm.cmbStyle = cmbStyle
'$$$$$$$$$Clear tab caption
For i = 0 To 4
    Individualfrm.SSTab2.TabCaption(i) = "---"
Next
If cmbCType.Text = "Special Protection" Then
    Individualfrm.Frame20.Visible = True
    Individualfrm.Frame19.Visible = False
Else
    Individualfrm.Frame20.Visible = False
    Individualfrm.Frame19.Visible = True
End If
If cmbStyle = "Rectangluar" Then
    Individualfrm.Frame12.Visible = True
    Individualfrm.Frame18.Visible = False
    Individualfrm.txtWL1 = txtWL1
    Individualfrm.txtWW1 = txtWW1
    Individualfrm.txtOL1 = txtOL1
    Individualfrm.txtOW1 = txtOW1
Else
    Individualfrm.Frame12.Visible = False
    Individualfrm.Frame18.Visible = True
    Individualfrm.txtOD1 = txtOD1
    Individualfrm.txtID1 = txtID1
End If

Individualfrm.txtBUP = txtBUP
Individualfrm.txtEArea = txtEArea
Individualfrm.txtMean = txtMean
Individualfrm.txtWt = txtWt
Individualfrm.txtArea = txtArea
Individualfrm.cmbStyle = cmbStyle
Individualfrm.cmbMetal = cmbMetal
Individualfrm.txtCoreType = txtCoreType

If txtCoreType = "Metering" Then
    Individualfrm.Label70.Caption = "ISF<=:"
ElseIf txtCoreType = "Protection" Then
    Individualfrm.Label70.Caption = "ALF:"
End If

Individualfrm.cmbFreq = cmbFreq

Individualfrm.txtVA1 = txtVA1

Individualfrm.txtClass1 = txtClass1
If Len(txtISF1) > 0 Then
    Individualfrm.txtISF1 = txtISF1
Else
    Individualfrm.txtISF1 = "-"
End If
Individualfrm.cmbSTC = cmbSTC
Individualfrm.cmbTime = cmbTime
Individualfrm.txtRCT1 = txtRCT1
Individualfrm.txtVK1(0) = txtVK1
Individualfrm.txtIE1 = txtIE1
Individualfrm.txtV1 = txtV1
'Individualfrm.cmbStyle = cmbStyle
Individualfrm.txtTurns = txtTurns

'$$$$$$$$Putting primary ratios both primary and secondary
If InStr(1, txtCoreRatio, "-") > 0 Then
    '#####Primary
'    ARatio = Individualfrm.txtRatio(cmbSelctCore - 1)



    For i = 1 To Len(txtRatio(cmbSelctCore - 1)) + 1
        value = Mid(txtRatio(cmbSelctCore - 1), i, i)
        If value = "-" Or i > Len(ARatio) Then
             If CDbl(str1) > Gnumber Then
                  Gnumber = CDbl(str1)
             End If
             If SWG.TextMatrix(K, 10) = "DA" Then
                Individualfrm.txtPCurrent(K) = str1
                Individualfrm.SSTab2.TabCaption(K) = str1 & "/" & txtRatioR2(cmbSelctCore - 1) & " A"
                Individualfrm.txtVA(K) = SWG.TextMatrix(K + 1, 3)
                Individualfrm.txtVK(K) = SWG.TextMatrix(K + 1, 6)
                Individualfrm.txtRct(K) = SWG.TextMatrix(K + 1, 7)
                Individualfrm.txtClass(K) = SWG.TextMatrix(K + 1, 4)
                Individualfrm.txtALF(K) = SWG.TextMatrix(K + 1, 5)
                Individualfrm.txtIe(K) = SWG.TextMatrix(K + 1, 8)
                Individualfrm.txtVAt(K) = SWG.TextMatrix(K + 1, 9)
                Individualfrm.txtSTurns(K) = str1 * txtTurns / txtRatioR2(cmbSelctCore - 1)
                Individualfrm.txtMinCsSec(K) = CTfrm.txtRateConti.Tag
             Else '
                Individualfrm.SSTab2.TabCaption(K) = str1 & "/" & txtRatioR2(cmbSelctCore - 1) & " A"
                Individualfrm.txtPCurrent(K) = str1
                Individualfrm.txtVA(K) = txtVA1
                Individualfrm.txtVK(K) = txtVK1
                Individualfrm.txtRct(K) = txtRCT1
                Individualfrm.txtClass(K) = txtClass1
                Individualfrm.txtIe(K) = txtIE1
                Individualfrm.txtALF(K) = txtISF1
                Individualfrm.txtVAt(K) = txtV1
                Individualfrm.txtSTurns(K) = str1 * txtTurns / txtRatioR2(cmbSelctCore - 1)
                Individualfrm.txtMinCsSec(K) = CTfrm.txtRateConti.Tag
'                MsgBox Individualfrm.txtSTurns(K)
             End If
             
             K = K + 1
             str1 = ""
        Else
            str1 = str1 & value
        End If
    Next
    
    
    
    Individualfrm.txtVA1.Tag = K
    
    '####secondary
    For i = 0 To K - 1
        Individualfrm.txtSecCurr(i) = txtRatioR2(cmbSelctCore - 1)
        If i = 0 Then
            Individualfrm.txtToTurns(i) = Individualfrm.txtSTurns(i)
'            MsgBox Individualfrm.txtToTurns(i)
        Else
            Individualfrm.txtToTurns(i) = Individualfrm.txtSTurns(i)
'            MsgBox Individualfrm.txtToTurns(i)
            For j = 0 To i - 1
                Individualfrm.txtToTurns(i) = Individualfrm.txtSTurns(i) - Individualfrm.txtSTurns(j)
'                MsgBox Individualfrm.txtToTurns(i)
            Next
        End If
    Next
    
Else
    
    Individualfrm.txtSTurns(0) = txtRatioDum(cmbSelctCore - 1) * txtTurns / txtRatioR2(cmbSelctCore - 1)
'    MsgBox Individualfrm.txtSTurns(0) txtMinCsSec
    Individualfrm.SSTab2.TabCaption(0) = txtRatio(cmbSelctCore - 1) & "/" & txtRatioR2(cmbSelctCore - 1) & " A"
    Individualfrm.txtPCurrent(0) = txtRatio(cmbSelctCore - 1)
    Individualfrm.txtVA(0) = txtVA1
    Individualfrm.txtVK(0) = txtVK1
    Individualfrm.txtRct(0) = txtRCT1
    Individualfrm.txtVAt(0) = txtV1
    Individualfrm.txtVA1.Tag = 0
    Individualfrm.txtSecCurr(0) = txtRatioR2(cmbSelctCore - 1)
    Individualfrm.txtMinCsSec(0) = MinCross
    Individualfrm.txtToTurns(0) = txtRatio(cmbSelctCore - 1)
    If txtNo = "S" Then
        Individualfrm.txtODn(0).Tag = txtODn
    ElseIf txtNo = "D" Then
        Individualfrm.txtODn(0).Tag = txtODn / 2
    ElseIf txtNo = "T" Then
        Individualfrm.txtODn(0).Tag = txtODn / 3
    ElseIf txtNo = "Q" Then
        Individualfrm.txtODn(0).Tag = txtODn / 4
    End If
    Individualfrm.txtODn(0) = txtODn
    Individualfrm.txtRct(0) = txtRCT1
    Individualfrm.txtTurns = txtTurns
    Individualfrm.txtHt = txtHt
    Individualfrm.txtSwg(0) = txtSwg
    Individualfrm.txtNo(0) = txtNo
    
'    MsgBox Individualfrm.txtPCurrent(0) ISF<=
End If
    Individualfrm.txtRatio = Individualfrm.txtRatio & " A"
For l = 0 To K - 1
    Individualfrm.txtHt = txtHt
    If ChK2.value = 1 Then
    Individualfrm.txtSwg(l) = SWG.TextMatrix(l + 1, 1)
    Individualfrm.txtNo(l) = SWG.TextMatrix(l + 1, 2)
    Else
    Individualfrm.txtSwg(l) = txtSwg
    Individualfrm.txtNo(l) = txtNo
       
    End If
Next
'Individualfrm.txtVA(0) = txtVA1
Individualfrm.TXTcOMP = txtCompen(0)
'If Type is metering then
If Individualfrm.txtCoreType = "Metering" Then
 'visible = FALSE
    For i = 0 To 4
        Individualfrm.ErrorRead(i).Visible = True
        Individualfrm.ErrorRead1(i).Visible = True
        Individualfrm.Label92(i).Visible = False
        Individualfrm.Label91(i).Visible = False
        Individualfrm.Label81(i).Visible = False
        Individualfrm.Label82(i).Visible = False
        Individualfrm.Label83(i).Visible = False
        Individualfrm.Label86(i).Visible = False
        Individualfrm.Label87(i).Visible = False
        
        Individualfrm.txtIeNor(i).Visible = False
        Individualfrm.txtIeALF(i).Visible = False
        Individualfrm.txtECatN(i).Visible = False
        Individualfrm.txtECatAlf(i).Visible = False
        Individualfrm.txtECurr(i).Visible = False
        Individualfrm.txtVKM1(i).Visible = False
        Individualfrm.txtRCTat75(i).Visible = False
        Individualfrm.txtC1(i).Visible = False
        Individualfrm.txtC2(i).Visible = False
        
        Individualfrm.Frame20.Visible = False
        Individualfrm.Frame19.Visible = True
        
    Next
    
ElseIf Individualfrm.txtCoreType = "Protection" Then
'If Type is Protection

    For i = 0 To 4
        Individualfrm.ErrorRead(i).Visible = True
        Individualfrm.ErrorRead1(i).Visible = False
        
        Individualfrm.Label92(i).Visible = True
        Individualfrm.Label91(i).Visible = True
        Individualfrm.Label86(i).Visible = True
        Individualfrm.Label87(i).Visible = True
        Individualfrm.txtIeNor(i).Visible = True
        Individualfrm.txtIeALF(i).Visible = True
        Individualfrm.txtECatN(i).Visible = True
        Individualfrm.txtECatAlf(i).Visible = True
        
        Individualfrm.Label81(i).Visible = False
        Individualfrm.Label82(i).Visible = False
        Individualfrm.Label83(i).Visible = False
        Individualfrm.Label85(i).Visible = False
        Individualfrm.txtECurr(i).Visible = False
        Individualfrm.txtVKM1(i).Visible = False
        Individualfrm.txtRCTat75(i).Visible = False
        Individualfrm.txtC1(i).Visible = False
        Individualfrm.txtC2(i).Visible = False
        Individualfrm.txtISF(i).Visible = False
        
        Individualfrm.Frame20.Visible = False
        Individualfrm.Frame19.Visible = True
    Next

ElseIf Individualfrm.txtCoreType = "Special Protection" Then
'iF Type is PS
    For i = 0 To 4
        Individualfrm.ErrorRead(i).Visible = False
        Individualfrm.ErrorRead1(i).Visible = False
        
        Individualfrm.Label92(i).Visible = False
        Individualfrm.Label91(i).Visible = False
        Individualfrm.Label86(i).Visible = False
        Individualfrm.Label87(i).Visible = False
        Individualfrm.txtIeNor(i).Visible = False
        Individualfrm.txtIeALF(i).Visible = False
        Individualfrm.txtECatN(i).Visible = False
        Individualfrm.txtECatAlf(i).Visible = False
        
        Individualfrm.Label81(i).Visible = True
        Individualfrm.Label82(i).Visible = True
        Individualfrm.Label83(i).Visible = True
        Individualfrm.Label85(i).Visible = False
        Individualfrm.txtECurr(i).Visible = True
        Individualfrm.txtVKM1(i).Visible = True
        Individualfrm.txtRCTat75(i).Visible = True
        Individualfrm.txtC1(i).Visible = True
        Individualfrm.txtC2(i).Visible = True
        Individualfrm.txtISF(i).Visible = False
        
        Individualfrm.Label91(i).Visible = True
        Individualfrm.Label92(i).Visible = True
        Individualfrm.Frame20.Visible = True
        Individualfrm.Frame19.Visible = False
        
    Next
    
End If

GNo = Gnumber
Dim ss As String
For i = 0 To K - 1
    ss = Individualfrm.txtNo(i)
    Individualfrm.txtNo(i) = ""
    Individualfrm.txtNo(i) = ss
Next
Dim TWH As Double
'Calculating total layers
    If CoreTap = "N" Then
    TWH = Round(Individualfrm.txtWHt(0), 2)
'    TWH = RoundUp(TWH)
    Individualfrm.txtTWh = TWH
    Individualfrm.txtTolayers = Individualfrm.txtLayers(0)
    Individualfrm.txtCoppWt = Individualfrm.txtCWt(0)
    Individualfrm.Label9(0) = "S1 - S2 = " & Individualfrm.txtToTurns(0)
    If txtnoofcores > 1 Then
        For f = 0 To 1
            If f = 0 Then
                Individualfrm.Label9(0).Caption = "(" & cmbSelctCore & "S" & (f + 1) & "-" & cmbSelctCore & "S" & f + 2 & ")"  ' &  "=" & Individualfrm.txtToTurns(f) & ")"
                Individualfrm.Label9(0).Visible = False
                Individualfrm.Label60(0) = "Turns : " & Individualfrm.Label9(0).Caption
                Individualfrm.Label0(0) = "Turns Between :" & Individualfrm.Label9(0).Caption & " :"
            Else
                For t = 0 To f
                    Marking = Marking & " , " & cmbSelctCore & "S" & (t + 1) & "-" & cmbSelctCore & "S" & (t + 2) & "=" & Individualfrm.txtToTurns(f)
                    If t = f Then
                        'Between Turn Marking
                        Individualfrm.Label0(f) = "Turns Between : (" & cmbSelctCore & "S" & (t + 1) & "-" & cmbSelctCore & "S" & (t + 2) & ")" '& "=" & Individualfrm.txtToTurns(f)
                    End If
                Next
                'Total Turns
                Individualfrm.Label9(f).Caption = cmbSelctCore & "S1" & "-" & cmbSelctCore & "S" & f + 2 & "=" & Individualfrm.txtSTurns(f) & " =>" & Marking
                Individualfrm.Label60(f) = "Turns : (" & cmbSelctCore & "S1" & "-" & cmbSelctCore & "S" & f + 2 & ")" '& "=" & Individualfrm.txtSTurns(f)
            End If
        Next
    Else
    'If only one core is present
        Dim Counterr As Integer
        If Individualfrm.txtVA1.Tag = 0 Then
            Counterr = Individualfrm.txtVA1.Tag
        Else
            Counterr = Individualfrm.txtVA1.Tag - 1
        End If
        For f = 0 To Counterr
            If f = 0 Then
                Individualfrm.Label9(0).Caption = "(" & "S" & (f + 1) & "-" & "S" & f + 2 & ")"  ' &  "=" & Individualfrm.txtToTurns(f) & ")"
                Individualfrm.Label60(0) = "Turns : " & Individualfrm.Label9(0).Caption
                Individualfrm.Label0(0) = "Turns Between :" & Individualfrm.Label9(0).Caption & " :"
            Else
                For t = 0 To f
                    Marking = Marking & " , " & "S" & (t + 1) & "-" & "S" & (t + 2) & "=" & Individualfrm.txtToTurns(t)
                    If t = f Then
                        'Between Turn Marking
                        Individualfrm.Label0(f) = "Turns Between : (" & "S" & (t + 1) & "-" & "S" & (t + 2) & ")" '& "=" & Individualfrm.txtToTurns(f)
                    End If
                Next
                'Total Turns
                Individualfrm.Label9(f).Caption = "S1" & "-" & "S" & f + 2 & "=" & Individualfrm.txtSTurns(f) & " =>" & Marking
                Individualfrm.Label60(f) = "Turns : (" & "S1" & "-" & "S" & f + 2 & ")" '& "=" & Individualfrm.txtSTurns(f)
            End If
        Next
    End If
Else
    
    Dim nooflayers As Integer
    If DiaType = "D" Then
        TWH = 1.5
        Individualfrm.txtCoppWt = 0
        For i = 0 To Individualfrm.txtVA1.Tag - 1
'            MsgBox (Individualfrm.txtLayers(i) * Individualfrm.txtODs(i))
            TWH = TWH + (Individualfrm.txtLayers(i) * Individualfrm.txtODs(i))
            Individualfrm.txtCoppWt = Individualfrm.txtCoppWt + Individualfrm.txtCWt(i)
            nooflayers = nooflayers + Individualfrm.txtLayers(i)
            Individualfrm.txtCoppWt = Individualfrm.txtCoppWt + Individualfrm.txtCWt(i).Text
        Next
        TWH = TWH + (0.25 * (nooflayers - 1)) + (1.5 + 3)
'        TWH = RoundUp(TWH)
        Individualfrm.txtTWh = TWH
        Individualfrm.txtTolayers = nooflayers
        '' Total Cu wt
        Individualfrm.txtCoppWt = 0
        For K = 0 To Individualfrm.txtVA1.Tag - 1
            Individualfrm.txtCoppWt = CDbl(Individualfrm.txtCoppWt) + CDbl(Individualfrm.txtCWt(K))
        Next
    Else
        Individualfrm.txtWHt(Individualfrm.txtVA1.Tag - 1) = Round(Individualfrm.txtWHt(Individualfrm.txtVA1.Tag - 1), 2)
'        TWH = RoundUp(Individualfrm.txtWHt(Individualfrm.txtVA1.Tag - 1))
        TWH = Individualfrm.txtWHt(Individualfrm.txtVA1.Tag - 1)
        Individualfrm.txtTWh = TWH 'Individualfrm.txtWHt(Individualfrm.txtVA1.Tag - 1)
        Individualfrm.txtTolayers = Individualfrm.txtLayers(Individualfrm.txtVA1.Tag - 1)
        If InStr(1, txtCoreRatio, "-") > 0 Then
            Individualfrm.txtCoppWt = Individualfrm.txtCWt(Individualfrm.txtVA1.Tag - 1)
        Else
            Individualfrm.txtCoppWt = Individualfrm.txtCWt(Individualfrm.txtVA1.Tag)
        End If
    End If
    Individualfrm.Label0(0).Visible = True
    If txtnoofcores > 1 Then
        For f = 0 To Individualfrm.txtVA1.Tag - 1
            If f = 0 Then
                Individualfrm.Label9(0).Caption = "(" & cmbSelctCore & "S" & (f + 1) & "-" & cmbSelctCore & "S" & f + 2 & ")"  ' &  "=" & Individualfrm.txtToTurns(f) & ")"
                Individualfrm.Label60(0) = "Turns : " & Individualfrm.Label9(0).Caption
                Individualfrm.Label0(0) = "Turns Between :" & Individualfrm.Label9(0).Caption & " :"
            Else
                For t = 0 To f
                    Marking = Marking & " , " & cmbSelctCore & "S" & (t + 1) & "-" & cmbSelctCore & "S" & (t + 2) & "=" & Individualfrm.txtToTurns(f)
                    If t = f Then
                        'Between Turn Marking
                        Individualfrm.Label0(f) = "Turns Between : (" & cmbSelctCore & "S" & (t + 1) & "-" & cmbSelctCore & "S" & (t + 2) & ")" '& "=" & Individualfrm.txtToTurns(f)
                    End If
                Next
                'Total Turns
                Individualfrm.Label9(f).Caption = cmbSelctCore & "S1" & "-" & cmbSelctCore & "S" & f + 2 & "=" & Individualfrm.txtSTurns(f) & " =>" & Marking
                Individualfrm.Label60(f) = "Turns : (" & cmbSelctCore & "S1" & "-" & cmbSelctCore & "S" & f + 2 & ")" '& "=" & Individualfrm.txtSTurns(f)
            End If
        Next
    Else
    'If only one core is present
        For f = 0 To Individualfrm.txtVA1.Tag - 1
            If f = 0 Then
                Individualfrm.Label9(0).Caption = "(" & "S" & (f + 1) & "-" & "S" & f + 2 & ")"  ' &  "=" & Individualfrm.txtToTurns(f) & ")"
                Individualfrm.Label60(0) = "Turns : " & Individualfrm.Label9(0).Caption
                Individualfrm.Label0(0) = "Turns Between :" & Individualfrm.Label9(0).Caption & " :"
            Else
                For t = 0 To f
                    Marking = Marking & " , " & "S" & (t + 1) & "-" & "S" & (t + 2) & "=" & Individualfrm.txtToTurns(t)
                    If t = f Then
                        'Between Turn Marking
                        Individualfrm.Label0(f) = "Turns Between : (" & "S" & (t + 1) & "-" & "S" & (t + 2) & ")" '& "=" & Individualfrm.txtToTurns(f)
                    End If
                Next
                'Total Turns
                Individualfrm.Label9(f).Caption = "S1" & "-" & "S" & f + 2 & "=" & Individualfrm.txtSTurns(f) & " =>" & Marking
                Individualfrm.Label60(f) = "Turns : (" & "S1" & "-" & "S" & f + 2 & ")" '& "=" & Individualfrm.txtSTurns(f)
            End If
        Next
    End If
End If

If Individualfrm.cmdOK.Enabled = True Then
    Individualfrm.cmdOK.SetFocus
End If


End Sub

Private Sub Chkoption_Click()

Exit Sub
'Individualfrm.Show
'prcFillInd


GCross.Rows = 1
Dim value As String * 1
If Chkoption.value = 1 Then
    ChK2.Enabled = True
'    ChK2.value = 0
    For i = 1 To Len(txtRatio(cmbSelctCore - 1)) + 1
        value = Mid(txtRatio(cmbSelctCore - 1), i, i)
        If value = "-" Or i > Len(txtRatio(cmbSelctCore - 1)) Then
             MaxCs = prcCalSecC(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), CDbl(str1), cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
             
             If MaxCs >= txtSecCS Then
             
                Set rst = New Recordset
                rst.Open "select * from secConductor where t_cs >=" & MaxCs & " ", con, adOpenStatic, adLockReadOnly
                If rst.EOF = False And rst.BOF = False Then
                    
                    GCross.Rows = GCross.Rows + 1
                    GCross.row = GCross.Rows - 1
                    If ChK2.value = 0 Then
                        GCross.TextMatrix(GCross.row, 0) = txtTurns * str1 / txtRatioR2(cmbSelctCore - 1)
                        GCross.TextMatrix(GCross.row, 1) = txtODs
                        GCross.TextMatrix(GCross.row, 2) = txtSwg
                        GCross.TextMatrix(GCross.row, 3) = txtODn
                        GCross.TextMatrix(GCross.row, 4) = txtNo
                        GCross.TextMatrix(GCross.row, 5) = txtODs
                        GCross.TextMatrix(GCross.row, 7) = txtTurns * str1 / txtRatioR2(cmbSelctCore - 1)
                        GCross.TextMatrix(GCross.row, 26) = str1
                    Else
                        GCross.TextMatrix(GCross.row, 0) = txtTurns * str1 / txtRatioR2(cmbSelctCore - 1)
                        GCross.TextMatrix(GCross.row, 1) = rst!t_ods
                        GCross.TextMatrix(GCross.row, 2) = rst!t_swg
                        GCross.TextMatrix(GCross.row, 3) = rst!t_cs
                        GCross.TextMatrix(GCross.row, 4) = "S"
                        GCross.TextMatrix(GCross.row, 5) = rst!t_ods
                        GCross.TextMatrix(GCross.row, 7) = txtTurns * str1 / txtRatioR2(cmbSelctCore - 1)
                        GCross.TextMatrix(GCross.row, 26) = str1
                    End If
                End If
                Set rst = Nothing
                str1 = ""
             Else
                
                Set rst = New Recordset
                rst.Open "select * from secConductor where t_cs >=" & txtSecCS & " ", con, adOpenStatic, adLockReadOnly
                If rst.EOF = False And rst.BOF = False Then
                                                  
                    GCross.Rows = GCross.Rows + 1
                    GCross.row = GCross.Rows - 1
                    If ChK2.value = 0 Then
                        GCross.TextMatrix(GCross.row, 0) = txtTurns * str1 / txtRatioR2(cmbSelctCore - 1)
                        GCross.TextMatrix(GCross.row, 1) = txtODs
                        GCross.TextMatrix(GCross.row, 2) = txtSwg
                        GCross.TextMatrix(GCross.row, 3) = txtODn
                        GCross.TextMatrix(GCross.row, 4) = txtNo
                        GCross.TextMatrix(GCross.row, 5) = txtODs
                        GCross.TextMatrix(GCross.row, 7) = txtTurns * str1 / txtRatioR2(cmbSelctCore - 1)
                        GCross.TextMatrix(GCross.row, 26) = str1
 
                    Else
                        GCross.TextMatrix(GCross.row, 0) = txtTurns * str1 / txtRatioR2(cmbSelctCore - 1)
                        GCross.TextMatrix(GCross.row, 1) = rst!t_ods
                        GCross.TextMatrix(GCross.row, 2) = rst!t_swg
                        GCross.TextMatrix(GCross.row, 3) = rst!t_cs
                        GCross.TextMatrix(GCross.row, 4) = "S"
                        GCross.TextMatrix(GCross.row, 5) = rst!t_ods
                        GCross.TextMatrix(GCross.row, 7) = txtTurns * str1 / txtRatioR2(cmbSelctCore - 1)
                        GCross.TextMatrix(GCross.row, 26) = str1
                    End If
                End If
                Set rst = Nothing
                
             End If

             str1 = ""
        Else
            str1 = str1 & value
        End If
        
    Next
    
    If Val(GCross.TextMatrix(1, 0)) < Val(GCross.TextMatrix(GCross.Rows - 1, 0)) Then
                
        Dim l As Integer
        l = GCross.Rows - 1
                
        Do Until l = 1
            If l > 1 Then
                GCross.TextMatrix(l, 0) = Val(GCross.TextMatrix(l, 0)) - Val(GCross.TextMatrix(l - 1, 0))
                l = l - 1
            End If
        Loop
        Grid.Tag = "A"
    ElseIf Val(GCross.TextMatrix(1, 0)) > Val(GCross.TextMatrix(GCross.Rows - 1, 0)) Then

        l = 1
                
        Do Until l = GCross.Rows - 1
            If l < GCross.Rows - 1 Then
                GCross.TextMatrix(l, 0) = Val(GCross.TextMatrix(l, 0)) - Val(GCross.TextMatrix(l + 1, 0))
                l = l + 1
            End If
        Loop
        Grid.Tag = ""
    End If
    
Else
    ChK2.Enabled = False
'    ChK2.value = 0
    Dim mincurr As Double
    For i = 1 To Len(txtRatio(cmbSelctCore - 1)) + 1
    value = Mid(txtRatio(cmbSelctCore - 1), i, i)
    If value = "-" Or i > Len(txtRatio(cmbSelctCore - 1)) Then
         If CDbl(str1) < mincurr Or mincurr = 0 Then
              mincurr = CDbl(str1)
         End If
         str1 = ""
    Else
        str1 = str1 & value
    End If
    Next
    txtMinCsSec = prcCalSecC(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), mincurr, cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
    If CDbl(txtSecCS) >= CDbl(txtMinCsSec) Then
        
        Set rst = New ADODB.Recordset
        rst.Open "select * from secconductor where t_cs >= " & txtMinCsSec & " ", con, adOpenStatic, adLockReadOnly
        
        If rst.EOF = False And rst.BOF = False Then
            txtSwg = rst!t_swg
            txtNo = "S"
            txtODs = rst!t_ods
            txtODn = rst!t_cs
            txtODn.Tag = rst!t_cs
        End If
            
        txtMinCsSec = txtSecCS
            
    Else
        
        Set rst = New ADODB.Recordset
        rst.Open "select * from secconductor where t_cs >= " & txtMinCsSec & " ", con, adOpenStatic, adLockReadOnly
        
        If rst.EOF = False And rst.BOF = False Then
            txtSwg = rst!t_swg
            txtNo = "S"
            txtODs = rst!t_ods
            txtODn = rst!t_cs
            txtODn.Tag = rst!t_cs
        End If
'        txtSecCS = txtMinCsSec
    End If
    
    Set rst = Nothing

End If
End Sub


Private Sub cmbClass_GotFocus()
'cmbClass.SelStart = 0
'cmbClass.SelLength = Len(cmbClass)
MDIDesign.StatusBar1.Panels(2).Text = "Set Class"
End Sub

Private Sub cmbClass_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbConStyle_Change()
Flag1 = True
txtTurns.ListIndex = -1
txtTop = 0
txtSide = 0
txtCondPer = "-"
txtTotalCon = "-"
If Len(Trim(cmbConStyle)) > 0 Then
    PrcCalPrimaryDetails
    If Len(cmbSTC) > 0 Then
        txtIDyn = CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1)) * 2.5
    End If
End If
Flag1 = False


Exit Sub

''
On Error Resume Next

If InStr(1, UCase(Trim(cmbEquipType)), "WOUND") > 0 And InStr(1, UCase(Trim(cmbEquipType)), "WINDOW") Then
    txtTurns = 1
    Exit Sub
End If
If cmbConStyle.Text = "Double Coil - Normal Wdg Ht." Then
    txtWidth = 48
    coiltype = "P"
    CoilName = "Double Coil - Normal Wdg Ht."
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "Double Coil - Equall Wdg Ht." Then
    txtWidth = 48
    coiltype = "P"
    CoilName = ""
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "Double Coil - Bare Conductor" Then
    txtWidth = 42
    txtTurns = 1
    coiltype = "B"
    txtTurns = 1
    CoilName = "Double Coil - Normal Wdg Ht."
ElseIf cmbConStyle.Text = "Single Coil - Bare Conductor" Then
    txtWidth = 21
    txtTurns = 1
    coiltype = "B"
    txtTurns = 1
    CoilName = ""
ElseIf cmbConStyle.Text = "Single Coil" Then
    txtWidth = 48
    coiltype = "P"
    CoilName = ""
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "Series Coil" Then
    txtWidth = 24
    coiltype = "P"
    CoilName = "Series Coil"
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "BAR-C" Then
    txtWidth = 48
    coiltype = "R"
    CoilName = "C"
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "BAR-R" Then
    txtWidth = 24
    coiltype = "R"
    CoilName = "R"
    txtTurns.SetFocus
End If
''SearchForStyle
'If Len(txtTurns) > 0 Then
'    PrcCalculate
'End If
End Sub

Private Sub cmbConStyle_Click()
Flag1 = True
txtTurns.ListIndex = -1
txtTop = 0
txtSide = 0
txtCondPer = ""
txtTotalCon = ""
If Len(Trim(cmbConStyle)) > 0 Then
    PrcCalPrimaryDetails
    If Len(cmbSTC) > 0 Then
        txtIDyn = CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1)) * 2.5
    End If
End If
Flag1 = False


Exit Sub

''
On Error Resume Next

If InStr(1, UCase(Trim(cmbEquipType)), "WOUND") > 0 And InStr(1, UCase(Trim(cmbEquipType)), "WINDOW") Then
    txtTurns = 1
    Exit Sub
End If
If cmbConStyle.Text = "Double Coil - Normal Wdg Ht." Then
    txtWidth = 48
    coiltype = "P"
    CoilName = "Double Coil - Normal Wdg Ht."
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "Double Coil - Equall Wdg Ht." Then
    txtWidth = 48
    coiltype = "P"
    CoilName = ""
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "Double Coil - Bare Conductor" Then
    txtWidth = 42
    txtTurns = 1
    coiltype = "B"
    txtTurns = 1
    CoilName = "Double Coil - Normal Wdg Ht."
ElseIf cmbConStyle.Text = "Single Coil - Bare Conductor" Then
    txtWidth = 21
    txtTurns = 1
    coiltype = "B"
    txtTurns = 1
    CoilName = ""
ElseIf cmbConStyle.Text = "Single Coil" Then
    txtWidth = 48
    coiltype = "P"
    CoilName = ""
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "Series Coil" Then
    txtWidth = 24
    coiltype = "P"
    CoilName = "Series Coil"
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "BAR-C" Then
    txtWidth = 48
    coiltype = "R"
    CoilName = "C"
    txtTurns.SetFocus
ElseIf cmbConStyle.Text = "BAR-R" Then
    txtWidth = 24
    coiltype = "R"
    CoilName = "R"
    txtTurns.SetFocus
End If
''SearchForStyle
'If Len(txtTurns) > 0 Then
'    PrcCalculate
'End If
End Sub

Private Sub cmbConStyle_GotFocus()
On Error GoTo err
If Len(Trim(txtDelStc)) = 0 Then
    Exit Sub
End If
If InStr(1, UCase(cmbEquipType), "WOUND") > 0 Or InStr(1, UCase(cmbEquipType), "BAR") > 0 Then
    If CDbl(txtDelStc) > 220 Then
        MsgBox "STC Current desity is greater than 180", vbCritical
        txtTurns.SetFocus
        Exit Sub
    End If
End If
cmbConStyle.SelStart = 0
cmbConStyle.SelLength = Len(cmbConStyle)
MDIDesign.StatusBar1.Panels(2).Text = "Set conductor style"
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub cmbConStyle_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

'Private Sub cmbConStyle_KeyPress(KeyAscii As Integer)
'If KeyAscii = 0 Then
'    KeyAscii = 0
'End If
'End Sub

Private Sub cmbCore_Change()
If OptCom(0).value = True Then
    Frame11.Caption = "For all Cores : "
    Frame10.Caption = "For all Cores : "
Else
    Frame11.Caption = "Size For " & cmbCore & " : "
    Frame10.Caption = "Size For " & cmbCore & " : "
End If
End Sub

Private Sub cmbCore_Click()
If OptCom(0).value = True Then
    Frame11.Caption = "For all Cores : "
    Frame10.Caption = "For all Cores : "
Else
    Frame11.Caption = "Size For " & cmbCore & " : "
    Frame10.Caption = "Size For " & cmbCore & " : "
End If
End Sub

Private Sub cmbCType_Change()

If Len(Trim(cmbnoofcores)) > 0 Then
    If Trim(cmbCType) = "-Select Core Type-" Then
        Frame7.Enabled = False
        Frame6.Enabled = False
    Else
        Frame7.Enabled = True
        Frame6.Enabled = True
    End If
End If

If cmbCType.Text = "Special Protection" Then
    Frame7.Visible = True
    Frame6.Visible = False
Else
    Frame7.Visible = False
    Frame6.Visible = True
End If
txtVK.Text = ""
txtRct.Text = ""
txtIe.Text = ""
txtIe2.Text = ""
'cmbVA.Text = ""
If cmbCType = "Metering" Then
    cmbClass.Clear
    cmbClass.AddItem "0.1"
    cmbClass.AddItem "0.2"
    cmbClass.AddItem "0.5"
    cmbClass.AddItem "1.0"
    cmbClass.AddItem "3.0"
    cmbClass.AddItem "5.0"
    cmbISF.Clear
    cmbISF.AddItem "5"
    cmbISF.AddItem "10"
    cmbISF.AddItem "15"
    cmbISF.AddItem "20"
    cmbISF.AddItem "25"
    Frame6.Caption = "Metering (Core " & cmbnoofcores & ") :"
ElseIf cmbCType = "Protection" Then
    cmbClass.Clear
    cmbClass.AddItem "5P"
    cmbClass.AddItem "10P"
    cmbClass.AddItem "15P"
    cmbISF.Clear
    cmbISF.AddItem "10"
    cmbISF.AddItem "15"
    cmbISF.AddItem "20"
    cmbISF.AddItem "30"
    Frame6.Caption = "Protection (Core " & cmbnoofcores & ") :"
ElseIf cmbCType = "Special Protection" Then
    Frame6.Caption = "Special Protection (Core " & cmbnoofcores & ") :"
End If
End Sub

Private Sub cmbCType_Click()

If Len(Trim(cmbnoofcores)) > 0 Then
    If Trim(cmbCType) = "-Select Core Type-" Then
        Frame7.Enabled = False
        Frame6.Enabled = False
    Else
        Frame7.Enabled = True
        Frame6.Enabled = True
    End If
End If

If cmbCType.Text = "Special Protection" Then
    Frame7.Visible = True
    Frame6.Visible = False
Else
    Frame7.Visible = False
    Frame6.Visible = True

End If
cmdVA.ListIndex = -1
cmbClass.ListIndex = -1
cmbISF.ListIndex = -1
txtVK.Text = ""
txtRct.Text = ""
txtIe.Text = ""
txtIe2.Text = ""
If cmbCType = "Metering" Then
    cmbClass.Clear
    cmbClass.AddItem "0.1"
    cmbClass.AddItem "0.2"
    cmbClass.AddItem "0.5"
    cmbClass.AddItem "1.0"
    cmbClass.AddItem "3.0"
    cmbClass.AddItem "5.0"
    cmbISF.Clear
    cmbISF.AddItem "5"
    cmbISF.AddItem "10"
    cmbISF.AddItem "15"
    cmbISF.AddItem "20"
    cmbISF.AddItem "25"
    Frame6.Caption = "Metering (Core " & cmbnoofcores & ") :"
    Label15.Caption = "ISF<=:"
ElseIf cmbCType = "Protection" Then
    cmbClass.Clear
    cmbClass.AddItem "5P"
    cmbClass.AddItem "10P"
    cmbClass.AddItem "15P"
    cmbISF.Clear
    cmbISF.AddItem "10"
    cmbISF.AddItem "15"
    cmbISF.AddItem "20"
    cmbISF.AddItem "30"
    Label15.Caption = "ALF :"
    Frame6.Caption = "Protection (Core " & cmbnoofcores & ") :"
ElseIf cmbCType = "Special Protection" Then
    Frame6.Caption = "Special Protection (Core " & cmbnoofcores & ") :"
End If
End Sub

Private Sub cmbCType_GotFocus()
cmbCType.SelStart = 0
cmbType.SelLength = Len(cmbCType)
MDIDesign.StatusBar1.Panels(2).Text = "Set Core Type"
End Sub

Private Sub cmbCType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

'cmbEquipType.AddItem
Private Sub cmbEquipType_Change()
LastSet = ""
'ChangeReadings
CleanAll
Frame1.Enabled = True
PcrConductor
'Wound HTCTs ( Wo-R)
'cmbConStyle.Enabled = False cmbEquipType.ListIndex
If UCase(Trim(cmbEquipType)) = UCase("Bar HTCTs ( B-C)") Then
    txtTurns = ""
    txtTurns = 1
    SetPTurnComm = "Y"
    DisApp (False)
    LastSet = "F"
    WindowSize.Visible = True
    cmbStyle.ListIndex = 0
ElseIf UCase(Trim(cmbEquipType)) = UCase("Bar HTCTs ( B-R )") Then
    txtTurns = ""
    txtTurns = 1
    SetPTurnComm = "Y"
    DisApp (False)
    LastSet = "F"
    WindowSize.Visible = True
    cmbStyle.ListIndex = 1
ElseIf UCase(Trim(cmbEquipType)) = UCase("Window HTCTs ( Win-C )") Then
    txtTurns = ""
    txtTurns = 1
    SetPTurnComm = "Y"
    DisApp (False)
    LastSet = "F"
    WindowSize.Visible = True
    cmbStyle.ListIndex = 0
ElseIf UCase(Trim(cmbEquipType)) = UCase("Window HTCTs ( Win-R )") Then
    txtTurns = ""
    txtTurns = 1
    SetPTurnComm = "Y"
    DisApp (False)
    LastSet = "F"
    WindowSize.Visible = True
    cmbStyle.ListIndex = 1
ElseIf UCase(Trim(cmbEquipType)) = UCase("Wound HTCTs ( Wo-C )") Then
    txtTurns = ""
    txtTurns = 1
    SetPTurnComm = "Y"
    DisApp (True)
    LastSet = ""
    cmbConStyle.Enabled = True
    cmbStyle.ListIndex = 0
ElseIf UCase(Trim(cmbEquipType)) = UCase("Wound HTCTs ( Wo-R)") Then
    txtTurns = ""
    txtTurns = 1
    SetPTurnComm = "Y"
    DisApp (True)
    LastSet = ""
    cmbConStyle.Enabled = True
    cmbStyle.ListIndex = 1
Else
    txtTurns = 0
End If
If InStr(1, UCase(cmbEquipType), "WOUND") > 0 Then
    MDIDesign.StatusBar1.Panels(5).Text = "Wound HTCT"
ElseIf InStr(1, UCase(cmbEquipType), "WINDOW") > 0 Then
    MDIDesign.StatusBar1.Panels(5).Text = "Window Type CT"
ElseIf InStr(1, UCase(cmbEquipType), "BAR") > 0 Then
    MDIDesign.StatusBar1.Panels(5).Text = "Bar Primary CT"
End If
End Sub

Private Sub cmbEquipType_Click()
'cmbConStyle.Enabled = False
LastSet = ""
CleanAll
Frame1.Enabled = True
PcrConductor
cmbStyle.ListIndex = 1
If UCase(Trim(cmbEquipType)) = UCase("Bar HTCTs ( B-C)") Then
    txtTurns = Val(txtTurns)
    txtTurns = 1
    SetPTurnComm = "Y"
    DisApp (False)
    LastSet = "F"
    WindowSize.Visible = True
    cmbStyle.ListIndex = 0
ElseIf UCase(Trim(cmbEquipType)) = UCase("Bar HTCTs ( B-R )") Then
    txtTurns = Val(txtTurns)
    txtTurns = 1
    SetPTurnComm = "Y"
    DisApp (False)
    WindowSize.Visible = True
    LastSet = "F"
ElseIf UCase(Trim(cmbEquipType)) = UCase("Window HTCTs ( Win-C )") Then
    txtTurns = Val(txtTurns)
    txtTurns = 1
    SetPTurnComm = "Y"
    WindowSize.Visible = True
    DisApp (False)
    LastSet = "F"
    cmbStyle.ListIndex = 0
ElseIf UCase(Trim(cmbEquipType)) = UCase("Window HTCTs ( Win-R )") Then
    txtTurns = Val(txtTurns)
    txtTurns = 1
    SetPTurnComm = "Y"
    WindowSize.Visible = True
    DisApp (False)
    LastSet = "F"
ElseIf UCase(Trim(cmbEquipType)) = UCase("Wound HTCTs (Wo-C)") Then
    txtTurns = Val(txtTurns)
    txtTurns = 1
    SetPTurnComm = "Y"
    DisApp (True)
    LastSet = ""
    cmbConStyle.Enabled = True
    FillConductorStyle
    cmbStyle.ListIndex = 0
ElseIf UCase(Trim(cmbEquipType)) = UCase("Wound HTCTs ( Wo-R)") Then
    txtTurns = Val(txtTurns)
    SetPTurnComm = "Y"
    DisApp (True)
    cmbConStyle.Enabled = True
    txtTurns = 1
    PcrConductor
    FillConductorStyle
    LastSet = ""
    
Else
    txtTurns = 0
End If
If InStr(1, UCase(cmbEquipType), "WOUND") > 0 Then
    MDIDesign.StatusBar1.Panels(5).Text = "Wound HTCT"
ElseIf InStr(1, UCase(cmbEquipType), "WINDOW") > 0 Then
    MDIDesign.StatusBar1.Panels(5).Text = "Window Type CT"
ElseIf InStr(1, UCase(cmbEquipType), "BAR") > 0 Then
    MDIDesign.StatusBar1.Panels(5).Text = "Bar Primary CT"
End If

End Sub

Private Sub cmbEquipType_GotFocus()
cmbEquipType.SelStart = 0
cmbEquipType.SelLength = Len(cmbEquipType)
MDIDesign.StatusBar1.Panels(2).Text = "Select Equipment Type"
End Sub

Private Sub cmbEquipType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbEquipType_LostFocus()
If Len(cmbEquipType) = 0 Then
    MsgBox "Please enter equipment Type", vbInformation
     cmbEquipType.SetFocus
End If
End Sub

Private Sub cmbFreq_Click()
    If G1.Cols > 1 Then
        MsgBox "Core data will changed", vbCritical
        Exit Sub
    End If
End Sub

Private Sub cmbFreq_GotFocus()
cmbFreq.SelStart = 0
cmbFreq.SelLength = Len(cmbFreq)
MDIDesign.StatusBar1.Panels(2).Text = "Select Freq Hz"
End Sub

Private Sub cmbFreq_LostFocus()
If Len(cmbFreq) = 0 Then
    MsgBox "Please enter freq", vbExclamation
    cmbFreq.SetFocus
End If
End Sub

Private Sub cmbHSV_Change()

Set rst = New ADODB.Recordset

rst.Open "select il from il where il like '" & cmbHSV & "%'  ", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    rst.Requery
    cmbIL.Clear
    Do Until rst.EOF = True
        
        cmbIL.AddItem rst!iL
        rst.MoveNext
        
    Loop
    cmbIL.ListIndex = 0
End If

rst.Close

End Sub

Private Sub cmbHSV_Click()
Set rst = New ADODB.Recordset

rst.Open "select il from il where il like '" & cmbHSV & "%'  ", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    rst.Requery
    cmbIL.Clear
    Do Until rst.EOF = True
        
        cmbIL.AddItem rst!iL & " kVp"
        rst.MoveNext
        
    Loop
    cmbIL.ListIndex = 0
End If
rst.Close
End Sub

Private Sub cmbHSV_GotFocus()
cmbHSV.SelStart = 0
cmbHSV.SelLength = Len(cmbHSV)
MDIDesign.StatusBar1.Panels(2).Text = "Select Short Time Rating (kA)"
End Sub


Private Sub cmbIL_GotFocus()
    cmbIL.SelStart = 0
    cmbIL.SelLength = Len(cmbIL)
    MDIDesign.StatusBar1.Panels(2).Text = "select Insulation level"
End Sub

Private Sub cmbISF_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Set ISF / ALF"
End Sub

Private Sub cmbISF_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub cmbMetal_GotFocus()
cmbMetal.SelStart = 0
cmbMetal.SelLength = Len(cmbMetal)
End Sub

Private Function prcCheckforInv(str11 As String) As Boolean
Dim king As String
For j = 1 To Grid.Rows - 1
    king = ""
    If Grid.TextMatrix(j, 1) = str11 Then
        king = "P"
        prcCheckforInv = False
        Exit Function
    End If
Next
prcCheckforInv = True
End Function

Private Sub cmbMetal_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbnoofcores_Change()

On Error GoTo err

Dim value As String * 1
Dim str1 As String

CheckCore

Frame4.Enabled = True
If OptInd.value = True Then
    If Len(Trim(cmbnoofcores)) = 0 Then
        OptComm.value = True
        Exit Sub
    End If
    For i = 1 To Len((txtRatio(cmbnoofcores - 1))) + 1
        value = Mid(txtRatio(cmbnoofcores - 1), i, i)
        If value = "-" Or Len(txtRatio(cmbnoofcores - 1)) < i Then
            king = ""
           If prcCheckforInv(str1 & "/" & txtRatioR2(cmbnoofcores - 1)) Then
                king = "P"
            End If
            str1 = ""
        Else
            str1 = str1 & value
        End If
    Next
End If

If king <> "P" Then
    OptComm.value = True
    IvAcc.Clear
    IvAcc.Enabled = True
End If
For i = 0 To Val(txtnoofcores) - 1
    
    If Len(txtRatio(i)) = 0 Or Len(txtRatioR2(i)) = 0 Then
'        MsgBox "Please enter all core ratio", vbInformation
'        txtRatioDum(0).SetFocus
        err.Description = "Error"
        err.Number = 12000
        GoTo err
        Exit Sub
    End If

Next
If Val(cmbnoofcores) > 0 Then
    txtRatio2 = txtRatio(cmbnoofcores - 1) & "/" & txtRatioR2(cmbnoofcores - 1) & " A"
    txtRatio2R2 = txtRatioR2(cmbnoofcores - 1)
End If
'Frame7 =]
'Dim value As String * 1
'Dim str1 As String

If InStr(1, txtRatio2, "-") > 0 Then
    Frame4.Enabled = True
    OptCom(0).value = True
    For i = 1 To Len(txtRatio2) + 1
        value = Mid(txtRatio2, i, i)
        If value = "-" And Len(txtRatio2) < i Then
        
            IvAcc.AddItem str1
            str1 = ""
            
        Else
        
            str1 = str1 + value
        
        End If
    Next
Else
    Frame4.Enabled = False
    OptComm.value = True
End If
If cmbCType = "Metering" Then
    Frame6.Caption = "Metering (Core " & cmbnoofcores & ") :"
ElseIf cmbCType = "Protection" Then
    Frame6.Caption = "Protection (Core " & cmbnoofcores & ") :"
ElseIf cmbCType = "Special Protection" Then
    Frame6.Caption = "Special Protection (Core " & cmbnoofcores & ") :"
End If
Frame4.Enabled = True
If InStr(1, txtRatio2, "-") = 0 Then
    Frame4.Enabled = False
    OptComm.value = True
Else
    Frame4.Enabled = True
End If
If Len(cmbnoofcores) > 0 Then
    txtRatio2.Text = txtRatioDum(cmbnoofcores - 1) & "/" & txtRatioR2Dum(cmbnoofcores - 1)
End If
'    txtRatio2.Text = txtRatioDum(cmbnoofcores - 1) & "/" & txtRatioR2Dum(cmbnoofcores - 1)

err:

    If Len(err.Description) > 0 Then
        If err.Number <> 12000 Then
            MsgBox err.Description, vbCritical
        End If
    End If

End Sub

Private Sub cmbnoofcores_click()

If Len(cmbnoofcores) = 0 Then
    Exit Sub
End If

IvAcc.Clear

Frame4.Enabled = True
OptComm.value = True
IvAcc.Enabled = False

Dim AscRatio As ChangeRa
Set AscRatio = New ChangeRa

For p = 0 To txtnoofcores - 1
    If Len(Trim(txtRatioDum(p))) = 0 Then
        txtRatioDum(p).SetFocus
        Exit Sub
    End If
    txtRatio(p) = AscRatio.prcInterChangeRatio(txtRatioDum(p))
Next
    
Set AscRatio = Nothing

For i = 0 To Val(txtnoofcores) - 1
    
    If Len(txtRatio(i)) = 0 Or Len(txtRatioR2(i)) = 0 Then
        MsgBox "Please enter all core ratio", vbInformation
        Exit Sub
    End If

Next

txtRatio2 = txtRatioDum(cmbnoofcores - 1) & "/" & txtRatioR2Dum(cmbnoofcores - 1) & " A"

'Frame7 =]
Dim value As String * 1
Dim str1 As String

If InStr(1, txtRatio2, "-") > 0 Then
    Frame4.Enabled = True
    OptCom(0).value = True
    For i = 1 To Len(txtRatio2) + 1
        value = Mid(txtRatio2, i, i)
        If value = "-" And Len(txtRatio2) < i Then
        
            IvAcc.AddItem str1
            str1 = ""
            
        Else
        
            str1 = str1 + value
        
        End If
    Next
Else
    Frame4.Enabled = False
    OptComm.value = True
End If
If cmbCType = "Metering" Then
    Frame6.Caption = "Metering (Core " & cmbnoofcores & ") :"
ElseIf cmbCType = "Protection" Then
    Frame6.Caption = "Protection (Core " & cmbnoofcores & ") :"
ElseIf cmbCType = "Special Protection" Then
    Frame6.Caption = "Special Protection (Core " & cmbnoofcores & ") :"
End If
Frame4.Enabled = True
If InStr(1, txtRatio2, "-") = 0 Then
    Frame4.Enabled = False
    OptComm.value = True
    cmbnoofcores.SetFocus
Else
    Frame4.Enabled = True
End If
End Sub

Private Sub cmbnoofcores_GotFocus()
If Val(txtnoofcores) = 1 And Grid.Rows = 1 Then
    cmbnoofcores.ListIndex = 0
    txtRatio2.Text = txtRatioDum(cmbnoofcores - 1) & "/" & txtRatioR2Dum(cmbnoofcores - 1)
ElseIf Val(txtnoofcores) > 1 And Grid.Rows = 1 Then
    cmbnoofcores.ListIndex = 0
    txtRatio2.Text = txtRatioDum(cmbnoofcores - 1) & "/" & txtRatioR2Dum(cmbnoofcores - 1)
Else
'    cmbCType.ListIndex = -1
'    cmbCType.Text = "-Select Core Type-"
    If Len(cmbnoofcores) > 0 Then
        txtRatio2.Text = txtRatioDum(cmbnoofcores - 1) & "/" & txtRatioR2Dum(cmbnoofcores - 1)
    End If
End If
MDIDesign.StatusBar1.Panels(2).Text = "Select Core No"
End Sub

Private Sub CheckCore()
On Error GoTo err
cmbnoofcores_Change

err:
    If Len(err.Description) > 0 Then
        If err.Number = 12000 Then
            MsgBox "Enter all core ration", vbExclamation
            txtRatioDum(0).SetFocus
        Else
            MsgBox err.Description, vbExclamation
        End If
    End If

End Sub

Private Sub cmbNSV_Change()

On Error GoTo err

Set rst = New ADODB.Recordset
If InStr(1, UCase(cmbNSV), "K") > 0 Then
    rst.Open "select hsv from hsv where nsv like '" & Mid(cmbNSV, 1, InStr(1, UCase(cmbNSV), "K") - 1) & "%'  ", con, adOpenStatic, adLockReadOnly
Else
    rst.Open "select hsv from hsv where nsv like '" & cmbNSV & "%'  ", con, adOpenStatic, adLockReadOnly
End If

If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    rst.Requery
    cmbHSV.Clear
    Do Until rst.EOF = True
        
        cmbHSV.AddItem rst!hsv
        rst.MoveNext
        
    Loop
    cmbHSV.ListIndex = 0
End If

''INSULATION LEVEL IS TAKEN FROM DATABASE
Set rst = New ADODB.Recordset
rst.Open "select insclass, insclass1 from hsv where nsv like '" & Val(cmbNSV) & "'  ", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
    rst.MoveFirst
    CMBINSULATION.Clear
    CMBINSULATION.AddItem rst!insclass
    CMBINSULATION.AddItem rst!insclass1
End If
Set rst = Nothing


Set rst = Nothing
Grid3.Rows = 1
prcCoreSize



err:

    If Len(err.Description) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbExclamation
    End If
    
End Sub

Private Sub cmbnsv_Click()
Set rst = New ADODB.Recordset

If InStr(1, UCase(cmbNSV), "K") > 0 Then
    rst.Open "select hsv from hsv where nsv like '" & Mid(cmbNSV, 1, InStr(1, UCase(cmbNSV), "K") - 1) & "%'  ", con, adOpenStatic, adLockReadOnly
Else
    rst.Open "select hsv from hsv where nsv like '" & cmbNSV & "%'  ", con, adOpenStatic, adLockReadOnly
End If
If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    rst.Requery
    cmbHSV.Clear
    Do Until rst.EOF = True
        
        cmbHSV.AddItem rst!hsv
        rst.MoveNext
        
    Loop
    cmbHSV.ListIndex = 0
End If

Set rst = Nothing
Grid3.Rows = 1
prcCoreSize
End Sub

Private Sub cmbnsv_GotFocus()
cmbNSV.SelStart = 0
cmbNSV.SelLength = Len(cmbNSV)
MDIDesign.StatusBar1.Panels(2).Text = "Select Nominal System Volt. (kV)"
End Sub

Private Sub cmbNSV_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9.]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub cmbNSV_LostFocus()

If Len(Trim(cmbNSV)) = 0 Or UCase(Trim(cmbNSV)) = "KV" Then
    MsgBox "Please select / enter NSV", vbInformation
    If Len(Trim(cmbEquipType)) = 0 Then Exit Sub
    cmbNSV.SetFocus
    Exit Sub
End If

If InStr(1, UCase(cmbNSV), "K") = 0 Then
    cmbNSV = cmbNSV & " kV"
End If

If InStr(1, UCase(cmbHSV), "K") = 0 Then
    cmbHSV = cmbHSV & " kV"
End If
End Sub

Private Function PrcCheck_Core_For_Mod(Core As Integer) As Boolean
''finding core no from g1(if any) and bringing  back the details
''details done for same conductor not done for diffewrent conductor

For i = 1 To G1.Cols - 1
    If G1.TextMatrix(1, i) = Core Then
        txtSwg = G1.TextMatrix(7, i)
        txtNo = G1.TextMatrix(6, i)
        txtHt = G1.TextMatrix(29, i)
        cmbMetal = G1.TextMatrix(107, i)
        PrcCheck_Core_For_Mod = False
        Exit Function
    End If
Next
PrcCheck_Core_For_Mod = True

End Function

Private Sub cmbSelctCore_Change()
'Label70

On Error GoTo err

'If G1.Cols = 1 Then
'    cmbSelctCore = 1
'End If

Dim RateCs As Double
txtCompen(0) = 0
If Len(cmbSelctCore) = 0 Then
    Exit Sub
End If

InCalcu = "Y"
DiaType = "S"
'txtODn.Tag = ""
GCross.Rows = 1

'Giving Error message on inproper core entry
Dim shibuCount  As Integer
If Val(cmbSelctCore) = 0 Then
        MsgBox "Please enter proper core", vbInformation
    Exit Sub
End If

'if tap present in primary
If InStr(1, txtRatio(cmbSelctCore - 1), "-") > 0 Then
    CoreTap = "Y"
'    Chkoption.value = 1
    txtSwg.Visible = False
    txtNo.Visible = False
    ChK2.Visible = True
'    Line1.Visible = False
    Label67.Visible = False
Else
    CoreTap = "N"
    ChK2.Visible = False
'    Chkoption.value = 0
    txtSwg.Visible = True
    txtNo.Visible = True
'    Line1.Visible = True
    Label67.Visible = True
End If

'Finding Core shape , Size , BU , Mean of selected core
If cmbSelctCore <> 0 Or Len(cmbSelctCore) <> 0 Then
    For coreno = 1 To GCore.Rows - 1
        If cmbSelctCore = GCore.TextMatrix(coreno, 0) Then
            If Trim(cmbStyle) <> "Circular" Then
                txtOL1 = GCore.TextMatrix(coreno, 1)
                txtOW1 = GCore.TextMatrix(coreno, 2)
                txtWL1 = GCore.TextMatrix(coreno, 3)
                txtWW1 = GCore.TextMatrix(coreno, 4)
                
                Frame12.Visible = True
                Frame18.Visible = False
            Else
                txtOD1 = GCore.TextMatrix(coreno, 8)
                txtID1 = GCore.TextMatrix(coreno, 9)
                
                Frame12.Visible = False
                Frame18.Visible = True
            End If
            
            txtBUP = GCore.TextMatrix(coreno, 11)
            txtMean = GCore.TextMatrix(coreno, 10)
        End If
    Next
End If
'secondary turns calculations
If Len(txtTurns) <> 0 Then
    If InStr(1, txtRatio(CInt(cmbSelctCore - 1)), "-") = 0 Then
        If cmbStyle = "Rectangluar" Then
            txtSecTurn = txtTurns * txtRatio(CInt(cmbSelctCore - 1)) / txtRatioR2(cmbSelctCore - 1)
        Else
            txtSecTurn = 1 * txtRatio(CInt(cmbSelctCore - 1)) / txtRatioR2(cmbSelctCore - 1)
        End If
'        txtMinCsSec = prcCalSecC(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), CDbl(mincurr), cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
        txtSecCS = prcCalNorSecC(txtRatioR2(CInt(cmbSelctCore - 1)))
'        If Len(txtRateConti) > 0 Then
'            RateCs = prcCalRate(txtRateConti, CDbl(mincurr), txtRatioR2(CInt(cmbSelctCore - 1)))
'        End If
    '######Short Circuit Current Density
'        txtDelta = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), txtSecTurn, txtRatioR2(cmbSelctCore - 1), CDbl(txtODn))
        GNo = txtRatio(cmbSelctCore - 1)
    Else
    
        Dim int1 As Double
        Dim value As String * 1
        Dim ARatio As String
        Dim str1 As String
        Dim Gnumber As Double
        ARatio = txtRatio(cmbSelctCore - 1)
        
        For i = 1 To Len(ARatio) + 1
            value = Mid(txtRatio(cmbSelctCore - 1), i, i)
            If value = "-" Or i > Len(ARatio) Then
                 If CDbl(str1) > Gnumber Then
                      Gnumber = CDbl(str1)
                 End If
                 str1 = ""
            Else
                str1 = str1 & value
            End If
        
        Next
        
        GNo = Gnumber
        
        'Setting tap turing
        txtSecTap = ""
        For i = 1 To Len(ARatio) + 1
            value = Mid(txtRatio(cmbSelctCore - 1), i, i)
            If value = "-" Or i > Len(ARatio) Then
                 If CDbl(str1) > Gnumber Then
                      Gnumber = CDbl(str1)
                 Else
                    If CDbl(str1) <> Gnumber Then
                        If cmbStyle = "Rectangluar" Then
                            txtSecTap = txtSecTap & "-" & Round(txtTurns * CDbl(str1) / txtRatioR2(cmbSelctCore - 1), 2)
                        Else
                            txtSecTap = txtSecTap & "-" & Round(1 * CDbl(str1) / txtRatioR2(cmbSelctCore - 1), 2)
                        End If
                    End If
                 End If
                 str1 = ""
            Else
                str1 = str1 & value
            End If
            
        Next
        GNo = Gnumber
        txtSecTap = Mid(txtSecTap, 2)
        If cmbStyle = "Rectangluar" Then
            txtSecTurn = Round(txtTurns * Gnumber / txtRatioR2(cmbSelctCore - 1), 2)
        Else
            txtSecTurn = Round(1 * Gnumber / txtRatioR2(cmbSelctCore - 1), 2)
        End If
    End If
    txtCoreRatio = txtRatio(CInt(cmbSelctCore - 1)) & "/" & txtRatioR2(cmbSelctCore - 1) & " A"
End If
'Find C/s section of lowest current in primary side

For i = 1 To Len(txtRatio(cmbSelctCore - 1)) + 1
    value = Mid(txtRatio(cmbSelctCore - 1), i, i)
    If value = "-" Or i > Len(txtRatio(cmbSelctCore - 1)) Then
         If CDbl(str1) < mincurr Or mincurr = 0 Then
              mincurr = CDbl(str1)
         End If
         str1 = ""
    Else
        str1 = str1 & value
    End If
Next
'MsgBox Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1)
txtMinCsSec = prcCalSecC(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1), CDbl(mincurr), cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
txtSecCS = prcCalNorSecC(txtRatioR2(CInt(cmbSelctCore - 1)))
CurrValue = txtSecCS
If Len(txtRateConti) > 0 Then
    RateCs = prcCalRate(txtRateConti, CDbl(mincurr), txtRatioR2(CInt(cmbSelctCore - 1)))
    CurrValue = RateCs
End If

'Dim CurrValue As Double
Dim txtODs As Double
Dim txtODn As Double
    If CDbl(CurrValue) >= CDbl(txtMinCsSec) Then
        txtRateConti.Tag = CurrValue
        MinCross = CurrValue
        Set rst = New ADODB.Recordset
        rst.Open "select * from secconductor where t_cs >= " & CurrValue & " ", con, adOpenStatic, adLockReadOnly
        If rst.EOF = False And rst.BOF = False Then
            txtSwg = rst!t_swg
            txtNo = "S"
            txtODs = rst!t_ods
            txtODn = rst!t_cs
'            txtODn.Tag = rst!t_cs
        End If
        txtMinCsSec = txtSecCS
    Else
        txtRateConti.Tag = CurrValue
        MinCross = txtMinCsSec
        Set rst = New ADODB.Recordset
        rst.Open "select * from secconductor where t_cs >= " & txtMinCsSec & " ", con, adOpenStatic, adLockReadOnly
        SHHH = "select * from secconductor where t_cs >= " & txtMinCsSec & " "
        If rst.EOF = False And rst.BOF = False Then
            txtSwg = rst!t_swg
            txtNo = "S"
            txtODs = rst!t_ods
            txtODn = rst!t_cs
'            txtODn.Tag = rst!t_cs
        End If
'        txtSecCS = txtMinCsSec
    End If

    Set rst = Nothing


'bringing core ratio
txtCoreRatio = txtRatio(cmbSelctCore - 1) & "/" & txtRatioR2(cmbSelctCore - 1) & " A"
If Len(txtHt) > 0 Then
    txtHt_LostFocus
End If

''C/S finding for different primary if "-" present
'If InStr(1, txtCoreRatio, "-") = 0 Then
'    Chkoption.Enabled = False
'End If

'For different Dia
If InStr(1, txtRatio(cmbSelctCore - 1), "-") > 0 Then
    txtSwg.Visible = True
    txtNo.Visible = True
    Label67.Visible = True

    SWG.Visible = False
    SWG.Cols = 11
    SWG.Rows = 1
    SWG.TextMatrix(0, 0) = "Ratio"
    SWG.TextMatrix(0, 1) = "SWG"
    SWG.ColWidth(1) = 1500
    SWG.TextMatrix(0, 2) = "S/D/T/Q"
    SWG.ColWidth(2) = 1000
    SWG.TextMatrix(0, 3) = "VA"
    SWG.ColWidth(3) = 1000
    SWG.TextMatrix(0, 4) = "Class"
    SWG.ColWidth(4) = 1000
    SWG.TextMatrix(0, 5) = "ISF/ALF"
    SWG.ColWidth(5) = 1000
    SWG.TextMatrix(0, 6) = "VK"
    SWG.ColWidth(6) = 1000
    SWG.TextMatrix(0, 7) = "RCT"
    SWG.ColWidth(7) = 1000
    SWG.TextMatrix(0, 8) = "Ie"
    SWG.ColWidth(8) = 1000
    SWG.TextMatrix(0, 9) = "Ie @"
    SWG.ColWidth(9) = 1000
    SWG.TextMatrix(0, 10) = "Ie @"
    SWG.ColWidth(10) = 1000
'    MsgBox cmbCType
    For i = 1 To Len(txtRatio(cmbSelctCore - 1)) + 1
        value = Mid(txtRatio(cmbSelctCore - 1), i, i)

        If value = "-" Or i > Len(txtRatio(cmbSelctCore - 1)) Then

             SWG.Rows = SWG.Rows + 1
             SWG.row = SWG.Rows - 1

             If ChK2.value = 1 Then

                SWG.TextMatrix(SWG.row, 0) = str1 & "/" & txtRatioR2(cmbSelctCore - 1)
                SWG.TextMatrix(SWG.row, 1) = ""
                SWG.ColAlignment(SWG.row) = vbLeftJustify
                SWG.TextMatrix(SWG.row, 2) = ""
                For K = 1 To Grid.Rows - 1
                    If Grid.TextMatrix(K, 0) = cmbSelctCore And SWG.TextMatrix(SWG.row, 0) = Grid.TextMatrix(K, 1) And Len(Grid.TextMatrix(K, 16)) > 0 Then
                    '+++++++++++++++++++
                    'This has been changed from orginal given below
                    '+++++++++++++++++++
'                    If Grid.TextMatrix(K, 0) = cmbSelctCore And SWG.TextMatrix(SWG.row, 0) = Grid.TextMatrix(K, 1) And Len(Grid.TextMatrix(K, 16)) > 0 Then
                        SWG.TextMatrix(SWG.row, 10) = "DA"

                        txtMinCsSec = prcCalSecC(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), CDbl(str1), cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
                        txtSecCS = prcCalNorSecC(txtRatioR2(CInt(cmbSelctCore - 1)))

                        Set rst = New ADODB.Recordset
                        rst.Open "select * from secconductor where t_cs >= " & txtSecCS & " ", con, adOpenStatic, adLockReadOnly
                        If rst.EOF = False And rst.BOF = False Then
                            txtODs = rst!t_ods
                            txtODn = rst!t_cs
'                            txtOdn.Tag = rst!t_cs
                            txtSwg = rst!t_swg
                            txtNo = "S"
                        End If
                        Set rst = Nothing

                        SWG.TextMatrix(SWG.row, 3) = Grid.TextMatrix(K, 7)
                        SWG.TextMatrix(SWG.row, 4) = Grid.TextMatrix(K, 8)
                        If Grid.TextMatrix(K, 19) = "Metering" Then
                            SWG.TextMatrix(SWG.row, 5) = Grid.TextMatrix(K, 9)
                        ElseIf Grid.TextMatrix(K, 19) = "Protection" Then
                            SWG.TextMatrix(SWG.row, 5) = Grid.TextMatrix(K, 10)
                        End If
                        SWG.TextMatrix(SWG.row, 6) = Grid.TextMatrix(K, 11)
                        SWG.TextMatrix(SWG.row, 7) = Grid.TextMatrix(K, 12)
                        SWG.TextMatrix(SWG.row, 8) = Grid.TextMatrix(K, 13)
                        SWG.TextMatrix(SWG.row, 9) = Grid.TextMatrix(K, 14)

                    End If
                Next
             Else

                SWG.TextMatrix(SWG.row, 0) = str1 & "/" & txtRatioR2(cmbSelctCore - 1)
                txtMinCsSec = prcCalSecC(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), CDbl(str1), cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
                txtSecCS = prcCalNorSecC(txtRatioR2(CInt(cmbSelctCore - 1)))

                Set rst = New ADODB.Recordset
                rst.Open "select * from secconductor where t_cs >= " & txtSecCS & " ", con, adOpenStatic, adLockReadOnly
                If rst.EOF = False And rst.BOF = False Then
                    txtSwg = rst!t_swg
                    txtNo = "S"
                    txtODs = rst!t_ods
                    txtODn = rst!t_cs
'                    txtODn.Tag = rst!t_cs
                End If
                Set rst = Nothing

                SWG.TextMatrix(SWG.row, 1) = txtSwg
                SWG.TextMatrix(SWG.row, 2) = txtNo
                For K = 1 To Grid.Rows - 1
                    If Grid.TextMatrix(K, 0) = cmbSelctCore And SWG.TextMatrix(SWG.row, 0) = Grid.TextMatrix(K, 1) Then
                    '+++++++++++++
                    'This has been change from orginal given below
                    '+++++++++++++
'                    If Grid.TextMatrix(K, 0) = cmbSelctCore And SWG.TextMatrix(SWG.row, 0) = Grid.TextMatrix(K, 1) And Len(Grid.TextMatrix(K, 16)) > 0 Then
                        SWG.TextMatrix(SWG.row, 10) = "DA"
                        SWG.TextMatrix(SWG.row, 3) = Grid.TextMatrix(K, 7)
                        SWG.TextMatrix(SWG.row, 4) = Grid.TextMatrix(K, 8)
                        If Grid.TextMatrix(K, 19) = "Metering" Then
                            SWG.TextMatrix(SWG.row, 5) = Grid.TextMatrix(K, 9)
                        ElseIf Grid.TextMatrix(K, 19) = "Protection" Then
                            SWG.TextMatrix(SWG.row, 5) = Grid.TextMatrix(K, 10)
                        End If
                        SWG.TextMatrix(SWG.row, 6) = Grid.TextMatrix(K, 11)
                        SWG.TextMatrix(SWG.row, 7) = Grid.TextMatrix(K, 12)
                        SWG.TextMatrix(SWG.row, 8) = Grid.TextMatrix(K, 13)
                        SWG.TextMatrix(SWG.row, 9) = Grid.TextMatrix(K, 14)
                    End If
                Next
             End If
             str1 = ""
        Else
            str1 = str1 & value
        End If
    Next
Else
    SWG.Visible = False
End If
    If CDbl(txtSecCS) >= CDbl(txtMinCsSec) Then
        txtMinCsSec = txtSecCS
    End If
    
    

    
'Providing Customer specification of core
Dim Aff, HT As Double
If Grid.Rows > 1 Then
    For i = 1 To Grid.Rows - 1
        If Grid.TextMatrix(i, 0) = cmbSelctCore Then
            txtCoreRatio = txtRatio(cmbSelctCore - 1) & "/" & txtRatioR2(cmbSelctCore - 1) & " A"
            If Grid.TextMatrix(i, 19) = "Metering" Then
                 Frame19.Caption = "Metering"
                 Frame19.Visible = True
                 Frame20.Visible = False
                 Label70.Caption = "ISF <="
                 txtVA1 = Grid.TextMatrix(i, 7)
                 txtClass1 = Grid.TextMatrix(i, 8)
                 txtISF1 = Grid.TextMatrix(i, 9)
                 txtCoreType = "Metering"
                 firstr = Val(txtCoreRatio)
'                 If SetPTurnComm = "Y" Then
'                    Exit Sub
'                 End If
                 we = CrossSection(CDbl(firstr), txtClass1, MinCross, txtNo, 1)
                If InStr(1, txtCoreRatio, "-") > 0 Then
                    txtSecTap.Visible = True
                    Label63.Visible = True
                Else
                    txtSecTap.Visible = False
                    Label63.Visible = False
                End If
                count1 = 0
                Dim firstratio As Integer
                
                If InStr(1, txtCoreRatio, "-") > 0 Then
                    Rationlength = Left(txtCoreRatio, InStr(1, (txtCoreRatio), "/") - 1)
                    count1 = 0
                    For p = 1 To InStr(1, (txtCoreRatio), "/")
                        value = Mid(Rationlength, p, p)
                        If value = "-" Or p > Len(Rationlength) Then
                             If count1 > 0 Then
                                txtCompen(count1) = firstratio / CDbl(str1) * txtCompen(0)
                                txtCompen(count1).Visible = True
                             Else
                                firstratio = str1
                             End If
                             count1 = count1 + 1
                             str1 = ""
                        Else
                            str1 = str1 & value
                        End If
                    Next
                End If
                If PrcCheck_Core_For_Mod(cmbSelctCore) = False Then
                    Exit Sub
                End If
                InCalcu = ""
                Exit Sub
'                 txtISF.Visible = True
                 Label85.Visible = True
                 txtECatAlf.Visible = False
                 Label87.Visible = False
                 txtECatN.Visible = False
                 Label86.Visible = False
                 txtIeNor.Visible = False
                 Label91.Visible = False
                 txtIeALF.Visible = False
                 Label92.Visible = False
                 txtBM50.Visible = True
                 txtBM60.Visible = True
                 txtC1.Visible = False
                 Label81.Visible = False
                 txtC2.Visible = False
                 Label82.Visible = False
                 txtECurr.Visible = False
                 Label83.Visible = False
                 
                 ErrorRead.Visible = True
                 ErrorRead1.Visible = True
                 
                 ErrorRead.Rows = 1
                 ErrorRead1.Rows = 1
                 
                 Exit Sub
            ElseIf Grid.TextMatrix(i, 19) = "Protection" Then
                 Frame19.Caption = "Protection"
                 Frame19.Visible = True
                 Frame20.Visible = False
                 Label70.Caption = "ALF "
                 txtVA1 = Grid.TextMatrix(i, 7)
                 txtClass1 = Grid.TextMatrix(i, 8)
                 txtISF1 = Grid.TextMatrix(i, 10)
                 txtCoreType = "Protection"
                                 
                If InStr(1, txtCoreRatio, "-") > 0 Then
                    txtSecTap.Visible = True
                    Label63.Visible = True
                Else
                    txtSecTap.Visible = False
                    Label63.Visible = False
                End If
'                 Exit Sub
                 'Temp
                 If Len(txtTurns) <> 0 Then
'
'                 'To calculate height protection
                 Dim BMMM As Double
                 Dim BMPass As comboPrc
                 
                 If PrcCheck_Core_For_Mod(cmbSelctCore) = False Then
                    Exit Sub
                 End If
                 
'                 Dim txtBM50 As Double
                For p = 2 To 100
                        
                    HT = p * 5
                    txtHt = HT
                    txtHt_LostFocus
                    Dim FF As Integer
                    If cmbStyle = "Circular" Then
                        FF = txtID1
                    Else
                        FF = txtWW1
                    End If
                    Set BMPass = New comboPrc
                    txtBM50 = BMPass.CalProtectingHt(txtBUP, txtEArea, cmbCType, txtMean, txtRatioR2(cmbSelctCore - 1), Val(txtRatio(cmbSelctCore - 1)), txtSecTurn, txtODn, txtODs, txtHt, FF, txtVA1)
                    
'                    txtNo_Change
                    If cmbFreq = 50 Then
                        If Val(txtISF1) = 10 Then
                            If CDbl(txtBM50) >= 1600 And CDbl(txtBM50) <= 1700 Then
                                Exit Sub
                            'OUT OF RANGE
                            ElseIf CDbl(txtBM50) < 1600 Then
                                txtHt = CDbl(txtHt) - 5
                                Exit Sub
                            End If
                        ElseIf Val(txtISF1) = 15 Then
                            If CDbl(txtBM50) >= 1000 And CDbl(txtBM50 <= 1100) Then
                                Exit Sub
                            'OUT OF RANGE
                            ElseIf CDbl(txtBM50) < 1000 Then
                                txtHt = CDbl(txtHt) - 5
                                Exit Sub
                            End If
                        ElseIf Val(txtISF1) = 20 Then
                            If CDbl(txtBM50) >= 800 And CDbl(txtBM50) <= 900 Then
                                Exit Sub
                            'OUT OF RANGE
                            ElseIf CDbl(txtBM50) < 800 Then
                                txtHt = CDbl(txtHt) - 5
                                Exit Sub
                            End If
                            
                        End If
                    ElseIf Val(cmbFreq) = 60 Then
                        If txtISF1 = 10 Then
                            If CDbl(txtBM50) >= 1600 And CDbl(txtBM50) <= 1700 Then
                                Exit Sub
                            'OUT OF RANGE
                            ElseIf CDbl(txtBM50) < 1600 Then
                                txtHt = CDbl(txtHt) - 5
                                Exit Sub
                            End If
                        ElseIf Val(txtISF1) = 15 Then
                            If CDbl(txtBM50) >= 1000 And CDbl(txtBM50 <= 1100) Then
                                Exit Sub
                            'OUT OF RANGE
                            ElseIf CDbl(txtBM50) < 1000 Then
                                txtHt = CDbl(txtHt) - 5
                                Exit Sub
                            End If
                        ElseIf Val(txtISF1) = 20 Then
                            If CDbl(txtBM50) >= 800 And CDbl(txtBM50) <= 900 Then
                                Exit Sub
                            'OUT OF RANGE
                            ElseIf CDbl(txtBM50) < 800 Then
                                txtHt = CDbl(txtHt) - 5
                                Exit Sub
                            End If
                            
                        End If
                    End If
                    
                 Next
                 End If
                 Exit Sub
            ElseIf Grid.TextMatrix(i, 19) = "Special Protection" Then
                'Dim mincurr As Double
                For j = 1 To Len(txtRatio(cmbSelctCore - 1)) + 1
                value = Mid(txtRatio(cmbSelctCore - 1), j, j)
                    If value = "-" Or j > Len(txtRatio(cmbSelctCore - 1)) Then
                         If CDbl(str1) < mincurr Or mincurr = 0 Then
                              mincurr = CDbl(str1)
                         End If
                         str1 = ""
                    Else
                        str1 = str1 & value
                    End If
                Next
                 Frame20.Caption = "Special Protection"
                 Frame19.Visible = False
                 Frame20.Visible = True
                 txtVK1 = Grid.TextMatrix(i, 11)
                 txtRCT1 = Grid.TextMatrix(i, 12)
                 txtIE1 = Grid.TextMatrix(i, 13)
                 txtV1 = Grid.TextMatrix(i, 14)
                 txtCoreType = "Special Protection"
                 
                If InStr(1, txtCoreRatio, "-") > 0 Then
                    txtSecTap.Visible = True
                    Label63.Visible = True
                Else
                    txtSecTap.Visible = False
                    Label63.Visible = False
                End If
'                 Exit Sub
'                 txtISF.Visible = False '1
'                 Label85.Visible = False '1
'                 txtECatAlf.Visible = False '1
'                 Label87.Visible = False '1
'                 txtECatN.Visible = False '1
'                 Label86.Visible = False '1
'                 txtIeNor.Visible = False '1
'                 Label91.Visible = False '1
'                 txtIeALF.Visible = False '1
'                 Label92.Visible = False '1
'                 txtBM50.Visible = False '1
'                 txtBM60.Visible = False '1
'                 txtECurr.Visible = True
'                 Label83.Visible = True
'                 txtC1.Visible = True
'                 Label81.Visible = True
'                 txtC2.Visible = True '
'                 Label82.Visible = True
'
'                 ErrorRead.Visible = False
'                 ErrorRead1.Visible = False
'
'                 ErrorRead.Rows = 1
'                 ErrorRead1.Rows = 1
'
                 If PrcCheck_Core_For_Mod(cmbSelctCore) = False Then
                    Exit Sub
                 End If
                 
                 
                     If Len(txtTurns) > 0 Then
                         If cmbFreq = "50" Then
                            'MsgBox ((MinCurr * txtTurns) / txtRatioR2(cmbSelctCore - 1))
                            Aff = (1.1 * txtVK1 * 3000) / ((mincurr * txtTurns) / txtRatioR2(cmbSelctCore - 1) * 96)
                            txtHt = Aff / (txtBUP * 0.95) * 100
                         ElseIf cmbFreq = "60" Then
                            Aff = (1.1 * txtVK1 * 2500) / ((mincurr * txtTurns) / txtRatioR2(cmbSelctCore - 1) * 96)
                            txtHt = Aff / (txtBUP * 0.95) * 100
                         End If
                     
'                     txtHt = 78.9
                         txtHt = Round(txtHt, 1)
    '                     MsgBox Right(txtHt, 3)
                         If InStr(1, txtHt, ".") > 0 Then
                            If Right(txtHt, 3) <= 2.5 And Right(txtHt, 3) > 0 Then
                               If InStr(1, txtHt, ".") = 0 Then
                                   txtHt = Left(txtHt, Len(txtHt) - 1) & 0
                               Else
                                   txtHt = Left(txtHt, Len(txtHt) - 3) & 0
                               End If
                            ElseIf Right(txtHt, 3) <= 5 And Right(txtHt, 3) > 2.5 Then
                               If InStr(1, txtHt, ".") = 0 Then
                                   txtHt = Left(txtHt, Len(txtHt) - 1) & 5
                               Else
                                   txtHt = Left(txtHt, Len(txtHt) - 3) & 5
                               End If
                            ElseIf Right(txtHt, 3) <= 7.5 And Right(txtHt, 3) > 5 Then
                               If InStr(1, txtHt, ".") = 0 Then
                                   txtHt = Left(txtHt, Len(txtHt) - 1) & 5
                               Else
                                   txtHt = Left(txtHt, Len(txtHt) - 3) & 5
                               End If
                            ElseIf Right(txtHt, 3) = 0 Or Right(txtHt, 3) > 7.5 Then
                               If InStr(1, txtHt, ".") = 0 Then
                                   txtHt = Left(txtHt, Len(txtHt) - 1) & 0
                               Else
                                   txtHt = Left(txtHt, Len(txtHt) - 3) & 0
                                   txtHt = txtHt + 10
                               End If
                            End If
                        Else
                            If Right(txtHt, 1) <= 2.5 And Right(txtHt, 1) > 0 Then
                               If InStr(1, txtHt, ".") = 0 Then
                                   txtHt = Left(txtHt, Len(txtHt) - 1) & 0
                               Else
                                   txtHt = Left(txtHt, Len(txtHt) - 3) & 0
                               End If
                            ElseIf Right(txtHt, 1) <= 5 And Right(txtHt, 1) > 2.5 Then
                               If InStr(1, txtHt, ".") = 0 Then
                                   txtHt = Left(txtHt, Len(txtHt) - 1) & 5
                               Else
                                   txtHt = Left(txtHt, Len(txtHt) - 3) & 5
                               End If
                            ElseIf Right(txtHt, 1) <= 7.5 And Right(txtHt, 1) > 5 Then
                               If InStr(1, txtHt, ".") = 0 Then
                                   txtHt = Left(txtHt, Len(txtHt) - 1) & 5
                               Else
                                   txtHt = Left(txtHt, Len(txtHt) - 3) & 5
                               End If
                            ElseIf Right(txtHt, 3) = 0 Or Right(txtHt, 3) > 7.5 Then
                               If InStr(1, txtHt, ".") = 0 Then
                                   txtHt = Left(txtHt, Len(txtHt) - 1) & 0
                               Else
                                   txtHt = Left(txtHt, Len(txtHt) - 3) & 0
                                   txtHt = txtHt + 10
                               End If
                            End If
                        End If
                    End If
                 Exit Sub
            End If
        End If
    Next
End If

we = CrossSection(txtSecTurn, txtClass1, MinCross, txtNo, 1)

err:
    
    If Len(Trim(err.Description)) Then
        MsgBox err.Description & vbCrLf & err.Number, vbExclamation
    End If
    
End Sub

Private Sub cmbSelctCore_Click()

'Exit Sub
On Error GoTo err

If Val(cmbSelctCore) = 0 Or Val(cmbSelctCore) > GCore.Rows - 1 Then
'    MsgBox "Please enter proper core", vbInformation
    cmbSelctCore.SetFocus
    Exit Sub
End If

If cmbSelctCore <> 0 Or Len(cmbSelctCore) <> 0 Then

    If Trim(cmbStyle) <> "Circular" Then
        txtOL1 = GCore.TextMatrix(cmbSelctCore, 1)
        txtOW1 = GCore.TextMatrix(cmbSelctCore, 2)
        txtWL1 = GCore.TextMatrix(cmbSelctCore, 3)
        txtWW1 = GCore.TextMatrix(cmbSelctCore, 4)

        Frame12.Visible = True
        Frame18.Visible = False
    Else
        txtOD1 = GCore.TextMatrix(cmbSelctCore, 8)
        txtID1 = GCore.TextMatrix(cmbSelctCore, 9)

        Frame12.Visible = False
        Frame18.Visible = True
    End If
    
    txtBUP = GCore.TextMatrix(cmbSelctCore, 11)
    txtMean = GCore.TextMatrix(cmbSelctCore, 10)
'            GCore.TextMatrix(GCore.Row, 1) = txtOL
'            GCore.TextMatrix(GCore.Row, 2) = txtOW
'            GCore.TextMatrix(GCore.Row, 3) = txtWL
'            GCore.TextMatrix(GCore.Row, 4) = txtWW
End If

If Len(txtHt) > 0 Then
    txtHt_LostFocus
End If
    Dim value As String * 1
    Dim str1 As String
    Dim mincurr As Double
    For i = 1 To Len(txtRatio(cmbSelctCore - 1)) + 1
        value = Mid(txtRatio(cmbSelctCore - 1), i, i)
        If value = "-" Or i > Len(txtRatio(cmbSelctCore - 1)) Then
            If Len(str1) = 0 Then
                str1 = 0
            End If
             If CDbl(str1) < mincurr Or mincurr = 0 Then
                  mincurr = CDbl(str1)
             End If
             str1 = ""
        Else
        str1 = str1 & value
    End If
    Next
'            MsgBox Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1)
    txtMinCsSec = prcCalSecC(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1), mincurr, cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))

If Grid.Rows > 1 Then
    For i = 1 To Grid.Rows - 1
        If Grid.TextMatrix(i, 0) = cmbSelctCore Then
            txtCoreRatio = txtRatio(cmbSelctCore - 1) & "/" & txtRatioR2(cmbSelctCore - 1) & " A"
            If Grid.TextMatrix(CInt(cmbSelctCore), 19) = "Metering" Then
                 Frame19.Caption = "Metering"
                 Frame19.Visible = True
                 Frame20.Visible = False
                 Label70.Caption = "ISF <="
                 txtVA1 = Grid.TextMatrix(i, 7)
                 txtClass1 = Grid.TextMatrix(i, 8)
                 txtISF1 = Grid.TextMatrix(i, 9)
                 txtCoreType = "Metering"
                 Exit Sub
            ElseIf Grid.TextMatrix(CInt(cmbSelctCore), 19) = "Protection" Then
                 Frame19.Caption = "Protection"
                 Frame19.Visible = True
                 Frame20.Visible = False
                 Label70.Caption = "ALF "
                 txtVA1 = Grid.TextMatrix(i, 7)
                 txtClass1 = Grid.TextMatrix(i, 8)
                 txtISF1 = Grid.TextMatrix(i, 10)
                 txtCoreType = "Protection"
                 Exit Sub
            ElseIf Grid.TextMatrix(CInt(cmbSelctCore), 19) = "Special Protection" Then
                 Frame20.Caption = "Special Protection"
                 Frame19.Visible = False
                 Frame20.Visible = True
                 txtVK1 = Grid.TextMatrix(i, 11)
                 txtRCT1 = Grid.TextMatrix(i, 12)
                 txtIE1 = Grid.TextMatrix(i, 13)
                 txtV1 = Grid.TextMatrix(i, 14)
                 txtCoreType = "Special Protection"
                 Exit Sub
            End If
        End If
    Next
End If

Exit Sub
err:
    
    MsgBox err.Description, vbCritical

End Sub

Private Sub cmbSelctCore_GotFocus()
cmbSelctCore.SelStart = 0
cmbSelctCore.SelLength = Len(cmbSelctCore)
MDIDesign.StatusBar1.Panels(2).Text = "Core Select"
End Sub

Private Sub cmbSelctCore_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub cmbSelctCore_LostFocus()
Exit Sub
''Check all core calculations
Dim Sta As Integer

For i = 1 To Val(txtnoofcores)
    Sta = 0
    For j = 1 To G1.Cols - 1
        If i = G1.TextMatrix(1, j) Then
            Sta = 1
        End If
        If Sta = 1 Then
            Exit For
        End If
    Next
    If Sta = 0 Then
        Exit For
    End If
Next

If Sta = 0 Then
    cmbSelctCore = i
    Exit Sub
    If Val(cmbSelctCore) = 0 Or Val(cmbSelctCore) > GCore.Rows - 1 Then
        MsgBox "Please enter proper core", vbInformation
        cmbSelctCore.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub cmbSTC_Change()
If Len(cmbSTC) > 0 And UCase(Trim(cmbSTC)) <> "KA" Then
    If Grid.Rows > 1 Then
        ChangeReadings
    End If
    If InStr(1, UCase(cmbSTC), "K") > 0 Then
        txtIDyn = CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1)) * 2.5
    Else
        txtIDyn = cmbSTC * 2.5
    End If
End If
End Sub

Private Sub cmbSTC_Click()
Dim X As String
If Len(cmbSTC) > 0 Then
    If Grid.Rows > 1 Then
        ChangeReadings
    End If
    txtIDyn = CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1)) * 2.5
End If
'x = cmbSTC
'cmbSTC.ListIndex = -1
'cmbSTC = x
End Sub

Private Sub cmbSTC_GotFocus()
cmbSTC.SelStart = 0
cmbSTC.SelLength = Len(cmbSTC)
End Sub

Private Sub cmbSTC_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9.]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub cmbSTC_LostFocus()
If InStr(1, UCase(cmbSTC), "K") = 0 Then
    cmbSTC = cmbSTC & " kA"
End If

If cmbSTC = "kA" Then
    MsgBox "Please enter STC", vbExclamation
    cmbSTC.SetFocus
End If

End Sub

Private Sub cmbStyle_Change()

If Trim(cmbStyle.Text) = "Rectangluar" Then
    Frame10.Visible = True
    Frame11.Visible = False
    Frame12.Visible = True
    Frame18.Visible = False
    Frame9.Enabled = True
    GCore.TextMatrix(0, 8) = "Over All"
    GCore.TextMatrix(0, 9) = "Window"
Else
    Frame10.Visible = False
    Frame11.Visible = True
    Frame12.Visible = False
    Frame18.Visible = True
    Frame9.Enabled = False
    GCore.TextMatrix(0, 8) = "OD"
    GCore.TextMatrix(0, 9) = "ID"
    Grid3.Rows = 1
End If



End Sub

Private Sub cmbStyle_Click()

If GCore.Rows > 1 Then
    MsgBox "Core Dimensions will be changed", vbExclamation
    GCore.Rows = 1
    G1.Cols = 1
ElseIf G1.Cols > 1 Then
    MsgBox "Core Dimensions will be changed", vbExclamation
    GCore.Rows = 1
    G1.Cols = 1

End If

If Trim(cmbStyle.Text) = "Rectangluar" Then

    Frame10.Visible = True
    Frame11.Visible = False
    
    Frame9.Enabled = True
    GCore.TextMatrix(0, 8) = "Over All"
    GCore.TextMatrix(0, 9) = "Window"
Else

    Frame10.Visible = False
    Frame11.Visible = True
    
    Frame9.Enabled = False
    GCore.TextMatrix(0, 8) = "OD"
    GCore.TextMatrix(0, 9) = "ID"

End If

End Sub

'Private Sub cmdCoreIn_Click()
'
'Dim value As String * 1
'Dim CorNo As Integer
'
''If Len(txtRatio) = 0 Or Len(txtRatioR2) = 0 Then
''
''    MsgBox "Ratio not set properly", vbInformation
''    txtRatio.SetFocus
''    Exit Sub
''
''End If
'
'GridRatio.Rows = GridRatio.Rows + 1
'GridRatio.row = GridRatio.Rows - 1
''Label6.Caption = "Core " & GridRatio.row + 1 & " Ratio :"
'GridRatio.RowHeight(GridRatio.row) = 300
'Dim intNo As Integer
''Checking the core no
'For i = 1 To GridRatio.Rows - 1
'
'    If Len(GridRatio.TextMatrix(i, 2)) > 0 Then
'
'        If InStr(1, GridRatio.TextMatrix(i, 2), "-") > 0 Then
'
'            For j = 1 To Len(GridRatio.TextMatrix(i, 2)) + 1
'                value = Mid(GridRatio.TextMatrix(i, 2), j, j)
'                If value = "-" Or j > Len(GridRatio.TextMatrix(i, 2)) Then
'                    CorNo = CorNo + 1
'                End If
'            Next
'            'if last value is - then is not counted
'            If Mid(GridRatio.TextMatrix(i, 2), Len(GridRatio.TextMatrix(i, 2)), Len(GridRatio.TextMatrix(i, 2))) = "-" Then
'                CorNo = CorNo - 1
'            End If
'
'            intNo = intNo + CorNo
'        Else
'
'            intNo = intNo + 1
'
'        End If
'
'    End If
'    CorNo = 0
'Next
'
'    For K = 1 To Len(txtRatioR2) + 1
'        value = Mid(txtRatioR2, K, K)
'        If value = "-" Or K > Len(txtRatioR2) Then
'            cor1 = cor1 + 1
'        End If
'    Next
'    'if last value is - then is not counted
'    If Mid(txtRatioR2, Len(txtRatioR2), Len(txtRatioR2)) = "-" Then
'        cor1 = cor1 - 1
'    End If
'
'    intNo = intNo + cor1
'
'GridRatio.TextMatrix(GridRatio.row, 0) = intNo
'GridRatio.TextMatrix(GridRatio.row, 1) = Trim(txtRatio)
'GridRatio.TextMatrix(GridRatio.row, 2) = Trim(txtRatioR2)
'
'txtRatio = ""
'txtRatioR2 = ""
'
'Frame2.Enabled = True
'Frame4.Enabled = True
'Frame6.Enabled = True
'
'End Sub

'Private Sub cmdCoreOut_Click()
'
'If GridRatio.Rows = 1 Then
'    Exit Sub
'End If
'
'If GridRatio.Rows = 2 Then
'    GridRatio.Rows = 1
'ElseIf GridRatio.Rows > 2 Then
'    GridRatio.RemoveItem (GridRatio.row)
'End If
'Dim value As String * 1
'Dim intNo As Integer
''Checking the core no
'For i = 1 To GridRatio.Rows - 1
'
'    If Len(GridRatio.TextMatrix(i, 2)) > 0 Then
'
'        If InStr(1, GridRatio.TextMatrix(i, 2), "-") > 0 Then
'
'            For j = 1 To Len(GridRatio.TextMatrix(i, 2)) + 1
'                value = Mid(GridRatio.TextMatrix(i, 2), j, j)
'                If value = "-" Or j > Len(GridRatio.TextMatrix(i, 2)) Then
'                    CorNo = CorNo + 1
'                End If
'            Next
'            'if last value is - then is not counted
'            If Mid(GridRatio.TextMatrix(i, 2), Len(GridRatio.TextMatrix(i, 2)), Len(GridRatio.TextMatrix(i, 2))) = "-" Then
'                CorNo = CorNo - 1
'            End If
'
'            intNo = intNo + CorNo
'        Else
'
'            intNo = intNo + 1
'
'        End If
'
'    End If
'    CorNo = 0
'    GridRatio.TextMatrix(i, 0) = intNo
'Next
'
'If GridRatio.Rows = 1 Then
'    Frame2.Enabled = False
'    Frame4.Enabled = False
'    Frame6.Enabled = False
'End If
'
'End Sub


Private Sub cmbStyle_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Set Core Style"
End Sub

Private Sub cmbTime_Change()
    If Grid.Rows > 1 Then
        ChangeReadings
    End If
End Sub

Private Sub cmbTime_Click()
    If Grid.Rows > 1 Then
        ChangeReadings
    End If
End Sub

Private Sub cmbTime_GotFocus()
cmbTime.SelStart = 0
cmbTime.SelLength = Len(cmbTime)
MDIDesign.StatusBar1.Panels(2).Text = "Select STC Tims (sec.)"
End Sub

Private Sub cmbTime_KeyPress(KeyAscii As Integer)
On Error GoTo err
If Len(Trim(cmbTime)) = 0 Then
    Exit Sub
End If
If cmbTime > 10 Then
    KeyAscii = 0
End If
If Not Chr(KeyAscii) Like "[0-9.]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If

End Sub

Private Sub cmbType_GotFocus()
cmbType.SelStart = 0
cmbType.SelLength = Len(cmbType)
MDIDesign.StatusBar1.Panels(2).Text = "Drawing Type"
End Sub

Private Function SwgCheck() As Boolean


''SWG size checking in database

Set rst = New Recordset
rst.Open "select * FROM SecConductor where t_swg >= " & txtSwg & " and  t_swg <= " & txtSwg & "", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
    SwgCheck = False
Else
    MsgBox txtSwg & " does not exits in the database.", vbCritical
    txtSwg.SetFocus
    SwgCheck = True
End If
Set rst = Nothing

End Function

Private Sub cmCal_Click()

On Error GoTo err

If Len(txtHt) = 0 Then
    MsgBox "Please enter height of core.", vbExclamation
    txtHt.SetFocus
    Exit Sub
End If
'Precausion for ht calculation eg. Aeff
txtHt_LostFocus

If SwgCheck Then
    Exit Sub
End If

If Len(txtSwg) = 0 Then
    MsgBox "Please enter SWG for core " + cmbSelctCore, vbExclamation
    txtSwg.SetFocus
    Exit Sub
End If
If Len(txtNo) = 0 Then
    MsgBox "Please enter S&ingle / D&ouble / T&riple / Quadrable. for core " + cmbSelctCore, vbExclamation
    txtNo.SetFocus
    Exit Sub
End If
If Len(cmbMetal) = 0 Then
    MsgBox "Please select core material.", vbExclamation
    cmbMetal.SetFocus
    Exit Sub
End If
If Val(txtTurns) = 0 Then
    MsgBox "Please Primary turns", vbExclamation
    SSTab1.Tab = 2
    txtTurns.SetFocus
    Exit Sub
End If

Individualfrm.Show
'txtRatio(cmbSelctCore - 1) = txtRatio(cmbSelctCore - 1) & "-"
'If InStr(1, txtRatio(cmbSelctCore - 1), "-") > 0 Then
prcFillInd
'    MsgBox err.Description
'End If

'======================
'Else
'
'    prcAddRem
'
'    txtNo_Change
'    Individualfrm.txtMinCsSec(0) = txtMinCsSec
'    Individualfrm.txtODn(0) = txtODn
'    Individualfrm.txtSwg(0) = txtSwg
'    Individualfrm.txtNo(0) = txtNo
'    Individualfrm.txtLayers(0) = txtLayers
'    Individualfrm.txtWHt(0) = txtWHt
'    If Len(txtBM50) = 0 Then
'        Individualfrm.txtBM50(0).Visible = False
'        Individualfrm.txtBM50(0).Visible = True
'    Else
'        Individualfrm.txtBM50(0).Visible = True
'        Individualfrm.txtBM50(0).Visible = False
'    End If
'    Individualfrm.txtBM50(0) = txtBM50
'    Individualfrm.txtBM60(0) = txtBM60
'    Individualfrm.txtRes(0) = txtRes
'    Individualfrm.txtRec75(0) = txtRec75
'    Individualfrm.txtCWt(0) = txtCWt
'    Individualfrm.txtVkCal(0) = txtVkCal
'    Individualfrm.txtC1(0) = txtC1
'    Individualfrm.txtMinCsSec(0) = txtC2
'    Individualfrm.txtDelta(0) = txtDelta
'    Individualfrm.txtISF(0) = txtISF
'    Individualfrm.txtECurr(0) = txtECurr
'    Individualfrm.txtECatN(0) = txtECatN
'    Individualfrm.txtECatAlf(0) = txtECatAlf
'    Individualfrm.txtSat(0) = txtSat
'    Individualfrm.txtSCTempR(0) = txtSCTempR
'
'
'    Individualfrm.ErrorRead(0).Cols = 17
'    For l = 0 To ErrorRead.Rows - 1
'        Individualfrm.ErrorRead(0).Rows = Individualfrm.ErrorRead(0).Rows + 1
'        Individualfrm.ErrorRead(0).Row = Individualfrm.ErrorRead(0).Rows - 1
'        For j = 0 To ErrorRead.Cols - 1
'            Individualfrm.ErrorRead(0).TextMatrix(l, j) = ErrorRead.TextMatrix(l, j)
'        Next
'    Next
'    Individualfrm.ErrorRead1(0).Cols = 17
'    For l = 0 To ErrorRead1.Rows - 1
'        Individualfrm.ErrorRead1(0).Rows = Individualfrm.ErrorRead1(0).Rows + 1
'        Individualfrm.ErrorRead1(0).Row = Individualfrm.ErrorRead1(0).Rows - 1
'        For j = 0 To ErrorRead1.Cols - 1
'            Individualfrm.ErrorRead1(0).TextMatrix(l, j) = ErrorRead1.TextMatrix(l, j)
'        Next
'    Next
'End If

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & " " & err.Number
    End If

End Sub

Private Sub cmCal_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Start Core Calculation"
End Sub

Private Sub cmdcancel_Click()

joel = MsgBox("Are you sure you want to cancel this design?", vbInformation + vbYesNo)
If vbYes = joel Then

    CancelCT = "1"
        
    Unload Me
    Set CTfrm = Nothing
    
    CTfrm.Show
    cmbEquipType.SetFocus
   

End If
End Sub

Private Sub cmdICore_Click()

On Error Resume Next

If GCore.Rows - 1 = CInt(txtnoofcores) Then
    MsgBox "All " & txtnoofcores & " core dimension added!", vbExclamation
    Exit Sub
End If

If Len(cmbCore) = 0 Then
    MsgBox "Please select core no", vbInformation
    Exit Sub
End If

cmbCore.ListIndex = -1

If Trim(cmbStyle) = "Circular" Then
    If Len(txtID) = 0 Or Len(txtOD) = 0 Then
        MsgBox "Please enter proper core size", vbInformation
        Exit Sub
    End If
Else
    If Len(txtOL) = 0 Or Len(txtOW) = 0 Or Len(txtWL) = 0 Or Len(txtWW) = 0 Then
        MsgBox "Please enter proper core size", vbInformation
        Exit Sub
    End If
End If

For j = 1 To GCore.Rows - 1
    If Trim(cmbCore) = Trim(GCore.TextMatrix(j, 0)) Then
        MsgBox "This core already exits", vbInformation
        Exit Sub
    End If
Next

Dim Cout As String
Dim CNo As Integer
CNo = 1
For K = 1 To txtnoofcores
    Cout = ""
    For j = 1 To GCore.Rows - 1
        If GCore.TextMatrix(j, 0) = K Then
            Cout = "K"
            Exit For
        End If
    Next
    If Cout <> "K" Then
        Exit For
    End If
Next

If Cout <> "K" Then
    CNo = K
Else
    CNo = GCore.Rows
End If

If cmbStyle.Text = "Rectangluar" Then
    If OptCom(0).value = True Then
        For i = 1 To CInt(txtnoofcores)
            GCore.Rows = GCore.Rows + 1
            GCore.row = GCore.Rows - 1
            GCore.RowHeight(GCore.row) = 300
            GCore.TextMatrix(GCore.row, 0) = i
            GCore.TextMatrix(GCore.row, 1) = txtOL
            GCore.TextMatrix(GCore.row, 2) = txtOW
            GCore.TextMatrix(GCore.row, 3) = txtWL
            GCore.TextMatrix(GCore.row, 4) = txtWW
            GCore.TextMatrix(GCore.row, 8) = txtOL & " x " & txtOW
            GCore.TextMatrix(GCore.row, 9) = txtWL & " x " & txtWW
            GCore.TextMatrix(GCore.row, 10) = txtMean
            GCore.TextMatrix(GCore.row, 11) = txtBUP
       Next
    Else
        GCore.Rows = GCore.Rows + 1
        GCore.row = GCore.Rows - 1
        GCore.RowHeight(GCore.row) = 300
        GCore.TextMatrix(GCore.row, 0) = CNo
        GCore.TextMatrix(GCore.row, 1) = txtOL
        GCore.TextMatrix(GCore.row, 2) = txtOW
        GCore.TextMatrix(GCore.row, 3) = txtWL
        GCore.TextMatrix(GCore.row, 4) = txtWW
        GCore.TextMatrix(GCore.row, 8) = txtOL & " x " & txtOW
        GCore.TextMatrix(GCore.row, 9) = txtWL & " x " & txtWW
        GCore.TextMatrix(GCore.row, 10) = txtMean
        GCore.TextMatrix(GCore.row, 11) = txtBUP
    End If
Else
    If OptCom(0).value = True Then
        For i = 1 To CInt(txtnoofcores)
            GCore.Rows = GCore.Rows + 1
            GCore.row = GCore.Rows - 1
            GCore.RowHeight(GCore.row) = 300
            GCore.TextMatrix(GCore.row, 0) = GCore.row
            GCore.TextMatrix(i, 5) = txtOD
            GCore.TextMatrix(i, 6) = txtID
            GCore.TextMatrix(i, 8) = txtOD
            GCore.TextMatrix(i, 9) = txtID
            GCore.TextMatrix(i, 10) = txtMean
            GCore.TextMatrix(i, 11) = txtBUP
        Next
    Else
            GCore.Rows = GCore.Rows + 1
            GCore.row = GCore.Rows - 1
            GCore.RowHeight(GCore.row) = 300
            GCore.TextMatrix(GCore.row, 0) = CNo
            GCore.TextMatrix(GCore.row, 5) = txtOD
            GCore.TextMatrix(GCore.row, 6) = txtID
            GCore.TextMatrix(GCore.row, 8) = txtOD
            GCore.TextMatrix(GCore.row, 9) = txtID
            GCore.TextMatrix(GCore.row, 10) = txtMean
            GCore.TextMatrix(GCore.row, 11) = txtBUP
    End If
End If
        If cmbCore.ListIndex < cmbCore.ListCount - 1 Then
            cmbCore.ListIndex = cmbCore.ListIndex + 1
        End If
        txtHt = ""
        txtBUP = ""
        txtArea = ""
        txtEArea = ""
        txtMean = ""
        txtWL = ""
        txtBUP = ""
        txtOL = ""
        txtOW = ""
        txtWL = ""
        txtWW = ""
        txtOD = ""
        txtID = ""
        Frame1.Enabled = False
        
CNo = 1
For K = 1 To txtnoofcores
    Cout = ""
    For j = 1 To GCore.Rows - 1
        If GCore.TextMatrix(j, 0) = K Then
            Cout = "K"
            Exit For
        End If
    Next
    If Cout <> "K" Then
        Exit For
    End If
Next
        
If Cout = "" Then
    cmbCore = "Core " & K
End If
txtBUP = ""
txtMean = ""

If OptCom(1).value = True Then
    If GCore.Rows < Val(txtnoofcores) Then
        cmbCore.Text = "Core " & GCore.Rows
    ElseIf GCore.Rows < Val(txtnoofcores) Then
        cmbCore.Text = ""
    End If
End If

If txtOL.Visible = True Then
    txtOL.SetFocus
Else
    txtOD.SetFocus
End If

End Sub

Private Function prcCheckIn() As Boolean

If UCase(Trim(cmbCType)) = "METERING" Then

If Len(cmbVA) = 0 Then
    MsgBox "Please select VA", vbExclamation
    prcCheckIn = True
    Exit Function
End If
If Len(cmbClass) = 0 Then
    MsgBox "Please select class", vbExclamation
    cmbClass.SetFocus
    prcCheckIn = True
    Exit Function
End If

ElseIf UCase(Trim(cmbCType)) = "PROTECTION" Then

If Len(cmbVA) = 0 Then
    MsgBox "Please select VA", vbExclamation
    prcCheckIn = True
    Exit Function
End If
If Len(cmbClass) = 0 Then
    MsgBox "Please select Class", vbExclamation
    cmbClass.SetFocus
    prcCheckIn = True
    Exit Function
End If
If Len(cmbISF) = 0 Then
    MsgBox "Please select ASF", vbExclamation
    cmbISF.SetFocus
    prcCheckIn = True
    Exit Function
End If
ElseIf UCase(Trim(cmbCType)) = "SPECIAL PROTECTION" Then
If Len(txtVK) = 0 Then
    MsgBox "Please select VK", vbExclamation
    txtVK.SetFocus
    prcCheckIn = True
    Exit Function
End If
If Len(txtRct) = 0 Then
    MsgBox "Please select rct", vbExclamation
    txtRct.SetFocus
    prcCheckIn = True
    Exit Function
End If
If Len(txtIe) = 0 Then
    MsgBox "Please select Ie", vbExclamation
    txtIe.SetFocus
    prcCheckIn = True
    Exit Function
End If
If Len(txtIe2) = 0 Then
    MsgBox "Please select Ie @ ", vbExclamation
    txtIe2.SetFocus
    prcCheckIn = True
    Exit Function
End If
End If
End Function

Private Sub cmdin_Click()

On Error Resume Next

If prcCoreCheck Then
    Exit Sub
End If
txtnoofcores.Locked = True
cmbCore.Text = "All Cores"
Grid.Rows = Grid.Rows + 1
Grid.row = Grid.Rows - 1
Grid.RowHeight(Grid.row) = 300
Grid.TextMatrix(Grid.row, 0) = cmbnoofcores
If OptComm.value = True Then
    If InStr(1, txtRatio2, "A") > 0 Then
        Grid.TextMatrix(Grid.row, 1) = Left(Trim(txtRatio2), InStr(1, txtRatio2, "A") - 2)
    Else
        Grid.TextMatrix(Grid.row, 1) = txtRatio2
    End If
    
Else
    Grid.TextMatrix(Grid.row, 1) = IvAcc '& "/" & txtRatio2R2
    If IvAcc.ListIndex = IvAcc.ListCount Then
'        IvAcc.ListIndex = IvAcc.ListIndex + 1
        IvAcc.ListIndex = -1
    Else
        IvAcc.ListIndex = IvAcc.ListIndex
    End If
    Grid.TextMatrix(Grid.row, 16) = "A"
End If
Grid.TextMatrix(Grid.row, 1 + 1) = Trim(cmbEquipType)
Grid.TextMatrix(Grid.row, 2 + 1) = Trim(cmbHSV)
Grid.TextMatrix(Grid.row, 3 + 1) = Trim(cmbIL)
Grid.TextMatrix(Grid.row, 4 + 1) = Trim(cmbSTC)
Grid.TextMatrix(Grid.row, 5 + 1) = Trim(cmbTime)
'Filtering for non Special Protection core
If UCase(cmbCType) <> UCase("Special Protection") Then
    Grid.TextMatrix(Grid.row, 6 + 1) = Trim(cmdVA)
    Grid.TextMatrix(Grid.row, 7 + 1) = Trim(cmbClass)
    'Filtering for non metereing
    If UCase(cmbCType) = "METERING" Then
        Grid.TextMatrix(Grid.row, 8 + 1) = Trim(cmbISF)
    'Filtering for non Protection
    ElseIf UCase(cmbCType) = UCase("Protection") Then
        Grid.TextMatrix(Grid.row, 9 + 1) = Trim(cmbISF)
    End If
End If

'locking core
'++++++++++++
    txtRatioDum(Trim(Val(Grid.TextMatrix(Grid.row, 0)) - 1)).Locked = True
    txtRatioR2Dum(Trim(Val(Grid.TextMatrix(Grid.row, 0)) - 1)).Locked = True
'++++++++++++

'For Specification for S P
If UCase(cmbCType) = UCase("Special Protection") Then
    Grid.TextMatrix(Grid.row, 10 + 1) = Trim(txtVK)
    Grid.TextMatrix(Grid.row, 11 + 1) = Trim(txtRct)
    Grid.TextMatrix(Grid.row, 12 + 1) = Trim(txtIe)
    Grid.TextMatrix(Grid.row, 13 + 1) = Trim(txtIe2)
End If

Grid.TextMatrix(Grid.row, 19) = Trim(cmbCType)
' IF SAME THEN 1 ELSEIF DIFFRENT ACCURACY THEN 1 AND MORE
If OptComm.value <> True Then
    Grid.TextMatrix(Grid.row, 17 + 1) = 1 + IvAcc.ListIndex
Else
'    Grid.TextMatrix(Grid.row, 17 + 1) = 1 + IvAcc.ListIndex
    Grid.TextMatrix(Grid.row, 17 + 1) = 1 + IvAcc.ListIndex
End If
    If IvAcc.ListIndex + 1 = IvAcc.ListCount Then
'        IvAcc.ListIndex = IvAcc.ListIndex + 1
        IvAcc.ListIndex = -1
    Else
        IvAcc.ListIndex = IvAcc.ListIndex + 1
'        IvAcc.ListIndex = IvAcc.ListIndex

    End If

If Grid.row > Val(txtnoofcores) Then
    
    Frame7.Caption = ""
    Frame6.Caption = ""
    Frame2.Caption = ""
    
End If

If OptComm.value <> True Then
 
    Dim king As String * 1
    Dim str1, str2 As String
    Dim str3 As Integer
    str3 = 0
    For l = 1 To Len(txtRatio2) + 1
    '    str1 = ""
        king = Mid(txtRatio2, l, l)
        If king = "-" Or Len(txtRatio2) < l Then
            str2 = str1
            If str3 = Grid.row Then
                txtRatio2 = str1
                cmbnoofcores.SetFocus
                cmbVA = ""
                cmbClass.Text = ""
                cmbISF.Text = ""
                txtVK.Text = ""
                txtRct.Text = ""
                txtIe.Text = ""
                txtIe2.Text = ""
                txtVK = ""
                txtRct = ""
                txtIe = ""
                txtIe2 = ""
                cmdVA.Text = ""
                txtRatio2.SetFocus
                Exit Sub
            End If
            str1 = ""
            str3 = str3 + 1
        Else
            str1 = str1 + king
        End If
    Next

Else


End If
If OptInd.value <> True Then
    Frame6 = cmbCType & " (Core " & cmbnoofcores & ") : "
Else
    Frame6 = cmbCType & " (Core " & cmbnoofcores & ") Ratio : " & IvAcc & "A"
End If
'Auto setting core
If OptInd.value = True Then
    
    If Len(IvAcc) > 0 Then
        cmbVA = ""
        cmbClass.Text = ""
        cmbISF.Text = ""
        txtVK.Text = ""
        txtRct.Text = ""
        txtIe.Text = ""
        txtIe2.Text = ""
        txtVK = ""
        txtRct = ""
        txtIe = ""
        txtIe2 = ""
        cmdVA.Text = ""
        Exit Sub
    End If
        
End If
    Dim game1 As Integer
    Dim GIndex As Integer
    For i = 0 To cmbnoofcores.ListCount - 1
        game1 = 0
        For j = 1 To Grid.Rows - 1
            If cmbnoofcores.List(i) = Grid.TextMatrix(j, 0) Then
                game1 = 1
            End If
        Next
        If game1 = 0 Then
            cmbnoofcores.ListIndex = i
            If Len(cmbnoofcores) > 0 Then
                txtRatio2.Text = txtRatioDum(cmbnoofcores - 1) & "/" & txtRatioR2Dum(cmbnoofcores - 1)
            End If
            cmdVA.ListIndex = -1
            ''**** if all the core details are present then blank
            ''**** ratio on specification
            For l = 1 To Grid.Rows - 1
                vinn = 0
                If l = cmbnoofcores Then
                    vinn = 1
                End If
            Next
            If vinn = 1 Then
                txtRatio2.Text = ""
            End If
            cmbnoofcores.Enabled = True
            cmbnoofcores.SetFocus
            Exit For
        End If
    Next
    
    If game1 = 1 Then
        cmbnoofcores.ListIndex = -1
        ''**** if all the core details are present then blank
        ''**** ratio on specification
        For l = 1 To Grid.Rows - 1
            vinn = 0
            If l = cmbnoofcores Then
                vinn = 1
            End If
        Next
        If vinn = 1 Then
            txtRatio2.Text = ""
        End If
        If vinn = 0 And Len(cmbnoofcores) = 0 Then
            txtRatio2.Text = ""
        End If
        If vinn = 0 And Len(cmbnoofcores) = 0 Then
            txtRatio2.Text = ""
        End If
        cmbCType.ListIndex = -1
        cmbCType.Text = "-Select Core Type-"
        cmbCType.SetFocus
        Exit Sub
    End If


cmbCType.ListIndex = -1
cmbCType.Text = "-Select Core Type-"
If Len(cmbnoofcores) > 0 Then
    txtRatio2.Text = txtRatioDum(cmbnoofcores - 1) & "/" & txtRatioR2Dum(cmbnoofcores - 1)
Else
    txtRatio2.Text = ""
End If

''**** if all the core details are present then blank
''**** ratio on specification
For l = 1 To Grid.Rows - 1
    vinn = 0
    If l = cmbnoofcores Then
        vinn = 1
    End If
Next
If vinn = 1 Then
    txtRatio2.Text = ""
End If
If vinn = 0 And Len(cmbnoofcores) = 0 Then
    txtRatio2.Text = ""
End If
'' ++++++++++++++++++++
'' Clear
cmbVA = ""
cmbClass.Text = ""
cmbISF.Text = ""
txtVK.Text = ""
txtRct.Text = ""
txtIe.Text = ""
txtIe2.Text = ""
txtVK = ""
txtRct = ""
txtIe = ""
txtIe2 = ""

End Sub

Private Sub cmdNext_Click(Index As Integer)
SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub cmdPre_Click(Index As Integer)
SSTab1.Tab = SSTab1.Tab - 1
End Sub

Private Sub cmdRemCore_Click()

On Error Resume Next
If Grid.Rows = 1 Then Exit Sub

joel = MsgBox("Do you want to delete All core or Specific core?" + "Press yes for All Press No for Specific", vbQuestion + vbYesNo)
If vbYes = joel Then
    Grid.Rows = 1
    txtRatioDum(0).Locked = False
    txtRatioR2Dum(0).Locked = False
    txtRatioDum(1).Locked = False
    txtRatioR2Dum(1).Locked = False
    txtRatioDum(2).Locked = False
    txtRatioR2Dum(2).Locked = False
    txtRatioDum(3).Locked = False
    txtRatioR2Dum(3).Locked = False
    txtRatioDum(4).Locked = False
    txtRatioR2Dum(4).Locked = False
    txtnoofcores.Locked = False
    Exit Sub
End If

If G1.Cols > 1 Or GCore.Rows > 1 Then
    joel = MsgBox("Changing the core specification will effect the core results" & vbCrLf & "Are you sure you want to delete this core dimension  ?", vbCritical + vbAbortRetryIgnore + vbYesNo)
    If joel = vbNo Then
        Exit Sub
    Else
        G1.Cols = 1
        wounded.Rows = 1
        Former.Rows = 1
        Casting.Rows = 1
        Mould.Rows = 1
'        GCore.Rows = 1
        txtdrg = ""
        cmbType.Clear
    End If
End If

If Grid.Rows = 1 Then
    MsgBox "No rows to delete", vbInformation
    Exit Sub
End If

If Grid.Rows = 2 Then
    Grid.Rows = 1
    txtRatioDum(0).Locked = False
    txtRatioR2Dum(0).Locked = False
    txtnoofcores.Locked = False
    cmdIn.Enabled = True
Else
    txtRatioDum(Grid.RowSel - 1).Locked = False
    txtRatioR2Dum(Grid.RowSel - 1).Locked = False
    Grid.RemoveItem (Grid.RowSel)
    cmdIn.Enabled = True
End If

If cmbCType.Text = "Protection" Then
    Label15 = "ALF :"
    Frame6.Caption = "Protection (Core " & Grid.row + 1 & ")"
ElseIf cmbCType.Text = "Metering" Then
    Label15 = "ISF :"
    Frame6.Caption = "Metering (Core " & Grid.row + 1 & ")"
ElseIf cmbCType.Text = "Special Protection " Then
    Frame7.Caption = "Special Protection (Core " & Grid.row + 1 & ")"
End If

'Frame2.Caption = "Core  Ratio " & Grid.row + 1
cmbnoofcores.SetFocus


End Sub

Private Function ModifyDes()

'Deleting details of design for mofifing details

Set rst = New Recordset
rst.Open "delete from designspec where des_code = '" & DesignNo & "'", con, adOpenDynamic, adLockOptimistic
Set rst = Nothing

Set rst = New Recordset
rst.Open "delete from designerror where des_code = '" & DesignNo & "'", con, adOpenDynamic, adLockOptimistic
Set rst = Nothing

Set rst = New Recordset
rst.Open "delete from designcore where deC_code = '" & DesignNo & "'", con, adOpenDynamic, adLockOptimistic
Set rst = Nothing

End Function

Private Sub cmdRemoCore_Click()

On Error Resume Next


If G1.Cols > 1 Then
    joel = MsgBox("Changing the core dimension will effect the core results" & vbCrLf & "Are you sure you want to delete this core dimension  ?", vbCritical + vbYesNo)
    If joel = vbNo Then
        Exit Sub
    Else
        G1.Cols = 1
        wounded.Rows = 1
        Former.Rows = 1
        Casting.Rows = 1
        Mould.Rows = 1
        txtdrg = ""
        cmbSelctCore.Text = ""
        txtOD1.Text = ""
        txtID1 = ""
        txtOL1.Text = ""
        txtOW.Text = ""
        txtWW.Text = ""
        txtWL.Text = ""
        txtHt.Text = ""
        txtSwg = ""
        txtNo = ""
        SWG.Rows = 1
        cmbType.Clear
    End If
End If
'Gcore has no data row then
If GCore.Rows = 1 Then
    Frame1.Enabled = True
    Exit Sub
End If

'Core dimenstion is comman the delete ass core details
If OptCom(0).value = True Then
    GCore.Rows = 1
    Frame1.Enabled = True
    Exit Sub
End If

If GCore.Rows = 2 Then
    GCore.Rows = 1
    Frame1.Enabled = True
    
    CNo = 1
    For K = 1 To txtnoofcores
        Cout = ""
        For j = 1 To GCore.Rows - 1
            If GCore.TextMatrix(j, 0) = K Then
                Cout = "K"
                Exit For
            End If
        Next
        If Cout <> "K" Then
            Exit For
        End If
    Next
            
    If Cout = "" Then
        cmbCore = "Core " & K
    End If
    
    Exit Sub
End If

'Dim Cout As String
Cout = """"
GCore.RemoveItem (GCore.RowSel)
CNo = 1
For K = 1 To txtnoofcores
    Cout = ""
    For j = 1 To GCore.Rows - 1
        If GCore.TextMatrix(j, 0) = K Then
            Cout = "K"
            Exit For
        End If
    Next
    If Cout <> "K" Then
        Exit For
    End If
Next
        
If Cout = "" Then
    cmbCore = "Core " & K
End If
'If GCore.Rows = 1 Then
'    OptCom(0).Enabled = True
'    OptCom(1).Enabled = True
'End If


End Sub


Private Sub cmdS1_Click()
SSTab1.Tab = 0
End Sub

Private Sub cmdS2_Click()
SSTab1.Tab = 1
End Sub

Private Sub cmdS3_Click()
If InStr(1, UCase(cmbEquipType), "WOUND") > 0 Then
    SSTab1.Tab = 2
Else
    MsgBox "When equipment BAR or WINDOW user cannot goto Primary Details", vbExclamation
End If
End Sub

Private Sub cmdS4_Click()
SSTab1.Tab = 3
End Sub

Private Function prcCheckingSave() As Boolean


'' Applicable for wound and bar
'If InStr(1, UCase(cmbEquipType), "WOUND") > 0 Or InStr(1, UCase(cmbEquipType), "BAR") > 0 Then
If InStr(1, UCase(cmbEquipType), "WOUND") > 0 Then
    If Trim(txtDelStc) > 180 Then
        MsgBox "Short circuit current exceds the limit. Please check", vbCritical
        prcCheckingSave = True
        Exit Function
    End If

End If

''Details of core diemension
If Trim(wounded.Rows - 1) < (Val(txtnoofcores) - 1) Then
    MsgBox "Wounded Dimension not set", vbCritical
    prcCheckingSave = True
    Exit Function
End If

''INSLUATION
If Len(Trim(CMBINSULATION)) = 0 Then
    MsgBox "Please set inslution Scheme", vbCritical
    CMBINSULATION.SetFocus
    prcCheckingSave = True
    Exit Function
End If

''Check all core calculations
Dim Sta As Integer
For i = 1 To Val(txtnoofcores)
    Sta = 0
    For j = 1 To G1.Cols - 1
        If i = G1.TextMatrix(1, j) Then
            Sta = 1
        End If
        If Sta = 1 Then
            Exit For
        End If
    Next
    If Sta = 0 Then
        Exit For
    End If
Next

If Sta = 0 Then
    MsgBox "Please check the calculations of core : " & i, vbCritical
    SSTab1.Tab = 3
    prcCheckingSave = True
    Exit Function
End If

''core size
If Trim(GCore.Rows - 1) < (Val(txtnoofcores) - 1) Then
    MsgBox "Please set core sizes", vbCritical
    prcCheckingSave = True
    Exit Function
End If

End Function

Private Sub cmdSave_Click()

On Error GoTo err

Dim Dno As String

If prcCheckingSave Then
    Exit Sub
End If

joel = MsgBox("Are you sure this design is OK?", vbYesNo + vbQuestion)
If joel = vbNo Then
    Exit Sub
End If

STCMemory = ""
STCTimeMemory = ""

If ModDesign = "MOD" Then
    ModifyDes
    Dno = txtDesNo
    ModDesign = ""
    DesignNo = ""
Else
    '++++++++++++++++++++++++++++++++++++ txtRateConti
    '+++Designo
    Dim Rep As AutoDNo
    Set Rep = New AutoDNo
       Dno = Rep.Autogen
    Set Rep = Nothing
    '+++End Design No
    '++++++++++++++++++++++++++++++++++++
End If
    
    Dim details As String
    '+++++++++++++++++++++++++++++++++
    'New Design code
        code = Dno

    '+++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++++++++++++++++
    'SPECIFICATION AND CORE DIMENSIONS , FORMER , CASTING , WEIGHT , MOULD DETAILS
    Set rst = New Recordset
    rst.Open "select * from designspec", con, adOpenDynamic, adLockOptimistic
    For i = 1 To Grid.Rows - 1
        
        rst.AddNew
        '+++++++++++++++++++++++++++++++++
        'Cus Specification 1
            rst!ratio = ratio
            rst!Des_code = code
            rst!freq = Trim(cmbFreq)
            rst!stc = Trim(cmbSTC)
            rst!hsv = Trim(cmbHSV)
            rst!NSV = Trim(cmbNSV)
            rst!iL = Trim(cmbIL)
            rst!Time = Trim(cmbTime)
            rst!ratedpri = Trim(txtRateConti)
            rst!core_ratio = Trim(Grid.TextMatrix(i, 1))
            rst!Equip = Trim(cmbEquipType.Text)
            rst!core_no = Trim(Grid.TextMatrix(i, 0))
            rst!tap_no = Trim(Grid.TextMatrix(i, 18))
            rst!core_type = Trim(Grid.TextMatrix(i, 19))
            rst!sp_va = Trim(Grid.TextMatrix(i, 7))
            rst!sp_class = Trim(Grid.TextMatrix(i, 8))
            rst!sp_isf = Trim(Grid.TextMatrix(i, 9))
            rst!sp_alf = Trim(Grid.TextMatrix(i, 10))
            rst!sp_vk = Trim(Grid.TextMatrix(i, 11))
            rst!sp_rct = Trim(Grid.TextMatrix(i, 12))
            rst!sp_ie = Trim(Grid.TextMatrix(i, 13))
            rst!sp_ie_at = Trim(Grid.TextMatrix(i, 14))
        'END Cus Specification 1 Aeff
        '+++++++++++++++++++++++++++++++++
                
        '+++++++++++++++++++++++++++++++++
        'RATIO SECONADARY AND PRIMARY
            rst!pri_ratio = txtRatioDum(rst!core_no - 1)
            rst!sec_ratio = txtRatioR2Dum(rst!core_no - 1)
        'END RATIO SECONADARY AND PRIMARY
        '+++++++++++++++++++++++++++++++++
        
        '+++++++++++++++++++++++++++++++++
        'CORE DIMENSIONS 2
            For g = 1 To GCore.Rows - 1
                If rst!core_no = GCore.TextMatrix(g, 0) Then
                    rst!OL = Val(Trim(GCore.TextMatrix(g, 1)))
                    rst!OW = Val(Trim(GCore.TextMatrix(g, 2)))
                    rst!WL = Val(Trim(GCore.TextMatrix(g, 3)))
                    rst!WW = Val(Trim(GCore.TextMatrix(g, 4)))
                    rst!oD = Val(Trim(GCore.TextMatrix(g, 5)))
                    rst!ID = Val(Trim(GCore.TextMatrix(g, 6)))
                    Exit For
                End If
            Next
            rst!CoreShape = Trim(cmbStyle.Text)
        'END CORE DIMENSIONS 2
        '+++++++++++++++++++++++++++++++++
        
        '+++++++++++++++++++++++++++++++++
        'PRIMARY DETAILS 3
            rst!pri_turns = Trim(txtTurns.Text)
            ''TO BE CALCULATED
            rst!amp_turns = Trim(txtTurns.Text)
            If InStr(1, Trim(UCase(cmbEquipType)), "WO-R") > 1 Then
                rst!pri_cond_style = Trim(cmbConStyle.List(cmbConStyle.ListIndex, 3))
            End If
            rst!pri_cross_sec = Trim(txtPriCros)
            rst!pri_coil = Trim(txtCondPer)
            rst!pri_cond = Trim(txtTotalCon)
            rst!pri_wind_ht_side = Trim(txtSide)
            rst!pri_wind_ht_top = Trim(txtTop)
            rst!pri_wind_ht_width = Trim(txtWidth)
            rst!pri_nor_curr_den = Trim(txtDelR)
            rst!pri_stc_curr_den = Trim(txtDelStc)
            rst!pri_i_dyn_amp = Trim(txtIDyn)
            rst!pri_i_dyn_amp_turns = Trim(txtIDATurns)
        'END PRIMARY DETAILS 3
        '+++++++++++++++++++++++++++++++++
        
        '+++++++++++++++++++++++++++++++++
        'FORMER DETAILS 3
            If Former.Rows > 1 Then
                rst!former_inside = Trim(Former.TextMatrix(1, 1))
                rst!former_inside2 = Trim(Former.TextMatrix(1, 2))
                rst!former_outside = Trim(Former.TextMatrix(2, 1))
                rst!former_outside2 = Trim(Former.TextMatrix(2, 2))
                'DISTANCE PIECE
                rst!Dist_piece = Trim(Former.TextMatrix(5, 1))
            End If
        'END FORMER DETAILS 3
        '+++++++++++++++++++++++++++++++++
        
        '+++++++++++++++++++++++++++++++++
        'INSULATION  ADD 3
            rst!insulation = Trim(CMBINSULATION)
        'END
        '+++++++++++++++++++++++++++++++++
        
        '+++++++++++++++++++++++++++++++++
        'CAST DETAILS 3
            If InStr(1, UCase(cmbEquipType), "WINDOW") = 0 And InStr(1, UCase(cmbEquipType), "BAR") = 0 Then
                'WINDING
                rst!cast_wind_len = Trim(Casting.TextMatrix(1, 1))
                rst!cast_wind_width = Trim(Casting.TextMatrix(1, 2))
                rst!cast_wind_ht = Trim(Casting.TextMatrix(1, 3))
                rst!cast_wind_seclen = Trim(Casting.TextMatrix(1, 4))
                'MOULD
                rst!cast_mould_len = Trim(Casting.TextMatrix(3, 1))
                rst!cast_mould_width = Trim(Casting.TextMatrix(3, 2))
                rst!cast_mould_ht = Trim(Casting.TextMatrix(3, 3))
                rst!cast_mould_seclen = Trim(Casting.TextMatrix(3, 4))
                'EPOXY
                rst!cast_epoxy_len = Trim(Casting.TextMatrix(4, 1))
                rst!cast_epoxy_width = Trim(Casting.TextMatrix(4, 2))
                rst!cast_epoxy_ht = Trim(Casting.TextMatrix(4, 3))
                rst!cast_epoxy_seclen = Trim(Casting.TextMatrix(4, 4))
                'MOULD NO. TO BE TAKEN FROM DATABASE XXXXXXXXXXX
                If Len(G1.Tag) > 0 Then
                    rst!mould_no = G1.Tag
                Else
                    rst!mould_no = "N.A"
                End If
            Else
                'WINDING
                rst!cast_wind_len = Trim(Casting.TextMatrix(1, 1))
                rst!cast_wind_width = Trim(Casting.TextMatrix(1, 2))
                rst!cast_wind_ht = Trim(Casting.TextMatrix(1, 3))
'                rst!cast_wind_seclen = Trim(Casting.TextMatrix(1, 4))
                'MOULD
                rst!cast_mould_len = Trim(Casting.TextMatrix(2, 1))
                rst!cast_mould_width = Trim(Casting.TextMatrix(2, 2))
                rst!cast_mould_ht = Trim(Casting.TextMatrix(2, 3))
'                rst!cast_mould_seclen = Trim(Casting.TextMatrix(3, 4))
                'EPOXY
                rst!cast_epoxy_len = Trim(Casting.TextMatrix(3, 1))
                rst!cast_epoxy_width = Trim(Casting.TextMatrix(3, 2))
                rst!cast_epoxy_ht = Trim(Casting.TextMatrix(3, 3))
'                rst!cast_epoxy_seclen = Trim(Casting.TextMatrix(4, 4))
                'MOULD NO. TO BE TAKEN FROM DATABASE XXXXXXXXXXX
    
                If Len(G1.Tag) > 0 Then
                    rst!mould_no = G1.Tag
                Else
                    rst!mould_no = "N.A"
                End If
            End If
        'END CAST DETAILS 3
        '+++++++++++++++++++++++++++++++++
        
        '+++++++++++++++++++++++++++++++++
        'DRG. NO(S)4
            rst!drg_no = Trim(txtdrg)
            rst!DRG_TYPE = Trim(cmbType)
            rst!WINBARLABEL = Trim(WindowSize)
            'DRG MOD TO BE SET ***************
            'rst!drg_MOD =
        'END DRG NO(S) 4
        '+++++++++++++++++++++++++++++++++
        
        '+++++++++++++++++++++++++++++++++
        'CALCULATED LENGTH AND WEIGHT(PRI) 5
            rst!tot_wt_core_wt = Trim(txtCoreWT)
            rst!tot_wt_scu_wt = Trim(txtSecCuWt)
            If InStr(1, UCase(cmbEquipType), "WOUND") > 0 Then
                rst!tot_wt_pricu_wt = Trim(txtPrimaryCopper)
                rst!pri_wind_mean_len = Trim(txtPriMeanL)
                rst!pri_wind_mean_len = Trim(txtPriMeanL)
                rst!Len = Trim(txtPriLength)
                rst!pri_cu_wt = Trim(txtpricuWt)
            Else
                rst!tot_wt_pricu_wt = Val(txtPrimaryCopper)
                rst!pri_wind_mean_len = Val(txtPriMeanL)
                rst!pri_wind_mean_len = Val(txtPriMeanL)
                rst!Len = Val(txtPriLength)
                rst!pri_cu_wt = Val(txtpricuWt)
            End If
            rst!tot_wt_epoxy_wt = Trim(txtResinWt)

'            rst!pri_wind_mean_len = Trim(txtPriMeanL)
'            rst!Len = Trim(txtPriLength)
'            rst!tot_len = Trim(txtPriTotLength)
'            rst!pri_cu_wt = Trim(txtpricuWt)
            
            'WINDOW CLEARANCES
            rst!window_clear_ep1 = Val(Trim(WOR1A.Caption))
            rst!window_clear_ep2 = Val(Trim(WOR2A.Caption))
            
        'END CALCULATED LENGTH AND WEIGHT(PRI) 5
        '+++++++++++++++++++++++++++++++++
        
        '+++++++++++++++++++++++++++++++++
        'FINAL DIMENSIONS 6 TOO BE CHECKED*************************************************
            For H = 1 To Val(txtnoofcores)
                If rst!core_no = wounded.TextMatrix(H, 0) Then
                    rst!final_dim1 = wounded.TextMatrix(H, 1)
                    rst!final_dim2 = wounded.TextMatrix(H, 2)
                    rst!final_ht = wounded.TextMatrix(H, 3)
                    rst!final_BU = wounded.TextMatrix(H, 4)
                    Exit For
                End If
            Next
            
            For f = 1 To G1.Cols - 1
                If G1.TextMatrix(1, f) = rst!core_no Then
                    rst!HT = G1.TextMatrix(32, f)
                    rst!wt = G1.TextMatrix(81, f)
                    rst!builtup = G1.TextMatrix(26, f)
                    rst!Aeff = G1.TextMatrix(27, f)
                    rst!Lm = G1.TextMatrix(28, f)
                    Exit For
                End If
            Next
            
            If InStr(1, wounded.TextMatrix(wounded.Rows - 1, 1), "X") > 0 Then
                rst!max_dim_ol = Mid(Trim(wounded.TextMatrix(wounded.Rows - 1, 1)), 1, InStr(1, wounded.TextMatrix(wounded.Rows - 1, 1), "X") - 1)
                rst!max_dim_OW = Mid(Trim(wounded.TextMatrix(wounded.Rows - 1, 1)), InStr(1, wounded.TextMatrix(wounded.Rows - 1, 1), "X") + 1)
            Else
                rst!max_dim_id = Trim(wounded.TextMatrix(wounded.Rows - 1, 1))
            End If
            If InStr(1, wounded.TextMatrix(wounded.Rows - 1, 2), "X") > 0 Then
                rst!max_dim_WL = Mid(Trim(wounded.TextMatrix(wounded.Rows - 1, 2)), 1, InStr(1, wounded.TextMatrix(wounded.Rows - 1, 2), "X") - 1)
                rst!max_dim_WW = Mid(Trim(wounded.TextMatrix(wounded.Rows - 1, 2)), InStr(1, wounded.TextMatrix(wounded.Rows - 1, 2), "X") + 1)
            Else
                rst!max_dim_od = Trim(wounded.TextMatrix(wounded.Rows - 1, 2))
            End If
            rst!max_dim_ht = Trim(wounded.TextMatrix(wounded.Rows - 1, 3))
            rst!max_dim_BU = Trim(wounded.TextMatrix(wounded.Rows - 1, 4))
            rst!Standard = 0
        'END FINAL DIMENSIONS 6
        '+++++++++++++++++++++++++++++++++
        rst.Update
    Next
    'END SPECIFICATION AND CORE DIMENSIONS , FORMER , CASTING , WEIGHT , MOULD DETAILS
    '+++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++++++++++++++++
    'ERRORS
    Set rst = New Recordset
'    rst.Open "DELETE from designerror", con, adOpenDynamic, adLockOptimistic
    rst.Open "select * from designerror", con, adOpenDynamic, adLockOptimistic
'    For Count1 = 1 To G1.Cols - 1
        For i = 1 To G1.Cols - 1
'            If Count1 = i Then
                rst.AddNew
                rst!Des_code = code 'to be generated
                rst!tap = Val(Trim(G1.TextMatrix(108, i)))
                rst!coreno = Trim(G1.TextMatrix(1, i))
                rst!Name1 = Trim(G1.TextMatrix(69, i))
                rst!NAME2 = Trim(G1.TextMatrix(70, i))
                'PERCENTAGE PRIMARY CURRENT
                rst!per_pri_curr = Trim(G1.TextMatrix(82, i))
                rst!per_pri_curr1 = Trim(G1.TextMatrix(87, i))
                rst!ppc1 = Trim(G1.TextMatrix(83, i))
                rst!ppc11 = Trim(G1.TextMatrix(88, i))
                rst!ppc2 = Trim(G1.TextMatrix(84, i))
                rst!ppc12 = Trim(G1.TextMatrix(89, i))
                rst!ppc3 = Trim(G1.TextMatrix(85, i))
                rst!ppc13 = Trim(G1.TextMatrix(90, i))
                rst!ppc4 = Trim(G1.TextMatrix(86, i))
                rst!ppc14 = Trim(G1.TextMatrix(91, i))
                'BM
                rst!BM = Trim(G1.TextMatrix(92, i))
                rst!B1 = Trim(G1.TextMatrix(97, i))
                rst!BM1 = Trim(G1.TextMatrix(93, i))
                rst!BM11 = Trim(G1.TextMatrix(98, i))
                rst!BM2 = Trim(G1.TextMatrix(94, i))
                rst!BM12 = Trim(G1.TextMatrix(99, i))
                rst!BM3 = Trim(G1.TextMatrix(95, i))
                rst!BM13 = Trim(G1.TextMatrix(100, i))
                rst!BM4 = Trim(G1.TextMatrix(96, i))
                rst!BM14 = Trim(G1.TextMatrix(101, i))
                'RATIO ERROR
'                rst!RERR = Trim(G1.TextMatrix(i, 36))
'                rst!R1ERR = Trim(G1.TextMatrix(i, 57))
                rst!RERR1 = Trim(G1.TextMatrix(36, i))
                rst!RERR11 = Trim(G1.TextMatrix(57, i))
                rst!RERR2 = Trim(G1.TextMatrix(37, i))
                rst!RERR12 = Trim(G1.TextMatrix(58, i))
                rst!RERR3 = Trim(G1.TextMatrix(38, i))
                rst!RERR13 = Trim(G1.TextMatrix(59, i))
                rst!RERR4 = Trim(G1.TextMatrix(39, i))
                rst!RERR14 = Trim(G1.TextMatrix(60, i))
                'com
'                rst!com_RERR = Trim(G1.TextMatrix(i, 36))
'                rst!com1_RERR = Trim(G1.TextMatrix(i, 57))
                rst!com_RERR1 = Trim(G1.TextMatrix(40, i))
                rst!com_RERR11 = Trim(G1.TextMatrix(61, i))
                rst!com_RERR2 = Trim(G1.TextMatrix(41, i))
                rst!com_RERR12 = Trim(G1.TextMatrix(62, i))
                rst!com_RERR3 = Trim(G1.TextMatrix(42, i))
                rst!com_RERR13 = Trim(G1.TextMatrix(63, i))
                rst!com_RERR4 = Trim(G1.TextMatrix(43, i))
                rst!com_RERR14 = Trim(G1.TextMatrix(64, i))
                'PERR
'                rst!PERR = Trim(G1.TextMatrix(i, 36))
'                rst!P1ERR = Trim(G1.TextMatrix(i, 57))
                rst!PERR1 = Trim(G1.TextMatrix(44, i))
                rst!PERR11 = Trim(G1.TextMatrix(65, i))
                rst!PERR2 = Trim(G1.TextMatrix(45, i))
                rst!PERR12 = Trim(G1.TextMatrix(66, i))
                rst!PERR3 = Trim(G1.TextMatrix(46, i))
                rst!PERR13 = Trim(G1.TextMatrix(67, i))
                rst!PERR4 = Trim(G1.TextMatrix(47, i))
                rst!PERR14 = Trim(G1.TextMatrix(68, i))
                rst.Update
'                Exit For
'            End If
        Next
'    Next
    Set rst = Nothing
    'ERRORS
    '+++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++++++++++++++++
    '
        Set rst = New Recordset
        rst.Open "select * from designcore", con, adOpenDynamic, adLockOptimistic
        For i = 1 To G1.Cols - 1
            rst.AddNew
            rst!dec_code = code
            rst!min_reg_ps = Trim(G1.TextMatrix(104, i))
            rst!sec_cond = Trim(G1.TextMatrix(102, i))
            rst!sec_cond_dia = Trim(G1.TextMatrix(103, i))
            rst!sec_cond = Trim(G1.TextMatrix(102, i))
            rst!r_turns = Trim(G1.TextMatrix(3, i))
            rst!Layers = Trim(G1.TextMatrix(8, i))
            rst!w_ht = Trim(G1.TextMatrix(9, i))
            rst!rct75 = Trim(G1.TextMatrix(10, i))
            rst!flux = Trim(G1.TextMatrix(24, i))
            rst!r_vk = Trim(G1.TextMatrix(12, i))
            rst!sec_cu_wt = Trim(G1.TextMatrix(80, i))
            rst!bet_turns = Trim(G1.TextMatrix(105, i)) 'TO BE ADDED
            rst!isf = Trim(G1.TextMatrix(17, i))
            rst!sh_cir_curr_den = Trim(G1.TextMatrix(18, i))
            rst!sat_curr_den = Trim(G1.TextMatrix(19, i))
            rst!sh_temp_rise = Trim(G1.TextMatrix(20, i))
            rst!ie_nor = Trim(G1.TextMatrix(15, i))
            rst!ie_alf = Trim(G1.TextMatrix(16, i))
            rst!ce_nor = Trim(G1.TextMatrix(22, i))
            rst!ce_alf = Trim(G1.TextMatrix(23, i))
            rst!compen = Trim(G1.TextMatrix(106, i)) 'TO BE ADDED
            rst!core_mat = Trim(G1.TextMatrix(107, i)) 'TO BE ADDED
            rst!core_ht = Trim(G1.TextMatrix(71, i)) 'TO BE checked
            rst!tot_layers = Trim(G1.TextMatrix(72, i))
            rst!TOT_CU_WT = Trim(G1.TextMatrix(80, i))
            rst!exc_curr = Trim(G1.TextMatrix(11, i))
            rst!C1 = Trim(G1.TextMatrix(13, i))
            rst!c2 = Trim(G1.TextMatrix(14, i))
            rst!exc_curr = Trim(G1.TextMatrix(21, i))
            rst!wind_ol = Val(Trim(G1.TextMatrix(75, i)))
            rst!wind_ow = Val(Trim(G1.TextMatrix(76, i)))
            rst!wind_wl = Val(Trim(G1.TextMatrix(77, i)))
            rst!wind_ww = Val(Trim(G1.TextMatrix(78, i)))
            rst!wind_ht = Val(Trim(G1.TextMatrix(79, i)))
            rst!wind_od = Val(Trim(G1.TextMatrix(73, i)))
            rst!wind_id = Val(Trim(G1.TextMatrix(74, i)))
            rst!tap = Val(Trim(G1.TextMatrix(108, i)))
            rst!coreno = Trim(G1.TextMatrix(1, i))
            rst!pri_curr = Trim(G1.TextMatrix(4, i))
            rst!sec_curr = Trim(G1.TextMatrix(5, i))
            rst!Area = Trim(G1.TextMatrix(27, i))
            rst!BUP = Trim(G1.TextMatrix(26, i))
            rst!Mean = Trim(G1.TextMatrix(28, i))
            rst!sp_alf = Trim(G1.TextMatrix(25, i))
            rst!sp_rct = Trim(G1.TextMatrix(30, i))
            rst!sp_ie = Trim(G1.TextMatrix(31, i))
            rst!sp_vk = Trim(G1.TextMatrix(32, i))
            rst!sp_va = Trim(G1.TextMatrix(33, i))
            rst!sp_class = Trim(G1.TextMatrix(34, i))
            rst!sp_isf = Trim(G1.TextMatrix(35, i))
            rst!core_type = Trim(G1.TextMatrix(48, i))
            rst!core_style = Trim(G1.TextMatrix(49, i))
            rst!core_ratio = Trim(G1.TextMatrix(0, i))
            rst!oD = Val(Trim(G1.TextMatrix(51, i)))
            rst!ID = Val(Trim(G1.TextMatrix(52, i)))
            rst!OL = Val(Trim(G1.TextMatrix(53, i)))
            rst!OW = Val(Trim(G1.TextMatrix(54, i)))
            rst!WL = Val(Trim(G1.TextMatrix(55, i)))
            rst!WW = Val(Trim(G1.TextMatrix(56, i)))
            rst!core_wt = Trim(G1.TextMatrix(81, i))
            rst!tap_ratio = Trim(G1.TextMatrix(2, i))
            rst!SWG = Trim(G1.TextMatrix(7, i))
            rst!SWGNO = Trim(G1.TextMatrix(6, i))
            rst!max_tap = Trim(G1.TextMatrix(109, i))
            rst.Update
        Next
    '
    '+++++++++++++++++++++++++++++++++
joel = MsgBox("Do you want to view the preiview of" + vbCrLf + "Design No :" & code, vbInformation + vbYesNo)
If joel = vbNo Then Exit Sub

'GENERATING REPORT
Dim Rep1 As Report
Set Rep1 = New Report
    Rep1.ReportDesign (code)
Set Rep1 = Nothing

DCode = code

Dim EQ As String
EQ = Trim(cmbEquipType)

''+++++++++++++++++++++++++++++++

'CancelCT = "1"
'Unload Me
'Set CTfrm = Nothing
'CancelCT = ""
'CTfrm.Show
cmbStyle = "Rectangluar"
cmbCore.Text = "All Cores"
cmbHSV = ""
prcclear
prcLastCal
txtnoofcores = ""
cmbEquipType.ListIndex = -1
cmbStyle.ListIndex = -1
prcClearALL
CleanAll
txtnoofcores = ""
''+++++++++++++++++++++++++++++++

If vbYes = joel Then
    ''++++++MOD = "" for mofify
    ModDesign = ""
    DesignNo = ""
    ''""++++++
    
    CrystalReport1.DiscardSavedData = True
    CrystalReport2.DiscardSavedData = True
    If InStr(1, UCase(EQ), "WOUND") > 0 Then
        CrystalReport1.ReportFileName = App.Path & "\REPORTs\ctdesign5.rpt"
'        CrystalReport1.ReportFileName = "\\server\c\reports\design\ctdesignwindow1.rpt"
    Else
        CrystalReport1.ReportFileName = App.Path & "\REPORTs\ctdesignwindow.rpt"
    End If
    CrystalReport2.ReportFileName = App.Path & "\REPORTs\Costing.rpt"
    CrystalReport1.Connect = "dsn=enhance;uid=sa;pwd=shilpraj;database=gem0607"
    
    CrystalReport2.ReportTitle = DCode
    CrystalReport2.Connect = "dsn=enhance;uid=sa;pwd=shilpraj;database=gem0607"
    CrystalReport1.Action = 1
    CrystalReport2.Action = 1
End If


err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If

End Sub

Private Sub CleanAll()
    
    Frame14.Visible = False
    Frame15.Visible = False
    Frame16.Visible = False
    Frame22.Visible = False
    Frame21.Visible = False
    Label90.Visible = False
    Label91.Visible = False
    FrameWound.Visible = False
    WindowSize.Visible = False
    Label85.Visible = False
    Label86.Visible = False
    Label87.Visible = False
    txtPRiRes20.Visible = False
    txtPRiRes75.Visible = False
    txtWattPri.Visible = False
    WindowSize.Caption = ""
    txtdrg = "PR-"
    cmbType.Clear
    txtPriCros.Text = ""
    txtTurns = ""
    txtWidth = ""
    txtTop = ""
    txtSide = ""
    txtIDyn = ""
    txtIDATurns = ""
    txtDelR = ""
    txtDelStc = ""
    lstCond.Clear
    txtCondPer = ""
    txtTotalCon = ""
End Sub


Private Sub prcClearALL()

SWG.Rows = 1
GCore.Rows = 1
Grid.Rows = 1
wounded.Rows = 1
Former.Rows = 1
Casting.Rows = 1
Mould.Rows = 1
G1.Cols = 1
For Each Control In CTfrm
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
txtnoofcores = ""

'cmbEquipType.ListIndex = cmbEquipType.ListCount - 1

lstCond.Clear
CMBINSULATION.Clear
cmbType.Clear

WOR1.Caption = ""
WOR2A.Caption = ""

txtminCs.Text = ""
txtPriCros = ""
txtTotalCon = ""
txtCondPer = ""
txtSide = ""
txtTop = ""
txtIDyn = ""
txtIDATurns = ""
txtDelR = ""
txtDelStc = ""
'txtdrg = "PR-"

End Sub

Private Sub cmdVA_GotFocus()
'cmbVA.SelStart = 0
'cmbVA.SelLength = Len(cmbVA)
MDIDesign.StatusBar1.Panels(2).Text = "Set VA "
End Sub

Private Sub CommandButton1_Click()
'    MsgBox Grid5.Cols
'    For i = 1 To Grid.Rows - 1
'        MsgBox prcCoreHight(Grid.TextMatrix(i, 9), Grid5.TextMatrix(i, 1), 30)
'    Next
    SearchDesign
End Sub

Private Sub SearchDesign()
'Search simillar design No
Dim WONO As String
Dim QUONO As String
Dim NoSearch As Integer
NoSearch = 0
    
    searchResultsfrm.lstDno.Clear
    
    Dim maxcoreno As Integer
    Dim nsvDetails As String
    Dim DesNo As String
    Dim Design As String
    Dim ctr As Integer
    Dim MAXCORENO_Find As Integer
    
    If Val(cmbNSV) <= 11 Then
        nsvDetails = "3.3 kV','6.6 kV','11 kV"
    Else
        nsvDetails = Trim(cmbNSV)
    End If
    
    maxcoreno = Val(txtnoofcores)
    Dim RT As Recordset
    Dim VKP As String
    If InStr(1, UCase(cmbIL), "K") > 0 Then
        VKP = Trim(Mid(cmbIL, 1, InStr(1, UCase(cmbIL), "K") - 1))
    Else
        VKP = Trim(cmbIL)
    End If
    
    Set RT = New Recordset
    sql = "select * from designspec where stc = '" & Trim(cmbSTC) & "' and freq = '" & Trim(cmbFreq) & "' and " _
          & "time = '" & Trim(cmbTime) & "' and equip = '" & Trim(cmbEquipType) & "' and nsv in ('" & nsvDetails & "') and core_no = " & maxcoreno & " and des_code in (select des_code from designspec where core_no = " & maxcoreno & ") order by des_code"
    RT.Open sql, con, adOpenKeyset, adLockReadOnly
    If RT.EOF = False And RT.BOF = False Then
        RT.MoveFirst
        Do Until RT.EOF
            If DesNo <> Trim(RT!Des_code) Then
                '++++++++++++++++++++++++
                'Find noof core in the matched design no
'                If Val(txtnoofcores) = maxcore(RT!Des_code) Then
'                    MsgBox "^^^^^^^^^^^^^^^^^^ " & RT!Des_code
                    'find ing noof core
                    sql = "select max(core_no) from designspec where des_code = '" & Trim(RT!Des_code) & "'"
                    Set rst = New Recordset
                        rst.Open sql, con, adOpenStatic, adLockReadOnly
                        If rst.EOF = False And rst.BOF = False Then
'                              MsgBox "Got it" & rst!Des_code
                              'Adding to find noof core matched
                              If IsNull(rst.Fields(0)) = False Then
                                MAXCORENO_Find = rst.Fields(0)
                              Else
                                MAXCORENO_Find = 1
                              End If
                        End If
                        Set rst = Nothing
                        
                    For i = 1 To Grid.Rows - 1
                        '++++++++++++++++++++++
                        'Finding specification
                        Set rst = New Recordset
                        'sql = "select * from designspec where des_code = '" & Trim(RT!Des_code) & "' and core_ratio = '" & Trim(Grid.TextMatrix(i, 1)) & "' and core_no = '" & Trim(Grid.TextMatrix(i, 0)) & "' and sp_va = '" & Trim(Grid.TextMatrix(i, 7)) & "' and sp_class = '" & Trim(Grid.TextMatrix(i, 8)) & "' and sp_isf = '" & Trim(Grid.TextMatrix(i, 9)) & "' and sp_alf = '" & Trim(Grid.TextMatrix(i, 10)) & "' and sp_vk = '" & Trim(Grid.TextMatrix(i, 11)) & "' and sp_rct like '" & Trim(Grid.TextMatrix(i, 12)) & "%' and sp_ie = '" & Trim(Grid.TextMatrix(i, 13)) & "' and sp_ie_at = '" & Trim(Grid.TextMatrix(i, 14)) & "'"
                        sql = "select * from designspec where des_code = '" & Trim(RT!Des_code) & "' and core_ratio = '" & Trim(Grid.TextMatrix(i, 1)) & "' and sp_va = '" & Trim(Grid.TextMatrix(i, 7)) & "' and sp_class = '" & Trim(Grid.TextMatrix(i, 8)) & "' and sp_alf = '" & Trim(Grid.TextMatrix(i, 10)) & "' and sp_vk = '" & Trim(Grid.TextMatrix(i, 11)) & "' and sp_rct like '" & Trim(Grid.TextMatrix(i, 12)) & "%' and sp_ie = '" & Trim(Grid.TextMatrix(i, 13)) & "' and sp_ie_at = '" & Trim(Grid.TextMatrix(i, 14)) & "' and core_type = '" & Grid.TextMatrix(i, 19) & "'"
                        rst.Open sql, con, adOpenStatic, adLockReadOnly
                        If rst.EOF = False And rst.BOF = False Then
'                              MsgBox "Got it" & rst!Des_code
                              'Adding to find noof core matched
                              ctr = ctr + 1
                        End If
                        Set rst = Nothing
                    Next
                    
'                End If
                If ctr = Val(txtnoofcores) And MAXCORENO_Find = Val(txtnoofcores) Then
                    searchResultsfrm.lstDno.AddItem RT!Des_code
                    WONO = PreOrder(RT!Des_code)
                    QUONO = QUOTATION(RT!Des_code)
                    modiii = FindRevesion(Trim(RT!drg_no))
                    
                    searchResultsfrm.lstDno.List(NoSearch, 1) = Trim(RT!drg_no) & " (" & modiii & ")" & vbNullString  'this is Add on 03_04_2005 by vishal
                    searchResultsfrm.lstDno.List(NoSearch, 2) = Trim(RT!DRG_TYPE) 'this is Add on 03_04_2005 by vishal
                    If Len(WONO) > 0 Then
                        searchResultsfrm.lstDno.List(NoSearch, 3) = WONO
                        WONO = ""
                    ElseIf Len(QUONO) > 0 Then
                        searchResultsfrm.lstDno.List(NoSearch, 3) = QUONO ' ADD ON 03_04_05
                        WONO = ""
                    End If
                    Design = Design & ", " & RT!Des_code
                    NoSearch = NoSearch + 1
                End If
                ctr = 0
                DesNo = Trim(RT!Des_code)
            End If
            RT.MoveNext
        Loop
    Else
    
        MsgBox "No match found!!!", vbExclamation
        
    End If
    
    If Len(Design) > 0 Then
        searchResultsfrm.Show vbModal
    End If

''Search simillar design No
'
'    searchResultsfrm.lstDno.Clear
'
'    Dim nsvDetails As String
'    Dim DesNo As String
'    Dim Design As String
'    Dim ctr As Integer
'
'    If Val(cmbNSV) <= 11 Then
'        nsvDetails = "3.3 kV','6.6 kV','11 kV"
'    Else
'        nsvDetails = Trim(cmbNSV)
'    End If
'
'
'    Dim RT As Recordset
'    Set RT = New Recordset
'    sql = "select * from designspec where stc = '" & Trim(cmbSTC) & "' and freq = '" & Trim(cmbFreq) & "' and il = '" & Trim(cmbIL) & "' and " _
'          & "time = '" & Trim(cmbTime) & "' and equip = '" & Trim(cmbEquipType) & "' and nsv in ('" & nsvDetails & "') order by des_code"
'    RT.Open sql, con, adOpenKeyset, adLockReadOnly
'    If RT.EOF = False And RT.BOF = False Then
'        RT.MoveFirst
'        Do Until RT.EOF
'            If DesNo <> Trim(RT!Des_code) Then
'                '++++++++++++++++++++++++
'                'Find noof core in the matched design no
'                If Val(txtnoofcores) = maxcore(RT!Des_code) Then
''                    MsgBox "^^^^^^^^^^^^^^^^^^ " & RT!Des_code
'                    For i = 1 To Grid.Rows - 1
'                        '++++++++++++++++++++++
'                        'Finding specification
'                        Set rst = New Recordset
'                        sql = "select * from designspec where des_code = '" & Trim(RT!Des_code) & "' and core_ratio = '" & Trim(Grid.TextMatrix(i, 1)) & "' and core_no = '" & Trim(Grid.TextMatrix(i, 0)) & "' and sp_va = '" & Trim(Grid.TextMatrix(i, 7)) & "' and sp_class = '" & Trim(Grid.TextMatrix(i, 8)) & "' and sp_isf = '" & Trim(Grid.TextMatrix(i, 9)) & "' and sp_alf = '" & Trim(Grid.TextMatrix(i, 10)) & "' and sp_vk = '" & Trim(Grid.TextMatrix(i, 11)) & "' and sp_rct like '" & Trim(Grid.TextMatrix(i, 12)) & "%' and sp_ie = '" & Trim(Grid.TextMatrix(i, 13)) & "' and sp_ie_at = '" & Trim(Grid.TextMatrix(i, 14)) & "'"
'                        rst.Open sql, con, adOpenStatic, adLockReadOnly
'                        If rst.EOF = False And rst.BOF = False Then
'                              MsgBox "Got it" & rst!Des_code
'                              'Adding to find noof core matched
'                              ctr = ctr + 1
'                        End If
'                        Set rst = Nothing
'                    Next
'                End If
'                If ctr = Val(txtnoofcores) Then
'                    searchResultsfrm.lstDno.AddItem RT!Des_code
'                    Design = Design & ", " & RT!Des_code
'                End If
'                ctr = 0
'                DesNo = Trim(RT!Des_code)
'            End If
'            RT.MoveNext
'        Loop
'    End If
'
'    If Len(Design) > 0 Then
'        searchResultsfrm.Show vbModal
'    End If
End Sub

Private Function FindRevesion(Drg As String) As String

    Dim rstD As Recordset
    Set rstD = New Recordset
    rstD.Open "Select mod from drg where drg = '" & Drg & "'", con, adOpenStatic, adLockReadOnly
    If rstD.EOF = False And rstD.BOF = False Then
        FindRevesion = rstD.Fields(0) & vbNullString
'    Else
'        Set rstD = New Recordset
'        rstD.Open "select mod from drg where drg = '" & drg & "' and drgt"
    End If
    Set rstD = Nothing
    
End Function

'Private Sub FillDno()
'
'lstDno.Clear
'Set rst = New Recordset
'rst.Open "select distinct des_code from designspec order by des_code desc", con, adOpenStatic, adLockReadOnly
'If rst.EOF = False And rst.BOF = False Then
'    rst.MoveFirst
''    lstdn
'    Do Until rst.EOF
'
'        rst.MoveNext
'    Loop
'End If
'Set rst = Nothing
'
'End Sub

Private Function maxcore(Des_code As String) As Integer
    Set rst = New Recordset
    rst.Open "select max(core_no) from designspec where des_code = '" & Trim(Des_code) & "'", con, adOpenKeyset, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        maxcore = rst.Fields(0)
    Else
        maxcore = 0
    End If
    Set rst = Nothing
End Function

Private Sub Command1_Click()
Shell App.Path & "\calc.exe"
End Sub

Private Sub cmdVA_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9.]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Form_GotFocus()
    MDIDesign.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub ChangeReadings()

    If ModDesign = "MOD" Then
        If MODTEST = 0 Then
            MODTEST = 1
            STCMemory = MemoStc
            STCTimeMemory = MemoTime
        End If
    End If
    
    If STCMemory <> Trim(cmbSTC) Or STCTimeMemory <> Trim(cmbTime) Then
    If G1.Cols > 1 And InStr(1, UCase(Trim(cmbEquipType)), "WINDOW") = 0 And InStr(1, UCase(Trim(cmbEquipType)), "BAR") = 0 Then
        Ans = MsgBox("All primary calculations will be changed!!!" & vbCrLf & "Are you sure you want to continue?", vbInformation + vbYesNo)
        If Ans = vbNo Then
            cmbSTC = STCMemory
            cmbTime = STCTimeMemory
            Exit Sub
        End If
    Else
        If Len(Trim(txtTurns)) > 0 And InStr(1, UCase(Trim(cmbEquipType)), "WINDOW") = 0 And InStr(1, UCase(Trim(cmbEquipType)), "BAR") = 0 Then
           Ans = MsgBox("All primary calculations will be changed!!!" & vbCrLf & "Are you sure you want to continue?", vbInformation + vbYesNo)
           If Ans = vbNo Then
                cmbSTC = STCMemory
                cmbTime = STCTimeMemory
                Exit Sub
            End If
        Else
            If Len(Trim(txtTotalCon)) > 0 Then
                Ans = MsgBox("All primary calculations will be changed!!!" & vbCrLf & "Are you sure you want to continue?", vbInformation + vbYesNo)
                If Ans = vbNo Then
                     cmbSTC = STCMemory
                     cmbTime = STCTimeMemory
                     Exit Sub
                 End If
            Else
                If InStr(1, UCase(Trim(cmbEquipType)), "WINDOW") > 0 And G1.Cols > 1 Or InStr(1, UCase(Trim(cmbEquipType)), "BAR") > 0 And G1.Cols > 1 Then
                    Ans = MsgBox("Secondary calculations will be changed!!!" & vbCrLf & "Are you sure you want to continue?", vbInformation + vbYesNo)
                    If Ans = vbNo Then
                        cmbSTC = STCMemory
                        cmbTime = STCTimeMemory
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                
            End If
       
        End If
        
    End If
    Else
        Exit Sub
    End If
    
    'Primary
    cmbConStyle.ListIndex = -1
    lstCond.Clear
    txtTurns = ""
    txtminCs = ""
    txtPriCros = ""
    txtTotalCon = ""
    txtCondPer = ""
    txtSide = ""
    txtTop = ""
    txtWidth = ""
    txtDelR = ""
    txtDelStc = ""
    txtIDyn = ""
    txtIDATurns = ""
    
    'Dimension
    wounded.Rows = 1
    txtdrg = ""
    cmbSelctCore = "1"
    cmbType.Clear
    G1.Cols = 1
    Frame21.Visible = False
    Frame22.Visible = False
    Frame16.Visible = False
    Frame14.Visible = False
    Frame15.Visible = False
    FrameWound.Visible = False
    
End Sub

Private Sub FillConductorStyle()
    Dim ctr As Integer
    Dim rst1 As Recordset
    
    cmbConStyle.Clear
    Set rst = New Recordset
    rst.Open "select distinct type from primarycond11 order by type", con, adOpenStatic, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        rst.MoveFirst
        Do Until rst.EOF
            Set rst1 = New Recordset
            rst1.Open "select distinct style from primarycond11 where type = '" & rst!Type & "' order by style", con, adOpenStatic, adLockReadOnly
            If rst1.EOF = False And rst1.BOF = False Then
                Do Until rst1.EOF
                    If rst!Type & vbNullString = "P" Then
                        cmbConStyle.AddItem "Paper"
                    Else
                        cmbConStyle.AddItem "Bare"
                    End If
                    cmbConStyle.List(ctr, 1) = rst1.Fields(0) & vbNullString
                    cmbConStyle.List(ctr, 2) = cmbConStyle.List(ctr, 0) & "-" & cmbConStyle.List(ctr, 1)
                    If cmbConStyle.List(ctr, 0) = "Bare" Then
                        cmbConStyle.List(ctr, 3) = cmbConStyle.List(ctr, 0) & " Conductor -" & cmbConStyle.List(ctr, 1) & " Coil"
                    Else
                        If InStr(cmbConStyle.List(ctr, 1), "Double") > 0 Then
                            cmbConStyle.List(ctr, 3) = cmbConStyle.List(ctr, 0) & " Covered -" & cmbConStyle.List(ctr, 1) & " Wdg. Ht."
                        ElseIf InStr(cmbConStyle.List(ctr, 1), "Series") > 0 Then
                            cmbConStyle.List(ctr, 3) = cmbConStyle.List(ctr, 0) & " Covered -" & cmbConStyle.List(ctr, 1) & " Wdg."
                        Else
                            cmbConStyle.List(ctr, 3) = cmbConStyle.List(ctr, 0) & " Covered -" & cmbConStyle.List(ctr, 1) & " Coil"
                        End If
                    End If
                    ctr = ctr + 1
                    rst1.MoveNext
                Loop
            End If
            rst.MoveNext
        Loop
    End If
    Set rst = Nothing
End Sub

Private Sub Form_Load()

'On Error GoTo err
'G1.Rows = 1
'prcConnect
'prcclear
'cmbEquipType.ListIndex = cmbEquipType.ListCount - 1
cmbEquipType_Change
PcrConductor
FillConductorStyle
'cmbConStyle.Clear
'cmbConStyle.AddItem "Double Coil - Bare Conductor"
'cmbConStyle.AddItem "Double Coil - Normal Wdg Ht."
'cmbConStyle.AddItem "Series Coil"

cmbCore.Text = "All Cores"
cmbCType.ListIndex = -1
cmbCType.Text = "-Select Core Type-"
'CMBINSULATION
''**** if all the core details are present then blank
''**** ratio on specification
For l = 1 To Grid.Rows - 1
    vinn = 0
    If l = cmbnoofcores Then
        vinn = 1
    End If
Next
If vinn = 1 Then
    txtRatio2.Text = ""
End If

    Frame7.Enabled = False
    Frame6.Enabled = False
    
cmbHSV = ""
prcclear
prcClearALL
''++++++++++++++++++++++++++++++++++++++++++++++
''Modification
If ModDesign = "MOD" Then
    
    Dim Rep As ModifyDes
    Set Rep = New ModifyDes
        Rep.ModifingDB (DesignNo)
    Set Modf = Nothing
    
    txtDesNo = DesignNo
    ModCal
'    SSTab1.Tab = 4
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Setting Ratio
    'No of Core Text
    Dim rstT As Recordset
    Set rstT = New Recordset
    rstT.Open "select drg_no, drg_type from designspec where des_code = '" & DesignNo & "' ", con, adOpenStatic, adLockReadOnly
    If rstT.EOF = False And rstT.BOF = False Then
    '    txtdrg = ""
        cmbType = rstT.Fields(1)
        txtdrg = rstT.Fields(0)
        
    End If
    Set rstT = Nothing
    'End Setting Ratio g1
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++
    NonStad = ""
End If

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & " " & err.Number
    End If

End Sub

Private Sub DisApp(Status As Boolean)

txtCondPer.Enabled = Status
txtTotalCon.Enabled = Status
txtSide.Enabled = Status
txtTop.Enabled = Status
txtWidth.Enabled = Status
lstCond.Enabled = Status

End Sub


Private Function Trig(Burden As Double, SecCurr As Double, Rec75 As Double)

Dim CosLamda, SinLamda, TanDelta, CosDelta, SinDelta As Double

If Burden > 5 Then
    CosLamda = 0.8
Else
    CosLamda = 1
End If

SinLamda = Sqr(1 - (CosLamda * CosLamda))

TanDelta = (Burden * SinLamda) / (SecCurr * SecCurr * Rec75 + Burden * CosLamda)

SinDelta = TanDelta / Sqr(1 + (TanDelta * TanDelta))
sinD = SinDelta

CosDelta = Sqr(1 - SinDelta * SinDelta)
cosD = CosDelta

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If CancelCT <> "1" Then
    joel = MsgBox("Are you sure you want to exit?", vbInformation + vbYesNo)
    If joel = vbNo Then
        Cancel = 1
    Else
        Set CTfrm = Nothing
    End If
Else
     
    CancelCT = ""
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

For i = 0 To 4
    RatioWatch(i) = ""
Next
Set CTfrm = Nothing

MDIDesign.StatusBar1.Panels(5).Text = "CT Design"

MDIDesign.mnuNDesi.Enabled = True
MDIDesign.Toolbar1.Buttons(1).Enabled = True
MDIDesign.StatusBar1.Panels(2).Text = ""
ModDesign = ""
DesignNo = ""
End Sub




Private Sub GCore_DblClick()
If GCore.Rows = 1 Then
    Exit Sub
End If

If GCore.Tag = "MOD" Then

    txtOL = GCore.TextMatrix(GCore.row, 1)
    txtOW = GCore.TextMatrix(GCore.row, 2)
    txtWL = GCore.TextMatrix(GCore.row, 3)
    txtWW = GCore.TextMatrix(GCore.row, 4)
    
    For i = 0 To GCore.Cols - 1
        GCore.TextMatrix(GCore.RowSel, i) = "Mod..."
    Next

End If

End Sub

Private Sub GCross_Click()
    
    txtGEdit.Top = GCross.Top + GCross.CellTop
    txtGEdit.Left = GCross.Left + GCross.CellLeft
    txtGEdit.Width = GCross.CellWidth
    txtGEdit.Height = GCross.CellHeight
    txtGEdit.Text = GCross.TextMatrix(GCross.RowSel, GCross.ColSel)
    txtGEdit.Visible = True
    txtGEdit.SetFocus
    
End Sub

Private Sub Grid3_DblClick()

If Grid3.Rows = 1 Then Exit Sub

txtWL = ""
txtWW = ""
txtOL = ""
txtOW = ""

If GCore.row - 1 = Val(txtnoofcores) Then
    Exit Sub
End If

If Len(cmbCore) = 0 Then
    MsgBox "Please select core no", vbInformation
End If

If GCore.Rows - 1 = Val(txtnoofcores) Then
    Exit Sub
End If

For j = 1 To GCore.Rows - 1
    If Trim(cmbCore) = Trim(GCore.TextMatrix(j, 0)) Then
        MsgBox "This core already exits", vbInformation
        Exit Sub
    End If
Next

Dim Cout As String
Dim CNo As Integer
CNo = 1
For K = 1 To txtnoofcores
    Cout = ""
    For j = 1 To GCore.Rows - 1
        If GCore.TextMatrix(j, 0) = K Then
            Cout = "K"
            Exit For
        End If
    Next
    If Cout <> "K" Then
        Exit For
    End If
Next

If Cout = "" Then
    CNo = K
End If

'txtOL.Text = Grid3.TextMatrix(Grid3.RowSel, 3)
'txtOW.Text = Grid3.TextMatrix(Grid3.RowSel, 4)
'txtWL.Text = Grid3.TextMatrix(Grid3.RowSel, 1)
'txtWW.Text = Grid3.TextMatrix(Grid3.RowSel, 2)


Dim Mean, BU As Double


BU = prcCalBU(Grid3.TextMatrix(Grid3.RowSel, 3), Grid3.TextMatrix(Grid3.RowSel, 1))
txtBUP = BU
Mean = prcCalMean(Grid3.TextMatrix(Grid3.RowSel, 1), Grid3.TextMatrix(Grid3.RowSel, 2), BU)
txtMean.Text = Mean

'If GCore.Rows > Grid2.Rows Then
'    cmbSelctCore = GCore.TextMatrix(Grid2.Rows, 0)
'End If
If OptCom(0).value = True Then
        
    For i = 1 To Val(txtnoofcores)
        GCore.Rows = GCore.Rows + 1
        GCore.row = GCore.Rows - 1
        GCore.RowHeight(GCore.row) = 300
        GCore.TextMatrix(GCore.row, 0) = GCore.row
        GCore.TextMatrix(GCore.row, 1) = Grid3.TextMatrix(Grid3.RowSel, 3)
        GCore.TextMatrix(GCore.row, 2) = Grid3.TextMatrix(Grid3.RowSel, 4)
        GCore.TextMatrix(GCore.row, 3) = Grid3.TextMatrix(Grid3.RowSel, 1)
        GCore.TextMatrix(GCore.row, 4) = Grid3.TextMatrix(Grid3.RowSel, 2)
        GCore.TextMatrix(GCore.row, 8) = Grid3.TextMatrix(Grid3.RowSel, 3) & " x " & Grid3.TextMatrix(Grid3.RowSel, 4)
        GCore.TextMatrix(GCore.row, 9) = Grid3.TextMatrix(Grid3.RowSel, 1) & " x " & Grid3.TextMatrix(Grid3.RowSel, 2)
        GCore.TextMatrix(GCore.row, 10) = txtMean
        GCore.TextMatrix(GCore.row, 11) = txtBUP
    Next
Else
    GCore.Rows = GCore.Rows + 1
    GCore.row = GCore.Rows - 1
    GCore.RowHeight(GCore.row) = 300
    GCore.TextMatrix(GCore.row, 0) = CNo
    GCore.TextMatrix(GCore.row, 1) = Grid3.TextMatrix(Grid3.RowSel, 3)
    GCore.TextMatrix(GCore.row, 2) = Grid3.TextMatrix(Grid3.RowSel, 4)
    GCore.TextMatrix(GCore.row, 3) = Grid3.TextMatrix(Grid3.RowSel, 1)
    GCore.TextMatrix(GCore.row, 4) = Grid3.TextMatrix(Grid3.RowSel, 2)
    GCore.TextMatrix(GCore.row, 8) = Grid3.TextMatrix(Grid3.RowSel, 3) & " x " & Grid3.TextMatrix(Grid3.RowSel, 4)
    GCore.TextMatrix(GCore.row, 9) = Grid3.TextMatrix(Grid3.RowSel, 1) & " x " & Grid3.TextMatrix(Grid3.RowSel, 2)
    GCore.TextMatrix(GCore.row, 10) = txtMean
    GCore.TextMatrix(GCore.row, 11) = txtBUP
End If
    
        If cmbCore.ListIndex < cmbCore.ListCount - 1 Then
            cmbCore.ListIndex = cmbCore.ListIndex + 1
        End If
If OptCom(0).value = True Then
    Frame11.Caption = "For all Cores : "
    Frame10.Caption = "For all Cores : "
Else
    Frame11.Caption = "Size For " & cmbCore & " : "
    Frame10.Caption = "Size For " & cmbCore & " : "
End If
Frame1.Enabled = False

CNo = 1
For K = 1 To txtnoofcores
    Cout = ""
    For j = 1 To GCore.Rows - 1
        If GCore.TextMatrix(j, 0) = K Then
            Cout = "K"
            Exit For
        End If
    Next
    If Cout <> "K" Then
        Exit For
    End If
Next

If Cout = "" Then
    cmbCore = "Core " & K
End If
If txtOL.Visible = True Then
    txtOL.SetFocus
Else
    txtID.SetFocus
End If
'If GCore.Rows > Grid2.Rows Then
'    cmbSelctCore = Val(GCore.TextMatrix(Grid2.Rows, 0))
'End If
'txtBUP = ""
'txtMean = ""
End Sub


Private Sub GridRatio_Click()

Frame2.Caption = "Core " & GridRatio.TextMatrix(GridRatio.RowSel, 0) & " Shape and Ratio : "
Frame6.Caption = "Core " & GridRatio.TextMatrix(GridRatio.RowSel, 0) & " Inputs : "

txtRatio2 = GridRatio.TextMatrix(GridRatio.RowSel, 1)
txtRatio2R2 = GridRatio.TextMatrix(GridRatio.RowSel, 2)

End Sub

Private Sub GridRatio_EnterCell()
Frame2.Caption = "Core " & GridRatio.TextMatrix(GridRatio.RowSel, 0) & " Shape and Ratio : "
Frame6.Caption = "Core " & GridRatio.TextMatrix(GridRatio.RowSel, 0) & " Inputs : "

End Sub


Private Sub IvAcc_Change()
Frame6 = cmbCType & "(Core " & cmbnoofcores & ") Ratio : " & IvAcc & "A"
End Sub

Private Sub IvAcc_Click()
Frame6 = cmbCType & "(Core " & cmbnoofcores & ") Ratio : " & IvAcc & "A"
End Sub

Private Sub IvAcc_GotFocus()
'IvAcc.SelStart = 0
'IvAcc.SelLength = Len(IvAcc)
End Sub

Private Sub lstCond_Click()

On Error GoTo err
    txtCondPer = ""
    txtTotalCon = ""
    LstValue = FindIndexConduct
    If LstValue = 1000 Then
        lstCond.ListIndex = -1
        MsgBox "Suitable primary section not found!!!", vbExclamation
        Exit Sub
    End If

    Dim str1 As String
    Dim j As String * 1
    If Len(lstCond) = 0 Then
        Exit Sub
    End If

        Dim value As Double

        'Calculate minimum primary cross section area at normal or continus current
        
        For i = 0 To Val(txtnoofcores) - 1
            If InStr(1, txtRatio(i), "-") = 0 Then
                If value < txtRatio(i) Then
                    value = txtRatio(i)
                End If
            Else
                For K = 1 To Len(txtRatio(i)) + 1
                    j = Mid(txtRatio(i), K, K)
                    If j = "-" Or K > Len(txtRatio(i)) Then
                        If value < CDbl(str1) Then
                            value = CDbl(str1)
                        End If
                        str1 = ""
                    Else
                        str1 = str1 & j
                    End If
                Next
            End If
        Next
        
        If Len(Trim(txtRateConti)) > 0 Then
                value = CDbl(txtRateConti)
        End If
        
        txtDelR = prcCalDElRate(CDbl(lstCond), value)
        If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
            txtDelStc = prcCalDElSTC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), CDbl(cmbTime), lstCond)
        End If
        txtPriCros = lstCond
    cmbConStyle.Enabled = True
    If cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Bare" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double" Then
            txtTurns.Clear
    
            ''Filling Combo
            Set rstCon = New Recordset
            rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
                rstCon.MoveFirst
                Do Until rstCon.EOF
                    txtTurns.AddItem rstCon.Fields(0) & vbNullString
                    rstCon.MoveNext
                Loop
                
                txtTurns.ListIndex = -1
                ''Details of primary
                Set rstCon = New Recordset
                rstCon.Open "select pri_wind_width, t_np, t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
                If rstCon.EOF = False And rstCon.BOF = False Then
    '                txtTurns = rstCon!t_np & vbNullString
                    txtWidth = rstCon!pri_wind_width & vbNullString
    '                txtPriCros = rstCon!t_csp & vbNullString
'                    txtTop = rstCon!t_whpt & vbNullString
'                    txtSide = rstCon!t_whps & vbNullString
                    txtCondPer = rstCon!t_COND_coi & vbNullString
                    txtTotalCon = rstCon!t_cond_siz & vbNullString
                    txtTop = "0"
                    txtSide = "0"
                Else
                    txtTop = "0"
                    txtSide = "0"
                    txtCondPer = "-"
                    txtTotalCon = "-"
                    txtWidth = "0"
                    MsgBox "Primary data not avaiable!!!", vbInformation
                End If
                Set rstCon = Nothing
            Else
                MsgBox "Primary data not avaiable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Bare" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Single" Then
            txtTurns.Clear
    
            ''Filling Combo
            Set rstCon = New Recordset
            rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
                rstCon.MoveFirst
                Do Until rstCon.EOF
                    txtTurns.AddItem rstCon.Fields(0) & vbNullString
                    rstCon.MoveNext
                Loop
                
                If rstCon.RecordCount > 1 Then
                    txtTurns.ListIndex = -1
                Else
                    txtTurns.ListIndex = 0
                End If
                
                ''Details of primary
                Set rstCon = New Recordset
                rstCon.Open "select pri_wind_width, t_np, t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
                If rstCon.EOF = False And rstCon.BOF = False Then
    '                txtTurns = rstCon!t_np & vbNullString
                    txtWidth = rstCon!pri_wind_width & vbNullString
    '                txtPriCros = rstCon!t_csp & vbNullString
'                    txtTop = rstCon!t_whpt & vbNullString
'                    txtSide = rstCon!t_whps & vbNullString
                    txtCondPer = rstCon!t_COND_coi & vbNullString
                    txtTotalCon = rstCon!t_cond_siz & vbNullString
                    txtTop = "0"
                    txtSide = "0"
                Else
                    txtTop = "0"
                    txtSide = "0"
                    txtCondPer = "-"
                    txtTotalCon = "-"
                    txtWidth = "0"
                    MsgBox "Primary data not avaiable!!!", vbInformation
                End If
                Set rstCon = Nothing
            Else
                MsgBox "Primary data not avaiable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double-Different" Then
            txtTurns.Clear
    
            ''Filling Combo
            Set rstCon = New Recordset
            rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
                rstCon.MoveFirst
                Do Until rstCon.EOF
                    txtTurns.AddItem rstCon.Fields(0) & vbNullString
                    rstCon.MoveNext
                Loop
                
                txtTurns.ListIndex = -1
                ''Details of primary
                Set rstCon = New Recordset
                rstCon.Open "select pri_wind_width, t_np, t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
                If rstCon.EOF = False And rstCon.BOF = False Then
    '                txtTurns = rstCon!t_np & vbNullString
                    txtWidth = rstCon!pri_wind_width & vbNullString
    '                txtPriCros = rstCon!t_csp & vbNullString
'                    txtTop = rstCon!t_whpt & vbNullString
'                    txtSide = rstCon!t_whps & vbNullString
'                    txtCondPer = rstCon!t_COND_coi & vbNullString
'                    txtTotalCon = rstCon!t_cond_siz & vbNullString
                    txtTop = "0"
                    txtSide = "0"
                Else
                    txtTop = "0"
                    txtSide = "0"
                    txtCondPer = "-"
                    txtTotalCon = "-"
                    txtWidth = "0"
                    MsgBox "Primary data not avaiable!!!", vbInformation
                End If
                Set rstCon = Nothing
            Else
                MsgBox "Primary data not avaiable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double-Equal" Then
        txtTurns.Clear
    
            ''Filling Combo
            Set rstCon = New Recordset
            rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
                rstCon.MoveFirst
                Do Until rstCon.EOF
                    txtTurns.AddItem rstCon.Fields(0) & vbNullString
                    rstCon.MoveNext
                Loop
                
                txtTurns.ListIndex = -1
                ''Details of primary
                Set rstCon = New Recordset
                rstCon.Open "select pri_wind_width, t_np, t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
                If rstCon.EOF = False And rstCon.BOF = False Then
    '                txtTurns = rstCon!t_np & vbNullString
                    txtWidth = rstCon!pri_wind_width & vbNullString
    '                txtPriCros = rstCon!t_csp & vbNullString
'                    txtTop = rstCon!t_whpt & vbNullString
'                    txtSide = rstCon!t_whps & vbNullString
'                    txtCondPer = rstCon!t_COND_coi & vbNullString
'                    txtTotalCon = rstCon!t_cond_siz & vbNullString
                    txtTop = "0"
                    txtSide = "0"
                Else
                    txtTop = "0"
                    txtSide = "0"
                    txtCondPer = "-"
                    txtTotalCon = "-"
                    txtWidth = "0"
                    MsgBox "Primary data not avaiable!!!", vbInformation
                End If
                Set rstCon = Nothing
            Else
                MsgBox "Primary data not avaiable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Series" Then
        txtTurns.Clear
    
            ''Filling Combo
            Set rstCon = New Recordset
            rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = '" & Trim(txtPriCros) & "'", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
                rstCon.MoveFirst
                Do Until rstCon.EOF
                    txtTurns.AddItem rstCon.Fields(0) & vbNullString
                    rstCon.MoveNext
                Loop
                
                txtTurns.ListIndex = -1
                ''Details of primary
                Set rstCon = New Recordset
                rstCon.Open "select pri_wind_width, t_np, t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = '" & Trim(txtPriCros) & "'", con, adOpenStatic, adLockReadOnly
                If rstCon.EOF = False And rstCon.BOF = False Then
    '                txtTurns = rstCon!t_np & vbNullString
                    txtWidth = rstCon!pri_wind_width & vbNullString
    '                txtPriCros = rstCon!t_csp & vbNullString
'                    txtTop = rstCon!t_whpt & vbNullString
'                    txtSide = rstCon!t_whps & vbNullString
'                    txtCondPer = rstCon!t_COND_coi & vbNullString
'                    txtTotalCon = rstCon!t_cond_siz & vbNullString
                    txtTop = "0"
                    txtSide = "0"
                Else
                    txtTop = "0"
                    txtSide = "0"
                    txtCondPer = "-"
                    txtTotalCon = "-"
                    txtWidth = "0"
                    MsgBox "Primary data not avaiable!!!", vbInformation
                End If
                Set rstCon = Nothing
            Else
                MsgBox "Primary data not avaiable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Single" Then
        txtTurns.Clear
    
            ''Filling Combo
            Set rstCon = New Recordset
            rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
                rstCon.MoveFirst
                Do Until rstCon.EOF
                    txtTurns.AddItem rstCon.Fields(0) & vbNullString
                    rstCon.MoveNext
                Loop
                
                txtTurns.ListIndex = -1
                ''Details of primary
                Set rstCon = New Recordset
                rstCon.Open "select pri_wind_width, t_np, t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
                If rstCon.EOF = False And rstCon.BOF = False Then
    '                txtTurns = rstCon!t_np & vbNullString
                    txtWidth = rstCon!pri_wind_width & vbNullString
    '                txtPriCros = rstCon!t_csp & vbNullString
'                    txtTop = rstCon!t_whpt & vbNullString
'                    txtSide = rstCon!t_whps & vbNullString
'                    txtCondPer = rstCon!t_COND_coi & vbNullString
'                    txtTotalCon = rstCon!t_cond_siz & vbNullString
                    txtTop = "0"
                    txtSide = "0"
                Else
                    txtTop = "0"
                    txtSide = "0"
                    txtCondPer = "-"
                    txtTotalCon = "-"
                    txtWidth = "0"
                    MsgBox "Primary data not avaiable!!!", vbInformation
                End If
                Set rstCon = Nothing
            Else
                MsgBox "Primary data not avaiable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    End If
        
If Len(Trim(txtDelR)) > 0 Then
    If Val(txtDelR) > 2 Then
        MsgBox "Nominal current density cannot be greater than 2!!!", vbExclamation
        txtTurns.SetFocus
        Exit Sub
    End If
End If

If Len(Trim(txtDelStc)) > 0 Then
    If Val(txtDelStc) > 180 Then
'        MsgBox "STC current density cannot be greater than 180!!!", vbExclamation
        txtTurns.SetFocus
        Exit Sub
    End If
End If


    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select * from primarycond11 where t_csp = " & lstCond.Text & "", con, adOpenKeyset, adLockReadOnly
    If rs.EOF = False And rs.BOF = False Then
        If NonStad = "" Then
            If InStr(Trim(rs!t_cond_siz), "12.5") > 0 Then
                MsgBox "This is a NON-STANDARD conductor!!!", vbInformation
            End If
        End If
    End If
    Set rs = Nothing
err:
    
    If Trim(Len(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number
    End If

End Sub

'Private Sub lstCond_DblClick()
''txtPriCros = lstCond
'End Sub

Private Sub OptCom_Click(Index As Integer)

If OptCom(0).value = True Then
    Frame11.Caption = "For all Cores : "
    Frame10.Caption = "For all Cores : "
    cmbCore.Enabled = False
    cmbCore.Text = "All Cores"
Else
    Frame11.Caption = "Size For " & cmbCore & " : "
    Frame10.Caption = "Size For " & cmbCore & " : "
    cmbCore.Text = "Core 1"
    cmbCore.Enabled = True
End If
End Sub

Private Sub OptCom_GotFocus(Index As Integer)
MDIDesign.StatusBar1.Panels(2).Text = "Option to set same or different dimension"
End Sub

Private Sub OptComm_Click()
IvAcc.Clear
If Grid.row = 0 Then
    Exit Sub
End If

For i = 1 To Grid.Rows - 1
    
    If Grid.TextMatrix(i, 18) = "" Then
    
        Frame6 = cmbCType & "(Core " & cmbnoofcores & ") : "
        txtinv = ""
        IvAcc.Enabled = False
        Exit Sub
    
    End If
    
    If Grid.TextMatrix(i, 18) >= 0 And Grid.TextMatrix(i, 1) = Trim(IvAcc) Then
    
        OptComm.value = False
        OptInd.value = True
        MsgBox "To be tested ", vbInformation
        Exit Sub
    
    End If

Next

If cmbCType = "Metering" Then
    Frame6.Caption = "Metering (Core " & cmbnoofcores & ") :"
ElseIf cmbCType = "Protection" Then
    Frame6.Caption = "Protection (Core " & cmbnoofcores & ") :"
ElseIf cmbCType = "Special Protection" Then
    Frame6.Caption = "Special Protection (Core " & cmbnoofcores & ") :"
End If

IvAcc.Clear
IvAcc.Enabled = False

End Sub

Private Sub OptInd_Click()

Dim value As String * 1
Dim str1 As String
Dim int11 As Integer
'MsgBox txtRatio(cmbnoofcores - 1)
If Len(cmbnoofcores) = 0 Then
    Exit Sub
End If
If InStr(1, txtRatio(cmbnoofcores - 1), "-") > 0 Then
    IvAcc.Clear
    For i = 1 To Len(txtRatio(cmbnoofcores - 1)) + 1
        value = Mid(txtRatio(cmbnoofcores - 1), i, i)
        If value = "-" Or Len(txtRatio(cmbnoofcores - 1)) < i Then
            
            IvAcc.AddItem str1 & "/" & txtRatioR2(cmbnoofcores - 1)
            str1 = ""
            
        Else
        
            str1 = str1 + value
        
        End If
    Next
    IvAcc.ListIndex = 0
End If

If IvAcc.ListCount > 0 Then
    If InStr(1, IvAcc.List(0), "/") <= 0 Then
        X = IvAcc.List(0)
    Else
        X = Left(IvAcc.List(0), InStr(1, IvAcc.List(0), "/") - 1)
    End If
    For i = 0 To IvAcc.ListCount - 1
        If i = IvAcc.ListCount - 1 Then
            Exit For
        End If
        If InStr(1, IvAcc.List(i + 1), "/") <= 0 Then
            y = IvAcc.List(i + 1)
        Else
            y = Left(IvAcc.List(i + 1), InStr(1, IvAcc.List(i + 1), "/") - 1)
        End If
'        Y = CDbl(IvAcc.List(i + 1))
        If X > y Then
            OptInd.Tag = i
            X = y
            
        End If
        
    Next
    
End If
Frame4.Enabled = True
txtinv = X
Frame6 = cmbCType & " (Core " & cmbnoofcores & ") Ratio : " & IvAcc & "A"
'IvAcc.ListIndex = 0
IvAcc.Enabled = True

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
STCMemory = cmbSTC
STCTimeMemory = cmbTime
'On Error GoTo err
If SSTab1.Tab = 2 Then
    If InStr(1, UCase(cmbEquipType), "BAR") > 0 Or InStr(1, UCase(cmbEquipType), "WINDOW") > 0 Then
        If PreviousTab <= 1 Then
            If Val(txtnoofcores) = 0 Then
                SSTab1.Tab = 0
                Exit Sub
            ElseIf Val(txtnoofcores) > Grid.Rows - 1 Then
                SSTab1.Tab = 0
                Exit Sub
            End If
            If (GCore.Rows - 1) < Val(txtnoofcores) Then
                SSTab1.Tab = 1
            Else
                SSTab1.Tab = 3
            End If
            Exit Sub
        ElseIf PreviousTab >= 3 Then
            If GCore.Rows - 1 < Val(txtnoofcores) Then
                SSTab1.Tab = 1
            Else
                SSTab1.Tab = 1
            End If
            Exit Sub
        End If
        Exit Sub
    End If
End If

If SSTab1.Tab = 4 Then
    If prcFINDCoreNo > 0 Then
        If Val(txtnoofcores) = 0 Then
            SSTab1.Tab = 0
            Exit Sub
        ElseIf Val(txtnoofcores) > Grid.Rows - 1 Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        SSTab1.Tab = 3
'        MsgBox "Please generate cores details", vbExclamation
    Else
        txtdrg.SetFocus
    End If
End If

''PREVOIUS TAB = 0 CHECK NOOF CORE AND CORE DETAILS
If PreviousTab = 0 Then
    If Val(txtnoofcores) = 0 Then
        SSTab1.Tab = 0
        Exit Sub
    ElseIf Grid.Rows = 1 Then
        SSTab1.Tab = 0
        Exit Sub
    End If
End If
STCMemory = cmbSTC

If SSTab1.Tab = 1 Then
    If cmbStyle.Text <> "Rectangluar" Then
        Grid3.Rows = 1
        txtOD.SetFocus
    Else
        txtOL.SetFocus
    End If
    
End If

Grid.Tag = ""
If PreviousTab = 0 Then
    For i = 1 To CInt(txtnoofcores)
        For j = 1 To Grid.Rows - 1
            If Grid.TextMatrix(j, 0) = i Then
                Grid.Tag = "Y"
            End If
        Next
        If Grid.Tag = "Y" Then
           Grid.Tag = ""
        Else
            MDIDesign.StatusBar1.Panels(2).Text = "Please enter specification of all cores"
            SSTab1.Tab = 0
            STCMemory = Trim(cmbSTC)
            Exit Sub
        End If
        
    Next
    If Len(cmbSTC) = 0 Or Trim(UCase(cmbSTC)) = "KA" Then
         MDIDesign.StatusBar1.Panels(2).Text = "Please select STC value"
        SSTab1.Tab = 0
        cmbSTC.SetFocus
        Exit Sub
    End If
    If Len(cmbTime) = 0 Then
        MDIDesign.StatusBar1.Panels(2).Text = "Please select STC Time"
        SSTab1.Tab = 0
        cmbTime.SetFocus
        Exit Sub
    End If
    PrcRatioCheck
    Dim AscRatio As ChangeRa
    Set AscRatio = New ChangeRa
    
    For p = 0 To txtnoofcores - 1
        txtRatio(p) = AscRatio.prcInterChangeRatio(txtRatioDum(p))
    Next
    
ElseIf PreviousTab = 1 Then
    'Allows changes in core details without enter dcore dimenstion
    If SSTab1.Tab = 0 Then
    
    ElseIf SSTab1.Tab = 2 Then
        If Len(txtTurns) > 0 Or Val(txtTurns) <> 0 Then
            Dim ss As String
            ss = txtTurns
            txtTurns = ""
            txtTurns = ss
        End If
    Else
    'Check for core dimenstion
        For i = 1 To CInt(txtnoofcores)
            For j = 1 To Grid.Rows - 1
                If Grid.TextMatrix(j, 0) = i Then
                    Grid.Tag = "Y"
                End If
            Next
            If Grid.Tag = "Y" Then
               Grid.Tag = ""
            Else
                MDIDesign.StatusBar1.Panels(2).Text = "Please enter specification of all cores"
                SSTab1.Tab = 0
                Exit Sub
            End If
        Next
        If Len(cmbSTC) = 0 Then
            MDIDesign.StatusBar1.Panels(2).Text = "Please select STC value"
            SSTab1.Tab = 0
            cmbSTC.SetFocus
            Exit Sub
        End If
        If Len(cmbTime) = 0 Then
            MDIDesign.StatusBar1.Panels(2).Text = "Please select STC Time"
            SSTab1.Tab = 0
            cmbTime.SetFocus
            Exit Sub
        End If
        If GCore.Rows - 1 < txtnoofcores Then
            MDIDesign.StatusBar1.Panels(2).Text = "Please enter all core sizes"
            SSTab1.Tab = 1
            Exit Sub
        End If
    End If
ElseIf PreviousTab = 3 Then
    If G1.Cols = 1 Then
        Exit Sub
    End If
    Dim GreaterOD, LowerID, GreaterOD1, LowerID1 As Double
    If SSTab1.Tab = 4 Then
        wounded.Rows = 1
        Dim GFormer As Double
        If G1.TextMatrix(48, 1) = "Rectangluar" Then
            GreaterOD = G1.TextMatrix(75, 1)
            GreaterOD1 = G1.TextMatrix(76, 1)
            LowerID = G1.TextMatrix(78, 1)
            LowerID1 = G1.TextMatrix(77, 1)
            wounded.AllowUserResizing = flexResizeBoth
            wounded.WordWrap = True
            Dim s As String
            wounded.ColAlignment(0) = vbCenter
            wounded.ColWidth(0) = 1600
            s = "OverAll (mm)" & vbCrLf & "Width X Heigth"
            wounded.TextMatrix(0, 1) = s
            wounded.ColAlignment(1) = vbCenter
            s = "Window (mm)" & vbCrLf & "Width X Heigth"
            wounded.ColWidth(1) = 1100
            wounded.TextMatrix(0, 2) = s
            wounded.ColAlignment(2) = vbCenter
            wounded.ColWidth(2) = 1100
            wounded.ColAlignment(3) = vbCenter
            wounded.ColWidth(3) = 1100
            wounded.ColAlignment(4) = vbCenter
            wounded.ColWidth(4) = 1100
            wounded.RowHeight(0) = 500
            
            For i = 1 To G1.Cols - 1
                If CDbl(G1.TextMatrix(75, i)) > CDbl(GreaterOD) Then
                    GreaterOD = G1.TextMatrix(75, i)
                    GreaterOD1 = G1.TextMatrix(76, i)
                End If
                If CDbl(G1.TextMatrix(78, i)) < CDbl(LowerID) Then
                    LowerID = G1.TextMatrix(78, i)
                    LowerID1 = G1.TextMatrix(77, i)
                End If
                wounded.Rows = wounded.Rows + 1
                wounded.row = wounded.Rows - 1
                wounded.TextMatrix(wounded.row, 0) = G1.TextMatrix(1, i)
                wounded.TextMatrix(wounded.row, 1) = Round(G1.TextMatrix(75, i), 0) & " X " & Round(G1.TextMatrix(76, i), 0)
                wounded.TextMatrix(wounded.row, 2) = Round(G1.TextMatrix(77, i), 0) & " X " & Round(G1.TextMatrix(78, i), 0)
                wounded.TextMatrix(wounded.row, 3) = Round(G1.TextMatrix(79, i), 0)
                WHt = RoundUp(G1.TextMatrix(71, i))
                wounded.TextMatrix(wounded.row, 4) = G1.TextMatrix(26, i) + (2 * G1.TextMatrix(71, i))
                wounded.TextMatrix(wounded.row, 4) = Round(wounded.TextMatrix(wounded.row, 4), 0)
                'BU + 2 * WindHt
            Next
            prcDeleteDup
            '********************Final Dimension
                wounded.Rows = wounded.Rows + 1
            wounded.row = wounded.Rows - 1
            wounded.TextMatrix(wounded.row, 0) = "Max-Dim (mm)"
            wounded.TextMatrix(wounded.row, 1) = Round(GreaterOD, 0) & " X " & Round(GreaterOD1, 0)
            OALLWin = GreaterOD1
            OALLLen = GreaterOD
            OALLWin1 = LowerID1
            OALLLen1 = LowerID
            wounded.TextMatrix(wounded.row, 2) = Round(LowerID1, 0) & " X " & Round(LowerID, 0)
            wounded.TextMatrix(wounded.row, 3) = 0
            wounded.TextMatrix(wounded.row, 4) = wounded.TextMatrix(1, 4)
            For l = 1 To wounded.row - 1
                wounded.TextMatrix(wounded.row, 3) = CDbl(wounded.TextMatrix(wounded.row, 3)) + CDbl(wounded.TextMatrix(l, 3))
                If CDbl(wounded.TextMatrix(wounded.row, 4)) < CDbl(wounded.TextMatrix(l, 4)) Then
                    wounded.TextMatrix(wounded.row, 4) = wounded.TextMatrix(l, 4)
                End If
            Next
            wounded.TextMatrix(wounded.row, 3) = wounded.TextMatrix(wounded.row, 3) + (2 * ((wounded.Rows - 2) - 1)) - (6 * ((wounded.Rows - 2) - 1))
            wounded.TextMatrix(wounded.row, 3) = Round(wounded.TextMatrix(wounded.row, 3), 2)
        Else
            wounded.Rows = 1
            wounded.ColWidth(0) = 1600
            wounded.ColAlignment(0) = vbCenter
            wounded.TextMatrix(0, 1) = "OD (mm)"
            wounded.ColAlignment(1) = vbCenter
            wounded.ColWidth(1) = 1100
            wounded.TextMatrix(0, 2) = "ID (mm)"
            wounded.ColAlignment(2) = vbCenter
            wounded.ColWidth(2) = 1100
            wounded.ColWidth(3) = 1100
            wounded.ColAlignment(3) = vbCenter
            wounded.ColWidth(4) = 1100
            wounded.ColAlignment(4) = vbCenter
            GreaterOD = G1.TextMatrix(73, 1)
            LowerID = G1.TextMatrix(74, 1)
            
            For i = 1 To G1.Cols - 1
                If CDbl(G1.TextMatrix(73, i)) > CDbl(GreaterOD) Then
                    GreaterOD = G1.TextMatrix(73, i)
                End If
                If Len(Trim(G1.TextMatrix(74, i))) > 0 Then
                    If CDbl(G1.TextMatrix(74, i)) < CDbl(LowerID) Then
                        LowerID = G1.TextMatrix(74, i)
                    End If
                Else
                    LowerID = G1.TextMatrix(74, i)
                End If
                wounded.Rows = wounded.Rows + 1
                wounded.row = wounded.Rows - 1
                wounded.TextMatrix(wounded.row, 0) = G1.TextMatrix(1, i)
                wounded.TextMatrix(wounded.row, 1) = Round(G1.TextMatrix(73, i), 0)
                wounded.TextMatrix(wounded.row, 2) = Round(G1.TextMatrix(74, i), 0)
                WHt = RoundUp(G1.TextMatrix(71, i))
                wounded.TextMatrix(wounded.row, 3) = Round(G1.TextMatrix(79, i), 0)
                wounded.TextMatrix(wounded.row, 4) = G1.TextMatrix(26, i) + (2 * G1.TextMatrix(71, i))
                wounded.TextMatrix(wounded.row, 4) = Round(wounded.TextMatrix(wounded.row, 4), 0)
            Next
            prcDeleteDup
            '**********************Final Dimension
            wounded.Rows = wounded.Rows + 1
            wounded.row = wounded.Rows - 1
            wounded.TextMatrix(wounded.row, 0) = "Max. Dim (mm)"
            wounded.TextMatrix(wounded.row, 1) = Round(GreaterOD, 0)
            wounded.TextMatrix(wounded.row, 2) = Round(LowerID, 0)
            OALLWin = GreaterOD
            OALLLen = LowerID
            wounded.TextMatrix(wounded.row, 3) = 0
            wounded.TextMatrix(wounded.row, 4) = wounded.TextMatrix(1, 4)
            For l = 1 To wounded.row - 1
                wounded.TextMatrix(wounded.row, 3) = CDbl(wounded.TextMatrix(wounded.row, 3)) + CDbl(wounded.TextMatrix(l, 3))
                If CDbl(wounded.TextMatrix(wounded.row, 4)) > CDbl(wounded.TextMatrix(l, 4)) Then
                    wounded.TextMatrix(wounded.row, 4) = wounded.TextMatrix(wounded.row, 4)
                Else
                    wounded.TextMatrix(wounded.row, 4) = wounded.TextMatrix(l, 4)
                End If
            Next
            wounded.TextMatrix(wounded.row, 3) = wounded.TextMatrix(wounded.row, 3) + (2 * ((wounded.Rows - 2) - 1)) - (6 * ((wounded.Rows - 2) - 1))
            wounded.TextMatrix(wounded.row, 3) = Round(wounded.TextMatrix(wounded.row, 3), 2)
        End If
    End If
End If

If SSTab1.Tab = 1 Then
    If Val(txtnoofcores) = 0 Then
        SSTab1.Tab = 0
        MDIDesign.StatusBar1.Panels(2).Text = "Please enter core nos and details"
'        StatusBar1.Panels(4).Picture = "\\server\D\Enhance\ICON IU\ARROW12B.ICO"
        txtnoofcores.SetFocus
        Exit Sub
    Else
        
'        OptCom(0).SetFocus
        Exit Sub
'        cmbCore.Text = "All Cores"
    End If
End If

If SSTab1.Tab = 2 Then
    If GCore.row < Val(txtnoofcores) Then
        SSTab1.Tab = 1
        MDIDesign.StatusBar1.Panels(2).Text = "Please enter core Dimension"
        Exit Sub
    End If
    
    If G1.Cols = 1 Then
        ss = cmbSelctCore
        cmbSelctCore = ""
        cmbSelctCore = ss
        SetPTurnComm = ""
    Else
    End If
'    txtturns.SetFocus
    cmbConStyle.SetFocus
ElseIf SSTab1.Tab = 3 Then
'    If InStr(1, UCase(cmbEquipType), "BAR") > 0 Or InStr(1, UCase(cmbEquipType), "WINDOW") > 0 Then
'        If PreviousTab <= 2 Then
'             SSTab1.Tab = 4
'        ElseIf PreviousTab >= 4 Then
'            SSTab1.Tab = 2
'        End If
'        Exit Sub
'    Else
        If prcFINDCoreNo > 0 Then
            If G1.Cols = 1 Then
                cmbSelctCore = 1
            End If
        Else
            txtHt = ""
        End If
        If GCore.row < Val(txtnoofcores) Then
            SSTab1.Tab = 1
            MDIDesign.StatusBar1.Panels(2).Text = "Please enter core Dimension"
            Exit Sub
        ElseIf Val(txtTurns) < 1 Then
            SSTab1.Tab = 2
            MDIDesign.StatusBar1.Panels(2).Text = "Please Set turns"
            Exit Sub
        End If
'    End If
    SSTab2.Tab = 0
    If G1.Cols = 1 Then
        cmbSelctCore = 1
    End If
    cmbSelctCore.SetFocus
'    cmbSelctCore = prcFINDCoreNo
ElseIf SSTab1.Tab = 4 Then
    
    If prcCheckAllCoreNo Then
        SSTab1.Tab = PreviousTab
        Exit Sub
    End If
    txtdrg.SetFocus
    Exit Sub
End If

If PreviousTab <> 2 Then
    If SSTab1.Tab = 2 Then
        If InStr(1, UCase(cmbEquipType), "WINDOW") > 0 Then
            If PreviousTab < 2 Then
                SSTab1.Tab = 3
            ElseIf PreviousTab > 2 Then
                SSTab1.Tab = 1
            End If
        End If
    End If
End If

'err:
'
'    If Len(Trim(err.Description)) > 0 Then
'        MsgBox err.Description & vbCrLf & err.Number, vbExclamation
'    End If
End Sub

Private Function prcFINDCoreNo() As Double
Dim Check As Integer
Check = 0
If G1.Cols = 1 Then
    If Val(cmbSelctCore) <> 1 Then
        cmbSelctCore = 1
    End If
    prcFINDCoreNo = 1
Else
    For j = 1 To Val(txtnoofcores)
        For i = 1 To G1.Cols - 1
            If G1.TextMatrix(1, i) = j Then
                Check = 1
            End If
        Next
        If Check = 0 Then
            prcFINDCoreNo = j
            Exit Function
        End If
    Next
'    prcClearALL
    txtdrg.SetFocus
End If
End Function

Private Function prcCheckAllCoreNo() As Boolean
Dim Check As Integer
Check = 0
For j = 1 To Val(txtnoofcores)
    For i = 1 To G1.Cols - 1
        If G1.TextMatrix(1, i) = j Then
            Check = 1
            Exit For
        End If
    Next
    If Check = 0 Then
        prcCheckAllCoreNo = True
        Exit Function
    End If
    Check = 0
Next
prcCheckAllCoreNo = False
End Function

Private Sub prcDeleteDup()
Dim value(10) As Double
Dim WoundV(10) As Double
Dim wCOUNT As Integer
Dim Count As Integer

Dim RSTW As Recordset

'Set RSTW = New Recordset
'RSTW.Open "DELETE FROM distwoundgrid", con, adOpenDynamic, adLockOptimistic
'Set RSTW = Nothing
'
'Set RSTW = New Recordset
'RSTW.Open "distwoundgrid", con, adOpenDynamic, adLockOptimistic
'    For K = 1 To wounded.Rows - 1
'        RSTW.AddNew
'        RSTW!valu = wounded.TextMatrix(K, 0)
'        RSTW.Update
'    Next
'Set rst = Nothing
'
'Set RSTW = New Recordset
'RSTW.Open "SELECT DISTINCT VALU FROM distwoundgrid ORDER BY VALU", con, adOpenDynamic, adLockOptimistic
'If RSTW.EOF = False = RSTW.BOF = False Then
'    Do Until RSTW.EOF
'        WoundV(K) = RSTW.Fields(0)
'        RSTW.MoveNext
'    Loop
'    wCOUNT = RSTW.RecordCount
'End If
'Set RSTW = Nothing

Count = 0

Set rst = New Recordset
rst.Open "delete from coresetdesign", con, adOpenDynamic, adLockOptimistic
Set rst = Nothing

Set rst = New Recordset
rst.Open "select * from coresetdesign", con, adOpenDynamic, adLockOptimistic
For d = 1 To wounded.Rows - 1
    rst.AddNew
    rst.Fields(0) = Trim(wounded.TextMatrix(d, 0))
    rst.Fields(1) = Trim(wounded.TextMatrix(d, 1))
    rst.Fields(2) = Trim(wounded.TextMatrix(d, 2))
    rst.Fields(3) = Trim(wounded.TextMatrix(d, 3))
    rst.Fields(4) = Trim(wounded.TextMatrix(d, 4))
    rst.Update
Next
wounded.Rows = 1
Set rst = New Recordset
rst.Open "select distinct [no] from coresetdesign order by [no]", con, adOpenDynamic, adLockOptimistic
Do Until rst.EOF
    Dim rst1 As Recordset
    Set rst1 = New Recordset
    rst1.Open "select * from coresetdesign where [no] = " & Trim(rst.Fields(0)) & "", con, adOpenDynamic, adLockOptimistic
    
        wounded.Rows = wounded.Rows + 1
        wounded.row = wounded.Rows - 1
        wounded.TextMatrix(wounded.row, 0) = rst1.Fields(0)
        wounded.TextMatrix(wounded.row, 1) = rst1.Fields(1)
        wounded.TextMatrix(wounded.row, 2) = rst1.Fields(2)
        wounded.TextMatrix(wounded.row, 3) = rst1.Fields(3)
        wounded.TextMatrix(wounded.row, 4) = rst1.Fields(4)
   
    Set rst1 = New Recordset
    rst.MoveNext
Loop
Set rst = New Recordset
'MsgBox wounded.TextMatrix(1, 5)
For K = 1 To wounded.Rows - 1
    For t = K + 1 To wounded.Rows - 1
        If wounded.TextMatrix(K, 0) = wounded.TextMatrix(t, 0) Then
            value(Count) = t
            Count = Count + 1
        End If
    Next
Next

'wounded.Rows = 1
For i = (Count - 1) To 0 Step -1
    wounded.RemoveItem (value(i))
Next

End Sub

Private Sub SSTab1_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub SSTab2_DblClick()
SSTab2.Tab = 0
End Sub

Private Sub SSTab2_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub SWG_Click()
    txtGEdit.Top = SWG.Top + SWG.CellTop
    txtGEdit.Left = SWG.Left + SWG.CellLeft
    txtGEdit.Width = SWG.CellWidth
    txtGEdit.Height = SWG.CellHeight
    txtGEdit.Text = SWG.TextMatrix(SWG.RowSel, SWG.ColSel)
    txtGEdit.Visible = True
    txtGEdit.SetFocus
End Sub








Private Sub txtCompen_Change(Index As Integer)
If Len(txtCompen(Index)) = 0 Then
    Exit Sub
'Else
''    txtCompen(Index) = Round(txtCompen(Index), 3)
'    txtCompen(Index) = Format(Round(txtCompen(Index), 3), "##,###.00")
End If
Dim value As String * 1
Dim str1 As String
If InStr(1, txtCoreRatio, "-") > 0 Then
    
    Rationlength = Left(txtCoreRatio, InStr(1, (txtCoreRatio), "/") - 1)
    count1 = 0
    For p = 1 To InStr(1, (txtCoreRatio), "/")
        value = Mid(Rationlength, p, p)
        If value = "-" Or p > Len(Rationlength) Then
                If Len(Trim(txtCompen(0))) = 0 Then
                    txtCompen(0) = 0
                End If
             If count1 > 0 Then
                txtCompen(count1) = Format(firstratio / CDbl(str1) * txtCompen(0), "##,##0.000")
                txtCompen(count1).Visible = True
             Else
                firstratio = str1
             End If
             count1 = count1 + 1
             str1 = ""
        Else
            str1 = str1 & value
        End If
    Next
End If
End Sub

Private Sub txtCompen_GotFocus(Index As Integer)
txtCompen(Index).SelStart = 0
txtCompen(Index).SelLength = Len(txtCompen(Index))
MDIDesign.StatusBar1.Panels(2).Text = "Set Compensation"
End Sub

Private Sub txtCompen_KeyPress(Index As Integer, KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9-.]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtCondPer_GotFocus()
txtCondPer.SelStart = 0
txtCondPer.SelLength = Len(txtCondPer)
MDIDesign.StatusBar1.Panels(2).Text = "Set Conductor Per Coil"
End Sub

Private Sub txtCoreRatio_Change()

If Len(txtCoreRatio) = 0 Then
    Exit Sub
End If

'If InStr(1, txtCoreRatio, "-") = 0 Then
'    Chkoption.value = False
'    Chkoption.Enabled = False
'Else
'    Chkoption.value = False
'    Chkoption.Enabled = True
'End If

End Sub


Private Sub txtCoreRatio_GotFocus()

MDIDesign.StatusBar1.Panels(2).Text = "Core Ratio"
End Sub

Private Sub txtCoreType_Change()
If UCase(txtCoreType) = "METERING" Then
    txtHt = 15
End If
End Sub

Private Sub txtCoreWT_GotFocus()
txtCoreWT.SelStart = 0
txtCoreWT.SelLength = Len(txtCoreWT)
MDIDesign.StatusBar1.Panels(2).Text = "Core Wt."
End Sub
Private Sub cmbCal_Click()


'frame22.top

If Len(Trim(txtdrg)) = 0 Then Exit Sub

If InStr(Trim(cmbEquipType), "Wo-R") > 1 Then
    If prcDrgCheck(txtdrg, cmbType, txtPriCros, txtTurns, cmbConStyle.List(cmbConStyle.ListIndex, 1), cmbConStyle.List(cmbConStyle.ListIndex, 0)) = True Then
        Casting.Rows = 1
        Former.Rows = 1
        Mould.Rows = 1
        WOR1A = 0
        WOR2A = 0
        txtPriMeanL = 0
        txtPriLength = 0
        txtPriTotLength = 0
        txtpricuWt = 0
        txtPRiRes20 = 0
        txtPRiRes75 = 0
        txtCoreWT = 0
        txtSecCuWt = 0
        txtPrimaryCopper = 0
        txtResinWt = 0
        
        MsgBox "No Primary data avalible for the selected Drg!!!" & vbCrLf & "Note : - Still if continue enter primary data for the selected drawing.", vbCritical
        Exit Sub
    End If
End If

On Error GoTo err

WOR1.Visible = True
WOR1 = "RESIN CLEARANCE WIDTH WISE (mm) :"
If Len(txtdrg) = 0 Then
    MsgBox "Please enter Drawing No.", vbExclamation
    Exit Sub
End If

Dim Sta As String * 1
Sta = ""
cmbEquipType.Tag = ""
If InStr(1, UCase(cmbEquipType), "BAR") > 0 Then
    If UCase(cmbStyle) = "CIRCULAR" Then
        Sta = "B"
        
            joel = MsgBox("Press -Yes- to Rectangular Bar and -No- For Circular Bar?", vbQuestion + vbYesNo)
            If joel = vbNo Then
                Sta = "C"
            Else
                Sta = "R"
            End If
            cmbEquipType.Tag = Sta
    End If
ElseIf InStr(1, UCase(cmbEquipType), "WINDOW") > 0 Then
    If UCase(cmbStyle) = "CIRCULAR" Then
        Sta = "W"
            joel = MsgBox("Press -Yes- to Rectangular Window" & vbCrLf & "and    -No- to Circular Window", vbQuestion + vbYesNo)
            If joel = vbNo Then
                Sta = "C"
            Else
                Sta = "R"
            End If
        cmbEquipType.Tag = Sta
    End If
End If

Former.Rows = 1
Mould.Rows = 1
Casting.Rows = 1

If Len(txtdrg) > 0 Then
    'Only For Wound
    If Len(Trim(LastSet)) = 0 Then
        Frame14.Visible = True
        Frame15.Caption = "Former Dimensions :"
        Frame15.Visible = True
        Frame14.Caption = "Clearances :"
        FrameWound.Visible = False
        Frame16.Visible = True
        prcDimen
'        Exit Sub
    'For Window
    ElseIf LastSet = "F" Then
        If cmbStyle = "Rectangluar" Then
            Frame14.Visible = True
            Frame15.Visible = False
            FrameWound.Visible = False
            prcWinddDiemRec
        ElseIf cmbStyle = "Circular" Then
            Frame14.Visible = True
            Frame15.Visible = False
            FrameWound.Visible = False
            prcWindDiemRound
        End If
    End If
End If

'Total Core Wt
Dim TotalCoreWt As Double
'Total Sec Copper Wt
Dim TotalSecCopperWt As Double

'Calulating Core wt for all cores
For i = 1 To txtnoofcores
    For j = 1 To G1.Cols - 1
        If G1.TextMatrix(1, j) = i Then
            TotalCoreWt = TotalCoreWt + G1.TextMatrix(81, j)
            Exit For
        End If
    Next
Next
txtCoreWT = TotalCoreWt

'Calulating secondary copper wt for all cores
For i = 1 To txtnoofcores
    For j = 1 To G1.Cols - 1
        If G1.TextMatrix(1, j) = i Then
            TotalSecCopperWt = TotalSecCopperWt + G1.TextMatrix(80, j)
            Exit For
        End If
    Next
Next

txtSecCuWt = TotalSecCopperWt
If InStr(1, UCase(cmbEquipType), "BAR") > 0 Then
'    txtPrimaryCopper = txtpricuWt
'    txtPrimaryCopper = txtpricuWt
    Set rst = New ADODB.Recordset
    If Len(Trim(cmbType)) = 0 Then
        rst.Open "SELECT t_id, t_winwid, t_winhei, t_pribarlen FROM mouldvol2 where t_drg = '" & txtdrg & "'", con, adOpenStatic, adLockReadOnly
    Else
        rst.Open "SELECT t_id, t_winwid, t_winhei, t_pribarlen FROM mouldvol2 where t_drg = '" & txtdrg & "' and t_drgtype = '" & cmbType & "'", con, adOpenStatic, adLockReadOnly
    End If
    If rst.EOF = False And rst.BOF = False Then
'        txt
        If Sta = "C" Then
            txtPriCros = (3.142 * (rst!t_id * rst!t_id)) / 4
        Else
            txtPriCros = (rst!t_winwid * rst!t_winhei)
            If IsNull(rst!t_pribarlen) = False Then
                If Len(Trim(Sta)) = 0 Then
                    txtPrimaryCopper = (rst!t_winwid * rst!t_winhei * rst!t_pribarlen) * 8.9 / 1000000
                End If
            Else
                If Len(Trim(Sta)) = 0 Then
                    txtPrimaryCopper = (rst!t_winwid * rst!t_winhei * 0) * 8.9 / 1000000
                End If
            End If
        End If
        'PRIMARY BAR LENGTH IS USED TO CALCULATE pRIMARY rESISTANCE AT 20 AND 75 DEGREE
        
        PRes20 = ((0.0177 * (Val(rst!t_pribarlen) / 1000)) / txtPriCros) * 1000
        ''mili ohms
        txtPRiRes20 = Round(Val(PRes20), 5)
        PRes = PRes20 * 1.2161
        txtPRiRes75 = Round(PRes, 5)
    '    txtPRiRes
        WattLoss = WatLos(CDbl(PRes))
        WattLoss = WattLoss / 1000
        txtWattPri = Round(WattLoss, 2)
    End If
'    Set rst = Nothing
ElseIf InStr(1, UCase(cmbEquipType), "WOUND") > 0 Then
'Primary Winding Resistance and Watt Loss
'Resi 20 = 0.0177 * primLen(Meter) / PriCrossSection Area
'@ 20 Degee Cen
'    Dim PRes20 As Double
'    Dim PRes As Double
'    Dim WattLoss As Double
    txtPrimaryCopper = txtpricuWt
    PRes20 = (0.0177 * (txtPriTotLength / 1000)) / txtPriCros * 1000
    txtPRiRes20 = Round(PRes20, 5)
    PRes = PRes20 * 1.2161
    txtPRiRes75 = Round(PRes, 5)
'    txtPRiRes
    WattLoss = WatLos(CDbl(PRes))
    WattLoss = WattLoss / 1000
    
    txtWattPri = Round(WattLoss, 2)
End If
Dim ResWt As Weights
Set ResWt = New Weights

If InStr(1, cmbEquipType, "Wound") > 0 Then
    txtResinWt = ResWt.CalResinWt(txtdrg, cmbType, txtSecCuWt, txtCoreWT, cmbEquipType, txtpricuWt)
    Frame21.Visible = True
    Frame22.Visible = True
    Frame22.Top = 3555
    Frame22.Left = 6975
    Label85.Visible = True
    Label87.Visible = True
    Label86.Visible = True
    Label90.Visible = True
    Label91.Visible = True
    txtPRiRes20.Visible = True
    txtPRiRes75.Visible = True
    txtWattPri.Visible = True
ElseIf InStr(1, cmbEquipType, "Bar") > 0 Then
    txtResinWt = ResWt.CalResinWt(txtdrg, cmbType, txtSecCuWt, txtCoreWT, cmbEquipType, 0)
    Frame21.Visible = False
    Frame22.Visible = True
    Frame22.Top = 495
    Frame22.Left = 6975
    Label85.Visible = True
    Label87.Visible = True
    Label86.Visible = True
    Label90.Visible = True
    Label91.Visible = True
    txtPRiRes20.Visible = True
    txtPRiRes75.Visible = True
    txtWattPri.Visible = True
Else
    txtResinWt = ResWt.CalResinWt(txtdrg, cmbType, txtSecCuWt, txtCoreWT, cmbEquipType, 0)
    Frame21.Visible = False
    Frame22.Visible = True
    Frame22.Top = 495
    Frame22.Left = 6975
    txtPrimaryCopper.Text = "-N.A.-"
    Label82.Enabled = False
    txtPrimaryCopper.Enabled = False
    Label85.Visible = False
    Label87.Visible = False
    Label86.Visible = False
    Label90.Visible = False
    Label91.Visible = False
    txtPRiRes20.Visible = False
    txtPRiRes75.Visible = False
    txtWattPri.Visible = False
End If

err:
    
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbExclamation
    End If

End Sub

Private Function WatLos(PR As Double) As Double

Dim value As String * 1
Dim str1 As String

Dim MaxValue As Double

'If primary rated continuse current then taking Rated continous else
If Len(txtRateConti) = 0 Then

    For i = 0 To txtnoofcores - 1
        For j = 1 To Len(txtRatioDum(i)) + 1
            value = Mid(txtRatioDum(i), j, j)
            If value = "-" Or Len(txtRatioDum(i)) < j Then
                If CDbl(str1) > MaxButton Then
                    MaxValue = CDbl(str1)
                End If
                str1 = ""
            Else
                str1 = str1 & value
            End If
        Next
    Next

Else
    MaxValue = txtRateConti
End If
''
WatLos = MaxValue * MaxValue * PR
WatLos = Format(WatLos, "##.00")

End Function


Private Sub txtDelR_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Normal Current Density (A/sq.mm)"
End Sub


Private Sub txtDelStc_Change()
If Len(Trim(txtDelStc)) = 0 Then
    Exit Sub
End If
If Trim(txtDelStc) > 180 Then
    MsgBox "Short circuit current density exceds the limit!", vbCritical
'    txtPriCros = ""
    lstCond.ListIndex = ListIndexofCondutor
    txtTurns.SetFocus
End If
End Sub

Private Sub txtDelStc_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "STC Current Desity (1sec) (A/Sq mm)"
End Sub

Private Sub txtdrg_GotFocus()
txtdrg.SelStart = 0
txtdrg.SelLength = Len(txtdrg)
MDIDesign.StatusBar1.Panels(2).Text = "Set Drawing No"
End Sub

Private Sub txtdrg_LostFocus()

On Error Resume Next

If Trim(txtdrg) = "PR-" Or Trim(txtdrg) = "" Then Exit Sub

If Len(Trim(txtdrg)) > 0 Then
    txtdrg = UCase(txtdrg)
Else
    Exit Sub
End If

cmbType.Clear

''+++++++++++++++++++
''drg type form mould vol
Set rst = New Recordset
rst.Open "select distinct t_drgtype from mouldvol2 where t_drg = '" & Trim(txtdrg) & "'", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
    rst.MoveFirst
    cmbType.Clear
    Do Until rst.EOF
        If IsNull(rst!T_DRGTYPE) = False Then
            cmbType.AddItem rst!T_DRGTYPE
        End If
        rst.MoveNext
    Loop
    If cmbType.ListIndex > -1 Then cmbType.ListIndex = 0
Else
    MsgBox "Drawing No : " & txtdrg & " doest not exit in the database.", vbInformation
    cmbType.Clear
'    txtdrg.SetFocus
End If
Set rst = Nothing

End Sub

Private Sub txtGEdit_Change()
SWG.TextMatrix(SWG.RowSel, SWG.ColSel) = txtGEdit.Text
End Sub

Private Sub txtGEdit_LostFocus()
txtGEdit.Visible = False
End Sub

Private Sub txtHt_GotFocus()
txtHt.SelStart = 0
txtHt.SelLength = Len(txtHt)
MDIDesign.StatusBar1.Panels(2).Text = "Core Height"
End Sub

Private Sub txtHt_LostFocus()

    On Error GoTo err

'    txtCoreRatio = ""
'    txtSecTap = ""
    
    If Val(txtHt) = 0 Then
        Exit Sub
    End If
    
    If Val(txtHt) < 10 Then
        MsgBox "Height should be greater than or equal to 10 mm.", vbExclamation
        txtHt = ""
        Exit Sub
    End If
    
    intre = Val(txtHt) Mod 5
    
    If intre <> 0 Then
        MsgBox "Height should be in the divisible of 5", vbExclamation
        txtHt.SetFocus
        txtHt = 0
    End If
    
    txtEArea = prcAreaEff(txtHt, txtBUP)
    txtArea = prcArea(txtHt, txtBUP)
    If UCase(Trim(cmbMetal)) = UCase("MU Metal") Then
        ''for mu metal density will be 8.8
        txtWt = Round((txtMean * txtEArea) * 8.8 / 1000, 2)
    Else
        ''density 7.65
        txtWt = Round((txtMean * txtEArea) * 7.65 / 1000, 2)
    End If
    Exit Sub
    'secondary turns calculations
    
    If InStr(1, txtRatio(CInt(cmbSelctCore - 1)), "-") = 0 Then
        If cmbStyle = "Rectangluar" Then
            txtSecTurn = txtTurns * txtRatio(CInt(cmbSelctCore - 1)) / txtRatioR2(cmbSelctCore - 1)
        Else
            txtSecTurn = 1 * txtRatio(CInt(cmbSelctCore - 1)) / txtRatioR2(cmbSelctCore - 1)
        End If
        txtMinCsSec = prcCalSecC(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), txtRatio(CInt(cmbSelctCore - 1)), cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
        txtSecCS = prcCalNorSecC(txtRatioR2(CInt(cmbSelctCore - 1)))
'    ######Short Circuit Current Density
        txtDelta = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), txtSecTurn, txtRatioR2(cmbSelctCore - 1), CDbl(txtODn))
        GNo = txtRatio(cmbSelctCore - 1)
    Else
    
        Dim int1 As Double
        Dim value As String * 1
        Dim ARatio As String
        Dim str1 As String
        Dim Gnumber As Double
        ARatio = txtRatio(cmbSelctCore - 1)
        
        For i = 1 To Len(ARatio) + 1
            value = Mid(txtRatio(cmbSelctCore - 1), i, i)
            If value = "-" Or i > Len(ARatio) Then
                 If CDbl(str1) > Gnumber Then
                      Gnumber = CDbl(str1)
                 End If
                 str1 = ""
            Else
                str1 = str1 & value
            End If
        
        Next
        
        GNo = Gnumber
        
        'Setting tap turing
        txtSecTap = ""
        For i = 1 To Len(ARatio) + 1
            value = Mid(txtRatio(cmbSelctCore - 1), i, i)
            If value = "-" Or i > Len(ARatio) Then
                 If CDbl(str1) > Gnumber Then
                      Gnumber = CDbl(str1)
                 Else
                    If CDbl(str1) <> Gnumber Then
                        If cmbStyle = "Rectangluar" Then
                            txtSecTap = txtSecTap & "-" & Round(txtTurns * CDbl(str1) / txtRatioR2(cmbSelctCore - 1), 2)
                        Else
                            txtSecTap = txtSecTap & "-" & Round(1 * CDbl(str1) / txtRatioR2(cmbSelctCore - 1), 2)
                        End If
                    End If
                 End If
                 str1 = ""
            Else
                str1 = str1 & value
            End If
            
        Next
        GNo = Gnumber
        txtSecTap = Mid(txtSecTap, 2)
        If cmbStyle = "Rectangluar" Then
            txtSecTurn = Round(txtTurns * Gnumber / txtRatioR2(cmbSelctCore - 1), 2)
        Else
            txtSecTurn = Round(1 * Gnumber / txtRatioR2(cmbSelctCore - 1), 2)
        End If
    End If
    txtCoreRatio = txtRatio(CInt(cmbSelctCore - 1)) & "/" & txtRatioR2(cmbSelctCore - 1)
        
    'setting C/section
    str1 = ""
    Dim MCS As Double
    Dim MaxCs As Double
    txtSecCS = prcCalNorSecC(txtRatioR2(CInt(cmbSelctCore - 1)))
    
    If CDbl(txtSecCS) >= CDbl(txtMinCsSec) Then
        txtMinCsSec = txtSecCS
        MinCross = txtSecCS
    Else
        MinCross = txtMinCsSec
    End If
    
    Exit Sub
'*************************************
    'if tap present in primary and and to find
    If InStr(1, txtRatio(CInt(cmbSelctCore - 1)), "-") > 0 Then
        
        If Chkoption.value = True Then
           
           'Find C/s section of all current in primary side
            For i = 1 To Len(ARatio) + 1
                value = Mid(txtRatio(cmbSelctCore - 1), i, i)
                If value = "-" Or i > Len(ARatio) Then
                     MaxCs = prcCalSecC(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), CDbl(str1), cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
                     
                     If MaxCs >= txtSecCS Then
                     
                        Set rst = New Recordset
                        rst.Open "select * from secConductor where t_cs >=" & MaxCs & " ", con, adOpenStatic, adLockReadOnly
                        If rst.EOF = False And rst.BOF = False Then
                            
                            GCross.Rows = GCross.Rows + 1
                            GCross.Rows = GCross.Rows - 1
                            GCross.TextMatrix(GCross.row, 0) = str1
                            GCross.TextMatrix(GCross.row, 1) = rst!t_cs
                            GCross.TextMatrix(GCross.row, 2) = rst!t_swg
                            GCross.TextMatrix(GCross.row, 3) = rst!t_ods
                            GCross.TextMatrix(GCross.row, 4) = rst!t_cs
                            GCross.TextMatrix(GCross.row, 5) = "S"
                        
                        End If
                        Set rst = Nothing
                        
                     Else
                        
                        Set rst = New Recordset
                        rst.Open "select * from secConductor where t_cs >=" & txtSecCS & " ", con, adOpenStatic, adLockReadOnly
                        If rst.EOF = False And rst.BOF = False Then
                                                          
                            GCross.Rows = GCross.Rows + 1
                            GCross.Rows = GCross.Rows - 1
                            GCross.TextMatrix(GCross.row, 0) = str1
                            
                            GCross.TextMatrix(GCross.row, 1) = rst!t_cs
                            GCross.TextMatrix(GCross.row, 2) = rst!t_swg
                            GCross.TextMatrix(GCross.row, 3) = rst!t_ods
                            GCross.TextMatrix(GCross.row, 4) = rst!t_cs
                            GCross.TextMatrix(GCross.row, 5) = "S"
                                          
                        End If
                        Set rst = Nothing
                        
                     End If
                     If MCS < MaxCs Then
                          MCS = CDbl(MaxCs)
                     End If
                     str1 = ""
                Else
                    str1 = str1 & value
                End If
            Next
            txtMinCsSec = MCS
        Else
        
        'Find C/s section of lowest current in primary side
            Dim mincurr As Double
            For i = 1 To Len(ARatio) + 1
            value = Mid(txtRatio(cmbSelctCore - 1), i, i)
            If value = "-" Or i > Len(ARatio) Then
                 If CDbl(str1) < mincurr Or mincurr = 0 Then
                      mincurr = CDbl(str1)
                 End If
                 str1 = ""
            Else
                str1 = str1 & value
            End If
            Next
'            MsgBox Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1)
            txtMinCsSec = prcCalSecC(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1), mincurr, cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
            
        End If
        
    Else
    
        'for  single current in primary
        txtMinCsSec = prcCalSecC(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), txtRatio(cmbSelctCore - 1), cmbTime, txtRatioR2(CInt(cmbSelctCore - 1)))
        
    End If
    
'''''''''''    If CDbl(txtSecCS) >= CDbl(txtMinCsSec) Then
'''''''''''
'''''''''''        Set rst = New ADODB.Recordset
'''''''''''        rst.Open "select * from secconductor where t_cs >= " & txtSecCS & " ", con, adOpenStatic, adLockReadOnly
'''''''''''        If rst.EOF = False And rst.BOF = False Then
'''''''''''            txtSwg = rst!t_swg
'''''''''''            txtNo = "S"
'''''''''''            txtODs = rst!t_ods
'''''''''''            txtODn = rst!t_cs
'''''''''''            txtODn.Tag = rst!t_cs
'''''''''''        End If
'''''''''''
'''''''''''        txtMinCsSec = txtSecCS
'''''''''''
'''''''''''    Else
'''''''''''
'''''''''''        Set rst = New ADODB.Recordset
'''''''''''        rst.Open "select * from secconductor where t_cs >= " & txtMinCsSec & " ", con, adOpenStatic, adLockReadOnly
'''''''''''        SHHH = "select * from secconductor where t_cs >= " & txtMinCsSec & " "
'''''''''''        If rst.EOF = False And rst.BOF = False Then
'''''''''''            txtSwg = rst!t_swg
'''''''''''            txtNo = "S"
'''''''''''            txtODs = rst!t_ods
'''''''''''            txtODn = rst!t_cs
'''''''''''            txtODn.Tag = rst!t_cs
'''''''''''        End If
''''''''''''        txtSecCS = txtMinCsSec
'''''''''''    End If
'''''''''''
'''''''''''    Set rst = Nothing
    
'    If cmbStyle = "Rectangluar" Then
'        txtLayers = CalLayers(txtODs)
'    End If

'    txtDelta = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), txtSecTurn, txtRatioR2(cmbSelctCore - 1), CDbl(txtODn))

err:
    
    MsgBox err.Description & " " & err.Number

End Sub

'Calculating layers RECTANGULAR
Private Function CalLayers(Dia As Double) As Double
        
    Dim value As Double
    
    Dim Turnss As Double
    Turnss = txtSecTurn
        
    Dim Count As Byte
    Dim Ins(30) As Double
    Ins(0) = 1.5
    Ins(1) = 0.25
    Ins(2) = 0.25
    Ins(3) = 0.25
    Ins(4) = 0.25
    Ins(5) = 0.25
    Ins(6) = 0.25
    Ins(7) = 0.25
    Ins(8) = 0.25
    Ins(9) = 0.25
    Ins(10) = 0.25
    Ins(11) = 0.25
    Ins(12) = 0.25
    Ins(13) = 0.25
    Ins(14) = 0.25
    Ins(15) = 0.25
    Ins(16) = 0.25
    Ins(17) = 0.25
    Ins(18) = 0.25
    Ins(19) = 0.25
    Ins(20) = 0.25
    Ins(21) = 0.25
    Ins(22) = 0.25
    Ins(23) = 0.25
    Ins(24) = 0.25
    Ins(25) = 0.25
    Ins(26) = 0.25
    Ins(27) = 0.25
    Ins(28) = 0.25
    Ins(29) = 0.25
    Ins(30) = 3
    Dim K As Double
    
'Type of Conductor
    If UCase(txtNo) = "S" Then
        Dia = Dia * 1
    ElseIf UCase(txtNo) = "D" Then
        Dia = Dia * 2
    ElseIf UCase(txtNo) = "T" Then
        Dia = Dia * 3
    ElseIf UCase(txtNo) = "Q" Then
        Dia = Dia * 4
    End If
'Counting No of turns
    
    Dim typ As Double
    
    For i = 0 To CInt(txtSecTurn)
        K = 0
'For First limb
        For l = 0 To Count
            K = K + Ins(l)
        Next
        If Count > 0 Then
            K = K + (txtODs * Count)
        End If
        
        'Rounding of k
        If InStr(1, K, ".") > 0 Then
            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
            Else
                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
            End If
        End If
        
        i = ((CDbl(txtWL1) - 2 * (K)) * 0.9) / Dia
        'Rounding of i
        If InStr(1, i, ".") > 0 Then
            i = Left(i, InStr(1, i, ".") - 1)
        End If
        
        value = value + i
        
'        If value >= CInt(txtSecTurn) Then
'             typ = 0.25
'             CalLayers = (Count + typ)
'             Exit Function
'        End If
        
'For secound limb

        value = value + i
'        If value >= CInt(txtSecTurn) Then
'             typ = 0.5
'             CalLayers = (Count + typ)
'             Exit Function
'        End If
        
'For third limb
        
        Dim s As Double
        
        s = (((CDbl(txtWW1) - 2 * (K)) * 0.9) / Dia)
        'Rounding of s
        If InStr(1, s, ".") > 0 Then
            s = Left(s, InStr(1, s, ".") - 1)
        End If
        value = value + s
'        If value >= CInt(txtSecTurn) Then
'             typ = 0.75
'             CalLayers = (Count + typ)
'             Exit Function
'        End If
        
'For Fourth Limb
        
        value = value + s
        turns1 = Turnss
        Turnss = Turnss - value
        
        If Turnss <= 0 Then
            MsgBox ((turns1 / value) * 100)
            finalt = ((turns1 / value) * 100)
            If finalt <= 25 Then
                typ = 0.25
                CalLayers = (Count + typ)
                Exit Function
            ElseIf finalt > 25 And finalt <= 50 Then
                typ = 0.5
                CalLayers = (Count + typ)
                Exit Function
            ElseIf finalt > 50 And finalt <= 75 Then
                typ = 0.75
                CalLayers = (Count + typ)
                Exit Function
            ElseIf finalt > 75 And finalt <= 100 Then
                typ = 1
                CalLayers = (Count + typ)
                Exit Function
            End If
        End If
        value = 0
        typ = 0
        Count = Count + 1
'        If value >= CInt(txtSecTurn) Then
'             CalLayers = (Count + typ)
'             Exit Function
'        End If

            
    Next
    
    CalLayers = i
    
End Function

'Calculate layers for tap RECTANGULAR Asc same dia
Private Function CalLayersTapS()
        
    Dim value As Double
    Dim Turnss As Double
    Dim C As Double
    Dim Count As Byte
    Dim Ins(20) As Double
    Ins(0) = 1.5
    Ins(1) = 0.25
    Ins(2) = 0.25
    Ins(3) = 0.25
    Ins(4) = 0.25
    Ins(5) = 0.25
    Ins(6) = 0.25
    Ins(7) = 0.25
    Ins(8) = 0.25
    Ins(9) = 0.25
    Ins(10) = 0.25
    Ins(11) = 0.25
    Ins(12) = 0.25
    Ins(13) = 0.25
    Ins(14) = 0.25
    Ins(15) = 0.25
    Ins(16) = 0.25
    Ins(17) = 0.25
    Ins(18) = 0.25
    Ins(19) = 0.25
    Ins(20) = 3
    Dim K As Double
    

'Counting No of turns
    
Dim typ As Double
    
For shibu = 1 To GCross.Rows - 1
    
'Type of Conductor
    If UCase(txtNo) = "S" Then
         Dia = GCross.TextMatrix(shibu, 1) * 1
    ElseIf UCase(txtNo) = "D" Then
        Dia = GCross.TextMatrix(shibu, 1) * 2
    ElseIf UCase(txtNo) = "T" Then
        Dia = GCross.TextMatrix(shibu, 1) * 3
    ElseIf UCase(txtNo) = "Q" Then
        Dia = GCross.TextMatrix(shibu, 1) * 4
    End If
    
'setting turns
    Turnss = GCross.TextMatrix(shibu, 7)
    For i = 0 To CInt(GCross.TextMatrix(shibu, 7))
        K = 0

'For First limb
        For l = 0 To Count
            K = K + Ins(l)
        Next
        If Count > 0 Then
            K = K + (GCross.TextMatrix(shibu, 1) * Count)
        End If
        
        'Rounding of k
        If InStr(1, K, ".") > 0 Then
            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
            Else
                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
            End If
        End If
        
        i = ((CDbl(txtWL1) - 2 * (K)) * 0.9) / Dia
        'Rounding of i
        If InStr(1, i, ".") > 0 Then
            i = Left(i, InStr(1, i, ".") - 1)
        End If
        
        value = value + i
        
'For secound limb

        value = value + i
        
'For third limb
        
        Dim s As Double
        
        s = (((CDbl(txtWW1) - 2 * (K)) * 0.9) / Dia)
        'Rounding of s
        If InStr(1, s, ".") > 0 Then
            s = Left(s, InStr(1, s, ".") - 1)
        End If
        value = value + s

'For Fourth Limb
        value = value + s
        turns1 = Turnss
        Turnss = Turnss - value
        
        If Turnss <= 0 Then
            MsgBox ((turns1 / value) * 100)
            finalt = ((turns1 / value) * 100)
            If finalt <= 25 Then
'@@@@@@@@################ Quarter
                typ = 0.25
                GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If typ > 0 Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                        Else
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 7))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                If txtCoreType <> "Special Protection" Then
                
                'BM
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 0)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Exit For
            ElseIf finalt > 25 And finalt <= 50 Then
'@@@@@@@@################ Half
                typ = 0.5
                GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If typ > 0 Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                        Else
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 7))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                
                'BM
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 0)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Exit For
            ElseIf finalt > 50 And finalt <= 75 Then
'@@@@@@@@################ 3Quarter
                typ = 0.75
                GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If typ > 0 Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                        Else
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 7))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                
                'BM
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                        
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                        
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 0)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Exit For
            ElseIf finalt > 75 And finalt <= 100 Then
'@@@@@@@@################ Full
                typ = 1
                GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If typ > 0 Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                        Else
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 7))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                
                'BM
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                        
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                        
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 0)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Exit For
            End If
        End If
        value = 0
        typ = 0
        Count = Count + 1
            
    Next
'    GCross.TextMatrix(shibu, 6) = i
Next
    
End Function

'Calculate layers for tap Round Asc same dia
Private Function CalLayersTapSRound()
        
    Dim value As Double
    Dim Turnss As Double
    Dim C As Double
    Dim Count As Byte
    Dim Ins(20) As Double
    Ins(0) = 1.5
    Ins(1) = 0.25
    Ins(2) = 0.25
    Ins(3) = 0.25
    Ins(4) = 0.25
    Ins(5) = 0.25
    Ins(6) = 0.25
    Ins(7) = 0.25
    Ins(8) = 0.25
    Ins(9) = 0.25
    Ins(10) = 0.25
    Ins(11) = 0.25
    Ins(12) = 0.25
    Ins(13) = 0.25
    Ins(14) = 0.25
    Ins(15) = 0.25
    Ins(16) = 0.25
    Ins(17) = 0.25
    Ins(18) = 0.25
    Ins(19) = 0.25
    Ins(20) = 3
    Dim K As Double
    

'Counting No of turns
    
Dim typ As Double
    
For shibu = 1 To GCross.Rows - 1
    
'Type of Conductor
    If UCase(txtNo) = "S" Then
        Dia = GCross.TextMatrix(shibu, 1) * 1
    ElseIf UCase(txtNo) = "D" Then
        Dia = GCross.TextMatrix(shibu, 1) * 2
    ElseIf UCase(txtNo) = "T" Then
        Dia = GCross.TextMatrix(shibu, 1) * 3
    ElseIf UCase(txtNo) = "Q" Then
        Dia = GCross.TextMatrix(shibu, 1) * 4
    End If
    
'setting turns
    Turnss = GCross.TextMatrix(shibu, 7)
    For i = 0 To CInt(GCross.TextMatrix(shibu, 7))
        K = 0

'For First limb
        For l = 0 To Count
            K = K + Ins(l)
        Next
        If Count > 0 Then
            K = K + (GCross.TextMatrix(shibu, 3) * Count)
        End If
        
        'Rounding of k
        If InStr(1, K, ".") > 0 Then
            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
            Else
                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
            End If
        End If
        
        i = ((CDbl(txtID1) - 2 * (K)) * 0.9 * 3.142) / Dia
        'Rounding of i
        If InStr(1, i, ".") > 0 Then
            i = Left(i, InStr(1, i, ".") - 1)
        End If
        
        value = value + i
        turns1 = Turnss
        Turnss = Turnss - value
        
        If Turnss <= 0 Then
            MsgBox ((turns1 / value) * 100)
            finalt = ((turns1 / value) * 100)
            If finalt <= 25 Then
                typ = 0.25
                GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If typ > 0 Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                        Else
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 7))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                
                'BM
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = Val(GCross.TextMatrix(shibu, 14)) * Val(txtISF1)
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 0)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Exit For
            ElseIf finalt > 25 And finalt <= 50 Then
                typ = 0.5
                GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If typ > 0 Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                        Else
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 7))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                
                'BM
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 0)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Exit For
            ElseIf finalt > 50 And finalt <= 75 Then
                typ = 0.75
                GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If typ > 0 Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                        Else
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 7))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                
                'BM
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 0)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Exit For
            ElseIf finalt > 75 And finalt <= 100 Then
                typ = 1
                GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If typ > 0 Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                        Else
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 7))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                
                'BM
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * GCross.TextMatrix(shibu, 12))) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 0)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Exit For
            End If
        End If
        value = 0
        typ = 0
        Count = Count + 1


            
    Next
'    GCross.TextMatrix(shibu, 6) = i
Next
    
End Function

'Calculate layers for tap ROUND Asc for different Dia
Private Function CalLayersTapDiffRound()
'    cmdIN.Tag = ""
    Dim value As Double
    Dim Turnss As Double
    Dim Layers As Byte
    Dim C As Double
    Dim Count As Byte
    Dim Ins(2) As Double
    Ins(0) = 1.5
    Ins(1) = 0.25

    Dim K As Double
    
'Counting No of turns
    
Dim typ As Double
    
For shibu = 1 To GCross.Rows - 1
    
'Type of Conductor
    If UCase(txtNo) = "S" Then
        Dia = GCross.TextMatrix(shibu, 1) * 1
    ElseIf UCase(txtNo) = "D" Then
        Dia = GCross.TextMatrix(shibu, 1) * 2
    ElseIf UCase(txtNo) = "T" Then
        Dia = GCross.TextMatrix(shibu, 1) * 3
    ElseIf UCase(txtNo) = "Q" Then
        Dia = GCross.TextMatrix(shibu, 1) * 4
    End If
    
'setting turns
    Turnss = GCross.TextMatrix(shibu, 0)
    For i = 0 To CInt(GCross.TextMatrix(shibu, 0))
        K = 0
'For First limb
        If Layers = 0 Then
            K = K + Ins(0)
        End If
        For l = 1 To Count
            If Layers <> 0 And Count <> 0 Then
                K = K + Ins(1)
            End If
        Next
        If Count > 0 Then
            K = K + (GCross.TextMatrix(shibu, 1) * Count)
        End If
        For j = 1 To shibu
            K = K + Val(GCross.TextMatrix(j, 8))
        Next
        'Rounding of k
        If InStr(1, K, ".") > 0 Then
            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
            Else
                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
            End If
        End If
        i = ((CDbl(txtID1) - 2 * (K)) * 0.9 * 3.142) / Dia
        'Rounding of i
        If InStr(1, i, ".") > 0 Then
            i = Left(i, InStr(1, i, ".") - 1)
        End If
        value = value + i
        turns1 = Turnss
        Turnss = Turnss - value
        If Turnss <= 0 Then
            MsgBox ((turns1 / value) * 100)
            finalt = ((turns1 / value) * 100)
            If finalt <= 25 Then
                typ = 0.25
                GCross.TextMatrix(shibu, 6) = (Count + 1) ' Cross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If cmdIn.Tag = "" Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                            cmdIn.Tag = "A"
                        Else
                            GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                'Mean Length of Turns
                If shibu = 1 Then
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                Else
                    GCross.TextMatrix(shibu, 9) = PrcGcross(CInt(shibu), CInt(shibu) - 1)
                End If
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 0))
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                If txtCoreType <> "Special Protection" Then
                'BM
                    p = 0
                    For l = 1 To shibu
                        p = p + GCross.TextMatrix(l, 12)
                    Next
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = Val(GCross.TextMatrix(shibu, 14)) * Val(txtISF1)
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = Val(GCross.TextMatrix(shibu, 14)) * Val(txtISF1)
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 2)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                        
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Layers = Layers + 1
                Exit For
            ElseIf finalt > 25 And finalt <= 50 Then
                typ = 0.5
                GCross.TextMatrix(shibu, 6) = (Count + 1) 'GCross.TextMatrix(shibu, 6) = (Count + typ )
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If cmdIn.Tag = "" Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                            cmdIn.Tag = "A"
                        Else
                            GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                If shibu = 1 Then
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                Else
                    GCross.TextMatrix(shibu, 9) = PrcGcross(CInt(shibu), CInt(shibu) - 1)
                End If
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 0))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                'BM
                    p = 0
                    For l = 1 To shibu
                        p = p + GCross.TextMatrix(l, 12)
                    Next
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 2)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                        
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Layers = Layers + 1
                Exit For
            ElseIf finalt > 50 And finalt <= 75 Then
                typ = 0.75
                GCross.TextMatrix(shibu, 6) = (Count + 1) 'GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If cmdIn.Tag = "" Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                            cmdIn.Tag = "A"
                        Else
                            GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                If shibu = 1 Then
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                Else
                    GCross.TextMatrix(shibu, 9) = PrcGcross(CInt(shibu), CInt(shibu) - 1)
                End If
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 0))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                'BM
                    p = 0
                    For l = 1 To shibu
                        p = p + GCross.TextMatrix(l, 12)
                    Next
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 2)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                        
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Layers = Layers + 1
                Exit For
            ElseIf finalt > 75 And finalt <= 100 Then
                typ = 1
                GCross.TextMatrix(shibu, 6) = (Count + typ) 'GCross.TextMatrix(shibu, 6) = (Count + 1)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If cmdIn.Tag = "" Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                            cmdIn.Tag = "A"
                        Else
                            GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                If shibu = 1 Then
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                Else
                    GCross.TextMatrix(shibu, 9) = PrcGcross(CInt(shibu), CInt(shibu) - 1)
                End If
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 0))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                'BM
                    p = 0
                    For l = 1 To shibu
                        p = p + GCross.TextMatrix(l, 12)
                    Next
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 2)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                        
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Layers = Layers + 1
                Exit For
            End If
        End If
        value = 0
        typ = 0
        Count = Count + 1
        Layers = Layers + 1

            
    Next
'    GCross.TextMatrix(shibu, 6) = i
Next

Dim d, n As Double
d = 1.5
n = 0
For i = 1 To GCross.Rows - 1
    d = d + (GCross.TextMatrix(i, 1) * GCross.TextMatrix(i, 6))
    n = n + GCross.TextMatrix(i, 6)
Next

n = n - 1
n = n * 0.25

d = d + n + 1.5 + 3
MsgBox d
Grid5.Tag = ""
For i = 1 To GCross.Rows - 1
    If GCross.TextMatrix(i, 1) <> txtODs Then
        Grid5.Tag = "A"
    End If
Next

If Grid5.Tag <> "A" Then
'    MsgBox "The Diameter of conductor is same for all," + vbCrLf + "Select the earlier one"
    ChK2.value = 0
End If
End Function


'Calculate layers for tap RECTANGULAR Asc for different Dia
Private Function CalLayersTapDiff()
'    cmdIN.Tag = ""
    Dim value As Double
    Dim Turnss As Double
    Dim Layers As Byte
    Dim C As Double
    Dim Count As Byte
    Dim Ins(2) As Double
    Ins(0) = 1.5
    Ins(1) = 0.25

    Dim K As Double
    
'Counting No of turns
    
Dim typ As Double
    
For shibu = 1 To GCross.Rows - 1
    
'Type of Conductor
    If UCase(txtNo) = "S" Then
        Dia = GCross.TextMatrix(shibu, 1) * 1
    ElseIf UCase(txtNo) = "D" Then
        Dia = GCross.TextMatrix(shibu, 1) * 2
    ElseIf UCase(txtNo) = "T" Then
        Dia = GCross.TextMatrix(shibu, 1) * 3
    ElseIf UCase(txtNo) = "Q" Then
        Dia = GCross.TextMatrix(shibu, 1) * 4
    End If
    
'setting turns
    Turnss = GCross.TextMatrix(shibu, 0)
    For i = 0 To CInt(GCross.TextMatrix(shibu, 0))
        K = 0

'For First limb
        
        If Layers = 0 Then
            K = K + Ins(0)
        End If
        For l = 1 To Count
            If Layers <> 0 And Count <> 0 Then
                K = K + Ins(1)
            End If
            
        Next
        
        If Count > 0 Then
            K = K + (GCross.TextMatrix(shibu, 1) * Count)
        End If
        
        For j = 1 To shibu
            K = K + Val(GCross.TextMatrix(j, 8))
        Next
        
        'Rounding of k
        If InStr(1, K, ".") > 0 Then
            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
            Else
                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
            End If
        End If
        
        i = ((CDbl(txtWL1) - 2 * (K)) * 0.9) / Dia
'Rounding of i
        If InStr(1, i, ".") > 0 Then
            i = Left(i, InStr(1, i, ".") - 1)
        End If
        
        value = value + i
        
'For secound limb

        value = value + i
        
'For third limb
        
        Dim s As Double
        s = (((CDbl(txtWW1) - 2 * (K)) * 0.9) / Dia)
        
'Rounding of s
        
        If InStr(1, s, ".") > 0 Then
            s = Left(s, InStr(1, s, ".") - 1)
        End If
        value = value + s
        
'For Fourth Limb

        value = value + s
        turns1 = Turnss
        Turnss = Turnss - value
        
        If Turnss <= 0 Then
            MsgBox ((turns1 / value) * 100)
            finalt = ((turns1 / value) * 100)
            If finalt <= 25 Then
                typ = 0.25
                GCross.TextMatrix(shibu, 6) = (Count + 1) ' Cross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If cmdIn.Tag = "" Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                            cmdIn.Tag = "A"
                        Else
                            GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                If shibu = 1 Then
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                Else
                    GCross.TextMatrix(shibu, 9) = PrcGcross(CInt(shibu), CInt(shibu) - 1)
                End If
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 0))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                If txtCoreType <> "Special Protection" Then
                'BM
                    p = 0
                    For l = 1 To shibu
                        p = p + GCross.TextMatrix(l, 12)
                    Next
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 2)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                        
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Layers = Layers + 1
                Exit For
'#########2
            ElseIf finalt > 25 And finalt <= 50 Then
                typ = 0.5
                GCross.TextMatrix(shibu, 6) = (Count + 1) 'GCross.TextMatrix(shibu, 6) = (Count + typ )
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If cmdIn.Tag = "" Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                            cmdIn.Tag = "A"
                        Else
                            GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                If shibu = 1 Then
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                Else
                    GCross.TextMatrix(shibu, 9) = PrcGcross(CInt(shibu), CInt(shibu) - 1)
                End If
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 0))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                'BM
                    p = 0
                    For l = 1 To shibu
                        p = p + GCross.TextMatrix(l, 12)
                    Next
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 2)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                        
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Layers = Layers + 1
                Exit For
            ElseIf finalt > 50 And finalt <= 75 Then
                typ = 0.75
                GCross.TextMatrix(shibu, 6) = (Count + 1) 'GCross.TextMatrix(shibu, 6) = (Count + typ)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If cmdIn.Tag = "" Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                            cmdIn.Tag = "A"
                        Else
                            GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                If shibu = 1 Then
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                Else
                    GCross.TextMatrix(shibu, 9) = PrcGcross(CInt(shibu), CInt(shibu) - 1)
                End If
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 0))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                'BM
                    p = 0
                    For l = 1 To shibu
                        p = p + GCross.TextMatrix(l, 12)
                    Next
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 2)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                        
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Layers = Layers + 1
                Exit For
            ElseIf finalt > 75 And finalt <= 100 Then
                typ = 1
                GCross.TextMatrix(shibu, 6) = (Count + typ) 'GCross.TextMatrix(shibu, 6) = (Count + 1)
                'Total Winding HT
                    If (GCross.Rows - 1) = shibu Then
                        GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
                    Else
                        If cmdIn.Tag = "" Then
                            GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                            cmdIn.Tag = "A"
                        Else
                            GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 8) = Round(GCross.TextMatrix(shibu, 8), 2)
                
                'Mean Length of Turns
                If shibu = 1 Then
                    GCross.TextMatrix(shibu, 9) = (GCross.TextMatrix(shibu, 8) * 3.142) + ((2 * txtBUP) + (2 * txtHt))
                Else
                    GCross.TextMatrix(shibu, 9) = PrcGcross(CInt(shibu), CInt(shibu) - 1)
                End If
                    GCross.TextMatrix(shibu, 9) = Round(GCross.TextMatrix(shibu, 9), 2)
                
                'Total Length of wire
                    GCross.TextMatrix(shibu, 10) = CDbl(GCross.TextMatrix(shibu, 9)) * CDbl(GCross.TextMatrix(shibu, 0))
                
                'Secondary Resistances
                    '@ 20 Degee Cen
                    GCross.TextMatrix(shibu, 11) = 0.0177 * (GCross.TextMatrix(shibu, 10) / 1000) / GCross.TextMatrix(shibu, 3)
                    GCross.TextMatrix(shibu, 11) = Round(GCross.TextMatrix(shibu, 11), 3)
                    
                '@ 75 Degee Cen
                    GCross.TextMatrix(shibu, 12) = GCross.TextMatrix(shibu, 11) * 1.2161
                    GCross.TextMatrix(shibu, 12) = Round(GCross.TextMatrix(shibu, 12), 3)
                    
                'Copper Weigth
                    GCross.TextMatrix(shibu, 13) = ((CDbl(GCross.TextMatrix(shibu, 9)) * GCross.TextMatrix(shibu, 7) * txtODn * 8.9) / 1000000) * 1.1
                    GCross.TextMatrix(shibu, 13) = Round(GCross.TextMatrix(shibu, 13), 2)
                    
                If txtCoreType <> "Special Protection" Then
                'BM
                    p = 0
                    For l = 1 To shibu
                        p = p + GCross.TextMatrix(l, 12)
                    Next
'                    GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                    If cmbFreq = "50" Then
                'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 4.5 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    ElseIf cmbFreq = "60" Then
                'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
                        GCross.TextMatrix(shibu, 14) = (txtVA1 + (Sqr(txtRatioR2(cmbSelctCore - 1)) * p)) * 3.75 * 100000 / (GCross.TextMatrix(shibu, 7) * txtEArea)
                        'Calculate ISF
                        If txtCoreType = "Metering" Then
                            GCross.TextMatrix(shibu, 19) = 21000 / GCross.TextMatrix(shibu, 14)
                            GCross.TextMatrix(shibu, 19) = Round(GCross.TextMatrix(shibu, 19), 2)
                        Else
                    '####Calculation of Composite Error at Normal
                            GCross.TextMatrix(shibu, 20) = CompNormal(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7))
                            GCross.TextMatrix(shibu, 21) = CompNormalPartII(GCross.TextMatrix(shibu, 20), txtRatioR2(cmbSelctCore - 1))
                    '####Calculation of Composite Error at ALF
                            GCross.TextMatrix(shibu, 22) = CompALF(GCross.TextMatrix(shibu, 14), txtMean, GCross.TextMatrix(shibu, 7), txtISF1)
                            C = GCross.TextMatrix(shibu, 14) * txtISF1
                            GCross.TextMatrix(shibu, 23) = CompALFPartII(GCross.TextMatrix(shibu, 22), txtRatioR2(cmbSelctCore - 1), CDbl(C), txtISF1)
                        End If
                    End If
                    GCross.TextMatrix(shibu, 14) = Round(GCross.TextMatrix(shibu, 14), 2)
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        GCross.TextMatrix(shibu, 15) = Round(GCross.TextMatrix(shibu, 15), 3)
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = GCross.TextMatrix(shibu, 15) * 96
                        
                        GCross.TextMatrix(shibu, 15) = ""
                        GCross.TextMatrix(shibu, 16) = ""
                        
                        
                Else
                        If cmbFreq = "50" Then
                    'FREQ @ 50 c1 = Aeff * no of turns / 3000
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 3000
                        ElseIf cmbFreq = "60" Then
                    'FREQ @ 60c1 = Aeff * no of turns / 2500
                            GCross.TextMatrix(shibu, 15) = txtEArea * GCross.TextMatrix(shibu, 7) / 2500
                        End If
                        
                        GCross.TextMatrix(shibu, 16) = CDbl(txtMean) / CDbl(GCross.TextMatrix(shibu, 7))
                        If InStr(1, GCross.TextMatrix(shibu, 16), "E") > 0 Then
                            GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        End If
                        'C2 with factor 1.3 * c2
                        GCross.TextMatrix(shibu, 16) = GCross.TextMatrix(shibu, 16) * 1.3
                        GCross.TextMatrix(shibu, 16) = Round(GCross.TextMatrix(shibu, 16), 3)
                        
                        GCross.TextMatrix(shibu, 17) = tGCross.TextMatrix(shibu, 15) * 96
                        
                        'Excitation Current various point
                        If txtV1 = "Vk/4" Then
                            GCross.TextMatrix(shibu, 18) = 0.092 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "Vk/2" Then
                            GCross.TextMatrix(shibu, 18) = 0.145 * GCross.TextMatrix(shibu, 16)
                        ElseIf txtV1 = "" Then
                            GCross.TextMatrix(shibu, 18) = 0.34 * GCross.TextMatrix(shibu, 16)
                        End If
                        
                End If
                GCross.TextMatrix(shibu, 24) = FindDelta(Left(cmbSTC, InStr(1, UCase(cmbSTC), "K") - 1), GCross.TextMatrix(shibu, 7), txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 3))
                If cmbMetal.ListIndex <> 3 Then
                    GCross.TextMatrix(shibu, 25) = Saturated(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                Else
                    GCross.TextMatrix(shibu, 25) = SaturatedMU(CInt(cmbFreq), GCross.TextMatrix(shibu, 26), txtTurns, txtEArea, txtRatioR2(cmbSelctCore - 1), GCross.TextMatrix(shibu, 12), GCross.TextMatrix(shibu, 3))
                End If
                If CDbl(GCross.TextMatrix(shibu, 24)) > CDbl(GCross.TextMatrix(shibu, 25)) Then
                    LowDen = GCross.TextMatrix(shibu, 25)
                Else
                    LowDen = GCross.TextMatrix(shibu, 24)
                End If
                GCross.TextMatrix(shibu, 27) = 6.34 * (LowDen * LowDen) * cmbTime
                value = 0
                Count = 0
                Layers = Layers + 1
                Exit For
            End If
        End If
        value = 0
        typ = 0
        Count = Count + 1
        Layers = Layers + 1

            
    Next
'    GCross.TextMatrix(shibu, 6) = i
Next

Dim d, n As Double
d = 1.5
n = 0
For i = 1 To GCross.Rows - 1
    d = d + (GCross.TextMatrix(i, 1) * GCross.TextMatrix(i, 6))
    n = n + GCross.TextMatrix(i, 6)
Next

n = n - 1
n = n * 0.25

d = d + n + 1.5 + 3
MsgBox d
Grid5.Tag = ""
For i = 1 To GCross.Rows - 1
    If GCross.TextMatrix(i, 1) <> txtODs Then
        Grid5.Tag = "A"
    End If
Next

If Grid5.Tag <> "A" Then
'    MsgBox "The Diameter of conductor is same for all," + vbCrLf + "Select the earlier one"
    ChK2.value = 0
End If

End Function

Private Function PrcGcross(shibu As Integer, shibu1 As Integer) As Double

Dim W1, W2, R As Double
W1 = 0
For i = 1 To shibu1
    W1 = W1 + CDbl(GCross.TextMatrix(i, 8))
Next
W2 = GCross.TextMatrix(shibu, 8)
R = ((W2 / 2) + W1) * (2 * 3.142) + (2 * txtBUP) + (2 * txtHt)
PrcGcross = R
End Function

'CalCulate layers Reound
Private Function CalLayersRound(Dia As Double) As Double
        
    Dim value As Double
    
    Dim Turnss As Double
    Turnss = txtSecTurn
        
    Dim Count As Byte
    Dim Ins(20) As Double
    Ins(0) = 1.5
    Ins(1) = 0.25
    Ins(2) = 0.25
    Ins(3) = 0.25
    Ins(4) = 0.25
    Ins(5) = 0.25
    Ins(6) = 0.25
    Ins(7) = 0.25
    Ins(8) = 0.25
    Ins(9) = 0.25
    Ins(10) = 0.25
    Ins(11) = 0.25
    Ins(12) = 0.25
    Ins(13) = 0.25
    Ins(14) = 0.25
    Ins(15) = 0.25
    Ins(16) = 0.25
    Ins(17) = 0.25
    Ins(18) = 0.25
    Ins(19) = 0.25
    Ins(20) = 3
    Dim K As Double
    
'Type of Conductor
    If UCase(txtNo) = "S" Then
        Dia = Dia * 1
    ElseIf UCase(txtNo) = "D" Then
        Dia = Dia * 2
    ElseIf UCase(txtNo) = "T" Then
        Dia = Dia * 3
    ElseIf UCase(txtNo) = "Q" Then
        Dia = Dia * 4
    End If
'Counting No of turns
    
    Dim typ As Double
    
    For i = 0 To CInt(txtSecTurn)
        K = 0
'For First limb
        For l = 0 To Count
            K = K + Ins(l)
        Next
        If Count > 0 Then
            K = K + (txtODs * Count)
        End If
        
        'Rounding of k
        If InStr(1, K, ".") > 0 Then
            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
            Else
                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
            End If
        End If
        
        i = ((CDbl(txtID1) - 2 * (K)) * 0.9 * 3.142) / Dia
        'Rounding of i
        If InStr(1, i, ".") > 0 Then
            i = Left(i, InStr(1, i, ".") - 1)
        End If
        
        value = value + i
        
'        If value >= CInt(txtSecTurn) Then
'             typ = 0.25
'             CalLayers = (Count + typ)
'             Exit Function
'        End If
        
'For secound limb

'        value = value + i

        turns1 = Turnss
        Turnss = Turnss - value
        
        If Turnss <= 0 Then
            MsgBox ((turns1 / value) * 100)
            finalt = ((turns1 / value) * 100)
            If finalt <= 25 Then
                typ = 0.25
                CalLayersRound = (Count + typ)
                Exit Function
            ElseIf finalt > 25 And finalt <= 50 Then
                typ = 0.5
                CalLayersRound = (Count + typ)
                Exit Function
            ElseIf finalt > 50 And finalt <= 75 Then
                typ = 0.75
                CalLayersRound = (Count + typ)
                Exit Function
            ElseIf finalt > 75 And finalt <= 100 Then
                typ = 1
                CalLayersRound = (Count + typ)
                Exit Function
            End If
        End If
        value = 0
        typ = 0
        Count = Count + 1
'        If value >= CInt(txtSecTurn) Then
'             CalLayers = (Count + typ)
'             Exit Function
'        End If

            
    Next
    
    CalLayersRound = i
    
End Function


'Protection Special
Private Sub txtID_Change()

'If Val(txtOD) <= Val(txtID) Then
'
'    MsgBox "Not a valid core size!" + vbCrLf + "OD should be greater than ID", vbExclamation
'    txtID = 0
'    txtID.SetFocus
'    Exit Sub
'
'End If

If Len(txtOD) = 0 Or Len(txtID) = 0 Then
    txtOD.SetFocus
    Exit Sub

End If

Dim Mean, BU As Double

'finding build up

BU = prcCalBU(txtOD, txtID)

txtBUP = BU

txtBUP.Text = BU

If Len(txtOD) = 0 Or Len(txtID) = 0 Then

    Exit Sub

End If

'finding lm mean

Mean = prcCalMeanCircular(txtOD, txtID)

txtMean.Text = Round(Mean, 2)

End Sub

Private Sub txtID_GotFocus()
txtID.SelStart = 0
txtID.SelLength = Len(txtID)
MDIDesign.StatusBar1.Panels(2).Text = "Set ID"
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtID_LostFocus()

If Val(txtOD) < Val(txtID) Then
    MsgBox "Please confirm core size!", vbCritical
    txtID.SetFocus
    Exit Sub
End If

If Len(txtOD) = 0 Or Len(txtID) = 0 Then
    
    Exit Sub

End If

Dim Mean, BU As Double

'finding build up

BU = prcCalBU(txtOD, txtID)

txtBUP = BU

txtBUP.Text = BU

If Len(txtOD) = 0 Or Len(txtID) = 0 Then

    Exit Sub

End If

'finding lm mean

Mean = prcCalMeanCircular(txtOD, txtID)

txtMean.Text = Round(Mean, 2)

End Sub



Private Sub txtIDATurns_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "I Dynamic Amper Turns (kA)"
End Sub

Private Sub txtIDyn_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "I Dynamic Amper(kA)"
End Sub

Private Sub txtIe_GotFocus()
txtIe.SelStart = 0
txtIe.SelLength = Len(txtIe)
MDIDesign.StatusBar1.Panels(2).Text = "Set Ie"
End Sub

Private Sub txtIe2_GotFocus()
txtIe2.SelStart = 0
txtIe2.SelLength = Len(txtIe2)
MDIDesign.StatusBar1.Panels(2).Text = "Set Ie @"
End Sub

Private Sub txtIe2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtMean_Change()
'prcAreaEff(Ht As Double, BUP As Double) As Double
End Sub


'Private Function FullBurdenProtection5p(BM As Double, mincurr As Double) As Double
'
'MsgBox Trig(txtVA1, txtRatioR2(cmbSelctCore - 1), txtRec75)
'    If cmbFreq = "50" Then
'        BM = txtBM50
'    ElseIf cmbFreq = "60" Then
'        BM = txtBM60
'    End If
'    ErrorRead.Visible = True
'    ErrorRead.Cols = 16
'    ErrorRead.Rows = 1
'    ErrorRead.AllowUserResizing = flexResizeBoth
'
'    For i = 0 To ErrorRead.Cols - 1
'        ErrorRead.ColWidth(i) = 0
'    Next
'
'    ErrorRead.TextMatrix(0, 0) = "%PriCurr"
'    ErrorRead.ColWidth(0) = 1000
'    ErrorRead.TextMatrix(0, 1) = "BM"
'    ErrorRead.ColWidth(1) = 1000
'    ErrorRead.TextMatrix(0, 9) = "Ratio Error"
'    ErrorRead.ColWidth(9) = 1000
'    ErrorRead.TextMatrix(0, 15) = "Phase Error"
'    ErrorRead.ColWidth(15) = 1000
'
'
'    ErrorRead.Rows = ErrorRead.Rows + 1
'    ErrorRead.row = ErrorRead.Rows - 1
'
'    ErrorRead.TextMatrix(ErrorRead.row, 0) = 1
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = BM * 1
'    ErrorRead.TextMatrix(ErrorRead.row, 2) = (mincurr * 1) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead.TextMatrix(ErrorRead.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead.TextMatrix(ErrorRead.row, 3) = rst!t_awm
'        ErrorRead.TextMatrix(ErrorRead.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = (ErrorRead.TextMatrix(ErrorRead.row, 3) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = Round(ErrorRead.TextMatrix(ErrorRead.row, 5), 5)
'    'Iw
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = (ErrorRead.TextMatrix(ErrorRead.row, 4) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = ErrorRead.TextMatrix(ErrorRead.row, 5) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = Round(ErrorRead.TextMatrix(ErrorRead.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = ErrorRead.TextMatrix(ErrorRead.row, 6) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = Round(ErrorRead.TextMatrix(ErrorRead.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = -(CDbl(ErrorRead.TextMatrix(ErrorRead.row, 7)) + CDbl(ErrorRead.TextMatrix(ErrorRead.row, 8)))
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = Round(ErrorRead.TextMatrix(ErrorRead.row, 9), 3)
'    'Im Cos D
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = ErrorRead.TextMatrix(ErrorRead.row, 5) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = Round(ErrorRead.TextMatrix(ErrorRead.row, 10), 5)
'    'Iw Sin D
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = ErrorRead.TextMatrix(ErrorRead.row, 6) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = Round(ErrorRead.TextMatrix(ErrorRead.row, 11), 5)
'    ErrorRead.TextMatrix(ErrorRead.row, 12) = 1
'    ErrorRead.TextMatrix(ErrorRead.row, 13) = 2
'    ErrorRead.TextMatrix(ErrorRead.row, 14) = 3
'    'Phase Error
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = 34.4 * (ErrorRead.TextMatrix(ErrorRead.row, 10) - ErrorRead.TextMatrix(ErrorRead.row, 11))
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = Round(ErrorRead.TextMatrix(ErrorRead.row, 15), 3)
'
'End Function

'Private Function FullBurdenProtection10p(BM As Double, mincurr As Double) As Double
'
'MsgBox Trig(txtVA1, txtRatioR2(cmbSelctCore - 1), txtRec75)
'    If cmbFreq = "50" Then
'        BM = txtBM50
'    ElseIf cmbFreq = "60" Then
'        BM = txtBM60
'    End If
'    ErrorRead.Visible = True
'    ErrorRead.Cols = 16
'    ErrorRead.Rows = 1
'    ErrorRead.AllowUserResizing = flexResizeBoth
'
'    For i = 0 To ErrorRead.Cols - 1
'        ErrorRead.ColWidth(i) = 0
'    Next
'
'    ErrorRead.TextMatrix(0, 0) = "%PriCurr"
'    ErrorRead.ColWidth(0) = 800
'    ErrorRead.TextMatrix(0, 1) = "BM"
'    ErrorRead.ColWidth(1) = 1000
'    ErrorRead.TextMatrix(0, 9) = "Ratio Error"
'    ErrorRead.ColWidth(9) = 1000
'    ErrorRead.TextMatrix(0, 15) = "Phase Error"
'    ErrorRead.ColWidth(15) = 1000
'
'    ErrorRead.Rows = ErrorRead.Rows + 1
'    ErrorRead.row = ErrorRead.Rows - 1
'
'    ErrorRead.TextMatrix(ErrorRead.row, 0) = 1
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = BM * 1
'    ErrorRead.TextMatrix(ErrorRead.row, 2) = (mincurr * 1) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead.TextMatrix(ErrorRead.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead.TextMatrix(ErrorRead.row, 3) = rst!t_awm
'        ErrorRead.TextMatrix(ErrorRead.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = (ErrorRead.TextMatrix(ErrorRead.row, 3) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = Round(ErrorRead.TextMatrix(ErrorRead.row, 5), 5)
'    'Iw
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = (ErrorRead.TextMatrix(ErrorRead.row, 4) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = ErrorRead.TextMatrix(ErrorRead.row, 5) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = Round(ErrorRead.TextMatrix(ErrorRead.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = ErrorRead.TextMatrix(ErrorRead.row, 6) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = Round(ErrorRead.TextMatrix(ErrorRead.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = -(CDbl(ErrorRead.TextMatrix(ErrorRead.row, 7)) + CDbl(ErrorRead.TextMatrix(ErrorRead.row, 8)))
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = Round(ErrorRead.TextMatrix(ErrorRead.row, 9), 3)
''    'Im Cos D
''    ErrorRead.TextMatrix(ErrorRead.Row, 10) = ErrorRead.TextMatrix(ErrorRead.Row, 5) * cosD
''    ErrorRead.TextMatrix(ErrorRead.Row, 10) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 10), 5)
''    'Iw Sin D
''    ErrorRead.TextMatrix(ErrorRead.Row, 11) = ErrorRead.TextMatrix(ErrorRead.Row, 6) * sinD
''    ErrorRead.TextMatrix(ErrorRead.Row, 11) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 11), 5)
''    ErrorRead.TextMatrix(ErrorRead.Row, 12) = 1
''    ErrorRead.TextMatrix(ErrorRead.Row, 13) = 2
''    ErrorRead.TextMatrix(ErrorRead.Row, 14) = 3
''    'Phase Error
''    ErrorRead.TextMatrix(ErrorRead.Row, 15) = 34.4 * (ErrorRead.TextMatrix(ErrorRead.Row, 10) - ErrorRead.TextMatrix(ErrorRead.Row, 11))
''    ErrorRead.TextMatrix(ErrorRead.Row, 15) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 15), 3)
'
'End Function

'Private Function FullBurdenMetering(BM As Double, mincurr As Double) As Double
'
'MsgBox Trig(txtVA1, txtRatioR2(cmbSelctCore - 1), txtRec75)
'    If cmbFreq = "50" Then
'        BM = txtBM50
'    ElseIf cmbFreq = "60" Then
'        BM = txtBM60
'    End If
'    ErrorRead.Visible = True
'    ErrorRead.Cols = 16
'    ErrorRead.Rows = 1
'    ErrorRead.AllowUserResizing = flexResizeBoth
'
'    For i = 0 To ErrorRead.Cols - 1
'        ErrorRead.ColWidth(i) = 0
'    Next
'
'    ErrorRead.TextMatrix(0, 0) = "%PriCurr"
'    ErrorRead.ColWidth(0) = 800
'    ErrorRead.TextMatrix(0, 1) = "BM"
'    ErrorRead.ColWidth(1) = 1000
'    ErrorRead.TextMatrix(0, 9) = "Ratio Error"
'    ErrorRead.ColWidth(9) = 1000
'    ErrorRead.TextMatrix(0, 15) = "Phase Error"
'    ErrorRead.ColWidth(15) = 1000
'
'    ErrorRead.Rows = ErrorRead.Rows + 1
'    ErrorRead.row = ErrorRead.Rows - 1
'
'    ErrorRead.TextMatrix(ErrorRead.row, 0) = 1.2
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = BM * 1.2
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = Round(ErrorRead.TextMatrix(ErrorRead.row, 1), 3)
'    ErrorRead.TextMatrix(ErrorRead.row, 2) = (mincurr * 1.2) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead.TextMatrix(ErrorRead.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead.TextMatrix(ErrorRead.row, 3) = rst!t_awm
'        ErrorRead.TextMatrix(ErrorRead.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = (ErrorRead.TextMatrix(ErrorRead.row, 3) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = Round(ErrorRead.TextMatrix(ErrorRead.row, 5), 5)
'    'Iw
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = (ErrorRead.TextMatrix(ErrorRead.row, 4) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = Round(ErrorRead.TextMatrix(ErrorRead.row, 6), 5)
'    'Im Sin Delta--------x
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = ErrorRead.TextMatrix(ErrorRead.row, 5) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = Round(ErrorRead.TextMatrix(ErrorRead.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = ErrorRead.TextMatrix(ErrorRead.row, 6) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = Round(ErrorRead.TextMatrix(ErrorRead.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = -(CDbl(ErrorRead.TextMatrix(ErrorRead.row, 7)) + CDbl(ErrorRead.TextMatrix(ErrorRead.row, 8)))
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = Round(ErrorRead.TextMatrix(ErrorRead.row, 9), 3)
'    'Im Cos D
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = ErrorRead.TextMatrix(ErrorRead.row, 5) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = Round(ErrorRead.TextMatrix(ErrorRead.row, 10), 5)
'    'Iw Sin D
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = ErrorRead.TextMatrix(ErrorRead.row, 6) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = Round(ErrorRead.TextMatrix(ErrorRead.row, 11), 5)
'    ErrorRead.TextMatrix(ErrorRead.row, 12) = 1
'    ErrorRead.TextMatrix(ErrorRead.row, 13) = 2
'    ErrorRead.TextMatrix(ErrorRead.row, 14) = 3
'    'Phase Error
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = 34.4 * (ErrorRead.TextMatrix(ErrorRead.row, 10) - ErrorRead.TextMatrix(ErrorRead.row, 11))
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = Round(ErrorRead.TextMatrix(ErrorRead.row, 15), 3)
''########2
'    ErrorRead.Rows = ErrorRead.Rows + 1
'    ErrorRead.row = ErrorRead.Rows - 1
'
'    ErrorRead.TextMatrix(ErrorRead.row, 0) = 1
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = BM * 1
'    ErrorRead.TextMatrix(ErrorRead.row, 2) = (mincurr * 1) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead.TextMatrix(ErrorRead.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead.TextMatrix(ErrorRead.row, 3) = rst!t_awm
'        ErrorRead.TextMatrix(ErrorRead.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = (ErrorRead.TextMatrix(ErrorRead.row, 3) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = Round(ErrorRead.TextMatrix(ErrorRead.row, 5), 5)
'    'Iw
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = (ErrorRead.TextMatrix(ErrorRead.row, 4) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = ErrorRead.TextMatrix(ErrorRead.row, 5) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = Round(ErrorRead.TextMatrix(ErrorRead.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = ErrorRead.TextMatrix(ErrorRead.row, 6) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = Round(ErrorRead.TextMatrix(ErrorRead.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = -(CDbl(ErrorRead.TextMatrix(ErrorRead.row, 7)) + CDbl(ErrorRead.TextMatrix(ErrorRead.row, 8)))
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = Round(ErrorRead.TextMatrix(ErrorRead.row, 9), 3)
'    'Im Cos D
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = ErrorRead.TextMatrix(ErrorRead.row, 5) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = Round(ErrorRead.TextMatrix(ErrorRead.row, 10), 5)
'    'Iw Sin D
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = ErrorRead.TextMatrix(ErrorRead.row, 6) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = Round(ErrorRead.TextMatrix(ErrorRead.row, 11), 5)
'    ErrorRead.TextMatrix(ErrorRead.row, 12) = 1
'    ErrorRead.TextMatrix(ErrorRead.row, 13) = 2
'    ErrorRead.TextMatrix(ErrorRead.row, 14) = 3
'    'Phase Error
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = 34.4 * (ErrorRead.TextMatrix(ErrorRead.row, 10) - ErrorRead.TextMatrix(ErrorRead.row, 11))
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = Round(ErrorRead.TextMatrix(ErrorRead.row, 15), 3)
''##########3
'    ErrorRead.Rows = ErrorRead.Rows + 1
'    ErrorRead.row = ErrorRead.Rows - 1
'
'    ErrorRead.TextMatrix(ErrorRead.row, 0) = 0.2
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = BM * 0.2
'    ErrorRead.TextMatrix(ErrorRead.row, 2) = (mincurr * 0.2) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead.TextMatrix(ErrorRead.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead.TextMatrix(ErrorRead.row, 3) = rst!t_awm
'        ErrorRead.TextMatrix(ErrorRead.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = (ErrorRead.TextMatrix(ErrorRead.row, 3) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = Round(ErrorRead.TextMatrix(ErrorRead.row, 5), 5)
'    'Iw
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = (ErrorRead.TextMatrix(ErrorRead.row, 4) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = ErrorRead.TextMatrix(ErrorRead.row, 5) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = Round(ErrorRead.TextMatrix(ErrorRead.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = ErrorRead.TextMatrix(ErrorRead.row, 6) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = Round(ErrorRead.TextMatrix(ErrorRead.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = -(CDbl(ErrorRead.TextMatrix(ErrorRead.row, 7)) + CDbl(ErrorRead.TextMatrix(ErrorRead.row, 8)))
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = Round(ErrorRead.TextMatrix(ErrorRead.row, 9), 3)
'    'Im Cos D
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = ErrorRead.TextMatrix(ErrorRead.row, 5) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = Round(ErrorRead.TextMatrix(ErrorRead.row, 10), 5)
'    'Iw Sin D
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = ErrorRead.TextMatrix(ErrorRead.row, 6) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = Round(ErrorRead.TextMatrix(ErrorRead.row, 11), 5)
'    ErrorRead.TextMatrix(ErrorRead.row, 12) = 1
'    ErrorRead.TextMatrix(ErrorRead.row, 13) = 2
'    ErrorRead.TextMatrix(ErrorRead.row, 14) = 3
'    'Phase Error
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = 34.4 * (ErrorRead.TextMatrix(ErrorRead.row, 10) - ErrorRead.TextMatrix(ErrorRead.row, 11))
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = Round(ErrorRead.TextMatrix(ErrorRead.row, 15), 3)
''########4
'    ErrorRead.Rows = ErrorRead.Rows + 1
'    ErrorRead.row = ErrorRead.Rows - 1
'
'    ErrorRead.TextMatrix(ErrorRead.row, 0) = 0.05
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = BM * 0.05
'    ErrorRead.TextMatrix(ErrorRead.row, 2) = (mincurr * 0.05) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead.TextMatrix(ErrorRead.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead.TextMatrix(ErrorRead.row, 3) = rst!t_awm
'        ErrorRead.TextMatrix(ErrorRead.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = (ErrorRead.TextMatrix(ErrorRead.row, 3) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = Round(ErrorRead.TextMatrix(ErrorRead.row, 5), 5)
'    'Iw
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = (ErrorRead.TextMatrix(ErrorRead.row, 4) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = ErrorRead.TextMatrix(ErrorRead.row, 5) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = Round(ErrorRead.TextMatrix(ErrorRead.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = ErrorRead.TextMatrix(ErrorRead.row, 6) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = Round(ErrorRead.TextMatrix(ErrorRead.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = -(CDbl(ErrorRead.TextMatrix(ErrorRead.row, 7)) + CDbl(ErrorRead.TextMatrix(ErrorRead.row, 8)))
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = Round(ErrorRead.TextMatrix(ErrorRead.row, 9), 3)
'    'Im Cos D
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = ErrorRead.TextMatrix(ErrorRead.row, 5) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 10) = Round(ErrorRead.TextMatrix(ErrorRead.row, 10), 5)
'    'Iw Sin D
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = ErrorRead.TextMatrix(ErrorRead.row, 6) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 11) = Round(ErrorRead.TextMatrix(ErrorRead.row, 11), 5)
'    ErrorRead.TextMatrix(ErrorRead.row, 12) = 1
'    ErrorRead.TextMatrix(ErrorRead.row, 13) = 2
'    ErrorRead.TextMatrix(ErrorRead.row, 14) = 3
'    'Phase Error
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = 34.4 * (ErrorRead.TextMatrix(ErrorRead.row, 10) - ErrorRead.TextMatrix(ErrorRead.row, 11))
'    ErrorRead.TextMatrix(ErrorRead.row, 15) = Round(ErrorRead.TextMatrix(ErrorRead.row, 15), 3)
'
'
'End Function

'Private Function FullBurdenMeteringC2(BM As Double, mincurr As Double) As Double
'
''############# 120 %
'
'MsgBox Trig(txtVA1, txtRatioR2(cmbSelctCore - 1), txtRec75)
'    If cmbFreq = "50" Then
'        BM = txtBM50
'    ElseIf cmbFreq = "60" Then
'        BM = txtBM60
'    End If
'    ErrorRead.Visible = True
'    ErrorRead.Cols = 16
'    ErrorRead.Rows = 1
'    ErrorRead.AllowUserResizing = flexResizeBoth
'
'    For i = 0 To ErrorRead.Cols - 1
'        ErrorRead.ColWidth(i) = 0
'    Next
'
'    ErrorRead.TextMatrix(0, 0) = "%PriCurr"
'    ErrorRead.ColWidth(0) = 800
'    ErrorRead.TextMatrix(0, 1) = "BM"
'    ErrorRead.ColWidth(1) = 1000
'    ErrorRead.TextMatrix(0, 9) = "Ratio Error"
'    ErrorRead.ColWidth(9) = 1000
'    ErrorRead.TextMatrix(0, 15) = "Phase Error"
'    ErrorRead.ColWidth(15) = 1000
'
'    ErrorRead.Rows = ErrorRead.Rows + 1
'    ErrorRead.row = ErrorRead.Rows - 1
'
'    ErrorRead.TextMatrix(ErrorRead.row, 0) = 1.2
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = BM * 1.2
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = Round(ErrorRead.TextMatrix(ErrorRead.row, 1), 3)
'    ErrorRead.TextMatrix(ErrorRead.row, 2) = (mincurr * 1.2) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead.TextMatrix(ErrorRead.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead.TextMatrix(ErrorRead.row, 3) = rst!t_awm
'        ErrorRead.TextMatrix(ErrorRead.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = (ErrorRead.TextMatrix(ErrorRead.row, 3) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = Round(ErrorRead.TextMatrix(ErrorRead.row, 5), 5)
'    'Iw
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = (ErrorRead.TextMatrix(ErrorRead.row, 4) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = Round(ErrorRead.TextMatrix(ErrorRead.row, 6), 5)
'    'Im Sin Delta--------x
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = ErrorRead.TextMatrix(ErrorRead.row, 5) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = Round(ErrorRead.TextMatrix(ErrorRead.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = ErrorRead.TextMatrix(ErrorRead.row, 6) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = Round(ErrorRead.TextMatrix(ErrorRead.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = -(CDbl(ErrorRead.TextMatrix(ErrorRead.row, 7)) + CDbl(ErrorRead.TextMatrix(ErrorRead.row, 8)))
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = Round(ErrorRead.TextMatrix(ErrorRead.row, 9), 3)
''    'Im Cos D
''    ErrorRead.TextMatrix(ErrorRead.Row, 10) = ErrorRead.TextMatrix(ErrorRead.Row, 5) * cosD
''    ErrorRead.TextMatrix(ErrorRead.Row, 10) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 10), 5)
''    'Iw Sin D
''    ErrorRead.TextMatrix(ErrorRead.Row, 11) = ErrorRead.TextMatrix(ErrorRead.Row, 6) * sinD
''    ErrorRead.TextMatrix(ErrorRead.Row, 11) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 11), 5)
''    ErrorRead.TextMatrix(ErrorRead.Row, 12) = 1
''    ErrorRead.TextMatrix(ErrorRead.Row, 13) = 2
''    ErrorRead.TextMatrix(ErrorRead.Row, 14) = 3
''    'Phase Error
''    ErrorRead.TextMatrix(ErrorRead.Row, 15) = 34.4 * (ErrorRead.TextMatrix(ErrorRead.Row, 10) - ErrorRead.TextMatrix(ErrorRead.Row, 11))
''    ErrorRead.TextMatrix(ErrorRead.Row, 15) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 15), 3)
''########2 50 %
'    ErrorRead.Rows = ErrorRead.Rows + 1
'    ErrorRead.row = ErrorRead.Rows - 1
'
'    ErrorRead.TextMatrix(ErrorRead.row, 0) = 1
'    ErrorRead.TextMatrix(ErrorRead.row, 1) = BM * 0.5
'    ErrorRead.TextMatrix(ErrorRead.row, 2) = (mincurr * 0.5) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead.TextMatrix(ErrorRead.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead.TextMatrix(ErrorRead.row, 3) = rst!t_awm
'        ErrorRead.TextMatrix(ErrorRead.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = (ErrorRead.TextMatrix(ErrorRead.row, 3) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    ErrorRead.TextMatrix(ErrorRead.row, 5) = Round(ErrorRead.TextMatrix(ErrorRead.row, 5), 5)
'    'Iw
'    ErrorRead.TextMatrix(ErrorRead.row, 6) = (ErrorRead.TextMatrix(ErrorRead.row, 4) / ErrorRead.TextMatrix(ErrorRead.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = ErrorRead.TextMatrix(ErrorRead.row, 5) * sinD
'    ErrorRead.TextMatrix(ErrorRead.row, 7) = Round(ErrorRead.TextMatrix(ErrorRead.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = ErrorRead.TextMatrix(ErrorRead.row, 6) * cosD
'    ErrorRead.TextMatrix(ErrorRead.row, 8) = Round(ErrorRead.TextMatrix(ErrorRead.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = -(CDbl(ErrorRead.TextMatrix(ErrorRead.row, 7)) + CDbl(ErrorRead.TextMatrix(ErrorRead.row, 8)))
'    ErrorRead.TextMatrix(ErrorRead.row, 9) = Round(ErrorRead.TextMatrix(ErrorRead.row, 9), 3)
''    'Im Cos D
''    ErrorRead.TextMatrix(ErrorRead.Row, 10) = ErrorRead.TextMatrix(ErrorRead.Row, 5) * cosD
''    ErrorRead.TextMatrix(ErrorRead.Row, 10) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 10), 5)
''    'Iw Sin D
''    ErrorRead.TextMatrix(ErrorRead.Row, 11) = ErrorRead.TextMatrix(ErrorRead.Row, 6) * sinD
''    ErrorRead.TextMatrix(ErrorRead.Row, 11) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 11), 5)
''    ErrorRead.TextMatrix(ErrorRead.Row, 12) = 1
''    ErrorRead.TextMatrix(ErrorRead.Row, 13) = 2
''    ErrorRead.TextMatrix(ErrorRead.Row, 14) = 3
''    'Phase Error
''    ErrorRead.TextMatrix(ErrorRead.Row, 15) = 34.4 * (ErrorRead.TextMatrix(ErrorRead.Row, 10) - ErrorRead.TextMatrix(ErrorRead.Row, 11))
''    ErrorRead.TextMatrix(ErrorRead.Row, 15) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 15), 3)
'
'
'End Function

'Private Function HalfBurdenMetering(BM As Double, mincurr As Double) As Double
'
'MsgBox Trig(txtVA1, txtRatioR2(cmbSelctCore - 1), txtRec75)
'    If cmbFreq = "50" Then
'        BM = txtBM50
'    ElseIf cmbFreq = "60" Then
'        BM = txtBM60
'    End If
'    ErrorRead1.Visible = True
'    ErrorRead1.Cols = 16
'    ErrorRead1.Rows = 1
'    ErrorRead1.AllowUserResizing = flexResizeBoth
'
'    For i = 0 To ErrorRead1.Cols - 1
'        ErrorRead1.ColWidth(i) = 0
'    Next
'
'    ErrorRead1.TextMatrix(0, 0) = "%PriCurr"
'    ErrorRead1.ColWidth(0) = 800
'    ErrorRead1.TextMatrix(0, 1) = "BM"
'    ErrorRead1.ColWidth(1) = 1000
'    ErrorRead1.TextMatrix(0, 9) = "Ratio Error"
'    ErrorRead1.ColWidth(9) = 1000
'    ErrorRead1.TextMatrix(0, 15) = "Phase Error"
'    ErrorRead1.ColWidth(15) = 1000
'
'    ErrorRead1.Rows = ErrorRead1.Rows + 1
'    ErrorRead1.row = ErrorRead1.Rows - 1
'
'    ErrorRead1.TextMatrix(ErrorRead1.row, 0) = 1.2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 1) = BM * 1.2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 2) = (mincurr * 1.2) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead1.TextMatrix(ErrorRead1.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead1.TextMatrix(ErrorRead1.row, 3) = rst!t_awm
'        ErrorRead1.TextMatrix(ErrorRead1.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = (ErrorRead1.TextMatrix(ErrorRead1.row, 3) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 5), 5)
'    'Iw
'    ErrorRead1.TextMatrix(ErrorRead1.row, 6) = (ErrorRead1.TextMatrix(ErrorRead1.row, 4) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    ErrorRead1.TextMatrix(ErrorRead1.row, 6) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 6), 5)
'    'Im Sin Delta--------x
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = -(CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 7)) + CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 8)))
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 9), 3)
''    'Im Cos D
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 10) = ErrorRead1.TextMatrix(ErrorRead1.Row, 5) * cosD
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 10) = Round(ErrorRead1.TextMatrix(ErrorRead1.Row, 10), 5)
''    'Iw Sin D
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 11) = ErrorRead1.TextMatrix(ErrorRead1.Row, 6) * sinD
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 11) = Round(ErrorRead1.TextMatrix(ErrorRead1.Row, 11), 5)
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 12) = 1
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 13) = 2
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 14) = 3
''    'Phase Error
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 15) = 34.4 * (ErrorRead1.TextMatrix(ErrorRead1.Row, 10) - ErrorRead1.TextMatrix(ErrorRead1.Row, 11))
''    ErrorRead.TextMatrix(ErrorRead.Row, 15) = Round(ErrorRead.TextMatrix(ErrorRead.Row, 9), 15)
''########2
'    ErrorRead1.Rows = ErrorRead1.Rows + 1
'    ErrorRead1.row = ErrorRead1.Rows - 1
'
'    ErrorRead1.TextMatrix(ErrorRead1.row, 0) = 1
'    ErrorRead1.TextMatrix(ErrorRead1.row, 1) = BM * 0.5
'    ErrorRead1.TextMatrix(ErrorRead1.row, 2) = (mincurr * 0.5) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead1.TextMatrix(ErrorRead1.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead1.TextMatrix(ErrorRead1.row, 3) = rst!t_awm
'        ErrorRead1.TextMatrix(ErrorRead1.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = (ErrorRead1.TextMatrix(ErrorRead1.row, 3) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 5), 5)
'    'Iw
'    ErrorRead1.TextMatrix(ErrorRead1.row, 6) = (ErrorRead1.TextMatrix(ErrorRead1.row, 4) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = -(CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 7)) + CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 8)))
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 9), 3)
''    'Im Cos D
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 10) = ErrorRead1.TextMatrix(ErrorRead1.Row, 5) * cosD
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 10) = Round(ErrorRead1.TextMatrix(ErrorRead1.Row, 10), 5)
''    'Iw Sin D
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 11) = ErrorRead1.TextMatrix(ErrorRead1.Row, 6) * sinD
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 11) = Round(ErrorRead1.TextMatrix(ErrorRead1.Row, 11), 5)
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 12) = 1
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 13) = 2
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 14) = 3
''    'Phase Error
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 15) = 34.4 * (ErrorRead1.TextMatrix(ErrorRead1.Row, 10) - ErrorRead1.TextMatrix(ErrorRead1.Row, 11))
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 15) = Round(ErrorRead1.TextMatrix(ErrorRead1.Row, 15), 3)
'End Function

'Private Function QuatarBurdenMetering(BM As Double, mincurr As Double) As Double
'
'b = Trig(txtVA1 / 4, txtRatioR2(cmbSelctCore - 1), txtRec75)
'    ErrorRead1.Visible = True
'    ErrorRead1.Cols = 16
'    ErrorRead1.Rows = 1
'    ErrorRead1.AllowUserResizing = flexResizeBoth
'
'    For i = 0 To ErrorRead1.Cols - 1
'        ErrorRead1.ColWidth(i) = 0
'    Next
'
'    ErrorRead1.TextMatrix(0, 0) = "%PriCurr"
'    ErrorRead1.ColWidth(0) = 800
'    ErrorRead1.TextMatrix(0, 1) = "BM"
'    ErrorRead1.ColWidth(1) = 1000
'    ErrorRead1.TextMatrix(0, 9) = "Ratio Error"
'    ErrorRead1.ColWidth(9) = 1000
'    ErrorRead1.TextMatrix(0, 15) = "Phase Error"
'    ErrorRead1.ColWidth(15) = 1000
'
'    ErrorRead1.Rows = ErrorRead1.Rows + 1
'    ErrorRead1.row = ErrorRead1.Rows - 1
'
'    ErrorRead1.TextMatrix(ErrorRead1.row, 0) = 1.2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 1) = BM * 1.2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 2) = (mincurr * 1.2) / txtMean
''    ErrorRead1.TextMatrix(ErrorRead1.Row, 1) = 900 '(MinCurr * 1.2) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead1.TextMatrix(ErrorRead1.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead1.TextMatrix(ErrorRead1.row, 3) = rst!t_awm
'        ErrorRead1.TextMatrix(ErrorRead1.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = (ErrorRead1.TextMatrix(ErrorRead1.row, 3) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 5), 5)
'    'Iw
'    ErrorRead1.TextMatrix(ErrorRead1.row, 6) = (ErrorRead1.TextMatrix(ErrorRead1.row, 4) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    ErrorRead1.TextMatrix(ErrorRead1.row, 6) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 6), 5)
'    'Im Sin Delta--------x
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = -(CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 7)) + CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 8)))
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 9), 3)
'    'Im Cos D
'    ErrorRead1.TextMatrix(ErrorRead1.row, 10) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 10) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 10), 5)
'    'Iw Sin D
'    ErrorRead1.TextMatrix(ErrorRead1.row, 11) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 11) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 11), 5)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 12) = 1
'    ErrorRead1.TextMatrix(ErrorRead1.row, 13) = 2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 14) = 3
'    'Phase Error
'    ErrorRead1.TextMatrix(ErrorRead1.row, 15) = 34.4 * (ErrorRead1.TextMatrix(ErrorRead1.row, 10) - ErrorRead1.TextMatrix(ErrorRead1.row, 11))
'     ErrorRead1.TextMatrix(ErrorRead1.row, 15) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 15), 3)
''########2
'    ErrorRead1.Rows = ErrorRead1.Rows + 1
'    ErrorRead1.row = ErrorRead1.Rows - 1
'
'    ErrorRead1.TextMatrix(ErrorRead1.row, 0) = 1
'    ErrorRead1.TextMatrix(ErrorRead1.row, 1) = BM * 1
'    ErrorRead1.TextMatrix(ErrorRead1.row, 2) = (mincurr * 1) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead1.TextMatrix(ErrorRead1.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead1.TextMatrix(ErrorRead1.row, 3) = rst!t_awm
'        ErrorRead1.TextMatrix(ErrorRead1.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = (ErrorRead1.TextMatrix(ErrorRead1.row, 3) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 5), 5)
'    'Iw
'    ErrorRead1.TextMatrix(ErrorRead1.row, 6) = (ErrorRead1.TextMatrix(ErrorRead1.row, 4) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = -(CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 7)) + CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 8)))
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 9), 3)
'    'Im Cos D
'    ErrorRead1.TextMatrix(ErrorRead1.row, 10) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 10) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 10), 5)
'    'Iw Sin D
'    ErrorRead1.TextMatrix(ErrorRead1.row, 11) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 11) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 11), 5)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 12) = 1
'    ErrorRead1.TextMatrix(ErrorRead1.row, 13) = 2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 14) = 3
'    'Phase Error
'    ErrorRead1.TextMatrix(ErrorRead1.row, 15) = 34.4 * (ErrorRead1.TextMatrix(ErrorRead1.row, 10) - ErrorRead1.TextMatrix(ErrorRead1.row, 11))
'    ErrorRead1.TextMatrix(ErrorRead1.row, 15) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 15), 3)
''##########3
'    ErrorRead1.Rows = ErrorRead1.Rows + 1
'    ErrorRead1.row = ErrorRead1.Rows - 1
'
'    ErrorRead1.TextMatrix(ErrorRead1.row, 0) = 0.2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 1) = BM * 0.2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 2) = (mincurr * 0.2) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead1.TextMatrix(ErrorRead1.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead1.TextMatrix(ErrorRead1.row, 3) = rst!t_awm
'        ErrorRead1.TextMatrix(ErrorRead1.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = (ErrorRead1.TextMatrix(ErrorRead1.row, 3) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 5), 5)
'    'Iw
'    ErrorRead1.TextMatrix(ErrorRead1.row, 6) = (ErrorRead1.TextMatrix(ErrorRead1.row, 4) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = -(CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 7)) + CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 8)))
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 9), 3)
'    'Im Cos D
'    ErrorRead1.TextMatrix(ErrorRead1.row, 10) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 10) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 10), 5)
'    'Iw Sin D
'    ErrorRead1.TextMatrix(ErrorRead1.row, 11) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 11) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 11), 5)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 12) = 1
'    ErrorRead1.TextMatrix(ErrorRead1.row, 13) = 2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 14) = 3
'    'Phase Error
'    ErrorRead1.TextMatrix(ErrorRead1.row, 15) = 34.4 * (ErrorRead1.TextMatrix(ErrorRead1.row, 10) - ErrorRead1.TextMatrix(ErrorRead1.row, 11))
'    ErrorRead1.TextMatrix(ErrorRead1.row, 15) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 15), 3)
''########4
'    ErrorRead1.Rows = ErrorRead1.Rows + 1
'    ErrorRead1.row = ErrorRead1.Rows - 1
'
'    ErrorRead1.TextMatrix(ErrorRead1.row, 0) = 0.05
'    ErrorRead1.TextMatrix(ErrorRead1.row, 1) = BM * 0.05
'    ErrorRead1.TextMatrix(ErrorRead1.row, 2) = (mincurr * 0.05) / txtMean
'    Set rst = New ADODB.Recordset
'    rst.Open "select t_awm, t_aww from losess where t_bm >= " & ErrorRead1.TextMatrix(ErrorRead1.row, 1) & " order by t_aw0", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        ErrorRead1.TextMatrix(ErrorRead1.row, 3) = rst!t_awm
'        ErrorRead1.TextMatrix(ErrorRead1.row, 4) = rst!t_aww
'    End If
'    Set rst = Nothing
'    'Im
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = (ErrorRead1.TextMatrix(ErrorRead1.row, 3) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    ErrorRead1.TextMatrix(ErrorRead1.row, 5) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 5), 5)
'    'Iw
'    ErrorRead1.TextMatrix(ErrorRead1.row, 6) = (ErrorRead1.TextMatrix(ErrorRead1.row, 4) / ErrorRead1.TextMatrix(ErrorRead1.row, 2)) * 100
'    'Im Sin Delta--------x
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 7) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 7), 5)
'    'Iw Cos Delta--------y
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 8) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 8), 5)
'    '% Ratio Error RE = -(x + y)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = -(CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 7)) + CDbl(ErrorRead1.TextMatrix(ErrorRead1.row, 8)))
'    ErrorRead1.TextMatrix(ErrorRead1.row, 9) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 9), 3)
'    'Im Cos D
'    ErrorRead1.TextMatrix(ErrorRead1.row, 10) = ErrorRead1.TextMatrix(ErrorRead1.row, 5) * cosD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 10) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 10), 5)
'    'Iw Sin D
'    ErrorRead1.TextMatrix(ErrorRead1.row, 11) = ErrorRead1.TextMatrix(ErrorRead1.row, 6) * sinD
'    ErrorRead1.TextMatrix(ErrorRead1.row, 11) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 11), 5)
'    ErrorRead1.TextMatrix(ErrorRead1.row, 12) = 1
'    ErrorRead1.TextMatrix(ErrorRead1.row, 13) = 2
'    ErrorRead1.TextMatrix(ErrorRead1.row, 14) = 3
'    'Phase Error
'    ErrorRead1.TextMatrix(ErrorRead1.row, 15) = 34.4 * (ErrorRead1.TextMatrix(ErrorRead1.row, 10) - ErrorRead1.TextMatrix(ErrorRead1.row, 11))
'    ErrorRead1.TextMatrix(ErrorRead1.row, 15) = Round(ErrorRead1.TextMatrix(ErrorRead1.row, 15), 3)
'
'
'End Function

'Private Function prcCalLower(Tur As Integer) As Double
'
'Dim value As Double
'
'    Dim Turnss As Double
'    Turnss = Tur
'
'    Dim Count As Double
'    Dim Ins(20) As Double
'    Ins(0) = 1.5
'    Ins(1) = 0.25
'    Ins(2) = 0.25
'    Ins(3) = 0.25
'    Ins(4) = 0.25
'    Ins(5) = 0.25
'    Ins(6) = 0.25
'    Ins(7) = 0.25
'    Ins(8) = 0.25
'    Ins(9) = 0.25
'    Ins(10) = 0.25
'    Ins(11) = 0.25
'    Ins(12) = 0.25
'    Ins(13) = 0.25
'    Ins(14) = 0.25
'    Ins(15) = 0.25
'    Ins(16) = 0.25
'    Ins(17) = 0.25
'    Ins(18) = 0.25
'    Ins(19) = 0.25
'    Ins(20) = 3
'    Dim K As Double
'
''Type of Conductor
'    If UCase(txtNo) = "S" Then
'        Dia = txtODs * 1
'    ElseIf UCase(txtNo) = "D" Then
'        Dia = txtODs * 2
'    ElseIf UCase(txtNo) = "T" Then
'        Dia = txtODs * 3
'    ElseIf UCase(txtNo) = "Q" Then
'        Dia = txtODs * 4
'    End If
''Counting No of turns
'
'    Dim typ As Double
'
'    For i = 0 To CInt(Tur)
'        K = 0
''For First limb
'        For l = 0 To Count
'            K = K + Ins(l)
'        Next
'        If Count > 0 Then
'            K = K + (txtODs * Count)
'        End If
'
'        'Rounding of k
'        If InStr(1, K, ".") > 0 Then
'            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
'                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
'            Else
'                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
'            End If
'        End If
'
'        i = ((CDbl(txtWL1) - 2 * (K)) * 0.9) / Dia
'        'Rounding of i
'        If InStr(1, i, ".") > 0 Then
'            i = Left(i, InStr(1, i, ".") - 1)
'        End If
'
'        value = value + i
'
''        If value >= CInt(txtSecTurn) Then
''             typ = 0.25
''             CalLayers = (Count + typ)
''             Exit Function
''        End If
'
''For secound limb
'
'        value = value + i
''        If value >= CInt(txtSecTurn) Then
''             typ = 0.5
''             CalLayers = (Count + typ)
''             Exit Function
''        End If
'
''For third limb
'
'        Dim s As Double
'
'        s = (((CDbl(txtWW1) - 2 * (K)) * 0.9) / Dia)
'        'Rounding of s
'        If InStr(1, s, ".") > 0 Then
'            s = Left(s, InStr(1, s, ".") - 1)
'        End If
'        value = value + s
''        If value >= CInt(txtSecTurn) Then
''             typ = 0.75
''             CalLayers = (Count + typ)
''             Exit Function
''        End If
'
''For Fourth Limb
'
'        value = value + s
'        turns1 = Turnss
'        Turnss = Turnss - value
'
'        If Turnss <= 0 Then
'            MsgBox ((turns1 / value) * 100)
'            finalt = ((turns1 / value) * 100)
'            If finalt <= 25 Then
'                typ = 0.25
'                Count = (Count + typ)
'                Exit For
'            ElseIf finalt > 25 And finalt <= 50 Then
'                typ = 0.5
'                Count = (Count + typ)
'                Exit For
'            ElseIf finalt > 50 And finalt <= 75 Then
'                typ = 0.75
'                Count = (Count + typ)
'                Exit For
'            ElseIf finalt > 75 And finalt <= 100 Then
'                typ = 1
'                Count = (Count + typ)
'                Exit For
'            End If
'        End If
'        value = 0
'        typ = 0
'        Count = Count + 1
''        If value >= CInt(txtSecTurn) Then
''             CalLayers = (Count + typ)
''             Exit Function
''        End If
'
'
'    Next
'
'    If InStr(1, Count, ".") > 0 Then
''    MsgBox Left(Count, InStr(1, Count, ".") - 1)
'        Count = Left(Count, InStr(1, Count, ".") - 1) + 1
'    End If
''Total Winding HT
'    WHt = 1.5 + (txtODs * Count) + ((Count - 1) * 0.25) + (1.5 + 3)
'    WHt = Round(WHt, 2)
'
''Mean Length of Turns
'    MLT = (WHt * 3.142) + ((2 * txtBUP) + (2 * txtHt))
'    MLT = Round(MLT, 2)
'
''Total Length of wire
'    Toleng = CDbl(MLT) * CDbl(Tur)
'    prcCalLower = Toleng
'
'
'End Function

'Private Function prcCalLowerCir(Tur As Integer) As Double
'
'Dim value As Double
'
'    Dim Turnss As Double
'    Turnss = Tur
'
'    Dim Count As Double
'    Dim Ins(20) As Double
'    Ins(0) = 1.5
'    Ins(1) = 0.25
'    Ins(2) = 0.25
'    Ins(3) = 0.25
'    Ins(4) = 0.25
'    Ins(5) = 0.25
'    Ins(6) = 0.25
'    Ins(7) = 0.25
'    Ins(8) = 0.25
'    Ins(9) = 0.25
'    Ins(10) = 0.25
'    Ins(11) = 0.25
'    Ins(12) = 0.25
'    Ins(13) = 0.25
'    Ins(14) = 0.25
'    Ins(15) = 0.25
'    Ins(16) = 0.25
'    Ins(17) = 0.25
'    Ins(18) = 0.25
'    Ins(19) = 0.25
'    Ins(20) = 3
'    Dim K As Double
'
''Type of Conductor
'    If UCase(txtNo) = "S" Then
'        Dia = txtODs * 1
'    ElseIf UCase(txtNo) = "D" Then
'        Dia = txtODs * 2
'    ElseIf UCase(txtNo) = "T" Then
'        Dia = txtODs * 3
'    ElseIf UCase(txtNo) = "Q" Then
'        Dia = txtODs * 4
'    End If
''Counting No of turns
'
'    Dim typ As Double
'
'    For i = 0 To CInt(Tur)
'        K = 0
''For First limb
'        For l = 0 To Count
'            K = K + Ins(l)
'        Next
'        If Count > 0 Then
'            K = K + (txtODs * Count)
'        End If
'
'        'Rounding of k
'        If InStr(1, K, ".") > 0 Then
'            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
'                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
'            Else
'                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
'            End If
'        End If
'        i = ((CDbl(txtID1) - 2 * (K)) * 0.9 * 3.142) / Dia
'        'Rounding of i
'        If InStr(1, i, ".") > 0 Then
'            i = Left(i, InStr(1, i, ".") - 1)
'        End If
'
'        value = value + i
'        turns1 = Turnss
'        Turnss = Turnss - value
'
'        If Turnss <= 0 Then
'            MsgBox ((turns1 / value) * 100)
'            finalt = ((turns1 / value) * 100)
'            If finalt <= 25 Then
'                typ = 0.25
'                Count = (Count + typ)
'                Exit For
'            ElseIf finalt > 25 And finalt <= 50 Then
'                typ = 0.5
'                Count = (Count + typ)
'                Exit For
'            ElseIf finalt > 50 And finalt <= 75 Then
'                typ = 0.75
'                Count = (Count + typ)
'                Exit For
'            ElseIf finalt > 75 And finalt <= 100 Then
'                typ = 1
'                Count = (Count + typ)
'                Exit For
'            End If
'        End If
'        value = 0
'        typ = 0
'        Count = Count + 1
''        If value >= CInt(txtSecTurn) Then
''             CalLayers = (Count + typ)
''             Exit Function
''        End If
'
'
'    Next
'
'    If InStr(1, Count, ".") > 0 Then
''    MsgBox Left(Count, InStr(1, Count, ".") - 1)
'        Count = Left(Count, InStr(1, Count, ".") - 1) + 1
'    End If
''Total Winding HT
'    WHt = 1.5 + (txtODs * Count) + ((Count - 1) * 0.25) + (1.5 + 3)
'    WHt = Round(WHt, 2)
'
''Mean Length of Turns
'    MLT = (WHt * 3.142) + ((2 * txtBUP) + (2 * txtHt))
'    MLT = Round(MLT, 2)
'
''Total Length of wire
'    Toleng = CDbl(MLT) * CDbl(Tur)
'    prcCalLowerCir = Toleng
'
'
'End Function




Private Sub txtminCs_GotFocus()
txtminCs.SelStart = 0
txtminCs.SelLength = Len(txtminCs)
MDIDesign.StatusBar1.Panels(2).Text = "Min. Required Cross Section (Sqmm)"
End Sub

Private Sub txtNo_Change()

Dim nos As Byte
'If UCase(txtNo) = "S" Then
'    nos = 1
'ElseIf UCase(txtNo) = "D" Then
'    nos = 2
'ElseIf UCase(txtNo) = "T" Then
'    nos = 3
'ElseIf UCase(txtNo) = "Q" Then
'    nos = 4
'End If
'
'txtCompen(0) = compensation / nos
'txtCompen(0) = Round(txtCompen(0), 2)

If InCalcu = "" Then
For i = 0 To 4
    txtCompen(i) = 0
Next
End If

End Sub

Private Sub txtNo_GotFocus()
    txtNo.SelLength = Len(txtNo)
    MDIDesign.StatusBar1.Panels(2).Text = "Conductor Nos"
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[SsDdTtQq]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Function RemoveItem(INo As Integer)

Dim C1 As Integer

'Core Specifiation
C1 = Grid.Rows
Do Until C1 = txtnoofcores
    If Grid.TextMatrix(C1 - 1, 0) > INo Then
        Grid.RemoveItem (C1 - 1)
    End If
    C1 = C1 - 1
Loop

'Core Dimenstion
C1 = GCore.Rows
Do Until C1 = txtnoofcores
    If GCore.TextMatrix(C1 - 1, 0) > INo Then
        GCore.RemoveItem (C1 - 1)
    End If
    C1 = C1 - 1
Loop

'Checking wheather details of this core alrady exits
For i = 1 To CTfrm.G1.Cols - 1
    If G1.TextMatrix(1, i) > txtnoofcores Then
        Dim copeTOgrid As Copying
        Set copeTOgrid = New Copying
        copeTOgrid.CopyGrid (txtnoofcores)
        Set copeTOgrid = Nothing
        Exit For
    End If
Next

wounded.Rows = 1
Former.Rows = 1
Casting.Rows = 1
Mould.Rows = 1

End Function

Private Sub txtNo_LostFocus()
If Len(txtNo) > 0 Then
    txtNo = UCase(txtNo)
End If
End Sub

Private Sub txtnoofcores_Change()
cmbnoofcores.Clear
cmbCore.Clear

If Grid.Rows > 1 Then
    joel = MsgBox("Changing no of core will create change in core description" _
    & "Are you sure you want to continue?", vbInformation + vbYesNo)
    If vbYes = joel Then
        G1.Cols = 1
        Grid.Rows = 1
        GCore.Rows = 1
'        RemoveItem (txtnoofcores)
    End If
ElseIf GCore.Rows > 1 Then
    joel = MsgBox("Changing no of core will create change in core description" _
    & "Are you sure you want to continue?", vbInformation + vbYesNo)
    If vbYes = joel Then
        RemoveItem (txtnoofcores)
    End If
ElseIf G1.Cols > 1 Then
    joel = MsgBox("Changing no of core will create change in core description" _
    & "Are you sure you want to continue?", vbInformation + vbYesNo)
    If vbYes = joel Then
        RemoveItem (txtnoofcores)
    End If
ElseIf wounded.Rows > 1 Then
    joel = MsgBox("Changing no of core will create change in core description" _
    & "Are you sure you want to continue?", vbInformation + vbYesNo)
    If vbYes = joel Then
        RemoveItem (txtnoofcores)
    End If
End If

For i = 0 To 4
    Label14.Visible = False
    Label17.Visible = False
    txtRatioDum(i).Visible = False
    txtRatioR2Dum(i).Visible = False
    Label6(i).Visible = False
Next

''+++++++++++++++++ newly added
    cmbCType.ListIndex = -1
    cmbCType.Text = "-Select Core Type-"
    ''**** if all the core details are present then blank
    ''**** ratio on specification
    For l = 1 To Grid.Rows - 1
        vinn = 0
        If l = cmbnoofcores Then
            vinn = 1
        End If
    Next
    If vinn = 1 Then
        txtRatio2.Text = ""
    End If
''+++++++++++++++++  end newly added

If Len(txtnoofcores) <> 0 Then
    If Val(txtnoofcores) <= 5 Then
        Label14.Visible = True
        Label17.Visible = True
        For i = 0 To Val(txtnoofcores) - 1
            txtRatioDum(i).Visible = True
            txtRatioR2Dum(i).Visible = True
            Label6(i).Visible = True
            If Val(txtnoofcores) = 1 Then
                Label6(i).Caption = "Ratio :"
            Else
                Label6(i).Caption = "Core " & i + 1 & " Ratio :"
            End If
            cmbnoofcores.AddItem i + 1
            cmbCore.AddItem "Core " & i + 1
            cmbCore.ListIndex = 0
        Next
    Else
        MsgBox "No of Cores cannot be greater than 5", vbCritical
        txtnoofcores = ""
        txtnoofcores.SetFocus
        Exit Sub
    End If
End If

End Sub

Private Sub txtnoofcores_GotFocus()
txtnoofcores.SelStart = 0
txtnoofcores.SelLength = Len(txtnoofcores)
MDIDesign.StatusBar1.Panels(2).Text = "Set no of core(s)"
End Sub

Private Sub txtnoofcores_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[1-5]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
txtnoofcores.Locked = False
End Sub

Private Sub txtOD_Change()
    
'If Val(txtOD) <= Val(txtID) Then
'
'    MsgBox "Not a valid core size!", vbInformation
'    txtID = 0
'    txtID.SetFocus
'    Exit Sub
'
'End If

If Len(txtOD) = 0 Or Len(txtID) = 0 Then

    Exit Sub

End If

Dim Mean, BU As Double

'finding build up

BU = prcCalBU(txtOD, txtID)

txtBUP = BU

txtBUP.Text = BU

If Len(txtOD) = 0 Or Len(txtID) = 0 Then

    Exit Sub

End If

'finding lm mean

Mean = prcCalMeanCircular(txtOD, txtID)

txtMean.Text = Round(Mean, 2)

End Sub

Private Sub txtOD_GotFocus()
txtOD.SelStart = 0
txtOD.SelLength = Len(txtOD)
MDIDesign.StatusBar1.Panels(2).Text = "Set OD"
End Sub

Private Sub txtOD_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtOD_LostFocus()

If Len(txtOD) = 0 Or Len(txtID) = 0 Then

    Exit Sub

End If

Dim Mean, BU As Double

'finding build up

BU = prcCalBU(txtOD, txtID)

txtBUP = BU

txtBUP.Text = BU

If Len(txtOD) = 0 Or Len(txtID) = 0 Then

    Exit Sub

End If

'finding lm mean

Mean = prcCalMeanCircular(txtOD, txtID)

txtMean.Text = Mean

End Sub

Private Sub txtOL_Change()

'If Val(txtOL) <= Val(txtWL) Then
'
'    MsgBox "Not a valid core size!", vbInformation
'    txtWL = 0
'    txtWL.SetFocus
'    Exit Sub
'
'End If


If Len(txtOL) = 0 Or Len(txtWL) = 0 Then

    Exit Sub

End If

Dim Mean, BU As Double

BU = prcCalBU(txtOL, txtWL)

txtBUP = BU

If Len(txtOL) = 0 Or Len(txtWL) = 0 Or Len(txtOW) = 0 Or Len(txtWW) = 0 Then
    Exit Sub
End If
Mean = prcCalMean(txtWL, txtWW, BU)
txtMean.Text = Mean

End Sub


Private Sub txtOL_GotFocus()
txtOL.SelStart = 0
txtOL.SelLength = Len(txtOL)
MDIDesign.StatusBar1.Panels(2).Text = "Set OL"
End Sub

Private Sub txtOL_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtOW_Change()

'If Val(txtOW) <= Val(txtWW) Then
'
'    MsgBox "Not a valid core size!", vbInformation
'    txtWW = 0
'    txtWW.SetFocus
'    Exit Sub
'
'End If

If Len(txtOL) = 0 Or Len(txtWL) = 0 Then

    Exit Sub

End If

Dim Mean, BU As Double

BU = prcCalBU(txtOL, txtWL)

txtBUP = BU

If Len(txtOL) = 0 Or Len(txtWL) = 0 Or Len(txtOW) = 0 Or Len(txtWW) = 0 Then

    Exit Sub

End If

Mean = prcCalMean(txtWL, txtWW, BU)

txtMean.Text = Mean

End Sub

'Private Sub txtRatio_Change()

'txtRatio2.Text = txtRatio.Text
'
'If InStr(1, Trim(txtRatio), "-") > 0 Then
'
'    Frame4.Enabled = True
'
'Else
'
'    Frame4.Enabled = False
'
'End If
'
'Dim king As String * 1
'Dim str1, str2 As String
'Dim str3 As Integer
'str3 = 0
'For l = 1 To Len(txtRatio)
''    str1 = ""
'    king = Mid(txtRatio, l, l)
'    If king = "-" Then
'
'        str3 = str3 + 1
'    Else
'
'        str1 = str1 + king
'
'    End If
'
'Next
'
'If InStr(1, Trim(txtRatio), "-") > 0 Then
'
'    If Mid(txtRatio, Len(txtRatio), 1) = "-" Then
'
'        If str3 <> 0 Then
'            str3 = str3 - 1
'        Else
'            str3 = 0
'        End If
'
'    End If
'
'Else
'
'    str3 = 0
'
'End If
'
'NoofPrim = str3

'End Sub

Public Sub prcFillCom()

'+++++++++++IL

Set rst = New ADODB.Recordset

rst.Open "select il from il", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    rst.Requery
    cmbIL.Clear
    Do Until rst.EOF = True
        
        cmbIL.AddItem rst!iL
        rst.MoveNext
        
    Loop

End If

rst.Close

'+++++++++++HSV
'
'Set rst = New ADODB.Recordset
'
'rst.Open "select hsv from hsv", con, adOpenStatic, adLockReadOnly
'
'If rst.EOF = False And rst.BOF = False Then
'
'    rst.MoveFirst
'    rst.Requery
'    cmbHSV.Clear
'    Do Until rst.EOF = True
'
'        cmbHSV.AddItem rst!hsv & " kV"
'        rst.MoveNext
'
'    Loop
'
'End If
'
'rst.Close

'++++++++++++STC

Set rst = New ADODB.Recordset

rst.Open "select stc from stc", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then
    rst.MoveFirst
    rst.Requery
    cmbSTC.Clear
    Do Until rst.EOF = True
        cmbSTC.AddItem rst!stc
        rst.MoveNext
    Loop
End If

rst.Close

'++++++++++++++Time
'
'Set rst = New ADODB.Recordset
'
'rst.Open "select distinct cus_sname from cus", con, adOpenStatic, adLockReadOnly
'
'If rst.EOF = False And rst.BOF = False Then
'
'    rst.Requery
'    rst.MoveFirst
'    cmbCusName.Clear
'    Do Until rst.EOF = True
'        If Len(rst!cus_sname) >= 1 Then
'        cmbCusName.AddItem rst!cus_sname
'        End If
'        rst.MoveNext
'
'    Loop
'
'End If
'
'rst.Close

'++++++++++++++Drg

'Set rst = New ADODB.Recordset
'
'rst.Open "select Drg from Drg", con, adOpenStatic, adLockReadOnly
'
'If rst.EOF = False And rst.BOF = False Then
'
'    rst.MoveFirst
'    rst.Requery
'    cmbDrg.Clear
'    Do Until rst.EOF = True
'
'        cmbDrg.AddItem rst!drg
'        rst.MoveNext
'
'    Loop
'
'End If
'
'rst.Close

'+++++++++++++Freq

'Set rst = New ADODB.Recordset
'
'rst.Open "select Freq from Freq", con, adOpenStatic, adLockReadOnly
'
'If rst.EOF = False And rst.BOF = False Then
'
'    rst.MoveFirst
'    rst.Requery
'    cmbFreq.Clear
'    Do Until rst.EOF = True
'
'        cmbFreq.AddItem rst!Freq
'        rst.MoveNext
'
'    Loop
'
'End If
'
'rst.Close

'+++++++++++++NSV

Set rst = New ADODB.Recordset

rst.Open "select nsv from hsv", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    rst.Requery
    cmbNSV.Clear
    Do Until rst.EOF = True

        cmbNSV.AddItem rst!NSV
        rst.MoveNext

    Loop

End If

rst.Close

'+++++++++++++++Equipment

Set rst = New ADODB.Recordset

rst.Open "select distinct [item] from Dtype where [mode] = '2'", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    rst.Requery
    cmbEquipType.Clear
    Do Until rst.EOF = True
        'for ltct
'        If "Window HTCTs ( Win-R )" = rst!Item Or "Window HTCTs ( Win-C )" = rst!Item Then
            cmbEquipType.AddItem rst![Item]
'        End If
        rst.MoveNext

    Loop

End If

rst.Close
'
End Sub

Public Sub AddRowGrid2()

Grid2.Rows = 1

For i = 1 To CInt(txtnoofcores)

    Grid2.Rows = Grid2.Rows + 1
    Grid2.row = Grid2.row - 1
    
    Grid2.RowHeight(i) = 700
    Grid2.TextMatrix(i, 0) = i

Next

End Sub

Public Sub prcCoreSize()

Grid3.Rows = 1
Dim NSV As Double
If Trim(UCase(cmbNSV)) = "KV" Or Len(cmbNSV) = 0 Then
    Exit Sub
End If
If InStr(1, cmbNSV, "k") > 0 Then
    If Mid(cmbNSV, 1, InStr(1, UCase(cmbNSV), "K") - 1) <= 11 Then
        NSV = 11
    Else
        NSV = Mid(cmbNSV, 1, InStr(1, UCase(cmbNSV), "K") - 1)
    End If
Else
    If cmbNSV <= 11 Then
        NSV = 11
    Else
        NSV = cmbNSV
    End If
End If
Set rst = New Recordset
rst.Open "select * from coresize where nsv = " & NSV & "", con, adOpenDynamic, adLockOptimistic

If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    Do Until rst.EOF = True

        Grid3.Rows = Grid3.Rows + 1
        Grid3.row = Grid3.Rows - 1
        Grid3.RowHeight(Grid3.row) = 300
        
        Grid3.TextMatrix(Grid3.row, 0) = Grid3.row
        Grid3.TextMatrix(Grid3.row, 1) = rst.Fields(0)
        Grid3.TextMatrix(Grid3.row, 2) = rst.Fields(1)
        Grid3.TextMatrix(Grid3.row, 3) = rst.Fields(2)
        Grid3.TextMatrix(Grid3.row, 4) = rst.Fields(3)
        Grid3.TextMatrix(Grid3.row, 5) = rst.Fields(2)
        Grid3.TextMatrix(Grid3.row, 6) = rst.Fields(3)
        Grid3.TextMatrix(Grid3.row, 7) = rst.Fields(0)
        Grid3.TextMatrix(Grid3.row, 8) = rst.Fields(1)
        rst.MoveNext

    Loop

rst.Close
Set rst = Nothing

End If

End Sub

'Private Sub txtRatio_LostFocus()
'
'Dim dashcount As String * 1
'
'If Len(txtRatio) <> 0 Then
'
'    withoutslash = Trim(txtRatio)
'    slashlen = Len(withoutslash)
'
'    For i = 1 To slashlen
'
'        dashcount = Mid(withoutslash, i, i + 1)
'        If dashcount = "-" Then
'
'            noofcore = noofcore + 1
'
'        End If
'
'    Next
'
'    NoofSec = noofcore + 1
'
'Else
'
'    NoofSec = 0
'
'End If
'
'End Sub

'Private Sub txtRatioR2_Change()

'''txtRatio2R2.Text = txtRatioR2.Text
'''txtNominal.Text = txtRatioR2.Text
'''If Grid.Row = 0 Then
'''    Frame2 = "Core " & Grid.Row + 1 & " Ratio : "
'''    If InStr(1, txtRatioR2, "-") > 0 Then
''''        MsgBox Left(txtRatioR2.Text, InStr(1, txtRatioR2, "-") - 1)
'''        txtRatio2R2.Text = Left(txtRatioR2.Text, InStr(1, txtRatioR2, "-") - 1)
'''    Else
'''        txtRatio2R2.Text = txtRatioR2
'''    End If
'''End If

'If Len(txtRatio2) <> 0 And Len(cmbSTC) <> 0 Then
'
'    Set rst = New Recordset
'    rst.Open "select ncd from normalcd", con, adOpenStatic, adLockReadOnly
'
'    If rst.EOF = False And rst.BOF = False Then
'
'        Dim j As String * 1
'        Dim value As String
'
'        If Len(txtTurns) <> 0 Then
'
'            txtAmper = prcPriAmper(txtRatio, txtRatioR2, txtTurns, rst!NCD)
'
'        End If
'
'    End If
'
''    rst.Close
'    Set rst = Nothing
'
'End If
'''Dim dashcount As String * 1
'''
'''If Len(txtRatioR2) <> 0 Then
'''
'''    withoutslash = Trim(txtRatioR2)
'''    slashlen = Len(withoutslash)
'''
'''    For i = 1 To slashlen
'''
'''        dashcount = Mid(withoutslash, i, i + 1)
'''        If dashcount = "-" Then
'''
'''            noofcore = noofcore + 1
'''
'''        End If
'''
'''    Next
'''
'''    txtnoofcores = noofcore + 1
''''    MsgBox Right(txtRatioR2, 1)
'''
'''Else
'''
'''    txtnoofcores = 0
'''
'''End If
'''
'''If Right(txtRatioR2, 1) = "-" And Val(txtnoofcores) <> 0 Then
'''        txtnoofcores = Val(txtnoofcores) - 1
'''End If

'End Sub

'Private Sub txtRatioR2_LostFocus()
'
'Dim dashcount As String * 1
'
'If Len(txtRatioR2) <> 0 Then
'
'    withoutslash = Trim(txtRatioR2)
'    slashlen = Len(withoutslash)
'
'    For i = 1 To slashlen
'
'        dashcount = Mid(withoutslash, i, i + 1)
'        If dashcount = "-" Then
'
'            noofcore = noofcore + 1
'
'        End If
'
'    Next
'
''    txtnoofcores = noofcore + 1
'
'Else
'
''    txtnoofcores = 0
'
'End If
'
'If Right(txtRatioR2, 1) = "-" And Val(txtnoofcores) <> 0 Then
'        txtnoofcores = Val(txtnoofcores) - 1
'End If
'
'Dim str3 As Integer
'str3 = 1
'For l = 1 To Len(txtRatioR2) + 1
'
'    king = Mid(txtRatioR2, l, l)
'    If king = "-" Or Len(txtRatioR2) < l Then
'
'        str2 = str1
'
'        If str3 = Grid.Row Then
'
'            txtRatio2R2 = str3
'            Exit Sub
'
'        End If
'        str1 = ""
'
'    Else
'
'        str1 = str1 + king
'
'    End If
'
'Next
'
'End Sub

Private Sub txtSecConductor_Change()

Set rst = New Recordset
rst.Open "select t_cs, t_ods from secconductor where t_swg = '" & txtSecConductor & "'", con, adOpenDynamic, adLockOptimistic

If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    txtt_cs.Text = rst!t_cs
    txtt_ods.Text = rst!t_ods

End If

End Sub

Private Sub txtSecConductor_Click()

Set rst = New Recordset
rst.Open "select t_cs, t_ods from secconductor where t_swg = '" & txtSecConductor & "'", con, adOpenDynamic, adLockOptimistic

If rst.EOF = False And rst.BOF = False Then

    rst.MoveFirst
    txtt_cs.Text = rst!t_cs
    txtt_ods.Text = rst!t_ods

End If

End Sub


Private Sub PrcCalculate()
'++++++++++++++
lstCond.Clear
cmbConStyle.Enabled = True
'If cmbConStyle.ListIndex = 0 Then
'    CoilType = "P"
'    CoilName = "Double Coil - Normal Wdg Ht."
'ElseIf cmbConStyle.ListIndex = 1 Then
'    CoilType = "P"
'    CoilName = ""
'ElseIf cmbConStyle.ListIndex = 2 Then
'    CoilType = "B"
'    CoilName = "Double Coil - Normal Wdg Ht."
'ElseIf cmbConStyle.ListIndex = 3 Then
'    CoilType = "B"
'    CoilName = "Series Coil"
'ElseIf cmbConStyle.ListIndex = 4 Then
'    CoilType = "P"
'    CoilName = ""
'ElseIf cmbConStyle.ListIndex = 5 Then
'    CoilType = "P"
'    CoilName = ""
'End If
If Len(txtTurns) = 0 Then
    Exit Sub
End If

'If cmbConStyle.ListIndex = 2 Or cmbConStyle.ListIndex = 3 Then
'    If CInt(txtTurns) > 1 Then
'        MsgBox "For bare conductor turns will always be one", vbCritical
'        txtTurns = ""
'        Exit Sub
'    End If
'ElseIf cmbConStyle.ListIndex = 0 Or cmbConStyle.ListIndex = 1 Or cmbConStyle.ListIndex = 4 Or cmbConStyle.ListIndex = 5 Then
'    If CInt(txtTurns) = 1 Then
'        MsgBox "For Paper conductor turns will always be grater than one", vbCritical
'        txtTurns = ""
'        Exit Sub
'    End If
'End If

Dim STCCurr, ConCurr As Double
Set rst = New Recordset
rst.Open "select ncd from normalcd", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then

    Dim j As String * 1
    Dim value As Double
    Dim str1 As String
    If Len(txtTurns) <> 0 Then
        'Calculate minimum primary cross section area at normal or continus current
        
        For i = 0 To Val(txtnoofcores) - 1
'        MsgBox txtRatio(i)
            If InStr(1, txtRatio(i), "-") = 0 Then
                If value < txtRatio(i) Then
                    value = txtRatio(i)
                End If
            Else
                For K = 1 To Len(txtRatio(i)) + 1
                    j = Mid(txtRatio(i), K, K)
                    If j = "-" Or K > Len(txtRatio(i)) Then
                        If value < CDbl(str1) Then
                            value = CDbl(str1)
                        End If
                        str1 = ""
                    Else
                        str1 = str1 & j
                    End If
                Next
            End If
        Next
        
        If Len(Trim(txtRateConti)) > 0 Then
                value = CDbl(txtRateConti)
        End If
'        txtAmper = prcPriAmper(value, txtRatioR2, txtTurns, rst!NCD)
'        txtAmper = prcPriAmper(Mid(value, 2), 10, txtTurns, rst!NCD)
        
        STCCurr = prcCalNorPC(value)
        If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
            ConCurr = prcCalPC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), 0, CDbl(cmbTime))
        End If
        
        If STCCurr >= ConCurr Then
            txtPriCros = STCCurr
        ElseIf STCCurr < ConCurr Then
            txtPriCros = ConCurr
        End If
        If Len(txtPriCros) = 0 Then
            txtPriCros = 0
        End If
        txtminCs = txtPriCros
        'Search for same or grater cross section
        Set rst = New ADODB.Recordset
        rst.Open "select t_csp, t_whpt, t_whps, t_cond_siz, t_cond_coi from primarycond where T_NP = " & txtTurns & " AND t_csp >= " & txtPriCros & " and type = '" & coiltype & "' and style = '" & Trim(cmbConStyle) & "' order by t_csp", con, adOpenStatic
        If rst.EOF = False And rst.BOF = False Then
            txtPriCros = rst.Fields(0)
            txtTop = rst!t_whpt
            txtSide = rst!t_whps
            txtCondPer = rst!t_COND_coi
            txtTotalCon = rst!t_cond_siz
        Else
            txtPriCros = ""
            txtTop = ""
            txtSide = ""
            txtCondPer = ""
            txtTotalCon = ""
            'When select equipment for the first time
            If Grid.Rows > 1 Then
                MsgBox "No Details available for " & cmbConStyle, vbInformation
            End If
            Exit Sub
        End If
        Set rst = Nothing
        
        txtDelR = prcCalDElRate(CDbl(txtPriCros), value)
        If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
            txtDelStc = prcCalDElSTC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), CDbl(cmbTime), txtPriCros)
        End If
        lstCond.Clear
        
        'Setting 2 less than and 2 less than C/s1
        Set rst = New ADODB.Recordset
        rst.Open "select distinct t_csp from primaryCond where t_csp < " & Trim(txtPriCros) & " and type = '" & coiltype & "'  and style = '" & CoilName & "' and t_np = " & txtTurns & " order by t_csp desc", con, adOpenStatic, adLockReadOnly
        
        If rst.EOF = False And rst.BOF = False Then
            
            rst.MoveFirst
            
            If rst.RecordCount < 2 Then
                lstCond.AddItem rst!t_csp
                lstCond.ListIndex = 0
            Else
                For i = 0 To 1
                      lstCond.AddItem rst!t_csp
                      rst.MoveNext
                Next
                lstCond.ListIndex = 0
            End If
        
        End If
        
        lstCond.AddItem txtPriCros
        lindex = lstCond.ListIndex
        'Setting 2 greater than and 2 less than C/s
        Set rst = New ADODB.Recordset
        rst.Open "select distinct t_csp from primaryCond where t_csp > " & Trim(txtPriCros) & " and type = '" & coiltype & "'  and style = '" & CoilName & "' and t_np = " & txtTurns & " order by t_csp", con, adOpenStatic, adLockReadOnly
        
        If rst.EOF = False And rst.BOF = False Then
            
            rst.MoveFirst
            
            If rst.RecordCount < 2 Then
                lstCond.AddItem rst!t_csp
                lstCond.ListIndex = 0
            Else
                For i = 0 To 1
                      lstCond.AddItem rst!t_csp
                      rst.MoveNext
                Next
                lstCond.ListIndex = 0
            End If
        
        End If
        
        For i = 0 To lstCond.ListCount
            If lstCond.List(i) = Trim(txtPriCros) Then
                lstCond.ListIndex = i
                Exit For
            End If
        Next
        
    End If

End If
        
'        Set rst = New ADODB.Recordset
'        rst.Open "select * from primarycond where t_NP = " & txtTurns & " and t_csp = '" & lstCond & "' ", con, adOpenStatic, adLockReadOnly
'
'        If rst.EOF = False And rst.BOF = False Then
'
'            txtTop = rst!t_whpt
'            txtSide = rst!t_whps
'
'        Else
'
'            txtTop = ""
'            txtSide = ""
'
'        End If
'
'        Set rst = Nothing

End Sub

Private Function FindIndexConduct() As Integer
    Dim FlagBool As Boolean
    Dim i As Integer
    For i = 0 To lstCond.ListCount - 1
        If lstCond.List(i) > CDbl(txtminCs) Then
'            MsgBox lstCond.List(i)
            FindIndexConduct = i
            FlagBool = True
            Exit Function
        End If
    Next
    
    FindIndexConduct = 1000
    
End Function

Private Sub PrcCalPrimaryDetails()
On Error GoTo err
Dim rstCon As Recordset
Dim FlagCounter As String

'++++++++++++++

''setting coil type as per conductor
''Note
''width will be set as per conductor (BAR TYPE IS TEMParary)
lstCond.Clear
cmbConStyle.Enabled = True
If cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Bare" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double" Then
    
    txtTurns.Clear
    
    ''Filling Combo
    Set rstCon = New Recordset
    rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' order by t_np", con, adOpenStatic, adLockReadOnly
    If rstCon.EOF = False And rstCon.BOF = False Then
        rstCon.MoveFirst
        Do Until rstCon.EOF
            txtTurns.AddItem rstCon.Fields(0) & vbNullString
            rstCon.MoveNext
        Loop
        
        If rstCon.RecordCount = 1 Then
            txtTurns.ListIndex = 0
        Else
            txtTurns.ListIndex = -1
        End If
        
        ''Details of primary
        Set rstCon = New Recordset
        rstCon.Open "select pri_wind_width, t_np, t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_np = " & Trim(txtTurns) & "", con, adOpenStatic, adLockReadOnly
        If rstCon.EOF = False And rstCon.BOF = False Then
'            txtturns = rstCon!t_np & vbNullString
            txtWidth = rstCon!pri_wind_width & vbNullString
            txtPriCros = rstCon!t_csp & vbNullString
            txtTop = rstCon!t_whpt & vbNullString
            txtSide = rstCon!t_whps & vbNullString
'            txtCondPer = rstCon!t_COND_coi & vbNullString
'            txtTotalCon = rstCon!t_cond_siz & vbNullString
        End If
        Set rstCon = Nothing
    End If
    Set rstCon = Nothing
    
ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Bare" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Single" Then
    
    txtTurns.Clear
    
    ''Filling Combo
    Set rstCon = New Recordset
    rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'  order by t_np", con, adOpenStatic, adLockReadOnly
    If rstCon.EOF = False And rstCon.BOF = False Then
        rstCon.MoveFirst
        Do Until rstCon.EOF
            txtTurns.AddItem rstCon.Fields(0) & vbNullString
            rstCon.MoveNext
        Loop
        
'        If rstCon.RecordCount = 1 Then
            txtTurns.ListIndex = 0
'        Else
'            txtturns.ListIndex = -1
'        End If
        
        ''Details of primary
        Set rstCon = New Recordset
        rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'", con, adOpenStatic, adLockReadOnly
        If rstCon.EOF = False And rstCon.BOF = False Then
'            txtturns = rstCon!t_np & vbNullString
            txtWidth = rstCon!pri_wind_width & vbNullString
            txtPriCros = rstCon!t_csp & vbNullString
            txtTop = rstCon!t_whpt & vbNullString
            txtSide = rstCon!t_whps & vbNullString
'            txtCondPer = rstCon!t_COND_coi & vbNullString
'            txtTotalCon = rstCon!t_cond_siz & vbNullString
        End If
        Set rstCon = Nothing
    End If
    Set rstCon = Nothing
    
ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double-Different" Then
        
    txtTurns.Clear
    
    ''Filling Combo
    Set rstCon = New Recordset
    rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' order by t_np", con, adOpenStatic, adLockReadOnly
    If rstCon.EOF = False And rstCon.BOF = False Then
        rstCon.MoveFirst
        Do Until rstCon.EOF
            txtTurns.AddItem rstCon.Fields(0) & vbNullString
            rstCon.MoveNext
        Loop
        
'        If rstCon.RecordCount = 1 Then
            txtTurns.ListIndex = 0
'        Else
'            txtturns.ListIndex = -1
'        End If
'
        ''Details of primary
        Set rstCon = New Recordset
        rstCon.Open "select pri_wind_width, t_np, t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'", con, adOpenStatic, adLockReadOnly
        If rstCon.EOF = False And rstCon.BOF = False Then
'            txtTurns = rstCon!t_np & vbNullString
            txtWidth = rstCon!pri_wind_width & vbNullString
            txtPriCros = rstCon!t_csp & vbNullString
            txtTop = rstCon!t_whpt & vbNullString
            txtSide = rstCon!t_whps & vbNullString
'            txtCondPer = rstCon!t_COND_coi & vbNullString
'            txtTotalCon = rstCon!t_cond_siz & vbNullString
        End If
        Set rstCon = Nothing
    End If
    Set rstCon = Nothing

ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double-Equal" Then
    
    txtTurns.Clear
    
    ''Filling Combo
    Set rstCon = New Recordset
    rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'  order by t_np", con, adOpenStatic, adLockReadOnly
    If rstCon.EOF = False And rstCon.BOF = False Then
        rstCon.MoveFirst
        Do Until rstCon.EOF
            txtTurns.AddItem rstCon.Fields(0) & vbNullString
            rstCon.MoveNext
        Loop
        
'        If rstCon.RecordCount = 1 Then
            txtTurns.ListIndex = 0
'        Else
'            txtturns.ListIndex = -1
'        End If
        
        ''Details of primary
        Set rstCon = New Recordset
        rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_np = " & Trim(txtTurns) & "", con, adOpenStatic, adLockReadOnly
        If rstCon.EOF = False And rstCon.BOF = False Then
'            txtTurns = rstCon!t_np & vbNullString
            txtWidth = rstCon!pri_wind_width & vbNullString
            txtPriCros = rstCon!t_csp & vbNullString
            txtTop = rstCon!t_whpt & vbNullString
            txtSide = rstCon!t_whps & vbNullString
'            txtCondPer = rstCon!t_COND_coi & vbNullString
'            txtTotalCon = rstCon!t_cond_siz & vbNullString
        End If
        Set rstCon = Nothing
    End If
    Set rstCon = Nothing
    
ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Series" Then
    
    txtTurns.Clear
    
    ''Filling Combo
    Set rstCon = New Recordset
    rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' order by t_np", con, adOpenStatic, adLockReadOnly
    If rstCon.EOF = False And rstCon.BOF = False Then
        rstCon.MoveFirst
        Do Until rstCon.EOF
            txtTurns.AddItem rstCon.Fields(0) & vbNullString
            rstCon.MoveNext
        Loop
        
'        If rstCon.RecordCount = 1 Then
            txtTurns.ListIndex = 0
'        Else
'            txtturns.ListIndex = -1
'        End If
        
        ''Details of primary
        Set rstCon = New Recordset
        rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'", con, adOpenStatic, adLockReadOnly
        If rstCon.EOF = False And rstCon.BOF = False Then
'            txtTurns = rstCon!t_np & vbNullString
            txtWidth = rstCon!pri_wind_width & vbNullString
            txtPriCros = rstCon!t_csp & vbNullString
            txtTop = rstCon!t_whpt & vbNullString
            txtSide = rstCon!t_whps & vbNullString
'            txtCondPer = rstCon!t_COND_coi & vbNullString
'            txtTotalCon = rstCon!t_cond_siz & vbNullString
        End If
        Set rstCon = Nothing
    End If
    Set rstCon = Nothing
    
ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Single" Then
    
    txtTurns.Clear
    
    ''Filling Combo
    Set rstCon = New Recordset
    rstCon.Open "select distinct t_np from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'  order by t_np", con, adOpenStatic, adLockReadOnly
    If rstCon.EOF = False And rstCon.BOF = False Then
        rstCon.MoveFirst
        Do Until rstCon.EOF
            txtTurns.AddItem rstCon.Fields(0) & vbNullString
            rstCon.MoveNext
        Loop
        
'        If rstCon.RecordCount = 1 Then
            txtTurns.ListIndex = 0
'        Else
'            txtturns.ListIndex = -1
'        End If
        
        ''Details of primary
        Set rstCon = New Recordset
        rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'", con, adOpenStatic, adLockReadOnly
        If rstCon.EOF = False And rstCon.BOF = False Then
'            txtTurns = rstCon!t_np & vbNullString
            txtWidth = rstCon!pri_wind_width & vbNullString
            txtPriCros = rstCon!t_csp & vbNullString
            txtTop = rstCon!t_whpt & vbNullString
            txtSide = rstCon!t_whps & vbNullString
'            txtCondPer = rstCon!t_COND_coi & vbNullString
'            txtTotalCon = rstCon!t_cond_siz & vbNullString
        End If
        Set rstCon = Nothing
    End If
    Set rstCon = Nothing
    
End If

Dim ConCurr As Double
Set rst = New Recordset
rst.Open "select ncd from normalcd", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then

    Dim j As String * 1
    Dim value As Double
    Dim str1 As String
    If Len(txtTurns) <> 0 Then
        'Calculate minimum primary cross section area at normal or continus current
        
        For i = 0 To Val(txtnoofcores) - 1
'        MsgBox txtRatio(i)
            If InStr(1, txtRatio(i), "-") = 0 Then
                If value < txtRatio(i) Then
                    value = txtRatio(i)
                End If
            Else
                For K = 1 To Len(txtRatio(i)) + 1
                    j = Mid(txtRatio(i), K, K)
                    If j = "-" Or K > Len(txtRatio(i)) Then
                        If value < CDbl(str1) Then
                            value = CDbl(str1)
                        End If
                        str1 = ""
                    Else
                        str1 = str1 & j
                    End If
                Next
            End If
        Next
        
        If Len(Trim(txtRateConti)) > 0 Then
            value = CDbl(txtRateConti)
        End If
'        txtAmper = prcPriAmper(value, txtRatioR2, txtTurns, rst!NCD)
'        txtAmper = prcPriAmper(Mid(value, 2), 10, txtTurns, rst!NCD)
        
        STCCurr = prcCalNorPC(value)
        If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
            ConCurr = prcCalPC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), 0, CDbl(cmbTime))
        End If
        
        If STCCurr >= ConCurr Then
            txtPriCros = STCCurr
        ElseIf STCCurr < ConCurr Then
            txtPriCros = ConCurr
        End If
        If Len(txtPriCros) = 0 Then
            txtPriCros = 0
        End If
        txtminCs = txtPriCros
        
        
        txtDelR = prcCalDElRate(CDbl(txtPriCros), value)
        If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
            txtDelStc = prcCalDElSTC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), CDbl(cmbTime), txtPriCros)
        End If
        lstCond.Clear
       ''Filling Conductor
        Set rst = New Recordset
        rst.Open "select distinct t_csp from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' order by t_csp", con, adOpenStatic, adLockReadOnly
        If rst.EOF = False And rst.BOF = False Then
            lstCond.Clear
            Do Until rst.EOF
                lstCond.AddItem rst.Fields(0) & vbNullString
                rst.MoveNext
            Loop
        End If
        Set rst = Nothing
'       MsgBox FindIndexConduct
        LstValue = FindIndexConduct
        If LstValue = 1000 Then
'            txtTurns.Clear
            txtPriCros = 0
            txtTotalCon = "-"
            txtCondPer = "-"
            txtWidth = 0
            txtTop = 0
            txtSide = 0
            txtDelR = 0
            txtDelStc = 0
'            lstCond.Clear
            MsgBox "Suitable primary section not avaiable!!!", vbCritical
            cmbConStyle.SetFocus
            Exit Sub
        Else
            lstCond.ListIndex = LstValue
            lstCond.TopIndex = lstCond.ListIndex
            ListIndexofCondutor = lstCond.ListIndex
        End If
        
'        For i = 0 To lstCond.ListCount
'            If lstCond.List(i) = Trim(txtPriCros) Then
'                lstCond.ListIndex = i
'                Exit For
'            End If
'        Next
        
    End If

End If

Flag1 = False

If Len(Trim(txtDelR)) > 0 Then
    If Val(txtDelR) > 2 Then
        MsgBox "Nominal current density cannot be greater than 2!!!", vbCritical
        txtTurns.SetFocus
        Exit Sub
    End If
End If

If Len(Trim(txtDelStc)) > 0 Then
    If Val(txtDelStc) > 180 Then
        MsgBox "STC current density cannot be greater than 180!!!", vbCritical
        txtTurns.SetFocus
        Exit Sub
    End If
End If

'    lstCond.ListIndex = 0
    If txtTurns.ListCount > 1 Then
        txtTurns.ListIndex = -1
    Else
        txtTurns.ListIndex = 0
    End If
    txtTop = 0
    txtSide = 0
'    txtWidth = 0

err:

    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbExclamation
    End If

End Sub


Private Sub txtOW_GotFocus()
txtOW.SelStart = 0
txtOW.SelLength = Len(txtOW)
MDIDesign.StatusBar1.Panels(2).Text = "Set OW"
End Sub

Private Sub txtOW_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtPriCros_GotFocus()
txtPriCros.SelStart = 0
txtPriCros.SelLength = Len(txtPriCros)
MDIDesign.StatusBar1.Panels(2).Text = "Primary Cross Section (Sqmm)"
prcLastCal
End Sub

Private Sub txtPriCros_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtPriCros_LostFocus()
If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
    Dim value As Double
    Dim j As String * 1
    'Calculate minimum primary cross section area at normal or continus current
    
    For i = 0 To Val(txtnoofcores) - 1
        If InStr(1, txtRatio(i), "-") = 0 Then
            If value < txtRatio(i) Then
                value = txtRatio(i)
            End If
        Else
            For K = 1 To Len(txtRatio(i)) + 1
                j = Mid(txtRatio(i), K, K)
                If j = "-" Or K > Len(txtRatio(i)) Then
                    If value < CDbl(str1) Then
                        value = CDbl(str1)
                    End If
                    str1 = ""
                Else
                    str1 = str1 & j
                End If
            Next
        End If
    Next
    
    If Len(Trim(txtRateConti)) > 0 Then
            value = CDbl(txtRateConti)
    End If
    If CDbl(txtPriCros) > 0 Then
        txtDelR = prcCalDElRate(CDbl(txtPriCros), value)
        If InStr(1, UCase(cmbSTC), UCase("k")) > 0 Then
            txtDelStc = prcCalDElSTC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), CDbl(cmbTime), Val(txtPriCros))
        Else
            txtDelStc = prcCalDElSTC(CDbl(cmbSTC), CDbl(cmbTime), lstCond)
        End If
    End If
'    If InStr(1, UCase(cmbSTC), UCase("k")) > 0 Then
'        txtDelStc = prcCalDElSTC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), CDbl(cmbTime), Val(txtPriCros))
'    Else
'        txtDelStc = prcCalDElSTC(CDbl(cmbSTC), CDbl(cmbTime), lstCond)
'    End If
End If

End Sub

Private Sub txtpricuWt_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Primary copper Wt."
End Sub

Private Sub txtPriLength_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Length"
End Sub

Private Sub txtPrimaryCopper_GotFocus()
    txtPrimaryCopper.SelStart = 0
    txtPrimaryCopper.SelLength = Len(txtPrimaryCopper)
    MDIDesign.StatusBar1.Panels(2).Text = "Primary copper Wt"
End Sub


Private Sub txtPriMeanL_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Primary Wind. Mean Length"
End Sub


Private Sub txtPRiRes20_GotFocus()
    MDIDesign.StatusBar1.Panels(2).Text = "Primary resistance @ 20"
End Sub


Private Sub txtPRiRes75_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Primary resistance @ 75"
End Sub

Private Sub txtPriTotLength_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "PRimary Total Length"
End Sub

Private Sub txtRateConti_Change()
    If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRateContiDum_GotFocus()
    MDIDesign.StatusBar1.Panels(2).Text = "Set Rated Continuous Current"
End Sub

Private Sub txtRateContiDum_LostFocus()
    
    Dim AscRatio As ChangeRa
    Set AscRatio = New ChangeRa
    
    For p = 0 To Val(txtnoofcores) - 1
        txtRateConti = AscRatio.prcInterChangeRaredCon(txtRateContiDum)
    Next
    
    If PriRateContinus <> "" Then
        If CDbl(PriRateContinus) <> CDbl(txtRateConti) And G1.Cols > 1 Or txtTurns > 0 Then
            joel = MsgBox("This will remove the primary and secondary calculation" + vbCrLf + "Are you sure you want to change Rate-Continouse-Current ?", vbCritical + vbYesNo)
            If joel = vbYes Then
                Mould.Rows = 1
                Casting.Rows = 1
                Former.Rows = 1
                wounded.Rows = 1
                
                txtdrg.Text = ""
                cmbType.Clear
                
                ''Primary
                txtPriMeanL = ""
                txtPriLength = ""
                txtPriTotLength = ""
                txtpricuWt = ""
                txtPRiRes20 = ""
                txtPRiRes75 = ""
                txtWattPri = ""
                txtCoreWT = ""
                txtSecCuWt = ""
                txtPrimaryCopper = ""
                txtResinWt = ""
                
                WOR1A = ""
                WOR2A = ""
                
                txtTurns = 0
                cmbConStyle = ""
                txtminCs = ""
                txtPriCros = ""
                txtTotalCon = ""
                txtCondPer = ""
                txtSide = ""
                txtTop = ""
                txtWidth = ""
                txtDelR = ""
                txtDelStc = ""
                txtIDyn = ""
                txtDelStc = ""
                lstCond.Clear
                
            End If
        End If
    End If
    
End Sub

Private Sub txtRatio_Change(Index As Integer)
If Not Chr(KeyAscii) Like "[0-9-]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtRatio_KeyPress(Index As Integer, KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9-]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtRatio2_GotFocus()
txtRatio2.SelStart = 0
txtRatio2.SelLength = Len(txtRatio2)
MDIDesign.StatusBar1.Panels(2).Text = "Core " & cmbnoofcores & "  Ratio"
End Sub

Private Sub txtRatioDum_Change(Index As Integer)
    txtRatio(Index) = txtRatioDum(Index)
End Sub

Private Sub txtRatioDum_GotFocus(Index As Integer)
txtRatioDum(Index).SelStart = 0
txtRatioDum(Index).SelLength = Len(txtRatioDum(Index))
MDIDesign.StatusBar1.Panels(2).Text = "Set Primary Current"
End Sub

Private Sub txtRatioDum_KeyPress(Index As Integer, KeyAscii As Integer)
txtRatioDum(Index).Locked = False
If Not Chr(KeyAscii) Like "[0-9-.]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtRatioR2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Chr(KeyAscii) Like "[0-9-.]" And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtRatioR2Dum_Change(Index As Integer)
txtRatioR2(Index) = txtRatioR2Dum(Index)
MDIDesign.StatusBar1.Panels(2).Text = "Set Secondary Current"
End Sub

Private Sub txtRatioR2Dum_GotFocus(Index As Integer)
txtRatioR2Dum(Index).SelStart = 0
txtRatioR2Dum(Index).SelLength = Len(txtRatioR2Dum(Index))
End Sub

Private Sub txtRatioR2Dum_KeyPress(Index As Integer, KeyAscii As Integer)
txtRatioR2Dum(Index).Locked = False
If Not Chr(KeyAscii) Like "[0-9.]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtRCT_GotFocus()
txtRct.SelStart = 0
txtRct.SelLength = Len(txtRct)
MDIDesign.StatusBar1.Panels(2).Text = "Set RCT"
End Sub


Private Sub txtResinWt_GotFocus()
txtResinWt.SelStart = 0
txtResinWt.SelLength = Len(txtResinWt)
MDIDesign.StatusBar1.Panels(2).Text = "Resin Wt."
End Sub

Private Sub txtSecCuWt_GotFocus()
txtSecCuWt.SelStart = 0
txtSecCuWt.SelLength = Len(txtSecCuWt)
MDIDesign.StatusBar1.Panels(2).Text = "Secondary Copper Wt."
End Sub


Private Sub txtSecTap_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Secondary Tap Turns"
End Sub

Private Sub txtSecTurn_GotFocus()
txtSecTurn.SelStart = 0
txtSecTurn.SelLength = Len(txtSecTurn)
MDIDesign.StatusBar1.Panels(2).Text = "Secondary Turns"
End Sub

Private Sub txtSide_GotFocus()
txtSide.SelStart = 0
txtSide.SelLength = Len(txtSide)
MDIDesign.StatusBar1.Panels(2).Text = "Side / Inside Winding Ht. (mm)"
End Sub

Private Sub txtSwg_Change()
'    'Finding Minimun C/S
'    If CDbl(txtSecCS) >= CDbl(txtMinCsSec) Then
'        txtMinCsSec = txtSecCS
'    End If
Dim rst1 As ADODB.Recordset
Set rst1 = New ADODB.Recordset
rst1.Open "select * from secconductor where t_swg like '" & txtSwg & "%' ", con, adOpenStatic, adLockReadOnly
If rst1.EOF = False And rst1.BOF = False Then
'    txtSwg = rst1!t_swg
'    txtNo = "S"
    txtODs = rst1!t_ods
    txtODn = rst1!t_cs
'    txtODn.Tag = rst1!t_cs
End If
Set rst1 = Nothing
End Sub

Private Sub txtSwg_GotFocus()
txtSwg.SelStart = 0
txtSwg.SelLength = Len(txtSwg)
MDIDesign.StatusBar1.Panels(2).Text = "Set conductor"
End Sub

Private Sub txtSwg_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtTop_GotFocus()
txtTop.SelStart = 0
txtTop.SelLength = Len(txtTop)
MDIDesign.StatusBar1.Panels(2).Text = "Top Winding Ht. (mm)"
End Sub

Private Sub txtTotalCon_GotFocus()
txtTotalCon.SelStart = 0
txtTotalCon.SelLength = Len(txtTotalCon)
MDIDesign.StatusBar1.Panels(2).Text = "Set Total Conductor"
End Sub

Private Sub txtTurns_Change()
    If Len(Trim(txtTurns)) = 0 Then Exit Sub
    If GCore.Rows = 1 Then Exit Sub
    If Grid.Rows = 1 Then Exit Sub
    If Flag1 = False Then
        CalculateDataForTurns1
    End If
End Sub

Private Sub txtTurns_Click()
    If Len(Trim(txtTurns)) = 0 Then Exit Sub
    If GCore.Rows = 1 Then Exit Sub
    If Grid.Rows = 1 Then Exit Sub
    If Flag1 = False Then
        CalculateDataForTurns1
        If Val(txtTurns) > 0 Then
            If Len(cmbSTC) > 0 Then
                txtIDATurns = CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1)) * 2.5 * txtTurns
            End If
        End If
    End If
End Sub

Private Sub txtTurns_GotFocus()
txtTurns.SelStart = 0
txtTurns.SelLength = Len(txtTurns)
MDIDesign.StatusBar1.Panels(2).Text = "Set Primary Turns (Nos)"
End Sub

Private Sub txtturns_KeyPress(KeyAscii As Integer)
'If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
'End If
End Sub


Private Sub txtturns_LostFocus()
On Error Resume Next
If Trim(txtTurns) = "" Then
    SSTab1.Tab = 2
    txtTurns.SetFocus
    Exit Sub
End If
If G1.Cols > 1 Then
    If Len(PriTurns) = 0 Then PriTurns = 0
    If Len(PriTurns) = 0 Or CInt(PriTurns) = 0 Then
        PriTurns = CInt(txtTurns)
    ElseIf Val(PriTurns) <> Val(txtTurns) Then
        joel = MsgBox("Change in turns will remove the prior calculations", vbInformation + vbYesNo)
        If joel = vbYes Then
            G1.Cols = 1
            prcLastCal
            txtpricuWt = ""
            txtPriLength = ""
            txtPriMeanL = ""
            txtPriTotLength = ""
            txtResinWt = ""
            txtPrimaryCopper = ""
            txtSecCuWt = ""
            txtCoreWT = ""
            txtdrg = ""
            cmbType = ""
             WOR1A = ""
            WOR2A = ""

            PriTurns = Val(txtTurns)
        Else
            txtTurns = PriTurns
        End If
    End If
Else
    If Len(txtTurns) > 0 Then
        PriTurns = CInt(txtTurns)
    End If
End If



'    Grid2.Rows = 1
'    If InStr(1, cmbEquipType, "WINDOW") > 0 Or InStr(1, cmbEquipType, "BAR") Then
'        If InStr(1, UCase(cmbEquipType), "BAR") Then
'            cmbConStyle.Clear
'            cmbConStyle.AddItem "BAR-C"
'            cmbConStyle.AddItem "BAR-R"
'        Else
'            cmbConStyle.Clear
'            cmbConStyle.AddItem "Double Coil - Bare Conductor"
'            cmbConStyle.AddItem "DSingle Coil - Bare Conductor"
'        End If
'    Else
'        If CDbl(txtTurns) = 1 Then
'            cmbConStyle.Clear
'            cmbConStyle.AddItem "Double Coil - Bare Conductor"
'            cmbConStyle.AddItem "Single Coil - Bare Conductor"
'            cmbConStyle.AddItem "Double Coil - Normal Wdg Ht."
'            cmbConStyle.AddItem "Double Coil - Equall Wdg Ht."
'            cmbConStyle.AddItem "Single Coil"
'            cmbConStyle.AddItem "Series Coil"
'        End If
'    End If
    
    'while adding new equipment during wound turns = 0
'    cmbConStyle.ListIndex = 0
    If Val(txtTurns) > 0 Then
        If Len(cmbSTC) > 0 Then
            txtIDATurns = CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), "K") - 1)) * 2.5 * txtTurns
        End If
    End If

'If Len(cmbConStyle) > 0 Then
''   Earler
''    PrcCalPrimaryDetails
''   New
'    CalculateDataForTurns
'End If
'cmbConStyle.ListIndex = 0


End Sub


Private Sub CalculateDataForTurns()
    On Error GoTo err
    Dim rstCon As Recordset
    Dim FlagCounter As String
    
    '++++++++++++++
    
    ''setting coil type as per conductor
    ''Note
    ''width will be set as per conductor (BAR TYPE IS TEMParary)
'    lstCond.Clear
    cmbConStyle.Enabled = True
    If cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Bare" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double" Then
                    
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = '" & Trim(txtPriCros) & "'", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
'                txtCondPer = rstCon!t_COND_coi & vbNullString
'                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
                
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Bare" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Single" Then
                        
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = '" & Trim(txtPriCros) & "'", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
'                txtCondPer = rstCon!t_COND_coi & vbNullString
'                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double-Different" Then
                               
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = '" & Trim(txtPriCros) & "'", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
'                txtCondPer = rstCon!t_COND_coi & vbNullString
'                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing

    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double-Equal" Then
        
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = '" & Trim(txtPriCros) & "'", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
'                txtCondPer = rstCon!t_COND_coi & vbNullString
'                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Series" Then
        
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = '" & Trim(txtPriCros) & "'", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
'                txtCondPer = rstCon!t_COND_coi & vbNullString
'                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Single" Then
        
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_csp = '" & Trim(txtPriCros) & "'", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
'                txtCondPer = rstCon!t_COND_coi & vbNullString
'                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
    
    End If
    
    Exit Sub
    ''+++++++++++++++++++++++++++++
      
    Dim ConCurr As Double
    Set rst = New Recordset
    rst.Open "select ncd from normalcd", con, adOpenStatic, adLockReadOnly
    
    If rst.EOF = False And rst.BOF = False Then
    
        Dim j As String * 1
        Dim value As Double
        Dim str1 As String
        If Len(txtTurns) <> 0 Then
            'Calculate minimum primary cross section area at normal or continus current
            
            For i = 0 To Val(txtnoofcores) - 1
    '        MsgBox txtRatio(i)
                If InStr(1, txtRatio(i), "-") = 0 Then
                    If value < txtRatio(i) Then
                        value = txtRatio(i)
                    End If
                Else
                    For K = 1 To Len(txtRatio(i)) + 1
                        j = Mid(txtRatio(i), K, K)
                        If j = "-" Or K > Len(txtRatio(i)) Then
                            If value < CDbl(str1) Then
                                value = CDbl(str1)
                            End If
                            str1 = ""
                        Else
                            str1 = str1 & j
                        End If
                    Next
                End If
            Next
            
            If Len(Trim(txtRateConti)) > 0 Then
                value = CDbl(txtRateConti)
            End If
    '        txtAmper = prcPriAmper(value, txtRatioR2, txtTurns, rst!NCD)
    '        txtAmper = prcPriAmper(Mid(value, 2), 10, txtTurns, rst!NCD)
            
            STCCurr = prcCalNorPC(value)
            If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
                ConCurr = prcCalPC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), 0, CDbl(cmbTime))
            End If
            
            If STCCurr >= ConCurr Then
                txtPriCros = STCCurr
            ElseIf STCCurr < ConCurr Then
                txtPriCros = ConCurr
            End If
            If Len(txtPriCros) = 0 Then
                txtPriCros = 0
            End If
            txtminCs = txtPriCros
            
            If Trim(txtPriCros) > 0 And value > 0 Then
                txtDelR = prcCalDElRate(CDbl(txtPriCros), value)
            End If
            If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
                txtDelStc = prcCalDElSTC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), CDbl(cmbTime), txtPriCros)
            End If
            lstCond.Clear
            
            'Setting 2 greater than and 2 less than C/s
            Set rst = New ADODB.Recordset
            rst.Open "select distinct t_csp from primaryCond11 where t_csp > " & Trim(txtPriCros) & " and type = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'  and style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and t_np = " & Trim(txtTurns) & " order by t_csp", con, adOpenStatic, adLockReadOnly
            If rst.EOF = False And rst.BOF = False Then
                rst.MoveFirst
                If rst.RecordCount < 2 Then
                    lstCond.AddItem rst!t_csp
                    lstCond.ListIndex = 0
                Else
                    For i = 0 To 1
                        FlagCounter = "1"
                        lstCond.AddItem rst!t_csp
                        rst.MoveNext
                    Next
                    lstCond.ListIndex = 0
                End If
                'Setting 2 less than and 2 less than C/s1
                Set rst = New ADODB.Recordset
                rst.Open "select distinct t_csp from primaryCond11 where t_csp <= " & Trim(txtPriCros) & " and type = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'  and style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and t_np = " & Trim(txtTurns) & " order by t_csp desc", con, adOpenStatic, adLockReadOnly
                If rst.EOF = False And rst.BOF = False Then
                    rst.MoveFirst
                    If rst.RecordCount < 2 Then
                        lstCond.AddItem rst!t_csp
                        lstCond.ListIndex = 0
                    Else
                        For i = 0 To 1
                              lstCond.AddItem rst!t_csp
                              rst.MoveNext
                        Next
                        lstCond.ListIndex = 0
                    End If
                Else
                    MsgBox "Cross-Section not available." & vbCrLf & "If required, enter details in database of primary conduct!!!", vbExclamation
                    txtPriCros = 0
                    Exit Sub
                End If
                If FlagCounter = "1" Then
'                    lstCond.ListIndex = 0
                End If
            Else
                MsgBox "Cross-Section not available." & vbCrLf & "If required, enter details in database of primary conduct!!!", vbExclamation
                txtPriCros = 0
                Exit Sub
            End If
            
    '        lstCond.AddItem txtPriCros
'            lindex = lstCond.ListIndex
                    
            For i = 0 To lstCond.ListCount
                If lstCond.List(i) = Trim(txtPriCros) Then
'                    lstCond.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
    
    End If
    
    If Len(Trim(txtDelR)) > 0 Then
        If Val(txtDelR) > 2 Then
            MsgBox "Nominal current density cannot be greater than 2!!!", vbCritical
            txtTurns.SetFocus
            Exit Sub
        End If
    End If
    
    If Len(Trim(txtDelStc)) > 0 Then
        If Val(txtDelStc) > 180 Then
            MsgBox "STC current density cannot be greater than 180!!!", vbCritical
            txtTurns.SetFocus
            Exit Sub
        End If
    End If
    
err:
    
        If Len(Trim(err.Description)) > 0 Then
            MsgBox err.Description & vbCrLf & err.Number, vbExclamation
        End If

End Sub

Private Sub CalculateDataForTurns1()
    On Error GoTo err
    Dim rstCon As Recordset
    Dim FlagCounter As String
    
    '++++++++++++++
    
    ''setting coil type as per conductor
    ''Note
    ''width will be set as per conductor (BAR TYPE IS TEMParary)
'    lstCond.Clear
    cmbConStyle.Enabled = True
    If cmbConStyle.ListIndex <> -1 Then
    If cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Bare" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double" Then
                    
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np, t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_np = " & Trim(txtTurns) & " and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
                txtCondPer = rstCon!t_COND_coi & vbNullString
                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
                
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Bare" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Single" Then
                        
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_np = " & Trim(txtTurns) & " and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
                txtCondPer = rstCon!t_COND_coi & vbNullString
                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double-Different" Then
                               
            ''Details of primary
            Set rstCon = New Recordset
            strr = "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_np = " & Trim(txtTurns) & " and t_csp = " & Trim(txtPriCros) & ""
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_np = " & Trim(txtTurns) & " and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
                txtCondPer = rstCon!t_COND_coi & vbNullString
                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing

    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Double-Equal" Then
        
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_np = " & Trim(txtTurns) & " and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
                txtCondPer = rstCon!t_COND_coi & vbNullString
                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Series" Then
        
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_np = '" & Trim(txtTurns) & "' and t_csp = '" & Trim(txtPriCros) & "'", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
                txtCondPer = rstCon!t_COND_coi & vbNullString
                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
        
    ElseIf cmbConStyle.List(cmbConStyle.ListIndex, 0) = "Paper" And cmbConStyle.List(cmbConStyle.ListIndex, 1) = "Single" Then
        
            ''Details of primary
            Set rstCon = New Recordset
            rstCon.Open "select pri_wind_width, t_np,t_cond_siz, t_COND_coi, t_whps, t_whpt, t_csp  from primarycond11 where style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and type  = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "' and t_np = " & Trim(txtTurns) & " and t_csp = " & Trim(txtPriCros) & "", con, adOpenStatic, adLockReadOnly
            If rstCon.EOF = False And rstCon.BOF = False Then
'                txtTurns = rstCon!t_np & vbNullString
'                txtWidth = rstCon!pri_wind_width & vbNullString
'                txtPriCros = rstCon!t_csp & vbNullString
                txtTop = rstCon!t_whpt & vbNullString
                txtSide = rstCon!t_whps & vbNullString
                txtCondPer = rstCon!t_COND_coi & vbNullString
                txtTotalCon = rstCon!t_cond_siz & vbNullString
            Else
                txtTop = "0"
                txtSide = "0"
'                txtCondPer = "0"
'                txtTotalCon = "0"
                MsgBox "Primary data not avilable!!!", vbInformation
            End If
            Set rstCon = Nothing
    
    End If
    End If
    Exit Sub
    ''+++++++++++++++++++++++++++++
      
    Dim ConCurr As Double
    Set rst = New Recordset
    rst.Open "select ncd from normalcd", con, adOpenStatic, adLockReadOnly
    
    If rst.EOF = False And rst.BOF = False Then
    
        Dim j As String * 1
        Dim value As Double
        Dim str1 As String
        If Len(txtTurns) <> 0 Then
            'Calculate minimum primary cross section area at normal or continus current
            
            For i = 0 To Val(txtnoofcores) - 1
    '        MsgBox txtRatio(i)
                If InStr(1, txtRatio(i), "-") = 0 Then
                    If value < txtRatio(i) Then
                        value = txtRatio(i)
                    End If
                Else
                    For K = 1 To Len(txtRatio(i)) + 1
                        j = Mid(txtRatio(i), K, K)
                        If j = "-" Or K > Len(txtRatio(i)) Then
                            If value < CDbl(str1) Then
                                value = CDbl(str1)
                            End If
                            str1 = ""
                        Else
                            str1 = str1 & j
                        End If
                    Next
                End If
            Next
            
            If Len(Trim(txtRateConti)) > 0 Then
                value = CDbl(txtRateConti)
            End If
    '        txtAmper = prcPriAmper(value, txtRatioR2, txtTurns, rst!NCD)
    '        txtAmper = prcPriAmper(Mid(value, 2), 10, txtTurns, rst!NCD)
            
            STCCurr = prcCalNorPC(value)
            If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
                ConCurr = prcCalPC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), 0, CDbl(cmbTime))
            End If
            
            If STCCurr >= ConCurr Then
                txtPriCros = STCCurr
            ElseIf STCCurr < ConCurr Then
                txtPriCros = ConCurr
            End If
            If Len(txtPriCros) = 0 Then
                txtPriCros = 0
            End If
            txtminCs = txtPriCros
            
            If Trim(txtPriCros) > 0 And value > 0 Then
                txtDelR = prcCalDElRate(CDbl(txtPriCros), value)
            End If
            If Len(cmbSTC) > 0 And Len(cmbTime) > 0 Then
                txtDelStc = prcCalDElSTC(CDbl(Mid(cmbSTC, 1, InStr(1, UCase(cmbSTC), UCase("k")) - 1)), CDbl(cmbTime), txtPriCros)
            End If
            lstCond.Clear
            
            'Setting 2 greater than and 2 less than C/s
            Set rst = New ADODB.Recordset
            rst.Open "select distinct t_csp from primaryCond11 where t_csp > " & Trim(txtPriCros) & " and type = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'  and style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and t_np = " & Trim(txtTurns) & " order by t_csp", con, adOpenStatic, adLockReadOnly
            If rst.EOF = False And rst.BOF = False Then
                rst.MoveFirst
                If rst.RecordCount < 2 Then
                    lstCond.AddItem rst!t_csp
                    lstCond.ListIndex = 0
                Else
                    For i = 0 To 1
                        FlagCounter = "1"
                        lstCond.AddItem rst!t_csp
                        rst.MoveNext
                    Next
                    lstCond.ListIndex = 0
                End If
                'Setting 2 less than and 2 less than C/s1
                Set rst = New ADODB.Recordset
                rst.Open "select distinct t_csp from primaryCond11 where t_csp <= " & Trim(txtPriCros) & " and type = '" & Mid(cmbConStyle.List(cmbConStyle.ListIndex, 0), 1, 1) & "'  and style = '" & cmbConStyle.List(cmbConStyle.ListIndex, 1) & "' and t_np = " & Trim(txtTurns) & " order by t_csp desc", con, adOpenStatic, adLockReadOnly
                If rst.EOF = False And rst.BOF = False Then
                    rst.MoveFirst
                    If rst.RecordCount < 2 Then
                        lstCond.AddItem rst!t_csp
                        lstCond.ListIndex = 0
                    Else
                        For i = 0 To 1
                              lstCond.AddItem rst!t_csp
                              rst.MoveNext
                        Next
                        lstCond.ListIndex = 0
                    End If
                Else
                    MsgBox "Cross-Section not available." & vbCrLf & "If required, enter details in database of primary conduct!!!", vbExclamation
                    txtPriCros = 0
                    Exit Sub
                End If
                If FlagCounter = "1" Then
'                    lstCond.ListIndex = 0
                End If
            Else
                MsgBox "Cross-Section not available." & vbCrLf & "If required, enter details in database of primary conduct!!!", vbExclamation
                txtPriCros = 0
                Exit Sub
            End If
            
    '        lstCond.AddItem txtPriCros
'            lindex = lstCond.ListIndex
                    
            For i = 0 To lstCond.ListCount
                If lstCond.List(i) = Trim(txtPriCros) Then
'                    lstCond.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
    
    End If
    
    If Len(Trim(txtDelR)) > 0 Then
        If Val(txtDelR) > 2 Then
            MsgBox "Nominal current density cannot be greater than 2!!!", vbCritical
            txtTurns.SetFocus
            Exit Sub
        End If
    End If
    
    If Len(Trim(txtDelStc)) > 0 Then
        If Val(txtDelStc) > 180 Then
            MsgBox "STC current density cannot be greater than 180!!!", vbCritical
            txtTurns.SetFocus
            Exit Sub
        End If
    End If
    
err:
    
        If Len(Trim(err.Description)) > 0 Then
            MsgBox err.Description & vbCrLf & err.Number, vbExclamation
        End If

End Sub

Private Sub txtVK_GotFocus()
txtVK.SelStart = 0
txtVK.SelLength = Len(txtVK)
MDIDesign.StatusBar1.Panels(2).Text = "Set VK"
End Sub


Private Sub txtWattPri_GotFocus()
MDIDesign.StatusBar1.Panels(2).Text = "Watt Loss"
End Sub

Private Sub txtWidth_GotFocus()
txtWidth.SelStart = 0
txtWidth.SelLength = Len(txtWidth)
MDIDesign.StatusBar1.Panels(2).Text = "Width (mm)"
End Sub

Private Sub txtWL_Change()

If Len(txtOL) = 0 Or Len(txtWL) = 0 Or Len(txtOW) = 0 Or Len(txtWW) = 0 Then
    Exit Sub
End If

Dim Mean, BU  As Double

'Built Up
BU = prcCalBU(txtOL, txtWL)
txtBUP = BU

'Mean
Mean = prcCalMean(txtWL, txtWW, BU)
txtMean.Text = Mean

End Sub

Private Sub txtWL_GotFocus()
txtWL.SelStart = 0
txtWL.SelLength = Len(txtWL)
MDIDesign.StatusBar1.Panels(2).Text = "Set WL"
End Sub

Private Sub txtWL_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtWL_LostFocus()
If Len(txtWL) = 0 Then
    Exit Sub
End If
Dim bal As Integer

bal = txtOL - txtOW
txtWW = txtWL - bal

End Sub

Private Sub txtWW_Change()

'If Val(txtOW) <= Val(txtWW) Then
'
'    MsgBox "Not a valid core size!", vbInformation
'    txtWW = 0
'    txtWW.SetFocus
'    Exit Sub
'
'End If

If Len(Trim(txtOW)) = 0 Or Len(Trim(txtWW)) = 0 Then

Else
If Val(txtOW) <= Val(txtWW) Then

    MsgBox "Not a valid core size!", vbInformation
    txtWW = 0
    txtWW.SetFocus
    Exit Sub

End If
End If
If Len(txtOL) = 0 Or Len(txtWL) = 0 Or Len(txtOW) = 0 Or Len(txtWW) = 0 Then
    Exit Sub
End If

Dim Mean, BU As Double

BU = prcCalBU(txtOL, txtWL)

txtBUP = BU

Mean = prcCalMean(txtWL, txtWW, BU)

txtMean.Text = Mean

End Sub

Public Function prcPriAmper(ratio As String, ratio2 As String, Turns As Double, n As Double) As String

Grid6.Cols = 6

Dim IIp As String

IIp = prcSmallPrim(ratio)

Dim stc As String
Dim intSTC As Integer

Dim IIs, Ip, IsAlp, IsValue, IpAlp, IpValue, Ns As String
Dim a, b, C As Double
Dim ISCarry As String * 1
Dim IPCarry As String * 1
Dim coll As Integer

intSTC = InStr(1, cmbSTC, "k")

stc = Trim(Left(cmbSTC, intSTC - 1))

Grid5.AllowUserResizing = flexResizeBoth
Grid5.WordWrap = True

Grid6.AllowUserResizing = flexResizeBoth
Grid6.WordWrap = True

Grid5.Rows = 1
Grid6.Rows = 1

Grid5.AllowUserResizing = flexResizeBoth

Grid5.Cols = NoofSec + 1

For i = 1 To Len(ratio2) + 1

    ISCarry = Mid(ratio2, i, i)
    
    If ISCarry = "-" Or i > Len(ratio2) Then
    
        Grid5.Rows = Grid5.Rows + 1
        Grid5.row = Grid5.Rows - 1
                
        Grid5.RowHeight(Grid5.row) = 350
        
                
        coll = 1
        
        For j = 1 To Len(ratio) + 1
        
            IPCarry = Mid(ratio, j, j)
            
            If IPCarry = "-" Or j > Len(ratio) Then
            
                a = Val(IpValue) * Turns / Val(IsValue)
                Ns = Ns & Round(a, 0) & "/"
                                            
                Grid5.TextMatrix(Grid5.row, 0) = "Core " & Grid5.row & " (" & IsValue & ")"
                Grid5.ColWidth(coll) = 1200
                                            
                Grid5.TextMatrix(Grid5.row, coll) = Round(a, 0)
                Grid5.ColWidth(coll) = 1500
                                
                
                
                coll = coll + 1
                IpValue = ""
                
            Else
            
                IpValue = IpValue + IPCarry
            
            End If
        Next
        ' IS / NCD
                
            Grid6.Rows = Grid6.Rows + 1
            Grid6.row = Grid6.Rows - 1
            Grid6.RowHeight(Grid6.row) = 350
            
            Grid6.ColWidth(Grid6.Col) = 800
            
            Grid6.TextMatrix(Grid6.row, 0) = "Core " & Grid6.row
'            Grid6.TextMatrix(Grid6.Row, 1) = Round(prcAreaCSNormal(CDbl(IsValue)), 2)
'            Grid6.TextMatrix(Grid6.Row, 2) = Round(prcAreaCSSCur(CDbl(stc), CDbl(IIp), CDbl(IsValue)), 2)
            Grid6.TextMatrix(Grid6.row, 1) = Round(prcAreaCSNormal(CDbl(IsValue)), 2)
            Grid6.TextMatrix(Grid6.row, 2) = Round(prcAreaCSSCur(CDbl(stc), CDbl(IIp), CDbl(IsValue)), 2)
            'finding suitable wire details
        
            If Grid6.TextMatrix(Grid6.row, 1) > Grid6.TextMatrix(Grid6.row, 2) Then
                            
                '+++++++++++
                
                Set rst = New ADODB.Recordset
                
                rst.Open "select * from secconductor where t_cs >= " & Grid6.TextMatrix(Grid6.row, 1) & "  ", con, adOpenStatic, adLockReadOnly
                
                If rst.EOF = False And rst.BOF = False Then
                
                    rst.MoveFirst
                        
                        Grid6.TextMatrix(Grid6.row, 3) = rst!t_cs
                        Grid6.TextMatrix(Grid6.row, 4) = rst!t_swg
                        Grid6.TextMatrix(Grid6.row, 5) = rst!t_odn
                                        
                End If
                
                rst.Close
                Set rst = Nothing
            
            Else
            
                Set rst = New ADODB.Recordset
                
                rst.Open "select * from secconductor where t_cs >= " & Grid6.TextMatrix(Grid6.row, 2) & "", con, adOpenStatic, adLockReadOnly
                
                If rst.EOF = False And rst.BOF = False Then
                
                    rst.MoveFirst
                        
                        Grid6.TextMatrix(Grid6.row, 3) = rst!t_cs
                        Grid6.TextMatrix(Grid6.row, 4) = rst!t_swg
                        Grid6.TextMatrix(Grid6.row, 5) = rst!t_odn
                                        
                End If
                
                rst.Close
                Set rst = Nothing
            
            End If
        
        IpValue = ""
        IsValue = ""
    
    Else
    
        IsValue = IsValue + ISCarry
        
    End If
    
  
    
Next

'MsgBox Left(Ns, Len(Ns) - 1)

prcPriAmper = Left(Ns, Len(Ns) - 1)

colu = coll

End Function

'prcSmallPrim(Prim As String) As Double
'Purpose : Provide Normal C Density for secondary
'Input : Secondary Core
'Returns : Area of Cur. Den. For Normal
Public Function prcSmallPrim(Prim As String) As String

TestGrid.Cols = 2
TestGrid.Rows = 1

Dim intLen As Integer
Dim value As String * 1
Dim PrimValue As String

Dim lowValue As String

For i = 1 To Len(Prim) + 1

    value = Mid(Prim, i, i)

    If value = "-" Or i > Len(Prim) Then

        TestGrid.Rows = TestGrid.Rows + 1
        TestGrid.row = TestGrid.Rows - 1
        
        TestGrid.TextMatrix(TestGrid.row, 1) = PrimValue
        
        PrimValue = ""

    Else
        
        PrimValue = PrimValue & value
    
    End If
    
Next

lowValue = TestGrid.TextMatrix(1, 1)

For j = 1 To TestGrid.Rows - 1

        If CDbl(lowValue) > CDbl(TestGrid.TextMatrix(j, 1)) Then
        
            lowValue = TestGrid.TextMatrix(j, 1)
                    
        End If

Next

prcSmallPrim = lowValue

End Function


Public Function prcCoreCheck() As Boolean

If OptInd.value = True Then
    If Len(IvAcc) = 0 Then
        MsgBox "Please select indiviual ratio", vbExclamation
        IvAcc.SetFocus
        Exit Function
    End If
End If
  
If OptInd.value = False Then
    For i = 0 To Grid.Rows - 1
        
        If Grid.TextMatrix(i, 0) = cmbnoofcores Then
            If OptInd.value = True Then
                If Val(Grid.TextMatrix(i, 18)) = IvAcc.ListIndex Then
                    MsgBox "This core already exits", vbInformation
                    prcCoreCheck = True
                    Exit Function
                End If
            Else
                    MsgBox "This core already exits", vbInformation
                    prcCoreCheck = True
                    Exit Function
            End If
        End If
    
    Next
Else
    For i = 0 To Grid.Rows - 1
        
        If Grid.TextMatrix(i, 0) = cmbnoofcores And Trim(Grid.TextMatrix(i, 1)) = Trim(txtRatio2) Then
            If OptInd.value = True Then
                If Val(Grid.TextMatrix(i, 18)) = IvAcc.ListIndex Then
                    MsgBox "This core already exits", vbInformation
                    prcCoreCheck = True
                    Exit Function
                End If
            Else
                    MsgBox "This core already exits", vbInformation
                    prcCoreCheck = True
                    Exit Function
            End If
        End If
    
    Next
End If

Dim jo As String * 1
Dim js As Integer
If Len(cmbnoofcores) = 0 Then
    MsgBox "Please select Core No.", vbInformation
    cmbnoofcores.SetFocus
    
    For j = 0 To cmbnoofcores.ListCount - 1
        js = 0
        If Grid.Rows > 1 Then
            For i = 0 To (Grid.Rows - 1)
               If Grid.TextMatrix(i, 0) = cmbnoofcores.List(j) Then
                jo = "a"
               Else
                jo = ""
                js = j
               End If
            Next
        End If
        If jo <> "a" Then
           cmbnoofcores.ListIndex = (js)
           If Len(cmbnoofcores) > 0 Then
                txtRatio2.Text = txtRatioDum(cmbnoofcores - 1) & "/" & txtRatioR2Dum(cmbnoofcores - 1)
            End If
        End If
    Next
    prcCoreCheck = True
    Exit Function
End If

If Trim(cmbCType) = "-Select Core Type-" Then
    MsgBox "Please select core type.", vbInformation
    cmbCType.SetFocus
    prcCoreCheck = True
    Exit Function
End If

prcCoreCheck = False

End Function

Public Sub prcGCore()

GCore.Rows = 1
GCore.Cols = 12
GCore.TextMatrix(GCore.row, 0) = "Core"
GCore.ColWidth(0) = 800
GCore.TextMatrix(GCore.row, 1) = "O-OD"
GCore.ColWidth(1) = 0
GCore.TextMatrix(GCore.row, 2) = "O-ID"
GCore.ColWidth(2) = 0
GCore.TextMatrix(GCore.row, 3) = "W-OD"
GCore.ColWidth(3) = 0
GCore.TextMatrix(GCore.row, 4) = "W-ID"
GCore.ColWidth(4) = 0
GCore.TextMatrix(GCore.row, 5) = "W-ID"
GCore.ColWidth(5) = 0
GCore.TextMatrix(GCore.row, 6) = "OD"
GCore.ColWidth(6) = 0
GCore.TextMatrix(GCore.row, 7) = "ID"
GCore.ColWidth(7) = 0
GCore.TextMatrix(GCore.row, 8) = ""
GCore.ColWidth(8) = 2000
GCore.TextMatrix(GCore.row, 9) = ""
GCore.ColWidth(9) = 2000
GCore.TextMatrix(GCore.row, 10) = "Mean"
GCore.ColWidth(10) = 0
GCore.TextMatrix(GCore.row, 11) = "BU"
GCore.ColWidth(11) = 0
GCore.AllowUserResizing = flexResizeBoth

End Sub

Public Sub prcFillGrid()

'Grid5.Rows = 1
'Grid6.Rows = 1

GCross.Rows = 1
GCross.Cols = 28


GCross.TextMatrix(0, 0) = "Between Tap"
GCross.TextMatrix(0, 1) = "ODs"
GCross.TextMatrix(0, 2) = "SWG"
GCross.TextMatrix(0, 3) = "C/S"
GCross.TextMatrix(0, 4) = "Nos"
GCross.TextMatrix(0, 5) = "C/S1"
GCross.TextMatrix(0, 6) = "Layer"
GCross.TextMatrix(0, 7) = "Total Turns"
GCross.TextMatrix(0, 8) = "HT"
GCross.TextMatrix(0, 9) = "Mean L T"
GCross.TextMatrix(0, 10) = "T L"
GCross.TextMatrix(0, 11) = "Res@20"
GCross.TextMatrix(0, 12) = "Res@75"
GCross.TextMatrix(0, 13) = "CWt."
GCross.TextMatrix(0, 14) = "BM"
GCross.TextMatrix(0, 15) = "C1"
GCross.TextMatrix(0, 16) = "C2"
GCross.TextMatrix(0, 17) = "VK"
GCross.TextMatrix(0, 18) = "Ex. Curr."
GCross.TextMatrix(0, 19) = "ISF"
GCross.TextMatrix(0, 20) = "IE1"
GCross.TextMatrix(0, 21) = "EC @ Nor"
GCross.TextMatrix(0, 22) = "IE2"
GCross.TextMatrix(0, 23) = "EC @ ALF"
GCross.TextMatrix(0, 24) = "SC Density"
GCross.TextMatrix(0, 25) = "Saturation"
GCross.TextMatrix(0, 26) = "Pri"
GCross.TextMatrix(0, 27) = "SC Temp Rise"

wounded.Rows = 1
wounded.Cols = 5
wounded.TextMatrix(0, 0) = "Core No"
wounded.ColWidth(0) = 1000
wounded.TextMatrix(0, 1) = ""
wounded.ColWidth(1) = 1500
wounded.TextMatrix(0, 2) = ""
wounded.ColWidth(2) = 1500
wounded.TextMatrix(0, 3) = "Heigth (mm)"
wounded.ColWidth(3) = 1000
wounded.TextMatrix(0, 4) = "Build Up (mm)"
wounded.ColWidth(4) = 1000

Grid3.Rows = 1
Grid3.Cols = 9
Grid3.TextMatrix(Grid3.row, 0) = "No"
Grid3.ColWidth(0) = 500
Grid3.TextMatrix(Grid3.row, 1) = "W-L"
Grid3.ColWidth(1) = 0
Grid3.TextMatrix(Grid3.row, 2) = "W-W"
Grid3.ColWidth(2) = 0
Grid3.TextMatrix(Grid3.row, 3) = "OA-L"
Grid3.ColWidth(3) = 0
Grid3.TextMatrix(Grid3.row, 4) = "OA-W"
Grid3.ColWidth(4) = 0
Grid3.TextMatrix(Grid3.row, 5) = "OA-L"
Grid3.ColWidth(5) = 1000
Grid3.TextMatrix(Grid3.row, 6) = "OA-W"
Grid3.ColWidth(6) = 1000
Grid3.TextMatrix(Grid3.row, 7) = "W-L"
Grid3.ColWidth(7) = 1000
Grid3.TextMatrix(Grid3.row, 8) = "W-W"
Grid3.ColWidth(8) = 1000
Grid3.AllowUserResizing = flexResizeBoth

'Grid2.Rows = 1
'Grid2.Cols = 22
'Grid2.TextMatrix(Grid2.Row, 0) = "Core No"
'Grid2.ColWidth(0) = 1000
'Grid2.TextMatrix(Grid2.Row, 1) = "Height"
'Grid2.ColWidth(1) = 800
'Grid2.TextMatrix(Grid2.Row, 2) = "Build Up"
'Grid2.ColWidth(2) = 800
'Grid2.TextMatrix(Grid2.Row, 3) = "Area"
'Grid2.ColWidth(3) = 800
'Grid2.TextMatrix(Grid2.Row, 4) = "Effective Area"
'Grid2.ColWidth(4) = 800
'Grid2.TextMatrix(Grid2.Row, 5) = "Mean Mag"
'Grid2.ColWidth(5) = 800
'Grid2.TextMatrix(Grid2.Row, 6) = "Weight"
'Grid2.ColWidth(6) = 800
'Grid2.TextMatrix(Grid2.Row, 7) = "Type"
'Grid2.ColWidth(7) = 800
'Grid2.TextMatrix(Grid2.Row, 8) = "OA Len"
'Grid2.ColWidth(8) = 800
'Grid2.TextMatrix(Grid2.Row, 9) = "OA Wit"
'Grid2.ColWidth(9) = 800
'Grid2.TextMatrix(Grid2.Row, 10) = "W Len"
'Grid2.ColWidth(10) = 800
'Grid2.TextMatrix(Grid2.Row, 11) = "W Wit"
'Grid2.ColWidth(11) = 800
'Grid2.TextMatrix(Grid2.Row, 12) = "OD"
'Grid2.ColWidth(2) = 800
'Grid2.TextMatrix(Grid2.Row, 13) = "ID"
'Grid2.ColWidth(13) = 800
'Grid2.TextMatrix(Grid2.Row, 14) = "WLsec"
'Grid2.ColWidth(14) = 800
'Grid2.TextMatrix(Grid2.Row, 15) = "Sec Turns"
'Grid2.ColWidth(15) = 800
'Grid2.TextMatrix(Grid2.Row, 16) = "Sec Cond"
'Grid2.ColWidth(16) = 800
'Grid2.TextMatrix(Grid2.Row, 17) = "Area C/S"
'Grid2.ColWidth(17) = 800
'Grid2.TextMatrix(Grid2.Row, 18) = "Rct 75"
'Grid2.ColWidth(18) = 800
'Grid2.TextMatrix(Grid2.Row, 19) = "Sec Cu wt"
'Grid2.ColWidth(19) = 800
'Grid2.TextMatrix(Grid2.Row, 20) = "Sec Layer"
'Grid2.ColWidth(20) = 800
'Grid2.TextMatrix(Grid2.Row, 21) = "Sec Wdg Ht"
'Grid2.ColWidth(21) = 800
'
'Grid2.AllowUserResizing = flexResizeBoth

Grid.Rows = 1
Grid.Cols = 20
Grid.TextMatrix(Grid.row, 0) = "Core No"
Grid.ColWidth(0) = 700
Grid.TextMatrix(Grid.row, 1) = "Ratio"
Grid.ColWidth(1) = 1000
Grid.TextMatrix(Grid.row, 1 + 1) = "Equipment"
Grid.ColWidth(1 + 1) = 0
Grid.TextMatrix(Grid.row, 2 + 1) = "HSV"
Grid.ColWidth(2 + 1) = 0
Grid.TextMatrix(Grid.row, 3 + 1) = "IL"
Grid.ColWidth(3 + 1) = 0
Grid.TextMatrix(Grid.row, 4 + 1) = "STC"
Grid.ColWidth(4 + 1) = 0
Grid.TextMatrix(Grid.row, 5 + 1) = "Time"
Grid.ColWidth(5 + 1) = 0
Grid.TextMatrix(Grid.row, 6 + 1) = "VA"
Grid.ColWidth(6 + 1) = 800
Grid.TextMatrix(Grid.row, 7 + 1) = "Class"
Grid.ColWidth(7 + 1) = 800
Grid.TextMatrix(Grid.row, 8 + 1) = "ISF"
Grid.ColWidth(8 + 1) = 800
Grid.TextMatrix(Grid.row, 9 + 1) = "ALF"
Grid.ColWidth(9 + 1) = 800
Grid.TextMatrix(Grid.row, 10 + 1) = "VK"
Grid.ColWidth(10 + 1) = 800
Grid.TextMatrix(Grid.row, 11 + 1) = "RCT(75°C)"
Grid.ColWidth(11 + 1) = 800
Grid.TextMatrix(Grid.row, 12 + 1) = "Ie"
Grid.ColWidth(12 + 1) = 800
Grid.TextMatrix(Grid.row, 13 + 1) = "Ie @"
Grid.ColWidth(13 + 1) = 800
Grid.TextMatrix(Grid.row, 15) = "Equip.Type"
Grid.ColWidth(15) = 500
Grid.TextMatrix(Grid.row, 16) = "Details"
Grid.ColWidth(16 + 1) = 0
Grid.TextMatrix(Grid.row, 17) = "Core Ratio Type"
Grid.ColWidth(17) = 500
Grid.TextMatrix(Grid.row, 18) = "Primary no"
Grid.ColWidth(18) = 500
Grid.TextMatrix(Grid.row, 19) = "Type"
Grid.ColWidth(19) = 500
Grid.AllowUserResizing = flexResizeBoth

End Sub

Private Sub GRIDCORE()
'G2 will be used for coping record from 1 to another
G2.Rows = 109
G2.Cols = 1
CTfrm.G1.ColAlignment(0) = vbLeftJustify
CTfrm.G1.ColWidth(0) = 2000
G1.Rows = 110
G1.Cols = 1
G1.TextMatrix(0, 0) = "Core ratio :"
G1.TextMatrix(1, 0) = "Core No :"
G1.TextMatrix(2, 0) = "Primary Ratio :"
G1.TextMatrix(3, 0) = "Primary Turns :"
G1.TextMatrix(4, 0) = "Primary Current :"
G1.TextMatrix(5, 0) = "Secondary Curent :"
G1.TextMatrix(6, 0) = "No :"
G1.TextMatrix(7, 0) = "Conductor :"
G1.TextMatrix(8, 0) = "Layer :"
G1.TextMatrix(9, 0) = "Winding Ht :"
G1.TextMatrix(10, 0) = "Rec75 :"
G1.TextMatrix(11, 0) = "Copper WT :"
G1.TextMatrix(12, 0) = "VK :"
G1.TextMatrix(13, 0) = "C1 :"
G1.TextMatrix(14, 0) = "C2 :"
G1.TextMatrix(15, 0) = "Ie @ Normal :"
G1.TextMatrix(16, 0) = "Ie @ ALF :"
G1.TextMatrix(17, 0) = "ISF :"
G1.TextMatrix(18, 0) = "SC Density :"
G1.TextMatrix(19, 0) = "Saturated Cur :"
G1.TextMatrix(20, 0) = "Temp Rise :"
G1.TextMatrix(11, 0) = "Excitation Cur :"
G1.TextMatrix(22, 0) = "C.E. @ Normal :"
G1.TextMatrix(23, 0) = "C.E. @ ALF :"
G1.TextMatrix(24, 0) = "BM :"
G1.TextMatrix(25, 0) = "ALF : "
G1.TextMatrix(26, 0) = "BUP :"
G1.TextMatrix(27, 0) = "Area :"
G1.TextMatrix(28, 0) = "Mean :"
G1.TextMatrix(29, 0) = "VK at (Given,0): "
G1.TextMatrix(30, 0) = "RCT (Given,0): "
G1.TextMatrix(31, 0) = "Ie (Given,0): "
G1.TextMatrix(32, 0) = "Ht :"
G1.TextMatrix(33, 0) = "VA (Given,0): "
G1.TextMatrix(34, 0) = "Class (Given,0): "
G1.TextMatrix(35, 0) = "ISF (Given,0): "
G1.TextMatrix(36, 0) = "R Er 1 :"
G1.TextMatrix(37, 0) = "R Er 2 :"
G1.TextMatrix(38, 0) = "R Er 3 :"
G1.TextMatrix(39, 0) = "R Er 4 :"
G1.TextMatrix(40, 0) = "RC Er 1 :"
G1.TextMatrix(41, 0) = "RC Er 2 :"
G1.TextMatrix(42, 0) = "RC Er 3 :"
G1.TextMatrix(43, 0) = "RC Er 4 :"
G1.TextMatrix(44, 0) = "P Er 1 :"
G1.TextMatrix(45, 0) = "P Er 2 :"
G1.TextMatrix(46, 0) = "P Er 3 :"
G1.TextMatrix(47, 0) = "P Er 4 :"
G1.TextMatrix(48, 0) = "Core Type"
G1.TextMatrix(49, 0) = "Core Style"
G1.TextMatrix(50, 0) = "Core ratio"
G1.TextMatrix(51, 0) = "OD"
G1.TextMatrix(52, 0) = "ID"
G1.TextMatrix(53, 0) = "OL"
G1.TextMatrix(54, 0) = "OW"
G1.TextMatrix(55, 0) = "WL"
G1.TextMatrix(56, 0) = "WW"
G1.TextMatrix(57, 0) = "R Er 1 :"
G1.TextMatrix(58, 0) = "R Er 2 :"
G1.TextMatrix(59, 0) = "R Er 3 :"
G1.TextMatrix(60, 0) = "R Er 4 :"
G1.TextMatrix(61, 0) = "RC Er 1 :"
G1.TextMatrix(62, 0) = "RC Er 2 :"
G1.TextMatrix(63, 0) = "RC Er 3 :"
G1.TextMatrix(64, 0) = "RC Er 4 :"
G1.TextMatrix(65, 0) = "P Er 1 :"
G1.TextMatrix(66, 0) = "P Er 2 :"
G1.TextMatrix(67, 0) = "P Er 3 :"
G1.TextMatrix(68, 0) = "P Er 4 :"
G1.TextMatrix(69, 0) = "Burden Type 1"
G1.TextMatrix(70, 0) = "Burden Type 2"
G1.TextMatrix(71, 0) = "Total HT"
G1.TextMatrix(72, 0) = "Total Layers"
G1.TextMatrix(73, 0) = "W+OD"
G1.TextMatrix(74, 0) = "W+ID"
G1.TextMatrix(75, 0) = "W+OL"
G1.TextMatrix(76, 0) = "W+OW"
G1.TextMatrix(77, 0) = "W+WL"
G1.TextMatrix(78, 0) = "W+WW"
G1.TextMatrix(79, 0) = "W+HT"
G1.TextMatrix(80, 0) = "Copper Wt"
G1.TextMatrix(81, 0) = "Core Wt"
G1.TextMatrix(82, 0) = "PPCURR1"
G1.TextMatrix(83, 0) = "PPC1"
G1.TextMatrix(84, 0) = "PPC2"
G1.TextMatrix(85, 0) = "PPC3"
G1.TextMatrix(86, 0) = "PPC4"
G1.TextMatrix(87, 0) = "1PPCURR1"
G1.TextMatrix(88, 0) = "1PPC1"
G1.TextMatrix(89, 0) = "1PPC2"
G1.TextMatrix(90, 0) = "1PPC3"
G1.TextMatrix(91, 0) = "1PPC4"
G1.TextMatrix(92, 0) = "BM"
G1.TextMatrix(93, 0) = "BM1"
G1.TextMatrix(94, 0) = "BM2"
G1.TextMatrix(95, 0) = "BM3"
G1.TextMatrix(96, 0) = "BM4"
G1.TextMatrix(97, 0) = "1BM"
G1.TextMatrix(98, 0) = "1BM1"
G1.TextMatrix(99, 0) = "1BM2"
G1.TextMatrix(100, 0) = "1BM3"
G1.TextMatrix(101, 0) = "1BM4"
G1.TextMatrix(102, 0) = "SEC.COND"
G1.TextMatrix(103, 0) = "SEC.COND DIA"
G1.TextMatrix(104, 0) = "MIN.COND DIA"
G1.TextMatrix(105, 0) = "BET-TURNS"
G1.TextMatrix(106, 0) = "COMPENS"
G1.TextMatrix(107, 0) = "CORE-MAT"
G1.TextMatrix(108, 0) = "TAP"
G1.TextMatrix(109, 0) = "MLay"
'G1.TextMatrix(57, 0) = "Core ratio"
G1.Cols = 1
End Sub

Public Sub prcAddRem()
'If Type is metering then
If Individualfrm.txtCoreType = "Metering" Then
 'visible = FALSE
        Individualfrm.ErrorRead(i).Visible = True
        Individualfrm.ErrorRead1(i).Visible = True
        
        Individualfrm.Label92(i).Visible = False
        Individualfrm.Label91(i).Visible = False
        Individualfrm.Label81(i).Visible = False
        Individualfrm.Label82(i).Visible = False
        Individualfrm.Label83(i).Visible = False
        Individualfrm.Label86(i).Visible = False
        Individualfrm.Label87(i).Visible = False
        
        Individualfrm.txtIeNor(i).Visible = False
        Individualfrm.txtIeALF(i).Visible = False
        Individualfrm.txtECatN(i).Visible = False
        Individualfrm.txtECatAlf(i).Visible = False
        Individualfrm.txtECurr(i).Visible = False
        Individualfrm.txtVKM1(i).Visible = False
        Individualfrm.txtRCTat75(i).Visible = False
        Individualfrm.txtC1(i).Visible = False
        Individualfrm.txtC2(i).Visible = False

ElseIf Individualfrm.cmbStyle = "Protection" Then
'If Type is Protection

        Individualfrm.ErrorRead(i).Visible = True
        Individualfrm.ErrorRead1(i).Visible = False
        
        Individualfrm.Label92(i).Visible = True
        Individualfrm.Label91(i).Visible = True
        Individualfrm.Label86(i).Visible = True
        Individualfrm.Label87(i).Visible = True
        Individualfrm.txtIeNor(i).Visible = True
        Individualfrm.txtIeALF(i).Visible = True
        Individualfrm.txtECatN(i).Visible = True
        Individualfrm.txtECatAlf(i).Visible = True
        
        Individualfrm.Label81(i).Visible = False
        Individualfrm.Label82(i).Visible = False
        Individualfrm.Label83(i).Visible = False
        Individualfrm.Label85(i).Visible = False
        Individualfrm.txtECurr(i).Visible = False
        Individualfrm.txtVKM1(i).Visible = False
        Individualfrm.txtRCTat75(i).Visible = False
        Individualfrm.txtC1(i).Visible = False
        Individualfrm.txtC2(i).Visible = False
        Individualfrm.txtISF(i).Visible = False
   
ElseIf Individualfrm.cmbStyle = "Special Protection" Then
'iF Type is PS

        Individualfrm.ErrorRead(i).Visible = False
        Individualfrm.ErrorRead1(i).Visible = False
        
        Individualfrm.Label92(i).Visible = False
        Individualfrm.Label91(i).Visible = False
        Individualfrm.Label86(i).Visible = False
        Individualfrm.Label87(i).Visible = False
        Individualfrm.txtIeNor(i).Visible = False
        Individualfrm.txtIeALF(i).Visible = False
        Individualfrm.txtECatN(i).Visible = False
        Individualfrm.txtECatAlf(i).Visible = False
        
        Individualfrm.Label81(i).Visible = True
        Individualfrm.Label82(i).Visible = True
        Individualfrm.Label83(i).Visible = True
        Individualfrm.Label85(i).Visible = False
        Individualfrm.txtECurr(i).Visible = True
        Individualfrm.txtVKM1(i).Visible = True
        Individualfrm.txtRCTat75(i).Visible = True
        Individualfrm.txtC1(i).Visible = True
        Individualfrm.txtC2(i).Visible = True
        Individualfrm.txtISF(i).Visible = False

End If
End Sub

Public Sub PrcControlNames()
  
Label9 = "VK>=            V,"
Label72 = "VK>=            V,"
Label10 = " RCT(75°C)<=            ohms,"
Label73 = " RCT(75°C)<=            ohms,"
End Sub

Public Function PrcRatioCheck()

Dim value As String * 1
Dim str1 As String

For i = 0 To CInt(txtnoofcores) - 1
    Set rst = New ADODB.Recordset
    rst.Open "select * from ratioselect", con, adOpenDynamic, adLockOptimistic
    For j = 1 To Len(txtRatio(i)) + 1
        value = Mid(txtRatio(i), j, j)
        If value = "-" Or j > Len(txtRatio(i)) Then
            rst.AddNew
            rst.Fields(0) = str1
            rst.Update
            str1 = ""
        Else
            str1 = str1 + value
        End If
    Next
    Set rst = Nothing
    RatioWatch(i) = ""
    Set rst = New ADODB.Recordset
    rst.Open "select * from ratioselect order by no", con, adOpenStatic, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        Do Until rst.EOF
            RatioWatch(i) = RatioWatch(i) & "-" & rst.Fields(0)
            rst.MoveNext
        Loop
    End If
    Set rst = Nothing
    
    RatioWatch(i) = Mid(RatioWatch(i), 2)
    
    Set rst = New ADODB.Recordset
    rst.Open "delete from ratioselect", con, adOpenDynamic, adLockOptimistic
    Set rst = Nothing
    
Next

End Function

Private Function prcResinThick(hsv As String) As Double

Dim hsv1 As String
If Len(cmbHSV) > 0 Then
    If InStr(1, UCase(cmbHSV), "K") > 0 Then
        hsv1 = Mid(cmbHSV, 1, InStr(1, UCase(cmbHSV), "K") - 1)
    Else
        hsv1 = cmbHSV
    End If
Else
    MsgBox "HSV Not set", vbExclamation
    hsv1 = 0
End If

Set rst = New ADODB.Recordset
rst.Open "select * from hsv where hsv = '" & CDbl(Val(hsv1)) & "'", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
    prcResinThick = rst!resin
End If

Set rst = Nothing

End Function


Public Sub prcDimen()
        
    '________Former
        Former.Cols = 3
        Former.Rows = 1
        Former.TextMatrix(Former.row, 0) = ""
        Former.ColWidth(0) = 2500
        Former.TextMatrix(Former.row, 1) = ""
        Former.ColAlignment(1) = vbCenter
        Former.ColWidth(1) = 1500
        Former.TextMatrix(Former.row, 2) = ""
        Former.ColWidth(2) = 0
        
    'FORMER INSIDE
        Former.Rows = Former.Rows + 1
        Former.row = Former.Rows - 1
        Former.RowHeight(Former.row) = 0
        Former.TextMatrix(Former.row, 0) = "Former inside"
        Former.TextMatrix(Former.row, 1) = wounded.TextMatrix(wounded.Rows - 1, 3) '- d
        Former.TextMatrix(Former.row, 2) = wounded.TextMatrix(wounded.Rows - 1, 4) '- e
                       
        Dim RThick As Double
        RThick = prcResinThick(cmbHSV)
                                          
    'Former Outer
        Former.Rows = Former.Rows + 1
        Former.row = Former.Rows - 1
        Former.RowHeight(Former.row) = 0
        Former.TextMatrix(Former.row, 0) = "Former Outside"
        Former.TextMatrix(Former.row, 1) = wounded.TextMatrix(wounded.Rows - 1, 3) + (2 * RThick)   '- d +( 2 * resin thickness)
        Former.TextMatrix(Former.row, 2) = wounded.TextMatrix(wounded.Rows - 1, 4) + (2 * RThick)   '- e +( 2 * resin thickness)
    
    '******taking it from earlier former
        Former.Rows = Former.Rows + 1
        Former.row = Former.Rows - 1
        Former.TextMatrix(Former.row, 0) = "Former Outside (mm x mm)"
        Former.TextMatrix(Former.row, 1) = Former.TextMatrix(1, 1) & " X " & Former.TextMatrix(1, 2)      '- d +( 2 * resin thickness)
        
        Former.Rows = Former.Rows + 1
        Former.row = Former.Rows - 1
        Former.TextMatrix(Former.row, 0) = "Former Inside (mm x mm)"
        Former.TextMatrix(Former.row, 1) = Former.TextMatrix(2, 1) & " X " & Former.TextMatrix(2, 2)
        
    '________Form Inside0
        Casting.Cols = 5
        Casting.Rows = 1
        Casting.TextMatrix(0, 0) = ""
        Casting.ColWidth(0) = 2000
        Casting.AllowUserResizing = flexResizeBoth
        Casting.WordWrap = True
        Casting.TextMatrix(0, 1) = "Length (mm)"
        Casting.ColAlignment(1) = vbCenter
        Casting.TextMatrix(0, 2) = "Width (mm)"
        Casting.ColAlignment(2) = vbCenter
        Casting.TextMatrix(0, 3) = "Height (mm)"
        Casting.ColAlignment(3) = vbCenter
        Casting.TextMatrix(0, 4) = "Sec.Length (mm)"
        Casting.ColAlignment(4) = vbCenter
    
    Dim Bracket As Double
    
    Set rst = New ADODB.Recordset
    If Len(Trim(cmbType)) = 0 Then
        rst.Open "select t_mbkt from mouldvol2 where t_drg = '" & txtdrg & "'", con, adOpenStatic, adLockReadOnly
    Else
        rst.Open "select t_mbkt from mouldvol2 where t_drg = '" & txtdrg & "' and t_drgtype = '" & cmbType & "'", con, adOpenStatic, adLockReadOnly
    End If
    If rst.EOF = False And rst.BOF = False Then
        Bracket = rst!t_mbkt
    End If
    Set rst = Nothing
    '______Winding Dim
    Casting.Rows = Casting.Rows + 1
    Casting.row = Casting.Rows - 1
        Casting.TextMatrix(Casting.row, 0) = "Wind. Dim(mm)"
        Casting.TextMatrix(Casting.row, 1) = Former.TextMatrix(2, 1) + (2 * txtSide) '----former outside +  2 * pri side/inside windHt
        Casting.TextMatrix(Casting.row, 2) = Round(OALLLen, 0)
        Casting.TextMatrix(Casting.row, 3) = Round(OALLWin, 0) + Bracket + txtTop + 24 'Core O All Bigest Length + BracketHT + Prim Top Wdg Ht + Prim Terminal Thickness
        Casting.TextMatrix(Casting.row, 4) = wounded.TextMatrix(wounded.Rows - 1, 3)
    
    '______MOULD NO
    Casting.Rows = Casting.Rows + 1
    Casting.row = Casting.Rows - 1
    Casting.RowHeight(Casting.row) = 0
        Casting.TextMatrix(Casting.row, 0) = "MOULD"
        Casting.TextMatrix(Casting.row, 1) = "Length"
        
    '______Mould
    Set rst = New ADODB.Recordset
    rst.Open "select t_length, t_height, t_width, t_lensec,t_moulddrg from mouldvol2 where t_drg = '" & txtdrg & "' and t_drgtype like '" & cmbType & "%'", con, adOpenDynamic, adLockOptimistic
    If rst.EOF = False And rst.BOF = False Then
    
        Casting.Rows = Casting.Rows + 1
        Casting.row = Casting.Rows - 1
            Casting.TextMatrix(Casting.row, 0) = "Mould Dim(mm)"
            Casting.TextMatrix(Casting.row, 1) = rst!t_length
            Casting.TextMatrix(Casting.row, 2) = rst!T_width
            Casting.TextMatrix(Casting.row, 3) = rst!t_height
            Casting.TextMatrix(Casting.row, 4) = rst!t_lensec
            If IsNull(rst!t_moulddrg) = True Then
                MsgBox "No Mould Drg available!!!", vbCritical
            End If
            G1.Tag = rst!t_moulddrg & vbNullString
    Else
        
        Set rst = New ADODB.Recordset
        rst.Open "select t_length, t_height, t_width, t_lensec ,t_moulddrg from mouldvol2 where t_drg = '" & txtdrg & "'", con, adOpenDynamic, adLockOptimistic
        If rst.EOF = False And rst.BOF = False Then
        
            Casting.Rows = Casting.Rows + 1
            Casting.row = Casting.Rows - 1
                Casting.TextMatrix(Casting.row, 0) = "Mould Dim"
                Casting.TextMatrix(Casting.row, 1) = rst!t_length
                Casting.TextMatrix(Casting.row, 2) = rst!T_width
                Casting.TextMatrix(Casting.row, 3) = rst!t_height
                Casting.TextMatrix(Casting.row, 4) = rst!t_lensec
                G1.Tag = rst!t_moulddrg
            
        Else
            MsgBox "Mould Data not available", vbCritical
            Mould.Rows = 1
            Casting.Rows = 1
            Former.Rows = 1
            Exit Sub
        End If
    End If
    
    Set rst = Nothing
     
    '______Epoxy
    Casting.Rows = Casting.Rows + 1
    Casting.row = Casting.Rows - 1
        Casting.TextMatrix(Casting.row, 0) = "Resin Thickness(mm)"
        Casting.TextMatrix(Casting.row, 1) = Round((Casting.TextMatrix(3, 1) - Casting.TextMatrix(1, 1)) / 2, 0)
        Casting.TextMatrix(Casting.row, 2) = Round((Casting.TextMatrix(3, 2) - Casting.TextMatrix(1, 2)) / 2, 0)
        Casting.TextMatrix(Casting.row, 3) = Round(Casting.TextMatrix(3, 3) - Casting.TextMatrix(1, 3), 0)
        Casting.TextMatrix(Casting.row, 4) = (Casting.TextMatrix(3, 4) - Casting.TextMatrix(1, 4)) / 2
        If InStr(1, (Casting.TextMatrix(Casting.row, 4)), ".") > 0 Then
            If Mid(Casting.TextMatrix(Casting.row, 4), InStr(1, (Casting.TextMatrix(Casting.row, 4)), ".") + 1) >= 5 Then
                Casting.TextMatrix(Casting.row, 4) = Mid((Casting.TextMatrix(Casting.row, 4)), 1, InStr(1, (Casting.TextMatrix(Casting.row, 4)), ".") - 1) + 1
            Else
                Casting.TextMatrix(Casting.row, 4) = Mid((Casting.TextMatrix(Casting.row, 4)), 1, InStr(1, (Casting.TextMatrix(Casting.row, 4)), ".") - 1)
            End If
        Else
            Casting.TextMatrix(Casting.row, 4) = (Casting.TextMatrix(3, 4) - Casting.TextMatrix(1, 4)) / 2
        End If
        
     'Distance Piece
        Former.Rows = Former.Rows + 1
        Former.row = Former.Rows - 1
        Former.TextMatrix(Former.row, 0) = "Distance Piece (mm)"
        Former.TextMatrix(Former.row, 1) = Casting.TextMatrix(Casting.row, 3) - RThick
        
'=============WINDOW CLEARANCES
            If Len(txtWidth) = 0 Then txtWidth = 0
            WOR1A = Round((Round(OALLWin1, 0) - txtWidth) / 2, 0)
            WOR2A = Round((Round(OALLLen1, 0) - txtSide) / 2, 0)
            
 '+++++++++++++++++++++++++++++++++++++++
 'Primary Winding Length , Total , Cu Weigth, Mean length
 Dim MenLen, leng, TotLen, CuWt As Double
 'if Primary turns = 1 or 2
 If txtTurns = 1 Or txtTurns = 2 Then
 
 'x - Mean Length = 2 ( Former outside(1) + Former outside(2) +  Distince Piece) + ( 8 * Primary side/inside height)
 MenLen = 2 * (CDbl(Former.TextMatrix(2, 1)) + CDbl(Former.TextMatrix(2, 2)) + CDbl(Former.TextMatrix(5, 1))) + (8 * CDbl(txtSide))
 txtPriMeanL = Format(MenLen, "##.00")
 'y - Length = x[upper formula] * no of primary turns
 leng = MenLen * CDbl(txtTurns)
 txtPriLength = Format(leng, "##.00")
 'z - Total Length = y + extra length || Extra length = 180
 TotLen = leng + 180
 txtPriTotLength = Format(TotLen, "##.00")
 'a -Pri Cu Weigth = (z * 8.9 * prim cross sectio area ) / 1000000
 CuWt = ((TotLen * 8.9 * txtPriCros) / 1000000) * 1.1
 txtpricuWt = Format(CuWt, "##.00")
 'if primary turns greater than 1 & 2 turns
 Else
 'x - Mean Length = 2 ( Former outside(1) + Former outside(2) +  Distince Piece) + ( 4 * Primary side/inside height)
 MenLen = 2 * (CDbl(Former.TextMatrix(2, 1)) + CDbl(Former.TextMatrix(2, 2)) + CDbl(Former.TextMatrix(5, 1))) + (4 * CDbl(txtSide))
 txtPriMeanL = Format(MenLen, "##.00")
 'y - Length = x[upper formula] * no of primary turns
 leng = MenLen * CDbl(txtTurns)
 txtPriLength = Format(leng, "##.00")
 'z - Total Length = y + extra length || Extra length = 180
 TotLen = leng + 180
 txtPriTotLength = Format(TotLen, "##.00")
 'a -Pri Cu Weigth = (z * 8.9 * prim cross sectio area ) / 1000000
 CuWt = ((TotLen * 8.9 * txtPriCros) / 1000000) * 1.1
 txtpricuWt = Format(CuWt, "##.00")
 End If
End Sub

Public Sub prcWindDiemRound()
    
    Casting.Cols = 4
    '________Form Inside0
        Casting.TextMatrix(0, 0) = ""
        Casting.ColWidth(0) = 1500
        Casting.TextMatrix(0, 1) = "OD"
        Casting.TextMatrix(0, 2) = "ID"
        Casting.TextMatrix(0, 3) = "Height"
    Casting.Rows = 1
    '______Winding Dim Filtering
    Casting.Rows = Casting.Rows + 1
    Casting.row = Casting.Rows - 1
        Casting.TextMatrix(Casting.row, 0) = "Wind. Dim(mm)"
        Casting.TextMatrix(Casting.row, 1) = Round(OALLWin, 0)
        Casting.TextMatrix(Casting.row, 2) = Round(OALLLen, 0)
        Casting.TextMatrix(Casting.row, 3) = wounded.TextMatrix(wounded.Rows - 1, 3)
        
    Set rst = New ADODB.Recordset
    If Len(Trim(cmbType)) = 0 Then
        rst.Open "SELECT t_moulddrg, t_length, t_id, t_od, t_winwid, t_winhei, t_pribarlen FROM mouldvol2 where t_drg = '" & txtdrg & "'", con, adOpenStatic, adLockReadOnly
    Else
        rst.Open "SELECT t_moulddrg, t_length, t_id, t_od, t_winwid, t_winhei, t_pribarlen FROM mouldvol2 where t_drg = '" & txtdrg & "' and t_drgtype = '" & cmbType & "'", con, adOpenStatic, adLockReadOnly
    End If
    If rst.EOF = False And rst.BOF = False Then
    Casting.Rows = Casting.Rows + 1
    Casting.row = Casting.Rows - 1
        Casting.TextMatrix(Casting.row, 0) = "Mould(mm)"
        Casting.TextMatrix(Casting.row, 1) = rst!t_od & vbNullString
        Casting.TextMatrix(Casting.row, 2) = rst!t_id & vbNullString
        Casting.TextMatrix(Casting.row, 3) = rst!t_length & vbNullString
        G1.Tag = rst!t_moulddrg & vbNullString
        If IsNull(rst!t_pribarlen) = True Or rst!t_pribarlen = 0 Then
            pribarlen = 0
        Else
            pribarlen = rst!t_pribarlen
        End If
        If cmbEquipType = "Bar HTCTs ( B-C)" Then
            WindowSize.Caption = "Primary Bar Size = " & rst!t_winwid & "mm X " & rst!t_winhei & "mm"
            If cmbEquipType.Tag = "R" Then
                WindowSize.Caption = "Bar Size : " & rst!t_winwid & "mm X " & rst!t_winhei & "mm"
                'Primary Bar Wt = (BarWidth * BarHeight * BarLength * density of copper)/ 1000 - x
                'converting to cm so x * 100
                txtPrimaryCopper = ((rst!t_winwid * rst!t_winhei * pribarlen * 8.9) / 1000000)
            Else
                WindowSize.Caption = "Bar Size = OD : " & rst!t_id & "mm"
                'Primary Bar Wt = (Pia  * ((D/2) * (D/2)) * len * 8.9) / 1000000
                'converting to cm so x * 100
                txtPrimaryCopper = (3.141 * ((rst!t_id / 2) * (rst!t_id / 2)) * pribarlen * 8.9) / 1000000
            End If
        ElseIf cmbEquipType = "Window HTCTs ( Win-C )" Then
            WindowSize.Caption = "Window Size =" ' " & rst!t_winwid & "mm X " & rst!t_winhei & "mm"
            If cmbEquipType.Tag = "R" Then
                WindowSize.Caption = "Window Size : " & rst!t_winwid & "mm X " & rst!t_winhei & "mm"
                'Primary Bar Wt = (BarWidth * BarHeight * BarLength * density of copper)/ 1000 - x
                'converting to cm so x * 100
                txtPrimaryCopper = ((rst!t_winwid * rst!t_winhei * pribarlen * 8.9) / 1000000)
            Else
                WindowSize.Caption = "Window Size   (OD) : " & rst!t_id & "mm"
                'Primary Bar Wt = (Pia  * ((D/2) * (D/2)) * len * 8.9) / 1000000
                'converting to cm so x * 100
                txtPrimaryCopper = (3.141 * ((rst!t_id / 2) * (rst!t_id / 2)) * pribarlen * 8.9) / 1000
            End If
        End If
    Else
        MsgBox "Drawing Specified does not exits in the Database", vbInformation
        Casting.Rows = 1
        txtCoreWT = ""
        txtSecCuWt = ""
        txtpricuWt = ""
        txtResinWt = ""
'        MsgBox txtpricuWt
        Exit Sub
    End If
    Set rst = Nothing
    
    Casting.Rows = Casting.Rows + 1
    Casting.row = Casting.Rows - 1
        Casting.TextMatrix(Casting.row, 0) = "Epoxy(mm)"
        Casting.TextMatrix(Casting.row, 1) = Round((Casting.TextMatrix(2, 1) - Casting.TextMatrix(1, 1)) / 2, 0)
        Casting.TextMatrix(Casting.row, 2) = Round((Casting.TextMatrix(1, 2) - Casting.TextMatrix(2, 2)) / 2, 0)
        Casting.TextMatrix(Casting.row, 3) = Round((Casting.TextMatrix(2, 3) - Casting.TextMatrix(1, 3)) / 2, 0)
        
End Sub

Public Sub prcWinddDiemRec()
    Casting.Cols = 4
        Casting.WordWrap = True
        Casting.AllowUserResizing = flexResizeBoth
    '________Form Inside0
        Casting.TextMatrix(0, 0) = ""
        Dim s As String
        s = "OverAll " & vbCrLf & "Width X Heigth"
        Casting.TextMatrix(0, 1) = s
        Casting.ColWidth(1) = 1500
        s = "Window " & vbCrLf & "Width X Height"
        Casting.TextMatrix(0, 2) = s
        Casting.ColWidth(2) = 1500
        Casting.TextMatrix(0, 3) = "Height"
        Casting.RowHeight(0) = 500
        Casting.ColWidth(3) = 1000
        
    Casting.Rows = 1
    '______Winding
    Casting.Rows = Casting.Rows + 1
    Casting.row = Casting.Rows - 1
        Casting.TextMatrix(Casting.row, 0) = "Wind. Dim"
        Casting.TextMatrix(Casting.row, 1) = Round(OALLLen, 0) & " X " & Round(OALLWin, 0)
        Casting.TextMatrix(Casting.row, 2) = Round(OALLWin1, 0) & " X " & Round(OALLLen1, 0)
'        Casting.TextMatrix(Casting.row, 1) = Round(OALLLen, 0) & " X " & Round(OALLWin, 0)
'        Casting.TextMatrix(Casting.row, 2) = Round(OALLWin1, 0) & " X " & Round(OALLLen1, 0)
        Casting.TextMatrix(Casting.row, 3) = wounded.TextMatrix(wounded.Rows - 1, 3)
        
    Dim OL, OW, WL, WW As Double
    Set rst = New ADODB.Recordset
    If Len(cmbType) = 0 Then
        rst.Open "SELECT t_length, t_id, t_od, t_winwid, t_winhei, oawid, oahie,T_moulddrg, t_pribarlen  FROM mouldvol2 where t_drg = '" & UCase(txtdrg) & "'", con, adOpenStatic, adLockReadOnly
    Else
        rst.Open "SELECT t_length, t_id, t_od, t_winwid, t_winhei, oawid, oahie,T_moulddrg , t_pribarlen   FROM mouldvol2 where t_drg = '" & UCase(txtdrg) & "' and t_drgtype = '" & cmbType & "' ", con, adOpenStatic, adLockReadOnly
    End If
    If rst.EOF = False And rst.BOF = False Then
    Casting.Rows = Casting.Rows + 1
    Casting.row = Casting.Rows - 1
        Casting.TextMatrix(Casting.row, 0) = "Mould"
        OL = rst!oawid
        OW = rst!oahie
        WL = rst!t_winwid
        WW = rst!t_winhei
        G1.Tag = rst!t_moulddrg & vbNullString
        Casting.TextMatrix(Casting.row, 1) = OL & " X " & OW
        Casting.TextMatrix(Casting.row, 2) = WL & " X " & WW
        If IsNull(rst!t_pribarlen) = True Or Len(rst!t_pribarlen) = 0 Then
            pribarlen = 0 '420
        Else
            pribarlen = rst!t_pribarlen
        End If
        Casting.TextMatrix(Casting.row, 3) = rst!t_length
        If cmbEquipType = "Bar HTCTs ( B-R )" Then
            If Len(cmbEquipType.Tag) = 0 Then
                WindowSize.Caption = "Primary Bar Size = " & rst!t_winwid & "mm X " & rst!t_winhei & "mm"
            Else
                If cmbEquipType.Tag = "R" Then
                    WindowSize.Caption = "Bar Size = " & rst!t_winwid & "mm X " & rst!t_winhei & "mm"
                    'Primary Bar Wt = (BarWidth * BarHeight * BarLength * density of copper)/ 1000 - x
                    'converting to cm so x * 100
                    txtPrimaryCopper = ((rst!t_winwid * rst!t_winhei * pribarlen * 8.9) / 1000000)
                Else
                    WindowSize.Caption = "OD = " & rst!t_id
                    'Primary Bar Wt = (Pia  * ((D/2) * (D/2)) * len * 8.9) / 1000000
                    'converting to cm so x * 100
                    txtPrimaryCopper = (3.141 * ((rst!t_id / 2) * (rst!t_id / 2)) * pribarlen * 8.9) / 1000
                End If
            End If
            
            
        ElseIf cmbEquipType = "Window HTCTs ( Win-R )" Then
            If Len(cmbEquipType.Tag) = 0 Then
                WindowSize.Caption = "Window Size = " & rst!t_winwid & "mm X " & rst!t_winhei & "mm"
            Else
                If cmbEquipType.Tag = "R" Then
                    WindowSize.Caption = "Window Size = " & rst!t_winwid & "mm X " & rst!t_winhei & "mm"
                Else
                    WindowSize.Caption = "OD =" & rst!t_id
                End If
            End If
        End If
    Else
        Casting.Rows = Casting.Rows + 1
        Casting.row = Casting.Rows - 1
        Casting.TextMatrix(Casting.row, 0) = "Mould"
        OL = 0
        OW = 0
        WL = 0
        WW = 0
'        G1.Tag = "-NA-"
        Casting.TextMatrix(Casting.row, 1) = OL & " X " & OW
        Casting.TextMatrix(Casting.row, 2) = WL & " X " & WW
        Casting.TextMatrix(Casting.row, 3) = 0
        If cmbEquipType = "Bar HTCTs ( B-R )" Then
            WindowSize.Caption = "Primary Bar Size =" & " 0 mm X 0 mm"
        ElseIf cmbEquipType = "Window HTCTs ( Win-R )" Then
            WindowSize.Caption = "Primary Bar Size =" & " 0 mm X 0 mm"
        End If
    End If
    Set rst = Nothing
    
    Casting.Rows = Casting.Rows + 1
    Casting.row = Casting.Rows - 1
        Dim Type1 As Double
        Dim Type2 As Double
        Dim Type3 As Double
        Dim Type4 As Double
        Dim Type5 As Double
        Dim Type6 As Double
        Dim Type7 As Double
        Dim Type8 As Double
        Type1 = Val(Casting.TextMatrix(1, 1))
        Type2 = Val(Mid(Casting.TextMatrix(1, 1), InStr(1, Casting.TextMatrix(1, 1), "X") + 1))
        Type3 = Val(Casting.TextMatrix(1, 2))
        Type4 = Val(Mid(Casting.TextMatrix(1, 2), InStr(1, Casting.TextMatrix(1, 2), "X") + 1))
        Type5 = Val(Casting.TextMatrix(2, 1))
        Type6 = Val(Mid(Casting.TextMatrix(2, 1), InStr(1, Casting.TextMatrix(2, 1), "X") + 1))
        Type7 = Val(Casting.TextMatrix(2, 2))
        Type8 = Val(Mid(Casting.TextMatrix(2, 2), InStr(1, Casting.TextMatrix(2, 2), "X") + 1))
        Casting.TextMatrix(Casting.row, 0) = "Epoxy"
'        Casting.TextMatrix(Casting.row, 1) = Format((OW - Round(OALLLen, 0)) / 2, 0) & " X " & Format((OL - Round(OALLWin, 0)) / 2, 0)
        Casting.TextMatrix(Casting.row, 1) = Format((Type5 - Round(Type1, 0)) / 2, 0) & " X " & Format((Type6 - Round(Type2, 0)) / 2, 0)
        Casting.TextMatrix(Casting.row, 2) = Format((Round(Type3, 0) - Type7) / 2, 0) & " X " & Format((Round(Type4, 0) - Type8) / 2, 0)
'        Casting.TextMatrix(Casting.row, 2) = Format((Round(OALLLen1, 0) - WL) / 2, 0) & " X " & Format((Round(OALLWin1, 0) - WW) / 2, 0)
        Casting.TextMatrix(Casting.row, 3) = Format((Casting.TextMatrix(2, 3) - Casting.TextMatrix(1, 3)) / 2, 0)
            
        

End Sub

Private Sub txtWW_GotFocus()
txtWW.SelStart = 0
txtWW.SelLength = Len(txtWW)
MDIDesign.StatusBar1.Panels(2).Text = "Set WW"
End Sub

Private Sub txtWW_KeyPress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "[0-9]" And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtWW_LostFocus()

Dim bal, bal1  As Integer
bal = Val(txtOL) - Val(txtOW)
bal1 = Val(txtWL) - Val(bal)

If Val(txtWW) <> bal1 Then
    MsgBox "Build-Up Mismatch!" + vbCrLf + "Confirm core size", vbCritical
    txtWL.SetFocus
    Exit Sub
End If
End Sub

Public Function CrossSection(Turns As Integer, ClassMet As Double, MinC As Double, nos As String, a As Double)

    Dim Comp As Double
    Dim CompCal As Double
    Dim turn As Double
    Dim Plus20 As Double
    Dim Sub20 As Double
    Dim DividedCon As Double
    Dim AccNo As Byte
    
    Dim RemovalComp As Double
    
    DividedCon = MinC
    txtCompen(0) = "0.000"
    If Val(txtTurns) > 0 Then
        Comp = 100 / ((Turns * txtTurns) / txtRatioR2(cmbSelctCore - 1)) * 1
    Else
        Exit Function
    End If
'    MsgBox Comp Mod ClassMet
    compensation = Comp
    If Comp < ClassMet Then
        
        If InStr(1, (ClassMet / Comp), ".") = 0 Then
'            MsgBox ClassMet / Comp
            txtCompen(0) = Format(ClassMet, "##,##.000")
            Exit Function
        End If
    End If
    
    Plus20 = ClassMet + (ClassMet * 0.2)
    Sub20 = ClassMet - (ClassMet * 0.2)
    'Calculate no of times of conductor eg S,D,T,Q
    For i = 1 To 4
        AccNo = i 'increase no oftime of cond
        CompCal = Comp / AccNo
        RemovalComp = CompCal
        count1 = 0
        Do Until RemovalComp >= Plus20
            'Searching aganist class from database
            'increase compensation of turns
            count1 = count1 + 1
            RemovalComp = CompCal * count1
            If RemovalComp >= Sub20 And RemovalComp <= Plus20 Then
                txtCompen(0) = Format(CompCal, "0.000")
                If AccNo = 2 Then
                    txtNo = "D"
                ElseIf AccNo = 3 Then
                    txtNo = "T"
                ElseIf AccNo = 4 Then
                    txtNo = "Q"
                End If
                Exit For
                
            End If
            
        Loop
    Next
    
'    MsgBox AccNo
    
    If AccNo > 1 Then
        
        DividedCon = MinC / AccNo
        
        Set rst = New ADODB.Recordset
        rst.Open "select * from secconductor where t_cs >= " & DividedCon & " ", con, adOpenStatic, adLockReadOnly
        If rst.EOF = False And rst.BOF = False Then
            txtSwg = rst!t_swg
            If AccNo = 2 Then
                txtNo = "D"
            ElseIf AccNo = 3 Then
                txtNo = "T"
            ElseIf AccNo = 4 Then
                txtNo = "Q"
            End If
            txtODs = rst!t_ods
            txtODn = rst!t_cs
        End If
        Set rst = Nothing
        
    End If
    If Len(txtCompen(0)) = 0 Then
        txtCompen(0) = 0
    End If
    txtCompen(0) = Format((txtCompen(0) * count1), "0.000")
    txtCompen(0) = Format(txtCompen(0), "0.000")
    'Label70
End Function

Private Sub ModCal()

 Dim GreaterOD, LowerID, GreaterOD1, LowerID1 As Double
'        MsgBox G1.Cols
        wounded.Rows = 1
        Dim GFormer As Double
        If G1.TextMatrix(48, 1) = "Rectangluar" Then
            GreaterOD = G1.TextMatrix(75, 1)
            GreaterOD1 = G1.TextMatrix(76, 1)
            LowerID = G1.TextMatrix(78, 1)
            LowerID1 = G1.TextMatrix(77, 1)
            wounded.AllowUserResizing = flexResizeBoth
            wounded.WordWrap = True
            Dim s As String
            s = "OverAll (mm)" & vbCrLf & "Width X Heigth"
            wounded.TextMatrix(0, 1) = s
            s = "Window (mm)" & vbCrLf & "Width X Heigth"
            wounded.TextMatrix(0, 2) = s
            
            wounded.RowHeight(0) = 500
            
            
            For i = 1 To G1.Cols - 1
                If G1.TextMatrix(75, i) > GreaterOD Then
                    GreaterOD = G1.TextMatrix(75, i)
                    GreaterOD1 = G1.TextMatrix(76, i)
                End If
                If G1.TextMatrix(78, i) < LowerID Then
                    LowerID = G1.TextMatrix(78, i)
                    LowerID1 = G1.TextMatrix(77, i)
                End If
                wounded.Rows = wounded.Rows + 1
                wounded.row = wounded.Rows - 1
                wounded.TextMatrix(wounded.row, 0) = G1.TextMatrix(1, i)
                wounded.TextMatrix(wounded.row, 1) = Round(G1.TextMatrix(75, i), 0) & " X " & Round(G1.TextMatrix(76, i), 0)
                wounded.TextMatrix(wounded.row, 2) = Round(G1.TextMatrix(77, i), 0) & " X " & Round(G1.TextMatrix(78, i), 0)
                wounded.TextMatrix(wounded.row, 3) = Round(G1.TextMatrix(79, i), 0)
                WHt = RoundUp(G1.TextMatrix(71, i))
                wounded.TextMatrix(wounded.row, 4) = G1.TextMatrix(26, i) + (2 * G1.TextMatrix(71, i))
                wounded.TextMatrix(wounded.row, 4) = Round(wounded.TextMatrix(wounded.row, 4), 0)
                'BU + 2 * WindHt
            Next
            prcDeleteDup
            '********************Final Dimension
                wounded.Rows = wounded.Rows + 1
            wounded.row = wounded.Rows - 1
            wounded.TextMatrix(wounded.row, 0) = "Max-Dim (mm)"
            wounded.TextMatrix(wounded.row, 1) = Round(GreaterOD, 0) & " X " & Round(GreaterOD1, 0)
            OALLWin = GreaterOD1
            OALLLen = GreaterOD
            OALLWin1 = LowerID1
            OALLLen1 = LowerID
            wounded.TextMatrix(wounded.row, 2) = Round(LowerID1, 0) & " X " & Round(LowerID, 0)
            wounded.TextMatrix(wounded.row, 3) = 0
            wounded.TextMatrix(wounded.row, 4) = wounded.TextMatrix(1, 4)
            For l = 1 To wounded.row - 1
                wounded.TextMatrix(wounded.row, 3) = CDbl(wounded.TextMatrix(wounded.row, 3)) + CDbl(wounded.TextMatrix(l, 3))
                If CDbl(wounded.TextMatrix(wounded.row, 4)) < CDbl(wounded.TextMatrix(l, 4)) Then
                    wounded.TextMatrix(wounded.row, 4) = wounded.TextMatrix(l, 4)
                End If
            Next
            wounded.TextMatrix(wounded.row, 3) = wounded.TextMatrix(wounded.row, 3) + (2 * ((wounded.Rows - 2) - 1)) - (6 * ((wounded.Rows - 2) - 1))
            wounded.TextMatrix(wounded.row, 3) = Round(wounded.TextMatrix(wounded.row, 3), 2)
        Else
            wounded.Rows = 1
            wounded.TextMatrix(0, 1) = "OD (mm)"
            wounded.TextMatrix(0, 2) = "ID (mm)"
            GreaterOD = G1.TextMatrix(73, 1)
            LowerID = G1.TextMatrix(74, 1)
            
            For i = 1 To G1.Cols - 1
                If G1.TextMatrix(73, i) > GreaterOD Then
                    GreaterOD = G1.TextMatrix(73, i)
                    LowerID = G1.TextMatrix(74, i)
                End If
                wounded.Rows = wounded.Rows + 1
                wounded.row = wounded.Rows - 1
                wounded.TextMatrix(wounded.row, 0) = G1.TextMatrix(1, i)
                wounded.TextMatrix(wounded.row, 1) = Round(G1.TextMatrix(73, i), 0)
                wounded.TextMatrix(wounded.row, 2) = Round(G1.TextMatrix(74, i), 0)
                WHt = RoundUp(G1.TextMatrix(71, i))
                wounded.TextMatrix(wounded.row, 3) = Round(G1.TextMatrix(79, i), 0)
                wounded.TextMatrix(wounded.row, 4) = G1.TextMatrix(26, i) + (2 * G1.TextMatrix(71, i))
                wounded.TextMatrix(wounded.row, 4) = Round(wounded.TextMatrix(wounded.row, 4), 0)
            Next
            prcDeleteDup
            '**********************Final Dimension
            wounded.Rows = wounded.Rows + 1
            wounded.row = wounded.Rows - 1
            wounded.TextMatrix(wounded.row, 0) = "Max. Dim (mm)"
            wounded.TextMatrix(wounded.row, 1) = Round(GreaterOD, 0)
            wounded.TextMatrix(wounded.row, 2) = Round(LowerID, 0)
            OALLWin = GreaterOD
            OALLLen = LowerID
            wounded.TextMatrix(wounded.row, 3) = 0
            wounded.TextMatrix(wounded.row, 4) = wounded.TextMatrix(1, 4)
            For l = 1 To wounded.row - 1
                wounded.TextMatrix(wounded.row, 3) = CDbl(wounded.TextMatrix(wounded.row, 3)) + CDbl(wounded.TextMatrix(l, 3))
                If wounded.TextMatrix(wounded.row, 4) > wounded.TextMatrix(l, 4) Then
                    wounded.TextMatrix(wounded.row, 4) = wounded.TextMatrix(l, 4)
                End If
            Next
            wounded.TextMatrix(wounded.row, 3) = wounded.TextMatrix(wounded.row, 3) + (2 * ((wounded.Rows - 2) - 1)) - (6 * ((wounded.Rows - 2) - 1))
            wounded.TextMatrix(wounded.row, 3) = Round(wounded.TextMatrix(wounded.row, 3), 2)
        End If
'    End If
End Sub

Private Sub EmdGrid()

wounded.Rows = 1
Former.Rows = 1
Casting.Rows = 1
Mould.Rows = 1


End Sub


Public Function prcCheckSave() As Boolean

If Len(txtdrg) = 0 Then
    MsgBox "Please set drg. no", vbExclamation
    txtdrg.SetFocus
    prcCheckSave = True
    Exit Function
End If

If wounded.Rows = 1 Then
    MsgBox "No wounded details available", vbExclamation
    wounded.SetFocus
    prcCheckSave = True
    Exit Function
End If

If G1.Cols = 1 Then
    MsgBox "Core details available", vbExclamation
    SSTab1.Tab = 3
    cmbSelctCore = 1
    txtHt.SetFocus
    prcCheckSave = True
    Exit Function
End If

prcCheckSave = False

End Function

Private Sub prcclear()

prcGCore
GRIDCORE
PrcControlNames

prcFillGrid
prcFillCom

cmbFreq.ListIndex = 0
cmbCType.ListIndex = 0

cmbConStyle.ListIndex = -1
txtTurns = ""
prcClearALL
OptComm.value = True
Frame4.Enabled = False
txtPriCros.Text = ""
txtTurns = ""
txtWidth = ""
txtTop = ""
txtSide = ""
txtIDyn = ""
txtIDATurns = ""
txtDelR = ""
txtDelStc = ""
lstCond.Clear
txtCondPer = ""
txtTotalCon = ""

For i = 0 To 4
    txtRatioDum(i).Visible = False
    txtRatioR2Dum(i).Visible = False
Next

cmbSTC = ""
cmbTime = ""
'cmbEquipType.ListIndex = cmbEquipType.ListCount - 1
SSTab1.Tab = 0
NoofPrim = 1
NoofSec = 1
OptCom(0).value = True
cmbMetal.ListIndex = 1
cmbStyle.ListIndex = 0
txtDesNo.Text = ""
SetPTurnComm = ""

'Grid
Grid.Rows = 1
G1.Cols = 1
wounded.Rows = 1
Mould.Rows = 1
Casting.Rows = 1
Former.Rows = 1
G2.Cols = 1
txtdrg = "PR-"
'Other details

End Sub

Public Sub PcrConductor()
    If InStr(1, UCase(cmbEquipType), "WINDOW") > 0 Or InStr(1, UCase(cmbEquipType), "BAR") Then
        If InStr(1, UCase(cmbEquipType), "BAR") Then
            cmbConStyle.Clear
            cmbConStyle.AddItem "BAR-C"
            cmbConStyle.AddItem "BAR-R"
            txtTurns = 1
        Else
            cmbConStyle.Clear
            cmbConStyle.AddItem "Double Coil - Bare Conductor"
            cmbConStyle.AddItem "Single Coil - Bare Conductor"
            txtTurns = 1
        End If
    Else
        If Val(txtTurns) = 1 Then
            cmbConStyle.Clear
            cmbConStyle.AddItem "Double Coil - Bare Conductor"
'            cmbConStyle.AddItem "Single Coil - Bare Conductor"
            cmbConStyle.AddItem "Double Coil - Normal Wdg Ht."
'            cmbConStyle.AddItem "Double Coil - Equall Wdg Ht."
'            cmbConStyle.AddItem "Single Coil"
            cmbConStyle.AddItem "Series Coil"
        End If
    End If
End Sub

Private Function PreOrder(DesCode As String) As String
Dim desRst As Recordset
Set desRst = New Recordset
desRst.CursorLocation = adUseClient
desRst.Open "select des_code,wo,it, ratio from tank where des_code='" & DesCode & "' order by wo desc", con, adOpenStatic, adLockReadOnly
    If desRst.EOF = True Then
        desRst.Close
        desRst.Open "select des_code,wo,it, ratio from tankhistory where des_code='" & DesCode & "' order by wo desc", con, adOpenStatic, adLockReadOnly
        If desRst.EOF = True And desRst.BOF = True Then
            PreOrder = ""
        Else
            PreOrder = desRst.Fields(1) & " (" & desRst.Fields(2) & ") " & vbNullString
        End If
    Else
        PreOrder = desRst.Fields(1) & " (" & desRst.Fields(2) & ") " & vbNullString
    End If
    Set desRst = Nothing
End Function
Private Function QUOTATION(DesCode As String) As String
Dim desRst As Recordset
Set desRst = New Recordset
desRst.CursorLocation = adUseClient
desRst.Open "select des_code,qno,QDATE, it from Quotation where des_code='" & DesCode & "' order by qdate desc", con, adOpenStatic, adLockReadOnly
    If desRst.EOF = True And desRst.BOF = True Then
        QUOTATION = ""
    Else
        desRst.MoveFirst
        QUOTATION = desRst!qNO & " (" & desRst!it & ")" & vbNullString
    End If
    'Set QUOTATION = Nothing
    desRst.Close
    Set desRst = Nothing
End Function
