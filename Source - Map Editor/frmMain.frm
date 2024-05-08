VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORE Map Editor"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   Visible         =   0   'False
   Begin VB.TextBox MapDescriptionTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   75
      Text            =   "New Map"
      Top             =   7440
      Width           =   6660
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   73
      Top             =   8625
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox ShowSpecialTilesChk 
      Caption         =   "Show Special Tiles"
      Height          =   255
      Left            =   4920
      TabIndex        =   52
      Top             =   480
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   43
      Top             =   8250
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox ShowBlockedTilesChk 
      Caption         =   "Show Blocked Tiles"
      Height          =   255
      Left            =   2880
      TabIndex        =   36
      Top             =   480
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox ShowStatsChk 
      Caption         =   "Show Stats"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1185
   End
   Begin VB.CheckBox WalkModeChk 
      Caption         =   "Walk Mode"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   240
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E0E
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1352
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1896
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DDA
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":231E
            Key             =   "tiles"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2642
            Key             =   "lights"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2996
            Key             =   "particle_groups"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CEA
            Key             =   "grh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":303E
            Key             =   "exits"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3392
            Key             =   "OBJs"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36E6
            Key             =   "NPCs"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.PictureBox MainView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6240
      Left            =   3615
      ScaleHeight     =   414
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   542
      TabIndex        =   2
      Top             =   855
      Width           =   8160
      Begin VB.Timer AutoSaveTimer 
         Enabled         =   0   'False
         Left            =   600
         Top             =   120
      End
   End
   Begin VB.CheckBox ShowTriggersChk 
      Caption         =   "Show Triggers"
      Height          =   255
      Left            =   6720
      TabIndex        =   81
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox GrhTools 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7275
      Left            =   45
      ScaleHeight     =   485
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   27
      Top             =   750
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton GrhPickCmd 
         Caption         =   "Pick Grh"
         Height          =   495
         Left            =   1800
         TabIndex        =   102
         Top             =   150
         Width           =   1575
      End
      Begin VB.CommandButton TileGroupsCmd 
         Caption         =   "Tile Groups"
         Height          =   495
         Left            =   0
         TabIndex        =   101
         Top             =   150
         Width           =   1695
      End
      Begin VB.Frame Frame7 
         Caption         =   "Centered"
         Height          =   855
         Left            =   0
         TabIndex        =   96
         Top             =   3840
         Width           =   1695
         Begin VB.CheckBox GrhVCenteredChk 
            Caption         =   "Vertically"
            Height          =   195
            Left            =   240
            TabIndex        =   98
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox GrhHCenteredChk 
            Caption         =   "Horizontally"
            Height          =   195
            Left            =   240
            TabIndex        =   97
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame DecorationPosition 
         Caption         =   "Decoration´s Position"
         Height          =   615
         Left            =   30
         TabIndex        =   99
         Top             =   4800
         Width           =   3375
         Begin VB.ComboBox DecorationPositionLst 
            Height          =   315
            ItemData        =   "frmMain.frx":3A3A
            Left            =   480
            List            =   "frmMain.frx":3A59
            TabIndex        =   100
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.ComboBox GrhLayerList 
         Height          =   315
         ItemData        =   "frmMain.frx":3AAF
         Left            =   1110
         List            =   "frmMain.frx":3AC2
         TabIndex        =   93
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CheckBox GrhAlphaBlendingChk 
         Caption         =   "Alpha Blending"
         Height          =   255
         Left            =   1830
         TabIndex        =   94
         Top             =   3840
         Width           =   1455
      End
      Begin VB.OptionButton GrhErase 
         Caption         =   "Erase One Layer"
         Height          =   255
         Index           =   0
         Left            =   1830
         TabIndex        =   92
         Top             =   4200
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton GrhErase 
         Caption         =   "Erase All Layers"
         Height          =   255
         Index           =   1
         Left            =   1830
         TabIndex        =   91
         Top             =   4440
         Width           =   1575
      End
      Begin MSComctlLib.TreeView tree 
         Height          =   2505
         Left            =   0
         TabIndex        =   90
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4419
         _Version        =   393217
         LineStyle       =   1
         Style           =   6
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Frame Frame1 
         Caption         =   "Angle"
         Height          =   1605
         Left            =   30
         TabIndex        =   28
         Top             =   5520
         Width           =   3375
         Begin VB.TextBox GrhAngleTxt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1920
            TabIndex        =   33
            Text            =   "0"
            Top             =   505
            Width           =   975
         End
         Begin VB.CommandButton GrhDecreaseAngleCmd 
            Caption         =   "-"
            Height          =   255
            Left            =   1320
            TabIndex        =   32
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton GrhIncreaseAngleCmd 
            Caption         =   "+"
            Height          =   255
            Left            =   1320
            TabIndex        =   31
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox picRotate 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   240
            ScaleHeight     =   825
            ScaleWidth      =   945
            TabIndex        =   30
            Top             =   240
            Width           =   975
            Begin VB.Line LineRotate 
               BorderColor     =   &H00FFFFFF&
               X1              =   500
               X2              =   500
               Y1              =   360
               Y2              =   0
            End
            Begin VB.Shape Shape1 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00FFFFFF&
               FillColor       =   &H00FFFFFF&
               Height          =   135
               Left            =   430
               Top             =   360
               Width           =   135
            End
         End
         Begin MSComctlLib.Slider GrhAngleSlider 
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            _Version        =   393216
            Max             =   360
            TickStyle       =   3
            TickFrequency   =   5
         End
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Layer"
         Height          =   195
         Left            =   630
         TabIndex        =   95
         Top             =   3540
         Width           =   390
      End
   End
   Begin VB.PictureBox ObjectsTools 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7155
      Left            =   45
      ScaleHeight     =   7155
      ScaleWidth      =   3495
      TabIndex        =   53
      Top             =   765
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton cmdnewitem 
         Caption         =   "Create Item"
         Height          =   375
         Left            =   1035
         TabIndex        =   104
         Top             =   4110
         Width           =   1335
      End
      Begin VB.ListBox OBJList 
         Height          =   1815
         Left            =   765
         TabIndex        =   72
         Top             =   720
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         Caption         =   "List Order"
         Height          =   615
         Left            =   45
         TabIndex        =   68
         Top             =   2520
         Width           =   3375
         Begin VB.OptionButton OBJOrderChk 
            Caption         =   "A to Z"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   71
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OBJOrderChk 
            Caption         =   "Z to A"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   70
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OBJOrderChk 
            Caption         =   "Items.ini"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   69
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton OBJRemoveAllCmd 
         Caption         =   "Remove All"
         Height          =   375
         Left            =   1020
         TabIndex        =   58
         Top             =   4530
         Width           =   1335
      End
      Begin VB.TextBox OBJAmountTxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   885
         TabIndex        =   55
         Text            =   "1"
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Amount:"
         Height          =   195
         Left            =   180
         TabIndex        =   57
         Top             =   3630
         Width           =   585
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   195
         Left            =   1605
         TabIndex        =   56
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.PictureBox ParticleGroupsTools 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7275
      Left            =   45
      ScaleHeight     =   7275
      ScaleWidth      =   3495
      TabIndex        =   23
      Top             =   750
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "Run Particle Editor"
         Height          =   495
         Left            =   405
         TabIndex        =   84
         Top             =   2370
         Width           =   2775
      End
      Begin VB.CommandButton RemoveAllParticleGroupsCmd 
         Caption         =   "Remove All"
         Height          =   375
         Left            =   1125
         TabIndex        =   26
         Top             =   1530
         Width           =   1335
      End
      Begin VB.ComboBox ParticleType 
         Height          =   315
         Left            =   765
         TabIndex        =   25
         Text            =   "Fountain"
         Top             =   930
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Stream Type"
         Height          =   195
         Left            =   1245
         TabIndex        =   24
         Top             =   690
         Width           =   900
      End
   End
   Begin VB.PictureBox ExitTools 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7275
      Left            =   45
      ScaleHeight     =   7275
      ScaleWidth      =   3495
      TabIndex        =   44
      Top             =   765
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton ExitRemoveAllCmd 
         Caption         =   "Remove All"
         Height          =   375
         Left            =   795
         TabIndex        =   59
         Top             =   3765
         Width           =   1455
      End
      Begin VB.CommandButton ExitPickCmd 
         Caption         =   "Pick Exit Data"
         Height          =   375
         Left            =   795
         TabIndex        =   54
         Top             =   3165
         Width           =   1455
      End
      Begin VB.TextBox ExitYCoordTxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1035
         TabIndex        =   51
         Text            =   "1"
         Top             =   2685
         Width           =   615
      End
      Begin VB.TextBox ExitXCoordTxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1035
         TabIndex        =   50
         Text            =   "1"
         Top             =   2205
         Width           =   615
      End
      Begin VB.ComboBox ExitMapsList 
         Height          =   315
         Left            =   1035
         TabIndex        =   49
         Text            =   "  ----- Choose One -----"
         Top             =   1725
         Width           =   1935
      End
      Begin VB.CommandButton AdjacentMapCmd 
         Caption         =   "Set adjacent maps..."
         Height          =   375
         Left            =   795
         TabIndex        =   45
         Top             =   1005
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   1875
         TabIndex        =   83
         Top             =   2745
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         Height          =   195
         Left            =   1875
         TabIndex        =   82
         Top             =   2265
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   555
         TabIndex        =   48
         Top             =   2745
         Width           =   195
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   555
         TabIndex        =   47
         Top             =   2265
         Width           =   195
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Map:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   435
         TabIndex        =   46
         Top             =   1725
         Width           =   435
      End
   End
   Begin VB.PictureBox LightTools 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7275
      Left            =   45
      ScaleHeight     =   7275
      ScaleWidth      =   3495
      TabIndex        =   4
      Top             =   750
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Frame Frame6 
         Caption         =   "Tool"
         Height          =   615
         Left            =   0
         TabIndex        =   85
         Top             =   5040
         Width           =   3495
         Begin VB.OptionButton LightToolChk 
            Caption         =   "Shadow"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   88
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton LightToolChk 
            Caption         =   "Base Light"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   87
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton LightToolChk 
            Caption         =   "Light"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.CommandButton PickBaseLightColorCmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pick Base Light Color"
         Height          =   375
         Left            =   1680
         MaskColor       =   &H00000080&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1628
      End
      Begin VB.TextBox Rangetxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Text            =   "1"
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton LightColorSelect 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Color..."
         Height          =   375
         Left            =   840
         MaskColor       =   &H00000080&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton PickColorCmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pick Light Color"
         Height          =   375
         Left            =   0
         MaskColor       =   &H00000080&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton EraseAllLightsCmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remove All"
         Height          =   375
         Left            =   840
         MaskColor       =   &H00000080&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4560
         Width           =   1575
      End
      Begin VB.PictureBox CurrentColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   2865
         TabIndex        =   13
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CommandButton BaseLightFillCmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Set Map Base Light"
         Height          =   375
         Left            =   840
         MaskColor       =   &H00000080&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CheckBox CornerChk 
         Caption         =   "Upper Left Corner"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   6480
         Width           =   1575
      End
      Begin VB.CheckBox CornerChk 
         Caption         =   "Upper Right Corner"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   7
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CheckBox CornerChk 
         Caption         =   "Lower Right Corner"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   6
         Top             =   6720
         Width           =   1695
      End
      Begin VB.CheckBox CornerChk 
         Caption         =   "Lower Left Corner"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   6720
         Width           =   1575
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   360
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   2760
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
         Max             =   255
         SelStart        =   255
         TickStyle       =   3
         Value           =   255
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   360
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   2160
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
         Max             =   255
         SelStart        =   255
         TickStyle       =   3
         Value           =   255
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   1560
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
         Max             =   255
         SelStart        =   255
         TickStyle       =   3
         Value           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Range"
         Height          =   195
         Left            =   840
         TabIndex        =   21
         Top             =   3960
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         Height          =   195
         Left            =   1440
         TabIndex        =   20
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Green"
         Height          =   195
         Left            =   1380
         TabIndex        =   19
         Top             =   1920
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         Height          =   195
         Left            =   1440
         TabIndex        =   18
         Top             =   2520
         Width           =   315
      End
   End
   Begin VB.PictureBox NPCTools 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7155
      Left            =   45
      ScaleHeight     =   7155
      ScaleWidth      =   3495
      TabIndex        =   60
      Top             =   750
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "Add Npc"
         Height          =   375
         Left            =   1170
         TabIndex        =   103
         Top             =   3795
         Width           =   1335
      End
      Begin VB.ListBox NPCList 
         Height          =   1815
         Left            =   795
         TabIndex        =   63
         Top             =   1050
         Width           =   2055
      End
      Begin VB.CommandButton NPCRemoveAllCmd 
         Caption         =   "Remove All"
         Height          =   375
         Left            =   1155
         TabIndex        =   61
         Top             =   4275
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "List Order"
         Height          =   615
         Left            =   75
         TabIndex        =   64
         Top             =   2850
         Width           =   3375
         Begin VB.OptionButton NPCOrderChk 
            Caption         =   "NPC.ini"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton NPCOrderChk 
            Caption         =   "Z to A"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   66
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton NPCOrderChk 
            Caption         =   "A to Z"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   65
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "NPC:"
         Height          =   195
         Left            =   1635
         TabIndex        =   62
         Top             =   690
         Width           =   375
      End
   End
   Begin VB.PictureBox TilesTools 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7275
      Left            =   45
      ScaleHeight     =   7275
      ScaleWidth      =   3495
      TabIndex        =   34
      Top             =   750
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton TriggerFillCmd 
         Caption         =   "Map Fill"
         Height          =   375
         Left            =   870
         TabIndex        =   89
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tool"
         Height          =   615
         Left            =   390
         TabIndex        =   78
         Top             =   2880
         Width           =   2655
         Begin VB.OptionButton TileToolChk 
            Caption         =   "Blocking Tool"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton TileToolChk 
            Caption         =   "Triggers"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   79
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ListBox TriggerList 
         Height          =   1425
         Left            =   870
         TabIndex        =   77
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton TriggersRemoveAllCmd 
         Caption         =   "Remove all"
         Height          =   375
         Left            =   870
         TabIndex        =   76
         Top             =   6120
         Width           =   1695
      End
      Begin VB.CommandButton UnblockAllCmd 
         Caption         =   "Unblock all"
         Height          =   375
         Left            =   870
         TabIndex        =   42
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Border´s size"
         Height          =   1095
         Left            =   870
         TabIndex        =   37
         Top             =   480
         Width           =   1695
         Begin VB.TextBox XBorderTxt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            TabIndex        =   39
            Text            =   "8"
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox YBorderTxt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            TabIndex        =   38
            Text            =   "6"
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "X"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   105
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Y"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   720
            Width           =   105
         End
      End
      Begin VB.CommandButton BlockBordersCmd 
         Caption         =   "Block Borders ON"
         Height          =   375
         Left            =   870
         TabIndex        =   35
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Map Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3675
      TabIndex        =   74
      Top             =   7460
      Width           =   1455
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&File"
      Begin VB.Menu NewMnu 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu OpenMnu 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu SaveMnu 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveNewMnu 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu EditMnu 
      Caption         =   "&Edit"
      Begin VB.Menu MnuUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MapMnu 
      Caption         =   "&Map"
      Begin VB.Menu GoToMnu 
         Caption         =   "&Go To..."
         Shortcut        =   ^G
      End
      Begin VB.Menu ResizeMnu 
         Caption         =   "&Resize"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu OptnsMnu 
      Caption         =   "&Options"
      Begin VB.Menu OptionsMnu 
         Caption         =   "&Configuration..."
      End
   End
   Begin VB.Menu ViewMnu 
      Caption         =   "&View"
      Begin VB.Menu GrhViewerMnuChk 
         Caption         =   "&Grh Viewer"
         Checked         =   -1  'True
      End
      Begin VB.Menu MiniMapMnuChk 
         Caption         =   "&Mini Map"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu AboutMnu 
      Caption         =   "&About..."
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AboutMnu_Click()

    'GrhViewer shouldn´t be on top of this
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, False
    End If
    'Neither should minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    frmAbout.Show vbModal

End Sub

Private Sub AdjacentMapCmd_Click()

    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    If Me.MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    frmAdjacentMap.Show vbModal

End Sub

Private Sub AutoSaveTimer_Timer()

    'Save file as BACKUP.map
    Engine.Map_Save_Map_To_File App.Path & "\BACKUP.Map"

End Sub

Private Sub BaseLightFillCmd_Click()

    store_action lights, fill, , , , , , , , base_light, Light_Color
    Engine.Map_Base_Light_Fill Light_Color
    Modified = True

End Sub

Private Sub BlockBordersCmd_Click()

    If BlockBordersCmd.Caption = "Block Borders ON" Then
        store_action blocking, fill, , , , , , , , , , True, Val(XBorderTxt.text), Val(YBorderTxt.text)
        Engine.Map_Edges_Blocked_Set Val(XBorderTxt.text), Val(YBorderTxt.text), True
        BlockBordersCmd.Caption = "Block Borders OFF"
    Else
        store_action blocking, fill, , , , , , , , , , False, Val(XBorderTxt.text), Val(YBorderTxt.text)
        Engine.Map_Edges_Blocked_Set Val(XBorderTxt.text), Val(YBorderTxt.text), False
        BlockBordersCmd.Caption = "Block Borders ON"
    End If
    
    Modified = True

End Sub

Private Sub cmdnewitem_Click()
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    frmNewItem.Show vbModal
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    frmDialog.Caption = "Loading particle Editor"
    frmDialog.lbldialog.Caption = "Loading Particle Editor, Please be patient."
    frmDialog.Show
    DoEvents
    Load frmParticleEditor
    frmDialog.Hide
End Sub

Private Sub Command2_Click()
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    frmNewNPC.Show vbModal
End Sub

Private Sub ExitMapsList_Click()

    If ExitMapsList.text = "Current Map" Then
        Engine.Map_Bounds_Get exit_map_max_x, exit_map_max_y
    Else
        Engine.Map_Bounds_Get_From_File ExitMapsList.text, exit_map_max_x, exit_map_max_y
    End If
    
    Label9.Caption = "Max X: " & exit_map_max_x
    Label10.Caption = "Max Y: " & exit_map_max_y
    
    If Val(ExitXCoordTxt.text) > exit_map_max_x Then
        ExitXCoordTxt.text = exit_map_max_x
    End If
    If Val(ExitYCoordTxt.text) > exit_map_max_y Then
        ExitYCoordTxt.text = exit_map_max_y
    End If

End Sub

Private Sub ExitPickCmd_Click()
    
    ExitPickCmd.Enabled = False

End Sub

Private Sub ExitRemoveAllCmd_Click()
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    
    store_action exits, Remove_all
    
    Engine.Map_Bounds_Get max_x, max_y
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            Engine.Map_Exit_Remove map_x, map_y
        Next map_y
    Next map_x

End Sub

Private Sub ExitXCoordTxt_Change()
    'MUST be an int
    ExitXCoordTxt.text = Int(Val(ExitXCoordTxt.text))
    
    'MUST be smaller than dest map´s max_x
    If Val(ExitXCoordTxt.text) > exit_map_max_x Then
        ExitXCoordTxt.text = exit_map_max_x
    End If

End Sub

Private Sub ExitYCoordTxt_Change()
    'MUST be an int
    ExitYCoordTxt.text = Int(Val(ExitYCoordTxt.text))
    
    'MUST be smaller than dest map´s max_y
    If Val(ExitYCoordTxt.text) > exit_map_max_y Then
        ExitYCoordTxt.text = exit_map_max_y
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Modified Then
        'Hide the GrhViewer, so it doesn´t cover the msgbox
        General_Form_On_Top_Set frmGrhViewer
        'Same with minimap
        General_Form_On_Top_Set frmMap
        
        Dim save As VbMsgBoxResult
        save = MsgBox("Changes have been made since this map was last saved. If you don´t save, changes will be lost. Do you want to save now?", vbYesNoCancel)
        Select Case save
            Case vbYes
                SaveMnu_Click
            Case vbCancel
                If GrhViewerMnuChk.Checked Then
                    General_Form_On_Top_Set frmGrhViewer, True
                End If
                If MiniMapMnuChk.Checked Then
                    General_Form_On_Top_Set frmMap, True
                End If
                Cancel = 1
            Case vbNo
                'Prevent re entering this loop
                Modified = False
        End Select
    End If
End Sub

Private Sub GoToMnu_Click()
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    'Neither should minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    frmGoToMapPos.Show vbModal
    
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    'Same with minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If
End Sub

Private Sub GrhLayerList_Click()
    'According to which layer was selected we enable / disable the decoration pos controls
    If GrhLayerList.ListIndex = 1 Or GrhLayerList.ListIndex = 3 Then
        DecorationPositionLst.Enabled = True
        DecorationPosition.Enabled = True
    Else
        DecorationPositionLst.Enabled = False
        DecorationPosition.Enabled = False
    End If
End Sub

Private Sub GrhViewerMnuChk_Click()

    If GrhViewerMnuChk.Checked Then
        Unload frmGrhViewer
    Else
        GrhViewerMnuChk.Checked = True
        frmGrhViewer.Show
    End If

End Sub

Private Sub MainView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Not Walk_Mode Then
        Mouse_React_to_Click
    End If
    
End Sub

Public Sub MainView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Walk or edit map
    Call MainView_MouseDown(Button, Shift, X, Y)
    
    'Update status bar
    Statusbar_Update

End Sub

Private Sub MainView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PickBaseLightColorCmd.Enabled = True
    PickColorCmd.Enabled = True
    GrhPickCmd.Enabled = True
    ExitPickCmd.Enabled = True

End Sub

Private Sub MapDescriptionTxt_Change()

    Engine.Map_Description_Set MapDescriptionTxt.text

End Sub

Private Sub MiniMapMnuChk_Click()
    If MiniMapMnuChk.Checked Then
        Unload frmMap
    Else
        frmMap.Show
    End If
End Sub

Private Sub MnuRedo_Click()

    undo_redo False

End Sub

Private Sub MnuUndo_Click()

    undo_redo True

End Sub

Private Sub NPCList_Click()

    If frmMain.GrhViewerMnuChk.Checked Then
        frmGrhViewer.Cls
        'This is really long, since it gets the char grh 1_1 based on the npc name
        Engine.Grh_Render_To_Hdc Val(General_Var_Get(resource_path & "\graphics\char.ini", "CharData" & Val(General_Var_Get(resource_path & "\scripts\npc.ini", "NPC" & NPC_Get_Index_From_Name(NPCList.text), "npc_char_data_index")), "1_1")), frmGrhViewer.hdc, 0, 0
    End If

End Sub

Private Sub NPCOrderChk_Click(index As Integer)
    Dim LoopC As Long
    Dim LoopC2 As Long
    
    Select Case index
        Case 0
            NPCList.Clear
            Load_NPC_Data NPCList
        Case 1
            'The first one has nothing to compare
            For LoopC = 1 To NPCList.ListCount
                For LoopC2 = 0 To LoopC
                    If NPCList.List(LoopC) <> "" And NPCList.List(LoopC) <> "" And NPCList.List(LoopC) < NPCList.List(LoopC2) Then
                        NPCList.AddItem NPCList.List(LoopC), LoopC2
                        NPCList.RemoveItem LoopC + 1
                    End If
                Next LoopC2
            Next LoopC
        Case 2
            'The first one has nothing to compare
            For LoopC = 1 To NPCList.ListCount
                For LoopC2 = 0 To LoopC
                    If NPCList.List(LoopC) <> "" And NPCList.List(LoopC) <> "" And NPCList.List(LoopC) > NPCList.List(LoopC2) Then
                        NPCList.AddItem NPCList.List(LoopC), LoopC2
                        NPCList.RemoveItem LoopC + 1
                    End If
                Next LoopC2
            Next LoopC
    End Select

End Sub

Private Sub NPCRemoveAllCmd_Click()
    Dim X As Long
    Dim Y As Long
    Dim max_x As Long
    Dim max_y As Long
    
    store_action NPC, Remove_all
    
    Engine.Map_Bounds_Get max_x, max_y
    
    For X = 1 To max_x
        For Y = 1 To max_y
            Engine.Map_NPC_Remove X, Y
        Next Y
    Next X

End Sub

Private Sub OBJAmountTxt_Change()
    If Val(OBJAmountTxt.text) < 0 Then
        OBJAmountTxt.text = 1
    End If
    'MUST be an int
    OBJAmountTxt.text = Int(Val(OBJAmountTxt.text))

End Sub

Private Sub OBJList_Click()

    If frmMain.GrhViewerMnuChk.Checked Then
        frmGrhViewer.Cls
        Engine.Grh_Render_To_Hdc Val(General_Var_Get(resource_path & "\scripts\item.ini", "ITEM" & Item_Get_Index_From_Name(OBJList.text), "item_grh_index")), frmGrhViewer.hdc, 0, 0
    End If

End Sub

Private Sub OBJOrderChk_Click(index As Integer)
    Dim LoopC As Long
    Dim LoopC2 As Long
    
    Select Case index
        Case 0
            OBJList.Clear
            Load_Items_Data OBJList
        Case 1
            'The first one has nothing to compare
            For LoopC = 1 To OBJList.ListCount
                For LoopC2 = 0 To LoopC
                    If OBJList.List(LoopC) <> "" And OBJList.List(LoopC) < OBJList.List(LoopC2) Then
                        OBJList.AddItem OBJList.List(LoopC), LoopC2
                        OBJList.RemoveItem LoopC + 1
                    End If
                Next LoopC2
            Next LoopC
        Case 2
            'The first one has nothing to compare
            For LoopC = 1 To OBJList.ListCount
                For LoopC2 = 0 To LoopC
                    If OBJList.List(LoopC) <> "" And OBJList.List(LoopC) > OBJList.List(LoopC2) Then
                        OBJList.AddItem OBJList.List(LoopC), LoopC2
                        OBJList.RemoveItem LoopC + 1
                    End If
                Next LoopC2
            Next LoopC
    End Select

End Sub

Private Sub OBJRemoveAllCmd_Click()
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    
    store_action object, Remove_all
    
    Engine.Map_Bounds_Get max_x, max_y
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            Engine.Map_Item_Remove map_x, map_y
        Next map_y
    Next map_x

End Sub

Private Sub PickBaseLightColorCmd_Click()

    'Can´t use it while in Walk Mode
    If Walk_Mode Then
        MsgBox "Can´t use this command while Walk Mode is on."
        Exit Sub
    End If
    
    PickColorCmd.Enabled = True
    PickBaseLightColorCmd.Enabled = False

End Sub

Private Sub ResizeMnu_Click()
    Dim view_x As Long
    Dim view_y As Long
    Dim max_x As Long
    Dim max_y As Long
    
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    'Neither should minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    frmMapResize.Show vbModal
    
    'Make sure the view remains in a good position
    Engine.Map_Bounds_Get max_x, max_y
    Engine.Engine_View_Pos_Get view_x, view_y
    
    If view_x > max_x Then view_x = max_x
    If view_y > max_y Then view_y = max_y
    
    Engine.Engine_View_Pos_Set view_x, view_y
    
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    'Same with minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If

End Sub

Private Sub ShowSpecialTilesChk_Click()
    
    Engine.Engine_Special_Tiles_Show_Toggle

End Sub

Private Sub ShowTriggersChk_Click()
    
    Engine.Engine_Triggers_Show_Toggle

End Sub

Private Sub Slider_Change(index As Integer)

    Call Slider_Scroll(index)

End Sub

Private Sub Slider_Scroll(index As Integer)
    'Set the new color to the pic
    CurrentColor.BackColor = RGB(slider(0).value, slider(1).value, slider(2).value)
    
    'Invert colors to fix bug
    Light_Color = RGB(slider(2).value, slider(1).value, slider(0).value)
    
    Dialog.color = RGB(slider(0).value, slider(1).value, slider(2).value)

End Sub

Private Sub EraseAllLightsCmd_Click()

    store_action lights, Remove_all, , , , , , , , Light
    
    'Destroy all Lights
    Engine.light_remove_all
    
    'If in walk mode, create the mouse light again
    If Walk_Mode Then
        Dim X As Long
        Dim Y As Long
        Engine.Engine_View_Pos_Get X, Y
        Cursor_Light_Index = Engine.Light_Create(X, Y, &HFFFFFF, 1)
    End If
    
    Modified = True

End Sub

Private Sub EraseAllPartGroupsCmd_Click()

    'Destroy all Particle Groups
    Engine.Particle_Group_Remove_All
    
    Modified = True

End Sub

Private Sub GrhAngleSlider_Scroll()
    'Code taked from FRedrik´s Map Editor and edited by Juan Martín Sotuyo Dodero
    angle = GrhAngleSlider.value
    
    With LineRotate
     .x2 = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(angle * PI / 180)
     .y2 = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(angle * PI / 180)
     .x1 = picRotate.width / 2
     .y1 = picRotate.height / 2
    End With
    GrhAngleTxt.text = Str(angle)

End Sub

Private Sub GrhAngleTxt_Change()

    angle = Val(GrhAngleTxt.text)
    GrhAngleSlider.value = angle
    
    With LineRotate
     .x2 = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(angle * PI / 180)
     .y2 = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(angle * PI / 180)
     .x1 = picRotate.width / 2
     .y1 = picRotate.height / 2
    End With

End Sub

Private Sub GrhDecreaseAngleCmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Code taked from FRedrik´s Map Editor and edited by Juan Martín Sotuyo Dodero
    If Button = vbLeftButton Then
        angle = angle - 1
    End If
    If Button = vbRightButton Then
        angle = angle - 5
    End If
    
    While angle < 0
        angle = 360 + angle
    Wend
    
    With LineRotate
     .x2 = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(angle * PI / 180)
     .y2 = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(angle * PI / 180)
     .x1 = picRotate.width / 2
     .y1 = picRotate.height / 2
    End With
    GrhAngleTxt.text = Str(angle)
    GrhAngleSlider.value = angle

End Sub

Private Sub GrhIncreaseAngleCmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Code taked from FRedrik´s Map Editor and edited by Juan Martín Sotuyo Dodero
    If Button = vbLeftButton Then
        angle = angle + 1
    End If
    If Button = vbRightButton Then
        angle = angle + 5
    End If
    
    While angle > 360
       angle = angle - 360
    Wend
    
    With LineRotate
     .x2 = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(angle * PI / 180)
     .y2 = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(angle * PI / 180)
     .x1 = picRotate.width / 2
     .y1 = picRotate.height / 2
    End With
    GrhAngleTxt.text = Str(angle)
    frmMain.GrhAngleSlider.value = angle

End Sub

Private Sub GrhPickCmd_Click()

    If GrhPickCmd.Enabled Then
        GrhPickCmd.Enabled = False
    Else
        GrhPickCmd.Enabled = True
    End If

End Sub

Private Sub LightColorSelect_Click()
On Error GoTo ErrHandler:

    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer
    
    Dialog.CancelError = True
    
    'GrhViewer shouldn´t be topmost anymore, or it will cover the dialog
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    Call ArrangeDialog(Dialog, 3)
    
    Light_Color = Dialog.color
    
    General_Long_Color_to_RGB Light_Color, red, green, blue
    
    'Set sliders to their new value
    slider(0).value = blue
    slider(1).value = green
    slider(2).value = red
    
    'Set GrhViewer back to it´s original state
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If
    
ErrHandler:

End Sub

Private Sub NewMnu_Click()

    'GrhViewer shouldn´t be on top of this
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    
    'Minimap shouldn't be on top of this
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    frmNewMap.Show vbModal

End Sub

Private Sub OpenMnu_Click()
On Error GoTo ErrHandler:
    
    Dialog.CancelError = True
    
    If Modified = True Then
        If MsgBox("Changes have been made since this map was last saved. If you don´t save, changes will be lost. Do you want to save now?", vbYesNo) = vbYes Then
            Call SaveMnu_Click
        End If
    End If
    
    'GrhViewer shouldn´t be topmost anymore, or it will cover the dialog
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    'Neither should minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    Call ArrangeDialog(Dialog, 1)
    
    'Disable Walk Mode
    If Walk_Mode Then
        Toggle_Walk_Mode
        WalkModeChk.value = 0
    End If
    
    If Engine.Map_Load_Map_From_File(Dialog.FileName, use_ini_files) = False Then
        MsgBox "An error ocurred when trying to load " & Dialog.FileName
        Exit Sub
    End If
    
    'Render the loaded map to the mini map
    Dim X As Long
    Dim Y As Long
    Engine.Map_Bounds_Get X, Y
    frmMap.picmain.width = X
    frmMap.picmain.height = Y
    frmMap.height = (frmMap.picmain.height + 28) * Screen.TwipsPerPixelY
    frmMap.width = (frmMap.picmain.width + 10) * Screen.TwipsPerPixelX
    
    frmMap.picmain.Cls
    
    Engine.Engine_Render_Mini_Map_To_hDC frmMap.picmain.hdc
    
    'Store current map id
    Current_Map = Dialog.FileName
    
    'Show Map´s description
    MapDescriptionTxt.text = Engine.Map_Description_Get
    
    'Load list of maps
    Load_Maps_To_ComboBox ExitMapsList
    
    'Clear action list (can´t undo/redo anything if it´s a different map)
    Clear_Action_List
    
    'Reset pos
    Engine.Engine_View_Pos_Set 5, 5
    
    'Reset mini map rect
    frmMap.shparea.top = -2
    frmMap.shparea.left = -4
    
    'Set modified flag
    Modified = False
    
ErrHandler:
    
    'Put GrhViewer to it´s initial state
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    'Same with minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If
    
    'Set the Grh tool as default
    Dim Button As Button
    Set Button = frmMain.Toolbar1.Buttons(8)
    frmMain.toolbar1_ButtonClick Button

End Sub

Private Sub OptionsMnu_Click()

    'GrhViewer shouldn´t be on top of this
    If frmMain.GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    'Neither should minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    frmOptions.Show vbModal

End Sub

Private Sub PickColorCmd_Click()

    'Can´t use it while in Walk Mode
    If Walk_Mode Then
        MsgBox "Can´t use this command while Walk Mode is on."
        Exit Sub
    End If
    
    PickColorCmd.Enabled = False
    PickBaseLightColorCmd.Enabled = True

End Sub

Private Sub Rangetxt_Change()
    If Val(Rangetxt.text) < 1 Then
        Rangetxt.text = 1
    End If
    'MUST be an int
    Rangetxt.text = Int(Val(Rangetxt.text))

End Sub

Private Sub RemoveAllParticleGroupsCmd_Click()

    store_action particle_stream, Remove_all
    
    Engine.Particle_Group_Remove_All
    Modified = True

End Sub

Public Sub SaveMnu_Click()
On Error GoTo ErrHandler:
    
    'Set Walk Mode off to avoid saving the cursor light
    If Walk_Mode Then
        Toggle_Walk_Mode
    End If
    
    DoEvents
    'GrhViewer shouldn´t be topmost anymore, or it will cover the dialog
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    'Neither should minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    If Current_Map = "" Then
        Call SaveNewMnu_Click
        Exit Sub
    Else
        frmDialog.Show
        frmDialog.Caption = "Saving"
        frmDialog.lbldialog.Caption = "Saving Please Wait"
        If Engine.Map_Save_Map_To_File(Current_Map, use_ini_files) = False Then
            
            MsgBox "An error ocurred when trying to save " & Current_Map
            frmDialog.Hide
            Exit Sub
        End If
    End If
    
    'Set changed flag
    Modified = False
ErrHandler:
    
    'Put GrhViewer back to it´s original state
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    'Same with minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If
    
    'Set the Grh tool as default
    Dim Button As Button
    Set Button = frmMain.Toolbar1.Buttons(8)
    frmMain.toolbar1_ButtonClick Button
    
    'Re-load the map list to add the new map
    Load_Maps_To_ComboBox frmMain.ExitMapsList
    frmDialog.Hide

End Sub

Private Sub SetBaselightCmd_Click()
    
    Engine.Map_Base_Light_Fill Light_Color

End Sub

Private Sub ShowBlockedTilesChk_Click()
    
    Engine.Engine_Blocked_Tiles_Show_Toggle

End Sub

Private Sub ShowStatsChk_Click()
    
    Engine.Engine_Stats_Show_Toggle

End Sub

Private Sub ExitMnu_Click()
    Dim Cancel As Integer
    
    Call Form_Unload(Cancel)

End Sub

Private Sub Form_Load()

    'Update main caption
    Me.Caption = frmMain.Caption & " V " & App.Major & "." & App.Minor & "." & App.Revision
    
    'Load toolbar
    LoadToolbar
    
    'Initiate status bars
    InitStatusBars
    
    'Set default list indexes
    GrhLayerList.ListIndex = 0
    DecorationPositionLst.ListIndex = 0
    
    'Set autosave delay
    frmMain.AutoSaveTimer.Interval = autosave_delay

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If prgRun Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub LoadMnu_Click()
On Error GoTo ErrHandler:
    
    Dialog.CancelError = True
    
    If Modified = True Then
        If MsgBox("Changes have been made since this map was last saved. If you don´t save, changes will be lost. Do you want to save now?", vbYesNo) = vbYes Then
            Engine.Map_Save_Map (Dialog.FileName)
        End If
    End If
    
    Call ArrangeDialog(Dialog, 1)
    Engine.Map_Load_Map (Dialog.FileName)
    
ErrHandler:

End Sub

Private Sub SaveNewMnu_Click()
On Error GoTo ErrHandler:
    
    'Set Walk Mode off to avoid saving the cursor light
    If Walk_Mode Then
        Toggle_Walk_Mode
    End If
    
    Dialog.CancelError = True
    
    'GrhViewer shouldn´t be topmost anymore, or it will cover the dialog
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    'Neither should minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap
    End If
    
    Call ArrangeDialog(Dialog, 2)
    frmDialog.Show
    frmDialog.Caption = "Saving"
    frmDialog.lbldialog.Caption = "Saving Please Wait"
    
    'If this is the first time the map is saved, chack for exits to the same map
    Dim map_x As Long
    Dim map_y As Long
    Dim max_map_x As Long
    Dim max_map_y As Long
    Dim dest_map As String
    Dim dest_map_x As Long
    Dim dest_map_y As Long
    Dim cur_pos As Long
    
    Engine.Map_Bounds_Get max_map_x, max_map_y
    For map_x = 1 To max_map_x
        For map_y = 1 To max_map_y
            If Engine.Map_Exit_Get(map_x, map_y, dest_map, dest_map_x, dest_map_y) Then
                If dest_map = "Current Map" Then
                    Engine.Map_Exit_Remove map_x, map_y
                    'Arrange dest map name
                    cur_pos = 1
                    Do Until Mid$(dest_map, Len(dest_map) - cur_pos, 1) = Chr(92)
                        cur_pos = cur_pos + 1
                    Loop
                    dest_map = Right$(dest_map, cur_pos)
                    Engine.Map_Exit_Add map_x, map_y, dest_map, dest_map_x, dest_map_y
                End If
            End If
        Next map_y
    Next map_x
    
    If Engine.Map_Save_Map_To_File(Dialog.FileName, use_ini_files) = False Then
        MsgBox "An error ocurred when trying to save " & Dialog.FileName
        frmDialog.Hide
        Exit Sub
    Else
        MsgBox "Map saved as " & Dialog.FileName
        frmDialog.Hide
    End If
    
    'Set modified flag
    Modified = False

ErrHandler:
    
    'Set GrhViewer to it´s original state
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    'Same with minimap
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If
    
    'Set the Grh tool as default
    Dim Button As Button
    Set Button = frmMain.Toolbar1.Buttons(8)
    frmMain.toolbar1_ButtonClick Button
    frmDialog.Hide
End Sub

Private Sub TileGroupsCmd_Click()
    frmTileGroups.Show
    
    'Set it to be allways visible
    General_Form_On_Top_Set frmTileGroups, True
End Sub

Public Sub toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim LoopC As Long
    'Check which button was clicked
    Select Case Button.Key
        Case Is = "new"     'Create new map
            frmNewMap.Show vbModal
        
        Case Is = "open"    'Open map
            OpenMnu_Click
        
        Case Is = "save"    'Save map
            SaveMnu_Click
        
        Case Is = "undo"
            undo_redo True  'Undo action
        
        Case Is = "redo"
            undo_redo False 'Redo action
        
        Case Is = "grh"
            show_grh_controls
            
        Case Is = "tiles"
            show_tiles_controls
            'Set border values to default
            XBorderTxt.text = x_border
            YBorderTxt.text = y_border
            If TriggerList.ListIndex < 0 Then
                TriggerList.ListIndex = 0
            End If
        
        Case Is = "lights"
            show_lights_controls
            'Set controls to default
            LightToolChk(0).value = vbChecked
            PickBaseLightColorCmd.Enabled = True
            PickColorCmd.Enabled = True
            
        Case Is = "particle_groups"
            show_PGs_controls
            
        Case Is = "exits"
            show_exits_controls
            ExitPickCmd.Enabled = True
        
        Case Is = "OBJs"
            show_OBJs_controls
            If OBJList.ListCount > 0 Then
                If OBJList.ListIndex < 0 Then
                    OBJList.ListIndex = 0
                End If
            End If
            
        Case Is = "NPCs"
            show_NPCs_controls
            If NPCList.ListCount > 0 Then
                If NPCList.ListIndex < 0 Then
                    NPCList.ListIndex = 0
                End If
            End If
        
    End Select
    
    'Set it as clicked
    For LoopC = 1 To Toolbar1.Buttons.count
        If Toolbar1.Buttons(LoopC).Key = Button.Key And LoopC > 7 Then
            Toolbar1.Buttons(LoopC).value = tbrPressed
            tool = Toolbar1.Buttons(LoopC).Key
        Else
            Toolbar1.Buttons(LoopC).value = tbrUnpressed
        End If
    Next LoopC
    
    'Refresh toolbar so that buttons are correctly drawn
    Toolbar1.Refresh
End Sub

Private Sub tree_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim tempidx As Long
    Dim temptxt As String
    
    'Check if it´s a grh index
    If left$(Node.Key, 3) = "grh" And left$(Node.text, 1) <> "<" Then
        tempidx = Node.index
        'Get rid of the "Grh " label
        current_grh = CLng(Right(Node.text, Len(Node.text) - 4))
        
        If GrhViewerMnuChk.Checked Then
            frmGrhViewer.Cls
            Engine.Grh_Render_To_Hdc current_grh, frmGrhViewer.hdc, 0, 0
        End If
        
        'Check if it has a specific Grh layer
        If Node.Children > 0 Then
            temptxt = Node.Child.text
            temptxt = left$(temptxt, Len(temptxt) - 8)
            temptxt = Right(temptxt, Len(temptxt) - 5)
            For tempidx = 0 To GrhLayerList.ListCount
                If GrhLayerList.List(tempidx) = temptxt Then
                    GrhLayerList.ListIndex = tempidx
                    Exit Sub
                End If
            Next tempidx
        End If
        
        'Check among siblings for Grh layer
        Dim LastSibling As Node
        Dim CurrentSibling As Node
        Set CurrentSibling = tree.Nodes(tempidx).FirstSibling
        Set LastSibling = tree.Nodes(tempidx).LastSibling
        
        Do Until CurrentSibling.Key = LastSibling.Key
            If left$(CurrentSibling.text, 1) = "<" Then
                temptxt = CurrentSibling.text
                Exit Do
            End If
            Set CurrentSibling = CurrentSibling.Next
        Loop
        
        If left$(LastSibling.text, 1) = "<" Then
            temptxt = LastSibling.text
        End If
        
        temptxt = left$(temptxt, Len(temptxt) - 8)
        temptxt = Right(temptxt, Len(temptxt) - 5)
        For tempidx = 0 To GrhLayerList.ListCount
            If GrhLayerList.List(tempidx) = temptxt Then
                GrhLayerList.ListIndex = tempidx
                Exit Sub
            End If
        Next tempidx
        
        Exit Sub
    End If

End Sub

Private Sub TriggerFillCmd_Click()
    Dim map_x As Long
    Dim max_x As Long
    Dim map_y As Long
    Dim max_y As Long
    
    If TriggerList.ListIndex < 0 Then
        MsgBox "Must select a trigger to place first.", vbOKOnly
        Exit Sub
    End If
    
    store_action trigger, fill
    
    Engine.Map_Bounds_Get max_x, max_y
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            Engine.Map_Trigger_Set map_x, map_y, TriggerList.ListIndex
        Next map_y
    Next map_x

End Sub

Private Sub TriggersRemoveAllCmd_Click()
    Dim X As Long
    Dim Y As Long
    Dim max_x As Long
    Dim max_y As Long
    
    store_action trigger, Remove_all
    
    Engine.Map_Bounds_Get max_x, max_y
    For X = 1 To max_x
        For Y = 1 To max_y
            Engine.Map_Trigger_Unset X, Y
        Next Y
    Next X
    
    Modified = True

End Sub

Private Sub UnblockAllCmd_Click()
    Dim X As Long
    Dim Y As Long
    Dim max_x As Long
    Dim max_y As Long
    
    store_action blocking, Remove_all
    
    Engine.Map_Bounds_Get max_x, max_y
    
    For X = 1 To max_x
        For Y = 1 To max_y
            Engine.Map_Blocked_Set X, Y, False
        Next Y
    Next X
    
    BlockBordersCmd.Caption = "Block Borders ON"
    
    Modified = True

End Sub

Private Sub WalkModeChk_Click()
    
    Toggle_Walk_Mode
    
    'Enable all buttons, since they can´t be used while in Walk Mode
    GrhPickCmd.Enabled = True
    PickBaseLightColorCmd.Enabled = True
    
    'Disable/Enable Map menu commands
    MapMnu.Enabled = Not Walk_Mode

End Sub

Private Sub LoadToolbar()
   Dim btn As Button
   'Adds all buttons to the toolbar
   'New
   Set btn = Toolbar1.Buttons.Add(, "new", , tbrDefault, "new")
   btn.ToolTipText = "New Map"
   btn.Description = btn.ToolTipText
   'Open
   Set btn = Toolbar1.Buttons.Add(, "open", , tbrDefault, "open")
   btn.ToolTipText = "Open Map"
   btn.Description = btn.ToolTipText
   'Save
   Set btn = Toolbar1.Buttons.Add(, "save", , tbrDefault, "save")
   btn.ToolTipText = "Save Map"
   btn.Description = btn.ToolTipText
   'Separator
   Toolbar1.Buttons.Add , , , tbrSeparator
   'Undo
   Set btn = Toolbar1.Buttons.Add(, "undo", , tbrDefault, "undo")
   btn.ToolTipText = "Undo"
   btn.Description = btn.ToolTipText
   btn.Enabled = False
   'Redo
   Set btn = Toolbar1.Buttons.Add(, "redo", , tbrDefault, "redo")
   btn.ToolTipText = "Redo"
   btn.Description = btn.ToolTipText
   btn.Enabled = False
   'Separator
   Toolbar1.Buttons.Add , , , tbrSeparator
   'Grhs
   Set btn = Toolbar1.Buttons.Add(, "grh", , tbrDefault, "grh")
   btn.ToolTipText = "Graphics"
   btn.Description = btn.ToolTipText
   'Tiles
   Set btn = Toolbar1.Buttons.Add(, "tiles", , tbrDefault, "tiles")
   btn.ToolTipText = "Block/Triggers"
   btn.Description = btn.ToolTipText
   'Lights
   Set btn = Toolbar1.Buttons.Add(, "lights", , tbrDefault, "lights")
   btn.ToolTipText = "Lights"
   btn.Description = btn.ToolTipText
   'Particle Groups
   Set btn = Toolbar1.Buttons.Add(, "particle_groups", , tbrDefault, "particle_groups")
   btn.ToolTipText = "Particle Groups"
   btn.Description = btn.ToolTipText
   'Exits
   Set btn = Toolbar1.Buttons.Add(, "exits", , tbrDefault, "exits")
   btn.ToolTipText = "Exits"
   btn.Description = btn.ToolTipText
   'OBJs
   Set btn = Toolbar1.Buttons.Add(, "OBJs", , tbrDefault, "OBJs")
   btn.ToolTipText = "Objects"
   btn.Description = btn.ToolTipText
   'NPCs
   Set btn = Toolbar1.Buttons.Add(, "NPCs", , tbrDefault, "NPCs")
   btn.ToolTipText = "NPCs"
   btn.Description = btn.ToolTipText
   Toolbar1.Appearance = ccFlat
   Toolbar1.BorderStyle = ccNone
   Toolbar1.style = tbrFlat
End Sub

Private Sub show_tiles_controls()
    'Shows tiles controls
    LightTools.Visible = False
    GrhTools.Visible = False
    TilesTools.Visible = True
    ParticleGroupsTools.Visible = False
    ExitTools.Visible = False
    ObjectsTools.Visible = False
    NPCTools.Visible = False
    
    'Updates tool in the status bar
    StatusBar1.Panels(1).text = "Tool: Tiles"

End Sub

Private Sub show_lights_controls()
    'Shows lights controls
    LightTools.Visible = True
    GrhTools.Visible = False
    TilesTools.Visible = False
    ParticleGroupsTools.Visible = False
    ExitTools.Visible = False
    ObjectsTools.Visible = False
    NPCTools.Visible = False
    
    'Updates tool in the status bar
    StatusBar1.Panels(1).text = "Tool: Lights"

End Sub

Private Sub show_PGs_controls()
    'Shows particle groups controls
    LightTools.Visible = False
    GrhTools.Visible = False
    TilesTools.Visible = False
    ParticleGroupsTools.Visible = True
    ExitTools.Visible = False
    ObjectsTools.Visible = False
    NPCTools.Visible = False
    
    'Updates tool in the status bar
    StatusBar1.Panels(1).text = "Tool: Particle Groups"

End Sub

Private Sub show_grh_controls()
    'Shows grhs controls
    LightTools.Visible = False
    GrhTools.Visible = True
    TilesTools.Visible = False
    ParticleGroupsTools.Visible = False
    ExitTools.Visible = False
    ObjectsTools.Visible = False
    NPCTools.Visible = False
    
    'Updates tool in the status bar
    StatusBar1.Panels(1).text = "Tool: Graphics"

End Sub

Private Sub show_exits_controls()
    'Shows exits controls
    LightTools.Visible = False
    GrhTools.Visible = False
    TilesTools.Visible = False
    ParticleGroupsTools.Visible = False
    ExitTools.Visible = True
    ObjectsTools.Visible = False
    NPCTools.Visible = False
    
    'Updates tool in the status bar
    StatusBar1.Panels(1).text = "Tool: Exits"

End Sub

Private Sub show_OBJs_controls()
    'Shows OBJs controls
    LightTools.Visible = False
    GrhTools.Visible = False
    TilesTools.Visible = False
    ParticleGroupsTools.Visible = False
    ExitTools.Visible = False
    ObjectsTools.Visible = True
    NPCTools.Visible = False
    
    'Updates tool in the status bar
    StatusBar1.Panels(1).text = "Tool: Objects"

End Sub

Private Sub show_NPCs_controls()
    'Shows NPCs controls
    LightTools.Visible = False
    GrhTools.Visible = False
    TilesTools.Visible = False
    ParticleGroupsTools.Visible = False
    ExitTools.Visible = False
    ObjectsTools.Visible = False
    NPCTools.Visible = True
    
    'Updates tool in the status bar
    StatusBar1.Panels(1).text = "Tool: NPCs"

End Sub

Public Sub arrange_light_color(ByVal color As Long)
    Dialog.color = color
    
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    General_Long_Color_to_RGB color, r, g, b
    
    'Invert r and b values to fix bug
    slider(0).value = r
    slider(1).value = g
    slider(2).value = b
End Sub

Private Sub XBorderTxt_Change()
    'MUST allways be an int
    XBorderTxt.text = Int(Val(XBorderTxt.text))
End Sub

Private Sub YBorderTxt_Change()
    'MUST allways be an int
    YBorderTxt.text = Int(Val(YBorderTxt.text))
End Sub

Public Sub InitStatusBars()
    Dim Panel As Panel
    Dim statusbar1_width As Long
    Dim statusbar2_width As Long
    
    statusbar1_width = frmMain.ScaleWidth / 6
    statusbar2_width = frmMain.ScaleWidth / 3
    
    'First panel already exists, just edit it
    Set Panel = StatusBar1.Panels(1)
    Panel.Key = "current_tool"
    Panel.width = statusbar1_width
    'Pos
    Set Panel = StatusBar1.Panels.Add(, "position")
    Panel.text = "Pos:"
    Panel.width = statusbar1_width
    'Blocked
    Set Panel = StatusBar1.Panels.Add(, "blocked")
    Panel.text = "Blocked: FALSE"
    Panel.width = statusbar1_width
    'Base Light
    Set Panel = StatusBar1.Panels.Add(, "base_light")
    Panel.text = "Base Light:"
    Panel.width = statusbar1_width
    'Lights
    Set Panel = StatusBar1.Panels.Add(, "light")
    Panel.text = "Light: NONE"
    Panel.width = statusbar1_width
    'Trigger
    Set Panel = StatusBar1.Panels.Add(, "trigger")
    Panel.text = "Trigger: NONE"
    Panel.width = statusbar1_width
    
    'These are in a separate StatusBar due to size
    'Exits
    Set Panel = StatusBar2.Panels(1)
    Panel.Key = "exits"
    Panel.text = "Exit: NONE"
    Panel.width = statusbar2_width
    'OBJs
    Set Panel = StatusBar2.Panels.Add(, "objects")
    Panel.text = "Item: NONE"
    Panel.width = statusbar2_width
    'NPCs
    Set Panel = StatusBar2.Panels.Add(, "NPCs")
    Panel.text = "NPC: NONE"
    Panel.width = statusbar2_width
End Sub

Public Sub Statusbar_Update()
    'Refresh the data on the Status Bar (except for the tool)
    Dim new_x As Long
    Dim new_y As Long
    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer
    Dim Light As Long
    Dim light_index As Long
    Dim blocked As Boolean
    
    Engine.Input_Mouse_Map_Get new_x, new_y
    
    If Not Engine.Map_In_Bounds(new_x, new_y) Then
        Exit Sub
    End If
    
    'Update pos
    StatusBar1.Panels(2).text = "Pos: " & new_x & ", " & new_y
    
    'Blocked
    blocked = Engine.Map_Blocked_Get(new_x, new_y)
    If blocked = False Then
        StatusBar1.Panels(3).text = "Blocked: FALSE"
    Else
        StatusBar1.Panels(3).text = "Blocked: TRUE"
    End If
    
    'Base Light
    General_Long_Color_to_RGB Engine.Map_Base_Light_Get(new_x, new_y), red, green, blue
    'Invert values
    StatusBar1.Panels(4).text = "Base Light: " & blue & ", " & green & ", " & red
    
    'Light
    light_index = Engine.Map_Light_Get(new_x, new_y)
    If light_index = 0 Then
        StatusBar1.Panels(5).text = "Light: NONE"
    Else
        Engine.Light_Color_Value_Get light_index, Light
        General_Long_Color_to_RGB Light, red, green, blue
        'Invert values
        StatusBar1.Panels(5).text = "Light: " & blue & ", " & green & ", " & red
    End If
    
    'Trigger
    StatusBar1.Panels(6).text = "Trigger: " & General_Var_Get(App.Path & "\Triggers.dat", "TRIG" & Engine.Map_Trigger_Get(new_x, new_y) + 1, "Name")
    
    'Exits
    Dim map_name As String
    Dim x_coord As Long
    Dim y_coord As Long
    Engine.Map_Exit_Get new_x, new_y, map_name, x_coord, y_coord
    If map_name <> "" Then
        StatusBar2.Panels(1).text = "Exit: " & map_name & "   Position: " & x_coord & ", " & y_coord
    Else
        StatusBar2.Panels(1).text = "Exit: NONE"
    End If
    
    'Items
    Dim item_data_index As Long
    Dim item_amount As Long
    Engine.Map_Item_Get new_x, new_y, item_data_index, item_amount
    If item_amount > 0 Then
        StatusBar2.Panels(2).text = "Item: " & General_Var_Get(resource_path & "\scripts\item.ini", "ITEM" & item_data_index, "item_name") & "     Amount: " & item_amount
    Else
        StatusBar2.Panels(2).text = "Item: NONE"
    End If
    
    'NPCs
    Dim npc_data_index As Long
    Engine.Map_NPC_Get new_x, new_y, npc_data_index
    If item_data_index > 0 Then
        StatusBar2.Panels(3).text = "NPC: " & General_Var_Get(resource_path & "\scripts\npc.ini", "NPC" & npc_data_index, "npc_name")
    Else
        StatusBar2.Panels(3).text = "NPC: NONE"
    End If
    
End Sub
