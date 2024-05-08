VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tool Settings"
   ClientHeight    =   2895
   ClientLeft      =   135
   ClientTop       =   7290
   ClientWidth     =   3870
   ControlBox      =   0   'False
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabItem 
      Height          =   2895
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   3855
      Visible         =   0   'False
      _ExtentX        =   6800
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Item Tool"
      TabPicture(0)   =   "frmSettings.frx":0CCE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label20"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label21"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstItem"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtIAmount"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtIAmount 
         Height          =   285
         Left            =   2040
         TabIndex        =   60
         Top             =   720
         Width           =   1575
      End
      Begin VB.ListBox lstItem 
         Height          =   1815
         Left            =   120
         TabIndex        =   58
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "Item List"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Amount:"
         Height          =   255
         Left            =   2040
         TabIndex        =   59
         Top             =   480
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab tabExit 
      Height          =   2895
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Exit Properties "
      TabPicture(0)   =   "frmSettings.frx":0CEA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Target Area"
         Height          =   2415
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Width           =   2415
         Begin VB.TextBox txtMap 
            Height          =   285
            Left            =   480
            TabIndex        =   65
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtEX 
            Height          =   285
            Left            =   360
            TabIndex        =   64
            Text            =   "15"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtEY 
            Height          =   285
            Left            =   360
            TabIndex        =   63
            Text            =   "15"
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "map"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   ".map"
            Height          =   255
            Left            =   1560
            TabIndex        =   68
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "X:"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label19 
            Caption         =   "Y:"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   1200
            Width           =   255
         End
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   3240
         Picture         =   "frmSettings.frx":0D06
         Top             =   480
         Width           =   480
      End
   End
   Begin TabDlg.SSTab tabNPC 
      Height          =   2895
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   3855
      Visible         =   0   'False
      _ExtentX        =   6800
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "NPC List "
      TabPicture(0)   =   "frmSettings.frx":1948
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstNPC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.ListBox lstNPC 
         Height          =   1815
         ItemData        =   "frmSettings.frx":1964
         Left            =   120
         List            =   "frmSettings.frx":1966
         TabIndex        =   53
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "As you see there are no options in this tool.        If you wan't to edit the list see the NPC.ini file in the scripts directory."
         Height          =   2055
         Left            =   2040
         TabIndex        =   55
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "NPC list:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab tabBlock 
      Height          =   2895
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Block Tool"
      TabPicture(0)   =   "frmSettings.frx":1968
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdBlockEdges"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.CommandButton cmdBlockEdges 
         Caption         =   "Block Edges OFF"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   3240
         Picture         =   "frmSettings.frx":1984
         Top             =   480
         Width           =   480
      End
   End
   Begin TabDlg.SSTab tabParticle 
      Height          =   2895
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "General Particel Settings"
      TabPicture(0)   =   "frmSettings.frx":25C6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image2"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(4)=   "chkAlpha2"
      Tab(0).Control(5)=   "txtNumParticels"
      Tab(0).Control(6)=   "txtParticelId"
      Tab(0).Control(7)=   "cmbStyle"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Particel List"
      TabPicture(1)   =   "frmSettings.frx":25E2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lstParticels"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lstPGrh"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "particelAdd"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdClear"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.ComboBox cmbStyle 
         Height          =   315
         ItemData        =   "frmSettings.frx":25FE
         Left            =   -74880
         List            =   "frmSettings.frx":2614
         TabIndex        =   46
         Text            =   "Star Burst"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtParticelId 
         Height          =   285
         Left            =   -74880
         TabIndex        =   45
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtNumParticels 
         Height          =   285
         Left            =   -74880
         TabIndex        =   44
         Text            =   "15"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox chkAlpha2 
         Caption         =   "Alpha Blending"
         Height          =   255
         Left            =   -74880
         TabIndex        =   43
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   615
         Left            =   1560
         TabIndex        =   40
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton particelAdd 
         Caption         =   "->"
         Height          =   375
         Left            =   1680
         TabIndex        =   39
         Top             =   600
         Width           =   495
      End
      Begin VB.ListBox lstPGrh 
         Height          =   1620
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1455
      End
      Begin VB.ListBox lstParticels 
         Height          =   1620
         ItemData        =   "frmSettings.frx":2650
         Left            =   2280
         List            =   "frmSettings.frx":2652
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Type"
         Height          =   255
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "ID(Optional)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   48
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Number of particels"
         Height          =   255
         Left            =   -74880
         TabIndex        =   47
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -73200
         Picture         =   "frmSettings.frx":2654
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Grh List"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Particel Graphic List"
         Height          =   255
         Left            =   2280
         TabIndex        =   41
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   120
         Picture         =   "frmSettings.frx":3296
         Top             =   2280
         Width           =   480
      End
   End
   Begin TabDlg.SSTab tabShadow 
      Height          =   2895
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Shadow Tool"
      TabPicture(0)   =   "frmSettings.frx":3ED8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SldRCS"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SldGCS"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "SldBCS"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ChkL4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ChkL3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ChkL2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ChkL1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "picColorColor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.PictureBox picColorColor 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   360
         ScaleHeight     =   435
         ScaleWidth      =   1635
         TabIndex        =   30
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox ChkL1 
         Caption         =   "Bottom Left"
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox ChkL2 
         Caption         =   "Top Left"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox ChkL3 
         Caption         =   "Bottom Right"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox ChkL4 
         Caption         =   "Top Right"
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   2160
         Width           =   1455
      End
      Begin MSComctlLib.Slider SldBCS 
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         SelStart        =   190
         Value           =   190
      End
      Begin MSComctlLib.Slider SldGCS 
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         SelStart        =   190
         Value           =   190
      End
      Begin MSComctlLib.Slider SldRCS 
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         SelStart        =   190
         Value           =   190
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   2280
         Picture         =   "frmSettings.frx":3EF4
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label12 
         Caption         =   "Corners:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   35
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "R          G          B"
         Height          =   1095
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   255
      End
   End
   Begin TabDlg.SSTab tabLight 
      Height          =   2895
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Light Tool"
      TabPicture(0)   =   "frmSettings.frx":4B36
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SlidB"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "SlidG"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SlidR"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdLFill"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtLightId"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbRadius"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "picLColor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.PictureBox picLColor 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   360
         ScaleHeight     =   915
         ScaleWidth      =   1515
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox cmbRadius 
         Height          =   315
         ItemData        =   "frmSettings.frx":4B52
         Left            =   2160
         List            =   "frmSettings.frx":4B65
         TabIndex        =   17
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLightId 
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdLFill 
         Caption         =   "Set Map Base Light"
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   2400
         Width           =   1575
      End
      Begin MSComctlLib.Slider SlidR 
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         SelStart        =   255
         Value           =   255
      End
      Begin MSComctlLib.Slider SlidG 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         SelStart        =   255
         Value           =   255
      End
      Begin MSComctlLib.Slider SlidB 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         SelStart        =   255
         Value           =   255
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2040
         Picture         =   "frmSettings.frx":4B78
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label7 
         Caption         =   "R          G        B"
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Radius"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Light ID"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   1200
         Width           =   615
      End
   End
   Begin TabDlg.SSTab tabGrh 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Grh Tool"
      TabPicture(0)   =   "frmSettings.frx":57BA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Slider1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "picRotate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdRotatUp"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdRotateDown"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LstGrh"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkAlpha"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbLayer"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkBlocked"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdFillMap"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.CommandButton cmdFillMap 
         Caption         =   "Fill Map"
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkBlocked 
         Caption         =   "Blocked"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cmbLayer 
         Height          =   315
         ItemData        =   "frmSettings.frx":57D6
         Left            =   600
         List            =   "frmSettings.frx":57E6
         TabIndex        =   8
         Text            =   "1"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox chkAlpha 
         Caption         =   "Use alpha blending"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   1695
      End
      Begin VB.ListBox LstGrh 
         Height          =   1230
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdRotateDown 
         Caption         =   "D"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdRotatUp 
         Caption         =   "U"
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox picRotate 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2040
         ScaleHeight     =   825
         ScaleWidth      =   945
         TabIndex        =   2
         Top             =   720
         Width           =   975
         Begin VB.Line LineRotate 
            BorderColor     =   &H00FFFFFF&
            X1              =   500
            X2              =   500
            Y1              =   360
            Y2              =   0
         End
         Begin VB.Label lblRotation 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   600
            Width           =   495
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
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Max             =   360
         TickFrequency   =   5
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rotation"
         Height          =   1695
         Left            =   1920
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Layer"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Grh number:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1320
         Picture         =   "frmSettings.frx":57F6
         Top             =   1920
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'frmSettings.frm - ORE Map Editor
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'***************************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'Fredrik Alexandersson (fredrik@oraklet.zzn.com) - 5/17/2003
'   -Second official/unofficial release
'
'Aaron Perkins(aaron@baronsoft.com) - 5/12/2003
'   -First offical release
'
'Fredrik Alexandersson (fredrik@oraklet.zzn.com) - 5/12/2003
'   -Last unoffical release
'
'*****************************************************************
Option Explicit

Public angle As Long 'The Grh Tool uses this
Dim x As Long
Dim y As Long


Private Sub cmdClear_Click()
lstParticels.Clear
frmEditor.ParticelCCount = 0
End Sub

Private Sub cmdRotateDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
angle = angle - 5
If angle < 0 Then
   angle = 360
End If
x = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(angle * 3.1415 / 180)
y = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(angle * 3.1415 / 180)
With LineRotate
 .x2 = x
 .y2 = y
End With
lblRotation.Caption = angle
Slider1.value = angle
End Sub

Private Sub cmdRotatUp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
angle = angle + 5
If angle > 361 Then
   angle = 0
End If
x = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(angle * 3.1415 / 180)
y = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(angle * 3.1415 / 180)
With LineRotate
 .x2 = x
 .y2 = y
End With
lblRotation.Caption = angle
Slider1.value = angle
End Sub



Private Sub LstGrh_Click()
If frmEditor.mnuGrhViewer.Checked = True Then
 frmGrhViewer.Cls
 frmEditor.engine.Grh_Render_To_Hdc LstGrh.List(LstGrh.ListIndex), frmGrhViewer.hdc, 1, 1, True
End If
End Sub

Private Sub LstGrh_KeyDown(KeyCode As Integer, Shift As Integer)
frmEditor.engine.Grh_Render_To_Hdc LstGrh.List(LstGrh.ListIndex), frmGrhViewer.hdc, 1, 1, True
End Sub

Private Sub lstParticels_Click()
If frmEditor.mnuGrhViewer.Checked = True Then
 frmGrhViewer.Cls
 frmEditor.engine.Grh_Render_To_Hdc lstParticels.List(lstParticels.ListIndex), frmGrhViewer.hdc, 1, 1, True
End If
End Sub

Private Sub lstPGrh_Click()
If frmEditor.mnuGrhViewer.Checked = True Then
 frmGrhViewer.Cls
 frmEditor.engine.Grh_Render_To_Hdc lstPGrh.List(lstPGrh.ListIndex), frmGrhViewer.hdc, 1, 1, True
End If
End Sub

Private Sub SldBCS_Scroll()
picColorColor.BackColor = RGB(SldRCS, SldGCS, SldBCS)
End Sub

Private Sub SldGCS_Scroll()
picColorColor.BackColor = RGB(SldRCS, SldGCS, SldBCS)
End Sub

Private Sub SldRCS_Scroll()
picColorColor.BackColor = RGB(SldRCS, SldGCS, SldBCS)
End Sub

Private Sub Slider1_Scroll()
angle = Slider1.value
If angle > 361 Then
   angle = 0
End If
x = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(angle * 3.1415 / 180)
y = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(angle * 3.1415 / 180)
With LineRotate
 .x2 = x
 .y2 = y
 .x1 = picRotate.width / 2
 .y1 = picRotate.height / 2
End With
lblRotation.Caption = angle
Slider1.value = angle
End Sub


Private Sub SlidR_Scroll()
 picLColor.BackColor = RGB(SlidR.value, SlidG.value, SlidB.value)
End Sub
Private Sub SlidG_Scroll()
 picLColor.BackColor = RGB(SlidR.value, SlidG.value, SlidB.value)
End Sub
Private Sub SlidB_Scroll()
 picLColor.BackColor = RGB(SlidR.value, SlidG.value, SlidB.value)
End Sub

Private Sub cmdBlockEdges_Click()
frmEditor.MapModified = True
If cmdBlockEdges.Caption = "Block Edges OFF" Then
 frmEditor.engine.Map_Edges_Blocked_Set 10, 7, True
 cmdBlockEdges.Caption = "Block Edges ON"
Else
 frmEditor.engine.Map_Edges_Blocked_Set 10, 7, False
 cmdBlockEdges.Caption = "Block Edges OFF"
End If
End Sub

Private Sub cmdFillMap_Click()
Dim Response As Integer
Response = MsgBox("Warning: Everything you painted on layer " & cmbLayer.text & " will be replaced with grh number" & LstGrh.List(LstGrh.ListIndex) & ". Are you shure?", 4)
   Select Case Response
      Case 6
      If LstGrh.ListIndex <> -1 Then
       frmEditor.engine.Map_Fill LstGrh.List(LstGrh.ListIndex), cmbLayer.text
       frmEditor.MapModified = True
      End If
      Case 7
      Exit Sub
   End Select

End Sub
Private Sub cmdLFill_Click()
frmEditor.MapModified = True
frmEditor.engine.Map_Base_Light_Fill RGB(SlidR, SlidG, SlidB)
frmEditor.MapBaseLight = RGB(SlidR, SlidG, SlidB)
End Sub

Private Sub Form_Load()
x = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(angle * 3.1415 / 180)
y = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(angle * 3.1415 / 180)
With LineRotate
 .x2 = x
 .y2 = y
 .x1 = picRotate.width / 2
 .y1 = picRotate.height / 2
End With
General_Form_On_Top_Set Me, True
End Sub
Private Sub particelAdd_Click()
lstParticels.AddItem lstPGrh.List(lstPGrh.ListIndex)
frmEditor.ParticelCCount = frmEditor.ParticelCCount + 1
End Sub

Private Sub txtEX_Change()
If txtEX.text = "" Then
 txtEX.text = 0
End If
txtEX.text = CLng(txtEX.text)
End Sub

Private Sub txtEY_Change()
If txtEY.text = "" Then
 txtEY.text = 0
End If
txtEY.text = CLng(txtEX.text)
End Sub

Private Sub txtIAmount_Change()
If txtIAmount.text = "" Then
 txtIAmount.text = 0
End If
txtIAmount.text = CLng(txtIAmount.text)
End Sub

Private Sub txtNumParticels_Change()
If txtNumParticels = "" Then
txtNumParticels = 0
End If
txtNumParticels = Int(txtNumParticels)
If Int(txtNumParticels) > 600 Then
MsgBox "Ok map editor user! It's not recomended to use this many particels in one particel system."
MsgBox "You might wan't lower the number of particels."
MsgBox "If you want more i can't stop you. If you do, it's recomended to save now!"
txtNumParticels.text = Val(txtNumParticels.text)
End If
End Sub
