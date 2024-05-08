VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParticleEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dark Sun Online Particle Editor"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12000
   Icon            =   "frmParticleEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.CommandButton cmdSync 
      Caption         =   "Synchronize"
      Height          =   255
      Left            =   240
      TabIndex        =   104
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Commands"
      Height          =   4815
      Left            =   9915
      TabIndex        =   90
      Top             =   195
      Width           =   2040
      Begin VB.CommandButton Command1 
         Caption         =   "&Save Particles"
         Height          =   375
         Left            =   45
         TabIndex        =   103
         Top             =   1875
         Width           =   1935
      End
      Begin VB.CommandButton cmdNewParticle 
         Caption         =   "Create &New Particle"
         Height          =   375
         Left            =   45
         TabIndex        =   102
         Top             =   1050
         Width           =   1935
      End
      Begin VB.CommandButton cmdNewDoubleParticle 
         Caption         =   "New &Double Particle"
         Enabled         =   0   'False
         Height          =   375
         Left            =   45
         TabIndex        =   93
         Top             =   1455
         Width           =   1935
      End
      Begin VB.CommandButton cmdEngineStats 
         Caption         =   "Toggle &Engine Stats"
         Height          =   375
         Left            =   60
         TabIndex        =   92
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdClearParticleGroups 
         Caption         =   "&Clear Particle Groups"
         Height          =   375
         Left            =   45
         TabIndex        =   91
         Top             =   645
         Width           =   1935
      End
   End
   Begin VB.Frame frmfade 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2235
      Left            =   90
      TabIndex        =   94
      Top             =   6570
      Width           =   7680
      Begin VB.TextBox txtfout 
         Height          =   300
         Left            =   1320
         TabIndex        =   98
         Text            =   "0"
         Top             =   405
         Width           =   645
      End
      Begin VB.TextBox txtfin 
         Height          =   285
         Left            =   1320
         TabIndex        =   96
         Text            =   "0"
         Top             =   90
         Width           =   630
      End
      Begin VB.Label Label31 
         Caption         =   "Note: The time a particle remains alive is set in the Duration Tab"
         Height          =   585
         Left            =   90
         TabIndex        =   99
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label30 
         Caption         =   "Fade out time"
         Height          =   300
         Left            =   60
         TabIndex        =   97
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Fade in time"
         Height          =   180
         Left            =   60
         TabIndex        =   95
         Top             =   120
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Particle Speed"
      Height          =   855
      Left            =   195
      TabIndex        =   86
      Top             =   6690
      Width           =   1935
      Begin VB.TextBox speed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   87
         Text            =   "0.5"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Render Delay:"
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame frameColorSettings 
      BorderStyle     =   0  'None
      Caption         =   "Color Tint Settings"
      Height          =   2175
      Left            =   135
      TabIndex        =   74
      Top             =   6555
      Width           =   3975
      Begin VB.HScrollBar RScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   79
         Top             =   1800
         Width           =   3015
      End
      Begin VB.HScrollBar GScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   78
         Top             =   1500
         Width           =   3015
      End
      Begin VB.HScrollBar BScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   77
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ListBox lstColorSets 
         Height          =   840
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1440
         ScaleHeight     =   795
         ScaleWidth      =   2355
         TabIndex        =   76
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   85
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtG 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   83
         Text            =   "0"
         Top             =   1500
         Width           =   375
      End
      Begin VB.TextBox txtB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   81
         Text            =   "0"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R:"
         Height          =   195
         Left            =   120
         TabIndex        =   84
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G:"
         Height          =   195
         Left            =   120
         TabIndex        =   82
         Top             =   1500
         Width           =   165
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   1800
         Width           =   150
      End
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   7995
      ScaleHeight     =   1335
      ScaleWidth      =   3810
      TabIndex        =   14
      Top             =   7515
      Width           =   3810
   End
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "&Save All Streams"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdOpenStreamFile 
      Caption         =   "&Open Stream File"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame frameGrhs 
      Caption         =   "Grh Parameters"
      Height          =   3795
      Left            =   7905
      TabIndex        =   10
      Top             =   5145
      Width           =   4050
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   255
         Left            =   1590
         TabIndex        =   2
         Top             =   780
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   255
         Left            =   1590
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox lstSelGrhs 
         Height          =   1620
         Left            =   2385
         TabIndex        =   4
         Top             =   450
         Width           =   1530
      End
      Begin VB.ListBox lstGrhs 
         Height          =   1620
         Left            =   45
         TabIndex        =   0
         Top             =   450
         Width           =   1500
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   255
         Left            =   1590
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Preview Window"
         Height          =   225
         Left            =   90
         TabIndex        =   89
         Top             =   2145
         Width           =   2115
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Grhs"
         Height          =   195
         Left            =   2370
         TabIndex        =   12
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grh List"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   255
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdNewStream 
      Caption         =   "&New Stream"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstStreamType 
      Height          =   4545
      Left            =   7920
      TabIndex        =   5
      Top             =   450
      Width           =   1935
   End
   Begin VB.PictureBox MainView 
      BackColor       =   &H00000000&
      Height          =   6210
      Left            =   15
      ScaleHeight     =   410
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   516
      TabIndex        =   8
      Top             =   45
      Width           =   7800
      Begin MSComDlg.CommonDialog ComDlg 
         Left            =   2040
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame frmSettings 
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   120
      TabIndex        =   43
      Top             =   6600
      Width           =   6600
      Begin VB.TextBox txtry 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   55
         Text            =   "0"
         Top             =   1635
         Width           =   495
      End
      Begin VB.CheckBox chkresize 
         Caption         =   "Resize"
         Height          =   195
         Left            =   1920
         TabIndex        =   56
         Top             =   1920
         Width           =   1245
      End
      Begin VB.CheckBox chkAlphaBlend 
         Caption         =   "Alpha Blend"
         Height          =   255
         Left            =   3930
         TabIndex        =   60
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox fric 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   59
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox life2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   58
         Text            =   "50"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox life1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   57
         Text            =   "10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox vecy2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   53
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox vecy1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   52
         Text            =   "-50"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox vecx2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   51
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox vecx1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   50
         Text            =   "-10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtAngle 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   49
         Text            =   "0"
         Top             =   1605
         Width           =   495
      End
      Begin VB.TextBox txtY2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   48
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtY1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   47
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtX2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   46
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtX1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   45
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtPCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   44
         Text            =   "20"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtrx 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   54
         Text            =   "0"
         Top             =   1395
         Width           =   495
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize X:"
         Height          =   195
         Left            =   1950
         TabIndex        =   101
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize Y:"
         Height          =   195
         Left            =   1950
         TabIndex        =   100
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   525
         Width           =   240
      End
      Begin VB.Label lblPCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# of Particles:"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Friction:"
         Height          =   195
         Left            =   3915
         TabIndex        =   68
         Top             =   885
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (H):"
         Height          =   195
         Left            =   3915
         TabIndex        =   67
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (L):"
         Height          =   195
         Left            =   3915
         TabIndex        =   66
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y2"
         Height          =   195
         Left            =   1950
         TabIndex        =   65
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   64
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X2:"
         Height          =   195
         Left            =   1950
         TabIndex        =   63
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   62
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angle:"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   1650
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Particle Duration"
      Height          =   855
      Left            =   90
      TabIndex        =   39
      Top             =   6645
      Width           =   1935
      Begin VB.CheckBox chkNeverDies 
         Caption         =   "Never Dies"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox life 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   40
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Life:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   525
         Width           =   300
      End
   End
   Begin VB.Frame frameSpinSettings 
      BorderStyle     =   0  'None
      Caption         =   "Spin Settings"
      Height          =   1095
      Left            =   105
      TabIndex        =   33
      Top             =   6615
      Width           =   1935
      Begin VB.TextBox spin_speedL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "1"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox spin_speedH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   35
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkSpin 
         Caption         =   "Spin"
         Height          =   255
         Left            =   105
         TabIndex        =   34
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (L):"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (H):"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   765
         Width           =   1125
      End
   End
   Begin VB.Frame frameMovement 
      BorderStyle     =   0  'None
      Caption         =   "Movement Settings"
      Height          =   1935
      Left            =   75
      TabIndex        =   22
      Top             =   6615
      Width           =   1935
      Begin VB.TextBox move_x1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   28
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox move_x2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   27
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox move_y1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   26
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox move_y2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   25
         Text            =   "0"
         Top             =   1560
         Width           =   375
      End
      Begin VB.CheckBox chkYMove 
         Caption         =   "Y Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkXMove 
         Caption         =   "X Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1605
         Width           =   1035
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1365
         Width           =   1035
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   765
         Width           =   1035
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   525
         Width           =   1035
      End
   End
   Begin VB.Frame frameGravity 
      BorderStyle     =   0  'None
      Caption         =   "Gravity Settings"
      Height          =   1095
      Left            =   90
      TabIndex        =   16
      Top             =   6630
      Width           =   1935
      Begin VB.TextBox txtGravStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   19
         Text            =   "5"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtBounceStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   18
         Text            =   "1"
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkGravity 
         Caption         =   "Gravity Influence"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravity Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   465
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bounce Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   705
         Width           =   1245
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2670
      Left            =   15
      TabIndex        =   15
      Top             =   6225
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   4710
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Particle Settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravity"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movement "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Spin "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Speed"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Duration "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Color "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fade"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStreamType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stream Types"
      Height          =   195
      Left            =   7965
      TabIndex        =   9
      Top             =   150
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_NewStream 
         Caption         =   "&New Stream"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile_OpenStreamFile 
         Caption         =   "&Open Stream File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile_SaveAll 
         Caption         =   "&Save All Streams"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuView_Toolbox 
         Caption         =   "&Toolbox"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmParticleEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Coded by Ryan Cain (OneZero)
'Edited by Juan Martín Sotuyo Dodero (Maraxus) to add speed and life

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'--> Public "particle_engine-Running" Variable <--
Public go As Boolean
'--> Current Stream File <--
Public CurStreamFile As String
'--> Holds Grh pos for Cursor <--
Public cursor_grh_index As Long
'--> Mouse Coords <--
Public Old_X As Long
Public Old_Y As Long

'Engine
Dim Particle_Engine As New clsTileEngineX

'***** Constants *****
Const MAX_STREAMS = 500

Private Sub chkresize_Click()
If chkresize.value = vbChecked Then
    StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).grh_resize = True
Else
   StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).grh_resize = False
End If
End Sub

Private Sub cmdAdd_Click()
Dim LoopC As Long

If lstGrhs.ListIndex >= 0 Then lstSelGrhs.AddItem lstGrhs.List(lstGrhs.ListIndex)

StreamData(lstStreamType.ListIndex + 1).NumGrhs = lstSelGrhs.ListCount

ReDim StreamData(lstStreamType.ListIndex + 1).grh_list(1 To lstSelGrhs.ListCount)

For LoopC = 1 To StreamData(lstStreamType.ListIndex + 1).NumGrhs
    StreamData(lstStreamType.ListIndex + 1).grh_list(LoopC) = lstSelGrhs.List(LoopC - 1)
Next LoopC

End Sub

Private Sub cmdClear_Click()

lstSelGrhs.Clear

StreamData(lstStreamType.ListIndex + 1).NumGrhs = 0

Erase StreamData(lstStreamType.ListIndex + 1).grh_list

End Sub

Private Sub cmdDelete_Click()
Dim LoopC As Long

If lstSelGrhs.ListIndex >= 0 Then lstSelGrhs.RemoveItem lstSelGrhs.ListIndex

StreamData(lstStreamType.ListIndex + 1).NumGrhs = lstSelGrhs.ListCount

If StreamData(lstStreamType.ListIndex + 1).NumGrhs = 0 Then
    Erase StreamData(lstStreamType.ListIndex + 1).grh_list
Else
    ReDim StreamData(lstStreamType.ListIndex + 1).grh_list(1 To lstSelGrhs.ListCount)
End If

For LoopC = 1 To StreamData(lstStreamType.ListIndex + 1).NumGrhs
    StreamData(lstStreamType.ListIndex + 1).grh_list(LoopC) = lstSelGrhs.List(LoopC - 1)
Next LoopC

End Sub

Private Sub cmdNewParticle_Click()
Call cmdNewStream_Click

End Sub

Private Sub cmdOpenStreamFile_Click()
Dim sFile As String

With ComDlg
    .Filter = "*.ini (Stream Data Files)|*.ini"
    .ShowOpen
    sFile = .FileName
End With

LoadStreamFile sFile
CurStreamFile = sFile

End Sub

Private Sub cmdSync_Click()

Dim llRetVal As Long
If General_File_Exists(App.Path & "\Update\Particles.ini", vbNormal) = True Then
    Dim retval
    retval = MsgBox("An updated particle file already exists, would you like to replace it?", vbYesNo Or vbQuestion)
    If retval = vbYes Then
        'llRetVal = URLDownloadToFile(0, "http://localhost/PEUpdate/Particles.ini", App.Path & "\Update\Particles.ini", 0, 0)
        llRetVal = URLDownloadToFile(0, "ftp://dso:tome@65.163.173.140/Tools/ParticleEditor/Particles.ini", App.Path & "\Particles.ini", 0, 0)
        MsgBox "Update complete!", vbExclamation
    ElseIf retval = vbNo Then
        MsgBox "Update aborted!", vbExclamation
    End If
ElseIf General_File_Exists(App.Path & "\Update\Particles.ini", vbNormal) = False Then
    'llRetVal = URLDownloadToFile(0, "http://localhost/PEUpdate/Particles.ini", App.Path & "\Update\Particles.ini", 0, 0)
    llRetVal = URLDownloadToFile(0, "ftp://dso:tome@65.163.173.140/Tools/ParticleEditor/Particles.ini", App.Path & "\Particles.ini", 0, 0)
    MsgBox "Update complete!", vbExclamation
End If

End Sub

Private Sub Command1_Click()
Call cmdSaveAll_Click
End Sub

Private Sub lstGrhs_DblClick()

Call cmdAdd_Click

End Sub

Private Sub lstGrhs_Click()
Dim GrhInfo

Dim filepath As String
Dim src_x As Long
Dim src_y As Long
Dim src_width As Long
Dim src_height As Long
Dim framecount As Long

GrhInfo = Particle_Engine.Grh_Info_Get(lstGrhs.List(lstGrhs.ListIndex), filepath, src_x, src_y, src_width, src_height, framecount)

If framecount <= 0 Then Exit Sub

picPreview.Cls
Particle_Engine.Grh_Render_To_Hdc lstGrhs.List(lstGrhs.ListIndex), picPreview.hdc, 2, 2

End Sub

Private Sub lstSelGrhs_Click()
Dim GrhInfo

Dim filepath As String
Dim src_x As Long
Dim src_y As Long
Dim src_width As Long
Dim src_height As Long
Dim framecount As Long

GrhInfo = Particle_Engine.Grh_Info_Get(lstSelGrhs.List(lstSelGrhs.ListIndex), filepath, src_x, src_y, src_width, src_height, framecount)

If framecount <= 0 Then Exit Sub

picPreview.Cls
Particle_Engine.Grh_Render_To_Hdc lstSelGrhs.List(lstSelGrhs.ListIndex), picPreview.hdc, 2, 2

End Sub

Private Sub lstSelGrhs_DblClick()

Call cmdDelete_Click

End Sub

Private Sub lstStreamType_Click()
Dim LoopC As Long
Dim DataTemp As Boolean
DataTemp = DataChanged

'Set the values
txtPCount.text = StreamData(lstStreamType.ListIndex + 1).NumOfParticles
txtX1.text = StreamData(lstStreamType.ListIndex + 1).x1
txtY1.text = StreamData(lstStreamType.ListIndex + 1).y1
txtX2.text = StreamData(lstStreamType.ListIndex + 1).x2
txtY2.text = StreamData(lstStreamType.ListIndex + 1).y2
txtAngle.text = StreamData(lstStreamType.ListIndex + 1).angle
vecx1.text = StreamData(lstStreamType.ListIndex + 1).vecx1
vecx2.text = StreamData(lstStreamType.ListIndex + 1).vecx2
vecy1.text = StreamData(lstStreamType.ListIndex + 1).vecy1
vecy2.text = StreamData(lstStreamType.ListIndex + 1).vecy2
life1.text = StreamData(lstStreamType.ListIndex + 1).life1
life2.text = StreamData(lstStreamType.ListIndex + 1).life2
fric.text = StreamData(lstStreamType.ListIndex + 1).friction
chkSpin.value = StreamData(lstStreamType.ListIndex + 1).spin
spin_speedL.text = StreamData(lstStreamType.ListIndex + 1).spin_speedL
spin_speedH.text = StreamData(lstStreamType.ListIndex + 1).spin_speedH
txtGravStrength.text = StreamData(lstStreamType.ListIndex + 1).grav_strength
txtBounceStrength.text = StreamData(lstStreamType.ListIndex + 1).bounce_strength
chkAlphaBlend.value = StreamData(lstStreamType.ListIndex + 1).AlphaBlend
chkGravity.value = StreamData(lstStreamType.ListIndex + 1).gravity
txtrx.text = StreamData(lstStreamType.ListIndex + 1).grh_resizex
txtry.text = StreamData(lstStreamType.ListIndex + 1).grh_resizey
chkXMove.value = StreamData(lstStreamType.ListIndex + 1).XMove
chkYMove.value = StreamData(lstStreamType.ListIndex + 1).YMove
move_x1.text = StreamData(lstStreamType.ListIndex + 1).move_x1
move_x2.text = StreamData(lstStreamType.ListIndex + 1).move_x2
move_y1.text = StreamData(lstStreamType.ListIndex + 1).move_y1
move_y2.text = StreamData(lstStreamType.ListIndex + 1).move_y2

If StreamData(lstStreamType.ListIndex + 1).grh_resize = True Then
    chkresize = vbChecked
Else
    chkresize = vbUnchecked
End If

If StreamData(lstStreamType.ListIndex + 1).life_counter = -1 Then
    life.Enabled = False
    chkNeverDies.value = vbChecked
Else
    life.Enabled = True
    life.text = StreamData(lstStreamType.ListIndex + 1).life_counter
    chkNeverDies.value = vbUnchecked
End If

speed.text = StreamData(lstStreamType.ListIndex + 1).speed

lstSelGrhs.Clear

For LoopC = 1 To StreamData(lstStreamType.ListIndex + 1).NumGrhs
    lstSelGrhs.AddItem StreamData(lstStreamType.ListIndex + 1).grh_list(LoopC)
Next LoopC

DataChanged = DataTemp
If DataChanged = True Then
    Caption = "Dark Sun Online Particle Editor* "
Else
    Caption = "Dark Sun Online Particle Editor"
End If

End Sub

Private Sub lstStreamType_KeyUp(KeyCode As Integer, Shift As Integer)

Dim LoopC As Long
Dim DataTemp As Boolean
DataTemp = DataChanged

'Set the values
txtPCount.text = StreamData(lstStreamType.ListIndex + 1).NumOfParticles
txtX1.text = StreamData(lstStreamType.ListIndex + 1).x1
txtY1.text = StreamData(lstStreamType.ListIndex + 1).y1
txtX2.text = StreamData(lstStreamType.ListIndex + 1).x2
txtY2.text = StreamData(lstStreamType.ListIndex + 1).y2
txtAngle.text = StreamData(lstStreamType.ListIndex + 1).angle
vecx1.text = StreamData(lstStreamType.ListIndex + 1).vecx1
vecx2.text = StreamData(lstStreamType.ListIndex + 1).vecx2
vecy1.text = StreamData(lstStreamType.ListIndex + 1).vecy1
vecy2.text = StreamData(lstStreamType.ListIndex + 1).vecy2
life1.text = StreamData(lstStreamType.ListIndex + 1).life1
life2.text = StreamData(lstStreamType.ListIndex + 1).life2
fric.text = StreamData(lstStreamType.ListIndex + 1).friction
chkSpin.value = StreamData(lstStreamType + 1).spin
spin_speedL.text = StreamData(lstStreamType.ListIndex + 1).spin_speedL
spin_speedH.text = StreamData(lstStreamType.ListIndex + 1).spin_speedH
txtGravStrength.text = StreamData(lstStreamType.ListIndex + 1).grav_strength
txtBounceStrength.text = StreamData(lstStreamType.ListIndex + 1).bounce_strength

chkAlphaBlend.value = StreamData(lstStreamType.ListIndex + 1).AlphaBlend
chkGravity.value = StreamData(lstStreamType.ListIndex + 1).gravity

chkXMove.value = StreamData(lstStreamType.ListIndex + 1).XMove
chkYMove.value = StreamData(lstStreamType.ListIndex + 1).YMove
move_x1.text = StreamData(lstStreamType.ListIndex + 1).move_x1
move_x2.text = StreamData(lstStreamType.ListIndex + 1).move_x2
move_y1.text = StreamData(lstStreamType.ListIndex + 1).move_y1
move_y2.text = StreamData(lstStreamType.ListIndex + 1).move_y2

lstSelGrhs.Clear

For LoopC = 1 To StreamData(lstStreamType.ListIndex + 1).NumGrhs
    lstSelGrhs.AddItem StreamData(lstStreamType.ListIndex + 1).grh_list(LoopC)
Next LoopC

DataChanged = DataTemp
If DataChanged = True Then
    Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
Else
    Caption = "Dark Sun Online Particle Editor - Created by Onezero"
End If

End Sub

Private Sub cmdClearParticleGroups_Click()

Particle_Engine.Particle_Group_Remove_All

End Sub

Private Sub cmdengineStats_Click()

Particle_Engine.Engine_Stats_Show_Toggle

End Sub

Private Sub cmdNewStream_Click()
Dim Name As String
Dim NewStreamNumber As Integer

'Get name for new stream
Name = InputBox("Please enter a Stream Name", "New Stream")

If Name = "" Then Exit Sub

'Set new stream #
NewStreamNumber = lstStreamType.ListCount + 1

'Add stream to combo box
lstStreamType.AddItem NewStreamNumber & " - " & Name

'Add 1 to TotalStreams
TotalStreams = TotalStreams + 1

'Add stream data to StreamData array
StreamData(NewStreamNumber).Name = Name
StreamData(NewStreamNumber).NumOfParticles = 20
StreamData(NewStreamNumber).x1 = 0
StreamData(NewStreamNumber).y1 = 0
StreamData(NewStreamNumber).x2 = 0
StreamData(NewStreamNumber).y2 = 0
StreamData(NewStreamNumber).angle = 0
StreamData(NewStreamNumber).vecx1 = -20
StreamData(NewStreamNumber).vecx2 = 20
StreamData(NewStreamNumber).vecy1 = -20
StreamData(NewStreamNumber).vecy2 = 20
StreamData(NewStreamNumber).life1 = 10
StreamData(NewStreamNumber).life2 = 50
StreamData(NewStreamNumber).friction = 8
StreamData(NewStreamNumber).spin_speedL = 0.1
StreamData(NewStreamNumber).spin_speedH = 0.1
StreamData(NewStreamNumber).grav_strength = 2
StreamData(NewStreamNumber).bounce_strength = -5

StreamData(NewStreamNumber).AlphaBlend = 1
StreamData(NewStreamNumber).gravity = 0

'Select the new stream type in the combo box
lstStreamType.ListIndex = NewStreamNumber - 1

End Sub

Private Sub cmdSaveAll_Click()
Dim LoopC As Long
Dim StreamFile As String
Dim Bypass As Boolean
Dim retval

If General_File_Exists(CurStreamFile, vbNormal) = True Then
    retval = MsgBox("The file " & CurStreamFile & " already exists!" & vbCrLf & "Would you like to overwrite it?", vbYesNoCancel Or vbQuestion)
    If retval = vbNo Then
        Bypass = False
    ElseIf retval = vbCancel Then
        Exit Sub
    ElseIf retval = vbYes Then
        StreamFile = CurStreamFile
        Bypass = True
    End If
End If

If Bypass = False Then
    With ComDlg
        .Filter = "*.ini (Stream Data Files)|*.ini"
        .ShowSave
        StreamFile = .FileName
    End With
    
    If General_File_Exists(StreamFile, vbNormal) = True Then
        retval = MsgBox("The file " & StreamFile & " already exists!" & vbCrLf & "Would you like to overwrite it?", vbYesNo Or vbQuestion)
        If retval = vbNo Then
            Exit Sub
        End If
    End If
End If

Dim GrhListing As String
Dim i As Long

'Check for existing data file and kill it
If General_File_Exists(StreamFile, vbNormal) Then Kill StreamFile

'Write particle data to Particles.ini
General_Var_Write StreamFile, "INIT", "Total", Val(TotalStreams)

For LoopC = 1 To TotalStreams
    General_Var_Write StreamFile, Val(LoopC), "Name", StreamData(LoopC).Name
    General_Var_Write StreamFile, Val(LoopC), "NumOfParticles", Val(StreamData(LoopC).NumOfParticles)
    General_Var_Write StreamFile, Val(LoopC), "X1", Val(StreamData(LoopC).x1)
    General_Var_Write StreamFile, Val(LoopC), "Y1", Val(StreamData(LoopC).y1)
    General_Var_Write StreamFile, Val(LoopC), "X2", Val(StreamData(LoopC).x2)
    General_Var_Write StreamFile, Val(LoopC), "Y2", Val(StreamData(LoopC).y2)
    General_Var_Write StreamFile, Val(LoopC), "Angle", Val(StreamData(LoopC).angle)
    General_Var_Write StreamFile, Val(LoopC), "VecX1", Val(StreamData(LoopC).vecx1)
    General_Var_Write StreamFile, Val(LoopC), "VecX2", Val(StreamData(LoopC).vecx2)
    General_Var_Write StreamFile, Val(LoopC), "VecY1", Val(StreamData(LoopC).vecy1)
    General_Var_Write StreamFile, Val(LoopC), "VecY2", Val(StreamData(LoopC).vecy2)
    General_Var_Write StreamFile, Val(LoopC), "Life1", Val(StreamData(LoopC).life1)
    General_Var_Write StreamFile, Val(LoopC), "Life2", Val(StreamData(LoopC).life2)
    General_Var_Write StreamFile, Val(LoopC), "Friction", Val(StreamData(LoopC).friction)
    General_Var_Write StreamFile, Val(LoopC), "Spin", Val(StreamData(LoopC).spin)
    General_Var_Write StreamFile, Val(LoopC), "Spin_SpeedL", Val(StreamData(LoopC).spin_speedL)
    General_Var_Write StreamFile, Val(LoopC), "Spin_SpeedH", Val(StreamData(LoopC).spin_speedH)
    General_Var_Write StreamFile, Val(LoopC), "Grav_Strength", Val(StreamData(LoopC).grav_strength)
    General_Var_Write StreamFile, Val(LoopC), "Bounce_Strength", Val(StreamData(LoopC).bounce_strength)
    
    General_Var_Write StreamFile, Val(LoopC), "AlphaBlend", Val(StreamData(LoopC).AlphaBlend)
    General_Var_Write StreamFile, Val(LoopC), "Gravity", Val(StreamData(LoopC).gravity)
    
    General_Var_Write StreamFile, Val(LoopC), "XMove", Val(StreamData(LoopC).XMove)
    General_Var_Write StreamFile, Val(LoopC), "YMove", Val(StreamData(LoopC).YMove)
    General_Var_Write StreamFile, Val(LoopC), "move_x1", Val(StreamData(LoopC).move_x1)
    General_Var_Write StreamFile, Val(LoopC), "move_x2", Val(StreamData(LoopC).move_x2)
    General_Var_Write StreamFile, Val(LoopC), "move_y1", Val(StreamData(LoopC).move_y1)
    General_Var_Write StreamFile, Val(LoopC), "move_y2", Val(StreamData(LoopC).move_y2)
    General_Var_Write StreamFile, Val(LoopC), "life_counter", Val(StreamData(LoopC).life_counter)
    General_Var_Write StreamFile, Val(LoopC), "Speed", Str(StreamData(LoopC).speed)
    
    General_Var_Write StreamFile, Val(LoopC), "resize", CInt(StreamData(LoopC).grh_resize)
    General_Var_Write StreamFile, Val(LoopC), "rx", StreamData(LoopC).grh_resizex
    General_Var_Write StreamFile, Val(LoopC), "ry", StreamData(LoopC).grh_resizey
    
    General_Var_Write StreamFile, Val(LoopC), "NumGrhs", Val(StreamData(LoopC).NumGrhs)
    
    GrhListing = vbNullString
    For i = 1 To StreamData(LoopC).NumGrhs
        GrhListing = GrhListing & StreamData(LoopC).grh_list(i) & ","
    Next i
    
    General_Var_Write StreamFile, Val(LoopC), "Grh_List", GrhListing
    
    General_Var_Write StreamFile, Val(LoopC), "ColorSet1", StreamData(LoopC).colortint(0).r & "," & StreamData(LoopC).colortint(0).g & "," & StreamData(LoopC).colortint(0).b
    General_Var_Write StreamFile, Val(LoopC), "ColorSet2", StreamData(LoopC).colortint(1).r & "," & StreamData(LoopC).colortint(1).g & "," & StreamData(LoopC).colortint(1).b
    General_Var_Write StreamFile, Val(LoopC), "ColorSet3", StreamData(LoopC).colortint(2).r & "," & StreamData(LoopC).colortint(2).g & "," & StreamData(LoopC).colortint(2).b
    General_Var_Write StreamFile, Val(LoopC), "ColorSet4", StreamData(LoopC).colortint(3).r & "," & StreamData(LoopC).colortint(3).g & "," & StreamData(LoopC).colortint(3).b
    
Next LoopC

'Report the results
If TotalStreams > 1 Then
    MsgBox TotalStreams & " particle stream types saved to: " & vbCrLf & StreamFile, vbInformation
Else
    MsgBox TotalStreams & " particle stream type saved to: " & vbCrLf & StreamFile, vbInformation
End If

'Set DataChanged variable to false
DataChanged = False
Caption = "Dark Sun Online Particle Editor - Created by Onezero"
CurStreamFile = StreamFile

End Sub

Private Sub BScroll_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"

StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).b = BScroll.value
txtB.text = BScroll.value

picColor.BackColor = RGB(txtB.text, txtG.text, txtR.text)

End Sub

Private Sub chkNeverDies_Click()

DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
If chkNeverDies.value = vbChecked Then
    life.Enabled = False
    StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).life_counter = -1
Else
    life.Enabled = True
    StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).life_counter = life.text
End If
End Sub

Private Sub chkSpin_Click()

DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).spin = chkSpin.value

If chkSpin.value = vbChecked Then
    spin_speedL.Enabled = True
    spin_speedH.Enabled = True
Else
    spin_speedL.Enabled = False
    spin_speedH.Enabled = False
End If

End Sub



Private Sub GScroll_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"

StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).g = GScroll.value
txtG.text = GScroll.value

picColor.BackColor = RGB(txtB.text, txtG.text, txtR.text)

End Sub

Private Sub life_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).life_counter = life.text
End Sub

Private Sub life_GotFocus()

life.SelStart = 0
life.SelLength = Len(life.text)

End Sub

Private Sub lstColorSets_Click()

Dim DataTemp As Boolean
DataTemp = DataChanged

RScroll.value = StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).r
GScroll.value = StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).g
BScroll.value = StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).b

DataChanged = DataTemp
If DataChanged = True Then
    frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
Else
    frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero"
End If

End Sub

Private Sub RScroll_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"

StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).r = RScroll.value
txtR.text = RScroll.value

picColor.BackColor = RGB(txtB.text, txtG.text, txtR.text)

End Sub

Private Sub speed_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
'Arrange decimal separator
Dim temp As String
temp = General_Field_Read(1, speed.text, 44)
If Not temp = "" Then
    speed.text = temp & "." & Right(speed.text, Len(speed.text) - Len(temp) - 1)
    speed.SelStart = Len(speed.text)
    speed.SelLength = 0
End If
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).speed = Val(speed.text)
End Sub

Private Sub speed_GotFocus()

speed.SelStart = 0
speed.SelLength = Len(speed.text)

End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.index
Case 1:
    frmSettings.Visible = True
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 2:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = True
    frmfade.Visible = False
Case 3:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = True
    frameGravity.Visible = False
    frmfade.Visible = False
Case 4:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = True
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 5:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = True
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 6:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = True
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 7:
    frmSettings.Visible = False
    frameColorSettings.Visible = True
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 8:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = True
End Select
End Sub


Private Sub txtrx_Change()
On Error Resume Next
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).grh_resizex = txtrx.text
End Sub

Private Sub txtry_Change()
On Error Resume Next
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).grh_resizey = txtry.text
End Sub

Private Sub vecx1_GotFocus()

vecx1.SelStart = 0
vecx1.SelLength = Len(vecx1.text)

End Sub

Private Sub vecx1_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).vecx1 = vecx1.text
End Sub

Private Sub vecx2_GotFocus()

vecx2.SelStart = 0
vecx2.SelLength = Len(vecx2.text)

End Sub

Private Sub vecx2_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).vecx2 = vecx2.text
End Sub

Private Sub vecy1_GotFocus()

vecy1.SelStart = 0
vecy1.SelLength = Len(vecy1.text)

End Sub

Private Sub vecy1_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).vecy1 = vecy1.text
End Sub

Private Sub vecy2_GotFocus()

vecy2.SelStart = 0
vecy2.SelLength = Len(vecy2.text)

End Sub

Private Sub vecy2_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).vecy2 = vecy2.text
End Sub

Private Sub life1_GotFocus()

life1.SelStart = 0
life1.SelLength = Len(life1.text)

End Sub

Private Sub life1_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).life1 = life1.text
End Sub

Private Sub life2_GotFocus()

life2.SelStart = 0
life2.SelLength = Len(life2.text)

End Sub

Private Sub life2_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).life2 = life2.text
End Sub

Private Sub fric_GotFocus()

fric.SelStart = 0
fric.SelLength = Len(fric.text)

End Sub

Private Sub fric_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).friction = fric.text
End Sub

Private Sub spin_speedL_GotFocus()

spin_speedL.SelStart = 0
spin_speedL.SelLength = Len(spin_speedH.text)

End Sub

Private Sub spin_speedL_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).spin_speedL = spin_speedL.text
End Sub

Private Sub spin_speedH_GotFocus()

spin_speedH.SelStart = 0
spin_speedH.SelLength = Len(spin_speedH.text)

End Sub

Private Sub spin_speedH_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).spin_speedH = spin_speedH.text
End Sub

Private Sub txtPCount_GotFocus()

txtPCount.SelStart = 0
txtPCount.SelLength = Len(txtPCount.text)

End Sub

Private Sub txtPCount_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).NumOfParticles = txtPCount.text
End Sub

Private Sub txtX1_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).x1 = txtX1.text
End Sub

Private Sub txtX1_GotFocus()

txtX1.SelStart = 0
txtX1.SelLength = Len(txtX1.text)

End Sub

Private Sub txtY1_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).y1 = txtY1.text
End Sub

Private Sub txtY1_GotFocus()

txtY1.SelStart = 0
txtY1.SelLength = Len(txtY1.text)

End Sub

Private Sub txtX2_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).x2 = txtX2.text
End Sub

Private Sub txtX2_GotFocus()

txtX2.SelStart = 0
txtX2.SelLength = Len(txtX2.text)

End Sub

Private Sub txtY2_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).y2 = txtY2.text
End Sub

Private Sub txtY2_GotFocus()

txtY2.SelStart = 0
txtY2.SelLength = Len(txtY2.text)

End Sub

Private Sub txtAngle_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).angle = txtAngle.text
End Sub

Private Sub txtAngle_GotFocus()

txtAngle.SelStart = 0
txtAngle.SelLength = Len(txtAngle.text)

End Sub

Private Sub txtGravStrength_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).grav_strength = txtGravStrength.text
End Sub

Private Sub txtGravStrength_GotFocus()

txtGravStrength.SelStart = 0
txtGravStrength.SelLength = Len(txtGravStrength.text)

End Sub

Private Sub txtBounceStrength_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).bounce_strength = txtBounceStrength.text
End Sub

Private Sub txtBounceStrength_GotFocus()

txtBounceStrength.SelStart = 0
txtBounceStrength.SelLength = Len(txtBounceStrength.text)

End Sub

Private Sub move_x1_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).move_x1 = move_x1.text
End Sub

Private Sub move_x1_GotFocus()

move_x1.SelStart = 0
move_x1.SelLength = Len(move_x1.text)

End Sub

Private Sub move_x2_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).move_x2 = move_x2.text
End Sub

Private Sub move_x2_GotFocus()

move_x2.SelStart = 0
move_x2.SelLength = Len(move_x2.text)

End Sub

Private Sub move_y1_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).move_y1 = move_y1.text
End Sub

Private Sub move_y1_GotFocus()

move_y1.SelStart = 0
move_y1.SelLength = Len(move_y1.text)

End Sub

Private Sub move_y2_Change()
On Error Resume Next
DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).move_y2 = move_y2.text
End Sub

Private Sub move_y2_GotFocus()

move_y2.SelStart = 0
move_y2.SelLength = Len(move_y2.text)

End Sub


Private Sub chkAlphaBlend_Click()

DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).AlphaBlend = chkAlphaBlend.value
End Sub

Private Sub chkGravity_Click()

DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).gravity = chkGravity.value

If chkGravity.value = vbChecked Then
    txtGravStrength.Enabled = True
    txtBounceStrength.Enabled = True
Else
    txtGravStrength.Enabled = False
    txtBounceStrength.Enabled = False
End If

End Sub

Private Sub chkXMove_Click()

DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).XMove = chkXMove.value

If chkXMove.value = vbChecked Then
    move_x1.Enabled = True
    move_x2.Enabled = True
Else
    move_x1.Enabled = False
    move_x2.Enabled = False
End If

End Sub

Private Sub chkYMove_Click()

DataChanged = True
frmParticleEditor.Caption = "Dark Sun Online Particle Editor - Created by Onezero*"
StreamData(frmParticleEditor.lstStreamType.ListIndex + 1).YMove = chkYMove.value

If chkYMove.value = vbChecked Then
    move_y1.Enabled = True
    move_y2.Enabled = True
Else
    move_y1.Enabled = False
    move_y2.Enabled = False
End If

End Sub
Private Sub Form_Load()
'On Error Resume Next
'*****************************************************************
'Author: Ryan Cain (Onezero)
'Last Modify Date: 4/10/2003
'Form Load
'*****************************************************************
    '**************************************************************
    'particle_engine Initialization
    '**************************************************************
    'Run windowed
    lstColorSets.AddItem "Bottom Left"
    lstColorSets.AddItem "Top Left"
    lstColorSets.AddItem "Bottom Right"
    lstColorSets.AddItem "Top Right"
    frmSettings.Visible = True
    frmfade.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    go = Particle_Engine.Engine_Initialize(frmParticleEditor.hwnd, MainView.hwnd, True, resource_path, , , , , MainView.width / 32, MainView.height / 32, 32)
    
    '****************************
    'set some particle_engine parameters
    '****************************
    Particle_Engine.Engine_Stats_Show_Toggle
    Particle_Engine.Engine_Base_Speed_Set 0.03 'Speed that the particle_engine should appear to run at (0.03 = roughly 30fps)
    DoEvents
    'Show forms
    If General_File_Exists(resource_path & "\Particles.ini", vbNormal) = True Then
        LoadStreamFile resource_path & "\Particles.ini"
        CurStreamFile = resource_path & "\Particles.ini"
    Else
        MsgBox "Default stream data file not found, please choose a file to open.", vbInformation
        cmdOpenStreamFile_Click
    End If
    'Particle_Engine.Fill_Grh_List_Particles_Only frmParticleEditor.lstGrhs
    DoEvents
     
    Me.Show
    'frmToolbox.Show vbModeless, Me
    
    '****************************
    'load stream file
    '****************************

    
    Particle_Engine.Fill_Grh_List frmParticleEditor.lstGrhs
    
    '****************************
    'Load Map
    '****************************
    Particle_Engine.Map_Create 25, 25
    'Create a map with any grh and remove them, so it´s xompletely black
    Particle_Engine.Map_Fill 1, 1
    Dim map_x As Long
    Dim map_y As Long
    For map_x = 1 To 25
        For map_y = 1 To 25
            Particle_Engine.Map_Grh_UnSet map_x, map_y, 1
        Next map_y
    Next map_x
    Particle_Engine.Map_Base_Light_Fill RGB(255, 255, 255)
    
    '****************************
    'set view pos
    '****************************
    Particle_Engine.Engine_View_Pos_Set 11, 11
    
    '****************************
    'show cursor pos
    '****************************
    cursor_grh_index = Particle_Engine.Map_Grh_Set(1, 1, 1, 2, False)
    
    'Set DataChanged variable to false
    DataChanged = False
    Me.Caption = "Dark Sun Online Particle Editor"
    
    '**************************************************************
    'Main Loop
    '**************************************************************
MainLoop:
    Do While go
       
        '****************************
        'Render next frame
        '****************************
        'Only run if the form is not minimized
        If Me.WindowState <> vbMinimized Then
            go = Particle_Engine.Engine_Render_Start
            go = Particle_Engine.Engine_Render_End
            
            '****************************
            'Handle Inputs
            '****************************
            Check_Keys
        End If
        
        '****************************
        'Do widnow's events
        '****************************
        DoEvents
    Loop
    
    If DataChanged Then
        Dim save As VbMsgBoxResult
        save = MsgBox("Particle Streams have been modified since the last time you saved." & vbCrLf & "Would you like to save now?", vbQuestion Or vbYesNoCancel)
        
        Select Case save
            Case vbYes
                cmdSaveAll_Click
                
            Case vbCancel
                go = True
                GoTo MainLoop
                
        End Select
    End If
    
    'Reload the particle data
    Load_Particle_Streams_To_ComboBox frmMain.ParticleType
    
    '**************************************************************
    'Clean Up
    '**************************************************************
    Particle_Engine.Engine_DeInitialize
    Set Particle_Engine = Nothing
    Particle_Editor_Unloaded = True
    'Unload frmToolbox
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*****************************************************************
'Form Unload
'*****************************************************************
    go = False
    DoEvents
    Cancel = 0
End Sub

Sub Check_Mouse(OldX As Long, OldY As Long)
'*****************************************************************
'Checks Mouse
'*****************************************************************
    Dim LoopC As Long
    'Make sure the mouse is in the view area
    If Particle_Engine.Input_Mouse_In_View Then
    
        'Check the mouse for movement
        If Particle_Engine.Input_Mouse_Moved_Get Then
            Dim temp_x As Long
            Dim temp_y As Long
            Particle_Engine.Input_Mouse_Map_Get temp_x, temp_y
            Particle_Engine.Map_Grh_UnSet OldX, OldY, 2
            Particle_Engine.Map_Grh_Set temp_x, temp_y, 1, 2, True
            Old_X = temp_x
            Old_Y = temp_y
        End If
        
        'Check left button
        If Particle_Engine.Input_Mouse_Button_Left_Get Then
            Debug.Print life1.text
            If fric.text < 1 Then Exit Sub
            
            If lstStreamType.ListIndex < 0 Then Exit Sub
            'Get mouse position
            Particle_Engine.Input_Mouse_Map_Get temp_x, temp_y

            Dim NumSelGrh As Integer
            NumSelGrh = lstSelGrhs.ListCount
            Dim temp_list() As Long
            ReDim temp_list(1 To NumSelGrh)
            For LoopC = 1 To NumSelGrh
                temp_list(LoopC) = lstSelGrhs.List(LoopC - 1)
                Debug.Print lstSelGrhs.List(LoopC - 1)
            Next LoopC

            'Create particle group
            Dim GravOn As Boolean
            If chkGravity.value = 1 Then
                GravOn = True
            Else
                GravOn = False
            End If
            
            Dim XMoveOn As Boolean
            If chkXMove.value = 1 Then
                XMoveOn = True
            Else
                XMoveOn = False
            End If
            
            Dim YMoveOn As Boolean
            If chkYMove.value = 1 Then
                YMoveOn = True
            Else
                YMoveOn = False
            End If
            
            Dim SpinOn As Boolean
            If chkSpin.value = 1 Then
                SpinOn = True
            Else
                SpinOn = False
            End If
            
            Dim alive_counter As Long
            If chkNeverDies.value = 1 Then
                alive_counter = -1
            Else
                alive_counter = Val(life.text)
            End If
            
            Dim rgb_list(0 To 3) As Long
            rgb_list(0) = RGB(StreamData(lstStreamType.ListIndex + 1).colortint(0).r, StreamData(lstStreamType.ListIndex + 1).colortint(0).g, StreamData(lstStreamType.ListIndex + 1).colortint(0).b)
            rgb_list(1) = RGB(StreamData(lstStreamType.ListIndex + 1).colortint(1).r, StreamData(lstStreamType.ListIndex + 1).colortint(1).g, StreamData(lstStreamType.ListIndex + 1).colortint(1).b)
            rgb_list(2) = RGB(StreamData(lstStreamType.ListIndex + 1).colortint(2).r, StreamData(lstStreamType.ListIndex + 1).colortint(2).g, StreamData(lstStreamType.ListIndex + 1).colortint(2).b)
            rgb_list(3) = RGB(StreamData(lstStreamType.ListIndex + 1).colortint(3).r, StreamData(lstStreamType.ListIndex + 1).colortint(3).g, StreamData(lstStreamType.ListIndex + 1).colortint(3).b)
            
            Particle_Engine.Particle_Group_Create temp_x, temp_y, temp_list, rgb_list(), Val(txtPCount.text), _
            3, chkAlphaBlend.value, alive_counter, Val(speed.text), , Val(txtX1.text), _
            Val(txtY1.text), txtAngle.text, Val(vecx1.text), Val(vecx2.text), _
            Val(vecy1.text), Val(vecy2.text), Val(life1.text), _
            Val(life2.text), Val(fric.text), Val(spin_speedL.text), _
            GravOn, txtGravStrength.text, txtBounceStrength.text, _
            Val(txtX2.text), Val(txtY2.text), XMoveOn, Val(move_x1.text), _
            Val(move_x2.text), Val(move_y1.text), Val(move_y2.text), _
            YMoveOn, Val(spin_speedH.text), SpinOn, CBool(chkresize.value), CLng(txtrx), CLng(txtry)
            Erase temp_list()
        ElseIf Particle_Engine.Input_Mouse_Button_Right_Get Then
            'Erase particle group
            Particle_Engine.Input_Mouse_Map_Get temp_x, temp_y
            Particle_Engine.Particle_Group_Remove Particle_Engine.Map_Particle_Group_Get(temp_x, temp_y)
        End If
    End If
End Sub

Sub Check_Keys()
'*****************************************************************
'Checks keys
'*****************************************************************

    'Escape
    If Particle_Engine.Input_Key_Get(vbKeyEscape) Then
        go = False
    End If

End Sub

Private Sub MainView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      'Check mouse
10    Check_Mouse Old_X, Old_Y
End Sub

Private Sub MainView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Check mouse
Check_Mouse Old_X, Old_Y
End Sub

Private Sub mnuFile_Exit_Click()
go = False
End Sub

Private Sub mnuFile_NewStream_Click()

Call cmdNewStream_Click

End Sub

Private Sub mnuFile_SaveAll_Click()

Call cmdSaveAll_Click

End Sub

Private Sub mnuHelp_About_Click()

frmParticleEditorAbout.Show vbModal, Me

End Sub

Private Sub mnuView_Toolbox_Click()

If mnuView_Toolbox.Checked = True Then
    'mnuView_Toolbox.Checked = False
Else
   ' mnuView_Toolbox.Checked = True
End If

If mnuView_Toolbox.Checked Then
    'frmToolbox.Show vbModeless, Me
Else
    'frmToolbox.Hide
End If

End Sub

Private Sub LoadStreamFile(StreamFile As String)
    Dim LoopC As Long
    
    '****************************
    'load stream types
    '****************************
    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To MAX_STREAMS) As Stream
    
    'clear combo box
    lstStreamType.Clear
    
    Dim i As Long
    Dim GrhListing As String
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).Name = General_Var_Get(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = General_Var_Get(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).x1 = General_Var_Get(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).y1 = General_Var_Get(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).x2 = General_Var_Get(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).y2 = General_Var_Get(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).angle = General_Var_Get(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = General_Var_Get(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = General_Var_Get(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = General_Var_Get(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = General_Var_Get(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = General_Var_Get(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = General_Var_Get(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).friction = General_Var_Get(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).spin = General_Var_Get(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).AlphaBlend = General_Var_Get(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = General_Var_Get(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = General_Var_Get(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = General_Var_Get(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = General_Var_Get(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = General_Var_Get(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = General_Var_Get(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = General_Var_Get(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = General_Var_Get(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = General_Var_Get(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = General_Var_Get(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).speed = Val(General_Var_Get(StreamFile, Val(LoopC), "Speed"))
        StreamData(LoopC).grh_resize = Val(General_Var_Get(StreamFile, Val(LoopC), "resize"))
        StreamData(LoopC).grh_resizex = Val(General_Var_Get(StreamFile, Val(LoopC), "rx"))
        StreamData(LoopC).grh_resizey = Val(General_Var_Get(StreamFile, Val(LoopC), "ry"))
        StreamData(LoopC).NumGrhs = General_Var_Get(StreamFile, Val(LoopC), "NumGrhs")
        
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(LoopC), "Grh_List")
        
        For i = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(i) = General_Field_Read(Str(i), GrhListing, 44)
        Next i
        
        Dim TempSet As String
        Dim ColorSet As Long
        
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).b = General_Field_Read(3, TempSet, 44)
        Next ColorSet
        
        'fill stream type combo box
        lstStreamType.AddItem LoopC & " - " & StreamData(LoopC).Name
    Next LoopC
    
    'set list box index to 1st item
    lstStreamType.ListIndex = 0

End Sub
