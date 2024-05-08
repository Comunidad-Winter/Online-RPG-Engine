VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progress..."
   ClientHeight    =   1425
   ClientLeft      =   3420
   ClientTop       =   3885
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Compressing..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1560
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

