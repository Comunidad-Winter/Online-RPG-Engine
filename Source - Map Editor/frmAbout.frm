VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About this Map Editor..."
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   4
      Top             =   360
      Width           =   540
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "The Script System was created by Murat Sütunc (FireStarter), and henaced by me."
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "The particle editor included in this Map Editor was coded by Ryan Cain (OneZero) and edited by me."
      Height          =   615
      Left            =   840
      TabIndex        =   7
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 0.7.0"
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dark Sun Online Map Editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   2865
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   870
      X2              =   4200
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label6 
      Caption         =   "Some of the angle code has been taken from Fredrik´s Map Editor."
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Please report any bug or send your suggestions to juansotuyo@hotmail.com"
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   3405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Coded by Juan Martín Sotuyo Dodero (Maraxus)."
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   3465
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKButton_Click()
    Unload Me
    'Place GrhViewer back on top
    If frmMain.GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    'Same with minimap
    If frmMain.MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If
End Sub
