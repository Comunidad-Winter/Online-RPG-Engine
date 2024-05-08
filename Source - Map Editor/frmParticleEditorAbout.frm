VERSION 5.00
Begin VB.Form frmParticleEditorAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   2250
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5070
   ClipControls    =   0   'False
   Icon            =   "frmParticleEditorAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1552.99
   ScaleMode       =   0  'User
   ScaleWidth      =   4760.992
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmParticleEditorAbout.frx":08CA
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1905
      TabIndex        =   0
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   1014.176
      X2              =   4507.448
      Y1              =   496.957
      Y2              =   496.957
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmParticleEditorAbout.frx":1194
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dark Sun Online Particle Editor"
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
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3210
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   195
      Left            =   1050
      TabIndex        =   4
      Top             =   480
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   1014.176
      X2              =   4507.448
      Y1              =   496.957
      Y2              =   496.957
   End
End
Attribute VB_Name = "frmParticleEditorAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
