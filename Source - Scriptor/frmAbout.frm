VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About ORE Scripter"
   ClientHeight    =   1260
   ClientLeft      =   2475
   ClientTop       =   2100
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   84
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   240
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label Label2 
      Caption         =   "Coded by Juan Martín Sotuyo Dodero (Maraxus). Post any comments, suggestion or report bugs to juansotuyo@hotmail.com"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ORE Scripter v 1.0"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

