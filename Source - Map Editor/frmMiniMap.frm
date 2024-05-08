VERSION 5.00
Begin VB.Form frmMiniMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generated MiniMap"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmMiniMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMap 
      BackColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   120
      ScaleHeight     =   7395
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmMiniMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
