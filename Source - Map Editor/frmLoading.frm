VERSION 5.00
Begin VB.Form frmLoading 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Loading..."
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   720
      Picture         =   "frmLoading.frx":0000
      ScaleHeight     =   1605
      ScaleWidth      =   1605
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label LoadingLbl 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadingLbl_Change()
    DoEvents
End Sub
