VERSION 5.00
Begin VB.Form frmGrhViewer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Grh Viewer"
   ClientHeight    =   2730
   ClientLeft      =   12030
   ClientTop       =   480
   ClientWidth     =   3135
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   182
   ScaleMode       =   0  'User
   ScaleWidth      =   3147.295
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmGrhViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    General_Form_On_Top_Set Me, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.GrhViewerMnuChk.Checked = False
End Sub
