VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mini Map"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   446
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6570
      Left            =   0
      ScaleHeight     =   438
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   439
      TabIndex        =   0
      Top             =   60
      Width           =   6585
      Begin VB.Shape shparea 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   30
         Top             =   45
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim X As Long
    Dim Y As Long
    Engine.Map_Bounds_Get X, Y
    frmMap.picmain.width = X
    frmMap.picmain.height = Y
    frmMap.height = (frmMap.picmain.height + 28) * Screen.TwipsPerPixelY
    frmMap.width = (frmMap.picmain.width + 10) * Screen.TwipsPerPixelX
    
    Engine.Engine_Render_Mini_Map_To_hDC picmain.hdc
    
    Engine.Engine_View_Pos_Get X, Y
    shparea.top = Y - 7
    shparea.left = X - 9
    General_Form_On_Top_Set Me, True
    
    frmMain.MiniMapMnuChk.Checked = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.MiniMapMnuChk.Checked = False
End Sub

Private Sub picmain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Get the tile which was clicked and move there
    If frmMain.WalkModeChk.value = 1 Then
        Toggle_Walk_Mode
    End If
    shparea.top = Fix(Y) - 7
    shparea.left = Fix(X) - 9
    Engine.Engine_View_Pos_Set Fix(X), Fix(Y)
    If frmMain.WalkModeChk.value = 1 Then
        Toggle_Walk_Mode
    End If
End Sub
