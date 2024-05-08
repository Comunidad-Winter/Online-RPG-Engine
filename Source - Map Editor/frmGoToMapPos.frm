VERSION 5.00
Begin VB.Form frmGoToMapPos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Go to map pos..."
   ClientHeight    =   975
   ClientLeft      =   4260
   ClientTop       =   3930
   ClientWidth     =   3000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton GoCmd 
      Caption         =   "Go"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox YPos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   540
      Width           =   615
   End
   Begin VB.TextBox XPos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   60
      Width           =   615
   End
   Begin VB.Label MaxX 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   150
   End
   Begin VB.Label MaxY 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   150
   End
End
Attribute VB_Name = "frmGoToMapPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim max_x As Long
Dim max_y As Long

Private Sub Form_Load()
    Dim X As Long
    Dim Y As Long
    
    Engine.Map_Bounds_Get max_x, max_y
    
    MaxX.Caption = "of " & max_x
    MaxY.Caption = "of " & max_y
    
    Engine.Engine_View_Pos_Get X, Y
    
    XPos.text = X
    YPos.text = Y
End Sub

Private Sub GoCmd_Click()
    If frmMain.WalkModeChk.value = 1 Then
        Toggle_Walk_Mode
    End If
    Engine.Engine_View_Pos_Set Val(XPos.text), Val(YPos.text)
    Unload Me
    If frmMain.WalkModeChk.value = 1 Then
        Toggle_Walk_Mode
    End If
End Sub

Private Sub XPos_Change()
    XPos.text = Int(Val(XPos.text))
    If Val(XPos.text) > max_x Then XPos.text = max_x
    If Val(XPos.text) < 1 Then XPos.text = 1
End Sub

Private Sub XPos_GotFocus()
    XPos.SelStart = 0
    XPos.SelLength = Len(XPos.text)
End Sub

Private Sub YPos_Change()
    YPos.text = Int(Val(YPos.text))
    If Val(YPos.text) > max_y Then YPos.text = max_y
    If Val(YPos.text) < 1 Then YPos.text = 1
End Sub

Private Sub YPos_GotFocus()
    YPos.SelStart = 0
    YPos.SelLength = Len(YPos.text)
End Sub
