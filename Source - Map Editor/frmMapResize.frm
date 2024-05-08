VERSION 5.00
Begin VB.Form frmMapResize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resize Map"
   ClientHeight    =   1695
   ClientLeft      =   3675
   ClientTop       =   3120
   ClientWidth     =   3255
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ResizeCmd 
      Caption         =   "Resize!"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "New Size"
      Height          =   975
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      Begin VB.TextBox NewHeight 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox NewWidth 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Height:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Size"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.Label CurrentHeight 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label CurrentWidth 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMapResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim map_x As Long
Dim map_y As Long

Private Sub Form_Load()
    'Load current values
    Engine.Map_Bounds_Get map_x, map_y
    
    CurrentWidth.Caption = "Width: " & Str(map_x)
    CurrentHeight.Caption = "Height: " & Str(map_y)
    
    NewWidth.text = map_x
    NewHeight.text = map_y
End Sub

Private Sub NewHeight_Change()
    'Make sure value is valid
    NewHeight.text = Int(Val(NewHeight.text))
    If NewHeight.text < 1 Then
        NewHeight.text = map_y
    End If
End Sub

Private Sub NewHeight_GotFocus()
    NewHeight.SelStart = 0
    NewHeight.SelLength = Len(NewHeight.text)
End Sub

Private Sub NewWidth_Change()
    'Make sure value is valid
    NewWidth.text = Int(Val(NewWidth.text))
    If NewWidth.text < 1 Then
        NewWidth.text = map_x
    End If
End Sub

Private Sub NewWidth_GotFocus()
    NewWidth.SelStart = 0
    NewWidth.SelLength = Len(NewWidth.text)
End Sub

Private Sub ResizeCmd_Click()
    If MsgBox("Are you sure you want to resize the current map?", vbOKCancel, "Resize Map") = vbOK Then
        Engine.Map_Resize Val(NewWidth.text), Val(NewHeight.text)
        Unload Me
    End If
End Sub
