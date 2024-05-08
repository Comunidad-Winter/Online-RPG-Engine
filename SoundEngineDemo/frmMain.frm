VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Sound Engine"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   120
      Max             =   0
      Min             =   -4000
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   2280
      Max             =   200
      Min             =   -4000
      TabIndex        =   7
      Top             =   840
      Value           =   200
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   0
      Min             =   -10000
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop Midi"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Play Midi"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop and Empty MP3"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load and Play MP3"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Wav"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlayWav 
      Caption         =   "Play Wav"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "This program demonstrates the diffirent functions in Sound Engine..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   10
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Sound Engine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SEngine As clsSoundEngine

Private Sub cmdPlayWav_Click()
SEngine.Sound_Play 1, False
SEngine.Sound_Play 2, True
End Sub

Private Sub Command1_Click()
SEngine.Sound_Stop (1)
SEngine.Sound_Stop (2)
End Sub

Private Sub Command2_Click()
SEngine.Music_MP3_Load App.Path & "\TLN.mp3"
SEngine.Music_MP3_Play
End Sub

Private Sub Command3_Click()
SEngine.Music_MP3_Stop
SEngine.Music_MP3_Empty
End Sub

Private Sub Command4_Click()
SEngine.Music_Midi_Play 1, False
End Sub

Private Sub Command5_Click()
SEngine.Music_Midi_Stop 1
End Sub

Private Sub Form_Load()
Set SEngine = New clsSoundEngine

SEngine.Engine_Initialize frmMain.hWnd, App.Path & "\..\Resources"
End Sub

Private Sub Form_Unload(Cancel As Integer)
SEngine.Engine_DeInitialzie
End Sub

Private Sub HScroll1_Change()
SEngine.Sound_Volume_Set 2, HScroll1.value
End Sub

Private Sub HScroll2_Change()
SEngine.Music_Midi_Volume_Set HScroll2.value
End Sub

Private Sub HScroll3_Change()
SEngine.Music_MP3_Volume_Set HScroll3.value
End Sub
