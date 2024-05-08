VERSION 5.00
Begin VB.Form frmBinary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORE Compressor"
   ClientHeight    =   2055
   ClientLeft      =   2325
   ClientTop       =   1620
   ClientWidth     =   3135
   Icon            =   "frmBinary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3135
   Begin VB.Frame Frame1 
      Caption         =   "File Type"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton Option1 
         Caption         =   "Patch"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Scripts"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Wav"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MP3"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MIDI"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Graphics"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Extract"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compress"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frmBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim File_Type_Index As Byte

Private Sub Command1_Click()
    frmProgress.Show
    frmProgress.Label1.Caption = "Compressing..."
    Compress_Files File_Type_Index, App.Path, App.Path & "\Output\", frmProgress.ProgressBar1
    Unload frmProgress
End Sub

Private Sub Command2_Click()
    Dim LoopC As Long
    
    frmProgress.Show
    frmProgress.Label1.Caption = "Decompressing..."
    If File_Type_Index <> Patch Then
        Extract_Files File_Type_Index, App.Path, frmProgress.ProgressBar1, Nothing, Nothing, True
    Else
        Extract_Patch App.Path & "\", "Patch.ORE", Nothing, frmProgress.ProgressBar1, Nothing
    End If
    Close_All
    Unload frmProgress
End Sub

Private Sub Option1_Click(Index As Integer)
    File_Type_Index = Index
End Sub
