VERSION 5.00
Begin VB.Form frmConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   3450
   ClientLeft      =   2475
   ClientTop       =   2145
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4110
   Begin VB.CommandButton CancelCmd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.DirListBox ScriptsDir 
      Height          =   2115
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.DriveListBox ScriptsDrive 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton OKCmd 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.DriveListBox GraphicsDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.DirListBox GraphicsDir 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Scripts path:"
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Graphics path:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If General_File_Exists(left$(graphics_path, 2), vbVolume) Then
        GraphicsDrive.Drive = left$(graphics_path, 2)
        If General_File_Exists(graphics_path, vbDirectory) Then
            GraphicsDir.Path = graphics_path
        End If
    End If
    If General_File_Exists(left$(scripts_path, 2), vbVolume) Then
        ScriptsDrive.Drive = left$(scripts_path, 2)
        If General_File_Exists(scripts_path, vbDirectory) Then
            ScriptsDir.Path = scripts_path
        End If
    End If
End Sub

Private Sub GraphicsDrive_Change()
    If General_File_Exists(GraphicsDrive.Drive, vbVolume) Then
        GraphicsDir.Path = GraphicsDrive.Drive
    Else
        GraphicsDrive.Drive = left$(graphics_path, 2)
        GraphicsDir.Path = graphics_path
    End If
End Sub

Private Sub OKCmd_Click()
    scripts_path = ScriptsDir.Path
    graphics_path = GraphicsDir.Path
    
    'Save changes to ini file
    General_Var_Write App.Path & "\ORE Scripter.ini", "INIT", "graphics_path", graphics_path
    General_Var_Write App.Path & "\ORE Scripter.ini", "INIT", "scripts_path", scripts_path
    
    Unload Me
End Sub

Private Sub ScriptsDrive_Change()
    If General_File_Exists(ScriptsDrive.Drive, vbVolume) Then
        ScriptsDir.Path = ScriptsDrive.Drive
    Else
        ScriptsDrive.Drive = left$(scripts_path, 2)
        ScriptsDir.Path = scripts_path
    End If
End Sub
