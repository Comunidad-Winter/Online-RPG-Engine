VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "MiniMap Color Finder"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   360
      Left            =   6600
      TabIndex        =   1
      Top             =   7170
      Width           =   1020
   End
   Begin VB.Label lblstatus 
      Caption         =   "grh's loaded!"
      Height          =   240
      Left            =   285
      TabIndex        =   0
      Top             =   7260
      Width           =   4320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim running As Boolean

Private Sub Command1_Click()
    If running Then Exit Sub
    
    running = True
    
    Dim i As Long
    
    If General_File_Exists(App.Path & "\minimap.dat", vbNormal) Then Kill App.Path & "\minimap.dat"
    
    Open App.Path & "\minimap.dat" For Binary As #1
    
    For i = 1 To grh_count
        Me.Cls
        lblstatus.Caption = "Grh " & i & "/" & UBound(grh_list())
        If Grh_Check(i) Then
            Put #1, , Grh_get_value(i, Me.hdc, 0, 0, False)
        End If
        DoEvents
    Next i
    
    Close #1
    
    lblstatus.Caption = "Done!"
    
    running = False
End Sub

Private Sub Form_Load()
    Extract_Files grh, App.Path & "\..\Resources", Nothing, Nothing, Nothing
    loadgrh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Delete_Resources App.Path & "\..\Resources"
End Sub
