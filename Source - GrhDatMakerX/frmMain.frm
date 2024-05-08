VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grh.dat Maker"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2640
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Height          =   405
      Left            =   1530
      TabIndex        =   0
      Top             =   150
      Width           =   1035
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Press Go..."
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'GrhDatMakerX - v0.5.0
'***************************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'***************************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'Aaron Perkins(aaron@baronsoft.com) - 3/06/2003
'   - First release
'*****************************************************************
Option Explicit

'***************************
'Variables
'***************************

Private ini_path As String

Private Sub cmdGo_Click()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Loads Grh.raw, parses and outputs Grh.dat
'*****************************************************************
On Error GoTo ErrorHandler
    Dim sX As Long
    Dim sY As Long
    Dim pixelWidth As Long
    Dim pixelHeight As Long
    Dim FileNum As Long
    Dim NumFrames As Long
    Dim Frames(1 To 16) As Long
    Dim Speed As Single
    
    Dim LastGrh As Long
    Dim TempInt As Long
    Dim Grh As Long
    Dim Frame As Long
    Dim ln As String
    
    LastGrh = Val(General_Var_Get(ini_path & "grh.ini", "INIT", "grh_count"))
    
    'Delete any old file
    If General_File_Exists(ini_path & "grh.dat", vbNormal) = True Then
        Kill ini_path & "grh.dat"
    End If
            
    'Open new file
    Open ini_path & "grh.dat" For Binary As #1
    Seek #1, 1
    
    'Header
    Put #1, , TempInt
    Put #1, , TempInt
    Put #1, , TempInt
    Put #1, , TempInt
    Put #1, , TempInt
    
    'Fill Grh List
    For Grh = 1 To LastGrh
    
        DoEvents
    
        lblStatus.Caption = Grh & "/" & LastGrh & " Grhs..."
        lblStatus.Refresh
    
        'Get line from fisrt file
        ln = General_Var_Get(ini_path & "grh1.raw", "Graphics", "Grh" & Grh)
        'If not found try other files
        If ln = "" Then
            ln = General_Var_Get(ini_path & "grh2.raw", "Graphics", "Grh" & Grh)
        End If
        If ln = "" Then
            ln = General_Var_Get(ini_path & "grh3.raw", "Graphics", "Grh" & Grh)
        End If
        If ln = "" Then
            ln = General_Var_Get(ini_path & "grh4.raw", "Graphics", "Grh" & Grh)
        End If
        If ln = "" Then
            ln = General_Var_Get(ini_path & "grh5.raw", "Graphics", "Grh" & Grh)
        End If
        
        If ln <> "" Then
        
            'Get number of frames and check
            NumFrames = Val(General_Field_Read(1, ln, 45))
            If NumFrames <= 0 Then GoTo ErrorHandler
        
            'Put grh number
            Put #1, , Grh
            'Put number of frames
            Put #1, , NumFrames
            
            If NumFrames > 1 Then
        
                'Read a animation GRH set
                For Frame = 1 To NumFrames
            
                    'Check and put each frame
                    Frames(Frame) = Val(General_Field_Read(Frame + 1, ln, 45))
                    If Frames(Frame) <= 0 Or Frames(Frame) > LastGrh Then GoTo ErrorHandler
                    Put #1, , Frames(Frame)
            
                Next Frame
        
                'Check and put speed
                Speed = CSng(General_Field_Read(NumFrames + 2, ln, 45))
                If Speed = 0 Then GoTo ErrorHandler
                Put #1, , Speed
            
            Else
        
                'check and put normal GRH data
                FileNum = Val(General_Field_Read(2, ln, 45))
                If FileNum <= 0 Then GoTo ErrorHandler
                Put #1, , FileNum
                
                sX = Val(General_Field_Read(3, ln, 45))
                If sX < 0 Then GoTo ErrorHandler
                Put #1, , sX
                
                sY = Val(General_Field_Read(4, ln, 45))
                If sY < 0 Then GoTo ErrorHandler
                Put #1, , sY
                
                pixelWidth = Val(General_Field_Read(5, ln, 45))
                If pixelWidth <= 0 Then GoTo ErrorHandler
                Put #1, , pixelWidth
                
                pixelHeight = Val(General_Field_Read(6, ln, 45))
                If pixelHeight <= 0 Then GoTo ErrorHandler
                Put #1, , pixelHeight
    
            End If
            
        End If
        
    Next Grh
    '************************************************
    
    Close #1
    lblStatus.Caption = "Done!"
    lblStatus.Refresh
Exit Sub
ErrorHandler:
    Close #1
    MsgBox "Error while loading the grhX.raw! Stopped at GRH number: " & Grh
    lblStatus.Caption = "Error!"
    lblStatus.Refresh
End Sub

Private Sub Form_Load()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Set ini path to app path
    ini_path = App.Path & "\"
End Sub




