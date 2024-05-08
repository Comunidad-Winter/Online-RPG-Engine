VERSION 5.00
Begin VB.Form DatMaker 
   Caption         =   "Raw Maker"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGo 
      Caption         =   "&Go"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "DatMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Grh.dat to Raw Converter
'Jonathan Valentin 2002
'http://www.visualbasiczone.com
'Ore Game at
'http://projectx.idlegames.net

'**********************************
'Edited by Maraxus to fit ORE 0.5 and 1.0
'Edited to use 5 .raw files
'Several bugs corrected
'**********************************
Option Explicit

Dim numgrhs As Integer

Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Long
Dim Frame As Long
Dim TempInt As Long
Dim GrhPath As String

'Get Number of Graphics
GrhPath = App.Path & "\"
numgrhs = Val(GetVar(App.Path & "\" & "Grh.ini", "INIT", "grh_count"))

'Resize arrays
ReDim GrhData(1 To numgrhs) As GrhData

'Open files
Open IniPath & "Grh.dat" For Binary As #1
Seek #1, 1

'Get Header
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
        
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > numgrhs Then GoTo ErrorHandler
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Private Sub CmdGo_Click()
Call LoadGrhData
Call MakeRaw
End Sub

Private Sub MakeRaw()
Dim i As Long
Dim file_id As Byte
Dim wrote As Long
Dim Animation As String
Dim Frame As Integer

Open "grh1.raw" For Output As #1
Print #1, "[Graphics]"
file_id = 1
For i = 1 To numgrhs
    If wrote = 2000 Then
        wrote = wrote
    End If
    
    If wrote = 2000 And file_id = 1 Then
        file_id = 2
        Close #1
        Open "grh2.raw" For Output As #1
        Print #1, "[Graphics]"
    ElseIf wrote = 4000 And file_id = 2 Then
        file_id = 3
        Close #1
        Open "grh3.raw" For Output As #1
        Print #1, "[Graphics]"
    ElseIf wrote = 6000 And file_id = 3 Then
        file_id = 4
        Close #1
        Open "grh4.raw" For Output As #1
        Print #1, "[Graphics]"
    ElseIf wrote = 8000 And file_id = 4 Then
        file_id = 5
        Close #1
        Open "grh5.raw" For Output As #1
        Print #1, "[Graphics]"
    End If

    If GrhData(i).NumFrames = 0 And GrhData(i).FileNum = 0 And GrhData(i).sX = 0 Then
    Else
        wrote = wrote + 1
        If GrhData(i).NumFrames > 1 Then
            Animation = ""
            For Frame = 1 To GrhData(i).NumFrames
                Animation = Animation & GrhData(i).Frames(Frame) & "-"
            Next Frame

            Print #1, "Grh" & i & "=" & GrhData(i).NumFrames & "-" & Animation & GrhData(i).Speed
        Else
            Print #1, "Grh" & i & "=" & GrhData(i).NumFrames & "-" & GrhData(i).FileNum & "-" & GrhData(i).sX & "-" & GrhData(i).sY & "-" & GrhData(GrhData(i).Frames(1)).pixelWidth & "-" & GrhData(GrhData(i).Frames(1)).pixelHeight
        End If
    End If
Next i
Close #1

End Sub
