VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TileEngineX Demo"
   ClientHeight    =   11325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   755
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Remove"
      Height          =   255
      Left            =   9540
      TabIndex        =   25
      Top             =   9300
      Width           =   1155
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Remove"
      Height          =   255
      Left            =   7860
      TabIndex        =   24
      Top             =   8160
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Remove"
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   9540
      TabIndex        =   22
      Text            =   "11"
      Top             =   8940
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   9540
      TabIndex        =   21
      Text            =   "11"
      Top             =   8580
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   9540
      TabIndex        =   20
      Text            =   "1"
      Top             =   8160
      Width           =   1035
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Create &Exit"
      Height          =   495
      Left            =   9540
      TabIndex        =   19
      Top             =   7620
      Width           =   1635
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Create I&tem"
      Height          =   495
      Left            =   7860
      TabIndex        =   18
      Top             =   7620
      Width           =   1515
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Create &NPC"
      Height          =   495
      Left            =   6000
      TabIndex        =   17
      Top             =   7620
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   1275
      Left            =   4020
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMain.frx":0000
      Top             =   9960
      Width           =   7395
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Show Grh"
      Height          =   375
      Left            =   4020
      TabIndex        =   16
      Top             =   7740
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   4020
      TabIndex        =   15
      Text            =   "1001"
      Top             =   7380
      Width           =   1515
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   120
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   14
      Top             =   7380
      Width           =   3840
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Block Edges"
      Height          =   495
      Left            =   9780
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Toggle &Blocked"
      Height          =   495
      Left            =   9780
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Toggle Show Blocked"
      Height          =   495
      Left            =   9780
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Create &Grh"
      Height          =   495
      Left            =   9780
      TabIndex        =   12
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Make New Map"
      Height          =   555
      Left            =   9780
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Make Map Dark"
      Height          =   495
      Left            =   9780
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Make Map Light"
      Height          =   495
      Left            =   9780
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Create &Light"
      Height          =   495
      Left            =   9780
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create &Star Burst"
      Height          =   495
      Left            =   9780
      TabIndex        =   10
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save As Map 1"
      Height          =   495
      Left            =   9780
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Map 1"
      Height          =   495
      Left            =   9780
      TabIndex        =   13
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Toggle Engine Stats"
      Height          =   495
      Left            =   9780
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   120
      ScaleHeight     =   478
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   638
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   9600
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'TileEngineXDemo - v0.5.0
'*****************************************************************
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
'Aaron Perkins(aaron@baronsoft.com) - 5/15/2003
'   - First release
'*****************************************************************

Option Explicit

'***************************
'Variables
'***************************
Dim engine As New clsTileEngineX
Dim go As Boolean

Dim user_char_index As Long
Dim cursor_light_index As Long

Private Sub Command12_Click()
    engine.Map_Edges_Blocked_Set 10, 7, True
End Sub

Private Sub Command13_Click()
    Picture2.Cls
    engine.Grh_Render_To_Hdc CLng(Text1.text), Picture2.hdc, 0, 0, True
End Sub

Private Sub Command14_Click()
    'Get mouse position
    Dim temp_x As Long
    Dim temp_y As Long
    engine.Input_Mouse_Map_Get temp_x, temp_y
    
    'Add a npc
    If Check1.value = 0 Then
        engine.Map_NPC_Add temp_x, temp_y, 1
    Else
        engine.Map_NPC_Remove temp_x, temp_y
    End If
End Sub

Private Sub Command16_Click()
    'Get mouse position
    Dim temp_x As Long
    Dim temp_y As Long
    engine.Input_Mouse_Map_Get temp_x, temp_y
    
    'Add a item
    If Check2.value = 0 Then
        engine.Map_Item_Add temp_x, temp_y, 1, 5
    Else
        engine.Map_Item_Remove temp_x, temp_y
    End If
End Sub

Private Sub Command17_Click()
    'Get mouse position
    Dim temp_x As Long
    Dim temp_y As Long
    engine.Input_Mouse_Map_Get temp_x, temp_y
    
    'Add a exit
    If Check3.value = 0 Then
        engine.Map_Exit_Add temp_x, temp_y, Text3.text, CLng(Text4.text), CLng(Text5.text)
    Else
        engine.Map_Exit_Remove temp_x, temp_y
    End If
End Sub

Private Sub Form_Load()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Form Load
'*****************************************************************
    Me.Show
    
    '**************************************************************
    'Engine Initialization
    '**************************************************************
    
    'Run windowed
    go = engine.Engine_Initialize(frmMain.hWnd, Picture1.hWnd, 1, App.Path & "\..\Resources")
    
    'Run 640 by 480
    'go = engine.Engine_Initialize(frmMain.hWnd, frmMain.hWnd, 0, App.Path & "\..\Resources", 640, 480)
    
    'Run 800 by 600
    'go = engine.Engine_Initialize(frmMain.hWnd, frmMain.hWnd, 0, App.Path & "\..\Resources", 800, 600, , , CInt(800 / 32), CInt(600 / 32))
    
    'Run 1024 by 768
    'go = engine.Engine_Initialize(frmMain.hWnd, frmMain.hWnd, 0, App.Path & "\..\Resources", 1024, 768, , , CInt(1024 / 32), CInt(768 / 32))
    
    'Run 1024 by 768 with a 640 by 480 game area
    'go = engine.Engine_Initialize(frmMain.hWnd, frmMain.hWnd, 0, App.Path & "\..\Resources", 1024, 768, , , CInt(640 / 32), CInt(480 / 32))
    
    '****************************
    'set some engine parameters
    '****************************
    engine.Engine_Stats_Show_Toggle
    engine.Engine_Border_Color_Set &H101010 'Only seen full screen mode
    engine.Engine_Base_Speed_Set 0.03 'Speed that the engine should appear to run at
                                      '0.03 = roughly 30fps
    
    'Show NPCs, items and exits
    engine.Engine_Special_Tiles_Show_Toggle
    
    '****************************
    'set user pos
    '****************************
    engine.Engine_View_Pos_Set 11, 11
    
    '****************************
    'make a user char
    '****************************
    user_char_index = engine.Char_Create(11, 11, 1, 1, 0)
    
    '****************************
    'make cursor light
    '****************************
    cursor_light_index = engine.Light_Create(1, 1, RGB(255, 255, 255), 1)
    
    '**************************************************************
    'Main Loop
    '**************************************************************
    Do While go
       
        '****************************
        'Render next frame
        '****************************
        'Only run if the form is not minimized
        If Me.WindowState <> vbMinimized Then
            go = engine.Engine_Render_Start()
            
            '****************************
            'Display some GUI stuff
            '****************************
            engine.GUI_Box_Outline_Render 0, 0, 200, 100, 5, &HFF555555
            engine.GUI_Box_Filled_Render 0, 0, 200, 100, &HFF555555, , , &HFF000055, True
            engine.GUI_Grh_Render 2004, 150, 25, , True
            engine.GUI_Text_Render Text2.text, 1, 5, 5, 195, 95, &HFFFFFFFF, fa_left, True
            
            go = engine.Engine_Render_End()
        End If
    
        '****************************
        'Handle Inputs
        '****************************
        Check_Keys
        Check_Mouse
        
        '****************************
        'Do widnow's events
        '****************************
        DoEvents
        
    Loop
    
    '**************************************************************
    'Clean Up
    '**************************************************************
    engine.Engine_Deinitialize
    Set engine = Nothing
    Unload Me
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Form Unload
'*****************************************************************
    go = False
    Cancel = 1
End Sub

Sub Check_Keys()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Checks keys
'*****************************************************************
    'Escape
    If engine.Input_Key_Get(vbKeyEscape) Then
        go = False
    End If
    
    'S
    If engine.Input_Key_Get(vbKeyS) Then
        Command4_Click
    End If
    
    'L
    If engine.Input_Key_Get(vbKeyL) Then
        Command5_Click
    End If
    
    'G
    If engine.Input_Key_Get(vbKeyG) Then
        Command9_Click
    End If
    
    'B
    If engine.Input_Key_Get(vbKeyB) Then
        Command11_Click
    End If
    
    'N
    If engine.Input_Key_Get(vbKeyN) Then
        Command14_Click
    End If
    
    'T
    If engine.Input_Key_Get(vbKeyT) Then
        Command16_Click
    End If

    'E
    If engine.Input_Key_Get(vbKeyE) Then
        Command17_Click
    End If

    'I
    If engine.Input_Key_Get(vbKeyI) Then
        'Get mouse position
        Dim temp_x As Long
        Dim temp_y As Long
        engine.Input_Mouse_Map_Get temp_x, temp_y
        Dim file_path As String
        Dim x As Long
        Dim y As Long
        Dim widht As Long
        Dim height As Long
        Dim frame_count As Long
        engine.Grh_Info_Get engine.Map_Grh_Get(temp_x, temp_y, 1), file_path, x, y, width, height, frame_count
        
        MsgBox "Grh1: " & file_path & " Light: " & engine.Map_Light_Get(temp_x, temp_y) & " PG: " & engine.Map_Particle_Group_Get(temp_x, temp_y)
    End If
    
    'Check arrow keys
    If engine.Input_Key_Get(vbKeyUp) Then
        If engine.Map_Legal_Char_Pos_By_Heading(user_char_index, 1) Then
            If engine.Engine_View_Move(1) Then
                engine.Char_Move user_char_index, 1
            End If
        End If
    End If
    
    If engine.Input_Key_Get(vbKeyPageUp) Then
        If engine.Map_Legal_Char_Pos_By_Heading(user_char_index, 2) Then
            If engine.Engine_View_Move(2) Then
                engine.Char_Move user_char_index, 2
            End If
        End If
    End If
    
    If engine.Input_Key_Get(vbKeyRight) Then
        If engine.Map_Legal_Char_Pos_By_Heading(user_char_index, 3) Then
            If engine.Engine_View_Move(3) Then
                engine.Char_Move user_char_index, 3
            End If
        End If
    End If
    
    If engine.Input_Key_Get(vbKeyPageDown) Then
        If engine.Map_Legal_Char_Pos_By_Heading(user_char_index, 4) Then
            If engine.Engine_View_Move(4) Then
                engine.Char_Move user_char_index, 4
            End If
        End If
    End If
    
    If engine.Input_Key_Get(vbKeyDown) Then
        If engine.Map_Legal_Char_Pos_By_Heading(user_char_index, 5) Then
            If engine.Engine_View_Move(5) Then
                engine.Char_Move user_char_index, 5
            End If
        End If
    End If
    
    If engine.Input_Key_Get(vbKeyEnd) Then
        If engine.Map_Legal_Char_Pos_By_Heading(user_char_index, 6) Then
            If engine.Engine_View_Move(6) Then
                engine.Char_Move user_char_index, 6
            End If
        End If
    End If
    
    If engine.Input_Key_Get(vbKeyLeft) Then
        If engine.Map_Legal_Char_Pos_By_Heading(user_char_index, 7) Then
            If engine.Engine_View_Move(7) Then
                engine.Char_Move user_char_index, 7
            End If
        End If
    End If
    
    If engine.Input_Key_Get(vbKeyHome) Then
        If engine.Map_Legal_Char_Pos_By_Heading(user_char_index, 8) Then
            If engine.Engine_View_Move(8) Then
                engine.Char_Move user_char_index, 8
            End If
        End If
    End If
End Sub

Sub Check_Mouse()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Checks Mouse
'*****************************************************************
    'Make sure the mouse is in the view area
    If engine.Input_Mouse_In_View Then
    
        'Check the mouse for movement
        If engine.Input_Mouse_Moved_Get Then
            'Move the light over the cursor
            Dim temp_x As Long
            Dim temp_y As Long
            engine.Input_Mouse_Map_Get temp_x, temp_y
            engine.Light_Map_Pos_Set cursor_light_index, temp_x, temp_y
        End If
        
        'Check left button
        If engine.Input_Mouse_Button_Left_Get Then
            'Move the view position and the user_char
            If engine.Map_Legal_Char_Pos_By_Heading(user_char_index, engine.Input_Mouse_Heading_Get) Then
                If engine.Engine_View_Move(engine.Input_Mouse_Heading_Get) Then
                    engine.Char_Move user_char_index, engine.Input_Mouse_Heading_Get
                End If
            End If
        End If
        
        'Check right button
        If engine.Input_Mouse_Button_Right_Get Then
            engine.Char_Data_Body_Set user_char_index, 3, True
        End If
        
    End If
End Sub

Private Sub Command1_Click()
    '****************************
    'Load Map
    '****************************
    engine.Map_Load_Map 1, True
    
    '****************************
    'reset user pos
    '****************************
    engine.Engine_View_Pos_Set 11, 11
    
    '****************************
    'remake a user char
    '****************************
    user_char_index = engine.Char_Create(11, 11, 1, 1, 0)

    '****************************
    'remake cursor light
    '****************************
    cursor_light_index = engine.Light_Create(1, 1, RGB(255, 255, 255), 1)
End Sub


Private Sub Command2_Click()
    engine.Engine_Stats_Show_Toggle
End Sub

Private Sub Command3_Click()
    '****************************************
    'destroy cursor light so it is not saved
    '****************************************
    cursor_light_index = engine.Light_Remove(cursor_light_index)

    '****************************
    'Save Map
    '****************************
    engine.Map_Save_Map 1, True
    
    '****************************
    'remake cursor light
    '****************************
    cursor_light_index = engine.Light_Create(1, 1, RGB(255, 255, 255), 1)
End Sub

Private Sub Command4_Click()
    'Get mouse position
    Dim temp_x As Long
    Dim temp_y As Long
    engine.Input_Mouse_Map_Get temp_x, temp_y
    
    'Hold the grh_indexs that will be in the streams
    Dim temp_list() As Long
    
    'Add Star Burst
    ReDim temp_list(1 To 4) As Long
    temp_list(1) = 5 'set up grh_index array
    temp_list(2) = 6
    temp_list(3) = 7
    temp_list(4) = 8
    engine.Particle_Group_Create temp_x, temp_y, temp_list(), 20, 2, True
    
    'Add fountain
    ReDim temp_list(1 To 4) As Long
    temp_list(1) = 6 'set up grh_index array
    temp_list(2) = 6
    temp_list(3) = 6
    temp_list(4) = 6
    'engine.Particle_Group_Create temp_x, temp_y, temp_list(), 40, 1, True
End Sub

Private Sub Command5_Click()
    'Get mouse position
    Dim temp_x As Long
    Dim temp_y As Long
    engine.Input_Mouse_Map_Get temp_x, temp_y
    
    'Add a light
    engine.Light_Create temp_x, temp_y, RGB(255, 255, 255), 1
End Sub

Private Sub Command6_Click()
    engine.Map_Base_Light_Fill RGB(200, 200, 200)
End Sub

Private Sub Command7_Click()
    engine.Map_Base_Light_Fill RGB(100, 100, 100)
End Sub

Private Sub Command8_Click()
    engine.Map_Create 50, 50
    
    engine.Map_Fill 10, 1

    '****************************
    'reset user pos
    '****************************
    engine.Engine_View_Pos_Set 11, 11
    
    '****************************
    'remake a user char
    '****************************
    user_char_index = engine.Char_Create(11, 11, 1, 1, 0)

    '****************************
    'remake cursor light
    '****************************
    cursor_light_index = engine.Light_Create(1, 1, RGB(255, 255, 255), 1)
End Sub

Private Sub Command9_Click()
    'Get mouse position
    Dim temp_x As Long
    Dim temp_y As Long
    engine.Input_Mouse_Map_Get temp_x, temp_y
    
    engine.Map_Grh_Set temp_x, temp_y, 1001, 4, False
End Sub

Private Sub Command10_Click()
    'Toggle show blocked
    engine.Engine_Blocked_Tiles_Show_Toggle
End Sub

Private Sub Command11_Click()
    'Get mouse position
    Dim temp_x As Long
    Dim temp_y As Long
    engine.Input_Mouse_Map_Get temp_x, temp_y
    
    If engine.Map_Blocked_Get(temp_x, temp_y) Then
        engine.Map_Blocked_Set temp_x, temp_y, False
    Else
        engine.Map_Blocked_Set temp_x, temp_y, True
    End If
End Sub
