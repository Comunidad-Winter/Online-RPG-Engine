VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORE Client"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   12720
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picMainView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   60
      ScaleHeight     =   478
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   638
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   9600
   End
   Begin VB.TextBox txtChatReceive 
      Height          =   2055
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7320
      Width           =   9615
   End
   Begin VB.TextBox txtChatSend 
      Height          =   375
      Left            =   60
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   9480
      Width           =   9615
   End
   Begin OREClient.ctlDirectPlayClient dp_client 
      Left            =   12120
      Top             =   9300
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Label lblServerStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Disconnected"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9720
      TabIndex        =   3
      Top             =   60
      Width           =   2955
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'frmMain.frm - ORE Client - v0.5.0
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
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************
Option Explicit

Public server_ip As String
Public server_port As Long
Public player_name As String
Public player_password As String
Public player_profile_name As String

Public main_loop_go As Boolean
Public tile_engine_go As Boolean
Public connection_mode_new As Boolean

Public tile_engine As New clsTileEngineX

Private on_screen_text_buffer As String

Private Sub Form_Load()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'
'**************************************************************
   Main
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'*****************************************************************
    main_loop_go = False
    Cancel = 1
End Sub

Private Sub Main()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/12/2003
'Main Loop
'*****************************************************************
    '**************************************************************
    'Client Initialization
    '**************************************************************
    'Initialize TileEngine
    main_loop_go = tile_engine.Engine_Initialize(frmMain.hWnd, picMainView.hWnd, 1, App.Path & "\..\Resources")
    'main_loop_go = tile_engine.Engine_Initialize(frmMain.hWnd, frmMain.hWnd, 0, App.Path & "\..\Resources", 1024, 768, , , CInt(1024 / 32), CInt(768 / 32))
    
    tile_engine.Engine_Border_Color_Set &H101010 'Only seen in full screen mode
    tile_engine.Engine_Base_Speed_Set 0.03 'Speed that the engine should appear to run at
                                            '0.03 = roughly 30fps
                                      
    'Initialize direct play client
    main_loop_go = dp_client.Client_Initialize(tile_engine, _
                        "{12345678-1234-1234-1234-123456789ABC}") 'This GUID needs to match the on in the server.ini
          
    'Turn on engine stats display
    'tile_engine.Engine_Stats_Show_Toggle

    'Show connect screen
    Me.Hide
    frmConnect.Show
    General_Form_On_Top_Set frmConnect, True

    'Enter main loop
    Do While main_loop_go
        '****************************
        'Render next frame
        '****************************
        If tile_engine_go Then
        
            'Only run if the form is not minimized
            If Me.WindowState <> vbMinimized Then
                main_loop_go = tile_engine.Engine_Render_Start
                main_loop_go = tile_engine.Engine_Render_End
            End If
            
            '****************************
            'Handle Inputs
            '****************************
            Check_Keys
            Check_Mouse
        End If
    
        '****************************
        'Do other events
        '****************************
        DoEvents
    Loop
 
   '**************************************************************
    'Clean Up
    '**************************************************************
    Set tile_engine = Nothing
    dp_client.Client_Deinitialize
    Unload frmConnect
    Unload frmNewPlayer
    Unload Me
    End
End Sub

Static Sub Check_Keys()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Checks keys
'*****************************************************************
    'Check arrow keys
    If tile_engine.Input_Key_Get(vbKeyUp) Then
        dp_client.Player_Move 1
    End If
    
    If tile_engine.Input_Key_Get(vbKeyPageUp) Then
        dp_client.Player_Move 2
    End If
    
    If tile_engine.Input_Key_Get(vbKeyRight) Then
        dp_client.Player_Move 3
    End If
    
    If tile_engine.Input_Key_Get(vbKeyPageDown) Then
        dp_client.Player_Move 4
    End If
    
    If tile_engine.Input_Key_Get(vbKeyDown) Then
        dp_client.Player_Move 5
    End If
    
    If tile_engine.Input_Key_Get(vbKeyEnd) Then
        dp_client.Player_Move 6
    End If
    
    If tile_engine.Input_Key_Get(vbKeyLeft) Then
        dp_client.Player_Move 7
    End If
    
    If tile_engine.Input_Key_Get(vbKeyHome) Then
        dp_client.Player_Move 8
    End If
    
    If tile_engine.Input_Key_Get(vbKeyHome) Then
        dp_client.Player_Move 8
    End If
End Sub

Sub Check_Mouse()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Checks Mouse
'*****************************************************************
    'Make sure the mouse is in the view area
    If tile_engine.Input_Mouse_In_View Then
        
        'Check left button
        If tile_engine.Input_Mouse_Button_Left_Get Then
            'Move
            dp_client.Player_Move tile_engine.Input_Mouse_Heading_Get
        End If
        
       'Check right button
        If tile_engine.Input_Mouse_Button_Right_Get Then
            'Attack
            dp_client.Player_Attack
        End If
        
    End If
End Sub

Private Sub dp_client_ClientConnected()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    'Set Status
    lblServerStatus.Caption = "Connected: Trying to authenticate..."
    
    'Try to authenticate
    If connection_mode_new Then
        'Send authentication for new player
        dp_client.Client_Authenticate_New_Player player_password, player_profile_name
    Else
        'Send authentication
        dp_client.Client_Authenticate player_password
    End If
End Sub

Private Sub dp_client_ClientConnectionFailed()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    'Send Message
    MsgBox "Error connecting to server."
    'Set Status to Disconnect
    Call dp_client_ClientDisconnected
End Sub

Private Sub dp_client_ClientAuthenticated()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    'Set Status
    lblServerStatus.Caption = "Connected: Loading ..."
End Sub

Private Sub dp_client_ClientDisconnected()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/7/2003
'
'**************************************************************
    'Set Status
    lblServerStatus.Caption = "Disconnected"
    
    'Reset everything
    tile_engine_go = False
    tile_engine.Map_Create 50, 50
    txtChatReceive.text = ""
    
    'Show connect screen
    Me.Hide
    frmConnect.Show
    General_Form_On_Top_Set frmConnect, True
End Sub

Private Sub dp_client_EngineStart()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    'Start engine
    tile_engine_go = True
    'Set Status
    lblServerStatus.Caption = "Ready"
End Sub

Private Sub dp_client_EngineStop()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    'Start engine
    tile_engine_go = False
    'Set Status
    lblServerStatus.Caption = "Loading ..."
End Sub

Private Sub dp_client_ReceiveChatText(ByVal chat_text As String, ByVal rgb_color As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    'Put text in chatbox and scroll to bottom
    txtChatReceive.SelStart = Len(txtChatReceive.text)
    txtChatReceive.SelLength = 0
    txtChatReceive.SelText = vbNewLine & chat_text
End Sub

Private Sub dp_client_ReceiveChatCritical(ByVal chat_text As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    MsgBox chat_text
End Sub

Private Sub txtChatSend_KeyPress(KeyAscii As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    If KeyAscii = 13 Then
        dp_client.Chat_Send txtChatSend.text
        txtChatSend.text = ""
        KeyAscii = 0
    End If
End Sub
