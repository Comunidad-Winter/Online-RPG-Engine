VERSION 5.00
Begin VB.UserControl ctlDirectPlayClient 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   Picture         =   "ctlDirectPlayClient.ctx":0000
   ScaleHeight     =   555
   ScaleWidth      =   495
   Windowless      =   -1  'True
End
Attribute VB_Name = "ctlDirectPlayClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'ctlDirectPlayClient.ctl - ORE DirectPlay 8 Client - v0.5.0
'
'Handles all the TCP/IP traffic and ties all the client objects
'together.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

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

'***************************
'Required Externals
'***************************
'Reference to dx8vb.dll
'   - URL: http://www.microsoft.com/directx
'***************************
Option Explicit

'***************************
'Constants and enumerations
'***************************

'***************************
'Types
'***************************

'***************************
'Variables and objects
'***************************
Private dx As DirectX8                          'Main DirectX8 object
Private dp_client As DirectPlay8Client          'Server object, for message handling
Private dp_client_address As DirectPlay8Address 'Client's own IP, port
Private dp_server_address As DirectPlay8Address 'Server's own IP, port

Private engine As clsTileEngineX
Private client_app_guid As String

Private client_char_id As Long

'***************************
'Arrays
'***************************

'***************************
'External Functions
'***************************
'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Very percise 64 bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'***************************
'Events
'***************************
Public Event ClientConnected()
Public Event ClientConnectionFailed()
Public Event ClientDisconnected()
Public Event ClientAuthenticated()
Public Event EngineStop()
Public Event EngineStart()
Public Event ReceiveChatText(ByVal chat_text As String, ByVal rgb_color As Long)
Public Event ReceiveChatCritical(ByVal chat_text As String)

'***************************
Implements DirectPlay8Event
'***************************

Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal message_id As Long, ByVal connection_id As Long, ByVal lGroupID As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_AppDesc(byreffRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(byrefdpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_ConnectComplete(ByRef dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    If dpnotify.hResultCode = 0 Then
        RaiseEvent ClientConnected
    Else
        RaiseEvent ClientConnectionFailed
    End If
End Sub

Private Sub DirectPlay8Event_CreateGroup(ByVal group_id As Long, ByVal owner_id As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal connection_id As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_DestroyGroup(ByVal group_id As Long, ByVal reason_code As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_DestroyPlayer(ByVal connection_id As Long, ByVal reason_code As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_EnumHostsQuery(ByRef dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_QUERY, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_EnumHostsResponse(ByRef dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_RESPONSE, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_HostMigrate(ByVal new_host_id As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_IndicateConnect(ByRef dpnotify As DxVBLibA.DPNMSG_INDICATE_CONNECT, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_IndicatedConnectAborted(ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_InfoNotify(ByVal message_id As Long, ByVal notify_id As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_Receive(ByRef dpnotify As DxVBLibA.DPNMSG_RECEIVE, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    'Get packet header
    Dim offset As Long
    
    'Get header
    Dim header As ServerPacketHeader
    Call GetDataFromBuffer(dpnotify.ReceivedData, header, SIZE_LONG, offset)
    
    'Get command
    Dim command As ServerPacketCommand
    Call GetDataFromBuffer(dpnotify.ReceivedData, command, SIZE_LONG, offset)
    
    'Get parameter(s)
    Dim received_data As String
    Dim parameters() As String
    Dim loopc As Long
    Dim count As Long
    received_data = GetStringFromBuffer(dpnotify.ReceivedData, offset)
    count = General_Field_Count(received_data, P_DELIMITER_CODE)
    If count > 0 Then
        ReDim parameters(1 To count) As String
        For loopc = 1 To count
            parameters(loopc) = General_Field_Read(loopc, received_data, P_DELIMITER_CODE)
        Next loopc
    Else
        ReDim parameters(0 To 0) As String
    End If
    
    'Handle the packet
    If header = s_Player Then
        Receive_Player command, parameters()
        Exit Sub
    End If
    
    If header = s_Chat Then
        Receive_Chat command, parameters()
        Exit Sub
    End If
    
    If header = s_Map Then
        Receive_Map command, parameters()
        Exit Sub
    End If
    
    If header = s_Char Then
        Receive_Char command, parameters()
        Exit Sub
    End If
End Sub

Private Sub DirectPlay8Event_SendComplete(ByRef dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_TerminateSession(ByRef dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    RaiseEvent ClientDisconnected
End Sub

Private Sub UserControl_Initialize()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
End Sub

Private Sub UserControl_Terminate()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    Client_Deinitialize
End Sub

Public Function Client_Initialize(ByRef tile_engine As clsTileEngineX, ByVal app_guid As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
'On Error GoTo Errorhandler
    Set dx = New DirectX8
    Set dp_client = dx.DirectPlayClientCreate
    Set dp_client_address = dx.DirectPlayAddressCreate
    Set dp_server_address = dx.DirectPlayAddressCreate
    dp_client.RegisterMessageHandler Me
    dp_client_address.SetSP DP8SP_TCPIP
    dp_server_address.SetSP DP8SP_TCPIP
    
    'Set app_guid
    client_app_guid = app_guid
    
    'Set pointer to engine
    Set engine = tile_engine
    
   'Set delimiter
    P_DELIMITER = Chr$(P_DELIMITER_CODE)
    
    Client_Initialize = True
Exit Function
ErrorHandler:
    Client_Initialize = False
End Function

Public Sub Client_Deinitialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
On Error Resume Next
    dp_client.Close
    dp_client.CancelAsyncOperation 0, DPNCANCEL_ALL_OPERATIONS
    dp_client.UnRegisterMessageHandler

    Set dp_client = Nothing
    Set dp_client_address = Nothing
    Set dp_server_address = Nothing
    Set dx = Nothing
End Sub

Public Function Client_Connect(ByVal server_ip As String, ByVal server_port As Long, ByVal player_name As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
On Error GoTo ErrorHandler:
    'Set player name
    Dim player_info As DPN_PLAYER_INFO
    player_info.Name = player_name
    player_info.lInfoFlags = DPNINFO_NAME
    dp_client.SetClientInfo player_info, DPNOP_SYNC

    'Set server address
    dp_server_address.AddComponentString DPN_KEY_HOSTNAME, server_ip
    dp_server_address.AddComponentLong DPN_KEY_PORT, server_port
    
    'Connect
    Dim app_desc As DPN_APPLICATION_DESC
    app_desc.guidApplication = client_app_guid
    dp_client.Connect app_desc, dp_server_address, dp_client_address, 0, player_info.Name, Len(player_info.Name)
    
    Client_Connect = True
Exit Function
ErrorHandler:

End Function

Public Sub Client_Authenticate(ByVal password As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    'Send authentication string
    Send_Command c_Authenticate, c_Authenticate_Login, password
End Sub

Public Sub Client_Authenticate_New_Player(ByVal password As String, ByVal profile_name As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    'Send authentication string
    Send_Command c_Authenticate, c_Authenticate_New, c_Packet_Player_New(password, profile_name)
End Sub

Public Function Client_Connection_Status() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
On Error GoTo ErrorHandler:
    If dp_client.GetServerInfo.Name = "" Then
    End If
    Client_Connection_Status = True
Exit Function
ErrorHandler:
    Client_Connection_Status = False
End Function

Public Function Player_Move(ByVal heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'
'**************************************************************
    'Find char_index based on server id
    Dim char_index As Long
    char_index = engine.Char_Find(client_char_id)

    'Move the view position and the user_char
     If engine.Map_Legal_Char_Pos_By_Heading(char_index, heading) Then
         If engine.Engine_View_Move(heading) Then
            'Move player
             engine.Char_Move char_index, heading
            'Send Move command
            Send_Command c_Move, c_Move_Moved, CStr(heading)
            Player_Move = True
         End If
     End If
End Function

Public Static Function Player_Attack() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim frame_last_count As Long

    'Find char_index based on server id
    Dim char_index As Long
    char_index = engine.Char_Find(client_char_id)

    'If it's been at least 16 frames since the last attack allow the command
    If engine.Engine_Frame_Counter_Get - frame_last_count > 16 Then
        frame_last_count = engine.Engine_Frame_Counter_Get
        Send_Command c_Action, c_Action_Attack, ""
        Player_Attack = True
    Else
        Player_Attack = False
    End If
End Function

Private Sub Receive_Player(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'
'**************************************************************
        'Player Authenticated
        If command = s_Player_Authenticated Then
            RaiseEvent ClientAuthenticated
            Exit Sub
        End If
        
        'Engine Start
        If command = s_Player_Engine_Start Then
            RaiseEvent EngineStart
            Exit Sub
        End If
        
        'Engine Stop
        If command = s_Player_Engine_Stop Then
            RaiseEvent EngineStop
            Exit Sub
        End If
End Sub

Private Sub Receive_Chat(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
        'Text Box Chat
        If command = s_Chat_Text Then
            RaiseEvent ReceiveChatText(parameters(1), 0)
            Exit Sub
        End If
        
        'Critical Chat
        If command = s_Chat_Critical Then
            RaiseEvent ReceiveChatCritical(parameters(1))
            Exit Sub
        End If
End Sub

Private Sub Receive_Map(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'
'**************************************************************
        'Load Map
        If command = s_Map_Load Then
            engine.Map_Load_Map CLng(parameters(1))
            Exit Sub
        End If
End Sub

Private Sub Receive_Char(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'
'**************************************************************
        Dim char_index As Long

        'Create Char
        If command = s_Char_Create Then
            engine.Char_Create CLng(parameters(2)), CLng(parameters(3)), CLng(parameters(4)), _
                                CLng(parameters(5)), CLng(parameters(1))
            Exit Sub
        End If

       'Char ID
        If command = s_Char_ID_Set Then
            'Set player char id
            client_char_id = parameters(1)
            'Find char_index based on server id
            char_index = engine.Char_Find(client_char_id)
            'Recenter view on player char
            Dim x As Long, y As Long
            engine.Char_Map_Pos_Get char_index, x, y
            engine.Engine_View_Pos_Set x, y
            Exit Sub
        End If

        'Set Label
        If command = s_Char_Label_Set Then
            'Find char_index based on server id
            char_index = engine.Char_Find(CLng(parameters(1)))
            'Set label
            engine.Char_Label_Set char_index, parameters(2), 1 'CLng(parameters(3))
            Exit Sub
        End If
        
        'Set Char Data
        If command = s_Char_Data_Set Then
            'Find char_index based on server id
            char_index = engine.Char_Find(CLng(parameters(1)))
            'Set
            engine.Char_Data_Set char_index, CLng(parameters(2))
            Exit Sub
        End If
        
       'Set Char Body Data
        If command = s_Char_Data_Body_Set Then
            'Find char_index based on server id
            char_index = engine.Char_Find(CLng(parameters(1)))
            'Set
            engine.Char_Data_Body_Set char_index, CLng(parameters(2)), CByte(parameters(3))
            Exit Sub
        End If
        
        'Set Map Pos
        If command = s_Char_Pos_Set Then
            'Find char_index based on server id
            char_index = engine.Char_Find(CLng(parameters(1)))
            'Set
            engine.Char_Map_Pos_Set char_index, CLng(parameters(2)), CLng(parameters(3))
            'If client's char, center view on it
            If client_char_id = CLng(parameters(1)) Then
                engine.Engine_View_Pos_Set CLng(parameters(2)), CLng(parameters(3))
            End If
            Exit Sub
        End If

        'Set Heading
        If command = s_Char_Heading_Set Then
            'Find char_index based on server id
            char_index = engine.Char_Find(CLng(parameters(1)))
            'Set
            engine.Char_Heading_Set char_index, CLng(parameters(2))
            Exit Sub
        End If

        'Move Char
        If command = s_Char_Move Then
            'Ignore client's own movment
            If client_char_id = CLng(parameters(1)) Then
                Exit Sub
            End If
            'Find char_index based on server id
            char_index = engine.Char_Find(CLng(parameters(1)))
            'Move
            engine.Char_Move char_index, CLng(parameters(2))
            Exit Sub
        End If
        
        'Remove Char
        If command = s_Char_Remove Then
            'Find char_index based on server id
            char_index = engine.Char_Find(CLng(parameters(1)))
            'Remove
            engine.Char_Remove char_index
            Exit Sub
        End If
End Sub

Public Sub Chat_Send(ByVal message As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/1/2003
'
'**************************************************************
    'TODO: More checks to make sure users aren't inputing garbage
    If message <> "" Then
        'Send chat string
        Send_Command c_Chat, c_Chat_Global, message
    End If
End Sub

Private Sub Send_Command(ByVal header As ClientPacketHeader, ByVal command As ClientPacketCommand, ByRef parameters As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    'Don't send anything if not connected
    If Client_Connection_Status = False Then
        Exit Sub
    End If
        
    'New packet
    Dim byte_buffer() As Byte
    Dim offset As Long
    offset = NewBuffer(byte_buffer)
    
    'Add header
    Call AddDataToBuffer(byte_buffer, header, SIZE_LONG, offset)
    
    'Add command
    Call AddDataToBuffer(byte_buffer, command, SIZE_LONG, offset)

    'Add send_data
    Call AddStringToBuffer(byte_buffer, parameters, offset)

   'Send the packet, guarenteed, no loopback
    dp_client.Send byte_buffer, 0, DPNSEND_GUARANTEED Or DPNSEND_NOLOOPBACK
End Sub
