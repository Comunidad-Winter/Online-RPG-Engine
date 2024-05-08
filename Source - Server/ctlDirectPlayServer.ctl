VERSION 5.00
Begin VB.UserControl ctlDirectPlayServer 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   Picture         =   "ctlDirectPlayServer.ctx":0000
   ScaleHeight     =   570
   ScaleWidth      =   525
   Windowless      =   -1  'True
   Begin VB.Timer timTickCounter 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   60
   End
End
Attribute VB_Name = "ctlDirectPlayServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'ctlDirectPlayServer.ctl - ORE DirectPlay 8 Server - v0.5.0
'
'Handles the TCP/IP traffic and ties all the server objects together
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
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 9/02/2004
'   - Add: WorldSave method
'   - Add: Character Creation
'   - Add: Chats
'   - Add: Accounts methods
'   - Add: Login / Logout (all 3 levels: char, Account, server)
'   - Add: Speeches are loaded
'   - Change: Maps are loaded by name
'   - Change: Fixed Sysop bugs and added logs
'   Sub Release Contributors:
'       David Justus - 8/14/2004
'           - Add: Sysop operations
'           - Add: GUMP code
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
'Constants
'***************************
Private Const PATH_PLAYERS = "\players"
Private Const PATH_MAPS = "\maps"
Private Const PATH_SCRIPTS = "\scripts"
'Added by Juan Martín Sotuyo Dodero
Private Const PATH_ACCOUNTS = "\accounts"
Private Const PATH_SPEECH = "\speechs"

Private Const SERVER_TICK_INTERVAL = 100  'In milliseconds

Public Enum Server_Status
    s_s_none = 0
    s_s_listening = 1
    s_s_shutting_down = 2
    s_s_closed = 3
End Enum

Public Enum Command_Send_Type
    to_id
    to_All
End Enum

'***************************
'Types
'***************************
Private Type Session_Variable
    variable_name As String
    variable_data As Variant
    variable_save As Boolean
End Type

'***************************
'Variables
'***************************
Private dx As DirectX8                          'Main DirectX8 object
Private dp_server As DirectPlay8Server          'Server object, for message handling
Private dp_server_address As DirectPlay8Address 'Server's own IP, port

Private server_state As Server_Status
Private server_ticks As Long

Private server_connection_id As Long
Private server_players_connection_id As Long

Private resource_path As String
Private players_path As String
Private maps_path As String
Private scripts_path As String
'Added by Juan Martín Sotuyo Dodero
Private accounts_path As String
Private speech_path As String

Private script_engine As New clsScriptEngine
Private script_interface As New clsScriptInterface
Private script_gump As New clsGump

Private item_count As Long

'***************************
'Arrays
'***************************
Dim player_list As clsList
Dim npc_list As clsList
Dim map_list As clsList
Dim char_list As clsList
'Added by Juan Martín Sotuyo Dodero
Dim speech_list As clsList
Dim gump_list As clsList


Private session_variable_list() As Session_Variable

'***************************
'External Functions
'***************************
'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'***************************
'Events
'***************************
Event ServerConnectionAdded(connection_id As Long)
Event ServerConnectionRemoved(connection_id As Long)

'***************************
Implements DirectPlay8Event
'***************************

'Sysop handler

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
'Last Modify Date: 4/30/2003
'
'*****************************************************************
    Deinitialize
End Sub

Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal message_id As Long, ByVal connection_id As Long, ByVal group_id As Long, fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_AppDesc(ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(ByRef dpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_ConnectComplete(ByRef dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_CreateGroup(ByVal group_id As Long, ByVal owner_id As Long, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Store ID's of each group as they come in
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    'Get new group info
    Dim group_info As DPN_GROUP_INFO
    group_info = dp_server.GetGroupInfo(group_id)
    
    'See if it's the all players group
    If group_info.Name = "PLAYERS" Then
        'Save it
        server_players_connection_id = group_id
        Exit Sub
    End If
    
    'See if it's a map group
    If Left(group_info.Name, 4) = "MAP " Then
        'Save it to map object
        Dim Name As String
        Name = Right(group_info.Name, Len(group_info.Name) - 4)
        Dim map As clsMap
        Set map = map_list.Find("name", Name)
        map.ConnectionID = group_id
        Exit Sub
    End If
End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal connection_id As Long, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/2/2003
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    'The first createplayer is always the server
    If server_connection_id = 0 Then
        'Save it
        server_connection_id = connection_id
        Exit Sub
    End If
    
    'Get the playerinfo
    Dim peer_info As DPN_PLAYER_INFO
    peer_info = dp_server.GetClientInfo(connection_id)
    
    'Check Name
    If General_String_Is_Alphanumeric(peer_info.Name) = False Then
        'Boot player
          Send_Command to_id, connection_id, s_Chat, s_Chat_Critical, "Invalid character found in Account name."
          dp_server.DestroyClient connection_id, 0, 0, 0
          Exit Sub
    End If

    'Don´t allow multiple login with one Account
    Dim tempplayer As clsPlayer
    Set tempplayer = player_list.Find("Name_Upper_Case", UCase$(peer_info.Name)) 'Use upper case so we are sure there isn't a match
    If Not (tempplayer Is Nothing) Then
        'Boot player
        Send_Command to_id, connection_id, s_Chat, s_Chat_Critical, "Player Account is already logged on."
        dp_server.DestroyClient connection_id, 0, 0, 0
        Exit Sub
    End If
    
    'Initialize player object and add to player list
    Dim player_id As Long
    Dim newplayer As New clsPlayer
    player_id = player_list.Add(newplayer)
    newplayer.Initialize Me, script_engine, map_list, player_list, npc_list, speech_list, player_id, accounts_path, peer_info.Name, players_path, resource_path
    newplayer.ConnectionID = connection_id
    newplayer.ConnectionStatus = p_cs_connected
    
    'Add player to players connection group
    dp_server.AddPlayerToGroup server_players_connection_id, newplayer.ConnectionID
    
    'Throw Event
    RaiseEvent ServerConnectionAdded(connection_id)
End Sub

Private Sub DirectPlay8Event_DestroyGroup(ByVal group_id As Long, ByVal reason_code As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_DestroyPlayer(ByVal connection_id As Long, ByVal reason_code As Long, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Edited by Juan Martín Sotuyo Dodero (Maraxus)
'**************************************************************
    'Make sure it's not the server
    If server_connection_id = connection_id Then
        Exit Sub
    End If
    
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Find("ConnectionID", connection_id)
    
    'See the connection_id has a player object
    If Not (player Is Nothing) Then
        'Logoff player properly
        Logoff_Session_Terminate player.id
    End If
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
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'**************************************************************
    'Check if a client´s info was changed
    If message_id = DPN_MSGID_CLIENT_INFO Then
        Dim player As clsPlayer
        Dim LoopC As Long
        Dim peer_info As DPN_PLAYER_INFO
        Dim temp_peer As DPN_PLAYER_INFO
        
        'Get player
        peer_info = dp_server.GetClientInfo(notify_id)
        For LoopC = player_list.LowerBound To player_list.UpperBound
            temp_peer = dp_server.GetClientInfo(CallByName(player_list.Item(LoopC), "ConnectionID", VbGet))
            If peer_info.Name = temp_peer.Name Then
                'We found it
                Set player = player_list.Item(LoopC)
                Exit For
            End If
        Next LoopC
        
        'Update name
        If Not player Is Nothing Then
            'If it´s logged in, log it off
            If player.Logged_In Then
                Logoff_Account player.id
            End If
            
            player.Name = peer_info.Name
        End If
    End If
End Sub

Private Sub DirectPlay8Event_Receive(ByRef dpnotify As DxVBLibA.DPNMSG_RECEIVE, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    'If the sender does not have a player object ignore the packet
    Dim player As clsPlayer
    Set player = player_list.Find("ConnectionID", dpnotify.idSender)
    If player Is Nothing Then
        Exit Sub
    End If

    'Get packet header
    Dim offset As Long
    
    'Get header
    Dim header As ClientPacketHeader
    Call GetDataFromBuffer(dpnotify.ReceivedData, header, SIZE_LONG, offset)
    
    'Get command
    Dim command As ClientPacketCommand
    Call GetDataFromBuffer(dpnotify.ReceivedData, command, SIZE_LONG, offset)
    
    'Get parameter(s)
    Dim received_data As String
    Dim parameters() As String
    Dim LoopC As Long
    Dim Count As Long
    received_data = GetStringFromBuffer(dpnotify.ReceivedData, offset)
    Count = General_Field_Count(received_data, P_DELIMITER_CODE)
    
    If Count > 0 Then
        ReDim parameters(1 To Count) As String
        For LoopC = 1 To Count
            parameters(LoopC) = General_Field_Read(LoopC, received_data, P_DELIMITER_CODE)
        Next LoopC
    Else
        ReDim parameters(0 To 0) As String
    End If
    
    'Handle the packet
    If header = c_Authenticate Then
        Receive_Authenticate player.id, command, parameters()
        Exit Sub
    End If
    
    If header = c_Account Then
        Receive_Account player.id, command, parameters()
        Exit Sub
    End If
    
    If header = c_Chat Then
        Receive_Chat player.id, command, parameters()
        Exit Sub
    End If
    
    If header = c_Move Then
        Receive_Move player.id, command, parameters()
        Exit Sub
    End If

    If header = c_Action Then
        Receive_Action player.id, command, parameters()
        Exit Sub
    End If
    
    If header = c_Request Then
        Receive_Request player.id, command, parameters()
        Exit Sub
    End If
    
    If header = c_Logoff Then
        Receive_Logoff player.id, command, parameters()
        Exit Sub
    End If
    
    If header = c_Gump Then
        Receive_Gump player.id, command, parameters()
        Exit Sub
    End If
    
    If header = c_Sysop Then
        Receive_Sysop player.id, command, parameters()
        Exit Sub
    End If
End Sub

Private Sub DirectPlay8Event_SendComplete(ByRef dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_TerminateSession(ByRef dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Public Property Get Player_Count() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'**************************************************************
    Player_Count = player_list.Count
End Property

Public Property Get ServerStatus() As Server_Status
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'**************************************************************
    ServerStatus = server_state
End Property

Public Function Initialize(ByVal app_guid As String, ByVal server_port As String, ByVal max_players As Long, ByVal session_name As String, ByVal s_resource_path As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
On Error GoTo ErrorHandler

    'Paths
    resource_path = s_resource_path
    players_path = resource_path & PATH_PLAYERS
    maps_path = resource_path & PATH_MAPS
    scripts_path = resource_path & PATH_SCRIPTS
    accounts_path = resource_path & PATH_ACCOUNTS
    speech_path = resource_path & PATH_SPEECH
    
    'Set delimiter
    P_DELIMITER = Chr$(P_DELIMITER_CODE)
        
    'DirectPlay
    Dim app_desc As DPN_APPLICATION_DESC     '
    With app_desc
        .guidApplication = app_guid
        .lMaxPlayers = max_players
        .SessionName = session_name
        .lFlags = DPNSESSION_CLIENT_SERVER
    End With
    Set dx = New DirectX8
    Set dp_server = dx.DirectPlayServerCreate
    Set dp_server_address = dx.DirectPlayAddressCreate
    dp_server.RegisterMessageHandler Me
    dp_server_address.SetSP DP8SP_TCPIP
    dp_server_address.AddComponentLong DPN_KEY_PORT, server_port
    dp_server.Host app_desc, dp_server_address
    
    'Lists
    Set player_list = New clsList
    Set npc_list = New clsList
    Set map_list = New clsList
    Set char_list = New clsList
    Set speech_list = New clsList
    Set gump_list = New clsList
    

    'Create All Player group
    Dim group_info As DPN_GROUP_INFO
    group_info.lInfoFlags = DPNINFO_NAME
    group_info.Name = "PLAYERS"
    dp_server.CreateGroup group_info

    'Check player directory
    If General_File_Exists(players_path, vbDirectory) = False Then
        'make the directory
        MkDir App.Path & PATH_PLAYERS
    End If
    
    'Check accounts directory
    If General_File_Exists(accounts_path, vbDirectory) = False Then
        'make the directory
        MkDir App.Path & PATH_ACCOUNTS
    End If
    
    'Load items
    Dim ini_path As String
    Dim LoopC As Long
    ini_path = scripts_path & "\item.ini"
    If General_File_Exists(ini_path, vbNormal) Then
        item_count = General_Var_Get(ini_path, "GENERAL", "item_count")
        If item_count > 0 Then
            ReDim item_list(1 To item_count)
            For LoopC = 1 To item_count
                item_list(LoopC).item_name = General_Var_Get(ini_path, "ITEM" & LoopC, "item_name")
                item_list(LoopC).item_grh = CLng(General_Var_Get(ini_path, "ITEM" & LoopC, "item_grh_index"))
            Next LoopC
        End If
    End If
    
    'Load speeches
    Dim speech As New clsSpeech
    ini_path = Dir(speech_path & "\*.spch", vbNormal)
    Do While ini_path <> ""
        speech.Initialize speech_path, ini_path
        speech_list.Add speech
        ini_path = Dir
    Loop
    
    'Load session vairables
    Session_Variables_Load
    
    'Scripting engine
    If Script_Engine_Initialize = False Then
        Exit Function
    End If
    
    
    '***********************
    'Script events
    '***********************
    'Server_Start_Up
    Dim command As New clsScriptCommand
    command.Initialize "Server_Start_Up"
    script_engine.Command_Add command
    '***********************
    
    'Load maps - should be the last thing done
    If Map_Load_All = False Then
        Exit Function
    End If
    
    'Start Tick Timer
    timTickCounter.Interval = SERVER_TICK_INTERVAL
    timTickCounter.Enabled = True
    
    'Server state
    server_state = s_s_listening
    
    'Log
    Log_Event "ctlDirectPlayServer", "Initialize", "Information - Server started ..."
    
    Initialize = True
Exit Function
ErrorHandler:
    Log_Event "ctlDirectPlayServer", "Initialize", "Unhandled Error - Number:" & Err.Number & " - Description: " & Err.Description
End Function

Public Function Deinitialize() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/25/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
On Error Resume Next
    Dim LoopC As Long
    Dim player As clsPlayer
    
    'Set status
    server_state = s_s_shutting_down
    
    'Save session variables
    Session_Variables_Save
    
    'Stop tick timer
    timTickCounter.Enabled = False

    'Deinit script engine
    Script_Engine_Deinitialize
    
    'Close sockets if needed
    If Not (dp_server Is Nothing) Then
        dp_server.CancelAsyncOperation 0, DPNCANCEL_ALL_OPERATIONS
        dp_server.Close
        dp_server.UnRegisterMessageHandler
        'Log
        Log_Event "ctlDirectPlayServer", "Deinitialize", "Information - Server stopped."
    End If
                
    'Destroy dx object
    Set dp_server = Nothing
    Set dp_server_address = Nothing
    Set dx = Nothing

    'Destroy lists
    Set player_list = Nothing
    Set npc_list = Nothing
    Set map_list = Nothing
    Set char_list = Nothing
    Set speech_list = Nothing

    'Reset variables
    server_connection_id = 0
    server_players_connection_id = 0
    server_ticks = 0
    
    'Set status to closed
    server_state = s_s_closed
    
    Deinitialize = True
Exit Function
ErrorHandler:
    Log_Event "ctlDirectPlayServer", "Deinitialize", "Unhandled Error - Number:" & Err.Number & " - Description: " & Err.Description
    Resume Next
End Function

Private Sub timTickCounter_Timer()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/05/2003
'
'**************************************************************
    Server_Tick
End Sub

Private Sub Server_Tick()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'Code that executes every server tick
'**************************************************************
    Dim LoopC As Long
    
    'Increment tick counter
    server_ticks = server_ticks + 1
    
    '***********************
    'Script events
    '***********************
    Dim command As New clsScriptCommand
    
    'Server_Tick
    command.Initialize "Server_Tick", 1
    command.Parameter_Set 1, server_ticks
    script_engine.Command_Add command
    Set command = Nothing
    
    'NPC AI
    For LoopC = npc_list.LowerBound To npc_list.UpperBound
        If npc_list.Item(LoopC).AiScript <> "" Then
            Set command = New clsScriptCommand
            command.Initialize npc_list.Item(LoopC).AiScript, 1
            command.Parameter_Set 1, LoopC
            script_engine.Command_Add command
            Set command = Nothing
        End If
    Next LoopC
    '***********************
    
    'Run all scripts batched this tick
    script_engine.Run_All
End Sub

Private Sub Receive_Chat(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/24/2004
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "Chat logged " & player.Name & ";" & player.Name_Char & ": " & parameters(1)
    'Player must be authenticated
    If player.AuthenticationStatus = p_as_none Then
        Exit Sub
    End If
    
    'Global Chat
    If command = c_Chat_Global Then
        If UBound(parameters()) = 1 Then
            'Check it's a GM
            If player.CharAuthenticationStatus > c_as_player Then
                Chat_To_All player.Name & " broadcasts: " & parameters(1)
            End If
        End If
        Exit Sub
    End If
    
    'Chat on the map only
    If command = c_Chat_Map Then
        If UBound(parameters()) = 1 Then
            Chat_To_Map player.MapID, player.Name & ": " & parameters(1)
        End If
        Exit Sub
    End If
    
    'Normal chatting
    If command = c_Chat_Normal Then
        If UBound(parameters()) = 1 Then
            Chat_Normal player_id, parameters(1)
        End If
        Exit Sub
    End If
    
    'Whisper
    If command = c_Chat_Whisper Then
        If UBound(parameters()) = 2 Then
            Chat_To_Player_Name player_id, player.Name_Char & ":" & parameters(2)
        End If
        Exit Sub
    End If
End Sub

Private Sub Receive_Move(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'Move
    If command = c_Move_Moved Then
        If player.Move_By_Heading(CLng(parameters(1))) Then
            'Movement OK
        Else
            'Movement not OK
            frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "Movement Error: " & player.Name & " , " & parameters(1) & ", " & player.MapX & " , " & player.MapY
        End If
        Exit Sub
    End If
End Sub

Private Sub Receive_Action(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 9/02/2004
'Modified Fredrik Alexandersson
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)

    'Chat with a NPC
    If command = c_Action_NPC_Chat Then
        'Check if we are starting a conversation or responding
        If UBound(parameters()) = 1 Then
            'Respond
            Send_Command to_id, player.ConnectionID, s_NPC, s_NPC_Chat, player.NPC_Talk_To_NPC(CLng(parameters(1)))
            Exit Sub
        ElseIf UBound(parameters()) = 2 Then
            'Start new conversation
            Dim temp_text As String
            temp_text = player.NPC_Start_Talk_To_NPC(CLng(parameters(1)), CLng(parameters(2)))
            If temp_text <> "" Then
                Send_Command to_id, player.ConnectionID, s_NPC, s_NPC_Chat, temp_text
            End If
            Exit Sub
        End If
    End If
    
    'Attack
    If command = c_Action_Attack Then
        player.Attack
        Exit Sub
    End If
    
    'Pick up an item.
    If command = c_Action_Item_Pickup Then
        player.Item_Pickup
        Exit Sub
    End If
    
    'Drop an item
    If command = c_Action_Item_Drop Then
        If UBound(parameters()) = 1 Then
            player.Item_Drop CLng(parameters(1))
        End If
        Exit Sub
    End If
    
    'Equip an item
    If command = c_Action_Item_Equip Then
        If UBound(parameters()) = 1 Then
            player.Item_Equip CLng(parameters(1))
        End If
        Exit Sub
    End If
End Sub

Public Sub Receive_Authenticate(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/20/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    frmcommand.txtlog = frmcommand.txtlog.text & vbNewLine & "Trying Auth of player"
    'Player must not already be authenticated to use these commands
    If player.AuthenticationStatus <> p_as_none Then
        Exit Sub
    End If
    
    'Login
    If command = c_Authenticate_Login Then
        If UBound(parameters()) = 1 Then
            'Set password
            player.password = parameters(1)
            'Try to login player
            Player_Login player_id
        End If
        Exit Sub
    End If
    
    'New
    If command = c_Authenticate_New Then
        If UBound(parameters()) = 5 Then
            'Set password
            player.password = parameters(1)
            'Try to login player
            Player_Login_New player_id, parameters(2), parameters(3), parameters(4), parameters(5)
        End If
        Exit Sub
    End If
    
    'Enter game using a char
    If command = c_Authenticate_Char Then
        If UBound(parameters()) = 1 Then
            'Check char is valid
            If Not player.Character_Check_Name(parameters(1)) Then
                'Char doesn´t belong to the Account
                Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Char name isn´t valid."
                frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "Char error logged"
            Else
                'Do it
                frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "Char logged in" & vbNewLine
                Char_Login player_id, parameters(1)
            End If
        End If
        Exit Sub
    End If
End Sub

Private Sub Player_Login(ByVal player_id As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/2/2003
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'Check name and password
    If Not player.Check_Account_Name_And_Password Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Account name does not exist or password does not match."
        Logoff_Account player_id
        Exit Sub
    End If
    
    'Send authentication confirmation to client, along with all account´s char´s names
    Send_Command to_id, player.ConnectionID, s_Player, s_Player_Authenticated, s_Packet_Character_Names(player.Get_Char_Name_From_Account(1), _
                        player.Get_Char_Name_From_Account(2), player.Get_Char_Name_From_Account(3), player.Get_Char_Name_From_Account(4))
End Sub

Private Sub Char_Login(ByVal PlayerID As Long, ByVal char_name As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/20/2004
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Dim LoopC As Long
    Dim Index As Long
    Dim Amount As Long
    Dim equiped As Boolean
    
    Set player = player_list.Item(PlayerID)
    
    'Load player info from file
    If player.Load_By_Name(char_name) = False Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Error loading player file."
        Logoff_Char PlayerID
        Exit Sub
    End If
    
    'Set to authenticated
    player.AuthenticationStatus = p_as_player
    
    'Give player a char_id
    Dim char As New clsChar
    char.PlayerID = player.id
    player.CharID = char_list.Add(char)
    
    'Add to Map
    If player.Map_Add(player.MapID, player.MapX, player.MapY) = False Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Error placing player on map."
        Logoff_Char PlayerID
        Exit Sub
    End If
    
    'Set player's general status to ready
    player.GeneralStatus = p_gs_ready
    
    'Tell the player it can start it's engine
    Send_Command to_id, player.ConnectionID, s_Player, s_Player_Engine_Start, ""
    
    'Send authentication confirmation to client
    Send_Command to_id, player.ConnectionID, s_Char, s_Char_Authenticated, ""
    
    'Send inventory items
    For LoopC = 0 To c_Player_Item_Slots
        player.Item_Get LoopC, Index, Amount, equiped
        If Index Then
            Send_Command to_id, player.ConnectionID, s_Char, s_Char_Set_Inventory_Slot, s_Packet_Inventory_Item(LoopC, Index, Amount, equiped)
        End If
    Next LoopC
    
    '***********************
    'Script events
    '***********************
    'Player_Login
    Dim command As New clsScriptCommand
    command.Initialize "Player_Login", 1
    command.Parameter_Set 1, PlayerID
    script_engine.Command_Add command
    '***********************
End Sub

Private Sub Player_Login_New(ByVal player_id As Long, ByVal Account_name As String, ByVal first_name As String, ByVal last_name As String, ByVal email As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/2/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'See if Account file already exists
    If player.Check_Account_Name(PATH_ACCOUNTS, Account_name) Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Account name already exists."
        Logoff_Account player_id
        Exit Sub
    End If
    
    'Load starting profile
    If Not player.Account_Create(Account_name, first_name, last_name, email) Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Invalid Account name."
        Logoff_Account player_id
        Exit Sub
    End If
    
    '***********************
    'Script events
    '***********************
    'Player_New
    Dim command As New clsScriptCommand
    command.Initialize "Player_New", 1
    command.Parameter_Set 1, player_id
    script_engine.Command_Add command
    '***********************
    
    'Login player
    Player_Login player_id
End Sub

Private Sub Logoff_Char(ByVal player_id As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/2/2003
'Edited by Jaun Martín Sotuyo Dodero
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'Check if player is fighting.
    If player.Fight <> -1 Then
        Send_Command to_id, player_id, s_Chat, s_Chat_Critical, "You will need to end the fight before loggin of."
        Exit Sub
    End If
    
    'Log off an authenticated player
    If player.AuthenticationStatus <> p_as_none Then
        'Save player
        player.Save_By_Name player.Name_Char
        
        'Remove from map
        player.Map_Remove
        
        'Remove char id
        If player.CharID Then
            char_list.Remove_Index player.CharID
        End If
    End If
    
    'Set authentication status
    player.AuthenticationStatus = p_as_none
    
    '***********************
    'Script events
    '***********************
    'Player_Logoff
    Dim command As New clsScriptCommand
    command.Initialize "Player_Off", 1
    command.Parameter_Set 1, player_id
    script_engine.Command_Add command
    '***********************
End Sub

Public Function Player_Map_Group_Add(ByVal s_player_id As Long, ByVal s_map_id As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(s_player_id)
    
   'Get map object
    Dim map As clsMap
    Set map = map_list.Item(s_map_id)
    
    'Add player to map group
    dp_server.AddPlayerToGroup map.ConnectionID, player.ConnectionID
    
    'Return true
    Player_Map_Group_Add = True
End Function

Public Function Player_Map_Group_Remove(ByVal s_player_id As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/2/2003
'
'**************************************************************
On Local Error Resume Next
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(s_player_id)

    'Get map object
    Dim map As clsMap
    Set map = map_list.Item(player.MapID)
    
    'Remove player from map group if needed
    If player.ConnectionStatus <> p_cs_disconnected And player.ConnectionStatus <> p_cs_none Then
        dp_server.RemovePlayerFromGroup map.ConnectionID, player.ConnectionID
    End If
    
    Player_Map_Group_Remove = True
End Function

Public Function NPC_Create(ByVal s_npc_data_index As Long, ByVal s_map As Long, ByVal s_map_x As Long, ByVal s_map_y As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Returns npc_id if success else 0
'**************************************************************
    'Create NPC object
    Dim npc_id As Long
    Dim new_npc As New clsNPC
    npc_id = npc_list.Add(new_npc)
    If new_npc.Initialize(Me, script_engine, map_list, player_list, npc_list, speech_list, npc_id, scripts_path & "\npc.ini") = False Then
        NPC_Remove npc_id
        Exit Function
    End If
    
    'Load npc data from ini file
    If new_npc.Load_From_Ini(s_npc_data_index) = False Then
        NPC_Remove npc_id
        Exit Function
    End If
    
    'Give npc a char_id
    Dim char As New clsChar
    char.NPCID = new_npc.id
    new_npc.CharID = char_list.Add(char)
    
    'Add to Map
    If new_npc.Map_Add(s_map, s_map_x, s_map_y) = False Then
        NPC_Remove npc_id
        Exit Function
    End If
    
    NPC_Create = npc_id
End Function

Public Function NPC_Remove(ByVal npc_id As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'
'**************************************************************
    'Get npc object
    Dim npc As clsNPC
    Set npc = npc_list.Item(npc_id)
    
    'Remove from map if needed
    If npc.MapID Then
        npc.Map_Remove
    End If
        
    'Remove char id needed
    If npc.CharID Then
        char_list.Remove_Index npc.CharID
    End If
        
    'Destroy npc object
    npc_list.Remove_Index npc_id
    
    NPC_Remove = True
End Function

Public Function Chat_To_All(ByVal s_message_string As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/24/2004
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'**************************************************************
    'Send Chat Packet
    Send_Command to_All, 0, s_Chat, s_Chat_Global, s_message_string
    Chat_To_All = True
End Function

Public Function Chat_To_Player_Name(ByVal s_player_name As String, ByVal s_message_string As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/24/2004
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Find("char_name", s_player_name)
    
    If player Is Nothing Then
        Exit Function
    End If
    
    'Send Chat Packet
    Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Whisper, s_message_string
    Chat_To_Player_Name = True
End Function

Public Function Chat_Normal(ByVal s_player_id As Long, ByVal s_message_string As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/24/2004
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(s_player_id)
    
    If player Is Nothing Then
        Exit Function
    End If
    
    Dim map As clsMap
    Set map = map_list.Item(player.MapID)
    
    'Send Chat Packet
    Send_Command to_id, map.ConnectionID, s_Chat, s_Chat_Normal, player.CharID & P_DELIMITER & s_message_string
    Chat_Normal = True
End Function

Public Function Chat_To_Player(ByVal s_player_id As Long, ByVal s_message_string As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/24/2004
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(s_player_id)
    
    If player Is Nothing Then
        Exit Function
    End If
    
    'Send Chat Packet
    Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Whisper, s_message_string
    Chat_To_Player = True
End Function

Public Function Chat_To_Map_Name(ByVal s_map_name As String, ByVal s_message_string As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/24/2004
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'**************************************************************
    'Get map object
    Dim map As clsMap
    Set map = map_list.Find("Name", s_map_name)
    If map Is Nothing Then
        Exit Function
    End If
    
    'Send Chat Packet
    Send_Command to_id, map.ConnectionID, s_Chat, s_Chat_Map, s_message_string
    Chat_To_Map_Name = True
End Function

Public Function Chat_To_Map(ByVal s_map_id As Long, ByVal s_message_string As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/24/2004
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'**************************************************************
    'Get map object
    Dim map As clsMap
    Set map = map_list.Item(s_map_id)
    If map Is Nothing Then
        Exit Function
    End If
    
    'Send Chat Packet
    Send_Command to_id, map.ConnectionID, s_Chat, s_Chat_Map, s_message_string
    Chat_To_Map = True
End Function

Public Function Send_Command(ByVal send_type As Command_Send_Type, ByVal connection_id As Long, ByVal header As ServerPacketHeader, ByVal command As ServerPacketCommand, ByRef parameters As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Send a command packet to client(s)
'**************************************************************
    Dim LoopC As Long

    'New packet
    Dim byte_buffer() As Byte
    Dim offset As Long
    offset = NewBuffer(byte_buffer)
    
    'Add header
    Call AddDataToBuffer(byte_buffer, header, SIZE_LONG, offset)
    
    'Add command
    Call AddDataToBuffer(byte_buffer, command, SIZE_LONG, offset)

    'Add parameters
    Call AddStringToBuffer(byte_buffer, parameters, offset)

    'To ID
    If send_type = to_id Then
        'Send the packet
        dp_server.SendTo connection_id, byte_buffer, 0, DPNSEND_GUARANTEED Or DPNSEND_NOLOOPBACK
        Exit Function
    End If
    
    'To All
    If send_type = to_All Then
        'Send the packet
        dp_server.SendTo server_players_connection_id, byte_buffer, 0, DPNSEND_GUARANTEED Or DPNSEND_NOLOOPBACK
        Exit Function
    End If
    
    Send_Command = True
End Function

Public Function Send_Char_Create(ByVal send_type As Command_Send_Type, ByVal s_connection_id As Long, ByVal s_char_id As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/23/2004
'Sends a series of commands to create a char to client(s)
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    'Get player or npc object
    Dim char As Object
    If char_list.Item(s_char_id).PlayerID Then
        Set char = player_list.Item(char_list.Item(s_char_id).PlayerID)
    Else
        Set char = npc_list.Item(char_list.Item(s_char_id).NPCID)
    End If

    'Create
    Send_Command send_type, s_connection_id, s_Char, s_Char_Create, _
        s_Packet_Char_Create(char.CharID, char.MapX, char.MapY, char.Heading, char.CharDataIndex)
    'Label
    Send_Command send_type, s_connection_id, s_Char, s_Char_Label_Set, _
        s_Packet_Char_Label_Set(char.CharID, char.Name_Char, 1)
    
    Send_Char_Create = True
End Function

Private Function Map_Load_All() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Load all maps in the map folder
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
On Error GoTo ErrorHandler:
    Dim map As clsMap
    Dim map_id As Long
    Dim map_name As String
    Dim LoopC As Long
    Dim group_info As DPN_GROUP_INFO
    
    'Load all the maps in the PATH_MAPS directory
    map_name = Dir(maps_path & "\*.map", vbNormal)
    
    Do Until map_name = ""
        'Create object, add to list, and initialize
        Set map = New clsMap
        map_id = map_list.Add(map)
        If map.Initialize(Me, script_engine, map_list, char_list, npc_list, map_id, Left$(map_name, Len(map_name) - 4), maps_path) = False Then
            Exit Function
        End If
        
        'Try to load map file
        If map.Load_By_Name(map_name) Then
            'Create map connection group
            group_info.lInfoFlags = DPNINFO_NAME
            group_info.Name = "MAP " & Left$(map_name, Len(map_name) - 4)
            dp_server.CreateGroup group_info
        Else
            'Remove from list
            map_list.Remove_Index map_id
        End If
        'Get next map name
        map_name = Dir
    Loop
    
    Set map = Nothing
    
    'Load ini files
    For LoopC = 1 To map_list.Count
        Set map = map_list.Item(LoopC)
        If Not (map Is Nothing) Then
            map.Load_Ini_By_Name map.Name
        End If
    Next LoopC
    Set map = Nothing
    
    Map_Load_All = True
Exit Function
ErrorHandler:
    Log_Event "ctlDirectPlayServer", "Map_Load_All", "Unhandled Error - Number:" & Err.Number & " - Description: " & Err.Description
End Function

Private Function Script_Engine_Initialize() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/7/2003
'
'**************************************************************
On Error GoTo ErrorHandler:
    'Initialize Scripting System
    Dim check As Boolean
    If script_engine.Initialize = False Then check = True
    script_gump.Initialize Me, gump_list
    If script_interface.Initialize(Me, script_engine, map_list, player_list, npc_list, char_list, gump_list) = False Then check = True
    If script_engine.Object_Add("ORE", script_interface) = False Then check = True
    If script_engine.Object_Add("GUMP", script_gump) = False Then check = True
    If script_engine.Load_From_File(scripts_path & "\main.vbs") = False Then check = True
    If script_engine.Load_From_File(scripts_path & "\tile_events.vbs") = False Then check = True
    If script_engine.Load_From_File(scripts_path & "\npc_ai.vbs") = False Then check = True
    If script_engine.Load_From_File(scripts_path & "\gump.vbs") = False Then check = True
    If check = False Then
        Script_Engine_Initialize = True
    Else
        Log_Event "clsDirectPlayServer", "Script_Engine_Initialize", "Error - Description: Error initializing script engine. Check script log for details."
    End If

Exit Function
ErrorHandler:
    Log_Event "ctlDirectPlayServer", "Script_Engine_Initialize", "Unhandled Error - Number:" & Err.Number & " - Description: " & Err.Description
End Function

Private Function Script_Engine_Deinitialize() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/7/2003
'
'**************************************************************
    script_engine.Deinitialize
    script_interface.Deinitialize
    
    Script_Engine_Deinitialize = True
End Function

Public Function Script_Engine_Reset() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/7/2003
'Unload and reload all scripts in the engine
'**************************************************************
    'Log
    Log_Event "ctlDirectPlayServer", "Script_Engine_Reset", "Information - Reset started ..."
    
    Script_Engine_Deinitialize
    If Script_Engine_Initialize Then
        Script_Engine_Reset = True
        Log_Event "ctlDirectPlayServer", "Script_Engine_Reset", "Information - Reset completed successfully."
    Else
        Log_Event "ctlDirectPlayServer", "Script_Engine_Reset", "Error - Reset failed."
    End If
    

End Function

Public Function Session_Variable_Create(ByVal s_variable_name As String, ByVal s_variable_data As Variant, ByVal s_variable_save As Boolean) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'
'*****************************************************************
    If Session_Variable_Check(s_variable_name) Then
        Exit Function
    End If

    ReDim Preserve session_variable_list(0 To UBound(session_variable_list) + 1)
    
    session_variable_list(UBound(session_variable_list)).variable_name = s_variable_name
    session_variable_list(UBound(session_variable_list)).variable_data = s_variable_data
    session_variable_list(UBound(session_variable_list)).variable_save = s_variable_save
    
    Session_Variable_Create = True
End Function

Public Function Session_Variable_Get(ByVal s_variable_name As String) As Variant
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'
'*****************************************************************
    Dim LoopC As Long
    
    For LoopC = LBound(session_variable_list) To UBound(session_variable_list)
        If session_variable_list(LoopC).variable_name = s_variable_name Then
            Session_Variable_Get = session_variable_list(LoopC).variable_data
            Exit Function
        End If
    Next LoopC
End Function

Public Function Session_Variable_Check(ByVal s_variable_name As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'
'*****************************************************************
    Dim LoopC As Long
    
    For LoopC = LBound(session_variable_list) To UBound(session_variable_list)
        If session_variable_list(LoopC).variable_name = s_variable_name Then
            Session_Variable_Check = True
            Exit Function
        End If
    Next LoopC
End Function

Public Function Session_Variable_Set(ByVal s_variable_name As String, ByVal s_variable_data) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'
'*****************************************************************
    Dim LoopC As Long
    
    For LoopC = LBound(session_variable_list) To UBound(session_variable_list)
        If session_variable_list(LoopC).variable_name = s_variable_name Then
             session_variable_list(LoopC).variable_data = s_variable_data
             Session_Variable_Set = True
            Exit Function
        End If
    Next LoopC
    
    Session_Variable_Set = False
End Function

Public Function Session_Variables_Save() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'
'*****************************************************************
    Dim LoopC As Long
    Dim counter  As Long
    Dim file_path As String
    
    file_path = App.Path & "\" & "session.ini"
        
    'SESSION
    counter = 1
    If UBound(session_variable_list()) <> 0 Then
        For LoopC = 1 To UBound(session_variable_list())
            If session_variable_list(LoopC).variable_save Then
                General_Var_Write file_path, "SESSION", CStr(counter), CStr(session_variable_list(LoopC).variable_name) & "-" & CStr(session_variable_list(LoopC).variable_data)
                counter = counter + 1
            End If
        Next LoopC
    End If
    General_Var_Write file_path, "SESSION", "count", CStr(counter - 1)
    
    Session_Variables_Save = True
End Function

Public Function Session_Variables_Load() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'
'*****************************************************************
    Dim LoopC As Long
    Dim t_count As Long
    Dim temp_string As String
    Dim file_path As String
    
    file_path = App.Path & "\" & "session.ini"
        
     'SESSION
    ReDim session_variable_list(0) As Session_Variable
    t_count = Val(General_Var_Get(file_path, "SESSION", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "SESSION", CStr(LoopC))
        If temp_string <> "" Then
            Session_Variable_Create General_Field_Read(1, temp_string, Asc("-")), CVar(General_Field_Read(2, temp_string, Asc("-"))), True
        End If
    Next LoopC
    
    Session_Variables_Load = True
End Function

Public Sub Log_Event(ByVal source_class As String, ByVal source_procedure As String, ByVal event_string As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/05/2003
'
'*****************************************************************
    Open App.Path & "\log_server.txt" For Append As #40
    Print #40, CStr(DateTime.Now) & " - " & source_class & " - " & source_procedure & " - " & event_string
    Close #40
End Sub

Private Sub Receive_Request(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/4/2004
'
'*****************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'Roll stats
    If command = c_Request_Roll_Stats Then
        Roll_Stats player.id
        Exit Sub
    End If
    
    'Get char stats
    If command = c_Request_Char_Stats Then
        If UBound(parameters()) = 1 Then
            Char_Get_Stats player.id, parameters(1)
        End If
        Exit Sub
    End If
End Sub

Private Sub Roll_Stats(ByVal player_id As Long)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'*****************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'Store total of points given to make sure player doesn´t cheat
    player.Stats_Total = Stats_Roll()
    
    'Send back stats
    Send_Command to_id, player.ConnectionID, s_Char, s_Char_Stats_Rolled, CStr(player.Stats_Total)
End Sub

Private Function Stats_Roll()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'*****************************************************************
    Dim result As Long
    Dim points As Long
    
    result = Round(General_Random_Number(1, 100), 0)
    
    Select Case result
        'See how many point correspond to the obtained number
        Case 1 To 40
            points = 66
        Case 41 To 70
            points = 72
        Case 71 To 86
            points = 78
        Case 87 To 96
            points = 86
        Case 97 To 99
            points = 92
        Case 100
            points = 98
    End Select
    
    Stats_Roll = points
End Function

Private Sub Char_Get_Stats(ByVal player_id As Long, ByVal char_name As String)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    If Not player.Character_Check_Name(char_name) Then
        'Char doesn´t exist in the player´s Account
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "The char " & char_name & " doesn´t exist in the current Account."
    End If
    
    If Not General_File_Exists(players_path & "\" & char_name & ".ini", vbNormal) Then
        'Char doesn´t exist in the player´s Account
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "The char " & char_name & " doesn´t exist in the current Account."
    End If
    
    'Send back the char´s stats
    Send_Command to_id, player.ConnectionID, s_Char, s_Char_Stats_Get, player.Char_Get_Stats(char_name)
End Sub

Private Sub Receive_Logoff(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'*****************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & player.Name & "Attemps Logoffcall"
    'Log Out from Account
    If command = c_Logoff_Account Then
        frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "call 1"
        Logoff_Account player_id
        Exit Sub
    End If
    
    'Log out a char from game
    If command = c_Logoff_Char Then
        Logoff_Char player_id
        frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "call 2"
        Exit Sub
    End If
    
    'Log out and finish session
    If command = c_Logoff_Session_Terminate Then
        frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "call 3"
        Logoff_Session_Terminate player_id
        Exit Sub
    End If
End Sub

Private Sub Receive_Gump(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'*****************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & player.Name & ";" & player.Name_Char & ": calls gump"
    'Send Gump
    If command = c_Gump_Page Then
        'Setsup the gump
        script_gump.Clear_Gump
        script_engine.Execute "gump_" & parameters(1)
        'send the gump
        Send_Command to_id, player.ConnectionID, S_Gump, S_Gump_page, script_gump.Compile_Gump
        Exit Sub
    End If
    If command = c_Gump_button Then
        script_engine.Execute "button_" & parameters(1) & "_" & parameters(2)
        Exit Sub
    End If
End Sub

Private Sub Receive_Sysop(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'*****************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Get player object
    Dim player As clsPlayer
    
    Set player = player_list.Item(player_id)
    
    General_Write_To_TextBox frmcommand.txtlog, player.Name & ";" & player.Name_Char & ": Sysop called"
    
    If player.CharAuthenticationStatus > c_as_player Then
        If command = c_Sysop_NewItem Then
            Dim map As clsMap
            Set map = map_list.Item(player.MapID)
            
            General_Write_To_TextBox frmcommand.txtlog, "Call NewItem"
            
            'Log event
            Log_Event "ctlDirectPlayServer", "Receive_Sysop", player.Name & ";" & player.Name_Char & ": Sysop create item " & parameters(1)
            
            map.Item_Create player.MapX, player.MapY, CLng(parameters(1)), 1
            Exit Sub
        End If
        
        If command = c_Sysop_Ban Then
            Dim acc As clsAccount
            acc.Account_Ban parameters(1), parameters(2)
            
            General_Write_To_TextBox frmcommand.txtlog, "Call Ban : " & parameters(1)
            
            'Log event
            Log_Event "ctlDirectPlayServer", "Receive_Sysop", player.Name & ";" & player.Name_Char & ": Sysop Ban Account " & parameters(1) & " " & parameters(2)
            Exit Sub
        End If
        
        If command = c_Sysop_Saveworld Then
            WorldSave
            
            General_Write_To_TextBox frmcommand.txtlog, "Call Saveworld"
            
            'Log event
            Log_Event "ctlDirectPlayServer", "Receive_Sysop", player.Name & ";" & player.Name_Char & ": Sysop World Save"
            Exit Sub
        End If
        
        If command = c_Sysop_Reset Then
            frmMain.Reset_Server
            
            General_Write_To_TextBox frmcommand.txtlog, "Call Reset Server"
            
            'Log event
            Log_Event "ctlDirectPlayServer", "Receive_Sysop", player.Name & ";" & player.Name_Char & ": Sysop Reset Server"
            Exit Sub
        End If
        
        If command = c_Sysop_Shutdown Then
            General_Write_To_TextBox frmcommand.txtlog, "Call Shutdown"
            
            'Log event
            Log_Event "ctlDirectPlayServer", "Receive_Sysop", player.Name & ";" & player.Name_Char & ": Sysop Shutdown Server"
            
            frmMain.Shutdown_Server
            Exit Sub
        End If
        
        If command = c_Sysop_Goto Then
            player.Map_Remove
            player.Map_Add map_list.Find_Index("map_name", parameters(1)), parameters(2), parameters(3)
            
            General_Write_To_TextBox frmcommand.txtlog, "Call Goto"
            
            'Log event
            Log_Event "ctlDirectPlayServer", "Receive_Sysop", player.Name & ";" & player.Name_Char & ": Sysop Goto " & parameters(1) & " - " & parameters(2) & " - " & parameters(3)
            Exit Sub
        End If
        
        If command = c_Sysop_Summon Then
            'Get player id
            Dim splayer As clsPlayer
            Set splayer = player_list.Find("Name_Char", parameters(1))
            
            player.Map_Remove
            player.Map_Add splayer.MapID, player.MapX + 1, player.MapY + 1
            
            General_Write_To_TextBox frmcommand.txtlog, "Call Summon"
            
            'Log event
            Log_Event "ctlDirectPlayServer", "Receive_Sysop", player.Name & ";" & player.Name_Char & ": Sysop Summon " & parameters(1)
            Exit Sub
        End If
    End If
End Sub

Private Sub Logoff_Account(ByVal player_id As Long)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'*****************************************************************
On Error Resume Next
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)

    'Log off the char first
    Logoff_Char player_id

    player.Account_Logoff
End Sub

Private Sub Logoff_Session_Terminate(ByVal player_id As Long)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
'We set this eeror handler to prevent the server from crushing if the client crushed
On Local Error Resume Next
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    Logoff_Account player_id
    
    'Destroy player object
    player_list.Remove_Index player_id
    
    'Add player to players connection group
    dp_server.RemovePlayerFromGroup server_players_connection_id, player.ConnectionID
    
    'Disconnect player if already didn't happen
    If player.ConnectionStatus <> p_cs_disconnected Then
        'Disconnect player
        dp_server.DestroyClient player.ConnectionID, 0, 0, 0
    End If
    
    'Throw Event
    RaiseEvent ServerConnectionRemoved(player.ConnectionID)
End Sub

Private Sub Receive_Account(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'*****************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'Create a char
    If command = c_Account_Add_Char Then
        'Check char name
        If UBound(parameters()) = 16 Then
            If Not General_String_Is_Alphanumeric(parameters(1)) Then
                'Boot player
                Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Invalid character found in player name."
                Exit Sub
            End If
            
            If player.Character_Create_Check_Name(parameters(1)) Then
                If Not player.Character_Create(parameters(1), CLng(parameters(2)), CLng(parameters(3)), CLng(parameters(4)), CLng(parameters(5)), _
                  CLng(parameters(6)), CLng(parameters(8)), CLng(parameters(9)), CLng(parameters(10)), CLng(parameters(11)), CLng(parameters(12)), _
                  CLng(parameters(13)), CLng(parameters(14)), CLng(parameters(15)), CLng(parameters(16))) Then
                    'Send message to the client. The char wasn´t valid. Reasons: a - You did something wrong in the client. b - The user is a cheater and edited the client himself.
                    Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "The char stats aren´t correct."
                     frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "Char create called and Failed 2"
                Else
                    'Save char
                    player.Save_By_Name player.Name_Char
                    'Log char into the game
                    Char_Login player_id, player.Name_Char
                    frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "Char create called and passed"
                End If
            Else
                'Char name already exists
                Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "A character already exists with the given name. Please choose another."
                frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "Char create called and Failed 1"
            End If
        End If
        
        Exit Sub
    End If
    
    'Delete a char
    If command = c_Account_Remove_Char Then
        If UBound(parameters()) = 1 Then
            player.Character_Remove CLng(parameters(1))
        End If
        Exit Sub
    End If
End Sub

Public Sub WorldSave()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/14/2004
'
'*****************************************************************
    Dim map As clsMap
    Dim player As clsPlayer
    Dim LoopC As Long
    
    For LoopC = map_list.LowerBound To map_list.UpperBound
        Set map = map_list.Item(LoopC)
        map.Save_Ini_By_Name (App.Path & "\" & map.Name & ".ini")
    Next LoopC
    
    'Save all chars
    For LoopC = player_list.LowerBound To player_list.UpperBound
        If player.AuthenticationStatus <> p_as_none Then
            player.Save_By_Name player.Name_Char
        End If
    Next LoopC
End Sub
