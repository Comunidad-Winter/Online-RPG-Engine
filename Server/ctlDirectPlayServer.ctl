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

Private script_engine As New clsScriptEngine
Private script_interface As New clsScriptInterface

'***************************
'Arrays
'***************************
Dim player_list As clsList
Dim npc_list As clsList
Dim map_list As clsList
Dim char_list As clsList

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
    If Left(group_info.Name, 3) = "MAP" Then
        'Save it to map object
        Dim ID As Long
        ID = CLng(Mid$(group_info.Name, 4, Len(group_info.Name)))
        Dim map As clsMap
        Set map = map_list.Find("ID", ID)
        map.ConnectionID = group_id
        Exit Sub
    End If
End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal connection_id As Long, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/2/2003
'
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
          Send_Command to_id, connection_id, s_Chat, s_Chat_Critical, "Invalid character found in player name."
          dp_server.DestroyClient connection_id, 0, 0, 0
          Exit Sub
    End If

    'See if there is already a player with the same name
    Dim tempplayer As clsPlayer
    Set tempplayer = player_list.Find("Name_Upper_Case", UCase$(peer_info.Name)) 'Use upper case so we are sure there isn't a match
    If Not (tempplayer Is Nothing) Then
        'Boot player
        Send_Command to_id, connection_id, s_Chat, s_Chat_Critical, "Player name is already logged on."
        dp_server.DestroyClient connection_id, 0, 0, 0
        Exit Sub
    End If
    
    'Initialize player object and add to player list
    Dim player_id As Long
    Dim newplayer As New clsPlayer
    player_id = player_list.Add(newplayer)
    newplayer.Initialize Me, script_engine, map_list, player_list, npc_list, player_id, peer_info.Name, players_path
    newplayer.ConnectionID = connection_id
    newplayer.ConnectionStatus = p_cs_connected
 
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
'
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
        'Set to disconnected
         player.ConnectionStatus = p_cs_disconnected
        'Logoff player properly
        Player_Logoff player.ID
    End If
    
    'Throw Event
    RaiseEvent ServerConnectionRemoved(connection_id)
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
    Dim loopc As Long
    Dim Count As Long
    received_data = GetStringFromBuffer(dpnotify.ReceivedData, offset)
    Count = General_Field_Count(received_data, P_DELIMITER_CODE)
    If Count > 0 Then
        ReDim parameters(1 To Count) As String
        For loopc = 1 To Count
            parameters(loopc) = General_Field_Read(loopc, received_data, P_DELIMITER_CODE)
        Next loopc
    Else
        ReDim parameters(0 To 0) As String
    End If
    
    'Handle the packet
    If header = c_Authenticate Then
        Receive_Authenticate player.ID, command, parameters()
        Exit Sub
    End If
    
    If header = c_Chat Then
        Receive_Chat player.ID, command, parameters()
        Exit Sub
    End If
    
    If header = c_Move Then
        Receive_Move player.ID, command, parameters()
        Exit Sub
    End If

    If header = c_Action Then
        Receive_Action player.ID, command, parameters()
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
'
'**************************************************************
On Error GoTo ErrorHandler

    'Paths
    resource_path = s_resource_path
    players_path = resource_path & PATH_PLAYERS
    maps_path = resource_path & PATH_MAPS
    scripts_path = resource_path & PATH_SCRIPTS
        
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
'Last Modify Date: 5/07/2003
'
'**************************************************************
On Error Resume Next
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
    Dim loopc As Long
    
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
    For loopc = npc_list.LowerBound To npc_list.UpperBound
        If npc_list.Item(loopc).AiScript <> "" Then
            Set command = New clsScriptCommand
            command.Initialize npc_list.Item(loopc).AiScript, 1
            command.Parameter_Set 1, loopc
            script_engine.Command_Add command
            Set command = Nothing
        End If
    Next loopc
    '***********************
    
    'Run all scripts batched this tick
    script_engine.Run_All
End Sub

Private Sub Receive_Chat(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'Player must be authenticated
    If player.AuthenticationStatus = p_as_none Then
        Exit Sub
    End If
    
    'Global Chat
    If command = c_Chat_Global Then
        If UBound(parameters()) = 1 Then
            Chat_To_All player.Name & " broadcasts: " & parameters(1)
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
            Debug.Print "Movement Error: " & player.Name & " , " & parameters(1) & ", " & player.MapX & " , " & player.MapY
        End If
        Exit Sub
    End If
End Sub

Private Sub Receive_Action(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)

    'Attack
    If command = c_Action_Attack Then
        player.Attack
        Exit Sub
    End If
End Sub

Private Sub Receive_Authenticate(ByVal player_id As Long, ByVal command As ClientPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
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
        If UBound(parameters()) = 2 Then
            'Set password
            player.password = parameters(1)
            'Try to login player
            Player_Login_New player_id, parameters(2)
        End If
        Exit Sub
    End If
End Sub

Private Sub Player_Login(ByVal player_id As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/2/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'Check name and password
    If player.Check_Name_And_Password = False Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Player name does not exist or password does not match."
        Player_Logoff player_id
        Exit Sub
    End If
    
    'Load player info from file
    If player.Load_By_Name(player.Name) = False Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Error loading player file."
        Player_Logoff player_id
        Exit Sub
    End If
    
    'Add player to players connection group
    dp_server.AddPlayerToGroup server_players_connection_id, player.ConnectionID
    
    'Set to authenticated
    player.AuthenticationStatus = p_as_player
    
    'Send authentication confirmation to client
    Send_Command to_id, player.ConnectionID, s_Player, s_Player_Authenticated, ""
    
    'Give player a char_id
    Dim char As New clsChar
    char.PlayerID = player.ID
    player.CharID = char_list.Add(char)
    
    'Add to Map
    If player.Map_Add(player.MapID, player.MapX, player.MapY) = False Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Error placing player on map."
        Player_Logoff player_id
        Exit Sub
    End If
    
    'Set player's general status to ready
    player.GeneralStatus = p_gs_ready
    
    'Tell the player it can start it's engine
    Send_Command to_id, player.ConnectionID, s_Player, s_Player_Engine_Start, ""
    
    '***********************
    'Script events
    '***********************
    'Player_Login
    Dim command As New clsScriptCommand
    command.Initialize "Player_Login", 1
    command.Parameter_Set 1, player_id
    script_engine.Command_Add command
    '***********************
End Sub

Private Sub Player_Login_New(ByVal player_id As Long, ByVal profile_name As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/2/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'See if player file already exists
    If player.Check_Name Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Player name already exists."
        Player_Logoff player_id
        Exit Sub
    End If
    
    'Load starting profile
    If player.Load_By_Name("_" & profile_name) = False Then
        Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Critical, "Invalid profile name."
        Player_Logoff player_id
        Exit Sub
    End If
    
    'Save player
    player.Save_By_Name player.Name
    
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

Private Sub Player_Logoff(ByVal player_id As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/2/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(player_id)
    
    'Log off an authenticated player
    If player.AuthenticationStatus <> p_as_none Then
        'Save player
        player.Save_By_Name player.Name
        
        'Remove from map
        player.Map_Remove
    End If

    'Disconnect player if already didn't happen
    If player.ConnectionStatus <> p_cs_disconnected Then
        'Disconnect player
        dp_server.DestroyClient player.ConnectionID, 0, 0, 0
    End If
        
    'Remove char id
    If player.CharID Then
        char_list.Remove_Index player.CharID
    End If
        
    'Destroy player object
    player_list.Remove_Index player_id
    
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

Public Function NPC_Create(ByVal s_npc_data_index As Long, ByVal s_map_id As Long, ByVal s_map_x As Long, ByVal s_map_y As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Returns npc_id if success else 0
'**************************************************************
    'Create NPC object
    Dim npc_id As Long
    Dim new_npc As New clsNPC
    npc_id = npc_list.Add(new_npc)
    If new_npc.Initialize(Me, script_engine, map_list, player_list, npc_list, npc_id, scripts_path & "\npc.ini") = False Then
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
    char.NPCID = new_npc.ID
    new_npc.CharID = char_list.Add(char)
    
    'Add to Map
    If new_npc.Map_Add(s_map_id, s_map_x, s_map_y) = False Then
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
'Last Modify Date: 5/6/2003
'
'**************************************************************
    'Send Chat Packet
    Send_Command to_All, 0, s_Chat, s_Chat_Text, s_message_string
    Chat_To_All = True
End Function

Public Function Chat_To_Player(ByVal s_player_id As Long, ByVal s_message_string As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    'Get player object
    Dim player As clsPlayer
    Set player = player_list.Item(s_player_id)
    If player Is Nothing Then
        Exit Function
    End If
    
    'Send Chat Packet
    Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Text, s_message_string
    Chat_To_Player = True
End Function

Public Function Chat_To_Map_Name(ByVal s_map_name As String, ByVal s_message_string As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    'Get map object
    Dim map As clsMap
    Set map = map_list.Find("Name", s_map_name)
    If map Is Nothing Then
        Exit Function
    End If
    
    'Send Chat Packet
    Send_Command to_id, map.ConnectionID, s_Chat, s_Chat_Text, s_message_string
    Chat_To_Map_Name = True
End Function

Public Function Chat_To_Map(ByVal s_map_id As Long, ByVal s_message_string As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    'Get map object
    Dim map As clsMap
    Set map = map_list.Item(s_map_id)
    If map Is Nothing Then
        Exit Function
    End If
    
    'Send Chat Packet
    Send_Command to_id, map.ConnectionID, s_Chat, s_Chat_Text, s_message_string
    Chat_To_Map = True
End Function

Public Function Send_Command(ByVal send_type As Command_Send_Type, ByVal connection_id As Long, ByVal header As ServerPacketHeader, ByVal command As ServerPacketCommand, ByRef parameters As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Send a command packet to client(s)
'**************************************************************
    Dim loopc As Long

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
'Last Modify Date: 5/12/2003
'Sends a series of commands to create a char to client(s)
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
        s_Packet_Char_Label_Set(char.CharID, char.Name, 1)
    'Heading
    Send_Command send_type, s_connection_id, s_Char, s_Char_Heading_Set, _
        s_Packet_Char_Heading_Set(char.CharID, char.Heading)
        
    Send_Char_Create = True
End Function

Private Function Map_Load_All() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Load all maps in the map folder
'**************************************************************
On Error GoTo ErrorHandler:
    Dim map As clsMap
    Dim map_id As Long
    Dim loopc As Long
    Dim group_info As DPN_GROUP_INFO
    
    'Load all the map in the PATH_MAPS directory
    'TODO: Make map size adjustable
    For loopc = 1 To 999
        'Create object, add to list, and initialize
        Set map = New clsMap
        map_id = map_list.Add(map)
        If map.Initialize(Me, script_engine, map_list, char_list, map_id, CStr(loopc), maps_path) = False Then
            Exit Function
        End If
        
        'Try to load map file
        If map.Load_By_Name(CStr(loopc)) Then
            'Create map connection group
            group_info.lInfoFlags = DPNINFO_NAME
            group_info.Name = "MAP" & loopc
            dp_server.CreateGroup group_info
        Else
            'Remove from list
            map_list.Remove_Index map_id
        End If
    Next loopc
    Set map = Nothing
    
    'Load ini files
    For loopc = 1 To map_list.Count
        Set map = map_list.Item(loopc)
        If Not (map Is Nothing) Then
            map.Load_Ini_By_Name map.Name
        End If
    Next loopc
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
    If script_interface.Initialize(Me, script_engine, map_list, player_list, npc_list, char_list) = False Then check = True
    If script_engine.Object_Add("ORE", script_interface) = False Then check = True
    If script_engine.Load_From_File(scripts_path & "\main.vbs") = False Then check = True
    If script_engine.Load_From_File(scripts_path & "\tile_events.vbs") = False Then check = True
    If script_engine.Load_From_File(scripts_path & "\npc_ai.vbs") = False Then check = True
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
    Dim loopc As Long
    
    For loopc = LBound(session_variable_list) To UBound(session_variable_list)
        If session_variable_list(loopc).variable_name = s_variable_name Then
            Session_Variable_Get = session_variable_list(loopc).variable_data
            Exit Function
        End If
    Next loopc
End Function

Public Function Session_Variable_Check(ByVal s_variable_name As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'
'*****************************************************************
    Dim loopc As Long
    
    For loopc = LBound(session_variable_list) To UBound(session_variable_list)
        If session_variable_list(loopc).variable_name = s_variable_name Then
            Session_Variable_Check = True
            Exit Function
        End If
    Next loopc
End Function

Public Function Session_Variable_Set(ByVal s_variable_name As String, ByVal s_variable_data) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'
'*****************************************************************
    Dim loopc As Long
    
    For loopc = LBound(session_variable_list) To UBound(session_variable_list)
        If session_variable_list(loopc).variable_name = s_variable_name Then
             session_variable_list(loopc).variable_data = s_variable_data
             Session_Variable_Set = True
            Exit Function
        End If
    Next loopc
    
    Session_Variable_Set = False
End Function

Public Function Session_Variables_Save() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'
'*****************************************************************
    Dim loopc As Long
    Dim counter  As Long
    Dim file_path As String
    
    file_path = App.Path & "\" & "session.ini"
        
    'SESSION
    counter = 1
    If UBound(session_variable_list()) <> 0 Then
        For loopc = 1 To UBound(session_variable_list())
            If session_variable_list(loopc).variable_save Then
                General_Var_Write file_path, "SESSION", CStr(counter), CStr(session_variable_list(loopc).variable_name) & "-" & CStr(session_variable_list(loopc).variable_data)
                counter = counter + 1
            End If
        Next loopc
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
    Dim loopc As Long
    Dim t_count As Long
    Dim temp_string As String
    Dim file_path As String
    
    file_path = App.Path & "\" & "session.ini"
        
     'SESSION
    ReDim session_variable_list(0) As Session_Variable
    t_count = Val(General_Var_Get(file_path, "SESSION", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "SESSION", CStr(loopc))
        If temp_string <> "" Then
            Session_Variable_Create General_Field_Read(1, temp_string, Asc("-")), CVar(General_Field_Read(2, temp_string, Asc("-"))), True
        End If
    Next loopc
    
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

