VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsPlayer.cls - ORE Server Player Object - v0.5.0
'
'Player object.
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
Option Explicit

'***************************
'Constants
'***************************
Public Enum PlayerConnectionStatus
    p_cs_none = 0
    p_cs_disconnected = 1
    p_cs_connected = 2
End Enum

Public Enum PlayerAuthenticationStatus
    p_as_none = 0
    p_as_player = 1
    p_as_guide = 2
    p_as_gm = 3
End Enum

Public Enum PlayerGeneralStatus
    p_gs_none = 0
    p_gs_loading = 1
    p_gs_ready = 2
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
Private dps As ctlDirectPlayServer
Private script_engine As clsScriptEngine
Private map_list As clsList
Private player_list As clsList
Private npc_list As clsList

Private players_path As String

Private player_id As Long
Private player_connection_id As Long
Private player_name As String
Private player_decription As String
Private player_password As String
Private player_char_id As Long
    
Private player_connection_status As PlayerConnectionStatus
Private player_authentication_status As PlayerAuthenticationStatus
Private player_general_status As PlayerGeneralStatus
    
Private player_map_id As Long
Private player_map_x As Long
Private player_map_y As Long
Private player_heading As Long

Private player_char_data_index As Long

Private session_variable_list() As Session_Variable

Private Sub Class_Initialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
End Sub

Private Sub Class_Terminate()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    Deinitialize
End Sub


Public Property Get ID() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
        ID = player_id
End Property

Public Property Get Name() As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
        Name = player_name
End Property

Public Property Let Name(vNewValue As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
        player_name = vNewValue
End Property

Public Property Get Name_Upper_Case() As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
        Name_Upper_Case = UCase$(player_name)
End Property

Public Property Get ConnectionID() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    ConnectionID = player_connection_id
End Property

Public Property Let ConnectionID(ByVal vNewValue As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    player_connection_id = vNewValue
End Property

Public Property Get CharID() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    CharID = player_char_id
End Property

Public Property Let CharID(ByVal vNewValue As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    player_char_id = vNewValue
End Property

Public Property Get password() As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    password = player_password
End Property

Public Property Let password(ByVal vNewValue As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
     player_password = vNewValue
End Property

Public Property Get ConnectionStatus() As PlayerConnectionStatus
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    ConnectionStatus = player_connection_status
End Property

Public Property Let ConnectionStatus(ByVal vNewValue As PlayerConnectionStatus)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    player_connection_status = vNewValue
End Property

Public Property Get AuthenticationStatus() As PlayerAuthenticationStatus
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    AuthenticationStatus = player_authentication_status
End Property

Public Property Let AuthenticationStatus(ByVal vNewValue As PlayerAuthenticationStatus)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    player_authentication_status = vNewValue
End Property

Public Property Get GeneralStatus() As PlayerGeneralStatus
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    GeneralStatus = player_general_status
End Property

Public Property Let GeneralStatus(ByVal vNewValue As PlayerGeneralStatus)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    player_general_status = vNewValue
End Property

Public Property Get MapID() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    MapID = player_map_id
End Property

Public Property Get MapX() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    MapX = player_map_x
End Property

Public Property Get MapY() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    MapY = player_map_y
End Property

Public Property Get Heading() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    Heading = player_heading
End Property

Public Property Let Heading(ByVal vNewValue As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Change the heading of the player
'**************************************************************
    'Check heading
    If vNewValue < 1 Or vNewValue > 8 Then
        Exit Property
    End If
    'Set
    player_heading = vNewValue
   'Send packet
    dps.Send_Command to_id, map_list.Item(player_map_id).ConnectionID, s_Char, s_Char_Heading_Set, _
        s_Packet_Char_Heading_Set(player_char_id, player_heading)
End Property

Public Property Get CharDataIndex() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    CharDataIndex = player_char_data_index
End Property

Public Property Let CharDataIndex(ByVal vNewValue As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    player_char_data_index = vNewValue
End Property

Private Sub Convert_Heading_to_Direction(ByVal s_heading As Long, ByRef r_map_x As Long, ByRef r_map_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Applies a heading to a map position
'**************************************************************
    Dim addY As Long
    Dim addX As Long
    
    'Figure out which way to move
    Select Case s_heading
    
        Case 1
            addY = -1
    
        Case 2
            addY = -1
            addX = 1
    
        Case 3
            addX = 1
            
        Case 4
            addX = 1
            addY = 1
    
        Case 5
            addY = 1
        
        Case 6
            addX = -1
            addY = 1
        
        Case 7
            addX = -1
            
        Case 8
            addX = -1
            addY = -1
            
    End Select
    
    r_map_x = r_map_x + addX
    r_map_y = r_map_y + addY
End Sub

Private Function Convert_Direction_To_Heading(ByVal s_direction_x As Long, ByVal s_direction_y As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Takes a direction and figures out a heading
'**************************************************************
    'North
    If s_direction_x = 0 And s_direction_y = -1 Then
        Convert_Direction_To_Heading = 1
    End If
    'NE
    If s_direction_x = 1 And s_direction_y = -1 Then
        Convert_Direction_To_Heading = 2
    End If
    'East
    If s_direction_x = 1 And s_direction_y = 0 Then
        Convert_Direction_To_Heading = 3
    End If
    'SE
    If s_direction_x = 1 And s_direction_y = 1 Then
        Convert_Direction_To_Heading = 4
    End If
    'South
    If s_direction_x = 0 And s_direction_y = 1 Then
        Convert_Direction_To_Heading = 5
    End If
    'SW
    If s_direction_x = -1 And s_direction_y = 1 Then
        Convert_Direction_To_Heading = 6
    End If
    'West
    If s_direction_x = -1 And s_direction_y = 0 Then
        Convert_Direction_To_Heading = 7
    End If
    'NW
    If s_direction_x = -1 And s_direction_y = -1 Then
        Convert_Direction_To_Heading = 8
    End If
End Function

Public Function Initialize(ByRef s_dps As ctlDirectPlayServer, ByRef s_script_engine As clsScriptEngine, ByRef s_map_list As clsList, ByRef s_player_list As clsList, ByRef s_npc_list As clsList, ByVal s_id As Long, Optional ByVal s_name As String, Optional ByVal s_players_path As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/28/2003
'
'*****************************************************************
    Set dps = s_dps
    Set script_engine = s_script_engine
    Set map_list = s_map_list
    Set player_list = s_player_list
    Set npc_list = s_npc_list
    player_id = s_id
    player_name = s_name
    players_path = s_players_path
    
    ReDim session_variable_list(0) As Session_Variable
    
    Initialize = True
End Function

Public Function Deinitialize() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/28/2003
'
'*****************************************************************
    Deinitialize = True
End Function

Public Function Save_By_Name(s_name As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Save a player by a name
'*****************************************************************
    Save_By_Name = Save_To_File(players_path & "\" & s_name & ".ini")
End Function

Public Function Save_To_File(ByVal file_path As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/07/2003
'Save a player to file
'*****************************************************************
    Dim loopc As Long
    Dim counter  As Long

    'See if file exists
    If General_File_Exists(file_path, vbNormal) Then
        'Kill it
        Kill file_path
    End If

    'Get Map Object
    Dim map As clsMap
    Set map = map_list.Item(player_map_id)

    '************
    'Save
    '************
    
    'GENERAL
    General_Var_Write file_path, "GENERAL", "password", player_password
    General_Var_Write file_path, "GENERAL", "map_name", map.Name
    General_Var_Write file_path, "GENERAL", "map_x", CStr(player_map_x)
    General_Var_Write file_path, "GENERAL", "map_y", CStr(player_map_y)
    General_Var_Write file_path, "GENERAL", "heading", CStr(player_heading)
    General_Var_Write file_path, "GENERAL", "char_data_index", CStr(player_char_data_index)
    
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

    Save_To_File = True
End Function

Public Function Load_By_Name(s_name As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Load a player by a name
'*****************************************************************
    Load_By_Name = Load_From_File(players_path & "\" & s_name & ".ini")
End Function

Public Function Load_From_File(ByVal file_path As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/7/2003
'Load a player from file
'*****************************************************************
    Dim loopc As Long
    Dim t_count As Long
    Dim temp_string As String

    'See if file exists
    If General_File_Exists(file_path, vbNormal) = False Then
        Load_From_File = False
        Exit Function
    End If
    
    '************
    'Load
    '************
    
    'GENERAL
    'Get Map Object
    Dim map As clsMap
    Set map = map_list.Find("Name", General_Var_Get(file_path, "GENERAL", "map_name"))
    player_map_id = map.ID
    player_map_x = General_Var_Get(file_path, "GENERAL", "map_x")
    player_map_y = General_Var_Get(file_path, "GENERAL", "map_y")
    player_heading = General_Var_Get(file_path, "GENERAL", "heading")
    player_char_data_index = General_Var_Get(file_path, "GENERAL", "char_data_index")
    
    'SESSION
    ReDim session_variable_list(0) As Session_Variable
    t_count = Val(General_Var_Get(file_path, "SESSION", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "SESSION", CStr(loopc))
        If temp_string <> "" Then
            Session_Variable_Create General_Field_Read(1, temp_string, Asc("-")), CVar(General_Field_Read(2, temp_string, Asc("-"))), True
        End If
    Next loopc
    
    Load_From_File = True
End Function

Public Function Check_Name_And_Password() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'*****************************************************************
    Dim file_path As String

    file_path = players_path & "\" & player_name & ".ini"

    'See if file exists
    If General_File_Exists(file_path, vbNormal) = False Then
        Exit Function
    End If
    
    If player_password <> General_Var_Get(file_path, "GENERAL", "password") Then
        Exit Function
    End If
    
    Check_Name_And_Password = True
End Function

Public Function Check_Name() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'*****************************************************************
    Dim file_path As String

    file_path = players_path & "\" & player_name & ".ini"

    'See if file exists
    If General_File_Exists(file_path, vbNormal) = False Then
        Exit Function
    End If
    
    Check_Name = True
End Function

Public Function Move_By_Heading(ByVal s_heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Move player along a heading
'*****************************************************************
    Dim x As Long
    Dim y As Long
    Dim nx As Long
    Dim ny As Long
    
    'Get Map Object
    Dim map As clsMap
    Set map = map_list.Item(player_map_id)
    
    'Try to move player
    If map.Char_Move_By_Heading(player_map_x, player_map_y, s_heading) Then
        'Moved
        'Update player position
        x = player_map_x
        y = player_map_y
        nx = x
        ny = y
        Convert_Heading_to_Direction s_heading, nx, ny
        player_map_x = nx
        player_map_y = ny
        player_heading = s_heading
        'Send move command to entire map
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Move, _
            s_Packet_Char_Move(player_char_id, s_heading)
            Move_By_Heading = True
        
        '***********************
        'Script events
        '***********************
        'Player_Move_By_Heading
        Dim command As New clsScriptCommand
        command.Initialize "Player_Move", 2
        command.Parameter_Set 1, player_id
        command.Parameter_Set 2, s_heading
        script_engine.Command_Add command
        '***********************
    Else
        'Didn't move
        'Force player to the proper position
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Pos_Set, _
            s_Packet_Char_Pos_Set(player_char_id, player_map_x, player_map_y)
        Move_By_Heading = False
    End If
End Function

Public Function Attack() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Have player attack
'*****************************************************************
    'Get Map Object
    Dim map As clsMap
    Set map = map_list.Item(player_map_id)

    'Send attack animation command
    dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Data_Body_Set, _
        s_Packet_Char_Data_Body_Set(player_char_id, 3, True)
        
    'TODO: Actually do something besides play the animation
    Attack = True
End Function

Public Function Position_Set(ByVal s_map_id As Long, ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'Change the position of the player
'**************************************************************
   'Get map object
    Dim map As clsMap
    Set map = map_list.Item(s_map_id)
    
    'See if new player's position is legal
    If map.Char_Position_Legal(s_map_x, s_map_y) = False Then
        'TODO: Recalculate player position finding a near by legal one
        Exit Function
    End If

    'See if we are changing maps
    If player_map_id <> s_map_id Then
        'Remove from old map
        If Map_Remove = False Then Exit Function
        'Add to new map
        If Map_Add(s_map_id, s_map_x, s_map_y) = False Then Exit Function
    Else
        'Erase char from old tile
        Dim old_tile As clsMapTile
        Set old_tile = map.Tile_Get(player_map_x, player_map_y)
        old_tile.CharID = 0
        'Update player pos
        player_map_x = s_map_x
        player_map_y = s_map_y
        'Add char to new tile
        Dim new_tile As clsMapTile
        Set new_tile = map.Tile_Get(player_map_x, player_map_y)
        new_tile.CharID = player_char_id
        'Send command to clients on map
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Pos_Set, _
            s_Packet_Char_Pos_Set(player_char_id, player_map_x, player_map_y)
    End If
    
    Position_Set = True
End Function

Public Function Map_Add(ByVal s_map_id As Long, ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Add a player to a map
'**************************************************************
   'Get map object
    Dim map As clsMap
    Set map = map_list.Item(s_map_id)
    
    'See if new player's position is legal
    If map.Char_Position_Legal(s_map_x, s_map_y) Then
        'Add map to player
        player_map_id = s_map_id
        player_map_x = s_map_x
        player_map_y = s_map_y
    Else
        Map_Add = False
        Exit Function
    End If
    
    'Tell player to load a map
    dps.Send_Command to_id, player_connection_id, s_Map, s_Map_Load, map.Name
    
    'Add player to map
    Dim tile As clsMapTile
    Set tile = map.Tile_Get(player_map_x, player_map_y)
    tile.CharID = player_char_id
    
    'Send all player chars, currently on the map, to new player
    Dim loopc As Long
    Dim tempplayer As clsPlayer
    For loopc = 1 To player_list.UpperBound
        Set tempplayer = player_list.Item(loopc)
        If Not (tempplayer Is Nothing) Then
            If tempplayer.MapID = player_map_id Then
                dps.Send_Char_Create to_id, player_connection_id, tempplayer.CharID
            End If
        End If
        Set tempplayer = Nothing
    Next loopc
    
    'Send all npc chars, currently on the map, to new player
    Dim tempnpc As clsNPC
    For loopc = 1 To npc_list.UpperBound
        Set tempnpc = npc_list.Item(loopc)
        If Not (tempnpc Is Nothing) Then
            If tempnpc.MapID = player_map_id Then
                dps.Send_Char_Create to_id, player_connection_id, tempnpc.CharID
            End If
        End If
        Set tempplayer = Nothing
    Next loopc
    
    'Set new player's char id
    dps.Send_Command to_id, player_connection_id, s_Char, s_Char_ID_Set, CStr(player_char_id)
    
    'Send the new player char to everyone else on the map
    'Create
    dps.Send_Char_Create to_id, map.ConnectionID, player_char_id
    
    'Add player to map connection group
    dps.Player_Map_Group_Add player_id, s_map_id
    
    'Return true
    Map_Add = True
End Function

Public Function Map_Remove() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Remove a player from a map
'**************************************************************
    'Get map object
    Dim map As clsMap
    Set map = map_list.Item(player_map_id)
    
    'Remove player from map
    Dim tile As clsMapTile
    Set tile = map.Tile_Get(player_map_x, player_map_y)
    tile.CharID = 0
    
    'Remove player from map connection group
    dps.Player_Map_Group_Remove player_id
    
    'Update everyone else on map
    If dps.ServerStatus = s_s_listening Then
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Remove, CStr(player_char_id)
    End If
    
    'Remove map from player
    player_map_id = 0
    player_map_x = 0
    player_map_y = 0
    
    Map_Remove = True
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
