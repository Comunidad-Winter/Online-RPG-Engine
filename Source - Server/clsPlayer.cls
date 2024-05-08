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
'           XX/XX/200X - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'David Justus - 14/8/2004
'   - Add: Sysop levels (1 = Player)
'   - Add: Item methods
'   - Add: Inventory
'   Subrelease contributors:
'      Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 9/02/2004
'       - Add: Item_Equip method
'       - Add: CharAuthenticationStatus properties
'       - Add: CharStatus enum
'       - Add: Item type
'       - Add: Targeting method
'       - Add: Attack method
'       - Add: All Level_Up methods
'       - Change: Item methods now use the new item type
'       - Change: Load_Stats_From_File and Save_Stats_To_File to store new vars
'
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 2/12/2004
'   - Add: All NPC speech related stuff
'   - Add: All Account stuff
'   - Add: Stats saving and loading methods
'   - Add: All character creation related stuff
'   - Add: All stats related stuff
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
End Enum

Public Enum PlayerGeneralStatus
    p_gs_none = 0
    p_gs_loading = 1
    p_gs_ready = 2
End Enum

Public Enum CharStatus
    c_as_player
    c_as_guide
    c_as_gm
End Enum

'***************************
'Types
'***************************
Private Type Session_Variable
    variable_name As String
    variable_data As Variant
    variable_save As Boolean
End Type

Private Type Item
    item_index As Long
    Amount As Long
    equiped As Boolean
End Type

'***************************
'Variables
'***************************
Private dps As ctlDirectPlayServer
Private script_engine As clsScriptEngine
Private map_list As clsList
Private player_list As clsList
Private npc_list As clsList
Private speech_list As clsList

Private players_path As String
Private resource_path As String

Private player_password As String

Private player_id As Long
Private player_connection_id As Long
Private player_name As String
Private player_decription As String
Private player_char_id As Long

Private char_connection_status As PlayerConnectionStatus
Private player_authentication_status As PlayerAuthenticationStatus
Private player_general_status As PlayerGeneralStatus
Private char_authentication_status As CharStatus

Private char_map_id As Long
Private char_map_x As Long
Private char_map_y As Long
Private char_heading As Long
Private char_in_fight As Long
Private char_talking_to_NPC As Long 'Keeps the NPC index to which it´s talking
Private NPC_greet_index As Long     'Keeps the last greet index used by the NPC

'***********
'Stats
'***********
Private char_name As String
Private char_race As races
Private char_class As classes
Private char_alignment As Alignment
Private char_sphere As spheres
Private char_psionic_power As psionic_powers
Private char_level As Long
Private char_portrait As Long

Private char_hp As Long
Private char_hp_max As Long
Private char_mp As Long
Private char_mp_max As Long
Private char_exp As Long
Private char_exp_next As Long
Private char_STR As Long
Private char_DEX As Long
Private char_CON As Long
Private char_INT As Long
Private char_WIS As Long
Private char_CHR As Long
Private player_target As String
Private player_target_npc As String

Private char_inventory(c_Player_Item_Slots) As Item
Private char_weapon_equipped_slot As Long

'Used to check stats are alright, and user is not cheating
Private char_stats_total As Long

'The player´s Account object
Private Account As New clsAccount

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


Public Property Get id() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
        id = player_id
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

Public Property Get ConnectionStatus() As PlayerConnectionStatus
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    ConnectionStatus = char_connection_status
End Property

Public Property Let ConnectionStatus(ByVal vNewValue As PlayerConnectionStatus)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    char_connection_status = vNewValue
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
    MapID = char_map_id
End Property

Public Property Get HP() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    HP = char_hp
End Property

Public Property Let HP(ByVal nNewValue As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    char_hp = nNewValue
End Property

Public Property Get HP_max() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    HP_max = char_hp_max
End Property

Public Property Get MP() As Long
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 1/12/2004
'
'**************************************************************
    MP = char_mp
End Property

Public Property Get MP_Max() As Long
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 1/12/2004
'
'**************************************************************
    MP_Max = char_mp_max
End Property

Public Property Get Charisma() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    Charisma = char_CHR
End Property

Public Property Get Dextery() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    Dextery = char_DEX
End Property

Public Property Get Wisdom() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    Wisdom = char_WIS
End Property

Public Property Get Intelligence() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    Intelligence = char_INT
End Property

Public Property Get Strength() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    Strength = char_STR
End Property

Public Property Get MapX() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    MapX = char_map_x
End Property

Public Property Get MapY() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    MapY = char_map_y
End Property

Public Property Get Heading() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    Heading = char_heading
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
    char_heading = vNewValue
   'Send packet
    dps.Send_Command to_id, map_list.Item(char_map_id).ConnectionID, s_Char, s_Char_Heading_Set, _
        s_Packet_Char_Heading_Set(player_char_id, char_heading)
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

Public Function Initialize(ByRef s_dps As ctlDirectPlayServer, ByRef s_script_engine As clsScriptEngine, ByRef s_map_list As clsList, ByRef s_player_list As clsList, ByRef s_npc_list As clsList, ByVal s_speech_list As clsList, ByVal s_id As Long, ByVal Account_path As String, Optional ByVal s_name As String, Optional ByVal s_players_path As String, Optional ByVal s_resource_path As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/28/2003
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    Set dps = s_dps
    Set script_engine = s_script_engine
    Set map_list = s_map_list
    Set player_list = s_player_list
    Set npc_list = s_npc_list
    Set speech_list = s_speech_list
    player_id = s_id
    player_name = s_name
    players_path = s_players_path
    resource_path = s_resource_path
    
    Account.Account_Initialize Account_path
    
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
'Last Modify Date: 8/14/2004
'Save a player to file
'Edited by Juan Martín Sotuyo Dodero
'*****************************************************************
    Dim LoopC As Long
    Dim counter  As Long

    'See if file exists
    If General_File_Exists(file_path, vbNormal) Then
        'Kill it
        Kill file_path
    End If

    'Get Map Object
    Dim map As clsMap
    Set map = map_list.Item(char_map_id)

    '************
    'Save
    '************
    
    'GENERAL
    General_Var_Write file_path, "GENERAL", "map_name", map.Name
    General_Var_Write file_path, "GENERAL", "map_x", char_map_x
    General_Var_Write file_path, "GENERAL", "map_y", char_map_y
    General_Var_Write file_path, "GENERAL", "heading", char_heading
    General_Var_Write file_path, "GENERAL", "char_data_index", CStr(player_char_data_index)
    
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

    'ITEM
    For LoopC = 0 To c_Player_Item_Slots
        General_Var_Write file_path, "ITEMS", LoopC & "itm", char_inventory(LoopC).item_index
        General_Var_Write file_path, "ITEMS", LoopC & "amt", char_inventory(LoopC).Amount
        General_Var_Write file_path, "ITEMS", LoopC & "equ", CStr(char_inventory(LoopC).equiped)
    Next LoopC
    
    Save_Stats_To_File file_path
    
    Save_To_File = True
End Function

Public Function Load_By_Name(ByVal s_name As String) As Boolean
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
'Last Modify Date: 8/14/2004
'Load a player from file
'Edited by Juan Martín Sotuyo Dodero
'*****************************************************************
    Dim LoopC As Long
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
    'Set name
    Dim temp_name As String
    temp_name = General_Field_Read(General_Field_Count(file_path, 92), file_path, 92)
    char_name = Left$(temp_name, Len(temp_name) - 4)
    
    'Get Map Object
    Dim map As clsMap
    Set map = map_list.Find("Name", General_Var_Get(file_path, "GENERAL", "map_name"))
    char_map_id = map.id
    char_map_x = Val(General_Var_Get(file_path, "GENERAL", "map_x"))
    char_map_y = Val(General_Var_Get(file_path, "GENERAL", "map_y"))
    char_heading = Val(General_Var_Get(file_path, "GENERAL", "heading"))
    player_char_data_index = Val(General_Var_Get(file_path, "GENERAL", "char_data_index"))
    
    'SESSION
    ReDim session_variable_list(0) As Session_Variable
    t_count = Val(General_Var_Get(file_path, "SESSION", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "SESSION", CStr(LoopC))
        If temp_string <> "" Then
            Session_Variable_Create General_Field_Read(1, temp_string, Asc("-")), CVar(General_Field_Read(2, temp_string, Asc("-"))), True
        End If
    Next LoopC
    
    'ITEM
    For LoopC = 0 To c_Player_Item_Slots
        char_inventory(LoopC).item_index = Val(General_Var_Get(file_path, "ITEMS", LoopC & "itm"))
        char_inventory(LoopC).Amount = Val(General_Var_Get(file_path, "ITEMS", LoopC & "amt"))
        char_inventory(LoopC).equiped = CBool(General_Var_Get(file_path, "ITEMS", LoopC & "equ"))
    Next LoopC
    
    Load_Stats_From_File file_path
    
    Load_From_File = True
End Function

Private Function Load_Stats_By_Name(ByVal char_name As String) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/4/2004
'Load a player by a name
'*****************************************************************
    Load_Stats_By_Name = Load_Stats_From_File(players_path & "\" & char_name & ".ini")
End Function

Private Function Load_Stats_From_File(ByVal file_path As String) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    'STATS
    char_race = Val(General_Var_Get(file_path, "STATS", "race"))
    char_class = Val(General_Var_Get(file_path, "STATS", "class"))
    char_alignment = Val(General_Var_Get(file_path, "STATS", "align"))
    char_sphere = Val(General_Var_Get(file_path, "STATS", "sphere"))
    char_psionic_power = Val(General_Var_Get(file_path, "STATS", "psionic"))
    char_level = Val(General_Var_Get(file_path, "STATS", "level"))
    char_portrait = Val(General_Var_Get(file_path, "STATS", "portrait"))
    
    char_hp = Val(General_Var_Get(file_path, "STATS", "hp"))
    char_hp_max = Val(General_Var_Get(file_path, "STATS", "hp_max"))
    char_mp = Val(General_Var_Get(file_path, "STATS", "mp"))
    char_mp_max = Val(General_Var_Get(file_path, "STATS", "mp_max"))
    char_exp = Val(General_Var_Get(file_path, "STATS", "exp"))
    char_exp_next = Val(General_Var_Get(file_path, "STATS", "exp_next"))
    char_STR = Val(General_Var_Get(file_path, "STATS", "str"))
    char_DEX = Val(General_Var_Get(file_path, "STATS", "dex"))
    char_CON = Val(General_Var_Get(file_path, "STATS", "con"))
    char_INT = Val(General_Var_Get(file_path, "STATS", "int"))
    char_WIS = Val(General_Var_Get(file_path, "STATS", "wis"))
    char_CHR = Val(General_Var_Get(file_path, "STATS", "cha"))
    
    char_authentication_status = CInt(General_Var_Get(file_path, "STATS", "Sysop"))
    
    char_weapon_equipped_slot = Val(General_Var_Get(file_path, "STATS", "EquippedWeaponSlot"))
    
    Load_Stats_From_File = True
End Function

Public Function Check_Account_Name_And_Password() As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'*****************************************************************
    Check_Account_Name_And_Password = Account.Account_Login(Name, password)
End Function

Public Function Check_Account_Name(ByVal accounts_path As String, ByVal Account_name As String) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/6/2004
'Returns true if an Account doesn´t exist, and is threrefore free to be created
'*****************************************************************
    Check_Account_Name = Not General_File_Exists(accounts_path & "\" & Account_name & ".acc", vbNormal)
End Function

Public Function Character_Check_Name(ByVal char_name As String) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/4/2004
'Checks if a char exists in the player´s Account
'*****************************************************************
    Character_Check_Name = Account.Account_Character_Check(char_name)
End Function

Public Function Move_By_Heading(ByVal s_heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/7/2003
'Move player along a heading
'*****************************************************************
    Dim X As Long
    Dim Y As Long
    Dim nx As Long
    Dim ny As Long
    
    'Get Map Object
    Dim map As clsMap
    Set map = map_list.Item(char_map_id)
    
    'Try to move player
    If map.Char_Move_By_Heading(char_map_x, char_map_y, s_heading) Then
        'Moved
        'Update player position
        X = char_map_x
        Y = char_map_y
        nx = X
        ny = Y
        Convert_Heading_to_Direction s_heading, nx, ny
        char_map_x = nx
        char_map_y = ny
        char_heading = s_heading
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
            s_Packet_Char_Pos_Set(player_char_id, char_map_x, char_map_y)
        Move_By_Heading = False
    End If
End Function

Public Property Get Fight() As Long
    If char_in_fight > 0 Then
        Fight = char_in_fight
        Exit Property
    End If
    Fight = -1
End Property

Public Property Let Fight(ByVal fight_id As Long)
        char_in_fight = fight_id
End Property

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
    If char_map_id <> s_map_id Then
        'Remove from old map
        If Map_Remove = False Then Exit Function
        'Add to new map
        If Map_Add(s_map_id, s_map_x, s_map_y) = False Then Exit Function
    Else
        'Erase char from old tile
        Dim old_tile As clsMapTile
        Set old_tile = map.Tile_Get(char_map_x, char_map_y)
        old_tile.CharID = 0
        'Update player pos
        char_map_x = s_map_x
        char_map_y = s_map_y
        'Add char to new tile
        Dim new_tile As clsMapTile
        Set new_tile = map.Tile_Get(char_map_x, char_map_y)
        new_tile.CharID = player_char_id
        'Send command to clients on map
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Pos_Set, _
            s_Packet_Char_Pos_Set(player_char_id, char_map_x, char_map_y)
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
        char_map_id = s_map_id
        char_map_x = s_map_x
        char_map_y = s_map_y
    Else
        Map_Add = False
        Exit Function
    End If
    
    'Tell player to load a map
    dps.Send_Command to_id, player_connection_id, s_map, s_Map_Load, map.Name
    
    'Add player to map
    Dim tile As clsMapTile
    Set tile = map.Tile_Get(char_map_x, char_map_y)
    tile.CharID = player_char_id
    
    'Send all player chars, currently on the map, to new player
    Dim LoopC As Long
    Dim tempplayer As clsPlayer
    For LoopC = 1 To player_list.UpperBound
        Set tempplayer = player_list.Item(LoopC)
        If Not (tempplayer Is Nothing) Then
            If tempplayer.MapID = char_map_id Then
                dps.Send_Char_Create to_id, player_connection_id, tempplayer.CharID
            End If
        End If
        Set tempplayer = Nothing
    Next LoopC
    
    'Send all npc chars, currently on the map, to new player
    Dim tempnpc As clsNPC
    For LoopC = 1 To npc_list.UpperBound
        Set tempnpc = npc_list.Item(LoopC)
        If Not (tempnpc Is Nothing) Then
            If tempnpc.MapID = char_map_id Then
                dps.Send_Char_Create to_id, player_connection_id, tempnpc.CharID
            End If
        End If
        Set tempplayer = Nothing
    Next LoopC
    
    'Sends all the items to player.
    Dim LoopD As Long
    For LoopC = map.Grid_X_Min To map.Grid_X_Max
        For LoopD = map.Grid_Y_Min To map.Grid_Y_Max
            Set tile = map.Tile_Get(LoopC, LoopD)
            If tile.ItemAmount Then
                dps.Send_Command to_id, player_connection_id, s_map, s_Map_Item_Add, s_Packet_Item_Create(LoopC, LoopD, tile.ItemAmount, tile.ItemDataIndex)
            End If
        Next LoopD
    Next LoopC
    
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
    Set map = map_list.Item(char_map_id)
    
    'Remove player from map
    Dim tile As clsMapTile
    Set tile = map.Tile_Get(char_map_x, char_map_y)
    tile.CharID = 0
    
    'Remove player from map connection group
    dps.Player_Map_Group_Remove player_id
    
    'Update everyone else on map
    If dps.ServerStatus = s_s_listening Then
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Remove, CStr(player_char_id)
    End If
    
    'Remove map from player
    char_map_id = 0
    char_map_x = 0
    char_map_y = 0
    
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

Public Function Item_Move(ByVal item_old_slot As Long, ByVal item_new_slot As Long) As Boolean
'*****************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 8/14/2004
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Make sure new slor is empty
    If char_inventory(item_new_slot).item_index Then Exit Function
    
    char_inventory(item_old_slot) = char_inventory(item_new_slot)
    char_inventory(item_old_slot).Amount = 0
    char_inventory(item_old_slot).item_index = 0
    char_inventory(item_old_slot).equiped = False
    
    Item_Move = True
End Function

Public Function Item_Get(ByVal Item_Slot As Long, ByRef Item_Data_Index As Long, ByRef Item_Amount As Long, ByRef Item_Equiped As Boolean) As Boolean
'*****************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 8/14/2004
'modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    If Item_Slot > c_Player_Item_Slots Or Item_Slot < 0 Then Exit Function
    
    Item_Data_Index = char_inventory(Item_Slot).item_index
    Item_Amount = char_inventory(Item_Slot).Amount
    Item_Equiped = char_inventory(Item_Slot).equiped
    
    Item_Get = True
End Function

Public Sub Item_Pickup()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/14/2004
'
'*****************************************************************
    Dim map As clsMap
    Dim tile As clsMapTile
    Dim LoopC As Long
    Dim parameters As String
    
    Set map = map_list.Item(char_map_id)
    Set tile = map.Tile_Get(char_map_x, char_map_y)
    
    For LoopC = 0 To c_Player_Item_Slots
        If char_inventory(LoopC).item_index = tile.ItemDataIndex Then
            If char_inventory(LoopC).Amount > 100 - tile.ItemAmount Then
                char_inventory(LoopC).Amount = char_inventory(LoopC).Amount + tile.ItemAmount
                tile.ItemAmount = 0
                tile.ItemDataIndex = 0
                dps.Send_Command to_All, player_connection_id, s_map, s_Map_Item_Remove, Generic_Packet_Map_Pos(char_map_x, char_map_y)
                Exit Sub
            Else
                tile.ItemAmount = tile.ItemAmount - (99 - char_inventory(LoopC).Amount)
                char_inventory(LoopC).Amount = 99
                'Keep looping until we add the rest of the items on the floor
            End If
        ElseIf char_inventory(LoopC).item_index = 0 Then
            char_inventory(LoopC).item_index = tile.ItemDataIndex
            char_inventory(LoopC).Amount = tile.ItemAmount
            char_inventory(LoopC).equiped = False
            tile.ItemAmount = 0
            tile.ItemDataIndex = 0
            dps.Send_Command to_All, player_connection_id, s_map, s_Map_Item_Remove, Generic_Packet_Map_Pos(char_map_x, char_map_y)
            Exit Sub
        End If
    Next LoopC
End Sub

Public Function Item_Drop(ByVal Item_Slot As Long) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/20/2004
'Drops the item in the given slot to the floor
'*****************************************************************
    If Item_Slot > c_Player_Item_Slots Or Item_Slot < 0 Then Exit Function
    
    'Check there is something in the slot
    If char_inventory(Item_Slot).item_index = 0 Then Exit Function
    
    Dim map As clsMap
    
    map.Item_Create char_map_x, char_map_y, char_inventory(Item_Slot).item_index, char_inventory(Item_Slot).Amount
    char_inventory(Item_Slot).item_index = 0
    
    Item_Drop = True
End Function

Public Function Char_Get_Stats(ByVal char_name As String) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/4/2004
'Retrieves a packet containing all  the char´s stats
'*****************************************************************
    Load_Stats_By_Name char_name
    
    Char_Get_Stats = s_Packet_Character_Stats(char_name, char_race, char_class, char_alignment, char_sphere, char_psionic_power, char_level, _
                                                char_STR, char_DEX, char_CON, char_INT, char_WIS, char_CHR, char_portrait, player_char_data_index)
End Function

Public Function Get_Char_Name_From_Account(ByVal slot As Long) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/4/2004
'Retrieves the names of a char in the current Account
'*****************************************************************
    Get_Char_Name_From_Account = Account.Account_Character_Get_Name(slot)
End Function

Public Sub Account_Logoff()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'*****************************************************************
    Account.Account_Logoff
End Sub

Public Property Get Logged_In() As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'*****************************************************************
    Logged_In = Account.LoggedIn
End Property

Public Function Character_Remove(ByVal slot As Long) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'*****************************************************************
    Dim char_name As String
    
    'Remove a char
    char_name = Account.Account_Character_Get_Name(slot)
    
    'Make sure it´s not empty
    If char_name <> "" Then
        'Remove it from Account
        Account.Account_Character_Remove slot
        'Delete the char file
        Kill players_path & "\" & char_name & ".ini"
        
        Character_Remove = True
        Exit Function
    End If
    
    Character_Remove = False
End Function

Public Function Character_Create(ByVal Name As String, ByVal race As Long, ByVal Class As Long, ByVal align As Long, ByVal sphere As Long, ByVal psionic_power As Long, ByVal new_STR As Long, ByVal new_DEX As Long, ByVal new_CON As Long, ByVal new_INT As Long, ByVal new_WIS As Long, ByVal new_CHR As Long, ByVal Portrait As Long, ByVal char_data_index As Long, ByVal slot As Long) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/11/2004
'
'*****************************************************************
    Dim CharCreation As New clsCharacterCreation
    
    'Initialize bject
    CharCreation.Initialize
    
    'Check the stats total wasn´t changed by the user
    If Not char_stats_total = new_STR + new_DEX + new_CON + new_INT + new_WIS + new_CHR Then
        'User is cheating!!!!
        Character_Create = False
        Exit Function
    End If
    
    'Check if the portrait is correct according to the char race
    If InStr(1, General_Var_Get(resource_path & "\Portraits.ini", "GENERAL", Portrait), race, vbTextCompare) = 0 Then
        Character_Create = False
        Exit Function
    End If
    
    'Set all stats
    CharCreation.Race_Set race
    CharCreation.Class_Set Class
    CharCreation.Alignment_Set align
    CharCreation.Sphere_Set sphere
    CharCreation.Psionic_Power_Set psionic_power
    CharCreation.Char_STR_Set new_STR
    CharCreation.Char_DEX_Set new_DEX
    CharCreation.Char_CON_Set new_CON
    CharCreation.Char_INT_Set new_INT
    CharCreation.Char_WIS_Set new_WIS
    CharCreation.Char_CHR_Set new_CHR
    
    'Set the level up expirience required according to race
    Select Case race
        Case Gladiator
            char_exp_next = 2250
        
        Case Fighter
            char_exp_next = 2000
        
        Case Thief
            char_exp_next = 1250
        
        Case Ranger
            char_exp_next = 2250
        
        Case Psionicist
            char_exp_next = 2200
        
        Case Preserver
            char_exp_next = 2500
        
        Case Defiler
            'char_exp_next
        
        Case Cleric
            char_exp_next = 1500
        
        Case Druid
            char_exp_next = 2000
    End Select
    
    'Check if it´s valid
    If CharCreation.Character_Check Then
        'Add char to Account
        If Account.Account_Character_Add(Name, slot) Then
            'Load stats
            char_name = Name
            char_race = race
            char_class = Class
            char_alignment = align
            char_sphere = sphere
            char_psionic_power = psionic_power
            char_STR = new_STR
            char_DEX = new_DEX
            char_CON = new_CON
            char_INT = new_INT
            char_WIS = new_WIS
            char_CHR = new_CHR
            char_portrait = Portrait
            player_char_data_index = char_data_index
            
'TODO: Get these right!!!! (They are just testing values)
            'Set default starting pos
            char_map_id = 1
            char_map_x = 15
            char_map_y = 15
            char_heading = 5
            
            Character_Create = True
            Exit Function
        End If
    End If
    
    Character_Create = False
End Function

Public Function Character_Create_Check_Name(ByVal char_name As String) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/6/2004
'If there is no char file with the given name returns true
'*****************************************************************
    Character_Create_Check_Name = Not General_File_Exists(players_path & "\" & char_name & ".ini", vbNormal)
End Function

Public Function Account_Create(ByVal Account_name As String, ByVal first_name As String, ByVal last_name As String, ByVal email As String) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/6/2004
'
'*****************************************************************
    Account_Create = Account.Account_Create(Account_name, player_password, first_name, last_name, email)
End Function

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

Private Function Save_Stats_To_File(ByVal file_path As String) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    'STATS
    General_Var_Write file_path, "STATS", "race", CStr(char_race)
    General_Var_Write file_path, "STATS", "class", CStr(char_class)
    General_Var_Write file_path, "STATS", "align", CStr(char_alignment)
    General_Var_Write file_path, "STATS", "sphere", CStr(char_sphere)
    General_Var_Write file_path, "STATS", "psionic", CStr(char_psionic_power)
    General_Var_Write file_path, "STATS", "level", CStr(char_level)
    General_Var_Write file_path, "STATS", "portrait", CStr(char_portrait)
    
    General_Var_Write file_path, "STATS", "hp", CStr(char_hp)
    General_Var_Write file_path, "STATS", "hp_max", CStr(char_hp_max)
    General_Var_Write file_path, "STATS", "mp", CStr(char_mp)
    General_Var_Write file_path, "STATS", "mp_max", CStr(char_mp_max)
    General_Var_Write file_path, "STATS", "exp", CStr(char_exp)
    General_Var_Write file_path, "STATS", "exp_next", CStr(char_exp_next)
    General_Var_Write file_path, "STATS", "str", CStr(char_STR)
    General_Var_Write file_path, "STATS", "dex", CStr(char_DEX)
    General_Var_Write file_path, "STATS", "con", CStr(char_CON)
    General_Var_Write file_path, "STATS", "int", CStr(char_INT)
    General_Var_Write file_path, "STATS", "wis", CStr(char_WIS)
    General_Var_Write file_path, "STATS", "cha", CStr(char_CHR)
    
    General_Var_Write file_path, "STATS", "Sysop", CStr(char_authentication_status)
    
    General_Var_Write file_path, "STATS", "EquippedWeaponSlot", CStr(char_weapon_equipped_slot)
    
    Save_Stats_To_File = True
End Function

Public Function NPC_Talk_To_NPC(ByVal response_index As Long) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/23/2004
'Returns the NPC´s greet and all responses.
'*****************************************************************
    If char_talking_to_NPC = 0 Then
        Exit Function
    End If
    
    Dim npc As clsNPC
    Dim speech As clsSpeech
    Set npc = npc_list.Item(char_talking_to_NPC)
    Set speech = speech_list.Item(npc.Speech_Index)
    
    NPC_greet_index = speech.Speech_Respond(NPC_greet_index, response_index)
    
    'Make sure conversation hasn´t ended
    If NPC_greet_index = -1 Then
        'Reset values
        char_talking_to_NPC = 0
        NPC_greet_index = 0
        'Tell the client conversation is over
        NPC_Talk_To_NPC = ""
        Exit Function
    End If
    
    NPC_Talk_To_NPC = speech.Speech_Get_Speech(NPC_greet_index)
End Function

Public Function NPC_Start_Talk_To_NPC(ByVal map_x As Long, ByVal map_y As Long) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/22/2004
'Returns the NPC´s greet and all responses.
'*****************************************************************
    Dim map As clsMap
    Dim npc As clsNPC
    Dim speech As clsSpeech
    
    'Get the NPC index
    Set map = map_list.Item(char_map_id)
    char_talking_to_NPC = map.Char_Get_NPC_ID(map_x, map_y)
    
    'Check if there is a char
    If char_talking_to_NPC = 0 Then
        NPC_Start_Talk_To_NPC = ""
        Exit Function
    End If
    
    Set npc = npc_list.Item(char_talking_to_NPC)
    'Check NPC has a speech set
    If npc.Speech_Index Then
        Set speech = speech_list.Item(npc.Speech_Index)
        NPC_greet_index = 0
        NPC_Start_Talk_To_NPC = speech.Speech_Get_Speech(0) & P_DELIMITER & npc.Name_Char & P_DELIMITER & npc.Portrait
    Else
    
        'Tell the client conversation is over
        NPC_Start_Talk_To_NPC = ""
    End If
End Function

Public Property Let Stats_Total(ByVal new_value As Long)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/23/2004
'
'*****************************************************************
    char_stats_total = new_value
End Property

Public Property Get Stats_Total() As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/23/2004
'
'*****************************************************************
    Stats_Total = char_stats_total
End Property

Public Property Get Name_Char() As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/26/2004
'
'*****************************************************************
    Name_Char = char_name
End Property

Public Function Target(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/14/2004
'
'*****************************************************************
    'Get Map Object
    Dim map As clsMap
    Set map = map_list.Item(char_map_id)
    
    'Check who is in the position clicked
    player_target = map.Char_Get_Player_ID(map_x, map_y)
    
    'Send Player/NPC name back to client for display
    If player_target <> "0" Then
        Dim tplayer As clsPlayer
        Set tplayer = player_list.Item(player_target)
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Target, player_target & P_DELIMITER & tplayer.Name_Char
        Target = True
        Exit Function
    End If
    
    player_target_npc = map.Char_Get_NPC_ID(map_x, map_y)
    
    If player_target_npc <> "0" Then
        Dim tnpc As clsNPC
        Set tnpc = npc_list.Item(player_target_npc)
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Target, player_target_npc & P_DELIMITER & tnpc.Name_Char
        Target = True
        Exit Function
    End If
    
    Target = False
End Function

Public Property Get CharAuthenticationStatus() As CharStatus
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/21/2004
'
'**************************************************************
    CharAuthenticationStatus = char_authentication_status
End Property

Public Property Let CharAuthenticationStatus(ByVal nNewValue As CharStatus)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/21/2004
'
'**************************************************************
    char_authentication_status = nNewValue
End Property

Public Function Attack() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    Dim map As clsMap
    Dim attacked_char As Long
    Dim damage As Long
    Dim user As Boolean
    
    'Check if there is something in front of the user
    Set map = map_list.Item(char_map_id)
    
    'Make usre there is someone in front
    attacked_char = map.Char_In_Front(char_map_x, char_map_y, char_heading, user)
    
    If attacked_char Then
        If char_weapon_equipped_slot Then
            damage = General_Random_Number(item_list(char_inventory(char_weapon_equipped_slot).item_index).min_dam, item_list(char_inventory(char_weapon_equipped_slot).item_index).max_dam)
        Else
            damage = General_Random_Number(1, 4)
        End If
        
        'Check if it was an NPC or user
        If user Then
            Dim player As clsPlayer
            Set player = player_list.Item(attacked_char)
            
            'Notify both users
            dps.Send_Command to_id, ConnectionID, s_Chat, s_Chat_Text, "You have hitted " & player.Name_Char & " for " & CStr(damage)
            dps.Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Text, char_name & "has hitted you for " & CStr(damage)
            'Make damage
            player.HP = player.HP - damage
            If player.HP < 1 Then
                'If it was killed, notify both users
                dps.Send_Command to_id, ConnectionID, s_Chat, s_Chat_Text, "You have killed " & player.Name_Char & ". You gained " & char_exp & " expirience points."
                dps.Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Text, char_name & "has killed you"
                
                'Give expirience
                'char_exp = char_exp +
                
                'Check if user leveled up
                Do While char_exp >= char_exp_next
                    Char_Level_Up
                    'Notify player
                    dps.Send_Command to_id, ConnectionID, s_Chat, s_Chat_Text, "You have leveled up. You are now level " & char_level
                Loop
            End If
        Else
            Dim npc As clsNPC
            Set npc = npc_list.Item(attacked_char)
            
            'Make damage
            npc.NPC_HP = npc.NPC_HP - damage
            dps.Send_Command to_id, ConnectionID, s_Chat, s_Chat_Text, "You have hitted " & npc.Name_Char & " for " & CStr(damage)
            
            If npc.NPC_HP < 1 Then
                char_exp = char_exp + npc.EXP
                dps.Send_Command to_id, ConnectionID, s_Chat, s_Chat_Text, "You have killed " & npc.Name_Char & ". You gained " & npc.EXP & "expirience points"
            End If
        End If
        
        Attack = True
    End If
End Function

Public Function Item_Equip(ByVal Item_Slot As Long) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'Equips / unequips the item in the selected slot
'*****************************************************************
    If Item_Slot > c_Player_Item_Slots Or Item_Slot < 0 Then Exit Function
    
    'Make sure there is something in the given inventory slot
    If char_inventory(Item_Slot).item_index Then
        char_inventory(Item_Slot).equiped = Not char_inventory(Item_Slot).equiped
        If item_list(char_inventory(Item_Slot).item_index).item_type = itWeapon Then
            If char_inventory(Item_Slot).equiped Then
                char_weapon_equipped_slot = Item_Slot
            Else
                char_weapon_equipped_slot = 0
            End If
        End If
    End If
End Function

Private Sub Char_Level_Up()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    'Update level and expirience points
    char_level = char_level + 1
    char_exp = char_exp - char_exp_next
    
    'Update next level required expirience
    Select Case char_class
        Case Gladiator
            Char_Level_Up_Gladiator
        
        Case Fighter
            Char_Level_Up_Figther
        
        Case Thief
            Char_Level_Up_Thief
        
        Case Ranger
            Char_Level_Up_Ranger
        
        Case Psionicist
            Char_Level_Up_Psionicist
        
        Case Preserver
            Char_Level_Up_Preserver
        
        Case Defiler
            Char_Level_Up_Defiler
        
        Case Cleric
            Char_Level_Up_Cleric
        
        Case Druid
            Char_Level_Up_Druid
    End Select
End Sub

Private Sub Char_Level_Up_Gladiator()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    Select Case char_level
        Case Is < 7
            char_exp_next = char_exp_next * 2
        
        Case 7
            char_exp_next = 75000
        
        Case Is < 10
            char_exp_next = char_exp_next * 2
        
        Case Is < 16
            char_exp_next = char_exp_next + 300000
    End Select
End Sub

Private Sub Char_Level_Up_Figther()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    Select Case char_level
        Case Is < 8
            char_exp_next = char_exp_next * 2
        
        Case 8
            char_exp_next = 125000
        
        Case Is < 11
            char_exp_next = char_exp_next * 2
        
        Case Is < 16
            char_exp_next = char_exp_next + 250000
    End Select
End Sub

Private Sub Char_Level_Up_Thief()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    Select Case char_level
        Case Is < 7
            char_exp_next = char_exp_next * 2
        
        Case Is < 9
            char_exp_next = char_exp_next * 3 / 2
        
        Case 9
            char_exp_next = 135000
        
        Case 10
            char_exp_next = 250000
        
        Case 11
            char_exp_next = 375000
        
        Case 12
            char_exp_next = 750000
        
        Case Is < 16
            char_exp_next = char_exp_next + 375000
    End Select
End Sub

Private Sub Char_Level_Up_Ranger()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    Select Case char_level
        Case Is < 7
            char_exp_next = char_exp_next * 2
        
        Case 7
            char_exp_next = 75000
        
        Case Is < 10
            char_exp_next = char_exp_next * 2
        
        Case Is < 16
            char_exp_next = char_exp_next + 300000
    End Select
End Sub

Private Sub Char_Level_Up_Psionicist()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    Select Case char_level
        Case Is < 5
            char_exp_next = char_exp_next * 2
        
        Case 5
            char_exp_next = 16500
        
        Case 6
            char_exp_next = 30000
        
        Case 7
            char_exp_next = 55000
        
        Case 8
            char_exp_next = 100000
        
        Case 9
            char_exp_next = 200000
        
        Case Is < 15
            char_exp_next = char_exp_next + 200000
        
        Case 15
            char_exp_next = 1500000
    End Select
End Sub

Private Sub Char_Level_Up_Preserver()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    Select Case char_level
        Case Is < 7
            char_exp_next = char_exp_next * 2
        
        Case Is < 8
            char_exp_next = char_exp_next * 3 / 2
        
        Case 9
            char_exp_next = 135000
        
        Case 10
            char_exp_next = 250000
        
        Case 11
            char_exp_next = 375000
        
        Case Is < 16
            char_exp_next = char_exp_next + 375000
    End Select
End Sub

Private Sub Char_Level_Up_Defiler()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************

End Sub

Private Sub Char_Level_Up_Cleric()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    Select Case char_level
        Case Is < 5
            char_exp_next = char_exp_next * 2
        
        Case 5
            char_exp_next = 13000
        
        Case 6
            char_exp_next = 27500
        
        Case Is < 9
            char_exp_next = char_exp_next * 2
        
        Case 9
            char_exp_next = 225000
        
        Case Is < 16
            char_exp_next = char_exp_next + 225000
    End Select
End Sub

Private Sub Char_Level_Up_Druid()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    Select Case char_level
        Case Is < 4
            char_exp_next = char_exp_next * 2
        
        Case 4
            char_exp_next = 7500
        
        Case 5
            char_exp_next = 12500
        
        Case 6
            char_exp_next = 20000
        
        Case 7
            char_exp_next = 35000
        
        Case 8
            char_exp_next = 60000
        
        Case 9
            char_exp_next = 90000
        
        Case 10
            char_exp_next = 125000
        
        Case 11
            char_exp_next = 200000
        
        Case 12
            char_exp_next = 300000
        
        Case 13
            char_exp_next = 750000
        
        Case Is < 16
            char_exp_next = char_exp_next * 2
    End Select
End Sub
