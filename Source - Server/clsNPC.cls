VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsNPC.cls - ORE Server NPC Object - v0.5.0
'
'NPC object.
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
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 9/02/2004
'   - Add: Scan_For_Player method
'   - Add: Attack method
'   - Add: NPC_HP properties
'   - Add: EXP property
'   - Add: Data_Index property
'   - Add: NPC speech related stuff
'   - Add: NPC portrait and Portrait property
'
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************
Option Explicit

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
'Added by Juan Martín Sotuyo Dodero
Private speech_list As clsList

Private npc_ini_path As String

Private npc_id As Long
Private npc_name As String
Private npc_decription As String
Private npc_char_id As Long
    
Private npc_map_id As Long
Private npc_map_x As Long
Private npc_map_y As Long
Private npc_heading As Long

Private npc_char_data_index As Long
Private npc_Data_Index As Long
Private npc_ai_script As String
'Added by Juan Martín Sotuyo Dodero
Private npc_speech_index As Long
Private npc_portrait As Long

'Added by Juan Martín Sotuyo Dodero
Private equiped_weapon_index As Long
Private npc_STR As Long

Private session_variable_list() As Session_Variable
Private npc_health As Integer
Private npc_exp As Integer
Private npc_original_health As Integer

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
        id = npc_id
End Property

Public Property Get Name_Char() As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
        Name_Char = npc_name
End Property

Public Property Let Name_Char(vNewValue As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
        npc_name = vNewValue
End Property

Public Property Get Name_Upper_Case() As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
        Name_Upper_Case = UCase$(npc_name)
End Property

Public Property Get CharID() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    CharID = npc_char_id
End Property

Public Property Let CharID(ByVal vNewValue As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    npc_char_id = vNewValue
End Property

Public Property Get MapID() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    MapID = npc_map_id
End Property

Public Property Get MapX() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    MapX = npc_map_x
End Property

Public Property Get MapY() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    MapY = npc_map_y
End Property

Public Property Get Heading() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    Heading = npc_heading
End Property

Public Property Let Heading(ByVal vNewValue As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Change the heading of the npc
'**************************************************************
    'Check heading
    If vNewValue < 1 Or vNewValue > 8 Then
        Exit Property
    End If
    'Set
    npc_heading = vNewValue
   'Send packet
    dps.Send_Command to_id, map_list.Item(npc_map_id).ConnectionID, s_Char, s_Char_Heading_Set, _
        s_Packet_Char_Heading_Set(npc_char_id, npc_heading)
End Property

Public Property Get CharDataIndex() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    CharDataIndex = npc_char_data_index
End Property

Public Property Let CharDataIndex(ByVal vNewValue As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    npc_char_data_index = vNewValue
End Property

Public Property Get AiScript() As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    AiScript = npc_ai_script
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

Public Function Initialize(ByRef s_dps As ctlDirectPlayServer, ByRef s_script_engine As clsScriptEngine, ByRef s_map_list As clsList, ByRef s_player_list As clsList, ByRef s_npc_list As clsList, ByRef s_speech_list As clsList, ByVal s_id As Long, ByVal s_npc_ini_path As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/28/2003
'Edited by Juan Martín Sotuyo Dodero
'*****************************************************************
    Set dps = s_dps
    Set script_engine = s_script_engine
    Set map_list = s_map_list
    Set player_list = s_player_list
    Set npc_list = s_npc_list
    Set speech_list = s_speech_list
    npc_id = s_id
    npc_ini_path = s_npc_ini_path
    
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

Public Function Load_From_Ini(ByVal s_npc_data_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Edited by Juan Martín Sotuyo Dodero
'*****************************************************************
    If General_File_Exists(npc_ini_path, vbNormal) = False Then
        Exit Function
    End If

    npc_Data_Index = s_npc_data_index
    npc_name = General_Var_Get(npc_ini_path, "NPC" & s_npc_data_index, "npc_name")
    npc_char_data_index = CLng(General_Var_Get(npc_ini_path, "NPC" & s_npc_data_index, "npc_char_data_index"))
    npc_ai_script = General_Var_Get(npc_ini_path, "NPC" & s_npc_data_index, "npc_ai_script")
    npc_speech_index = speech_list.Find_Index("speech_name", General_Var_Get(npc_ini_path, "NPC" & s_npc_data_index, "npc_speech_file"))
    npc_heading = 5
    npc_portrait = CLng(General_Var_Get(npc_ini_path, "NPC" & s_npc_data_index, "npc_portrait"))
    
    Load_From_Ini = True
End Function

Public Function Move_By_Heading(ByVal s_heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Move npc along a heading
'*****************************************************************
    Dim X As Long
    Dim Y As Long
    Dim nx As Long
    Dim ny As Long
    
    'Get Map Object
    Dim map As clsMap
    Set map = map_list.Item(npc_map_id)
    
    'Try to move npc
    If map.Char_Move_By_Heading(npc_map_x, npc_map_y, s_heading) Then
        'Moved
        'Update npc position
        X = npc_map_x
        Y = npc_map_y
        nx = X
        ny = Y
        Convert_Heading_to_Direction s_heading, nx, ny
        npc_map_x = nx
        npc_map_y = ny
        npc_heading = s_heading
        'Send move command to entire map
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Move, _
            s_Packet_Char_Move(npc_char_id, s_heading)
        Move_By_Heading = True
    Else
        Move_By_Heading = False
    End If
End Function

Public Function Position_Set(ByVal s_map_id As Long, ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'Change the position of the npc
'**************************************************************
   'Get map object
    Dim map As clsMap
    Set map = map_list.Item(s_map_id)
    
    'See if new player's position is legal
    If map.Char_Position_Legal(s_map_x, s_map_y) = False Then
        'TODO: Recalculate position finding a near by legal one
        Exit Function
    End If

    'See if we are changing maps
    If npc_map_id <> s_map_id Then
        'Remove from old map
        If Map_Remove = False Then Exit Function
        'Add to new map
        If Map_Add(s_map_id, s_map_x, s_map_y) = False Then Exit Function
    Else
        'Erase char from old tile
        Dim old_tile As clsMapTile
        Set old_tile = map.Tile_Get(npc_map_x, npc_map_y)
        old_tile.CharID = 0
        'Update player pos
        npc_map_x = s_map_x
        npc_map_y = s_map_y
        'Add char to new tile
        Dim new_tile As clsMapTile
        Set new_tile = map.Tile_Get(npc_map_x, npc_map_y)
        new_tile.CharID = npc_char_id
        'Send command to clients on map
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Pos_Set, _
            s_Packet_Char_Pos_Set(npc_char_id, npc_map_x, npc_map_y)
    End If
    
    Position_Set = True
End Function

Public Function Map_Add(ByVal s_map_id As Long, ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Add a npc to a map
'**************************************************************
   'Get map object
    Dim map As clsMap
    Set map = map_list.Item(s_map_id)
    
    'See if new npc's position is legal
    If map.Char_Position_Legal(s_map_x, s_map_y) Then
        'Add map to npc
        npc_map_id = s_map_id
        npc_map_x = s_map_x
        npc_map_y = s_map_y
    Else
        Map_Add = False
        Exit Function
    End If
    
    'Add npc to map
    Dim tile As clsMapTile
    Set tile = map.Tile_Get(npc_map_x, npc_map_y)
    tile.CharID = npc_char_id
    
    'Send the new npc char to everyone else on the map
    'Create
    dps.Send_Char_Create to_id, map.ConnectionID, npc_char_id
    
    'Return true
    Map_Add = True
End Function

Public Function Map_Remove() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Remove a npc from a map
'**************************************************************
    'Get map object
    Dim map As clsMap
    Set map = map_list.Item(npc_map_id)
    
    'Remove npc from map
    Dim tile As clsMapTile
    Set tile = map.Tile_Get(npc_map_x, npc_map_y)
    tile.CharID = 0
    
    'Update everyone else on map
    If dps.ServerStatus = s_s_listening Then
        dps.Send_Command to_id, map.ConnectionID, s_Char, s_Char_Remove, CStr(npc_char_id)
    End If
    
    'Remove map from player
    npc_map_id = 0
    npc_map_x = 0
    npc_map_y = 0
    
    Map_Remove = True
End Function

Public Function Session_Variable_Create(ByVal s_variable_name As String, ByVal s_variable_data As Variant) As Boolean
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

Public Property Get Speech_Index() As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/18/2004
'
'*****************************************************************
    Speech_Index = npc_speech_index
End Property

Public Property Get Portrait() As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/23/2004
'
'*****************************************************************
    Portrait = npc_portrait
End Property

Public Property Get Data_Index() As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/14/2004
'
'*****************************************************************
    Data_Index = npc_Data_Index
End Property

Public Property Get NPC_HP() As Integer
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    NPC_HP = npc_health
End Property

Public Property Let NPC_HP(ByVal nNewValue As Integer)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    npc_health = NPC_HP
End Property

Public Property Get EXP() As Integer
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    EXP = npc_exp
End Property

Public Function Attack() As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'*****************************************************************
    Dim map As clsMap
    Dim player As clsPlayer
    Dim player_id As Long
    Dim user As Boolean
    Dim damage As Long
    
    Set map = map_list.Item(npc_map_id)
    
    player_id = map.Char_In_Front(npc_map_x, npc_map_y, npc_heading, user)
    
    'NPCs will only attack users
    If player_id Then
        Set player = player_list.Item(player_id)
        'Calculate damage
        If equiped_weapon_index Then
            damage = General_Random_Number(item_list(equiped_weapon_index).min_dam, item_list(equiped_weapon_index).max_dam)
        Else
            damage = General_Random_Number(1, 4)
        End If
        'Do damage
        player.HP = player.HP - damage
        dps.Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Text, npc_name & "has hitted you for " & CStr(damage)
        
        'Check if user was killed
        If player.HP < 1 Then
            'Notify user
            dps.Send_Command to_id, player.ConnectionID, s_Chat, s_Chat_Text, npc_name & "killed you"
        End If
        
        Attack = True
    End If
End Function

Public Function Scan_For_Player(ByVal s_scan_range As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'Looks around NPC for a player, and returns it's index
'*****************************************************************
    Dim map As clsMap
    
    Set map = map_list.Item(npc_map_id)
    
    Scan_For_Player = map.Grid_Find_Nearest_Player(npc_map_x, npc_map_y, s_scan_range)
End Function
