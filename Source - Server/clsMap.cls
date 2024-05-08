VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsMap.cls - ORE Server Map Object - v0.5.0
'
'Holds a grid of MapTile objects which represents a whole map.
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
'   - Add: Item_Create method
'   - Add: Save_Ini_By_Name method
'   - Add: Char_In_Front method
'   - Change: Load_From_File and Load_ini_From_File methods were completely
'               rewritten to work with new map format
'
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 2/23/2004
'   - Add: Char_Get_NPC_ID and Char_Get_Player_ID methods
'
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************
Option Explicit

'***************************
'Variables
'***************************
Private dps As ctlDirectPlayServer
Private script_engine As clsScriptEngine
Private map_list As clsList
Private char_list As clsList
Private npc_list As clsList

Private map_id As Long
Private map_connection_id As Long
Private map_name As String
Private map_description As String
Private maps_path As String

Private map_grid_x_max As Long
Private map_grid_x_min As Long
Private map_grid_y_max As Long
Private map_grid_y_min As Long
Private map_grid() As clsMapTile

Private Sub Class_Initialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'**************************************************************
End Sub

Private Sub Class_Terminate()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
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
    id = map_id
End Property

Public Property Get Name() As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    Name = map_name
End Property

Public Function Grid_X_Max()
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: Men orka titta.
'
'**************************************************************
    Grid_X_Max = map_grid_x_max
End Function

Public Function Grid_X_Min()
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: Men orka titta.
'
'**************************************************************
    Grid_X_Min = map_grid_x_min
End Function

Public Function Grid_Y_Max()
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: Men orka titta.
'
'**************************************************************
    Grid_Y_Max = map_grid_y_max
End Function

Public Function Grid_Y_Min()
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: Men orka titta.
'
'**************************************************************
    Grid_Y_Min = map_grid_y_min
End Function

Public Property Let Name(ByVal vNewValue As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    map_name = vNewValue
End Property

Public Property Get ConnectionID() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    ConnectionID = map_connection_id
End Property

Public Property Let ConnectionID(ByVal vNewValue As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    map_connection_id = vNewValue
End Property

Public Function Save_Ini_By_Name(ByVal file_path As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'
'**************************************************************
    Dim X As Long
    Dim Y As Long
    Dim LoopC As Long
    Dim tile As clsMapTile
    Dim char As clsChar
    Dim counter As Long
    
   'If file already exists kill it
    If General_File_Exists(file_path, vbNormal) Then
        Kill file_path
    End If
    
    General_Var_Write file_path, "GENERAL", "map_description", map_description
    
    'NPCs
    counter = 1
    If npc_list.Count Then
        Dim npc As clsNPC
        
        For LoopC = npc_list.LowerBound To npc_list.UpperBound
            Set npc = npc_list.Item(LoopC)
            If npc.MapID = id Then
                General_Var_Write file_path, "NPC", CStr(counter), CStr(X) & "-" & CStr(Y) & "-" & CStr(npc.Data_Index)
                counter = counter + 1
            End If
        Next LoopC
    End If
    General_Var_Write file_path, "NPC", "count", CStr(counter - 1)
    
    'Exits
    counter = 1
    For X = Grid_X_Min To Grid_X_Max
        For Y = Grid_Y_Min To Grid_Y_Max
            If map_grid(X, Y).TileExitMapID Then
                General_Var_Write file_path, "EXIT", CStr(counter), CStr(X) & "-" & CStr(Y) & "-" & map_list.Find("map_id", map_grid(X, Y).TileExitMapID).map_name & "-" & CStr(map_grid(X, Y).TileExitMapX) & "-" & CStr(map_grid(X, Y).TileExitMapY)
                counter = counter + 1
            End If
        Next Y
    Next X
    General_Var_Write file_path, "EXIT", "count", CStr(counter - 1)
    
    'Items
    counter = 1
    For X = Grid_X_Min To Grid_X_Max
        For Y = Grid_Y_Min To Grid_Y_Max
            If map_grid(X, Y).ItemDataIndex Then
                General_Var_Write file_path, "ITEM", CStr(counter), CStr(X) & "-" & CStr(Y) & "-" & CStr(map_grid(X, Y).ItemDataIndex) & "-" & CStr(map_grid(X, Y).ItemAmount)
                counter = counter + 1
            End If
        Next Y
    Next X
    General_Var_Write file_path, "ITEM", "count", CStr(counter - 1)
    
    'Triggers are not saved since they are static data
    
    Save_Ini_By_Name = True
End Function

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

Public Function Initialize(ByRef s_dps As ctlDirectPlayServer, ByRef s_script_engine As clsScriptEngine, ByRef s_map_list As clsList, ByRef s_char_list As clsList, ByRef s_npc_list As clsList, ByVal s_id As Long, Optional ByVal s_name As String, Optional ByRef s_maps_path As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'**************************************************************
    Set dps = s_dps
    Set script_engine = s_script_engine
    Set map_list = s_map_list
    Set char_list = s_char_list
    Set npc_list = s_npc_list
    
    map_id = s_id
    map_name = s_name
    maps_path = s_maps_path
    
    Initialize = True
End Function

Public Function Deinitialize() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'
'*****************************************************************
End Function

Public Function Grid_Initialize(ByVal s_x_min As Long, ByVal s_x_max As Long, ByVal s_y_min As Long, ByVal s_y_max As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Setup up the map grid
'*****************************************************************
    Dim Y As Long
    Dim X As Long
    
    'Set new size
    map_grid_x_max = s_x_max
    map_grid_x_min = s_x_min
    map_grid_y_max = s_y_max
    map_grid_y_min = s_y_min
    ReDim map_grid(map_grid_x_min To map_grid_x_max, map_grid_y_min To map_grid_y_max)
    
    'Initialize grid
    For Y = map_grid_y_min To map_grid_y_max
        For X = map_grid_x_min To map_grid_x_max
            Set map_grid(X, Y) = New clsMapTile
            map_grid(X, Y).Initialize script_engine, map_list, char_list, map_id, X, Y
        Next X
    Next Y
    
    Grid_Initialize = True
End Function

Public Function Grid_Deinitialize() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Clear map grid
'*****************************************************************
    Dim Y As Long
    Dim X As Long
    
    If map_grid_x_max <> 0 Then
        For Y = map_grid_y_min To map_grid_y_max
             For X = map_grid_x_min To map_grid_x_max
                 map_grid(X, Y).Deinitialize
                 Set map_grid(X, Y) = Nothing
             Next X
         Next Y
     End If
     Grid_Deinitialize = True
End Function

Public Function Grid_Find_Nearest_Player(ByVal s_map_x As Long, ByVal s_map_y As Long, ByVal s_scan_range As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Returns the nearest player id to the given position
'*****************************************************************
    Dim X As Long
    Dim Y As Long
    
    Dim t_id As Long
    Dim t_distance As Single
    
    For Y = s_map_y - s_scan_range To s_map_y + s_scan_range
        For X = s_map_x - s_scan_range To s_map_x + s_scan_range
        
            If Tile_Position_Legal(X, Y) Then
                If map_grid(X, Y).CharID Then
                    If char_list.Item(map_grid(X, Y).CharID).PlayerID Then
                        If t_id = 0 Then
                            t_id = char_list.Item(map_grid(X, Y).CharID).PlayerID
                            t_distance = General_Distance(s_map_x, s_map_y, X, Y)
                        Else
                            If General_Distance(s_map_x, s_map_y, X, Y) < t_distance Then
                                t_id = char_list.Item(map_grid(X, Y).CharID).PlayerID
                                t_distance = General_Distance(s_map_x, s_map_y, X, Y)
                            End If
                        End If
                    End If
                End If
            End If
            
        Next X
    Next Y
    
    Grid_Find_Nearest_Player = t_id
End Function

Public Function Load_By_Name(ByVal s_name As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Load a map by name
'Edited by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Get map file path
    Load_By_Name = Load_From_File(maps_path & "\" & s_name)
End Function

Public Function Load_From_File(ByVal file_path As String) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'Load map data from file using an filepath
'*****************************************************************
    Dim Y As Long
    Dim X As Long
    Dim map_path As String
    Dim tile_flags As Byte
    Dim layer_flags As Byte
    Dim deco_flags As Byte
    Dim TempInt As Integer
    Dim TempInt2 As Integer
    Dim TempLng As Long
    Dim TempByte As Byte
    Dim TempByte2 As Byte
    Dim LoopC As Long
    
    'Get map file path
    map_path = file_path
    
    'If file doesn't exists, exit
    If Not (General_File_Exists(map_path, vbNormal)) Then
        Exit Function
    End If
    
    map_grid_x_min = 1
    map_grid_y_min = 1
    
    'Open .map file
    Open map_path For Binary As #1
    Seek #1, 1
    
    'map Header
    Get #1, , map_grid_x_max
    Get #1, , map_grid_y_max
    Get #1, , TempLng
    
    'Clear out and resize map
    Grid_Initialize map_grid_x_min, map_grid_x_max, map_grid_y_min, map_grid_y_max
    
    'Read .map file
    For Y = map_grid_y_min To map_grid_y_max
        For X = map_grid_x_min To map_grid_x_max
            
            '.map file
            
            'Tile Flags
            Get #1, , tile_flags
            
            'Blocked
            map_grid(X, Y).Blocked = (tile_flags And 4)
            
            'Layer 1
            Get #1, , TempInt
            
            'Layer 2
            If tile_flags And 1 Then
                Get #1, , TempInt
                Get #1, , layer_flags
                'Check for angle
                If layer_flags And 8 Then
                    Get #1, , TempInt2
                End If
            End If
            
            'Layer 3
            If tile_flags And 2 Then
                Get #1, , TempInt
                Get #1, , layer_flags
                'Check for angle
                If layer_flags And 8 Then
                    Get #1, , TempInt2
                End If
            End If
            
            'Base lights
            If tile_flags And 8 Then
                For LoopC = 0 To 3
                    Get #1, , TempLng
                Next LoopC
            End If
            
            'Lights
            If tile_flags And 16 Then
                Get #1, , TempLng
                Get #1, , TempByte
            End If
            
            'Particle Streams
            If tile_flags And 64 Then
                Get #1, , TempByte
            End If
            
            'Decorations
            If tile_flags And 32 Then
                Get #1, , TempByte
                For LoopC = 1 To TempByte
                    Get #1, , TempInt
                    Get #1, , deco_flags
                    Get #1, , TempByte2
                    If deco_flags And 8 Then
                        Get #1, , TempInt2
                    End If
                Next LoopC
            End If
            
            'Non combat zone trigger
            If tile_flags And 128 Then map_grid(X, Y).Trigger = 3
        Next X
    Next Y
    
    'Close .map file
    Close #1
    
    Load_From_File = True
End Function

Public Function Load_Ini_By_Name(ByVal s_name As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 9/02/2004
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    Dim file_path As String
    Dim LoopC As Long
    Dim t_count As Long
    Dim X As Long
    Dim Y As Long
    Dim temp_string As String

    'Load ini file
    file_path = maps_path & "\" & s_name & ".ini"
    
    map_description = General_Var_Get(file_path, "GENERAL", "map_description")
        
    'NPC
    t_count = Val(General_Var_Get(file_path, "NPC", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "NPC", CStr(LoopC))
        If temp_string <> "" Then
            dps.NPC_Create CLng(General_Field_Read(3, temp_string, Asc("-"))), map_id, CLng(General_Field_Read(1, temp_string, Asc("-"))), CLng(General_Field_Read(2, temp_string, Asc("-")))
        End If
    Next LoopC
    
    'Item
    t_count = Val(General_Var_Get(file_path, "ITEM", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "ITEM", CStr(LoopC))
        If temp_string <> "" Then
             X = CLng(General_Field_Read(1, temp_string, Asc("-")))
             Y = CLng(General_Field_Read(2, temp_string, Asc("-")))
             map_grid(X, Y).ItemDataIndex = CLng(General_Field_Read(3, temp_string, Asc("-")))
             map_grid(X, Y).ItemAmount = CLng(General_Field_Read(4, temp_string, Asc("-")))
        End If
    Next LoopC
    
    'Exit
    t_count = Val(General_Var_Get(file_path, "EXIT", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "EXIT", CStr(LoopC))
        If temp_string <> "" Then
            X = CLng(General_Field_Read(1, temp_string, Asc("-")))
            Y = CLng(General_Field_Read(2, temp_string, Asc("-")))
            map_grid(X, Y).TileExitMapID = map_list.Find("Name", General_Field_Read(3, temp_string, Asc("-"))).id
            map_grid(X, Y).TileExitMapX = CLng(General_Field_Read(4, temp_string, Asc("-")))
            map_grid(X, Y).TileExitMapY = CLng(General_Field_Read(5, temp_string, Asc("-")))
        End If
    Next LoopC
    
    'Triggers
    t_count = Val(General_Var_Get(file_path, "TRIGGER", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "TRIGGER", CStr(LoopC))
        If temp_string <> "" Then
            X = CLng(General_Field_Read(1, temp_string, Asc("-")))
            Y = CLng(General_Field_Read(2, temp_string, Asc("-")))
            map_grid(X, Y).Trigger = CLng(General_Field_Read(3, temp_string, Asc("-")))
        End If
    Next LoopC
    
    Load_Ini_By_Name = True
End Function

Public Function Tile_Position_Legal(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/21/2003
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If map_x < map_grid_x_min Or map_x > map_grid_x_max Or map_y < map_grid_y_min Or map_y > map_grid_y_max Then
        Tile_Position_Legal = False
        Exit Function
    End If
    
    Tile_Position_Legal = True
End Function

Public Function Char_Position_Legal(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/21/2003
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    If Tile_Position_Legal(map_x, map_y) = False Then
        Exit Function
    End If
    If map_grid(map_x, map_y).Blocked Then
        Exit Function
    End If
    If map_grid(map_x, map_y).CharID Then
        Exit Function
    End If
    Char_Position_Legal = True
End Function

Public Function Tile_Get(ByVal s_map_x As Long, ByVal s_map_y As Long) As clsMapTile
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'Returns a tile object
'**************************************************************
    Set Tile_Get = map_grid(s_map_x, s_map_y)
End Function

Public Function Char_Move_By_Heading(ByVal s_map_x As Long, ByVal s_map_y As Long, ByVal s_heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/21/2003
'Move a character along a heading
'*****************************************************************
    Dim X As Long
    Dim Y As Long
    Dim nx As Long
    Dim ny As Long
    Dim char_id As Long
    
    X = s_map_x
    Y = s_map_y
    char_id = map_grid(X, Y).CharID
    If char_id = 0 Then
        Exit Function
    End If
    
    nx = X
    ny = Y
    Convert_Heading_to_Direction s_heading, nx, ny
    
    If Char_Position_Legal(nx, ny) = False Then
        Exit Function
    End If
    
    map_grid(X, Y).CharID = 0
    map_grid(nx, ny).CharID = char_id
    
    Char_Move_By_Heading = True
End Function

Public Function Char_In_Front(ByVal s_map_x As Long, ByVal s_map_y As Long, ByVal s_heading As Long, ByRef user As Boolean) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/02/2004
'Returns the PlayerID or NPCID of the char right in front of the given tile,
'in the given heading.
'*****************************************************************
    Dim X As Long
    Dim Y As Long
    
    X = s_map_x
    Y = s_map_y
    
    Convert_Heading_to_Direction s_heading, X, Y
    
    If char_list.Item(map_grid(X, Y).CharID).PlayerID Then
        user = True
        Char_In_Front = char_list.Item(map_grid(X, Y).CharID).PlayerID
    Else
        user = False
        Char_In_Front = char_list.Item(map_grid(X, Y).CharID).NPCID
    End If
End Function

Public Function Char_Get_NPC_ID(ByVal map_x As Long, ByVal map_y As Long) As Long
'**************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/23/2003
'
'**************************************************
    Dim char As clsChar
    Dim char_index As Long
    
    char_index = map_grid(map_x, map_y).CharID
    
    'Check if there is a char
    If char_index Then
        Set char = char_list.Item(char_index)
        Char_Get_NPC_ID = char.NPCID
    Else
        Char_Get_NPC_ID = 0
    End If
End Function
 
Public Function Char_Get_Player_ID(ByVal map_x As Long, ByVal map_y As Long) As Long
'**************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/23/2003
'
'**************************************************
    Dim char As clsChar
    Dim char_index As Long
    
    char_index = map_grid(map_x, map_y).CharID
    
    'Check if there is a char
    If char_index Then
        Set char = char_list.Item(char_index)
        Char_Get_Player_ID = char.PlayerID
    Else
        Char_Get_Player_ID = 0
    End If
End Function

Public Function Item_Create(ByVal map_x As Long, ByVal map_y As Long, ByVal item_index As Long, ByVal Item_Amount As Long) As Boolean
'**************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/14/2004
'
'**************************************************
    map_grid(map_x, map_y).ItemDataIndex = item_index
    map_grid(map_x, map_y).ItemAmount = Item_Amount
End Function
