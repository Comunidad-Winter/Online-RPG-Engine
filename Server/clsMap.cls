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
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
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

Public Property Get ID() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    ID = map_id
End Property

Public Property Get Name() As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/6/2003
'
'**************************************************************
    Name = map_name
End Property

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

Public Function Initialize(ByRef s_dps As ctlDirectPlayServer, ByRef s_script_engine As clsScriptEngine, ByRef s_map_list As clsList, ByRef s_char_list As clsList, ByVal s_id As Long, Optional ByVal s_name As String, Optional ByRef s_maps_path As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'**************************************************************
    Set dps = s_dps
    Set script_engine = s_script_engine
    Set map_list = s_map_list
    Set char_list = s_char_list
    
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
    Dim y As Long
    Dim x As Long
    
    'Set new size
    map_grid_x_max = s_x_max
    map_grid_x_min = s_x_min
    map_grid_y_max = s_y_max
    map_grid_y_min = s_y_min
    ReDim map_grid(map_grid_x_min To map_grid_x_max, map_grid_y_min To map_grid_y_max)
    
    'Initialize grid
    For y = map_grid_y_min To map_grid_y_max
        For x = map_grid_x_min To map_grid_x_max
            Set map_grid(x, y) = New clsMapTile
            map_grid(x, y).Initialize script_engine, map_list, char_list, map_id, x, y
        Next x
    Next y
    
    Grid_Initialize = True
End Function

Public Function Grid_Deinitialize() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Clear map grid
'*****************************************************************
    Dim y As Long
    Dim x As Long
    
    If map_grid_x_max <> 0 Then
        For y = map_grid_y_min To map_grid_y_max
             For x = map_grid_x_min To map_grid_x_max
                 map_grid(x, y).Deinitialize
                 Set map_grid(x, y) = Nothing
             Next x
         Next y
     End If
     Grid_Deinitialize = True
End Function

Public Function Grid_Find_Nearest_Player(ByVal s_map_x As Long, ByVal s_map_y As Long, ByVal s_scan_range As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Returns the nearest player id to the given position
'*****************************************************************
    Dim x As Long
    Dim y As Long
    
    Dim t_id As Long
    Dim t_distance As Single
    
    For y = s_map_y - s_scan_range To s_map_y + s_scan_range
        For x = s_map_x - s_scan_range To s_map_x + s_scan_range
        
            If Tile_Position_Legal(x, y) Then
                If map_grid(x, y).CharID Then
                    If char_list.Item(map_grid(x, y).CharID).PlayerID Then
                        If t_id = 0 Then
                            t_id = char_list.Item(map_grid(x, y).CharID).PlayerID
                            t_distance = General_Distance(s_map_x, s_map_y, x, y)
                        Else
                            If General_Distance(s_map_x, s_map_y, x, y) < t_distance Then
                                t_id = char_list.Item(map_grid(x, y).CharID).PlayerID
                                t_distance = General_Distance(s_map_x, s_map_y, x, y)
                            End If
                        End If
                    End If
                End If
            End If
            
        Next x
    Next y
    
    Grid_Find_Nearest_Player = t_id
End Function

Public Function Load_By_Name(ByVal s_name As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/30/2003
'Load a map by name
'*****************************************************************
    'Get map file path
    Load_By_Name = Load_From_File(maps_path & "\" & "map" & s_name & ".map")
End Function

Public Function Load_From_File(ByVal file_path As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'Load map data from file
'*****************************************************************
    Dim loopc As Long
    Dim LoopC2 As Long
    Dim TempLng As Long
    Dim TempLng2 As Long
    Dim TempLng3 As Long
    Dim TempLng4 As Long
    Dim TempLng5 As Long
    Dim TempLng6 As Long
    Dim TempSgl As Single
    Dim TempBln As Boolean
    Dim TempLngList() As Long
    Dim y As Long
    Dim x As Long
    Dim map_path As String
    
    'Get map file path
    map_path = file_path
    
    'See if map_path exists
    If General_File_Exists(map_path, vbNormal) = False Then
        Load_From_File = False
        Exit Function
    End If
    
    'Open .map file
    Open map_path For Binary As #1
    Seek #1, 1
    
    'Map Header
    
    'map size
    Get #1, , TempLng2
    Get #1, , TempLng3
    Get #1, , TempLng4
    Get #1, , TempLng5
    
    'place holders
    Get #1, , TempLng
    Get #1, , TempLng
    Get #1, , TempLng
    Get #1, , TempLng
    Get #1, , TempLng
    
    'Create Map
    Grid_Initialize TempLng2, TempLng3, TempLng4, TempLng5
    
    'Body File
    For y = map_grid_y_min To map_grid_y_max
        For x = map_grid_x_min To map_grid_x_max
            'Blocked
            Get #1, , TempBln
            map_grid(x, y).Blocked = TempBln
            
            'Layers
            For loopc = 1 To 4
                Get #1, , TempLng
                Get #1, , TempBln
                Get #1, , TempSgl
            Next loopc
            
            'Light values
            Get #1, , TempLng
            Get #1, , TempLng
            Get #1, , TempLng
            Get #1, , TempLng
            
            'Empty place holders for future expansion
            Get #1, , TempLng
            Get #1, , TempLng
            Get #1, , TempLng
            Get #1, , TempLng
            Get #1, , TempLng
            Get #1, , TempLng
            Get #1, , TempLng
        Next x
    Next y
    
    'Read footer
    'Lights
    Get #1, , TempLng
    For loopc = 1 To TempLng
            Get #1, , TempLng2
            Get #1, , TempLng3
            Get #1, , TempLng4
            Get #1, , TempLng5
    Next loopc
    
    'Particle Groups
    Get #1, , TempLng
    For loopc = 1 To TempLng
            Get #1, , TempLng2
            Get #1, , TempLng3
            Get #1, , TempLng4
            Get #1, , TempLng5
            Get #1, , TempBln
            Get #1, , TempSgl
            Get #1, , TempLng6
            ReDim TempLngList(1 To TempLng6)
            For LoopC2 = 1 To TempLng6
                Get #1, , TempLngList(LoopC2)
            Next LoopC2
    Next loopc
    
    'Close .map file
    Close #1
    Load_From_File = True
End Function

Public Function Load_Ini_By_Name(ByVal s_name As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    Dim file_path As String
    Dim loopc As Long
    Dim t_count As Long
    Dim x As Long
    Dim y As Long
    Dim temp_string As String

    'Load ini file
    file_path = maps_path & "\" & "map" & s_name & ".ini"
    
    map_description = General_Var_Get(file_path, "GENERAL", "map_description")
        
    'NPC
    t_count = Val(General_Var_Get(file_path, "NPC", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "NPC", CStr(loopc))
        If temp_string <> "" Then
            dps.NPC_Create CLng(General_Field_Read(3, temp_string, Asc("-"))), map_id, CLng(General_Field_Read(1, temp_string, Asc("-"))), CLng(General_Field_Read(2, temp_string, Asc("-")))
        End If
    Next loopc
    
    'Item
    t_count = Val(General_Var_Get(file_path, "ITEM", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "ITEM", CStr(loopc))
        If temp_string <> "" Then
             x = CLng(General_Field_Read(1, temp_string, Asc("-")))
             y = CLng(General_Field_Read(2, temp_string, Asc("-")))
             map_grid(x, y).ItemDataIndex = CLng(General_Field_Read(3, temp_string, Asc("-")))
             map_grid(x, y).ItemAmount = CLng(General_Field_Read(4, temp_string, Asc("-")))
        End If
    Next loopc
    
    'Exit
    t_count = Val(General_Var_Get(file_path, "EXIT", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "EXIT", CStr(loopc))
        If temp_string <> "" Then
            x = CLng(General_Field_Read(1, temp_string, Asc("-")))
            y = CLng(General_Field_Read(2, temp_string, Asc("-")))
            map_grid(x, y).TileExitMapID = map_list.Find("Name", General_Field_Read(3, temp_string, Asc("-"))).ID
            map_grid(x, y).TileExitMapX = CLng(General_Field_Read(4, temp_string, Asc("-")))
            map_grid(x, y).TileExitMapY = CLng(General_Field_Read(5, temp_string, Asc("-")))
        End If
    Next loopc
    
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
    Dim x As Long
    Dim y As Long
    Dim nx As Long
    Dim ny As Long
    Dim char_id As Long
    
    x = s_map_x
    y = s_map_y
    char_id = map_grid(x, y).CharID
    If char_id = 0 Then
        Exit Function
    End If
    
    nx = x
    ny = y
    Convert_Heading_to_Direction s_heading, nx, ny
    
    If Char_Position_Legal(nx, ny) = False Then
        Exit Function
    End If
    
    map_grid(x, y).CharID = 0
    map_grid(nx, ny).CharID = char_id
    
    Char_Move_By_Heading = True
End Function
