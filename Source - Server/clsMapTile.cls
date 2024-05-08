VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsMapTile.cls - ORE Server Map Tile Object - v0.5.0
'
'Represents a individual tile on a map
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
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 2/9/2004
'   - Add: Trigger properties
'
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************
Option Explicit

'***************************
'Variables
'***************************
Private script_engine As clsScriptEngine
Private map_list As clsList
Private char_list As clsList

Private tile_map_id As Long
Private tile_map_x As Long
Private tile_map_y As Long

Private tile_blocked As Boolean

Private tile_char_id As Long

Private tile_exit_map_id As Long
Private tile_exit_map_x As String
Private tile_exit_map_y As String

Private tile_item_data_index As Long
Private tile_item_amount As Long

'Added by Juan Martín Sotuyo Dodero
Private tile_trigger As Long

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

Public Property Get Blocked() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'*****************************************************************
    Blocked = tile_blocked
End Property

Public Property Let Blocked(ByVal vNewValue As Boolean)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    tile_blocked = vNewValue
End Property

Public Property Get TileExitMapID() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    TileExitMapID = tile_exit_map_id
End Property

Public Property Let TileExitMapID(ByVal vNewValue As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    tile_exit_map_id = vNewValue
End Property

Public Property Get TileExitMapX() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    TileExitMapX = tile_exit_map_x
End Property

Public Property Let TileExitMapX(ByVal vNewValue As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    tile_exit_map_x = vNewValue
End Property

Public Property Get TileExitMapY() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    TileExitMapY = tile_exit_map_y
End Property

Public Property Let TileExitMapY(ByVal vNewValue As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    tile_exit_map_y = vNewValue
End Property

Public Property Get ItemDataIndex() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    ItemDataIndex = tile_item_data_index
End Property

Public Property Let ItemDataIndex(ByVal vNewValue As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    tile_item_data_index = vNewValue
End Property

Public Property Get ItemAmount() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    ItemAmount = tile_item_amount
End Property

Public Property Let ItemAmount(ByVal vNewValue As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    tile_item_amount = vNewValue
End Property

Public Property Get CharID() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'*****************************************************************
    CharID = tile_char_id
End Property

Public Property Let CharID(ByVal vNewValue As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'*****************************************************************
    Dim command As New clsScriptCommand
    
    If vNewValue Then
        'Char enters
        If char_list.Item(vNewValue).PlayerID Then
            '***********************
            'Script events
            '***********************
            If tile_exit_map_id Then
                'Map_Tile_Exit_Activated
                command.Initialize "Map_Tile_Exit_Activated", 4
                command.Parameter_Set 1, char_list.Item(vNewValue).PlayerID
                command.Parameter_Set 2, map_list.Item(tile_exit_map_id).Name
                command.Parameter_Set 3, tile_exit_map_x
                command.Parameter_Set 4, tile_exit_map_y
                script_engine.Command_Add command
            Else
                'Tile_MapName_X_Y_Player_Enter
                command.Initialize "", 1
                command.ProcedureName = "Tile_" & map_list.Item(tile_map_id).Name & "_" & tile_map_x & "_" & tile_map_y & "_Player_Enter"
                command.Parameter_Set 1, char_list.Item(vNewValue).PlayerID
                script_engine.Command_Add command
            End If
        Else
            '***********************
            'Script events
            '***********************
            'Tile_MapName_X_Y_NPC_Enter
            command.Initialize "", 1
            command.ProcedureName = "Tile_" & map_list.Item(tile_map_id).Name & "_" & tile_map_x & "_" & tile_map_y & "_NPC_Enter"
            command.Parameter_Set 1, char_list.Item(vNewValue).NPCID
            script_engine.Command_Add command
        End If
    Else
        'Char leaves
        If tile_char_id Then
            If char_list.Item(tile_char_id).PlayerID Then
                '***********************
                'Script events
                '***********************
                'Tile_MapName_X_Y_Player_Leave
                command.Initialize "", 1
                command.ProcedureName = "Tile_" & map_list.Item(tile_map_id).Name & "_" & tile_map_x & "_" & tile_map_y & "_Player_Leave"
                command.Parameter_Set 1, char_list.Item(tile_char_id).PlayerID
                script_engine.Command_Add command
            Else
                '***********************
                'Script events
                '***********************
                'Tile_MapName_X_Y_NPC_Leave
                command.Initialize "", 1
                command.ProcedureName = "Tile_" & map_list.Item(tile_map_id).Name & "_" & tile_map_x & "_" & tile_map_y & "_NPC_Leave"
                command.Parameter_Set 1, char_list.Item(tile_char_id).NPCID
                script_engine.Command_Add command
            End If
        End If
    End If

    tile_char_id = vNewValue
End Property

Public Function Initialize(ByRef s_script_engine As clsScriptEngine, ByRef s_map_list As clsList, ByRef s_char_list As clsList, ByVal s_map_id As Long, ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'*****************************************************************
    Set script_engine = s_script_engine
    Set map_list = s_map_list
    Set char_list = s_char_list
    tile_map_id = s_map_id
    tile_map_x = s_map_x
    tile_map_y = s_map_y
    Initialize = True
End Function

Public Function Deinitialize() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'*****************************************************************
    Deinitialize = True
End Function

Public Property Let Trigger(ByVal trigger_id As Long)
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/9/2004
'
'*****************************************************************
    tile_trigger = trigger_id
End Property

Public Property Get Trigger() As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/9/2004
'
'*****************************************************************
    Trigger = tile_trigger
End Property
