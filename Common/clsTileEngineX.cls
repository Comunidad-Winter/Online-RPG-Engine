VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTileEngineX"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsTileEngineX - ORE Direct3D 8 Tile Engine - v0.5.0
'
'Graphics Engine.
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
'05/27/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Change: Light_Move now called Light_Map_Pos_Set and Light_Move is a new method.
'   -Add: Particle_Group_Map_Pos_Set and Move
'   Sub Release Contributors:
'       05/25/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Add: Map_Bounds_Get_From_File
'           - Add: Light_Move
'
'05/15/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Add: GUI_Grh_Render
'   -Change: Text_Render is now called GUI_Text_Render
'       -Added word_wrap option
'   Sub Release Contributors:
'       05/15/2003 - Fredrik Alexandersson (fredrik@oraklet.zzn.com)
'           - Add: GUI_Box_Filled_Render
'           - Add: GUI_Box_Outline_Render
'
'05/12/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Add: Char_Data_Grh_Index_Get
'
'05/08/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Change: Moved all the General_X methods to a external modGeneral.bas
'
'05/07/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Fix: Map_Save and Map_Load format changed. Added angle and alpha values.
'   Sub Release Contributors:
'       05/07/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Fix: Geometry_Create_Sqaure - fixed rotation
'
'05/02/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - Fix: bug with Light_Move and base light values
'   - Fix: Grh_Get_Info and Grh_Render_To_Hdc when using animated grhs
'   - Add: Base light values can now be specfied for each corner
'      - Changed .map format. Old maps will not work anymore
'      - Modified Map_Base_Light_Set
'   - Remove: Light_Remove_XY (Use Light_Remove Map_Light_Get(X,Y) instead)
'
'04/29/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - Add: Grh_Render_To_Hdc
'   - Add: Map_Save_To_File and Map_Load_From_File
'   - Add: Text_Render method
'   - Change: file name to clsTileEngineX.cls
'   - Change: Split Engine_Render into Engine_Render_Start and Engine_Render_End
'   - Change: Method Names
'       -Engine_Stats_Render
'       -Engine_Blocked_Tiles_Show_Toggle
'   Sub Release Contributors:
'       04/28/2003 - Fredrik Alexandersson (fredrik@oraklet.zzn.com)
'           - Add: Grh_Add_GrhList_To_ListBox method
'           - Change: Particle_Render (added 3 new stream types)
'
'04/07/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - First release
'   Sub Release Contributors:
'       03/31/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Add: Light_Erase method
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
Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

Private Const COLOR_KEY As Long = &HFF000000
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'Paths
Private Const PATH_GRAPHICS = "\graphics"
Private Const PATH_MAPS = "\maps"
Private Const PATH_SOUNDS = "\sounds"
Private Const PATH_SCRIPTS = "\scripts"

'***************************
'Types
'***************************

'Structure to hold all the graphics for the game
Private Type DX_Texture
    inuse As Boolean
    dimension As Long
    d3dtexture As Direct3DTexture8
End Type

'This structure describes a transformed and lit vertex.
Private Type TLVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

'Holds data about where a bmp can be found,
'How big it is and animation info
Private Type Grh_Data
    active As Boolean
    texture_index As Long
    src_x As Long
    src_y As Long
    src_width As Long
    src_height As Long
    
    frame_count As Long
    frame_list(1 To 16) As Long
    frame_speed As Single
End Type

'Points to a Grh_Data and keeps animation info
Private Type grh
    grh_index As Long
    alpha_blend As Boolean
    angle As Single
    frame_speed As Single
    frame_counter As Single
    started As Boolean
    noloop As Boolean
End Type

'Char Body
Private Type Char_Data_Body
    body(1 To 8) As grh
End Type

'Char Data
Private Type Char_Data
    body_list() As Char_Data_Body
    body_count As Long
    label_offset_x As Long
    label_offset_y As Long
End Type

'Hold info about a character
Private Type char
    active As Boolean
    heading As Long
    id As Long
    map_x As Long
    map_y As Long

    chr_data As Char_Data
    chr_data_body_index As Long
    
    label As String
    label_font_index As Long
    label_offset_x As Long
    label_offset_y As Long
    
    scroll_on As Boolean
    scroll_offset_counter_x As Single
    scroll_offset_counter_y As Single
    scroll_direction_x As Long
    scroll_direction_y As Long
End Type

Private Type Map_Exit
    exit_map_name As String
    exit_map_x As Long
    exit_map_y As Long
    
    c_map_x As Long
    c_map_y As Long
End Type

Private Type Map_NPC
    npc_data_index As Long
    
    c_char_data_index As Long
    c_map_x As Long
    c_map_y As Long
End Type

Private Type Map_Item
    item_data_index As Long
    item_amount As Long
    
    c_grh_index As Long
    c_map_x As Long
    c_map_y As Long
End Type

'Map Tile structure
Private Type Map_Tile
    grh(1 To 5) As grh
    blocked As Boolean
    particle_group_index As Long
    char_index As Long
    light_base_value(0 To 3) As Long
    light_value(0 To 3) As Long
    
    exit_index As Long
    npc_index As Long
    item_index As Long
End Type

'Map structure
Private Type Map
    map_grid() As Map_Tile
    map_x_max As Long
    map_x_min As Long
    map_y_max As Long
    map_y_min As Long
    map_description As String
End Type

'The light structure
Private Type Light
    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Long 'Coordinates
    map_y As Long
    color As Long 'Start colour
    range As Long
End Type

'Particle types
Private Type Particle
    friction As Single
    x As Single
    y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    grh As grh
    alive_counter As Long
End Type

Private Type Particle_Group
    active As Boolean
    id As Long
    map_x As Long
    map_y As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Long

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
End Type

'***************************
'Variables
'***************************
'Major DX Objects
Dim dx As DirectX8
Dim d3d As Direct3D8
Dim ddevice As Direct3DDevice8
Dim d3dx As D3DX8
Dim d3dpp As D3DPRESENT_PARAMETERS
Dim d3dcaps As D3DCAPS8
Dim d3ddm As D3DDISPLAYMODE
Dim di As DirectInput8
Dim ddevice_mouse As DirectInputDevice8

'The app path
Dim resource_path As String

'Main form handle
Dim form_hwnd As Long

'Display variables
Dim screen_hwnd As Long
Dim screen_width As Long
Dim screen_height As Long

'Map view area (where the game is played)
Dim view_screen_top As Long 'In pixels
Dim view_screen_left As Long 'In pixels
Dim view_screen_bottom As Long
Dim view_screen_right As Long
Dim view_screen_tile_width As Long 'In tiles
Dim view_screen_tile_height As Long 'In tiles
Dim view_screen_width As Long
Dim view_screen_height As Long

'Buffer area (used to draw object outside the map area but may still show up on the screen)
Dim view_tile_buffer As Long 'In tiles

'Base tile size (smallest possible tile size: must be square)
Dim base_tile_size As Long 'In pixels

'View position: In tiles
Dim view_pos_x As Long
Dim view_pos_y As Long

'Scrolling stuff
Dim scroll_on As Boolean
Dim scroll_direction_x As Long
Dim scroll_direction_y As Long
Dim scroll_offset_counter_x As Single
Dim scroll_offset_counter_y As Single
Dim scroll_pixels_per_frame As Long

'Clip rect (used to cleanup around the edges of the map view area after rendering)
Dim clip_rect(0 To 3) As D3DRECT

'FPS Counters
Dim fps_last_time As Long 'When did we last check the frame rate?
Dim fps_frame_counter As Long 'How many frames have been drawn
Dim fps As Long 'What the current frame rate is.....

'time between frames
Dim timer_elapsed_time As Single
'ticks per frame
Dim timer_ticks_per_frame As Single

'base speed for the engine
Dim engine_base_speed As Single

'total frame counter
Dim total_frame_counter As Long

Dim engine_render_started As Boolean

'windowed or not windowed
Dim engine_windowed As Boolean

'clip border color
Dim engine_clip_border_color As Long

'show engine stats
Dim engine_show_stats As Boolean

'show blocked tiles on map
Dim engine_show_blocked_tiles As Boolean
Dim engine_show_special_tiles As Boolean

'mouse input
Dim input_mouse_button_left As Boolean
Dim input_mouse_button_right As Boolean
Dim input_mouse_moved As Boolean
Dim input_mouse_screen_x As Long
Dim input_mouse_screen_y As Long
Dim input_mouse_view_x As Long
Dim input_mouse_view_y As Long
Dim input_mouse_map_x As Long
Dim input_mouse_map_y As Long
Dim input_mouse_heading As Long

'***************************
'Arrays
'***************************
'Array of surfaces
Dim texture_list() As DX_Texture
Dim texture_count As Long

'Grh Data Array
Dim grh_list() As Grh_Data
Dim grh_count As Long

'Char list
Dim char_list() As char
Dim char_count As Long
Dim char_last As Long

'Char data list
Dim char_data_list() As Char_Data
Dim char_data_count As Long

'Current Map
Dim map_current As Map

'Light list
Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long

'Particle system
Dim particle_group_list() As Particle_Group
Dim particle_group_count As Long
Dim particle_group_last As Long

'Font List
Dim font_list() As D3DXFont
Dim font_count As Long
Dim font_last As Long

Dim exit_list() As Map_Exit
Dim npc_list() As Map_NPC
Dim item_list() As Map_Item

'***************************
'External Functions
'***************************
'KeyInput
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

'MouseInput
Private Type PointAPI
    x As Long
    y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

'For getting the display size in windowed mode
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Old fashion BitBlt functions
Private Const SRCCOPY = &HCC0020
Private Const SRCPAINT = &HEE0086
Private Const SRCAND = &H8800C6
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcsrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Sub Class_Initialize()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
End Sub

Private Sub Class_Terminate()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    Engine_Deinitialize
End Sub

Private Function Convert_Tile_To_View_Y(ByVal y As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_Y = ((y - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_Y = view_screen_top + ((y - 1) * base_tile_size)
    End If
End Function

Private Sub Convert_Screen_To_View(ByVal screen_x As Long, ByVal screen_y As Long, ByRef view_x As Long, ByRef view_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    view_x = screen_x - view_screen_left
    view_y = screen_y - view_screen_top
End Sub

Private Sub Convert_View_To_Map(ByVal view_x As Long, ByVal view_y As Long, ByRef map_x As Long, ByRef map_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim half_view_tile_width As Long
    Dim half_view_tile_height As Long
    
    half_view_tile_width = (view_screen_tile_width \ 2)
    half_view_tile_height = (view_screen_tile_height \ 2)
    
    map_x = (view_x \ base_tile_size)
    map_y = (view_y \ base_tile_size)
    
    If map_x > half_view_tile_width Then
        map_x = (map_x - half_view_tile_width)
    
    Else
        If map_x < half_view_tile_width Then
            map_x = (0 - (half_view_tile_width - map_x))
        Else
            map_x = 0
        End If
    End If
    
    If map_y > half_view_tile_height Then
        map_y = (0 - (half_view_tile_height - map_y))
    Else
        If map_y < half_view_tile_height Then
            map_y = (map_y - half_view_tile_height)
        Else
            map_y = 0
        End If
    End If
    
    map_x = view_pos_x + map_x
    map_y = view_pos_y + map_y
End Sub

Private Sub Convert_Map_To_Direction(ByVal map_x As Long, ByVal map_y As Long, ByRef direction_x As Long, ByRef direction_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = map_x - view_pos_x
    temp_y = map_y - view_pos_y
    
    If temp_x <> 0 Then
        direction_x = temp_x \ Abs(temp_x)
    Else
        direction_x = 0
    End If
    If temp_y <> 0 Then
        direction_y = temp_y \ Abs(temp_y)
    Else
        direction_y = 0
    End If
End Sub

Private Function Convert_Direction_To_Heading(ByVal direction_x As Long, ByVal direction_y As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'North
    If direction_x = 0 And direction_y = -1 Then
        Convert_Direction_To_Heading = 1
    End If
    'NE
    If direction_x = 1 And direction_y = -1 Then
        Convert_Direction_To_Heading = 2
    End If
    'East
    If direction_x = 1 And direction_y = 0 Then
        Convert_Direction_To_Heading = 3
    End If
    'SE
    If direction_x = 1 And direction_y = 1 Then
        Convert_Direction_To_Heading = 4
    End If
    'South
    If direction_x = 0 And direction_y = 1 Then
        Convert_Direction_To_Heading = 5
    End If
    'SW
    If direction_x = -1 And direction_y = 1 Then
        Convert_Direction_To_Heading = 6
    End If
    'West
    If direction_x = -1 And direction_y = 0 Then
        Convert_Direction_To_Heading = 7
    End If
    'NW
    If direction_x = -1 And direction_y = -1 Then
        Convert_Direction_To_Heading = 8
    End If
End Function

Private Sub Convert_Heading_to_Direction(ByVal heading As Long, ByRef direction_x As Long, ByRef direction_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim addY As Long
    Dim addX As Long
    
    'Figure out which way to move
    Select Case heading
    
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
    
    direction_x = direction_x + addX
    direction_y = direction_y + addY
End Sub

Private Function Convert_Tile_To_View_X(ByVal x As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_X = ((x - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_X = view_screen_left + ((x - 1) * base_tile_size)
    End If
End Function

Private Sub Engine_Stats_Render()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim x As Long
    Dim y As Long
    
    Dim offset_x As Long
    
    offset_x = screen_width - 110
    
    'fps
    Device_Text_Render font_list(1), fps & " FPS", 0, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    'delay
    Device_Text_Render font_list(1), timer_ticks_per_frame & " tpf", 20, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    
    'cursor screen pos
    Device_Text_Render font_list(1), input_mouse_screen_x & " sx, " & input_mouse_screen_y & " sy", 40, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    
    Device_Text_Render font_list(1), input_mouse_view_x & " vx, " & input_mouse_view_y & " vy", 60, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    
    Device_Text_Render font_list(1), input_mouse_map_x & " mx, " & input_mouse_map_y & " my", 80, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    
    Device_Text_Render font_list(1), input_mouse_heading & " heading", 100, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
End Sub

Public Sub Engine_Border_Color_Set(ByVal b_color As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sets the color used to clip around the view area
'*****************************************************************
    engine_clip_border_color = b_color
End Sub

Public Sub Engine_Stats_Show_Toggle()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Toggles engine stats
'*****************************************************************
    If engine_show_stats Then
        engine_show_stats = False
    Else
        engine_show_stats = True
    End If
End Sub

Public Sub Engine_Blocked_Tiles_Show_Toggle()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003

'*****************************************************************
    If engine_show_blocked_tiles Then
        engine_show_blocked_tiles = False
    Else
        engine_show_blocked_tiles = True
    End If
End Sub

Public Sub Engine_Special_Tiles_Show_Toggle()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003

'*****************************************************************
    If engine_show_special_tiles Then
        engine_show_special_tiles = False
    Else
        engine_show_special_tiles = True
    End If
End Sub

Public Function Engine_View_Pos_Set(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sets the user postion
'If valid, returns True, else False
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        view_pos_x = map_x
        view_pos_y = map_y
        
    'Update mouse position
    If Input_Mouse_In_View Then
        'Update pos
        Convert_View_To_Map input_mouse_view_x, input_mouse_view_y, input_mouse_map_x, input_mouse_map_y
    End If

    End If
End Function

Public Sub Engine_View_Pos_Get(ByRef map_x As Long, ByRef map_y As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the user postion
'*****************************************************************
    map_x = view_pos_x
    map_y = view_pos_y
End Sub

Public Sub Engine_Base_Speed_Set(ByVal speed As Single)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'*****************************************************************
    engine_base_speed = speed
End Sub

Public Function Engine_Base_Speed_Get() As Single
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the target fps that the engine
'*****************************************************************
    Engine_Base_Speed_Get = engine_base_speed
End Function

Public Function Engine_FPS_Get() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the target fps that the engine
'*****************************************************************
    Engine_FPS_Get = fps
End Function

Public Function Engine_Frame_Counter_Get() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Gets the total number of frames since the engine started
'*****************************************************************
    Engine_Frame_Counter_Get = total_frame_counter
End Function

Public Sub Engine_Deinitialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next
    
    Dim i As Long
    
    For i = 1 To UBound(texture_list)
        Set texture_list(i).d3dtexture = Nothing
    Next i
    
    'ddevice_mouse.Unacquire
    Set ddevice_mouse = Nothing
    Set di = Nothing
    
    Set d3dx = Nothing
    Set ddevice = Nothing
    Set d3d = Nothing
    Set dx = Nothing
End Sub

Public Function Engine_Initialize(ByVal f_hwnd As Long, ByVal s_hwnd As Long, ByVal windowed As Boolean, _
                                    ByVal r_path As String, Optional ByVal s_width As Long, Optional ByVal s_height As Long, _
                                    Optional ByVal v_left As Long = 0, Optional ByVal v_top As Long = 0, Optional ByVal v_width_in_tiles As Long = 0, _
                                    Optional ByVal v_height_in_tiles As Long = 0, Optional ByVal tile_size As Long = 32) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'**************************************************************
'On Error GoTo ErrorHandler:
    Engine_Initialize = True
    
    '****************************************************
    'Fill in global variables
    '****************************************************
    'Main display
    screen_hwnd = s_hwnd
    form_hwnd = f_hwnd
    
    'App path
    resource_path = r_path
    
    'Fill in base tile size (must always be square)
    base_tile_size = tile_size 'In pixels
    
    'Fill in viewport sizes. How large the game area is.
    If v_width_in_tiles = 0 Or v_height_in_tiles = 0 Then
        view_screen_tile_width = 20 'In tiles
        view_screen_tile_height = 15 'In tiles
    Else
        view_screen_tile_width = v_width_in_tiles 'In tiles
        view_screen_tile_height = v_height_in_tiles 'In tiles
    End If
    If windowed Then
        Dim target As RECT
        GetWindowRect screen_hwnd, target
        view_screen_top = target.top
        view_screen_left = target.left
        view_screen_right = target.Right
        view_screen_bottom = target.bottom
        view_screen_width = target.Right - target.left
        view_screen_height = target.bottom - target.top
        screen_width = view_screen_width
        screen_height = view_screen_height
        engine_windowed = True
    Else
        screen_width = s_width
        screen_height = s_height
        view_screen_left = v_left 'In pixels
        view_screen_top = v_top 'In pixels
        view_screen_width = view_screen_tile_width * base_tile_size
        view_screen_height = view_screen_tile_height * base_tile_size
        view_screen_right = view_screen_left + view_screen_width - 1
        view_screen_bottom = view_screen_top + view_screen_height - 1
        engine_windowed = False
        
        'Figure out clip plane
        clip_rect(0).x1 = view_screen_left
        clip_rect(0).y1 = 0
        clip_rect(0).x2 = view_screen_left + view_screen_tile_width * base_tile_size
        clip_rect(0).y2 = view_screen_top
        clip_rect(1).x1 = view_screen_left
        clip_rect(1).y1 = view_screen_top + view_screen_tile_height * base_tile_size
        clip_rect(1).x2 = view_screen_left + view_screen_tile_width * base_tile_size
        clip_rect(1).y2 = screen_height + 480
        clip_rect(2).x1 = 0
        clip_rect(2).y1 = 0
        clip_rect(2).x2 = view_screen_left
        clip_rect(2).y2 = screen_height
        clip_rect(3).x1 = view_screen_left + view_screen_tile_width * base_tile_size
        clip_rect(3).y1 = 0
        clip_rect(3).x2 = screen_width
        clip_rect(3).y2 = screen_height
        
    End If
    
    '****************************************************
    'Get external data
    '****************************************************
    'Load Grh List
    Grh_Load_All
    
    'Load body data for characters
    Char_Load_Char_Data
    
    '****************************************************
    'Setup Map
    '****************************************************
    'Buffer area
    view_tile_buffer = 6
    
    'How many pixels to move per frame when scrolling
    scroll_pixels_per_frame = 4
    
    'User start position
    view_pos_x = 1
    view_pos_y = 1
    
    'Create default map
    Map_Create 50, 50
    
    '*******************************
    'Initialize root DirectX8 objects
    '*******************************
    Set dx = New DirectX8
    'Create the Direct3D object
    Set d3d = dx.Direct3DCreate
    'Create helper class
    Set d3dx = New D3DX8
    'Input object
    Set di = dx.DirectInputCreate
    
    '*******************************
    'Initialize video device
    '*******************************
    Dim DevType As CONST_D3DDEVTYPE
    DevType = D3DDEVTYPE_HAL
    'Get the capabilities of the Direct3D device that we specify. In this case,
    'we'll be using the adapter default (the primiary card on the system).
    Call d3d.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, d3dcaps)
    'Grab some information about the current display mode.
    Call d3d.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, d3ddm)
    'Now we'll go ahead and fill the D3DPRESENT_PARAMETERS type.
    With d3dpp
        If engine_windowed Then
            .windowed = 1
            .SwapEffect = D3DSWAPEFFECT_COPY
            .BackBufferFormat = d3ddm.format 'current display depth
        Else
            .SwapEffect = D3DSWAPEFFECT_FLIP
            'Set refresh rate
            .FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
            'Turn off vsync
            .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
            'set color depth
            .BackBufferFormat = D3DFMT_R5G6B5 '16 bit
            'Back buffer size
            .BackBufferWidth = screen_width
            .BackBufferHeight = screen_height
        End If
    End With
    'create device
    Set ddevice = d3d.CreateDevice(D3DADAPTER_DEFAULT, DevType, screen_hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    'setup device
    With ddevice
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        Call .SetVertexShader(FVF)
        'Turn off lighting
        Call .SetRenderState(D3DRS_LIGHTING, 0)
        'Set the render state that uses the alpha component as the source for blending.
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        'Set the render state that uses the inverse alpha component as the destination blend.
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    End With
    
    '*******************************
    'Initialize input devices
    '*******************************
    Set ddevice_mouse = di.CreateDevice("guid_SysMouse")
    Call ddevice_mouse.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call ddevice_mouse.SetCooperativeLevel(form_hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
    ' Set the buffer size
    Dim diProp As DIPROPLONG
    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = 1000
    Call ddevice_mouse.SetProperty("DIPROP_BUFFERSIZE", diProp)
    ddevice_mouse.Acquire
    
    '****************************************************
    'Load textures
    '****************************************************
    Texture_Load_All
    
    '****************************************************
    'Misc stuff
    '****************************************************
    'Set up default font
    Font_Create "Arial", 11, False, False
    
    'Clears the buffer
    Device_Clear
    
Exit Function
ErrorHandler:
    MsgBox "Error in Engine_Initialization: " & Err.Number & ": " & Err.Description
    Engine_Initialize = False
End Function

Public Function Engine_View_Move(ByVal heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim x As Long
    Dim y As Long
    Dim nx As Long
    Dim ny As Long
    
    'Don't move if we are already moving....
    If scroll_on Then
        Engine_View_Move = False
        Exit Function
    End If
    
    'Invalid heading
    If heading < 1 Or heading > 8 Then
        Engine_View_Move = False
        Exit Function
    End If
    
    x = view_pos_x
    y = view_pos_y
    nx = x
    ny = y
    Convert_Heading_to_Direction heading, nx, ny
    
    'See if out new position is legal
    If Map_In_Bounds(nx, ny) Then
        'start the scrolling process
        view_pos_x = nx
        view_pos_y = ny
        
        scroll_offset_counter_x = (base_tile_size * (x - nx))
        scroll_offset_counter_y = (base_tile_size * (y - ny))
        scroll_direction_x = nx - x
        scroll_direction_y = ny - y
        scroll_on = True

        'Update mouse position
        If Input_Mouse_In_View Then
            'Update pos
            Convert_View_To_Map input_mouse_view_x, input_mouse_view_y, input_mouse_map_x, input_mouse_map_y
        End If

        Engine_View_Move = True
    Else
        'not legal don't move
        scroll_direction_x = 0
        scroll_direction_y = 0
    End If
End Function

Public Function Engine_Render_Start() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/29/2003
'
'**************************************************************
'On Error GoTo ErrorHandler:
    Engine_Render_Start = True
    
    '****************************************************
    'Handle Alt-Tab (not quite working yet)
    '****************************************************
    'Call TestCooperativeLevel to see what state the device is in.
    Dim hr As Long
    hr = ddevice.TestCooperativeLevel
    If hr = D3DERR_DEVICELOST Then
        'If the device is lost, exit and wait for it to come back.
        Exit Function
    ElseIf hr = D3DERR_DEVICENOTRESET Then
        'The device became lost for some reason (probably an alt-tab) and now
        'Reset() needs to be called to try and get the device back.
        hr = 0
        hr = Device_Reset
        'If the device failed to be reset, exit the sub.
        If hr Then Exit Function
    End If
    
    '****************************************************
    'Render
    '****************************************************
    
    '*******************************
    'get the screen_rect if windowed
    If engine_windowed Then
        Dim target As RECT
        GetWindowRect screen_hwnd, target
        view_screen_top = target.top
        view_screen_left = target.left
        view_screen_right = target.Right
        view_screen_bottom = target.bottom
        view_screen_width = target.Right - target.left
        view_screen_height = target.bottom - target.top
    End If
    '*******************************

    '*******************************
    'Erase the backbuffer so that it can be drawn on again
    Device_Clear
    '*******************************
    
    '*******************************
    'Start the scene
    ddevice.BeginScene
    '*******************************
    
    '*******************************
    'Render lights
    Light_Render_All
    '*******************************
    
    '*******************************
    'Draw Map
    Map_Render
    '*******************************
    
    engine_render_started = True
Exit Function
ErrorHandler:
    Engine_Render_Start = False
    MsgBox "Error in Engine_Render_Start: " & Err.Number & ": " & Err.Description
End Function

Public Function Engine_Render_End() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/29/2003
'
'**************************************************************
'On Error GoTo ErrorHandler:
    Engine_Render_End = True

    If engine_render_started = False Then
        Exit Function
    End If

    '*******************************
    'Draw engine stats
    If engine_show_stats Then
        Engine_Stats_Render
    End If
    '*******************************
    
    '*******************************
    'End scene
    ddevice.EndScene
    '*******************************
    
    '*******************************
    'Flip the backbuffer to the screen
    Device_Flip
    '*******************************

    '*******************************
    'Calculate current frames per second
    If GetTickCount >= (fps_last_time + 1000) Then
        fps = fps_frame_counter
        fps_frame_counter = 0
        fps_last_time = GetTickCount
    Else
        fps_frame_counter = fps_frame_counter + 1
    End If
    '*******************************
    
    '*******************************
    'Add one to total frame counter
    total_frame_counter = total_frame_counter + 1
    'If, for some reason, it actually gets to 2 billon, reset to avoid overflow
    If total_frame_counter = 2000000000 Then
        total_frame_counter = 0
    End If
    '*******************************
    
    'Get timing info
    timer_elapsed_time = General_Get_Elapsed_Time()
    timer_ticks_per_frame = (timer_elapsed_time * engine_base_speed)
    
    '*******************************
    'Poll Inputs
    Input_Mouse_Poll
    '*******************************
    
    engine_render_started = False
Exit Function
ErrorHandler:
    Engine_Render_End = False
    MsgBox "Error in Engine_Render_End: " & Err.Number & ": " & Err.Description
End Function

Private Sub Device_Flip()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Draw the graphics to the front buffer.
    ddevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
End Sub

Private Sub Device_Box_Textured_Render(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal src_width As Long, _
                                            ByVal src_height As Long, ByRef rgb_list() As Long, ByVal src_x As Long, _
                                            ByVal src_y As Long, ByVal dest_width As Long, ByVal dest_height As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height - 1
        .left = src_x
        .Right = src_x + src_width - 1
        .top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, rgb_list(), texture_list(texture_index).dimension, angle
    
    'Set Texture
    'ddevice.SetTexture 0, Nothing
    ddevice.SetTexture 0, texture_list(texture_index).d3dtexture
    
    'Enable alpha-blending
    ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Else
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    'Turn off alphablending after we're done
    ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub

Private Sub Device_Clear()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Clear the back buffer
    ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 0, 0
End Sub

Private Function Device_Reset() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Resets the device
'**************************************************************
'On Error GoTo ErrHandler:
'On Error Resume Next
    ddevice.Reset d3dpp
    
    If Err.Number Then
        Device_Reset = Err.Number
        Exit Function
    End If
    
    With ddevice
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        Call .SetVertexShader(FVF)
        
        'Turn off lighting
        Call .SetRenderState(D3DRS_LIGHTING, 0)
                                    
        'Set the render state that uses the alpha component as the source for blending.
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            
        'Set the render state that uses the inverse alpha component as the destination blend.
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    End With
Exit Function
ErrHandler:
    Device_Reset = Err.Number
End Function

Private Sub Device_Text_Render(ByVal font As D3DXFont, ByVal text As String, ByVal top As Long, ByVal left As Long, _
                                ByVal width As Long, ByVal height As Long, ByVal color As Long, ByVal format As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim TextRect As RECT 'This defines where it will be
    
    TextRect.top = top
    TextRect.left = left
    TextRect.bottom = top + height
    TextRect.Right = left + width
    d3dx.DrawText font, color, text, TextRect, format
End Sub

Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.z = z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.specular = specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef texture_size As Long, Optional ByVal angle As Single)
'**************************************************************
'Authors: Aaron Perkins;
'Last Modify Date: 5/07/2002
'
' * v1 *    v3
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' * v0 *    v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    'Center coordinates on screen of the square
    x_center = dest.left + (dest.Right - dest.left - 1) / 2
    y_center = dest.top + (dest.bottom - dest.top - 1) / 2
    
    'Calculate radius
    radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
    
    'Calculate left and right points
    temp = (dest.Right - x_center) / radius
    right_point = Atn(temp / Sqr(-temp * temp + 1))
    left_point = PI - right_point
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If texture_size Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / texture_size, (src.bottom + 1) / texture_size)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If texture_size Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / texture_size, src.top / texture_size)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If texture_size Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / texture_size, (src.bottom + 1) / texture_size)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
    End If
        'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If texture_size Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / texture_size, src.top / texture_size)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 0, 0)
    End If
End Sub

Private Sub Texture_Load_From_File(ByVal texture_index As Long, ByVal file As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim surface_desc As D3DSURFACE_DESC
    
    Set texture_list(texture_index).d3dtexture = d3dx.CreateTextureFromFileEx(ddevice, file, D3DX_DEFAULT, _
    D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
    D3DX_FILTER_POINT, D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)
    
    'Get the width and height of the texture
    texture_list(texture_index).d3dtexture.GetLevelDesc 0, surface_desc
    texture_list(texture_index).dimension = surface_desc.width - 1
End Sub

Private Sub Texture_Load_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'Loads all the graphics into memory
'**************************************************************
    Dim i As Long
    Dim inipath As String
    Dim filepath As String
    
    inipath = resource_path & PATH_GRAPHICS & "\"
    
    'Get Number of Graphics
    texture_count = General_Var_Get(inipath & "grh.ini", "INIT", "texture_count")
    
    ReDim Preserve texture_list(1 To texture_count)
    
    For i = 1 To texture_count
        filepath = resource_path & PATH_GRAPHICS & "\grh" & i & ".bmp"
        If General_File_Exists(filepath, vbNormal) Then
            Texture_Load_From_File i, filepath
        End If
    Next i
End Sub

Private Sub Grh_Initialize(ByRef grh As grh, ByVal grh_index As Long, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    'Copy of parameters
    grh.grh_index = grh_index
    grh.alpha_blend = alpha_blend
    grh.angle = angle
    
    'Start it if it's a animated grh
    If grh_list(grh.grh_index).frame_count > 1 Then
        grh.started = True
    Else
        grh.started = False
    End If
    
    'Set frame counters
    grh.frame_counter = 1
    grh.frame_speed = grh_list(grh.grh_index).frame_speed
End Sub

Private Sub Grh_Load_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Loads Grh.dat
'**************************************************************
On Error GoTo ErrorHandler
    Dim grh As Long
    Dim frame As Long
    Dim tempint As Long
    
    Dim inipath As String
    inipath = resource_path & PATH_GRAPHICS & "\"
    
    'Get number of grhs
    grh_count = Val(General_Var_Get(inipath & "grh.ini", "INIT", "grh_count"))
    
    'Resize arrays
    ReDim grh_list(1 To grh_count) As Grh_Data
    
    'Open files
    Open inipath & "grh.dat" For Binary As #1
    Seek #1, 1
    
    'Get Header
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    
    'Fill Grh List
    
    'Get first Grh Number
    Get #1, , grh
    
    Do Until grh <= 0
        
        grh_list(grh).active = True
        
        'Get number of frames
        Get #1, , grh_list(grh).frame_count
        If grh_list(grh).frame_count <= 0 Then GoTo ErrorHandler
        
        If grh_list(grh).frame_count > 1 Then
        
            'Read a animation GRH set
            For frame = 1 To grh_list(grh).frame_count
            
                Get #1, , grh_list(grh).frame_list(frame)
                If grh_list(grh).frame_list(frame) <= 0 Or grh_list(grh).frame_list(frame) > grh_count Then GoTo ErrorHandler
            
            Next frame
        
            Get #1, , grh_list(grh).frame_speed
            If grh_list(grh).frame_speed = 0 Then GoTo ErrorHandler
            
            'Compute width and height
            grh_list(grh).src_height = grh_list(grh_list(grh).frame_list(1)).src_height
            If grh_list(grh).src_height <= 0 Then GoTo ErrorHandler
            
            grh_list(grh).src_width = grh_list(grh_list(grh).frame_list(1)).src_width
            If grh_list(grh).src_width <= 0 Then GoTo ErrorHandler
        
        Else
        
            'Read in normal GRH data
            Get #1, , grh_list(grh).texture_index
            If grh_list(grh).texture_index <= 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_x
            If grh_list(grh).src_x < 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_y
            If grh_list(grh).src_y < 0 Then GoTo ErrorHandler
                
            Get #1, , grh_list(grh).src_width
            If grh_list(grh).src_width <= 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_height
            If grh_list(grh).src_height <= 0 Then GoTo ErrorHandler
            
            grh_list(grh).frame_list(1) = grh
                
        End If
    
        'Get Next Grh Number
        Get #1, , grh
    
    Loop
    '************************************************
    
    Close #1
Exit Sub
ErrorHandler:
    Close #1
    MsgBox "Error while loading the grh.dat! Stopped at GRH number: " & grh
End Sub

Public Sub Grh_Add_GrhList_To_ListBox(ListboxName As ListBox)
'*****************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 4/25/2003
'Add Grh List To Listbox
'*****************************************************************
On Local Error GoTo Cancel
    Dim grh As Long
    If grh_count > 0 Then
      For grh = 1 To grh_count
        If grh_list(grh).frame_count > 0 Then
            ListboxName.AddItem grh
        End If
      Next grh
    End If
Exit Sub
Cancel:
MsgBox "Some sort of error", vbCritical
End Sub


Public Function Grh_Count_Get() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Gets the total number of grhs
'*****************************************************************
    Grh_Count_Get = grh_count
End Function

Public Function Grh_Info_Get(ByVal grh_index As Long, ByRef file_path As String, ByRef src_x As Long, ByRef src_y As Long, ByRef src_width As Long, ByRef src_height As Long, ByRef frame_count As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/02/2003
'Gets information about a grh
'*****************************************************************
    If Grh_Check(grh_index) = False Then
        Exit Function
    End If
    
    frame_count = grh_list(grh_index).frame_count
    
    'If it's animated switch grh_index to first frame
    If grh_list(grh_index).frame_count <> 1 Then
        grh_index = grh_list(grh_index).frame_list(1)
    End If

    file_path = resource_path & PATH_GRAPHICS & "\grh" & grh_list(grh_index).texture_index & ".bmp"
    src_x = grh_list(grh_index).src_x
    src_y = grh_list(grh_index).src_y
    src_width = grh_list(grh_index).src_width
    src_height = grh_list(grh_index).src_height
       
    Grh_Info_Get = True
End Function

Private Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grh_count Then
        If grh_list(grh_index).active Then
            Grh_Check = True
        End If
    End If
End Function

Private Sub Grh_Uninitialize(ByRef grh As grh)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Resets a Grh
'*****************************************************************
    'Copy of parameters
    grh.grh_index = 0
    grh.alpha_blend = False
    grh.angle = 0
    grh.started = False
    'Set frame counters
    grh.frame_counter = 0
    grh.frame_speed = 0
End Sub

Private Sub Grh_Render(ByRef grh As grh, ByVal screen_x As Long, ByVal screen_y As Long, ByRef rgb_list() As Long, Optional ByVal center As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Long
    Dim tile_height As Long
    Dim grh_index As Long
    
    'Animation
    If grh.started Then
        grh.frame_counter = grh.frame_counter + (timer_ticks_per_frame * grh.frame_speed)
        If grh.frame_counter > grh_list(grh.grh_index).frame_count Then
            If grh.noloop Then
                grh.frame_counter = grh_list(grh.grh_index).frame_count
            Else
                grh.frame_counter = 1
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If grh.frame_counter = 0 Then grh.frame_counter = 1
    grh_index = grh_list(grh.grh_index).frame_list(grh.frame_counter)
    
    'Center Grh over X,Y pos
    If center Then
        tile_width = grh_list(grh_index).src_width / base_tile_size
        tile_height = grh_list(grh_index).src_height / base_tile_size
        If tile_width <> 1 Then
            screen_x = screen_x - Int(tile_width * base_tile_size / 2) + base_tile_size / 2
        End If
        If tile_height <> 1 Then
            screen_y = screen_y - Int(tile_height * base_tile_size) + base_tile_size
        End If
    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_list(grh_index).texture_index, _
        screen_x, screen_y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        rgb_list, _
        grh_list(grh_index).src_x, grh_list(grh_index).src_y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        grh.alpha_blend, _
        grh.angle
End Sub

Public Sub Grh_Render_To_Hdc(ByVal grh_index As Long, desthdc As Long, ByVal screen_x As Long, ByVal screen_y As Long, Optional transparent As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/02/2003
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************
    If Grh_Check(grh_index) = False Then
        Exit Sub
    End If

    Dim file_path As String
    Dim src_x As Long
    Dim src_y As Long
    Dim src_width As Long
    Dim src_height As Long
    Dim hdcsrc As Long

    'If it's animated switch grh_index to first frame
    If grh_list(grh_index).frame_count <> 1 Then
        grh_index = grh_list(grh_index).frame_list(1)
    End If

    file_path = resource_path & PATH_GRAPHICS & "\grh" & grh_list(grh_index).texture_index & ".bmp"
    src_x = grh_list(grh_index).src_x
    src_y = grh_list(grh_index).src_y
    src_width = grh_list(grh_index).src_width
    src_height = grh_list(grh_index).src_height

    hdcsrc = CreateCompatibleDC(desthdc)
    SelectObject hdcsrc, LoadPicture(file_path)
    
    If transparent = False Then
        BitBlt desthdc, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, SRCCOPY
    Else
        TransparentBlt desthdc, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, src_width, src_height, COLOR_KEY
    End If
        
    DeleteDC hdcsrc

End Sub

Public Function Input_Mouse_Moved_Get() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Moved_Get = input_mouse_moved
    input_mouse_moved = False
End Function

Private Sub Input_Mouse_Poll()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next
    'Get the mouse event buffer
    Dim DevData(1 To 5000) As DIDEVICEOBJECTDATA
    Dim nEvents As Long
    nEvents = ddevice_mouse.GetDeviceData(DevData, DIGDD_DEFAULT)
    
    'Check buffer for clicks
    Dim i As Long
    For i = 1 To nEvents
        Select Case DevData(i).lOfs
            Case DIMOFS_BUTTON0
                If input_mouse_button_left Then
                    input_mouse_button_left = False
                Else
                    input_mouse_button_left = True
                End If
                
            Case DIMOFS_BUTTON1
                If input_mouse_button_right Then
                    input_mouse_button_right = False
                Else
                    input_mouse_button_right = True
                End If
        End Select
    Next i
    
    'Use a API to get the mouse cordinates.
    Dim temp_point As PointAPI
    Dim temp_x As Long
    Dim temp_y As Long
    GetCursorPos temp_point
    
    'See if we moved
    If input_mouse_screen_x <> temp_point.x Or input_mouse_screen_y <> temp_point.y Then
        
        'We moved
        'Store new screen position
        input_mouse_screen_x = temp_point.x
        input_mouse_screen_y = temp_point.y
        
        'See if we are in the view area
        Convert_Screen_To_View input_mouse_screen_x, input_mouse_screen_y, input_mouse_view_x, input_mouse_view_y
                  
        'Figure out where on the map was clicked
        Convert_View_To_Map input_mouse_view_x, input_mouse_view_y, input_mouse_map_x, input_mouse_map_y
     
        'Figure out what direction from center of screen to where was clicked
        Convert_Map_To_Direction input_mouse_map_x, input_mouse_map_y, temp_x, temp_y
        
        'Figure out heading from center of screen
        input_mouse_heading = Convert_Direction_To_Heading(temp_x, temp_y)
           
        input_mouse_moved = True
    End If
End Sub

Public Sub Input_Mouse_Screen_Get(ByRef screen_x As Long, ByRef screen_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    screen_x = input_mouse_screen_x
    screen_y = input_mouse_screen_y
End Sub

Public Sub Input_Mouse_View_Get(ByRef view_x As Long, ByRef view_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    view_x = input_mouse_view_x
    view_y = input_mouse_view_y
End Sub

Public Function Input_Mouse_In_View() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    If input_mouse_view_x >= 0 And input_mouse_view_x < view_screen_width And input_mouse_view_y >= 0 And input_mouse_view_y < view_screen_height Then
        Input_Mouse_In_View = True
    End If
End Function

Public Function Input_Mouse_Button_Left_Get() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Button_Left_Get = input_mouse_button_left
End Function

Public Function Input_Mouse_Button_Right_Get() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Button_Right_Get = input_mouse_button_right
End Function

Public Function Input_Mouse_Heading_Get() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Heading_Get = input_mouse_heading
End Function

Public Sub Input_Mouse_Map_Get(ByRef map_x As Long, ByRef map_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    map_x = input_mouse_map_x
    map_y = input_mouse_map_y
End Sub

Public Function Input_Key_Get(ByVal key_code As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    If GetKeyState(key_code) < 0 Then
        Input_Key_Get = True
    End If
End Function

Private Function NPC_Ini_Char_Data_Index_Get(ByVal s_npc_data_index As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    NPC_Ini_Char_Data_Index_Get = CLng(General_Var_Get(resource_path & PATH_SCRIPTS & "\npc.ini", "NPC" & s_npc_data_index, "npc_char_data_index"))
End Function

Private Function Item_Ini_Grh_Index_Get(ByVal s_item_data_index As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    Item_Ini_Grh_Index_Get = CLng(General_Var_Get(resource_path & PATH_SCRIPTS & "\item.ini", "ITEM" & s_item_data_index, "item_grh_index"))
End Function

Public Function Map_NPC_Add(ByVal s_map_x As Long, ByVal s_map_y As Long, ByVal s_npc_data_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_Legal_Char_Pos(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).npc_index = 0 Then
            ReDim Preserve npc_list(0 To UBound(npc_list) + 1)
            npc_list(UBound(npc_list)).npc_data_index = s_npc_data_index
            npc_list(UBound(npc_list)).c_char_data_index = NPC_Ini_Char_Data_Index_Get(s_npc_data_index)
            npc_list(UBound(npc_list)).c_map_x = s_map_x
            npc_list(UBound(npc_list)).c_map_y = s_map_y
            map_current.map_grid(s_map_x, s_map_y).npc_index = UBound(npc_list)
            Map_NPC_Add = True
        End If
    End If
End Function

Public Function Map_NPC_Remove(ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).npc_index <> 0 Then
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).npc_data_index = 0
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).c_char_data_index = 0
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).c_map_x = 0
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).c_map_y = 0
            map_current.map_grid(s_map_x, s_map_y).npc_index = 0
            Map_NPC_Remove = True
        End If
    End If
End Function

Public Function Map_NPC_Get(ByVal s_map_x As Long, ByVal s_map_y As Long, ByRef r_npc_data_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'Returns NPC's data index
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).npc_index <> 0 Then
            r_npc_data_index = npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).npc_data_index
            Map_NPC_Get = True
        End If
    End If
End Function

Public Function Map_Exit_Add(ByVal s_map_x As Long, ByVal s_map_y As Long, ByVal s_exit_map_name As String, ByVal s_exit_map_x As Long, ByVal s_exit_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).exit_index = 0 Then
            ReDim Preserve exit_list(0 To UBound(exit_list) + 1)
            exit_list(UBound(exit_list)).exit_map_name = s_exit_map_name
            exit_list(UBound(exit_list)).exit_map_x = s_exit_map_x
            exit_list(UBound(exit_list)).exit_map_y = s_exit_map_y
            exit_list(UBound(exit_list)).c_map_x = s_map_x
            exit_list(UBound(exit_list)).c_map_y = s_map_y
            map_current.map_grid(s_map_x, s_map_y).exit_index = UBound(exit_list)
            Map_Exit_Add = True
        End If
    End If
End Function

Public Function Map_Exit_Remove(ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).exit_index <> 0 Then
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_name = ""
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_x = 0
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_y = 0
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).c_map_x = 0
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).c_map_y = 0
            map_current.map_grid(s_map_x, s_map_y).exit_index = 0
            Map_Exit_Remove = True
        End If
    End If
End Function

Public Function Map_Exit_Get(ByVal s_map_x As Long, ByVal s_map_y As Long, ByRef r_exit_map_name As String, ByRef r_exit_map_x As Long, ByRef r_exit_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'Returns exit information
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).exit_index <> 0 Then
            r_exit_map_name = exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_name
            r_exit_map_x = exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_x
            r_exit_map_y = exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_y
            Map_Exit_Get = True
        End If
    End If
End Function

Public Function Map_Item_Add(ByVal s_map_x As Long, ByVal s_map_y As Long, ByVal s_item_data_index As Long, ByVal s_item_amount As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).item_index = 0 Then
            ReDim Preserve item_list(0 To UBound(item_list) + 1)
            item_list(UBound(item_list)).item_data_index = s_item_data_index
            item_list(UBound(item_list)).item_amount = s_item_amount
            item_list(UBound(item_list)).c_grh_index = Item_Ini_Grh_Index_Get(s_item_data_index)
            item_list(UBound(item_list)).c_map_x = s_map_x
            item_list(UBound(item_list)).c_map_y = s_map_y
            map_current.map_grid(s_map_x, s_map_y).item_index = UBound(item_list)
            Map_Item_Add = True
        End If
    End If
End Function

Public Function Map_Item_Remove(ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).item_index <> 0 Then
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_data_index = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_amount = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).c_grh_index = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).c_map_x = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).c_map_y = 0
            map_current.map_grid(s_map_x, s_map_y).item_index = 0
            Map_Item_Remove = True
        End If
    End If
End Function

Public Function Map_Item_Get(ByVal s_map_x As Long, ByVal s_map_y As Long, ByRef r_item_data_index As Long, ByRef r_item_amount As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'Returns item information
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).item_index <> 0 Then
            r_item_data_index = item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_data_index
            r_item_amount = item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_amount
            Map_Item_Get = True
        End If
    End If
End Function

Public Function Map_Save_Ini_To_File(ByVal file_path As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    Dim loopc As Long
    Dim counter As Long

   'If file already exists kill it
    If General_File_Exists(file_path, vbNormal) Then
        Kill file_path
    End If
    
    General_Var_Write file_path, "GENERAL", "map_description", map_current.map_description
    
    'NPCs
    counter = 1
    If UBound(npc_list()) <> 0 Then
        For loopc = 1 To UBound(npc_list())
            If npc_list(loopc).npc_data_index Then
                General_Var_Write file_path, "NPC", CStr(counter), CStr(npc_list(loopc).c_map_x) & "-" & CStr(npc_list(loopc).c_map_y) & "-" & CStr(npc_list(loopc).npc_data_index)
                counter = counter + 1
            End If
        Next loopc
    End If
    General_Var_Write file_path, "NPC", "count", CStr(counter - 1)
    
    'Exits
    counter = 1
    If UBound(exit_list()) <> 0 Then
        For loopc = 1 To UBound(exit_list())
            If exit_list(loopc).exit_map_name <> "" Then
                General_Var_Write file_path, "EXIT", CStr(counter), CStr(exit_list(loopc).c_map_x) & "-" & CStr(exit_list(loopc).c_map_y) & "-" & exit_list(loopc).exit_map_name & "-" & CStr(exit_list(loopc).exit_map_x) & "-" & CStr(exit_list(loopc).exit_map_y)
                counter = counter + 1
            End If
        Next loopc
    End If
    General_Var_Write file_path, "EXIT", "count", CStr(counter - 1)
    
    'Items
    counter = 1
    If UBound(item_list()) <> 0 Then
        For loopc = 1 To UBound(item_list())
            If item_list(loopc).item_data_index Then
                General_Var_Write file_path, "ITEM", CStr(counter), CStr(item_list(loopc).c_map_x) & "-" & CStr(item_list(loopc).c_map_y) & "-" & CStr(item_list(loopc).item_data_index) & "-" & CStr(item_list(loopc).item_amount)
                counter = counter + 1
            End If
        Next loopc
    End If
    General_Var_Write file_path, "ITEM", "count", CStr(counter - 1)
    
    Map_Save_Ini_To_File = True
End Function

Public Function Map_Load_Ini_From_File(ByVal file_path As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    Dim loopc As Long
    Dim t_count As Long
    Dim temp_string As String
    
    map_current.map_description = General_Var_Get(file_path, "GENERAL", "map_description")
        
    'NPC
    t_count = Val(General_Var_Get(file_path, "NPC", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "NPC", CStr(loopc))
        If temp_string <> "" Then
            Map_NPC_Add CLng(General_Field_Read(1, temp_string, Asc("-"))), CLng(General_Field_Read(2, temp_string, Asc("-"))), CLng(General_Field_Read(3, temp_string, Asc("-")))
        End If
    Next loopc
    
    'Item
    t_count = Val(General_Var_Get(file_path, "ITEM", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "ITEM", CStr(loopc))
        If temp_string <> "" Then
            Map_Item_Add CLng(General_Field_Read(1, temp_string, Asc("-"))), CLng(General_Field_Read(2, temp_string, Asc("-"))), CLng(General_Field_Read(3, temp_string, Asc("-"))), CLng(General_Field_Read(4, temp_string, Asc("-")))
        End If
    Next loopc
    
    'Exit
    t_count = Val(General_Var_Get(file_path, "EXIT", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "EXIT", CStr(loopc))
        If temp_string <> "" Then
            Map_Exit_Add CLng(General_Field_Read(1, temp_string, Asc("-"))), CLng(General_Field_Read(2, temp_string, Asc("-"))), General_Field_Read(3, temp_string, Asc("-")), CLng(General_Field_Read(4, temp_string, Asc("-"))), CLng(General_Field_Read(5, temp_string, Asc("-")))
        End If
    Next loopc
    
    Map_Load_Ini_From_File = True
End Function

Public Function Map_Save_Map(ByVal map_name As String, Optional ByVal save_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Save map data to files
'*****************************************************************
    Dim map_path As String
    
    'Get map file path
    map_path = resource_path & PATH_MAPS & "\map" & map_name & ".map"
    
    Map_Save_Map = Map_Save_Map_To_File(map_path, save_ini)
End Function

Public Function Map_Save_Map_To_File(ByVal file_path As String, Optional ByVal save_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Save map data to files
'*****************************************************************
    Dim loopc As Long
    Dim LoopC2 As Long
    Dim TempLng As Long
    Dim y As Long
    Dim x As Long
    Dim map_path As String
    
    'Get map file path
    map_path = file_path
    
    'If file already exists kill it
    If General_File_Exists(map_path, vbNormal) Then
        Kill map_path
    End If
    
    'Open .map file
    Open map_path For Binary As #1
    Seek #1, 1
    
    'map Header
    Put #1, , map_current.map_x_min
    Put #1, , map_current.map_x_max
    Put #1, , map_current.map_y_min
    Put #1, , map_current.map_y_max
    Put #1, , TempLng
    Put #1, , TempLng
    Put #1, , TempLng
    Put #1, , TempLng
    Put #1, , TempLng
    
    'Write .map file
    For y = map_current.map_y_min To map_current.map_y_max
        For x = map_current.map_x_min To map_current.map_x_max
            
            '.map file
            
            'Blocked
            Put #1, , map_current.map_grid(x, y).blocked
            
            'Layers
            For loopc = 1 To 4
                Put #1, , map_current.map_grid(x, y).grh(loopc).grh_index
                Put #1, , map_current.map_grid(x, y).grh(loopc).alpha_blend
                Put #1, , map_current.map_grid(x, y).grh(loopc).angle
            Next loopc
            
            'Light base values
            For loopc = 0 To 3
                Put #1, , map_current.map_grid(x, y).light_base_value(loopc)
            Next loopc
            
            'Empty place holders for future expansion
            Put #1, , TempLng
            Put #1, , TempLng
            Put #1, , TempLng
            Put #1, , TempLng
            Put #1, , TempLng
            Put #1, , TempLng
            Put #1, , TempLng
            
        Next x
    Next y
    
    'Write footer
    
    'Lights
    Put #1, , light_count
    For loopc = 1 To light_last
        If light_list(loopc).active Then
            Put #1, , light_list(loopc).map_x
            Put #1, , light_list(loopc).map_y
            Put #1, , light_list(loopc).color
            Put #1, , light_list(loopc).range
        End If
    Next loopc
    
    'Particle Groups
    Put #1, , particle_group_count
    For loopc = 1 To particle_group_last
        If particle_group_list(loopc).active And particle_group_list(loopc).never_die Then
            Put #1, , particle_group_list(loopc).map_x
            Put #1, , particle_group_list(loopc).map_y
            Put #1, , particle_group_list(loopc).particle_count
            Put #1, , particle_group_list(loopc).stream_type
            Put #1, , particle_group_list(loopc).alpha_blend
            Put #1, , particle_group_list(loopc).frame_speed
            Put #1, , particle_group_list(loopc).grh_index_count
            For LoopC2 = 1 To particle_group_list(loopc).grh_index_count
                Put #1, , particle_group_list(loopc).grh_index_list(LoopC2)
            Next LoopC2
        End If
    Next loopc
    
    'Save ini if needed
    If save_ini Then
        Map_Save_Ini_To_File left$(map_path, Len(map_path) - 3) & "ini"
    End If
    
    'Close .map file
    Close #1
    Map_Save_Map_To_File = True
End Function

Public Function Map_Load_Map(ByVal map_name As String, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/25/2003
'Load map data from file using an map id number
'*****************************************************************
    Dim map_path As String

    'Get map file path
    map_path = resource_path & PATH_MAPS & "\map" & map_name & ".map"
    
    Map_Load_Map = Map_Load_Map_From_File(map_path, load_ini)
End Function

Public Function Map_Load_Map_From_File(ByVal file_path As String, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/25/2003
'Load map data from file using an filepath
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
    
    'If file doesn't exists, exit
    If Not (General_File_Exists(map_path, vbNormal)) Then
        Exit Function
    End If
    
    'Erase Chars, Lights, and Particle Groups
    Char_Remove_All
    Light_Remove_All
    Particle_Group_Remove_All
    
    'Open .map file
    Open map_path For Binary As #1
    Seek #1, 1
    
    'map Header
    Get #1, , map_current.map_x_min
    Get #1, , map_current.map_x_max
    Get #1, , map_current.map_y_min
    Get #1, , map_current.map_y_max
    Get #1, , TempLng
    Get #1, , TempLng
    Get #1, , TempLng
    Get #1, , TempLng
    Get #1, , TempLng
    
    'Clear out and resize map
    ReDim map_current.map_grid(map_current.map_x_min To map_current.map_x_max, _
                                map_current.map_y_min To map_current.map_y_max) As Map_Tile
    
    'Read .map file
    For y = map_current.map_y_min To map_current.map_y_max
        For x = map_current.map_x_min To map_current.map_x_max
            
            '.map file
            
            'Blocked
            Get #1, , map_current.map_grid(x, y).blocked
            
            'Layers
            For loopc = 1 To 4
                Get #1, , TempLng
                Get #1, , TempBln
                Get #1, , TempSgl
                If TempLng <> 0 Then
                    Grh_Initialize map_current.map_grid(x, y).grh(loopc), TempLng, TempBln, TempSgl
                End If
            Next loopc
            
            'Light values
            For loopc = 0 To 3
                Get #1, , TempLng
                Map_Base_Light_Set x, y, TempLng, loopc
            Next loopc
                        
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
            
            Light_Create TempLng2, TempLng3, TempLng4, TempLng5
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
    
            Particle_Group_Create TempLng2, TempLng3, TempLngList(), TempLng4, TempLng5, TempBln, , TempSgl
    Next loopc
    
    'Load Ini file
    ReDim npc_list(0)
    ReDim item_list(0)
    ReDim exit_list(0)
    If load_ini Then
        Map_Load_Ini_From_File left$(map_path, Len(map_path) - 3) & "ini"
    End If
    
    'Close .map file
    Close #1
    Map_Load_Map_From_File = True
End Function

Public Function Map_Fill(ByVal grh_index As Long, ByVal layer As Long, Optional ByVal light_base_color As Long = -1, _
                        Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    Dim x As Long
    Dim y As Long
    Dim loop_counter As Long
    
    Dim temp_list() As Long
    
    For y = map_current.map_y_min To map_current.map_y_max
        For x = map_current.map_x_min To map_current.map_x_max
        
            'Grh
            If Map_Grh_Set(x, y, grh_index, layer, alpha_blend, angle) = False Then
                Exit Function
            End If
        
            'Base light color
            If light_base_color <> -1 Then
                If Map_Base_Light_Set(x, y, light_base_color) = False Then
                    Exit Function
                End If
            End If
        
        Next x
    Next y
    
    Map_Fill = True
End Function

Public Function Map_Edges_Blocked_Set(ByVal edge_distance_x As Long, ByVal edge_distance_y As Long, ByVal blocked As Boolean) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'
'**************************************************************
    Dim x As Long
    Dim y As Long
    
    For y = map_current.map_y_min To map_current.map_y_max
        For x = map_current.map_x_min To map_current.map_x_max
            If x <= edge_distance_x Or y <= edge_distance_y Then
                map_current.map_grid(x, y).blocked = blocked
            End If

            If x > map_current.map_x_max - edge_distance_x Or y > map_current.map_y_max - edge_distance_y Then
                map_current.map_grid(x, y).blocked = blocked
            End If
        Next x
    Next y
    
    Map_Edges_Blocked_Set = True
End Function

Public Function Map_Create(ByVal map_x_max As Long, ByVal map_y_max As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Set size
    map_current.map_x_min = 1
    map_current.map_x_max = map_x_max
    map_current.map_y_min = 1
    map_current.map_y_max = map_y_max
    
    'Erase Chars, Lights, and Particle Groups
    Char_Remove_All
    Light_Remove_All
    Particle_Group_Remove_All
    
    ReDim npc_list(0)
    ReDim item_list(0)
    ReDim exit_list(0)
    
    'Erase map
    ReDim map_current.map_grid(map_current.map_x_min To map_current.map_x_max, map_current.map_y_min To map_current.map_y_max) As Map_Tile
    
    'Fill in the map with grh 1 so ther is something to render
    Map_Fill 1, 1, &HAAAAAA
    
    Map_Create = True
End Function

Public Function Map_Bounds_Get(ByRef map_x_max As Long, ByRef map_y_max As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'
'**************************************************************
    'Get size
    map_x_max = map_current.map_x_max
    map_y_max = map_current.map_y_max
    Map_Bounds_Get = True
End Function

Public Function Map_Bounds_Get_From_File(ByVal map_name As String, ByRef max_x As Long, ByRef max_y As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/27/2003
'
'**************************************************************
    Dim TempLng As Long
    
    map_name = resource_path & PATH_MAPS & "map" & map_name & ".map"
    
    'If file doesn't exists, exit
    If Not (General_File_Exists(map_name, vbNormal)) Then
        Exit Function
    End If
    
    'Open .map file
    Open map_name For Binary As #1
    Seek #1, 1
    
    'map Header
    Get #1, , TempLng
    Get #1, , max_x
    Get #1, , TempLng
    Get #1, , max_y
    
    Close #1
    
    Map_Bounds_Get_From_File = True
End Function


Public Function Map_Base_Light_Fill(ByVal light_base_color As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim x As Long
    Dim y As Long
    
    For y = map_current.map_y_min To map_current.map_y_max
        For x = map_current.map_x_min To map_current.map_x_max
            
            'Base light color
            If Map_Base_Light_Set(x, y, light_base_color) = False Then
                Exit Function
            End If
    
        Next x
    Next y
    
    Map_Base_Light_Fill = True
End Function

Public Function Map_Grh_Set(ByVal map_x As Long, ByVal map_y As Long, ByVal grh_index As Long, _
                            ByVal layer As Long, Optional ByVal alpha_blend As Boolean, _
                            Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'**************************************************************
    'Check
    If layer < 1 Or layer > 5 Then
        Exit Function
    End If
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    If Grh_Check(grh_index) = False Then
        Exit Function
    End If
    
    'Do it
    Grh_Initialize map_current.map_grid(map_x, map_y).grh(layer), grh_index, alpha_blend, angle
    
    Map_Grh_Set = True
End Function

Public Function Map_Grh_UnSet(ByVal map_x As Long, ByVal map_y As Long, ByVal layer As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'**************************************************************
    Dim grh_index As Long
    
    'Check
    If layer < 1 Or layer > 5 Then
        Exit Function
    End If
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    grh_index = map_current.map_grid(map_x, map_y).grh(layer).grh_index
    If Grh_Check(grh_index) = False Then
        Exit Function
    End If
    
    'Do it
    Grh_Uninitialize map_current.map_grid(map_x, map_y).grh(layer)
    
    Map_Grh_UnSet = True
End Function

Public Function Map_Base_Light_Set(ByVal map_x As Long, ByVal map_y As Long, ByVal light_base_value As Long, Optional corner As Long = -1) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
    
    'Check
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    If corner < -1 Or corner > 3 Then
        Exit Function
    End If
    
    'Do it
    If corner = -1 Then
        'Set all corners
        For loop_counter = 0 To 3
            map_current.map_grid(map_x, map_y).light_base_value(loop_counter) = light_base_value
            map_current.map_grid(map_x, map_y).light_value(loop_counter) = light_base_value
        Next loop_counter
    Else
        'Set just one
        map_current.map_grid(map_x, map_y).light_base_value(corner) = light_base_value
        map_current.map_grid(map_x, map_y).light_value(corner) = light_base_value
    End If
    
    Map_Base_Light_Set = True
End Function


Public Function Map_Base_Light_Get(ByVal map_x As Long, ByVal map_y As Long, Optional corner As Long = 0) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Base_Light_Get = map_current.map_grid(map_x, map_y).light_base_value(corner)
    End If
End Function

Public Function Map_In_Bounds(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If map_x < map_current.map_x_min Or map_x > map_current.map_x_max Or map_y < map_current.map_y_min Or map_y > map_current.map_y_max Then
        Map_In_Bounds = False
        Exit Function
    End If
    
    Map_In_Bounds = True
End Function

Public Function Map_Legal_Char_Pos(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    If map_current.map_grid(map_x, map_y).blocked Then
        Exit Function
    End If
    If map_current.map_grid(map_x, map_y).char_index Then
        Exit Function
    End If
    Map_Legal_Char_Pos = True
End Function

Public Function Map_Legal_Char_Pos_By_Heading(ByVal char_index As Long, ByVal heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    'Invalid heading
    If heading < 1 Or heading > 8 Then
        Exit Function
    End If
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Dim nx As Long
        Dim ny As Long
        nx = char_list(char_index).map_x
        ny = char_list(char_index).map_y
        Convert_Heading_to_Direction heading, nx, ny
            
        Map_Legal_Char_Pos_By_Heading = Map_Legal_Char_Pos(nx, ny)
    End If
End Function
Public Function Map_Char_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a char_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Char_Get = map_current.map_grid(map_x, map_y).char_index
    Else
        Map_Char_Get = 0
    End If
End Function

Public Function Map_Description_Get() As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    Map_Description_Get = map_current.map_description
End Function

Public Function Map_Description_Set(ByVal s_map_description As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    map_current.map_description = s_map_description
    Map_Description_Set = True
End Function

Public Function Map_Blocked_Get(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is blocked
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Blocked_Get = map_current.map_grid(map_x, map_y).blocked
    Else
        Map_Blocked_Get = True
    End If
End Function

Public Function Map_Blocked_Set(ByVal map_x As Long, ByVal map_y As Long, ByVal blocked As Boolean) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sets a tile position to blocked
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        map_current.map_grid(map_x, map_y).blocked = blocked
        Map_Blocked_Set = True
    End If
End Function

Public Function Map_Particle_Group_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Particle_Group_Get = map_current.map_grid(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

Public Function Map_Grh_Get(ByVal map_x As Long, ByVal map_y As Long, ByVal layer As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a grh and return it
'*****************************************************************
    'Check
    If layer < 1 Or layer > 4 Then
        Map_Grh_Get = 0
        Exit Function
    End If
    
    If Map_In_Bounds(map_x, map_y) Then
        Map_Grh_Get = map_current.map_grid(map_x, map_y).grh(layer).grh_index
    Else
        Map_Grh_Get = 0
    End If
End Function

Public Function Map_Light_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a light_index and return it
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).map_x = map_x And light_list(loopc).map_y = map_y
        If loopc = light_last Then
            Map_Light_Get = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Map_Light_Get = loopc
Exit Function
ErrorHandler:
    Map_Light_Get = 0
End Function

Private Sub Map_Render()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim view_x As Long
    Dim view_y As Long
    Dim screen_x As Long
    Dim screen_y As Long
    
    Dim view_min_y As Long
    Dim view_max_y As Long
    Dim view_min_x As Long
    Dim view_max_x As Long
    
    Dim view_buffer_min_y As Long
    Dim view_buffer_max_y As Long
    Dim view_buffer_min_x As Long
    Dim view_buffer_max_x As Long
    
    '*********************
    'Handle scrolling
    'counters
    '*********************
    If scroll_on Then
        '****** Move screen Left and Right if needed ******
        If scroll_direction_x <> 0 Then
            scroll_offset_counter_x = scroll_offset_counter_x + (scroll_pixels_per_frame * timer_ticks_per_frame * scroll_direction_x)
            If Sgn(scroll_offset_counter_x) = scroll_direction_x Then
                scroll_offset_counter_x = 0
                scroll_direction_x = 0
            End If
        End If
        '****** Move screen Up and Down if needed ******
        If scroll_direction_y <> 0 Then
            scroll_offset_counter_y = scroll_offset_counter_y + (scroll_pixels_per_frame * timer_ticks_per_frame * scroll_direction_y)
            If Sgn(scroll_offset_counter_y) = scroll_direction_y Then
                scroll_offset_counter_y = 0
                scroll_direction_y = 0
            End If
        End If
        'End scrolling if needed
        If scroll_direction_x = 0 And scroll_direction_y = 0 Then
            scroll_on = False
        End If
    End If
    
    'Figure out ends and starts of view area
    view_min_y = ((view_pos_y) - (view_screen_tile_height \ 2))
    view_max_y = ((view_pos_y) + (view_screen_tile_height \ 2))
    view_min_x = ((view_pos_x) - (view_screen_tile_width \ 2))
    view_max_x = ((view_pos_x) + (view_screen_tile_width \ 2))
    
    'Add the buffer
    view_buffer_min_y = view_min_y - view_tile_buffer
    view_buffer_max_y = view_max_y + view_tile_buffer
    view_buffer_min_x = view_min_x - view_tile_buffer
    view_buffer_max_x = view_max_x + view_tile_buffer
    
    '*********************
    'Layer 1
    '*********************
    view_y = 0
    For map_y = view_min_y - 1 To view_max_y + 1
        view_x = 0
        For map_x = view_min_x - 1 To view_max_x + 2
    
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start Layer 1 ***
                If map_current.map_grid(map_x, map_y).grh(1).grh_index Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(1), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value()
                End If
                '*** End Layer 1 ***
            End If
            
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Layer 2 and 5
    '*********************
    view_y = -1 * view_tile_buffer + 1
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = -1 * view_tile_buffer + 1
        For map_x = view_buffer_min_x To view_buffer_max_x
    
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start Layer 2 ***
                If map_current.map_grid(map_x, map_y).grh(2).grh_index Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(2), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value(), True
                End If
                '*** End Layer 2 ***
                '*** Start Layer 5 *** 'Special layer that is not saved and used for items
               If map_current.map_grid(map_x, map_y).grh(5).grh_index Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(5), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value(), True
                End If
                '*** End Layer 5 ***
            End If
            
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Middle layer
    '*********************
    view_y = -1 * view_tile_buffer + 1
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = -1 * view_tile_buffer + 1
        For map_x = view_buffer_min_x To view_buffer_max_x
    
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start Layer 3 ***
                If map_current.map_grid(map_x, map_y).grh(3).grh_index Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(3), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value(), True
                End If
                '*** End Layer 3 ***
                '*** Start Characters ***
                If map_current.map_grid(map_x, map_y).char_index Then
                    'Figure out screen position
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Char_Render char_list(map_current.map_grid(map_x, map_y).char_index), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value
                End If
                '*** End Characters ***
            End If
                
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Layer 4
    '*********************
    view_y = -1 * view_tile_buffer + 1
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = -1 * view_tile_buffer + 1
        For map_x = view_buffer_min_x To view_buffer_max_x
    
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start particle effects ***
                If map_current.map_grid(map_x, map_y).particle_group_index Then
                    screen_x = (Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x) + base_tile_size / 2
                    screen_y = (Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y) + base_tile_size / 2
                    Particle_Group_Render map_current.map_grid(map_x, map_y).particle_group_index, screen_x, screen_y
                End If
                '*** End particle effects ***
                '*** Start Layer 4 ***
                If map_current.map_grid(map_x, map_y).grh(4).grh_index Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(4), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value(), True
                End If
                '*** End Layer 4 ***
            End If
            
            view_x = view_x + 1
        Next map_x
        
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Special Layer
    '*********************
    If engine_show_blocked_tiles Or engine_show_special_tiles Then
        Dim grh As grh
        Dim temp_array(3) As Long
        temp_array(0) = &HFFFFFF
        temp_array(1) = &HFFFFFF
        temp_array(2) = &HFFFFFF
        temp_array(3) = &HFFFFFF
        view_y = 0
        For map_y = view_min_y - 1 To view_max_y + 1
            view_x = 0
            For map_x = view_min_x - 1 To view_max_x + 2
        
                If Map_In_Bounds(map_x, map_y) Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    
                    If engine_show_blocked_tiles Then
                        '*** Start Blocked Tiles ***
                        If map_current.map_grid(map_x, map_y).blocked Then
                            grh.grh_index = 2
                            Grh_Render grh, screen_x, screen_y, temp_array
                        End If
                        '*** End Blocked Tiles ***
                    End If
                    
                   If engine_show_special_tiles Then
                        '*** Start Special Tiles ***
                        If map_current.map_grid(map_x, map_y).exit_index Then
                            grh.grh_index = 3
                            Grh_Render grh, screen_x, screen_y, temp_array
                        End If
                        If map_current.map_grid(map_x, map_y).item_index Then
                            grh.grh_index = item_list(map_current.map_grid(map_x, map_y).item_index).c_grh_index
                            Grh_Render grh, screen_x, screen_y, temp_array, True
                        End If
                        If map_current.map_grid(map_x, map_y).npc_index Then
                            grh.grh_index = Char_Data_Grh_Index_Get(npc_list(map_current.map_grid(map_x, map_y).npc_index).c_char_data_index, 1, 5)
                            Grh_Render grh, screen_x, screen_y, temp_array, True
                        End If
                        '*** End Special Tiles ***
                    End If
                End If
                
                view_x = view_x + 1
            Next map_x
            view_y = view_y + 1
        Next map_y
    End If
    
    '*********************
    'Clip around edges
    'to clean up
    '*********************
    If engine_windowed = False Then
        ddevice.Clear 4, clip_rect(0), D3DCLEAR_TARGET, engine_clip_border_color, 0, 0
    End If
End Sub

Public Function Font_Create(ByVal style As String, ByVal size As Long, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'Returns the font_index if successful, else 0
'**************************************************************
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    
    Font_Make Font_Create, style, bold, italic, size
ErrorHandler:
    Font_Create = 0
End Function

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal size As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont 'We use this temporarily to setup the font
    Dim fnt As New StdFont 'This is also used to describe and setup the font
    fnt.Name = style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    
    Set font_desc = fnt
    Set font_list(font_index) = d3dx.CreateFont(ddevice, font_desc.hFont)
End Sub

Private Function Font_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'
'**************************************************************
    'check font_index
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
End Function

Public Function Char_Create(ByVal map_x As Long, ByVal map_y As Long, ByVal heading As Long, _
                            ByVal char_data_index As Long, Optional ByVal id As Long = 0) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Returns the char_index if successful, else 0
'**************************************************************
    'Invalid heading
    If heading < 1 Or heading > 8 Then
        Exit Function
    End If

    If Map_Char_Get(map_x, map_y) = 0 Then
        Char_Create = Char_Next_Open
        Char_Make Char_Create, map_x, map_y, heading, char_data_index, 1, id
    End If
End Function

Private Function Char_Check(ByVal char_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= char_last Then
        If char_list(char_index).active Then
            Char_Check = True
        End If
    End If
End Function

Private Function Char_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until char_list(loopc).active = False
        If loopc = char_last Then
            Char_Next_Open = char_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Next_Open = loopc
Exit Function
ErrorHandler:
    Char_Next_Open = 1
End Function

Public Function Char_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until char_list(loopc).id = id
        If loopc = char_last Then
            Char_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Find = loopc
Exit Function
ErrorHandler:
    Char_Find = 0
End Function

Public Function Char_Move(ByVal char_index As Long, ByVal heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    Dim temp_x As Long
    Dim temp_y As Long
    
    'Invalid heading
    If heading < 1 Or heading > 8 Then
        Char_Move = False
        Exit Function
    End If
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Make sure it's a legal move
        temp_x = char_list(char_index).map_x
        temp_y = char_list(char_index).map_y
        Convert_Heading_to_Direction heading, temp_x, temp_y
        If Map_In_Bounds(temp_x, temp_y) Then
        
            'check for another char_index
            If map_current.map_grid(temp_x, temp_y).char_index = 0 Then
                'Move it
                Char_Move_By_Heading char_index, heading
                Char_Move = True
            End If
            
        End If
    End If
End Function

Public Function Char_Remove(ByVal char_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Char_Check(char_index) Then
        Char_Destroy char_index
        Char_Remove = True
    End If
End Function

Public Function Char_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
    
    For index = 1 To char_last
        'Make sure it's a legal index
        If Char_Check(index) Then
            Char_Destroy index
        End If
    Next index
    
    Char_Remove_All = True
End Function

Private Sub Char_Make(ByVal char_index As Long, ByVal map_x As Long, ByVal map_y As Long, ByVal heading As Long, _
                        ByVal char_data_index As Long, ByVal body_index As Long, ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Makes a new character and puts it on the map
'*****************************************************************
    'Update array size
    If char_index > char_last Then
        char_last = char_index
        ReDim Preserve char_list(1 To char_last)
    End If
    char_count = char_count + 1
    
    'Make active
    char_list(char_index).active = True
    
    'Heading
    char_list(char_index).heading = heading
    
    'Update char data
    char_list(char_index).chr_data = char_data_list(char_data_index)
    char_list(char_index).chr_data_body_index = body_index

    'Label Offset
    char_list(char_index).label_offset_x = char_data_list(char_data_index).label_offset_x
    char_list(char_index).label_offset_y = char_data_list(char_data_index).label_offset_y
    
    'Reset moving stats
    char_list(char_index).scroll_on = False
    char_list(char_index).scroll_direction_x = 0
    char_list(char_index).scroll_direction_y = 0
    char_list(char_index).scroll_offset_counter_y = 0
    char_list(char_index).scroll_offset_counter_x = 0
    
    'Update position
    char_list(char_index).map_x = map_x
    char_list(char_index).map_y = map_y
    
    'id
    char_list(char_index).id = id
    
    'Plot on map
    map_current.map_grid(map_x, map_y).char_index = char_index
End Sub

Public Function Char_Label_Set(ByVal char_index As Long, ByVal label As String, ByVal label_font_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'Changes the character label
'*****************************************************************
    'Check font
    If Font_Check(label_font_index) = False Then
        Exit Function
    End If
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Label
        char_list(char_index).label = label
        char_list(char_index).label_font_index = label_font_index
        Char_Label_Set = True
    End If
End Function

Public Function Char_Data_Grh_Index_Get(ByVal char_data_index As Long, ByVal s_body_index As Long, ByVal s_heading As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/12/2003
'Returns the grh_index for a character given a body_index and heading
'*****************************************************************
    'Invalid heading
    If s_heading < 1 Or s_heading > 8 Then
        Exit Function
    End If

    'Make sure it's a legal char_index
    If char_data_index <= UBound(char_data_list) And char_data_index >= LBound(char_data_list) Then
        Char_Data_Grh_Index_Get = char_data_list(char_data_index).body_list(s_body_index).body(s_heading).grh_index
    End If
End Function

Public Function Char_Data_Set(ByVal char_index As Long, ByVal char_data_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Changes the character data
'*****************************************************************
    Dim loopc As Long
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Update char data
        char_list(char_index).chr_data = char_data_list(char_data_index)
        char_list(char_index).chr_data_body_index = 1
        
        'Label Offset
        char_list(char_index).label_offset_x = char_data_list(char_data_index).label_offset_x
        char_list(char_index).label_offset_y = char_data_list(char_data_index).label_offset_y
        
        Char_Data_Set = True
    End If
End Function

Public Function Char_Data_Body_Set(ByVal char_index As Long, ByVal body_index As Long, ByVal noloop As Boolean) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Changes the character label
'*****************************************************************
    Dim loopc As Long
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Update char data
        char_list(char_index).chr_data_body_index = body_index
        'Set loop status
        For loopc = 1 To 8
            char_list(char_index).chr_data.body_list(body_index).body(loopc).frame_counter = 1
            char_list(char_index).chr_data.body_list(body_index).body(loopc).noloop = noloop
        Next loopc
        
        Char_Data_Body_Set = True
    End If
End Function

Public Function Char_Heading_Set(ByVal char_index As Long, ByVal heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Changes the character heading
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        char_list(char_index).heading = heading
        Char_Heading_Set = True
        Exit Function
    End If
End Function

Public Function Char_Map_Pos_Get(ByVal char_index As Long, ByRef map_x As Long, ByRef map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Changes the character label
'*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Get map pos
        map_x = char_list(char_index).map_x
        map_y = char_list(char_index).map_y
        
        Char_Map_Pos_Get = True
    End If
End Function

Public Function Char_Map_Pos_Set(ByVal char_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Changes the character label
'*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Check map pos
        If Map_In_Bounds(map_x, map_y) Then
            'Move char
            map_current.map_grid(char_list(char_index).map_x, char_list(char_index).map_y).char_index = 0
             char_list(char_index).map_x = map_x
             char_list(char_index).map_y = map_y
            map_current.map_grid(char_list(char_index).map_x, char_list(char_index).map_y).char_index = char_index
            Char_Map_Pos_Set = True
        End If
    End If
End Function

Private Sub Char_Move_By_Heading(ByVal char_index As Long, ByVal heading As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim x As Long
    Dim y As Long
    Dim nx As Long
    Dim ny As Long
    
    x = char_list(char_index).map_x
    y = char_list(char_index).map_y
    
    nx = x
    ny = y
    Convert_Heading_to_Direction heading, nx, ny
    
    map_current.map_grid(nx, ny).char_index = char_index
    char_list(char_index).map_x = nx
    char_list(char_index).map_y = ny
    map_current.map_grid(x, y).char_index = 0
    
    char_list(char_index).scroll_offset_counter_x = (base_tile_size * (x - nx))
    char_list(char_index).scroll_offset_counter_y = (base_tile_size * (y - ny))
    char_list(char_index).scroll_direction_x = nx - x
    char_list(char_index).scroll_direction_y = ny - y
    
    char_list(char_index).scroll_on = True
    char_list(char_index).heading = heading
    'Set char to walk
    char_list(char_index).chr_data_body_index = 2
End Sub

Private Sub Char_Destroy(ByVal char_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As char
    
    map_current.map_grid(char_list(char_index).map_x, char_list(char_index).map_y).char_index = 0
    char_list(char_index) = temp
    
    'Update array size
    If char_index = char_last Then
        Do Until char_list(char_last).active
            char_last = char_last - 1
            If char_last = 0 Then
                char_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve char_list(1 To char_last)
    End If
    char_count = char_count - 1
End Sub

Private Sub Char_Load_Char_Data()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Loads Body.dat
'*****************************************************************
    Dim loopc As Long
    Dim LoopC2 As Long
    Dim loopc3 As Long
    Dim inipath As String
    
    inipath = resource_path & PATH_GRAPHICS & "\" & "char.ini"
    
    'Get number of bodies
    char_data_count = Val(General_Var_Get(inipath, "INIT", "char_data_count"))
    
    'Resize array
    ReDim char_data_list(1 To char_data_count) As Char_Data
    
    'Fill list
    For loopc = 1 To char_data_count
    
        'Get list count
        char_data_list(loopc).body_count = Val(General_Var_Get(inipath, "CharData" & loopc, "body_count"))
        
        'Get label offset
        char_data_list(loopc).label_offset_x = Val(General_Var_Get(inipath, "CharData" & loopc, "label_offset_x"))
        char_data_list(loopc).label_offset_y = Val(General_Var_Get(inipath, "CharData" & loopc, "label_offset_y"))
        
        'Redim body list
        ReDim char_data_list(loopc).body_list(1 To char_data_list(loopc).body_count)
        
        'Fill body list
        For LoopC2 = 1 To char_data_list(loopc).body_count
            For loopc3 = 1 To 8
                Grh_Initialize char_data_list(loopc).body_list(LoopC2).body(loopc3), Val(General_Var_Get(inipath, "CharData" & loopc, LoopC2 & "_" & loopc3))
            Next loopc3
        Next LoopC2
    
    Next loopc
End Sub

Private Sub Char_Render(ByRef temp_char As char, ByVal screen_x As Long, ByVal screen_y As Long, ByRef light_value() As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/28/2002
'Renders a character at a paticular screen point
'*****************************************************************
    Dim char_moved As Boolean
    
    'Scrolling
    If temp_char.scroll_on Then
        '****** Move Left and Right if needed ******
        If temp_char.scroll_direction_x <> 0 Then
            temp_char.scroll_offset_counter_x = temp_char.scroll_offset_counter_x + (scroll_pixels_per_frame * timer_ticks_per_frame * Sgn(temp_char.scroll_direction_x))
            If Sgn(temp_char.scroll_offset_counter_x) = temp_char.scroll_direction_x Then
                temp_char.scroll_offset_counter_x = 0
                temp_char.scroll_direction_x = 0
            End If
        End If
        '****** Move Up and Down if needed ******
        If temp_char.scroll_direction_y <> 0 Then
            temp_char.scroll_offset_counter_y = temp_char.scroll_offset_counter_y + (scroll_pixels_per_frame * timer_ticks_per_frame * Sgn(temp_char.scroll_direction_y))
            If Sgn(temp_char.scroll_offset_counter_y) = temp_char.scroll_direction_y Then
                temp_char.scroll_offset_counter_y = 0
                temp_char.scroll_direction_y = 0
            End If
        End If
        'End scrolling if needed
        If temp_char.scroll_direction_x = 0 And temp_char.scroll_direction_y = 0 Then
            'Turn off scrolling
            temp_char.scroll_on = False
            'Set char to stand
            temp_char.chr_data_body_index = 1
        End If
    End If
    
    'Find screen position
    screen_x = screen_x + temp_char.scroll_offset_counter_x
    screen_y = screen_y + temp_char.scroll_offset_counter_y
    
    'Render Grh
    Grh_Render temp_char.chr_data.body_list(temp_char.chr_data_body_index).body(temp_char.heading), screen_x, screen_y, light_value(), True
    
    'Do label
    If temp_char.label <> "" Then
        Device_Text_Render font_list(temp_char.label_font_index), _
                         temp_char.label, _
                         screen_y + temp_char.label_offset_y, _
                         screen_x + temp_char.label_offset_x - base_tile_size, _
                         100, 20, _
                         &HFFFFFFFF, DT_TOP Or DT_CENTER
    End If
End Sub

Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True
    End If
End Function

Public Function Light_Create(ByVal map_x As Long, ByVal map_y As Long, Optional ByVal color_value As Long = &HFFFFFF, _
                            Optional ByVal range As Long = 1, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function

Public Function Light_Map_Pos_Set(ByVal light_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If Map_In_Bounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index ', , , map_current.map_grid(map_x, map_y).light_base_value
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Map_Pos_Set = True
            
        End If
    End If
End Function

Public Function Light_Move(ByVal light_index As Long, ByVal heading As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/25/2003
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim nx As Long
    Dim ny As Long
    
    'Check for valid heading
    If heading < 1 Or heading > 8 Then
        Light_Move = False
        Exit Function
    End If
    
    'Make sure it's a legal index
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        
        nx = map_x
        ny = map_y
        
        Convert_Heading_to_Direction heading, nx, ny
        
        'Make sure it's a legal move
        If Map_In_Bounds(nx, ny) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = nx
            light_list(light_index).map_y = ny
            
            Light_Move = True
        
        End If
    End If
End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Long, ByVal map_y As Long, ByVal rgb_value As Long, _
                        ByVal range As Long, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).active = True
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function

Private Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).active Then
            Light_Render loop_counter
        End If
    
    Next loop_counter
End Sub

Private Sub Light_Render(ByVal light_index, Optional ByVal map_x As Long = -1, Optional ByVal map_y As Long = -1, _
                        Optional ByVal rgb_value As Long = -1, Optional ByVal range As Long = -1)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim x2 As Long
    Dim y2 As Long
    
    If map_x = -1 Then
        map_x = light_list(light_index).map_x
    End If
    If map_y = -1 Then
        map_y = light_list(light_index).map_y
    End If
    If rgb_value = -1 Then
        rgb_value = light_list(light_index).color
    End If
    If range = -1 Then
        range = light_list(light_index).range
    End If
    
    map_current.map_grid(map_x, map_y).light_value(0) = rgb_value
    map_current.map_grid(map_x, map_y).light_value(1) = rgb_value
    map_current.map_grid(map_x, map_y).light_value(2) = rgb_value
    map_current.map_grid(map_x, map_y).light_value(3) = rgb_value
    
    x2 = map_x - 1
    y2 = map_y - 1
    If Map_In_Bounds(x2, y2) Then
        map_current.map_grid(x2, y2).light_value(2) = rgb_value
        If range > 1 Then
            Light_Render light_index, x2, y2, rgb_value, range - 1
        End If
    End If
    
    x2 = map_x - 1
    y2 = map_y
    If Map_In_Bounds(x2, y2) Then
        map_current.map_grid(x2, y2).light_value(2) = rgb_value
        map_current.map_grid(x2, y2).light_value(3) = rgb_value
        If range > 1 Then
            Light_Render light_index, x2, y2, rgb_value, range - 1
        End If
    End If
         
    x2 = map_x - 1
    y2 = map_y + 1
    If Map_In_Bounds(x2, y2) Then
        map_current.map_grid(x2, y2).light_value(3) = rgb_value
        If range > 1 Then
            Light_Render light_index, x2, y2, rgb_value, range - 1
        End If
    End If
    
    x2 = map_x
    y2 = map_y + 1
    If Map_In_Bounds(x2, y2) Then
        map_current.map_grid(x2, y2).light_value(1) = rgb_value
        map_current.map_grid(x2, y2).light_value(3) = rgb_value
        If range > 1 Then
            Light_Render light_index, x2, y2, rgb_value, range - 1
        End If
    End If
    
    x2 = map_x + 1
    y2 = map_y + 1
    If Map_In_Bounds(x2, y2) Then
        map_current.map_grid(x2, y2).light_value(1) = rgb_value
        If range > 1 Then
            Light_Render light_index, x2, y2, rgb_value, range - 1
        End If
    End If
    
    x2 = map_x + 1
    y2 = map_y
    If Map_In_Bounds(x2, y2) Then
        map_current.map_grid(x2, y2).light_value(1) = rgb_value
        map_current.map_grid(x2, y2).light_value(0) = rgb_value
        If range > 1 Then
            Light_Render light_index, x2, y2, rgb_value, range - 1
        End If
    End If
    
    x2 = map_x + 1
    y2 = map_y - 1
    If Map_In_Bounds(x2, y2) Then
        map_current.map_grid(x2, y2).light_value(0) = rgb_value
        If range > 1 Then
            Light_Render light_index, x2, y2, rgb_value, range - 1
        End If
    End If
    
    x2 = map_x
    y2 = map_y - 1
    If Map_In_Bounds(x2, y2) Then
        map_current.map_grid(x2, y2).light_value(0) = rgb_value
        map_current.map_grid(x2, y2).light_value(2) = rgb_value
        If range > 1 Then
            Light_Render light_index, x2, y2, rgb_value, range - 1
        End If
    End If
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
    
    For index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(index) Then
            Light_Destroy index
        End If
    Next index
    
    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Derenders a light on a map
'***************************************'
    Dim loopc As Long
    Dim min_x As Long
    Dim min_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim x As Long
    Dim y As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If Map_In_Bounds(min_x, min_y) Then
        map_current.map_grid(min_x, min_y).light_value(2) = map_current.map_grid(min_x, min_y).light_base_value(2)
    End If
    'NW
    If Map_In_Bounds(max_x, min_y) Then
        map_current.map_grid(max_x, min_y).light_value(0) = map_current.map_grid(max_x, min_y).light_base_value(0)
    End If
    'SW
    If Map_In_Bounds(max_x, max_y) Then
        map_current.map_grid(max_x, max_y).light_value(1) = map_current.map_grid(max_x, max_y).light_base_value(1)
    End If
    'SE
    If Map_In_Bounds(min_x, max_y) Then
        map_current.map_grid(min_x, max_y).light_value(3) = map_current.map_grid(min_x, max_y).light_base_value(3)
    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1
        If Map_In_Bounds(x, min_y) Then
            map_current.map_grid(x, min_y).light_value(0) = map_current.map_grid(x, min_y).light_base_value(0)
            map_current.map_grid(x, min_y).light_value(2) = map_current.map_grid(x, min_y).light_base_value(2)
        End If
    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1
        If Map_In_Bounds(x, max_y) Then
            map_current.map_grid(x, max_y).light_value(1) = map_current.map_grid(x, max_y).light_base_value(1)
            map_current.map_grid(x, max_y).light_value(3) = map_current.map_grid(x, max_y).light_base_value(3)
        End If
    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, y) Then
            map_current.map_grid(min_x, y).light_value(2) = map_current.map_grid(min_x, y).light_base_value(2)
            map_current.map_grid(min_x, y).light_value(3) = map_current.map_grid(min_x, y).light_base_value(3)
        End If
    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, y) Then
            map_current.map_grid(max_x, y).light_value(0) = map_current.map_grid(max_x, y).light_base_value(0)
            map_current.map_grid(max_x, y).light_value(1) = map_current.map_grid(max_x, y).light_base_value(1)
        End If
    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1
            If Map_In_Bounds(x, y) Then
                For loopc = 0 To 3
                    map_current.map_grid(x, y).light_value(loopc) = map_current.map_grid(x, y).light_base_value(loopc)
                Next loopc
            End If
        Next y
    Next x

End Sub


Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Public Function Particle_Group_Create(ByVal map_x As Long, ByVal map_y As Long, ByRef grh_index_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Returns the particle_group_index if successful, else 0
'**************************************************************
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), alpha_blend, alive_counter, frame_speed, id
    End If
End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
    
    For index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(index) Then
            Particle_Group_Destroy index
        End If
    Next index
    
    Particle_Group_Remove_All = True
End Function

Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Particle_Group
    
    map_current.map_grid(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub

Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Long, ByVal map_y As Long, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Makes a new particle effect
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Map pos
    particle_group_list(particle_group_index).map_x = map_x
    particle_group_list(particle_group_index).map_y = map_y
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on map
    map_current.map_grid(map_x, map_y).particle_group_index = particle_group_index
End Sub

Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Long, ByVal screen_y As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    'Set color to white
    temp_rgb(0) = &HFFFFFF
    temp_rgb(1) = &HFFFFFF
    temp_rgb(2) = &HFFFFFF
    temp_rgb(3) = &HFFFFFF
    
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
    
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).stream_type, _
                            particle_group_list(particle_group_index).grh_index_list(Int(General_Random_Number(1, particle_group_list(particle_group_index).grh_index_count))), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move
        Next loopc
        
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    
    Else
        'If it's dead destroy it
        Particle_Group_Destroy particle_group_index
    
    End If
End Sub

Public Function Particle_Group_Map_Pos_Set(ByVal particle_group_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        'Make sure it's a legal move
        If Map_In_Bounds(map_x, map_y) Then
            'Move it
            particle_group_list(particle_group_index).map_x = map_x
            particle_group_list(particle_group_index).map_y = map_y
    
            Particle_Group_Map_Pos_Set = True
        End If
    End If
End Function

Public Function Particle_Group_Move(ByVal particle_group_index As Long, ByVal heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim nx As Long
    Dim ny As Long
    
    'Check for valid heading
    If heading < 1 Or heading > 8 Then
        Particle_Group_Move = False
        Exit Function
    End If
    
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
    
        map_x = particle_group_list(particle_group_index).map_x
        map_y = particle_group_list(particle_group_index).map_y
        
        nx = map_x
        ny = map_y
        
        Convert_Heading_to_Direction heading, nx, ny
        
        'Make sure it's a legal move
        If Map_In_Bounds(nx, ny) Then
            'Move it
            particle_group_list(particle_group_index).map_x = nx
            particle_group_list(particle_group_index).map_y = ny
            
            Particle_Group_Move = True
        End If
    End If
End Function

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Long, ByVal screen_y As Long, _
                            ByVal particle_type As Long, ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/24/2003
'
'**************************************************************
    If no_move = False Then
        Select Case particle_type
            'Fountain
            Case 1
                If temp_particle.alive_counter = 0 Then
                    'Start new particle
                    Grh_Initialize temp_particle.grh, grh_index, alpha_blend
                    temp_particle.x = 0
                    temp_particle.y = 0
                    temp_particle.vector_x = General_Random_Number(-20, 20)
                    temp_particle.vector_y = General_Random_Number(-10, -50)
                    temp_particle.angle = 0
                    temp_particle.alive_counter = General_Random_Number(10, 50)
                    temp_particle.friction = 8
                Else
                    'Continue old particle
                    'Do gravity
                    temp_particle.vector_y = temp_particle.vector_y + 2
                    If temp_particle.y > 0 Then
                        'bounce
                         temp_particle.vector_y = -5
                    End If
                    'Do rotation
                    temp_particle.grh.angle = temp_particle.grh.angle + 0.1
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                End If
        
            'Star burst
            Case 2
                If temp_particle.alive_counter = 0 Then
                    'Start new particle
                    Grh_Initialize temp_particle.grh, grh_index, alpha_blend
                    temp_particle.x = 0
                    temp_particle.y = 0
                    temp_particle.vector_x = General_Random_Number(-10, 10)
                    temp_particle.vector_y = General_Random_Number(-10, 10)
                    temp_particle.angle = 0
                    temp_particle.alive_counter = General_Random_Number(10, 50)
                    temp_particle.friction = 8
                Else
                    'Continue old particle
                    'Do rotation
                    temp_particle.grh.angle = temp_particle.grh.angle + 0.1
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                End If
                
            '*********************************************************
            '* Created by: Fredrik Alexandersson  Date: 24 april 2003*
            '* Name: Insect                                          *
            '*********************************************************
            Case 3
                If temp_particle.alive_counter = 0 Then
                    Grh_Initialize temp_particle.grh, grh_index, alpha_blend
                    temp_particle.x = -5
                    temp_particle.y = -5
                    temp_particle.vector_x = 0
                    temp_particle.vector_y = 0
                    temp_particle.angle = 0
                    temp_particle.alive_counter = General_Random_Number(10, 100)
                    temp_particle.friction = 8
                Else
                    temp_particle.grh.angle = temp_particle.grh.angle + 0.1
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                    temp_particle.vector_x = General_Random_Number(-10, 10)
                    temp_particle.vector_y = General_Random_Number(-10, 10)
                End If
                
            '*********************************************************
            '* Created by: Fredrik Alexandersson  Date: 24 april 2003*
            '* Name: Water Fall                                      *
            '*********************************************************
            Case 4
                If temp_particle.alive_counter = 0 Then
                    Grh_Initialize temp_particle.grh, grh_index, alpha_blend
                    temp_particle.x = -5
                    temp_particle.y = -5
                    temp_particle.vector_x = 0
                    temp_particle.vector_y = General_Random_Number(5, 20)
                    temp_particle.angle = 0
                    temp_particle.alive_counter = General_Random_Number(10, 100)
                    temp_particle.friction = 8
                Else
                    temp_particle.grh.angle = temp_particle.grh.angle + 0.1
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                    temp_particle.vector_x = General_Random_Number(-10, 10)
                End If
                
             '*********************************************************
             '* Created by: Fredrik Alexandersson  Date: 24 april 2003*
             '* Name: Smoke                                           *
             '*********************************************************
             Case 5
                If temp_particle.alive_counter = 0 Then
                    Grh_Initialize temp_particle.grh, grh_index, alpha_blend
                    temp_particle.x = -5
                    temp_particle.y = -5
                    temp_particle.vector_x = 0
                    temp_particle.vector_y = General_Random_Number(-5, -20)
                    temp_particle.angle = 0
                    temp_particle.alive_counter = General_Random_Number(10, 100)
                    temp_particle.friction = 8
                Else
                    temp_particle.grh.angle = temp_particle.grh.angle + 0.1
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                    temp_particle.vector_x = General_Random_Number(-10, 10)
                End If
                
            '*********************************************************
            '* Created by: Fredrik Alexandersson  Date: 24 april 2003*
            '* Name: Fire?!?!                                        *
            '*********************************************************
            Case 6
                If temp_particle.alive_counter = 0 Then
                    Grh_Initialize temp_particle.grh, grh_index, alpha_blend
                    temp_particle.x = General_Random_Number(-10, 10)
                    temp_particle.y = -1
                    temp_particle.vector_x = 0
                    temp_particle.vector_y = General_Random_Number(-10, -11)
                    temp_particle.angle = 0
                    temp_particle.alive_counter = General_Random_Number(10, 100)
                    temp_particle.friction = 9
                Else
                    temp_particle.grh.angle = temp_particle.grh.angle + 0.1
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                    temp_particle.vector_y = General_Random_Number(-5, -12)
                End If
                
        End Select
        'Add in vector
        temp_particle.x = temp_particle.x + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.y = temp_particle.y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
    'Draw it
    If temp_particle.grh.grh_index Then
        Grh_Render temp_particle.grh, _
                temp_particle.x + screen_x, _
                temp_particle.y + screen_y, _
                rgb_list(), _
                False
    End If
End Sub

Public Function GUI_Box_Filled_Render(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, _
                            ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional alpha_blend As Boolean) As Boolean
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/15/2003
'Draws a filled box
'**************************************************************
    If Not Color2 <> 0 Then
        Color2 = Color1
    End If
    If Not Color3 <> 0 Then
        Color3 = Color1
    End If
    If Not Color4 <> 0 Then
        Color4 = Color1
    End If
    
    Dim Vertex(3) As TLVERTEX
    Dim box_rect As RECT
    
    With box_rect
        .bottom = y + height - 1
        .left = x
        .Right = x + width - 1
        .top = y
    End With
    
    Vertex(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 1, Color1, 0, 0, 0)
    Vertex(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color2, 0, 0, 0)
    Vertex(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 1, Color3, 0, 0, 0)
    Vertex(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color4, 0, 0, 0)
    
    'Enable alpha-blending
    ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    ddevice.SetTexture 0, Nothing
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
    
    'Turn off alphablending after we're done
    ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    
    GUI_Box_Filled_Render = True
End Function

Public Function GUI_Box_Outline_Render(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, _
                            ByVal border_width As Long, ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional alpha_blend As Boolean) As Boolean
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/15/2003
'Draws a box outline
'**************************************************************
    If Not Color2 <> 0 Then
        Color2 = Color1
    End If
    If Not Color3 <> 0 Then
        Color3 = Color1
    End If
    If Not Color4 <> 0 Then
        Color4 = Color1
    End If
    
    Dim VertexU(3) As TLVERTEX
    Dim VertexL(3) As TLVERTEX
    Dim VertexR(3) As TLVERTEX
    Dim VertexB(3) As TLVERTEX
    Dim box_rect As RECT
    
    With box_rect
        .bottom = y + height - 1
        .left = x
        .Right = x + width - 1
        .top = y
    End With
    
    'Enable alpha-blending
    ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    ddevice.SetTexture 0, Nothing
    'Upper Line
    VertexU(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 1, Color1, 0, 0, 0)
    VertexU(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color1, 0, 0, 0)
    VertexU(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top + border_width, 0, 1, Color1, 0, 0, 0)
    VertexU(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top + border_width, 0, 1, Color1, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexU(0), Len(VertexU(0))
    'Left Line
    VertexL(0) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.top, 0, 1, Color2, 0, 0, 0)
    VertexL(1) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.bottom, 0, 1, Color2, 0, 0, 0)
    VertexL(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 2, Color1, 0, 0, 0)
    VertexL(3) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 2, Color1, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexL(0), Len(VertexL(0))
    'Right Border
    VertexR(0) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color3, 0, 0, 0)
    VertexR(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color3, 0, 0, 0)
    VertexR(2) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.top, 0, 3, Color1, 0, 0, 0)
    VertexR(3) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.bottom, 0, 3, Color1, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexR(0), Len(VertexR(0))
    'Bottom Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom - border_width, 0, 1, Color4, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom - border_width, 0, 1, Color4, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 1, Color4, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color4, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    
    'Turn off alphablending after we're done
    ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    
    GUI_Box_Outline_Render = True
End Function

Public Function GUI_Text_Render(ByVal text As String, ByVal font_index As Long, ByVal x As Long, _
                                ByVal y As Long, ByVal width As Long, _
                                ByVal height As Long, ByVal color As Long, ByVal alignment As FontAlignment, Optional ByVal word_wrap As Boolean) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'
'**************************************************************
    If Font_Check(font_index) = False Then
        Exit Function
    End If
    
    If word_wrap Then
        alignment = alignment Or DT_WORDBREAK Or DT_EDITCONTROL
    End If
    
    Device_Text_Render font_list(font_index), text, x, y, width, height, color, alignment
    
    GUI_Text_Render = True
End Function

Public Function GUI_Grh_Render(ByVal grh_index As Long, x As Long, y As Long, Optional ByVal angle As Single, Optional ByVal alpha_blend As Boolean) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Dim temp_grh As grh
    Dim rpg_list(3) As Long

    If Grh_Check(grh_index) = False Then
        Exit Function
    End If

    rpg_list(0) = &HFFFFFFFF
    rpg_list(1) = &HFFFFFFFF
    rpg_list(2) = &HFFFFFFFF
    rpg_list(3) = &HFFFFFFFF

    Grh_Initialize temp_grh, grh_index, alpha_blend, angle
    
    Grh_Render temp_grh, x, y, rpg_list
    
    GUI_Grh_Render = True
End Function
