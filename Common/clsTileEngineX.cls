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
'11/11/2004 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: Grh_Render_To_hDC_Advance method
'   -Add: Engine_Render_Mini_Map_To_hDC method
'   -Add: Input_Mouse_Button_Left_Click and Input_Mouse_Button_Right_Click
'   -Add: Input_Mouse_Button_Left_Double_Click_Get, Input_Mouse_Button_Right_Double_Click_Get and
'           Input_Mouse_Double_Click_Time_Set, allowing the engine to check for double clicks
'   -Change: Light_Render method was rewritten and now is much faster
'   -Change: Changed the way mouse button events were handled to avoid errors
'   -Change: Map_Base_Light_Fill was edited to store the base light color to the map structure (also edited)
'   -Change: Several var's type was changed to reduce memory used, and make maps smaller
'   -Change: Map format has been totally changed for a great reduction in map sizes
'   -Fix: Grh_Render_To_hDC had a memory leak issue
'
'4/22/2004 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: Overhead chats to chars
'   -Add: Engine_Resolution_Change method
'   -Change: Device_Reset now also resets fonts and enables alpha-blending
'   -Change: Font_Make now saves the StdFont so it can be recreated after changing resolution
'   -Change: Engine_Render_Start now calls the OnDeviceLost method for each font
'   -Change: Engine_Initialize now takes 2 new parameters to allow
'it to unlock grhs as they are loaded and destroy them afterwards to grant security
'   -Change: Texture_Load_All unclocks files if needed and deletes them if asked to do so
'   -Fix: Fonts are now destroyed to avoid memory leak
'   -Fix: Font alignments were completely edited to make them work properly
'
'2/26/2004 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: Device_Box_Textured_Render improved for extra speed is faster
'   -Change: Alphablending is enabled only once, instead of once every time a texture is rendered
'   -Change: Private Sub Device_Box_Textured_Render is now called Device_Box_Textured_Render_Advance, allows resizing but is slower
'   -Change: Lights can only be created if the tile had no light previously set
'   -Change: Engine_Render_Start now is able to avoid being called without having ended the scene
'   -Fix: Totally improved the way Particle_Streams are rendered
'   -Fix: Input_Mouse_Poll is set to 1000 data items, since the buffer size was set to 1000 (improved FPS)
'   Sub Release Contributors:
'       2/26/2004 - Fredrik Alexandersson (fredrik@oraklet.zzn.com)
'           - Add: Path_Resources_Get method
'
'2/11/2004 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: Map_Load_Map_By_Name method
'   -Add: GUI_Box_Filled_Render_Advance method
'   -Change: Map format changed to make .map files smaller
'   -Change: Char_Render to make labels be allways centered, no matter the lenght of the text (offset_x will have to be big enough)
'
'12/08/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: Several methods relating subtiles
'   -Add: Map_Grh_Info_Get and Map_Decoration_Grh_Info_Get methods
'   -Add: Map_Resize method
'   -Add: Grhs can be centered in layers 2 and 3
'   -Add: All Trigger related methods
'   -Add: Particle_Type_Get method
'   -Add: Map_Bounds_Get_From_File method
'   -Add: GUI_Grh_Render_Advance and Grh_Render_Advance methods
'   -Fix: The mouse pos is kept updated
'   -Fix: Geometry_Create_Box method was slightly modified to improve speed
'   -Change: Included the Particle_Groups_Load_All method (which originally belonged to the Map Editor)
'   -Change: There are only 3 formal layers (floor, walls and roof), and a decorative layer
'            This decorative layer can be rendered either above or beneath layer 2
'   -Change: Map format includes triggers and centered grhs
'   -Change: Particle_Render no longer checks for particle type, since any particle
'            stream can be done with OneZero´s modifications
'   -Change: Removed Layer 5 since it wasn´t used at all
'   -Change: Several more small fixes to origninal code
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
'           - Fix: Geometry_Create_Box - fixed rotation
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
    fa_center = DT_CENTER Or DT_VCENTER
    fa_top = DT_TOP Or DT_CENTER
    fa_left = DT_LEFT Or DT_VCENTER
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM Or DT_CENTER
    fa_right = DT_RIGHT Or DT_VCENTER
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

Private Const COLOR_KEY As Long = &HFF000000
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'Bitmap header
Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Bitmap info header
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

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
    X As Single
    Y As Single
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
    mini_map_color As Long
End Type

'Points to a Grh_Data and keeps animation info
'Edited by Juan Martín Sotuyo Dodero
Private Type grh
    grh_index As Integer
    alpha_blend As Boolean
    angle As Single
    frame_speed As Single
    frame_counter As Single
    started As Boolean
    noloop As Boolean
    h_centered As Boolean
    v_centered As Boolean
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
'Edited by Juan Martín Sotuyo Dodero
Private Type char
    active As Boolean
    Heading As Long
    id As Long
    map_x As Long
    map_y As Long

    chr_data As Char_Data
    chr_data_body_index As Long
    
    label As String
    label_font_index As Long
    label_offset_x As Long
    label_offset_y As Long
    
    Chat As String
    chat_death_time As Long
    chat_color As Long
    chat_font_index As Long
    
    scroll_on As Boolean
    scroll_offset_counter_x As Single
    scroll_offset_counter_y As Single
    scroll_direction_x As Long
    scroll_direction_y As Long
End Type

'Added by Juan Martín Sotuyo Dodero
Private Type decoration
    grh As grh
    Render_On_Top As Boolean
    subtile_pos As Byte
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
'Edited by Juan Martín Sotuyo Dodero
Private Type Map_Tile
    grh(1 To 3) As grh
    decoration(1 To 5) As decoration
    decoration_count As Byte
    blocked As Boolean
    particle_group_index As Long
    char_index As Long
    light_base_value(0 To 3) As Long
    light_value(0 To 3) As Long
    
    exit_index As Long
    npc_index As Long
    Item_Index As Long
    
    trigger As Byte
End Type

'Map structure
Private Type Map
    map_grid() As Map_Tile
    map_x_max As Long
    map_x_min As Long
    map_y_max As Long
    map_y_min As Long
    map_description As String
    'Added by Juan Martín Sotuyo Dodero
    base_light_color As Long
End Type

'The light structure
Private Type Light
    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Long 'Coordinates
    map_y As Long
    color As Long 'Start colour
    range As Byte
End Type

'Particle types
'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
'Particle types
'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    grh As grh
    alive_counter As Long
    x1 As Long
    x2 As Long
    y1 As Long
    y2 As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type Particle_Group
    active As Boolean
    id As Long
    map_x As Long
    map_y As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    x1 As Long
    x2 As Long
    y1 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martín Sotuyo Dodero
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Long
    grh_resizey As Long
End Type

'Added by Juan Martín Sotuyo Dodero
'This is only used by Map_Render to store decorations rendered
'on top, and avoid having to check for decorations twice
Private Type decoration_list
    map_x As Long
    map_y As Long
    decoration_index As Byte
    screen_x As Long
    screen_y As Long
End Type

'***************************
'Variables
'***************************
'Major DX Objects
Dim DX As DirectX8
Dim d3d As Direct3D8
Dim ddevice As Direct3DDevice8
Dim d3dx As D3DX8
Dim d3dpp As D3DPRESENT_PARAMETERS
Dim d3dcaps As D3DCAPS8
Dim d3ddm As D3DDISPLAYMODE
Dim di As DirectInput8
Dim ddevice_mouse As DirectInputDevice8

'The app path
Dim Path_Resources As String

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

'show engine stats
Dim engine_show_stats As Boolean

'show blocked tiles on map
Dim engine_show_blocked_tiles As Boolean
Dim engine_show_special_tiles As Boolean
Dim engine_show_triggers As Boolean

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
'Added by Juan Martín Sotuyo Dodero
Dim input_mouse_subtile As Long
Dim input_mouse_left_click_last As Long
Dim input_mouse_left_double_click As Boolean
Dim input_mouse_right_click_last As Long
Dim input_mouse_right_double_click As Boolean
Dim input_mouse_double_click_time As Long
Dim input_mouse_left_click As Boolean
Dim input_mouse_right_click As Boolean


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
Dim font_primitive_list() As StdFont  'This contains all basic data on fonts
Dim font_list() As D3DXFont
Dim font_count As Long
Dim font_last As Long

Dim exit_list() As Map_Exit
Dim npc_list() As Map_NPC
Dim item_list() As Map_Item

'Loaded Particle groups list
Dim StreamData() As Particle_Group
Dim TotalStreams As Long

'***************************
'External Functions
'***************************
'KeyInput
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

'MouseInput
Private Type PointAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

'For getting the display size in windowed mode
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Old fashion BitBlt function
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Added by Juan Martín Sotuyo Dodero
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

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
    Engine_DeInitialize
End Sub

Private Function Convert_Tile_To_View_Y(ByVal Y As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_Y = ((Y - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_Y = view_screen_top + ((Y - 1) * base_tile_size)
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

Private Sub Convert_Heading_to_Direction(ByVal Heading As Long, ByRef direction_x As Long, ByRef direction_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim addY As Long
    Dim addX As Long
    
    'Figure out which way to move
    Select Case Heading
    
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

Private Function Convert_Tile_To_View_X(ByVal X As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_X = ((X - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_X = view_screen_left + ((X - 1) * base_tile_size)
    End If
End Function

Private Sub Engine_Stats_Render()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim X As Long
    Dim Y As Long
    
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
        'Update mouse pos in map
        If Input_Mouse_In_View Then
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

Public Sub Engine_DeInitialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/28/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
On Error Resume Next
    
    Dim i As Long
    
    'Set no texture to standard stage to avoid memory leak
    ddevice.SetTexture 0, Nothing
    
    'Destroy all textures
    For i = 1 To UBound(texture_list)
        Set texture_list(i).d3dtexture = Nothing
    Next i
    
    'Destroy all fonts
    For i = 1 To UBound(font_list)
        Set font_list(i) = Nothing
    Next i
    
    For i = 1 To UBound(font_primitive_list)
        Set font_primitive_list(i) = Nothing
    Next i
    
    'ddevice_mouse.Unacquire
    Set ddevice_mouse = Nothing
    Set di = Nothing
    
    Set d3dx = Nothing
    Set ddevice = Nothing
    Set d3d = Nothing
    Set DX = Nothing
End Sub

Public Function Engine_Initialize(ByVal f_hwnd As Long, ByVal s_hwnd As Long, ByVal windowed As Boolean, _
                                    ByVal r_path As String, Optional ByVal s_width As Long, Optional ByVal s_height As Long, _
                                    Optional ByVal v_left As Long = 0, Optional ByVal v_top As Long = 0, Optional ByVal v_width_in_tiles As Long = 0, _
                                    Optional ByVal v_height_in_tiles As Long = 0, Optional ByVal tile_size As Long = 32, Optional grh_files_need_unlocking As Boolean = False, _
                                    Optional delete_grh_files_after_load As Boolean = False) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 7/12/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
On Error GoTo ErrorHandler:
    Engine_Initialize = True
    
    '****************************************************
    'Fill in global variables
    '****************************************************
    'Main display
    screen_hwnd = s_hwnd
    form_hwnd = f_hwnd
    
    'App path
    Path_Resources = r_path
    
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
    End If
    
    '****************************************************
    'Get external data
    '****************************************************
    'Load Grh List
    Grh_Load_All
    
    'Load body data for characters
    Char_Load_Char_Data
    
    'Load Particle Streams data
    Particle_Group_Load_All
    
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
    Set DX = New DirectX8
    'Create the Direct3D object
    Set d3d = DX.Direct3DCreate
    'Create helper class
    Set d3dx = New D3DX8
    'Input object
    Set di = DX.DirectInputCreate
    
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
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
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
        .hDeviceWindow = screen_hwnd
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
        'Enable alpha blending
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
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
    Call Texture_Load_All(grh_files_need_unlocking, delete_grh_files_after_load)
    
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

Public Function Engine_View_Move(ByVal Heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim X As Long
    Dim Y As Long
    Dim nx As Long
    Dim ny As Long
    
    'Don't move if we are already moving....
    If scroll_on Then
        Engine_View_Move = False
        Exit Function
    End If
    
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then
        Engine_View_Move = False
        Exit Function
    End If
    
    X = view_pos_x
    Y = view_pos_y
    nx = X
    ny = Y
    Convert_Heading_to_Direction Heading, nx, ny
    
    'See if out new position is legal
    If Map_In_Bounds(nx, ny) Then
        'start the scrolling process
        view_pos_x = nx
        view_pos_y = ny
        
        scroll_offset_counter_x = (base_tile_size * (X - nx))
        scroll_offset_counter_y = (base_tile_size * (Y - ny))
        scroll_direction_x = nx - X
        scroll_direction_y = ny - Y
        scroll_on = True
        
        'See if mouse is in view
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
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
'On Error GoTo ErrorHandler:
    If engine_render_started Then Exit Function
    
    Engine_Render_Start = True
    
    '****************************************************
    'Handle Alt-Tab (not quite working yet)
    '****************************************************
    'Call TestCooperativeLevel to see what state the device is in.
    Dim hr As Long
    hr = ddevice.TestCooperativeLevel
    If hr = D3DERR_DEVICELOST Then
        Dim LoopC As Long
        For LoopC = 1 To font_count
            font_list(LoopC).OnLostDevice
        Next LoopC
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
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
'On Error GoTo ErrorHandler:
    Engine_Render_End = True

    If Not engine_render_started Then Exit Function

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

Private Sub Device_Box_Textured_Render_Advance(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal src_width As Long, _
                                            ByVal src_height As Long, ByRef rgb_list() As Long, ByVal src_x As Long, _
                                            ByVal src_y As Long, ByVal dest_width As Long, Optional ByVal dest_height As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the texture allowing resizing
'Modified by Juan Martín Sotuyo Dodero
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
    ddevice.SetTexture 0, texture_list(texture_index).d3dtexture
    
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
End Sub

Private Sub Device_Box_Textured_Render(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal src_width As Long, _
                                            ByVal src_height As Long, ByRef rgb_list() As Long, ByVal src_x As Long, _
                                            ByVal src_y As Long, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the texture
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
        .bottom = dest_y + src_height
        .left = dest_x
        .Right = dest_x + src_width
        .top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, rgb_list(), texture_list(texture_index).dimension, angle
    
    'Set Texture
    ddevice.SetTexture 0, texture_list(texture_index).d3dtexture
    
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
'Last Modify Date: 3/30/2004
'Resets the device
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    'Destroy all fonts
    Dim LoopC As Long
    For LoopC = 1 To font_last
        Set font_list(LoopC) = Nothing
    Next LoopC
    
    font_last = 0
    
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
                                    
        'Enable alpha blending
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        
        'Set the render state that uses the alpha component as the source for blending.
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            
        'Set the render state that uses the inverse alpha component as the destination blend.
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    End With
    
    'Recreate all fonts
    For LoopC = 1 To UBound(font_primitive_list())
        Font_Create font_primitive_list(LoopC).name, font_primitive_list(LoopC).size, font_primitive_list(LoopC).bold, font_primitive_list(LoopC).italic
    Next LoopC
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

Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
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
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.Right - dest.left - 1) / 2
        y_center = dest.top + (dest.bottom - dest.top - 1) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
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

Private Sub Texture_Load_All(ByVal grh_files_need_unlocking As Boolean, ByVal delete_grh_files_after_load As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 7/12/2004
'Loads all the graphics into memory
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Dim i As Long
    Dim inipath As String
    Dim filepath As String
    Dim LoopC As Long
    
    inipath = Path_Resources & PATH_GRAPHICS & "\"
    
    'Get Number of Graphics
    texture_count = General_Var_Get(inipath & "grh.ini", "INIT", "texture_count")
    
    ReDim Preserve texture_list(1 To texture_count)
    
    For i = 1 To texture_count
        filepath = Path_Resources & PATH_GRAPHICS & "\grh" & i & ".bmp"
        If General_File_Exists(filepath, vbNormal) Then
            If grh_files_need_unlocking Then
                LoopC = 0
                Do Until LCase(Trim(GRH_Handles(LoopC).file_name)) = "grh" & i & ".bmp"
                    LoopC = LoopC + 1
                Loop
                Close GRH_Handles(LoopC).handle
            End If
            Texture_Load_From_File i, filepath
            If delete_grh_files_after_load Then
                Kill filepath
            Else
                Debug.Print filepath
                Open filepath For Binary As GRH_Handles(LoopC).handle
            End If
        End If
    Next i
End Sub

Private Sub Grh_Initialize(ByRef grh As grh, ByVal grh_index As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sets up a grh. MUST be done before rendering
'Modified by Juan Martín Sotuyo Dodero
'Added centering
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
    
    'Set centered state
    grh.h_centered = h_centered
    grh.v_centered = v_centered
End Sub

Private Sub Grh_Load_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/13/2004
'Loads Grh.dat
'Modified by David Justus
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
On Error GoTo ErrorHandler
    Dim grh As Long
    Dim frame As Long
    Dim TempInt As Long
    
    Dim inipath As String
    inipath = Path_Resources & PATH_GRAPHICS & "\"
    
    'Get number of grhs
    grh_count = Val(General_Var_Get(inipath & "grh.ini", "INIT", "grh_count"))
    
    'Resize arrays
    ReDim grh_list(1 To grh_count) As Grh_Data
    
    'Open files
    Open inipath & "grh.dat" For Binary As #1
    Seek #1, 1
    
    'Get Header
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    
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
    
    Dim count As Long
    Open inipath & "minimap.dat" For Binary As #1
        Seek #1, 1
        For count = 1 To grh_count
            If Grh_Check(count) Then
                Get #1, , grh_list(count).mini_map_color
            End If
        Next count
    Close #1
Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error while loading the grh.dat! Stopped at GRH number: " & grh
End Sub

Public Sub Grh_Add_GrhList_To_ListBox(ListBoxName As ListBox)
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
            ListBoxName.AddItem grh
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

    file_path = Path_Resources & PATH_GRAPHICS & "\grh" & grh_list(grh_index).texture_index & ".bmp"
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

Private Sub Grh_Render(ByRef grh As grh, ByVal screen_x As Long, ByVal screen_y As Long, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
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
    
    'Center Grh over X, Y pos
    If grh.h_centered Or h_centered Then
        tile_width = grh_list(grh_index).src_width / base_tile_size
        If tile_width <> 1 Then
            screen_x = screen_x - Int(tile_width * base_tile_size / 2) + base_tile_size / 2
        End If
    End If
    If grh.v_centered Or v_centered Then
        tile_height = grh_list(grh_index).src_height / base_tile_size
        If tile_height <> 1 Then
            screen_y = screen_y - Int(tile_height * base_tile_size / 2) + base_tile_size / 2
        End If
    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_list(grh_index).texture_index, _
        screen_x, screen_y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        rgb_list, _
        grh_list(grh_index).src_x, grh_list(grh_index).src_y, _
        grh.alpha_blend, _
        grh.angle
End Sub

Public Sub Grh_Render_To_Hdc(ByVal grh_index As Long, destHDC As Long, ByVal screen_x As Long, ByVal screen_y As Long, Optional transparent As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/30/2004
'This method is SLOW... Don't use in a loop if you care about
'speed!
'Modified by Juan Martín Sotuyo Dodero
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
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long

    'If it's animated switch grh_index to first frame
    If grh_list(grh_index).frame_count <> 1 Then
        grh_index = grh_list(grh_index).frame_list(1)
    End If

    file_path = Path_Resources & PATH_GRAPHICS & "\grh" & grh_list(grh_index).texture_index & ".bmp"
    src_x = grh_list(grh_index).src_x
    src_y = grh_list(grh_index).src_y
    src_width = grh_list(grh_index).src_width
    src_height = grh_list(grh_index).src_height
    
    hdcsrc = CreateCompatibleDC(destHDC)
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
    
    If transparent = False Then
        BitBlt destHDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
    Else
        MaskDC = CreateCompatibleDC(destHDC)
        
        PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
        
        Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
        
        'Render tranparently
        BitBlt destHDC, 0, 0, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
        BitBlt destHDC, 0, 0, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint
        
        Call DeleteObject(SelectObject(MaskDC, PrevObj2))
        
        DeleteDC MaskDC
    End If
    
    Call DeleteObject(SelectObject(hdcsrc, PrevObj))
    
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
'Last Modify Date: 8/05/2004
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
On Error Resume Next
    'Get the mouse event buffer
    Dim DevData(1 To 1000) As DIDEVICEOBJECTDATA
    Dim nEvents As Long
    nEvents = ddevice_mouse.GetDeviceData(DevData, DIGDD_DEFAULT)
    
    'Check buffer for clicks
    Dim i As Long
    For i = 1 To nEvents
        Select Case DevData(i).lOfs
            Case DIMOFS_BUTTON0
                'Check if mouse button was pressed or released
                input_mouse_button_left = (DevData(i).lData And &H80)
                'If released it was a click
                If Not input_mouse_button_left Then
                    input_mouse_left_click = True
                Else
                    input_mouse_left_click = False
                End If
                'Check if it was a double click
                If input_mouse_left_click_last - GetTickCount <= input_mouse_double_click_time Then
                    input_mouse_left_double_click = True
                Else
                    input_mouse_left_double_click = False
                    input_mouse_left_click_last = GetTickCount
                End If
                
            Case DIMOFS_BUTTON1
                'Check if mouse button was pressed or released
                input_mouse_button_right = (DevData(i).lData And &H80)
                'If released it was a click
                If Not input_mouse_button_right Then
                    input_mouse_right_click = True
                Else
                    input_mouse_right_click = False
                End If
                'Check if it was a double click
                If input_mouse_right_click_last - GetTickCount <= input_mouse_double_click_time Then
                    input_mouse_right_double_click = True
                Else
                    input_mouse_right_double_click = False
                    input_mouse_right_click_last = GetTickCount
                End If
                
        End Select
    Next i
    
    'Use a API to get the mouse cordinates.
    Dim temp_point As PointAPI
    Dim temp_x As Long
    Dim temp_y As Long
    GetCursorPos temp_point
    
    'See if we moved
    If input_mouse_screen_x <> temp_point.X Or input_mouse_screen_y <> temp_point.Y Then
        
        'We moved
        'Store new screen position
        input_mouse_screen_x = temp_point.X
        input_mouse_screen_y = temp_point.Y
        
        'See if we are in the view area
        Convert_Screen_To_View input_mouse_screen_x, input_mouse_screen_y, input_mouse_view_x, input_mouse_view_y
                  
        'Figure out where on the map was clicked
        Convert_View_To_Map input_mouse_view_x, input_mouse_view_y, input_mouse_map_x, input_mouse_map_y
        
        'Figure out on which subtile
        input_mouse_subtile = Convert_View_To_Subtile(input_mouse_view_x, input_mouse_view_y)
        
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
    NPC_Ini_Char_Data_Index_Get = CLng(General_Var_Get(Path_Resources & PATH_SCRIPTS & "\npc.ini", "NPC" & s_npc_data_index, "npc_char_data_index"))
End Function

Private Function Item_Ini_Grh_Index_Get(ByVal s_item_data_index As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    Item_Ini_Grh_Index_Get = CLng(General_Var_Get(Path_Resources & PATH_SCRIPTS & "\item.ini", "ITEM" & s_item_data_index, "item_grh_index"))
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
        If map_current.map_grid(s_map_x, s_map_y).Item_Index = 0 Then
            ReDim Preserve item_list(0 To UBound(item_list) + 1)
            item_list(UBound(item_list)).item_data_index = s_item_data_index
            item_list(UBound(item_list)).item_amount = s_item_amount
            item_list(UBound(item_list)).c_grh_index = Item_Ini_Grh_Index_Get(s_item_data_index)
            item_list(UBound(item_list)).c_map_x = s_map_x
            item_list(UBound(item_list)).c_map_y = s_map_y
            map_current.map_grid(s_map_x, s_map_y).Item_Index = UBound(item_list)
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
        If map_current.map_grid(s_map_x, s_map_y).Item_Index <> 0 Then
            item_list(map_current.map_grid(s_map_x, s_map_y).Item_Index).item_data_index = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).Item_Index).item_amount = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).Item_Index).c_grh_index = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).Item_Index).c_map_x = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).Item_Index).c_map_y = 0
            map_current.map_grid(s_map_x, s_map_y).Item_Index = 0
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
        If map_current.map_grid(s_map_x, s_map_y).Item_Index <> 0 Then
            r_item_data_index = item_list(map_current.map_grid(s_map_x, s_map_y).Item_Index).item_data_index
            r_item_amount = item_list(map_current.map_grid(s_map_x, s_map_y).Item_Index).item_amount
            Map_Item_Get = True
        End If
    End If
End Function

Public Function Map_Save_Ini_To_File(ByVal file_path As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 9/02/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Dim LoopC As Long
    Dim counter As Long
    Dim map_x As Long
    Dim map_y As Long

   'If file already exists kill it
    If General_File_Exists(file_path, vbNormal) Then
        Kill file_path
    End If
    
    General_Var_Write file_path, "GENERAL", "map_description", map_current.map_description
    
    'NPCs
    counter = 1
    If UBound(npc_list()) <> 0 Then
        For LoopC = 1 To UBound(npc_list())
            If npc_list(LoopC).npc_data_index Then
                General_Var_Write file_path, "NPC", CStr(counter), CStr(npc_list(LoopC).c_map_x) & "-" & CStr(npc_list(LoopC).c_map_y) & "-" & CStr(npc_list(LoopC).npc_data_index)
                counter = counter + 1
            End If
        Next LoopC
    End If
    General_Var_Write file_path, "NPC", "count", CStr(counter - 1)
    
    'Exits
    counter = 1
    If UBound(exit_list()) <> 0 Then
        For LoopC = 1 To UBound(exit_list())
            If exit_list(LoopC).exit_map_name <> "" Then
                General_Var_Write file_path, "EXIT", CStr(counter), CStr(exit_list(LoopC).c_map_x) & "-" & CStr(exit_list(LoopC).c_map_y) & "-" & exit_list(LoopC).exit_map_name & "-" & CStr(exit_list(LoopC).exit_map_x) & "-" & CStr(exit_list(LoopC).exit_map_y)
                counter = counter + 1
            End If
        Next LoopC
    End If
    General_Var_Write file_path, "EXIT", "count", CStr(counter - 1)
    
    'Items
    counter = 1
    If UBound(item_list()) <> 0 Then
        For LoopC = 1 To UBound(item_list())
            If item_list(LoopC).item_data_index Then
                General_Var_Write file_path, "ITEM", CStr(counter), CStr(item_list(LoopC).c_map_x) & "-" & CStr(item_list(LoopC).c_map_y) & "-" & CStr(item_list(LoopC).item_data_index) & "-" & CStr(item_list(LoopC).item_amount)
                counter = counter + 1
            End If
        Next LoopC
    End If
    General_Var_Write file_path, "ITEM", "count", CStr(counter - 1)
    
    'Triggers
    counter = 1
    For map_y = map_current.map_y_min To map_current.map_y_max
        For map_x = map_current.map_x_min To map_current.map_x_max
            If map_current.map_grid(map_x, map_y).trigger > 0 And map_current.map_grid(map_x, map_y).trigger <> 3 Then
                General_Var_Write file_path, "TRIGGER", CStr(counter), CStr(map_x) & "-" & CStr(map_y) & "-" & CStr(map_current.map_grid(map_x, map_y).trigger)
                counter = counter + 1
            End If
        Next map_x
    Next map_y
    General_Var_Write file_path, "TRIGGER", "count", CStr(counter - 1)
    
    Map_Save_Ini_To_File = True
End Function

Public Function Map_Load_Ini_From_File(ByVal file_path As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 9/02/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Dim LoopC As Long
    Dim t_count As Long
    Dim temp_string As String
    
    map_current.map_description = General_Var_Get(file_path, "GENERAL", "map_description")
        
    'NPC
    t_count = Val(General_Var_Get(file_path, "NPC", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "NPC", CStr(LoopC))
        If temp_string <> "" Then
            Map_NPC_Add CLng(General_Field_Read(1, temp_string, Asc("-"))), CLng(General_Field_Read(2, temp_string, Asc("-"))), CLng(General_Field_Read(3, temp_string, Asc("-")))
        End If
    Next LoopC
    
    'Item
    t_count = Val(General_Var_Get(file_path, "ITEM", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "ITEM", CStr(LoopC))
        If temp_string <> "" Then
            Map_Item_Add CLng(General_Field_Read(1, temp_string, Asc("-"))), CLng(General_Field_Read(2, temp_string, Asc("-"))), CLng(General_Field_Read(3, temp_string, Asc("-"))), CLng(General_Field_Read(4, temp_string, Asc("-")))
        End If
    Next LoopC
    
    'Exit
    t_count = Val(General_Var_Get(file_path, "EXIT", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "EXIT", CStr(LoopC))
        If temp_string <> "" Then
            Map_Exit_Add CLng(General_Field_Read(1, temp_string, Asc("-"))), CLng(General_Field_Read(2, temp_string, Asc("-"))), General_Field_Read(3, temp_string, Asc("-")), CLng(General_Field_Read(4, temp_string, Asc("-"))), CLng(General_Field_Read(5, temp_string, Asc("-")))
        End If
    Next LoopC
    
    'Triggers
    t_count = Val(General_Var_Get(file_path, "TRIGGER", "count"))
    For LoopC = 1 To t_count
        temp_string = General_Var_Get(file_path, "TRIGGER", CStr(LoopC))
        If temp_string <> "" Then
            Map_Trigger_Set CLng(General_Field_Read(1, temp_string, Asc("-"))), CLng(General_Field_Read(2, temp_string, Asc("-"))), CLng(General_Field_Read(3, temp_string, Asc("-")))
        End If
    Next LoopC
    
    Map_Load_Ini_From_File = True
End Function

Public Function Map_Save_Map(ByVal map_id As Long, Optional ByVal save_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Save map data to files
'*****************************************************************
    Dim map_path As String
    
    'Get map file path
    map_path = Path_Resources & PATH_MAPS & "\map" & map_id & ".map"
    
    Map_Save_Map = Map_Save_Map_To_File(map_path, save_ini)
End Function

Public Function Map_Save_Map_To_File(ByVal file_path As String, Optional ByVal save_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/26/2004
'Modified by Juan Martín Sotuyo Dodero
'Save map data to files
'*****************************************************************
    Dim LoopC As Long
    Dim Y As Long
    Dim X As Long
    Dim map_path As String
    Dim tile_flags As Byte
    Dim layer_flags As Byte
    Dim deco_flags As Byte
    Dim light_index As Long
    Dim particle_index As Long
    
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
    Put #1, , map_current.map_x_max
    Put #1, , map_current.map_y_max
    Put #1, , map_current.base_light_color
    
    'Write .map file
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
            '************************
            'Tile Flags START
            '************************
            tile_flags = 0
            
            'Layer 2
            If map_current.map_grid(X, Y).grh(2).grh_index Then tile_flags = tile_flags Or 1
            'Layer 3
            If map_current.map_grid(X, Y).grh(3).grh_index Then tile_flags = tile_flags Or 2
            'Blocked
            If map_current.map_grid(X, Y).blocked Then tile_flags = tile_flags Or 4
            'Base lights
            For LoopC = 0 To 3
                If map_current.map_grid(X, Y).light_base_value(LoopC) <> map_current.base_light_color Then
                    tile_flags = tile_flags Or 8
                    Exit For
                End If
            Next LoopC
            'Lights
            light_index = Map_Light_Get(X, Y)
            If light_index Then tile_flags = tile_flags Or 16
            'Decorations
            If map_current.map_grid(X, Y).decoration_count Then tile_flags = tile_flags Or 32
            'Particle Streams
            particle_index = Map_Particle_Group_Get(X, Y)
            If particle_index Then tile_flags = tile_flags Or 64
            'Non-Combat Zone trigger
            If map_current.map_grid(X, Y).trigger = 3 Then tile_flags = tile_flags Or 128
            '************************
            'Tile Flags END
            '************************
            
            '.map file
            
            'Store flags
            Put #1, , tile_flags
            
            'Layer 1
            Put #1, , map_current.map_grid(X, Y).grh(1).grh_index
            
            'Layers 2 and 3
            For LoopC = 2 To 3
                If map_current.map_grid(X, Y).grh(LoopC).grh_index Then
                    '************************
                    'Layer Flags START
                    '************************
                    layer_flags = 0
                    If map_current.map_grid(X, Y).grh(LoopC).alpha_blend Then layer_flags = layer_flags Or 1
                    If map_current.map_grid(X, Y).grh(LoopC).h_centered Then layer_flags = layer_flags Or 2
                    If map_current.map_grid(X, Y).grh(LoopC).v_centered Then layer_flags = layer_flags Or 4
                    If map_current.map_grid(X, Y).grh(LoopC).angle Then layer_flags = layer_flags Or 8
                    '************************
                    'Layer Flags END
                    '************************
                    
                    'Store grh index
                    Put #1, , map_current.map_grid(X, Y).grh(LoopC).grh_index
                    
                    'Store layer flags
                    Put #1, , layer_flags
                    
                    'Store angle (only if needed)
                    If map_current.map_grid(X, Y).grh(LoopC).angle Then
                        Put #1, , General_Convert_Radians_To_Degrees(map_current.map_grid(X, Y).grh(LoopC).angle)
                    End If
                End If
            Next LoopC
            
            'Base lights
            If tile_flags And 8 Then
                For LoopC = 0 To 3
                    Put #1, , map_current.map_grid(X, Y).light_base_value(LoopC)
                Next LoopC
            End If
            
            'Lights
            If light_index Then
                Put #1, , light_list(light_index).color
                Put #1, , light_list(light_index).range
            End If
            
            'Particle groups
            If particle_index Then
                Put #1, , particle_group_list(particle_index).stream_type
            End If
            
            'Decorations
            If tile_flags And 32 Then
                Put #1, , map_current.map_grid(X, Y).decoration_count
                For LoopC = 1 To map_current.map_grid(X, Y).decoration_count
                    '************************
                    'Deco Flags START
                    '************************
                    deco_flags = 0
                    If map_current.map_grid(X, Y).decoration(LoopC).grh.alpha_blend Then deco_flags = deco_flags Or 1
                    If map_current.map_grid(X, Y).decoration(LoopC).grh.h_centered Then deco_flags = deco_flags Or 2
                    If map_current.map_grid(X, Y).decoration(LoopC).grh.v_centered Then deco_flags = deco_flags Or 4
                    If map_current.map_grid(X, Y).decoration(LoopC).grh.angle Then deco_flags = deco_flags Or 8
                    If map_current.map_grid(X, Y).decoration(LoopC).Render_On_Top Then deco_flags = deco_flags Or 16
                    '************************
                    'Deco Flags END
                    '************************
                    
                    'Store grh index
                    Put #1, , map_current.map_grid(X, Y).decoration(LoopC).grh.grh_index
                    
                    'Store deco flags
                    Put #1, , deco_flags
                    
                    'Store subtile pos
                    Put #1, , map_current.map_grid(X, Y).decoration(LoopC).subtile_pos
                Next LoopC
            End If
            
        Next X
    Next Y
    
    'Save ini if needed
    If save_ini Then
        Map_Save_Ini_To_File left$(map_path, Len(map_path) - 3) & "ini"
    End If
    
    'Close .map file
    Close #1
    Map_Save_Map_To_File = True

End Function

Public Function Map_Load_Map(ByVal map_id As Long, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/25/2003
'Load map data from file using an map id number
'*****************************************************************
    Dim map_path As String

    'Get map file path
    map_path = Path_Resources & PATH_MAPS & "\map" & map_id & ".map"
    
    Map_Load_Map = Map_Load_Map_From_File(map_path, load_ini)
End Function

Public Function Map_Load_Map_From_File(ByVal file_path As String, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/26/2004
'Modified by Juan Martín Sotuyo Dodero
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
    
    'Erase Chars, Lights, and Particle Groups
    Char_Remove_All
    Light_Remove_All
    Particle_Group_Remove_All
    
    map_current.map_x_min = 1
    map_current.map_y_min = 1
    
    'Open .map file
    Open map_path For Binary As #1
    Seek #1, 1
    
    'map Header
    Get #1, , map_current.map_x_max
    Get #1, , map_current.map_y_max
    Get #1, , map_current.base_light_color
    
    'Clear out and resize map
    ReDim map_current.map_grid(map_current.map_x_min To map_current.map_x_max, _
                                map_current.map_y_min To map_current.map_y_max) As Map_Tile
    
    'Set map's base light
    Map_Base_Light_Fill map_current.base_light_color
    
    'Read .map file
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
            
            '.map file
            
            'Tile Flags
            Get #1, , tile_flags
            
            'Blocked
            map_current.map_grid(X, Y).blocked = (tile_flags And 4)
            
            'Layer 1
            Get #1, , TempInt
            If TempInt Then Grh_Initialize map_current.map_grid(X, Y).grh(1), TempInt
            
            'Layer 2
            If tile_flags And 1 Then
                Get #1, , TempInt
                Get #1, , layer_flags
                'Check for angle
                If layer_flags And 8 Then
                    Get #1, , TempInt2
                Else
                    TempInt2 = 0
                End If
                
                Grh_Initialize map_current.map_grid(X, Y).grh(2), TempInt, layer_flags And 1, General_Convert_Degrees_To_Radians(CSng(TempInt2)), layer_flags And 1, layer_flags And 2
            End If
            
            'Layer 3
            If tile_flags And 2 Then
                Get #1, , TempInt
                Get #1, , layer_flags
                'Check for angle
                If layer_flags And 8 Then
                    Get #1, , TempInt2
                Else
                    TempInt2 = 0
                End If
                
                Grh_Initialize map_current.map_grid(X, Y).grh(3), TempInt, layer_flags And 1, General_Convert_Degrees_To_Radians(CSng(TempInt2)), layer_flags And 1, layer_flags And 2
            End If
            
            'Base lights
            If tile_flags And 8 Then
                For LoopC = 0 To 3
                    Get #1, , TempLng
                    Map_Base_Light_Set X, Y, TempLng, LoopC
                Next LoopC
            End If
            
            'Lights
            If tile_flags And 16 Then
                Get #1, , TempLng
                Get #1, , TempByte
                Light_Create X, Y, TempLng, TempByte
            End If
            
            'Particle Streams
            If tile_flags And 64 Then
                Get #1, , TempByte
                Particle_Group_Create X, Y, StreamData(TempByte).grh_index_list(), StreamData(TempByte).rgb_list(), StreamData(TempByte).particle_count, TempByte, StreamData(TempByte).alpha_blend, _
                                    StreamData(TempByte).alive_counter, StreamData(TempByte).frame_speed, , StreamData(TempByte).x1, StreamData(TempByte).y1, StreamData(TempByte).angle, _
                                    StreamData(TempByte).vecx1, StreamData(TempByte).vecx2, StreamData(TempByte).vecy1, StreamData(TempByte).vecy2, StreamData(TempByte).life1, StreamData(TempByte).life2, _
                                    StreamData(TempByte).fric, StreamData(TempByte).spin_speedL, StreamData(TempByte).gravity, StreamData(TempByte).grav_strength, StreamData(TempByte).bounce_strength, _
                                    StreamData(TempByte).x2, StreamData(TempByte).y2, StreamData(TempByte).XMove, StreamData(TempByte).move_x1, StreamData(TempByte).move_x2, StreamData(TempByte).move_y1, _
                                    StreamData(TempByte).move_y2, StreamData(TempByte).YMove, StreamData(TempByte).spin_speedH, StreamData(TempByte).spin, StreamData(TempByte).grh_resize, _
                                    StreamData(TempByte).grh_resizex, StreamData(TempByte).grh_resizey
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
                    Else
                        TempInt2 = 0
                    End If
                    Map_Decoration_Add X, Y, TempByte2, TempInt, deco_flags And 16, deco_flags And 1, General_Convert_Degrees_To_Radians(CSng(TempInt2)), deco_flags And 2, deco_flags And 4
                Next LoopC
            End If
            
            'Non combat zone trigger
            If tile_flags And 128 Then Map_Trigger_Set X, Y, 3
        Next X
    Next Y
    
    'Close .map file
    Close #1
    
    'Load Ini file
    ReDim npc_list(0)
    ReDim item_list(0)
    ReDim exit_list(0)
    If load_ini Then
        Map_Load_Ini_From_File left$(map_path, Len(map_path) - 3) & "ini"
    End If
    
    'Set map base light
    map_current.base_light_color = map_current.map_grid(1, 1).light_base_value(0)
    
    Map_Load_Map_From_File = True
End Function

Public Function Map_Fill(ByVal grh_index As Long, ByVal layer As Long, Optional ByVal light_base_color As Long = -1, Optional ByVal alpha_blend As Boolean, _
                        Optional ByVal angle As Single, Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Modified by Juan Martín Sotuyo Dodero
'Added security checks and centered options
'**************************************************************
    Dim X As Long
    Dim Y As Long
    Dim loop_counter As Long
    
    If layer < 1 Or layer > 3 Then
        Exit Function
    End If
    If Grh_Check(grh_index) = False Then
        Exit Function
    End If
    
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
        
            'Grh
            If Map_Grh_Set(X, Y, grh_index, layer, alpha_blend, angle, h_centered, v_centered) = False Then
                Exit Function
            End If
        
            'Base light color
            If light_base_color <> -1 Then
                If Map_Base_Light_Set(X, Y, light_base_color) = False Then
                    Exit Function
                End If
            End If
        
        Next X
    Next Y
    
    Map_Fill = True
End Function

Public Function Map_Edges_Blocked_Set(ByVal edge_distance_x As Long, ByVal edge_distance_y As Long, ByVal blocked As Boolean) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'
'**************************************************************
    Dim X As Long
    Dim Y As Long
    Dim loop_counter As Long
    
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
            If X <= edge_distance_x Or Y <= edge_distance_y Then
                map_current.map_grid(X, Y).blocked = blocked
            End If

            If X > map_current.map_x_max - edge_distance_x Or Y > map_current.map_y_max - edge_distance_y Then
                map_current.map_grid(X, Y).blocked = blocked
            End If
        Next X
    Next Y
    
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

Public Function Map_Base_Light_Fill(ByVal light_base_color As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/29/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Dim X As Long
    Dim Y As Long
    
    'Store it
    map_current.base_light_color = light_base_color
    
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
            
            'Base light color
            If Map_Base_Light_Set(X, Y, light_base_color) = False Then
                Exit Function
            End If
    
        Next X
    Next Y
    
    Map_Base_Light_Fill = True
End Function

Public Function Map_Grh_Set(ByVal map_x As Long, ByVal map_y As Long, ByVal grh_index As Long, ByVal layer As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, _
                            Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    'Check
    If layer < 1 Or layer > 3 Then
        Exit Function
    End If
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    If Grh_Check(grh_index) = False Then
        Exit Function
    End If
    
    'Do it
    Grh_Initialize map_current.map_grid(map_x, map_y).grh(layer), grh_index, alpha_blend, angle, h_centered, v_centered
    
    Map_Grh_Set = True
End Function

Public Function Map_Grh_UnSet(ByVal map_x As Long, ByVal map_y As Long, ByVal layer As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Dim grh_index As Long
    
    'Check
    If layer < 1 Or layer > 3 Then
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

Public Function Map_Legal_Char_Pos_By_Heading(ByVal char_index As Long, ByVal Heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then
        Exit Function
    End If
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Dim nx As Long
        Dim ny As Long
        nx = char_list(char_index).map_x
        ny = char_list(char_index).map_y
        Convert_Heading_to_Direction Heading, nx, ny
            
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
    Dim LoopC As Long
    
    'We start from the back, to get the last light to be placed on the tile first
    LoopC = light_last
    Do Until light_list(LoopC).map_x = map_x And light_list(LoopC).map_y = map_y
        If LoopC = 0 Then
            Map_Light_Get = 0
            Exit Function
        End If
        LoopC = LoopC - 1
    Loop
    
    Map_Light_Get = LoopC
Exit Function
ErrorHandler:
    Map_Light_Get = 0
End Function

Private Sub Map_Render()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'Modified by Juan Martín Sotuyo Dodero
'Added decorations, removed layer 2
'Added Centering
'Some other small improvements
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
    
    Dim LoopC As Long
    Dim Decorations_On_Top() As decoration_list
    Dim Decorations_On_Top_Count As Long
    
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
    'Layer 1 (Floor level)
    '*********************
    view_y = 0
    For map_y = view_min_y - 1 To view_max_y + 1
        view_x = 0
        For map_x = view_min_x - 1 To view_max_x + 1
    
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
    
    '****************************
    'Decoration´s layer
    '****************************
    view_y = -view_tile_buffer + 1
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = -view_tile_buffer + 1
        For map_x = view_buffer_min_x To view_buffer_max_x
    
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start Decorations ***
                If map_current.map_grid(map_x, map_y).decoration_count > 0 Then
                    'Check each decoration and where it´s rendered
                    For LoopC = 1 To 5
                        If map_current.map_grid(map_x, map_y).decoration(LoopC).grh.grh_index > 0 Then
                            screen_x = Convert_Tile_To_View_X(view_x) + Convert_Subtile_To_View_X(map_current.map_grid(map_x, map_y).decoration(LoopC).subtile_pos) - scroll_offset_counter_x
                            screen_y = Convert_Tile_To_View_Y(view_y) + Convert_Subtile_To_View_Y(map_current.map_grid(map_x, map_y).decoration(LoopC).subtile_pos) - scroll_offset_counter_y
                            If map_current.map_grid(map_x, map_y).decoration(LoopC).Render_On_Top Then
                                'Store data to render later (it´s faster than checking each tile and subtile twice
                                Decorations_On_Top_Count = Decorations_On_Top_Count + 1
                                ReDim Preserve Decorations_On_Top(1 To Decorations_On_Top_Count)
                                Decorations_On_Top(Decorations_On_Top_Count).map_x = map_x
                                Decorations_On_Top(Decorations_On_Top_Count).map_y = map_y
                                Decorations_On_Top(Decorations_On_Top_Count).screen_x = screen_x
                                Decorations_On_Top(Decorations_On_Top_Count).screen_y = screen_y
                                Decorations_On_Top(Decorations_On_Top_Count).decoration_index = LoopC
                            Else
                                Grh_Render map_current.map_grid(map_x, map_y).decoration(LoopC).grh, screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value()
                            End If
                        End If
                    Next LoopC
                End If
                '*** End Decorations ***
            End If
            
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Middle layer (wall´s level)
    '*********************
    view_y = -view_tile_buffer + 1
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = -view_tile_buffer + 1
        For map_x = view_buffer_min_x To view_buffer_max_x
    
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start Layer 2 ***
                If map_current.map_grid(map_x, map_y).grh(2).grh_index Then
                    'Figure out screen position
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(2), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value()
                End If
                '*** End Layer 2 ***
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
    'Decoration´s Layer - Higher Z-order
    '*********************
    If Decorations_On_Top_Count > 0 Then
        For LoopC = 1 To Decorations_On_Top_Count
            Grh_Render map_current.map_grid(Decorations_On_Top(LoopC).map_x, Decorations_On_Top(LoopC).map_y).decoration(Decorations_On_Top(LoopC).decoration_index).grh, _
                        Decorations_On_Top(LoopC).screen_x, Decorations_On_Top(LoopC).screen_y, map_current.map_grid(Decorations_On_Top(LoopC).map_x, Decorations_On_Top(LoopC).map_y).light_value()
        Next LoopC
    End If
    
    '*********************
    'Layer 3 (Roof level)
    '*********************
    view_y = -view_tile_buffer + 1
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = -view_tile_buffer + 1
        For map_x = view_buffer_min_x To view_buffer_max_x
    
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start particle effects ***
                If map_current.map_grid(map_x, map_y).particle_group_index Then
                    screen_x = (Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x) + base_tile_size / 2
                    screen_y = (Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y) + base_tile_size / 2
                    Particle_Group_Render map_current.map_grid(map_x, map_y).particle_group_index, screen_x, screen_y
                End If
                '*** End particle effects ***
                '*** Start Layer 3 ***
                If map_current.map_grid(map_x, map_y).grh(3).grh_index Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(3), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value()
                End If
                '*** End Layer 3 ***
            End If
            
            view_x = view_x + 1
        Next map_x
        
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Special Layer
    '*********************
    If engine_show_blocked_tiles Or engine_show_special_tiles Or engine_show_triggers Then
        Dim grh As grh
        Dim temp_array(3) As Long
        temp_array(0) = &HFFFFFF
        temp_array(1) = &HFFFFFF
        temp_array(2) = &HFFFFFF
        temp_array(3) = &HFFFFFF
        view_y = 0
        For map_y = view_min_y - 1 To view_max_y + 1
            view_x = 0
            For map_x = view_min_x - 1 To view_max_x + 1
        
                If Map_In_Bounds(map_x, map_y) Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    
                    If engine_show_blocked_tiles Then
                        '*** Start Blocked Tiles ***
                        If map_current.map_grid(map_x, map_y).blocked Then
                            grh.grh_index = 256
                            Grh_Render grh, screen_x, screen_y, temp_array
                        End If
                        '*** End Blocked Tiles ***
                    End If
                    
                   If engine_show_special_tiles Then
                        '*** Start Special Tiles ***
                        If map_current.map_grid(map_x, map_y).exit_index Then
                            grh.grh_index = 255
                            Grh_Render grh, screen_x, screen_y, temp_array
                        End If
                        If map_current.map_grid(map_x, map_y).Item_Index Then
                            grh.grh_index = item_list(map_current.map_grid(map_x, map_y).Item_Index).c_grh_index
                            Grh_Render grh, screen_x, screen_y, temp_array, True, True
                        End If
                        If map_current.map_grid(map_x, map_y).npc_index Then
                            grh.grh_index = Char_Data_Grh_Index_Get(npc_list(map_current.map_grid(map_x, map_y).npc_index).c_char_data_index, 1, 5)
                            Grh_Render grh, screen_x, screen_y, temp_array, True, True
                        End If
                        '*** End Special Tiles ***
                    End If
                    
                    If engine_show_triggers Then
                        '*** Start Trigger ***
                        Device_Text_Render font_list(1), CStr(map_current.map_grid(map_x, map_y).trigger), screen_y, screen_x, base_tile_size, base_tile_size, &HFFFFFFFF, DT_VCENTER Or DT_CENTER
                    End If
                    
                End If
                
                view_x = view_x + 1
            Next map_x
            view_y = view_y + 1
        Next map_y
    End If
End Sub

Public Sub Engine_Render_Mini_Map_To_hDC(ByVal destHDC As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/17/2004
'Renders the mini map to the given hDC
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim screen_x As Long
    Dim screen_y As Long
    Dim grh_index As Long
    
    '*********************
    'Layer 1 (Floor level) and walls
    '*********************
    For map_y = map_current.map_x_min To map_current.map_x_max
        For map_x = map_current.map_y_min To map_current.map_y_max
            screen_x = (map_x - 1) * 2
            screen_y = (map_y - 1) * 2
            '*** Start Layer 1 ***
            If map_current.map_grid(map_x, map_y).grh(1).grh_index Then
                grh_index = map_current.map_grid(map_x, map_y).grh(1).grh_index
                SetPixel destHDC, map_x, map_y, grh_list(grh_index).mini_map_color
            End If
            '*** End Layer 1 ***
        Next map_x
    Next map_y
    
    For map_y = map_current.map_x_min To map_current.map_x_max
        For map_x = map_current.map_y_min To map_current.map_y_max
            screen_x = (map_x - 1) * 2
            screen_y = (map_y - 1) * 2
            '*** Start Layer 2 ***
            If map_current.map_grid(map_x, map_y).grh(2).grh_index Then
                grh_index = map_current.map_grid(map_x, map_y).grh(2).grh_index
                SetPixel destHDC, map_x, map_y, grh_list(grh_index).mini_map_color
            End If
            '*** End Layer 2 ***
        Next map_x
    Next map_y
End Sub

Public Function Font_Create(ByVal style As String, ByVal size As Long, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'Returns the font_index if successful, else 0
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    
    Font_Make Font_Create, style, bold, italic, size
    
    Exit Function
ErrorHandler:
    Font_Create = 0
End Function

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal size As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/28/2004
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont 'We use this temporarily to setup the font
    Dim fnt As New StdFont 'This is also used to describe and setup the font
    fnt.name = style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    
    If font_last = 1 And font_count = 1 Then
        ReDim Preserve font_primitive_list(1 To font_last)
        Set font_primitive_list(font_last) = fnt
    ElseIf font_last > UBound(font_primitive_list()) Then
        ReDim Preserve font_primitive_list(1 To font_last)
        Set font_primitive_list(font_last) = fnt
    End If
    
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

Public Function Char_Create(ByVal map_x As Long, ByVal map_y As Long, ByVal Heading As Long, _
                            ByVal char_data_index As Long, Optional ByVal id As Long = 0) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Returns the char_index if successful, else 0
'**************************************************************
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then
        Exit Function
    End If

    If Map_Char_Get(map_x, map_y) = 0 Then
        Char_Create = Char_Next_Open
        Char_Make Char_Create, map_x, map_y, Heading, char_data_index, 1, id
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
    Dim LoopC As Long
    
    LoopC = 1
    Do Until char_list(LoopC).active = False
        If LoopC = char_last Then
            Char_Next_Open = char_last + 1
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    Char_Next_Open = LoopC
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
    Dim LoopC As Long
    
    LoopC = 1
    Do Until char_list(LoopC).id = id
        If LoopC = char_last Then
            Char_Find = 0
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    Char_Find = LoopC
Exit Function
ErrorHandler:
    Char_Find = 0
End Function

Public Function Char_Move(ByVal char_index As Long, ByVal Heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    Dim temp_x As Long
    Dim temp_y As Long
    
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then
        Char_Move = False
        Exit Function
    End If
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Make sure it's a legal move
        temp_x = char_list(char_index).map_x
        temp_y = char_list(char_index).map_y
        Convert_Heading_to_Direction Heading, temp_x, temp_y
        If Map_In_Bounds(temp_x, temp_y) Then
        
            'check for another char_index
            If map_current.map_grid(temp_x, temp_y).char_index = 0 Then
                'Move it
                Char_Move_By_Head char_index, Heading
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

Private Sub Char_Make(ByVal char_index As Long, ByVal map_x As Long, ByVal map_y As Long, ByVal Heading As Long, _
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
    char_list(char_index).Heading = Heading
    
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
    Dim LoopC As Long
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
    Dim LoopC As Long
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Update char data
        char_list(char_index).chr_data_body_index = body_index
        'Set loop status
        For LoopC = 1 To 8
            char_list(char_index).chr_data.body_list(body_index).body(LoopC).frame_counter = 1
            char_list(char_index).chr_data.body_list(body_index).body(LoopC).noloop = noloop
        Next LoopC
        
        Char_Data_Body_Set = True
    End If
End Function

Public Function Char_Heading_Set(ByVal char_index As Long, ByVal Heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Changes the character heading
'*****************************************************************
    'Make sure it´s a legal heading
    If Heading < 1 Or Heading > 8 Then
        Char_Heading_Set = False
        Exit Function
    End If
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        char_list(char_index).Heading = Heading
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

Private Sub Char_Move_By_Head(ByVal char_index As Long, ByVal Heading As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim X As Long
    Dim Y As Long
    Dim nx As Long
    Dim ny As Long
    
    X = char_list(char_index).map_x
    Y = char_list(char_index).map_y
    
    nx = X
    ny = Y
    Convert_Heading_to_Direction Heading, nx, ny
    
    map_current.map_grid(nx, ny).char_index = char_index
    char_list(char_index).map_x = nx
    char_list(char_index).map_y = ny
    map_current.map_grid(X, Y).char_index = 0
    
    char_list(char_index).scroll_offset_counter_x = (base_tile_size * (X - nx))
    char_list(char_index).scroll_offset_counter_y = (base_tile_size * (Y - ny))
    char_list(char_index).scroll_direction_x = nx - X
    char_list(char_index).scroll_direction_y = ny - Y
    
    char_list(char_index).scroll_on = True
    char_list(char_index).Heading = Heading
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
    Dim LoopC As Long
    Dim LoopC2 As Long
    Dim loopc3 As Long
    Dim inipath As String
    
    inipath = Path_Resources & PATH_GRAPHICS & "\" & "char.ini"
    
    'Get number of bodies
    char_data_count = Val(General_Var_Get(inipath, "INIT", "char_data_count"))
    
    'Resize array
    ReDim char_data_list(1 To char_data_count) As Char_Data
    
    'Fill list
    For LoopC = 1 To char_data_count
    
        'Get list count
        char_data_list(LoopC).body_count = Val(General_Var_Get(inipath, "CharData" & LoopC, "body_count"))
        
        'Get label offset
        char_data_list(LoopC).label_offset_x = Val(General_Var_Get(inipath, "CharData" & LoopC, "label_offset_x"))
        char_data_list(LoopC).label_offset_y = Val(General_Var_Get(inipath, "CharData" & LoopC, "label_offset_y"))
        
        'Redim body list
        ReDim char_data_list(LoopC).body_list(1 To char_data_list(LoopC).body_count)
        
        'Fill body list
        For LoopC2 = 1 To char_data_list(LoopC).body_count
            For loopc3 = 1 To 8
                Grh_Initialize char_data_list(LoopC).body_list(LoopC2).body(loopc3), Val(General_Var_Get(inipath, "CharData" & LoopC, LoopC2 & "_" & loopc3))
            Next loopc3
        Next LoopC2
    
    Next LoopC
End Sub

Private Sub Char_Render(ByRef temp_char As char, ByVal screen_x As Long, ByVal screen_y As Long, ByRef light_value() As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/05/2004
'Renders a character at a paticular screen point
'Edited By Juan Martín Sotuyo Dodero
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
    Grh_Render temp_char.chr_data.body_list(temp_char.chr_data_body_index).body(temp_char.Heading), screen_x, screen_y, light_value(), True

    'Do label
    If temp_char.label <> "" Then
        Device_Text_Render font_list(temp_char.label_font_index), _
                         temp_char.label, _
                         screen_y + temp_char.label_offset_y, _
                         screen_x + temp_char.label_offset_x + base_tile_size / 4 - 150, _
                         300, 20, _
                         &HFFFFFFFF, DT_TOP Or DT_CENTER
    End If
    
    'Do overhead chats
    If temp_char.Chat <> "" Then
        'Check life time
        If temp_char.chat_death_time <= GetTickCount Then
            temp_char.Chat = ""
            Exit Sub
        End If
        
        Device_Text_Render font_list(temp_char.chat_font_index), _
                            temp_char.Chat, _
                            screen_y - 30, _
                            screen_x - 150, _
                            300, 20, _
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
                            Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        If Map_Light_Get(map_x, map_y) <> 0 Then
            Light_Create = 0
            Exit Function
        End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function

Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal char_index
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If Map_In_Bounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim nx As Long
    Dim ny As Long
    
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal char_index
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        
        nx = map_x
        ny = map_y
        
        Convert_Heading_to_Direction Heading, nx, ny
        
        'Make sure it's a legal move
        If Map_In_Bounds(nx, ny) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_x = nx
            light_list(light_index).map_y = ny
    
            Light_Move_By_Head = True
            
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

Private Sub Light_Render(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 11/11/2004
'Renders a light
'***************************************'
    Dim min_x As Long
    Dim min_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim X As Long
    Dim Y As Long
    Dim color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    color = light_list(light_index).color
    
    'Arrange corners
    'NE
    If Map_In_Bounds(min_x, min_y) Then
        map_current.map_grid(min_x, min_y).light_value(2) = color
    End If
    'NW
    If Map_In_Bounds(max_x, min_y) Then
        map_current.map_grid(max_x, min_y).light_value(0) = color
    End If
    'SW
    If Map_In_Bounds(max_x, max_y) Then
        map_current.map_grid(max_x, max_y).light_value(1) = color
    End If
    'SE
    If Map_In_Bounds(min_x, max_y) Then
        map_current.map_grid(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, min_y) Then
            map_current.map_grid(X, min_y).light_value(0) = color
            map_current.map_grid(X, min_y).light_value(2) = color
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, max_y) Then
            map_current.map_grid(X, max_y).light_value(1) = color
            map_current.map_grid(X, max_y).light_value(3) = color
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, Y) Then
            map_current.map_grid(min_x, Y).light_value(2) = color
            map_current.map_grid(min_x, Y).light_value(3) = color
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, Y) Then
            map_current.map_grid(max_x, Y).light_value(0) = color
            map_current.map_grid(max_x, Y).light_value(1) = color
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If Map_In_Bounds(X, Y) Then
                map_current.map_grid(X, Y).light_value(0) = color
                map_current.map_grid(X, Y).light_value(1) = color
                map_current.map_grid(X, Y).light_value(2) = color
                map_current.map_grid(X, Y).light_value(3) = color
            End If
        Next Y
    Next X
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
    
    LoopC = 1
    Do Until light_list(LoopC).active = False
        If LoopC = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    Light_Next_Open = LoopC
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
    Dim LoopC As Long
    
    LoopC = 1
    Do Until light_list(LoopC).id = id
        If LoopC = light_last Then
            Light_Find = 0
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    Light_Find = LoopC
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
'Correctly erases a light
'***************************************'
    Dim min_x As Long
    Dim min_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim X As Long
    Dim Y As Long
    
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
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, min_y) Then
            map_current.map_grid(X, min_y).light_value(0) = map_current.map_grid(X, min_y).light_base_value(0)
            map_current.map_grid(X, min_y).light_value(2) = map_current.map_grid(X, min_y).light_base_value(2)
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, max_y) Then
            map_current.map_grid(X, max_y).light_value(1) = map_current.map_grid(X, max_y).light_base_value(1)
            map_current.map_grid(X, max_y).light_value(3) = map_current.map_grid(X, max_y).light_base_value(3)
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, Y) Then
            map_current.map_grid(min_x, Y).light_value(2) = map_current.map_grid(min_x, Y).light_base_value(2)
            map_current.map_grid(min_x, Y).light_value(3) = map_current.map_grid(min_x, Y).light_base_value(3)
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, Y) Then
            map_current.map_grid(max_x, Y).light_value(0) = map_current.map_grid(max_x, Y).light_base_value(0)
            map_current.map_grid(max_x, Y).light_value(1) = map_current.map_grid(max_x, Y).light_base_value(1)
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If Map_In_Bounds(X, Y) Then
                map_current.map_grid(X, Y).light_value(0) = map_current.map_grid(X, Y).light_base_value(0)
                map_current.map_grid(X, Y).light_value(1) = map_current.map_grid(X, Y).light_base_value(1)
                map_current.map_grid(X, Y).light_value(2) = map_current.map_grid(X, Y).light_base_value(2)
                map_current.map_grid(X, Y).light_value(3) = map_current.map_grid(X, Y).light_base_value(3)
            End If
        Next Y
    Next X
End Sub

Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
    
    LoopC = 1
    Do Until particle_group_list(LoopC).active = False
        If LoopC = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    Particle_Group_Next_Open = LoopC
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

Public Function Particle_Group_Create(ByVal map_x As Long, ByVal map_y As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Long, Optional ByVal move_x2 As Long, _
                                        Optional ByVal move_y1 As Long, Optional ByVal move_y2 As Long, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Long, Optional grh_resizey As Long)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
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
    Dim LoopC As Long
    
    LoopC = 1
    Do Until particle_group_list(LoopC).id = id
        If LoopC = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    Particle_Group_Find = LoopC
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
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Long, Optional ByVal move_x2 As Long, _
                                Optional ByVal move_y1 As Long, Optional ByVal move_y2 As Long, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Long, Optional grh_resizey As Long)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
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
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
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
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim LoopC As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    'Set colors
    temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
    temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
    temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
    temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    
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
        For LoopC = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(LoopC), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(General_Random_Number(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
                            particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
        Next LoopC
        
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

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Long, ByVal screen_y As Long, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Long, Optional ByVal move_x2 As Long, _
                            Optional ByVal move_y1 As Long, Optional ByVal move_y2 As Long, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Long, Optional grh_resizey As Long)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            Grh_Initialize temp_particle.grh, grh_index, alpha_blend
            temp_particle.X = General_Random_Number(x1, x2) - (base_tile_size / 2)
            temp_particle.Y = General_Random_Number(y1, y2) - (base_tile_size / 2)
            temp_particle.vector_x = General_Random_Number(vecx1, vecx2)
            temp_particle.vector_y = General_Random_Number(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = General_Random_Number(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.grh.angle = temp_particle.grh.angle + (General_Random_Number(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
            
            If XMove = True Then temp_particle.vector_x = General_Random_Number(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = General_Random_Number(move_y1, move_y2)
        End If
        
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
    'Draw it
    If grh_resize = True Then
        If temp_particle.grh.grh_index Then
             Grh_Render_Advance temp_particle.grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True
            Exit Sub
        End If
    End If
    If temp_particle.grh.grh_index Then
        Grh_Render temp_particle.grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, rgb_list(), True, True
    End If
End Sub

Public Function GUI_Box_Filled_Render(ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, _
                            ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional alpha_blend As Boolean) As Boolean
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/15/2003
'Draws a filled box
'Modified by Juan Martín Sotuyo Dodero
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
        .bottom = Y + height - 1
        .left = X
        .Right = X + width - 1
        .top = Y
    End With
    
    Vertex(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 1, Color1, 0, 0, 0)
    Vertex(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color2, 0, 0, 0)
    Vertex(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 1, Color3, 0, 0, 0)
    Vertex(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color4, 0, 0, 0)
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Else
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    ddevice.SetTexture 0, Nothing
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
    
    GUI_Box_Filled_Render = True
End Function

Public Function GUI_Box_Outline_Render(ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, _
                            ByVal border_width As Long, ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional alpha_blend As Boolean) As Boolean
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/15/2003
'Draws a box outline
'Modified by Juan Martín Sotuyo Dodero
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
        .bottom = Y + height - 1
        .left = X
        .Right = X + width - 1
        .top = Y
    End With
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Else
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
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
    
    GUI_Box_Outline_Render = True
End Function

Public Function GUI_Text_Render(ByVal text As String, ByVal font_index As Long, ByVal X As Long, _
                                ByVal Y As Long, ByVal width As Long, _
                                ByVal height As Long, ByVal color As Long, ByVal Alignment As FontAlignment, Optional ByVal word_wrap As Boolean) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'
'**************************************************************
    If Font_Check(font_index) = False Then
        Exit Function
    End If
    
    If word_wrap Then
        Alignment = Alignment Or DT_WORDBREAK Or DT_EDITCONTROL
    End If
    
    Device_Text_Render font_list(font_index), text, Y, X, width, height, color, Alignment
    
    GUI_Text_Render = True
End Function

Public Function GUI_Grh_Render(ByVal grh_index As Long, X As Long, Y As Long, Optional ByVal angle As Single, Optional ByVal alpha_blend As Boolean) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Dim temp_grh As grh
    Dim rgb_list(3) As Long

    If Grh_Check(grh_index) = False Then
        Exit Function
    End If

    rgb_list(0) = &HFFFFFFFF
    rgb_list(1) = &HFFFFFFFF
    rgb_list(2) = &HFFFFFFFF
    rgb_list(3) = &HFFFFFFFF

    Grh_Initialize temp_grh, grh_index, alpha_blend, angle
    
    Grh_Render temp_grh, X, Y, rgb_list
    
    GUI_Grh_Render = True
End Function

Public Function Map_Bounds_Get_From_File(ByVal map_name As String, ByRef max_x As Long, ByRef max_y As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/23/2003
'
'**************************************************************
    Dim TempLng As Long
    
    map_name = Path_Resources & PATH_MAPS & "\" & map_name & ".map"
    
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

Public Function Map_Trigger_Set(ByVal map_x As Long, ByVal map_y As Long, ByVal trigger_id As Long) As Boolean
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/27/03
'Sets a trigger at the given position
'*************************************
    If Map_In_Bounds(map_x, map_y) Then
        map_current.map_grid(map_x, map_y).trigger = trigger_id
        Map_Trigger_Set = True
    End If
End Function

Public Function Map_Trigger_Unset(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/27/03
'Erases a trigger from the given pos
'*************************************
    If Map_In_Bounds(map_x, map_y) Then
        map_current.map_grid(map_x, map_y).trigger = 0
        Map_Trigger_Unset = True
    End If
End Function

Public Function Map_Trigger_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/27/03
'Returns the trigger value of the given tile
'*************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Trigger_Get = map_current.map_grid(map_x, map_y).trigger
    End If
End Function

Public Function Engine_Triggers_Show_Toggle()
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/27/03
'
'*************************************
    If engine_show_triggers Then
        engine_show_triggers = False
    Else
        engine_show_triggers = True
    End If
End Function

Public Function Map_Grh_Info_Get(ByVal map_x As Long, ByVal map_y As Long, ByVal layer As Long, ByRef alpha_blending As Boolean, ByRef angle As Single, ByRef h_centered As Boolean, ByRef v_centered As Boolean) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 12/10/2003
'If the tile has a Grh returns it´s index, and sets the angle
'and alpha_blending status to the given variables
'*****************************************************************
    'Check
    If layer < 1 Or layer > 4 Then
        Map_Grh_Info_Get = 0
        Exit Function
    End If
    
    If Map_In_Bounds(map_x, map_y) Then
        Map_Grh_Info_Get = map_current.map_grid(map_x, map_y).grh(layer).grh_index
        angle = map_current.map_grid(map_x, map_y).grh(layer).angle
        alpha_blending = map_current.map_grid(map_x, map_y).grh(layer).alpha_blend
        h_centered = map_current.map_grid(map_x, map_y).grh(layer).h_centered
        v_centered = map_current.map_grid(map_x, map_y).grh(layer).v_centered
    Else
        Map_Grh_Info_Get = 0
    End If
End Function

Public Function Light_Info_Get(ByVal light_index As Long, ByRef map_x As Long, ByRef map_y As Long, ByRef color As Long, ByRef range As Byte) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 7/19/2003
'Gets the data of the given light
'*****************************************************************
    If Not Light_Check(light_index) Then
        Light_Info_Get = False
        Exit Function
    End If
    
    map_x = light_list(light_index).map_x
    map_y = light_list(light_index).map_y
    color = light_list(light_index).color
    range = light_list(light_index).range
    
    Light_Info_Get = True
End Function

Public Function Light_Count_Get() As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 7/19/2003
'Gets the number of existing lights
'*****************************************************************
    Light_Count_Get = UBound(light_list)
End Function

Public Sub Fill_Grh_List(GrhList As ListBox)
'****************************
'Fill Grh List
'Author: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'****************************
Dim LoopC As Long

For LoopC = 1 To grh_count
    If grh_list(LoopC).frame_count > 0 Then
        GrhList.AddItem LoopC
    End If
Next LoopC

End Sub

Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    Else
        Particle_Type_Get = 0
    End If
End Function

Public Function GUI_Grh_Render_Advance(ByVal grh_index As Long, X As Long, Y As Long, ByVal width As Long, ByVal height As Long, Optional ByVal angle As Single, Optional ByVal alpha_blend As Boolean) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/19/2003
'Similar to GUI_Grh_Rebder, but let´s you resize the Grh
'**************************************************************
    Dim temp_grh As grh
    Dim rgb_list(3) As Long

    If Grh_Check(grh_index) = False Then
        Exit Function
    End If

    rgb_list(0) = &HFFFFFFFF
    rgb_list(1) = &HFFFFFFFF
    rgb_list(2) = &HFFFFFFFF
    rgb_list(3) = &HFFFFFFFF

    Grh_Initialize temp_grh, grh_index, alpha_blend, angle
    
    Grh_Render_Advance temp_grh, X, Y, height, width, rgb_list
    
    GUI_Grh_Render_Advance = True
End Function

Private Sub Grh_Render_Advance(ByRef grh As grh, ByVal screen_x As Long, ByVal screen_y As Long, ByVal height As Long, ByVal width As Long, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
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
    
    'Center Grh over X, Y pos
    If h_center Then
        tile_width = grh_list(grh_index).src_width / base_tile_size
        If tile_width <> 1 Then
            screen_x = screen_x - Int(tile_width * base_tile_size / 2) + base_tile_size / 2
        End If
    End If
    If v_center Then
        tile_height = grh_list(grh_index).src_height / base_tile_size
        If tile_height <> 1 Then
            screen_y = screen_y - Int(tile_height * base_tile_size / 2) + base_tile_size / 2
        End If
    End If
    
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_list(grh_index).texture_index, _
        screen_x, screen_y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        rgb_list, _
        grh_list(grh_index).src_x, grh_list(grh_index).src_y, _
        width, height, grh.alpha_blend, grh.angle
End Sub

Public Function Map_Resize(ByVal max_x As Long, ByVal max_y As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/29/2003
'Changes the map´s size
'**************************************************************
    Dim old_max_x As Long
    Dim old_max_y As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim LoopC As Long
    Dim TempMap As Map
    Dim X As Long
    Dim Y As Long
    
    'Get old map size
    old_max_x = map_current.map_x_max
    old_max_y = map_current.map_y_max
    
    'Check if map was shrinked
    If max_x < old_max_x Then
        For map_x = old_max_x To max_x
            For map_y = 1 To max_y
                'Remove Particle streams
                If map_current.map_grid(map_x, map_y).particle_group_index > 0 Then
                    Particle_Group_Remove map_current.map_grid(map_x, map_y).particle_group_index
                End If
                'Remove exits
                Map_Exit_Remove map_x, map_y
            Next map_y
        Next map_x
    End If
    If max_y < old_max_y Then
        For map_y = old_max_y To max_x
            For map_x = 1 To max_x
                'Remove Particle streams
                If map_current.map_grid(map_x, map_y).particle_group_index > 0 Then
                    Particle_Group_Remove map_current.map_grid(map_x, map_y).particle_group_index
                End If
                'Remove exits
                Map_Exit_Remove map_x, map_y
            Next map_x
        Next map_y
    End If
    'Remove lights which will be left outside of the map
    If max_y < old_max_y Or max_x < old_max_x Then
        LoopC = light_last
        If light_last > 0 Then
            Do Until LoopC = 0
                If light_list(LoopC).map_x > max_x Or light_list(LoopC).map_y > max_y Then
                    Light_Remove LoopC
                End If
            Loop
        End If
    End If
    
    'Set new map size
    'Store data to a temp map
    TempMap.map_x_min = 1
    TempMap.map_y_min = 1
    TempMap.map_x_max = max_x
    TempMap.map_y_max = max_y
    
    ReDim TempMap.map_grid(1 To max_x, 1 To max_y)
    
    TempMap.map_description = map_current.map_description
    
    If max_x > old_max_x Then
        X = old_max_x
    Else
        X = max_x
    End If
    If max_y > old_max_y Then
        Y = old_max_y
    Else
        Y = max_y
    End If
    For map_x = 1 To X
        For map_y = 1 To Y
            TempMap.map_grid(map_x, map_y) = map_current.map_grid(map_x, map_y)
        Next map_y
    Next map_x
    
    'Resize map and copy evertyhing back
    map_current.map_x_min = 1
    map_current.map_y_min = 1
    map_current.map_x_max = max_x
    map_current.map_y_max = max_y
    
    map_current.map_description = TempMap.map_description
    
    ReDim map_current.map_grid(1 To max_x, 1 To max_y)
    
    For map_x = 1 To X
        For map_y = 1 To Y
            map_current.map_grid(map_x, map_y) = TempMap.map_grid(map_x, map_y)
        Next map_y
    Next map_x
    
    'If map was enlarged set all new tiles new info
    If max_x > old_max_x Then
        For map_x = old_max_x + 1 To max_x
            For map_y = 1 To max_y
                Grh_Initialize map_current.map_grid(map_x, map_y).grh(1), 1
                Map_Base_Light_Set map_x, map_y, &HAAAAAA
            Next map_y
        Next map_x
    End If
    If max_y > old_max_y Then
        For map_y = old_max_y + 1 To max_x
            For map_x = 1 To max_x
                Grh_Initialize map_current.map_grid(map_x, map_y).grh(1), 1
                Map_Base_Light_Set map_x, map_y, &HAAAAAA
            Next map_x
        Next map_y
    End If
End Function

Public Function Map_Decoration_Add(ByVal map_x As Long, ByVal map_y As Long, ByVal subtile As Byte, ByVal grh_index As Integer, _
                                    ByVal Render_On_Top As Boolean, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, _
                                    Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/07/2003
'
'**************************************************************
    Dim decoration_index As Long
    
    'Check
    If subtile < 0 Or subtile > 8 Then
        Exit Function
    End If
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    If Grh_Check(grh_index) = False Then
        Exit Function
    End If
    
    decoration_index = Decoration_Next_Open(map_x, map_y)
    If decoration_index = -1 Then
        Exit Function
    End If
    
    'Do it
    Grh_Initialize map_current.map_grid(map_x, map_y).decoration(decoration_index).grh, grh_index, alpha_blend, angle, h_centered, v_centered
    
    'Store other data
    map_current.map_grid(map_x, map_y).decoration(decoration_index).Render_On_Top = Render_On_Top
    map_current.map_grid(map_x, map_y).decoration(decoration_index).subtile_pos = subtile
    
    map_current.map_grid(map_x, map_y).decoration_count = map_current.map_grid(map_x, map_y).decoration_count + 1
    
    Map_Decoration_Add = True
End Function

Private Function Decoration_Next_Open(ByVal map_x As Long, ByVal map_y As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/07/2003
'Returns the next available index for decoration in a tile.
'If there is no free index returns -1
'**************************************************************
    Dim LoopC As Long
    
    If map_current.map_grid(map_x, map_y).decoration_count < 5 Then
        For LoopC = 1 To 5
            If map_current.map_grid(map_x, map_y).decoration(LoopC).subtile_pos = 0 Then
                Decoration_Next_Open = LoopC
                Exit Function
            End If
        Next LoopC
    End If
    
    Decoration_Next_Open = -1
End Function

Public Function Map_Decoration_Remove(ByVal map_x As Long, ByVal map_y As Long, ByVal subtile As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/07/2003
'
'**************************************************************
    Dim LoopC As Long
    
    'Check
    If subtile < 0 Or subtile > 8 Then
        Exit Function
    End If
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    
    For LoopC = 1 To 5
        If map_current.map_grid(map_x, map_y).decoration(LoopC).subtile_pos = subtile Then
            If map_current.map_grid(map_x, map_y).decoration(LoopC).grh.grh_index > 0 Then
                'Do it
                Grh_Uninitialize map_current.map_grid(map_x, map_y).decoration(LoopC).grh
                'Remove other data
                map_current.map_grid(map_x, map_y).decoration(LoopC).Render_On_Top = 0
                map_current.map_grid(map_x, map_y).decoration(LoopC).subtile_pos = 0
                map_current.map_grid(map_x, map_y).decoration_count = map_current.map_grid(map_x, map_y).decoration_count - 1
            End If
            Map_Decoration_Remove = True
            Exit Function
        End If
    Next LoopC
End Function

Public Function Map_Decoration_Info_Get(ByVal map_x As Long, ByVal map_y As Long, ByVal subtile As Long, ByRef alpha_blending As Boolean, ByRef angle As Single, ByRef h_centered As Boolean, ByRef v_centered As Boolean) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/10/2003
'Returns grh index of the given pos decoration or -1
'**************************************************************
    Dim LoopC As Long
    
    'Check
    If subtile < 0 Or subtile > 8 Then
        Exit Function
    End If
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    
    For LoopC = 1 To 5
        If map_current.map_grid(map_x, map_y).decoration(LoopC).subtile_pos = subtile Then
            alpha_blending = map_current.map_grid(map_x, map_y).decoration(LoopC).grh.alpha_blend
            angle = map_current.map_grid(map_x, map_y).decoration(LoopC).grh.angle
            h_centered = map_current.map_grid(map_x, map_y).decoration(LoopC).grh.h_centered
            v_centered = map_current.map_grid(map_x, map_y).decoration(LoopC).grh.v_centered
            Map_Decoration_Info_Get = map_current.map_grid(map_x, map_y).decoration(LoopC).grh.grh_index
            Exit Function
        End If
    Next LoopC
    
    Map_Decoration_Info_Get = -1
End Function

Private Sub Particle_Group_Load_All()
'*************************************
'Author: OneZero (onezero_ss@hotmail.com)
'Last Modified: 12/08/03
'Loads the Particles.ini file
'*************************************
    Dim LoopC As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim StreamFile As String

    StreamFile = Path_Resources & "\Particles.ini"
    
    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'Resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Particle_Group
    
    'Fill StreamData array
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).particle_count = General_Var_Get(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).x1 = General_Var_Get(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).y1 = General_Var_Get(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).x2 = General_Var_Get(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).y2 = General_Var_Get(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).angle = General_Var_Get(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = General_Var_Get(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = General_Var_Get(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = General_Var_Get(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = General_Var_Get(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = General_Var_Get(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = General_Var_Get(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).fric = General_Var_Get(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).spin = General_Var_Get(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).alpha_blend = General_Var_Get(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = General_Var_Get(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = General_Var_Get(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = General_Var_Get(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = General_Var_Get(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = General_Var_Get(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = General_Var_Get(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = General_Var_Get(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = General_Var_Get(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = General_Var_Get(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = General_Var_Get(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).speed = Val(General_Var_Get(StreamFile, Val(LoopC), "Speed"))
        
        StreamData(LoopC).grh_index_count = General_Var_Get(StreamFile, Val(LoopC), "NumGrhs")
        
        ReDim StreamData(LoopC).grh_index_list(1 To StreamData(LoopC).grh_index_count)
        GrhListing = General_Var_Get(StreamFile, Val(LoopC), "Grh_List")
        
        For i = 1 To StreamData(LoopC).grh_index_count
            StreamData(LoopC).grh_index_list(i) = General_Field_Read(str(i), GrhListing, 44)
        Next i
        
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).rgb_list(ColorSet - 1) = RGB(Val(General_Field_Read(1, TempSet, 44)), Val(General_Field_Read(2, TempSet, 44)), Val(General_Field_Read(3, TempSet, 44)))
        Next ColorSet
    Next LoopC
    
End Sub

Private Function Convert_Subtile_To_View_X(ByVal subtile As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/08/2003
'Convert tile position into position in tile
'**************************************************************
    Convert_Subtile_To_View_X = (subtile Mod 3) * base_tile_size / 3
End Function

Private Function Convert_Subtile_To_View_Y(ByVal subtile As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/08/2003
'Convert tile position into position in tile
'**************************************************************
    Convert_Subtile_To_View_Y = Fix(subtile / 3) * base_tile_size / 3
End Function

Public Function Map_Decoration_Get_Centered(ByVal map_x As Long, ByVal map_y As Long, ByVal subtile As Long, ByRef h_centered As Boolean, ByRef v_centered As Boolean)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/09/2003
'Returns a grh centered value
'**************************************************************
    Dim LoopC As Long
    
    'Check
    If Not Map_In_Bounds(map_x, map_y) Then
        Exit Function
    End If
    
    For LoopC = 1 To 5
        If map_current.map_grid(map_x, map_y).decoration(LoopC).subtile_pos = subtile Then
            h_centered = map_current.map_grid(map_x, map_y).decoration(LoopC).grh.h_centered
            v_centered = map_current.map_grid(map_x, map_y).decoration(LoopC).grh.v_centered
        End If
    Next LoopC
End Function

Private Function Convert_View_To_Subtile(ByVal view_x As Long, ByVal view_y As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/09/2003
'Calculates the subtile beneath the mouse pos
'**************************************************************
    Convert_View_To_Subtile = Fix((view_x Mod base_tile_size) * 3 / base_tile_size) + Fix((view_y Mod base_tile_size) * 3 / base_tile_size) * 3
End Function

Public Function Input_Mouse_Subtile_Get() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/09/2003
'Returns the subtile beneath the mouse pos
'**************************************************************
    Input_Mouse_Subtile_Get = input_mouse_subtile
End Function

Public Function Map_Decoration_Get(ByVal map_x As Long, ByVal map_y As Long, ByVal subtile As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 12/10/2003
'Returns grh index of the given pos decoration or -1
'**************************************************************
    Dim LoopC As Long
    
    'Check
    If subtile < 0 Or subtile > 8 Then
        Exit Function
    End If
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    
    For LoopC = 1 To 5
        If map_current.map_grid(map_x, map_y).decoration(LoopC).subtile_pos = subtile Then
            Map_Decoration_Get = map_current.map_grid(map_x, map_y).decoration(LoopC).grh.grh_index
            Exit Function
        End If
    Next LoopC
    
    Map_Decoration_Get = -1
End Function

Public Function GUI_Box_Filled_Render_Advance(ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, _
                            ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Juan MArtín Sotuyo Dodero
'Last Modify Date: 2/4/2004
'Draws a filled box with  some extra options
'Modified by Juan Martín Sotuyo Dodero
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
    Dim rgb_list(1 To 4) As Long
    
    With box_rect
        .bottom = Y + height - 1
        .left = X
        .Right = X + width - 1
        .top = Y
    End With
    
    rgb_list(1) = Color1
    rgb_list(2) = Color2
    rgb_list(3) = Color3
    rgb_list(4) = Color4
    
    'Add angle
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), , angle
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Else
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    ddevice.SetTexture 0, Nothing
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
    
    GUI_Box_Filled_Render_Advance = True
End Function

Public Function Map_Load_Map_By_Name(ByVal map_name As String, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/11/2004
'Load map data from file using a map name
'*****************************************************************
    Dim map_path As String

    'Get map file path
    map_path = Path_Resources & PATH_MAPS & "\" & map_name & ".map"
    
    Map_Load_Map_By_Name = Map_Load_Map_From_File(map_path, load_ini)
End Function

Public Function Path_Resources_Get() As String
'*****************************************************************
'Author: Fredrik Alexandersson
'*****************************************************************
    Path_Resources_Get = Path_Resources
End Function

Public Function Engine_Resolution_Change(ByVal windowed As Boolean, Optional ByVal s_width As Long, Optional ByVal s_height As Long, _
                                        Optional ByVal v_left As Long = 0, Optional ByVal v_top As Long = 0, Optional ByVal v_width_in_tiles As Long = 0, _
                                        Optional ByVal v_height_in_tiles As Long = 0) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/29/2004
'Changes the resolution while in-game
'*****************************************************************
On Error GoTo ErrHandler
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
        view_screen_left = v_left 'In pixels
        view_screen_top = v_top 'In pixels
        view_screen_width = view_screen_tile_width * base_tile_size
        view_screen_height = view_screen_tile_height * base_tile_size
        view_screen_right = view_screen_left + view_screen_width - 1
        view_screen_bottom = view_screen_top + view_screen_height - 1
        screen_width = s_width
        screen_height = s_height
        engine_windowed = False
    End If
    
    Dim d3dpp2 As D3DPRESENT_PARAMETERS
    
    With d3dpp2
        If engine_windowed Then
            .windowed = 1
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
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
        .hDeviceWindow = screen_hwnd
    End With
    
    d3dpp = d3dpp2
    
    Device_Reset
Exit Function

ErrHandler:
    MsgBox "Couldn't change reolution."
End Function

Public Function Char_Chat_Set(ByVal char_index As Long, ByVal Chat As String, ByVal font_color As Long, ByVal font_index As Long) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2004
'Sets a text as chat over head to the given char index
'*****************************************************************
    If Char_Check(char_index) Then
        char_list(char_index).Chat = Chat
        char_list(char_index).chat_color = font_color
        char_list(char_index).chat_font_index = font_index
        'Give 10 secs of life to message
        char_list(char_index).chat_death_time = GetTickCount + 10000
        
        Char_Chat_Set = True
    End If
End Function

Public Sub Grh_Render_To_Hdc_Advance(ByVal grh_index As Long, destHDC As Long, ByVal screen_x As Long, ByVal screen_y As Long, ByVal width As Long, ByVal height As Long, Optional ByVal transparent As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/30/2004
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
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    
    'If it's animated switch grh_index to first frame
    If grh_list(grh_index).frame_count <> 1 Then
        grh_index = grh_list(grh_index).frame_list(1)
    End If

    file_path = Path_Resources & PATH_GRAPHICS & "\grh" & grh_list(grh_index).texture_index & ".bmp"
    src_x = grh_list(grh_index).src_x
    src_y = grh_list(grh_index).src_y
    src_width = grh_list(grh_index).src_width
    src_height = grh_list(grh_index).src_height
    
    hdcsrc = CreateCompatibleDC(destHDC)
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
    
    If Not transparent Then
        StretchBlt destHDC, screen_x, screen_y, width, height, hdcsrc, src_x, src_y, src_width, src_height, vbSrcCopy
    Else
        MaskDC = CreateCompatibleDC(destHDC)
        
        PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
        
        Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
        
        'Render tranparently
        StretchBlt destHDC, screen_x, screen_y, width, height, MaskDC, src_x, src_y, src_width, src_height, vbSrcAnd
        StretchBlt destHDC, screen_x, screen_y, width, height, hdcsrc, src_x, src_y, src_width, src_height, vbSrcPaint
        
        Call DeleteObject(SelectObject(MaskDC, PrevObj2))
        
        DeleteDC MaskDC
    End If
    
    Call DeleteObject(SelectObject(hdcsrc, PrevObj))
    
    DeleteDC hdcsrc
End Sub

Public Function Input_Mouse_Button_Left_Double_Click_Get() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Input_Mouse_Button_Left_Double_Click_Get = input_mouse_left_double_click
    input_mouse_left_double_click = False
End Function

Public Function Input_Mouse_Button_Right_Double_Click_Get() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Input_Mouse_Button_Right_Double_Click_Get = input_mouse_right_double_click
    input_mouse_right_double_click = False
End Function

Public Function Input_Mouse_Double_Click_Time_Set(ByVal time As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    input_mouse_double_click_time = time
End Function

Public Function Input_Mouse_Button_Left_Click_Get() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Input_Mouse_Button_Left_Click_Get = input_mouse_left_click
    input_mouse_left_click = False
End Function

Public Function Input_Mouse_Button_Right_Click_Get() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Input_Mouse_Button_Right_Click_Get = input_mouse_right_click
    input_mouse_right_click = False
End Function

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, ByRef MaskDC As Long, ByVal src_x As Long, ByVal src_y As Long, ByVal src_width As Long, ByVal src_height As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/30/2004
'Creates a Mask hDC, and sets the source hDC to work for trans bliting.
'**************************************************************
    Dim X As Long
    Dim Y As Long
    Dim TransColor As Long
    Dim ColorKey As String
    
    ColorKey = Hex(COLOR_KEY)
    
    'Check if it has an alpha component
    If Len(ColorKey) > 6 Then
        'get rid of alpha
        ColorKey = "&H" & Right$(ColorKey, 6)
    End If
    
    TransColor = Val(ColorKey)
    
    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For Y = src_y To src_height - src_y
        For X = src_x To src_width - src_x
            If GetPixel(hdcsrc, X, Y) = TransColor Then
                SetPixel MaskDC, X, Y, vbWhite
                SetPixel hdcsrc, X, Y, vbBlack
            Else
                SetPixel MaskDC, X, Y, vbBlack
            End If
        Next X
    Next Y
End Sub
