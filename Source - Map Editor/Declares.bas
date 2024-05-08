Attribute VB_Name = "Declares"
Option Explicit

'The Engine object
Public Engine As New clsTileEngineX

'User-defined
Public resource_path As String
Public autosave_delay As Long
Public map_height As Long
Public map_width As Long
Public x_border As Integer
Public y_border As Integer
Public tile_size As Long 'In Pixels
Public base_speed As Single
Public char_label As String
Public use_ini_files As Boolean
Public use_resource_files As Boolean

'Tile Groups
Public Type TileGroup
    GrhIndexes() As Long
    Name As String
End Type

Public Current_Group As TileGroup
Public TileGroupOffset As Long

'Other variables
Public prgRun As Boolean 'When false the program ends
Public Modified As Boolean 'Whether the map has been modified or not
Public Walk_Mode As Boolean
Public User_Char_Index As Long
Public Cursor_Light_Index As Long
Public Light_Color As Long
Public Current_Map As String
Public angle As Long
Public exit_map_max_x As Long
Public exit_map_max_y As Long
Public current_grh As Long
Public tool As String

'Everything from now on was coded by OneZero unless it´s said do
'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream

'RGB Type
Public Type RGB
    r As Long
    g As Long
    b As Long
End Type

Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    'Added by Maraxus
    speed As Single
    life_counter As Long
    
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

'Vars needed for the Particle Editor which MUST be public
Public DataChanged As Boolean

'Added by Maraxus
Public Particle_Editor_Unloaded As Boolean

'Added by FireStarter
Public fso As New FileSystemObject, f As Scripting.TextStream
Public Tree_Cur_Grh As Long
