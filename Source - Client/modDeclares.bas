Attribute VB_Name = "modDeclares"
Option Explicit

Public Engine As New clsTileEngineX
Public Sound As New clsSoundEngine
Public GUI As New clsGui
Public Video As New clsVideoEngine
Public ChrCreation As New clsCharacterCreation
Public Chats As New clsChats

'The this values should be true or else the game will quit.
Public engine_initialized As Boolean
Public Run As Boolean
Public game_running As Boolean

'The resource path
Public resource_path As String

'Connection status
Public connected As Boolean
Public attempting_to_connect As Boolean

'Used for Non-Click Movement
Public NonClickMovement As Boolean

'The text typed by the player for chats
Public Chat_Text As String

'True if the user is typing a chat at the moment
Public Writing_Chat As Boolean

'Window IDs
Public winCharCreation As Long
Public winSplashScreen As Long
Public winAccount As Long
Public winLogin As Long
Public winNPCSpeech1 As Long    'NPC greet
Public winNPCSpeech2 As Long    'Reponses
Public GUMP(1 To 10) As Long
Public gumpcount As Long

'Mouse location
Public MouseX As Long
Public MouseY As Long
Public MouseHitX As Long
Public MouseHitY As Long
Public MouseHit As Long
Public MouseHitText As Long
Public MouseHitButton As Long
Public MouseHitComboBox As Long
Public MouseHitComboBoxY As Long    'X coord is not necessary
Public MouseHitLabel As Long

'Chars in the account
Public Type char
    name As String
    race As races
    Class As classes
    align As Alignment
    sphere As spheres
    psionic_power As psionic_powers
    level As Long
    char_STR As Long
    char_DEX As Long
    char_CON As Long
    char_INT As Long
    char_WIS As Long
    char_CHR As Long
    portrait As Long
    char_data_index As Long
End Type

Public chars(1 To 4) As char

'The View Info
Public ViewHeight As Long
Public ViewWidth As Long
Public fullscreen As Boolean

'Used for splashscreen
Public SplashScreenAlphaBlend As Integer
Public Inc As Boolean

'....................................
Public server_ip As String
Public server_port As Long
Public player_name As String
Public player_password As String
Public player_profile_name As String

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
    name As String
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
    
    speed As Single
    life_counter As Long
End Type

Public Type Lines
    text_line() As String
End Type
    
'NPC Speech stuff
Public Type NPC_Speech
    Responses() As Lines
    NPC_greet() As String
    NPC_name As String
    NPC_portrait_pic As Long
End Type

Public Current_Speech As NPC_Speech
Public player_talking_to_NPC As Boolean

Public NPCSpeechOffset1 As Long
Public NPCSpeechOffset2 As Long

'Used for NPC Chating System
Public MAX_CHARS_PER_LINE_NPC1 As Long
Public MAX_CHARS_PER_LINE_NPC2 As Long

'Inventory stuff
Public Type item
    item_data_index As Long
    amount As Long
    equiped As Boolean
End Type

Public User_Inventory(1 To 99) As item

'Outside functions
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Boolean) As Integer
'Gets number of ticks since windows started
Public Declare Function GetTickCount Lib "kernel32" () As Long
