Attribute VB_Name = "modGame"
Option Explicit

Public Sub Main()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 7/12/2004
'**************************************************************
    game_running = True
    gumpcount = 0
    frmMain.show
    
    'Initialize the TileEngine
    ViewHeight = General_Var_Get(App.Path & "\Client.ini", "Graphics", "Height")
    ViewWidth = General_Var_Get(App.Path & "\Client.ini", "Graphics", "Width")
    fullscreen = (General_Var_Get(App.Path & "\Client.ini", "Graphics", "Fullscreen") = "true")
    If fullscreen Then
        engine_initialized = Engine.Engine_Initialize(frmMain.hwnd, frmMain.hwnd, False, App.Path & resource_path, ViewWidth, ViewHeight, 0, 0, ViewWidth / 32, ViewHeight / 32, 32, True, True)
    Else
        engine_initialized = Engine.Engine_Initialize(frmMain.hwnd, frmMain.hwnd, True, App.Path & resource_path, , , , , ViewWidth / 32, ViewHeight / 32, 32, True, True)
    End If
    
    Load_Particle_Streams App.Path & General_Var_Get(App.Path & "\client.ini", "GENERAL", "resource_path") & "\Particles.ini"
    
    'Initialize the chat object
    Chats.Initialize Engine, ViewWidth, ViewHeight
    
    'Initialize the sound engine.
    Sound.Engine_Initialize frmMain.hwnd, App.Path
    
    'Initialize the video engine.
    frmMain.MoviePic.width = frmMain.ScaleWidth
    frmMain.MoviePic.height = frmMain.ScaleHeight
    frmMain.MoviePic.top = 0
    frmMain.MoviePic.left = 0
    Video.Engine_Initialize frmMain.MoviePic.hwnd, frmMain.MoviePic.width, frmMain.MoviePic.height
    
    'Initialize the DirectPlay Client
    engine_initialized = frmMain.dp_client.Client_Initialize(Engine, _
                        "{12345678-1234-1234-1234-123456789ABC}")
    
    'Set some data in the tile engine.
    Engine.Engine_Base_Speed_Set 0.03
    
    'Font used for almost everything in our game
    Engine.Font_Create "MS Sans Serif", 8, False, False
    
    'Font used to display NPC chat
    Engine.Font_Create "MS Sans Serif", 10, False, False
    
    'Font used to display NPC´s name while chatting with them
    Engine.Font_Create "MS Sans Serif", 12, True, False
    
    'SplashScreen
    Game_Splashscreen
    
    'Play intro video file
    'Play_Video App.Path & "\test.avi"
    
    'Initialize the character creation object
    ChrCreation.Initialize
    
    'Show login screen.
    Do While game_running
        Game_Login
        If connected And game_running Then
            frmMain.dp_client.Client_Request_Gump_Page 1
            Game_Main
        End If
    Loop
    
    'Terminate session with server
    If connected Then
        frmMain.dp_client.Client_Logoff_Session_Terminate
    End If
    
    'Delete resource files
    Delete_Resources App.Path & resource_path
    
    'Clean up, logout.
    frmMain.dp_client.Client_Deinitialize
    Engine.Engine_DeInitialize
    Video.Engine_DeInitialize
    Sound.Engine_DeInitialize
    
    Set Engine = Nothing
    Set Video = Nothing
    Set Sound = Nothing
    Set GUI = Nothing
    Set ChrCreation = Nothing
    Set Chats = Nothing
    
    'Show cursor
    Do Until ShowCursor(True) > -1
        ShowCursor True
    Loop
    
    Unload frmMain
    End
End Sub

Public Sub Game_Login()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    Dim current_char As Long    'Selected char
    Dim name As String
    Dim password As String
    
    'Create Login and Account Windows
    Windows_Create_Login_Window
    Windows_Create_Account_Window
    
    'Load .ini file
    Load_Ini_File name, password
    'Write name and password
    GUI.Text_Set_Text winLogin, 0, name
    GUI.Text_Set_Text winLogin, 1, password
    'If they were valid set the checkbox to true
    If name <> "" And password <> "" Then
        GUI.CheckBox_Set_Value winLogin, 0, True
    End If
    
    'Check if we have just started the game, or we return from the game
    If connected Then
        GUI.Window_Show winAccount, True
        'Update char stats
        Dim LoopC As Long
        For LoopC = 1 To 4
            If chars(LoopC).name <> "" Then
                frmMain.dp_client.Client_Request_Char_Stats chars(LoopC).name
            End If
        Next LoopC
    Else
        'Show the login screen
        GUI.Window_Show winLogin, True
    End If
    
    If engine_initialized Then
        Run = True
        Do While Run
            If frmMain.WindowState <> vbMinimized Then
                'Start Rendering
                Engine.Engine_Render_Start
                
                'Hide the map
                Engine.GUI_Box_Filled_Render 0, 0, ViewWidth + 1, ViewHeight + 1, D3DColorARGB(255, 0, 0, 0)

                'Render the GUI
                Engine.GUI_Grh_Render_Advance 20371, 0, 0, ViewWidth, ViewHeight
                GUI.Render_Visible
                
                'Render the Mouse
                ShowCursor False
                Engine.GUI_Grh_Render 30000, MouseX, MouseY
                
                Engine.Engine_Render_End
                
                'Check mouse
                CheckGUIControls
                
                'Check the Login buttons
                If GUI.Window_Get_Active = winLogin Then
                    Select Case MouseHitButton
                        Case 0
                            'Exit game
                            Unload frmMain
                        Case 1
                            'Login (make sure the message is sent only once)
                            If Not attempting_to_connect Then
                                'Save .ini file if needed
                                If GUI.CheckBox_Get_Value(winLogin, 0) Then
                                    Save_Ini_File GUI.Text_Get_Text(winLogin, 0), GUI.Text_Get_Text(winLogin, 1)
                                Else
                                    'Clear the correspoding fields in the .ini file
                                    Save_Ini_File "", ""
                                    'Clear the corresponding fields in the GUI controls
                                    GUI.Text_Set_Text winLogin, 0, ""
                                    GUI.Text_Set_Text winLogin, 1, ""
                                End If
                                
                                If Not connected Then
                                    'Connect to server
                                    frmMain.dp_client.Client_Connect server_ip, server_port, GUI.Text_Get_Text(winLogin, 0)
                                Else
                                    'Change client info and login using account given
                                    frmMain.dp_client.Client_Info_Change GUI.Text_Get_Text(winLogin, 0)
                                    'Authenticate
                                    frmMain.dp_client.Client_Authenticate GUI.Text_Get_Text(winLogin, 1)
                                End If
                                'Make sure the message is sent only once.
                                attempting_to_connect = True
                            End If
                    End Select
                
                'Check the Account buttons
                ElseIf GUI.Window_Get_Active = winAccount Then
                    Select Case MouseHitButton
                        Case 0 To 3
                            current_char = MouseHitButton + 1
                            'Check if it´s a an existing char to arrange buttons 4 and 5
                            If chars(current_char).name <> "" Then
                                GUI.Button_Set_Unpressed_Grh winAccount, 4, 20457
                                GUI.Button_Set_Pressed_Grh winAccount, 4, 20458
                                GUI.Button_Set_Pressed_Grh winAccount, 5, 20456
                                GUI.Button_Set_Unpressed_Grh winAccount, 5, 20455
                                
                                GUI.Button_Set_Enabled winAccount, 7, True
                                
                                'Display the chars data
                                GUI.Label_Set_Text winAccount, 0, chars(current_char).name
                                GUI.Label_Set_Text winAccount, 1, ChrCreation.Race_Get_Name(chars(current_char).race)
                                GUI.Label_Set_Text winAccount, 2, ChrCreation.Class_Get_Name(chars(current_char).Class)
                                GUI.Label_Set_Text winAccount, 3, chars(current_char).level
                                GUI.Label_Set_Text winAccount, 4, ChrCreation.Alignment_Get_Name(chars(current_char).align)
                                GUI.Label_Set_Text winAccount, 5, chars(current_char).char_STR
                                GUI.Label_Set_Text winAccount, 6, chars(current_char).char_DEX
                                GUI.Label_Set_Text winAccount, 7, chars(current_char).char_CON
                                GUI.Label_Set_Text winAccount, 8, chars(current_char).char_INT
                                GUI.Label_Set_Text winAccount, 9, chars(current_char).char_WIS
                                GUI.Label_Set_Text winAccount, 10, chars(current_char).char_CHR
                            Else
                                GUI.Button_Set_Unpressed_Grh winAccount, 4, 20459
                                GUI.Button_Set_Pressed_Grh winAccount, 4, 20460
                                GUI.Button_Set_Pressed_Grh winAccount, 5, 0
                                GUI.Button_Set_Unpressed_Grh winAccount, 5, 0
                                
                                GUI.Button_Set_Enabled winAccount, 7, False
                                
                                'Erase labels
                                GUI.Label_Set_Text winAccount, 0, ""
                                GUI.Label_Set_Text winAccount, 1, ""
                                GUI.Label_Set_Text winAccount, 2, ""
                                GUI.Label_Set_Text winAccount, 3, ""
                                GUI.Label_Set_Text winAccount, 4, ""
                                GUI.Label_Set_Text winAccount, 5, ""
                                GUI.Label_Set_Text winAccount, 6, ""
                                GUI.Label_Set_Text winAccount, 7, ""
                                GUI.Label_Set_Text winAccount, 8, ""
                                GUI.Label_Set_Text winAccount, 9, ""
                                GUI.Label_Set_Text winAccount, 10, ""
                            End If
                        Case 4
                            'Create / Delete char
                            'Check which one was set
                            If chars(current_char).name = "" Then
                                'Create
                                GUI.Window_Show winAccount, False
                                'Send slot into which we will add the char
                                Game_Character_Creation current_char
                            Else
                                'Delete
'TODO: ask for confirmation before deleting
                                frmMain.dp_client.Client_Account_Char_Remove current_char
                                'Update view
                                chars(current_char).name = ""
                                GUI.Label_Set_Text winAccount, 0, ""
                                GUI.Label_Set_Text winAccount, 1, ""
                                GUI.Label_Set_Text winAccount, 2, ""
                                GUI.Label_Set_Text winAccount, 3, ""
                                GUI.Label_Set_Text winAccount, 4, ""
                                GUI.Label_Set_Text winAccount, 5, ""
                                GUI.Label_Set_Text winAccount, 6, ""
                                GUI.Label_Set_Text winAccount, 7, ""
                                GUI.Label_Set_Text winAccount, 8, ""
                                GUI.Label_Set_Text winAccount, 9, ""
                                GUI.Label_Set_Text winAccount, 10, ""
                                
                                GUI.Button_Set_Unpressed_Grh winAccount, 4, 20459
                                GUI.Button_Set_Pressed_Grh winAccount, 4, 20460
                                GUI.Button_Set_Pressed_Grh winAccount, 5, 0
                                GUI.Button_Set_Unpressed_Grh winAccount, 5, 0
                                
                                GUI.Button_Set_Unpressed_Grh winAccount, current_char - 1, 0
                                GUI.Button_Set_Pressed_Grh winAccount, current_char - 1, 0
                            End If
                        
                        Case 5
                        'Modify Class
                            
                        Case 6
                            'Cancel
                            frmMain.dp_client.Client_Logoff_Account
                            GUI.Window_Show winAccount, False
                            GUI.Window_Show winLogin, True
                            'Make sure we don´t hit this code again next loop
                            MouseHitButton = -1
                            GUI.Window_Set_Active winLogin
                        Case 7
                            'Start Game
                            If current_char > 0 And current_char < 5 Then
                                If chars(current_char).name <> "" Then
                                    frmMain.dp_client.Client_Authenticate_With_Char chars(current_char).name
                                Else
                                    MsgBox "A valid char must be selected before starting the game."
                                End If
                            Else
                                MsgBox "A valid char must be selected before starting the game."
                            End If
                            'Make sure we don´t hit this code again next loop
                            MouseHitButton = -1
                    End Select
                End If
                
                'Check Key input.
                KeysCheck False
            End If
            
            'Check if the song need looping.
            Sound.Music_MP3_Get_Loop
            
            'Let window think...
            DoEvents
        Loop
    End If
End Sub

Public Sub Game_Character_Creation(ByVal slot As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    Dim current_pic As Long
    Dim last_portrait As Long
    Dim first_portrait As Long
    Dim current_portrait As Long
    
    'Create window
    Windows_Create_CharCreation_Window
    
    'Disable group 1 (spheres), since Gladiator is set as default
    GUI.OptionButton_Group_Disable winCharCreation, 1
    
    'Set race and alignment labels to default
    GUI.Label_Set_Text winCharCreation, 0, ChrCreation.Race_Get_Name(ChrCreation.Race_Get)
    GUI.Label_Set_Text winCharCreation, 1, ChrCreation.Alignment_Get_Name(ChrCreation.Alignment_Get)
    
    'Show the Char Creation Screen
    GUI.Window_Show winCharCreation, True
    
    'Load first and last portrait
    first_portrait = Val(General_Var_Get(App.Path & resource_path & "\portraits.ini", "GENERAL", "first_portrait"))
    last_portrait = Val(General_Var_Get(App.Path & resource_path & "\portraits.ini", "GENERAL", "last_portrait"))
    
    'Look for the first human (0) portrait
    current_portrait = first_portrait
    
    If Not InStr(1, General_Var_Get(App.Path & resource_path & "\portraits.ini", "GENERAL", CStr(current_portrait)), "0", vbTextCompare) Then
        'Start searching
        Do Until InStr(1, General_Var_Get(App.Path & resource_path & "\portraits.ini", "GENERAL", CStr(current_portrait)), 0, vbTextCompare)
            current_portrait = current_portrait + 1
        Loop
    End If
    
    'Request a dice roll
    frmMain.dp_client.Client_Request_Stats_Roll
    
    If engine_initialized Then
        Run = True
        Do While Run
            If frmMain.WindowState <> vbMinimized Then
                'Start Rendering
                Engine.Engine_Render_Start
                
                'Hide the map
                Engine.GUI_Box_Filled_Render 0, 0, ViewWidth + 1, ViewHeight + 1, D3DColorARGB(255, 0, 0, 0)

                'Render the GUI
                Engine.GUI_Grh_Render_Advance 20371, 0, 0, ViewWidth, ViewHeight
                GUI.Render_Visible
                
                'Render the Mouse
                ShowCursor False
                Engine.GUI_Grh_Render 30000, MouseX, MouseY
                
                Engine.Engine_Render_End
                
                'Check mouse
                CheckGUIControls
                
                'Check all controls
                If GUI.Window_Get_Active = winCharCreation Then
                    'Option Buttons
                    'Set class
                    ChrCreation.Class_Set GUI.OptionButton_Get_Selected_From_Group(winCharCreation, 0)
                    
                    'Set sphere
                    If GUI.OptionButton_Get_Enabled(winCharCreation, 9) Then
                        'If they are enabled set the currently selected
                        ChrCreation.Sphere_Set GUI.OptionButton_Get_Selected_From_Group(winCharCreation, 1) - 8
                    Else
                        'else set none
                        ChrCreation.Sphere_Set None
                    End If
                    
                    'Set psionic powers
                    If ChrCreation.Class_Get = Psionicist Then
                        ChrCreation.Psionic_Power_Set All
                    Else
                        ChrCreation.Psionic_Power_Set GUI.OptionButton_Get_Selected_From_Group(winCharCreation, 2) - 13
                    End If
                    
                    Select Case GUI.OptionButton_Get_Selected_From_Group(winCharCreation, 0)
                        Case 0 To 2
                            'Disable group 1 (spheres)
                            GUI.OptionButton_Group_Disable winCharCreation, 1
                            'Just allow 1 psionic power
                            Windows_CharCreation_Enable_Psionic_Powers
                        Case 3
                            'Enable group 1 (spheres)
                            GUI.OptionButton_Group_Enable winCharCreation, 1
                            'Just allow 1 psionic power
                            Windows_CharCreation_Enable_Psionic_Powers
                        Case 4
                            'Disable group 1 (spheres)
                            GUI.OptionButton_Group_Disable winCharCreation, 1
                            'All psionic powers must be selected
                            Windows_CharCreation_Enable_Psionic_Powers False
                        Case 5 To 6
                            'Disable group 1 (spheres)
                            GUI.OptionButton_Group_Disable winCharCreation, 1
                            'Just allow 1 psionic power
                            Windows_CharCreation_Enable_Psionic_Powers
                        Case 7 To 8
                            'Enable group 1 (spheres)
                            GUI.OptionButton_Group_Enable winCharCreation, 1
                            'Just allow 1 psionic power
                            Windows_CharCreation_Enable_Psionic_Powers
                    End Select
                    
                    'Buttons
                    Select Case MouseHitButton
                        Case 2
                            'Portrait pic (+)
                            current_portrait = current_portrait + 1
                            Do Until InStr(1, General_Var_Get(App.Path & resource_path & "\Portraits.ini", "GENERAL", CStr(current_portrait)), CStr(ChrCreation.Race_Get), vbTextCompare)
                                'Check we don´t go past the last portrait
                                current_portrait = current_portrait + 1
                                If current_portrait > last_portrait Then
                                    'Start over
                                    current_portrait = first_portrait
                                End If
                            Loop
                            'Show the pic
                            GUI.Button_Set_Unpressed_Grh winCharCreation, 0, current_portrait
                            GUI.Button_Set_Pressed_Grh winCharCreation, 0, current_portrait
                            
                        Case 3
                            'Portrait pic (-)
                            current_portrait = current_portrait - 1
                            Do Until InStr(1, General_Var_Get(App.Path & resource_path & "\Portraits.ini", "GENERAL", CStr(current_portrait)), CStr(ChrCreation.Race_Get), vbTextCompare)
                                'Check we don´t go past the last portrait
                                current_portrait = current_portrait - 1
                                If current_portrait < first_portrait Then
                                    'Start over
                                    current_portrait = last_portrait
                                End If
                            Loop
                            'Show the pic
                            GUI.Button_Set_Unpressed_Grh winCharCreation, 0, current_portrait
                            GUI.Button_Set_Pressed_Grh winCharCreation, 0, current_portrait
                            
'TODO: Add controls for char_data_index selection
                        Case 6
                            'Race (-)
                            GUI.Label_Set_Text winCharCreation, 0, ChrCreation.Race_Change(True)
                            'Roll all stats again
                            frmMain.dp_client.Client_Request_Stats_Roll
                            Windows_CharCreation_Arrange_Classes
                            
                            'Look for a valid portrait pic
                            If Not InStr(1, General_Var_Get(App.Path & resource_path & "\Portraits.ini", "GENERAL", CStr(current_portrait)), CStr(ChrCreation.Race_Get), vbTextCompare) Then
                                current_portrait = current_portrait + 1
                                Do Until InStr(1, General_Var_Get(App.Path & resource_path & "\Portraits.ini", "GENERAL", CStr(current_portrait)), CStr(ChrCreation.Race_Get), vbTextCompare)
                                    'Check we don´t go past the last portrait
                                    current_portrait = current_portrait + 1
                                    If current_portrait > last_portrait Then
                                        'Start over
                                        current_portrait = first_portrait
                                    End If
                                Loop
                                'Show the pic
                                GUI.Button_Set_Unpressed_Grh winCharCreation, 0, current_portrait
                                GUI.Button_Set_Pressed_Grh winCharCreation, 0, current_portrait
                            End If
                            
                        Case 7
                            'Race (+)
                            GUI.Label_Set_Text winCharCreation, 0, ChrCreation.Race_Change(False)
                            'Roll all stats again
                            frmMain.dp_client.Client_Request_Stats_Roll
                            Windows_CharCreation_Arrange_Classes
                            
                            'Look for a valid portrait pic
                            If Not InStr(1, General_Var_Get(App.Path & resource_path & "\Portraits.ini", "GENERAL", CStr(current_portrait)), CStr(ChrCreation.Race_Get), vbTextCompare) Then
                                current_portrait = current_portrait + 1
                                Do Until InStr(1, General_Var_Get(App.Path & resource_path & "\Portraits.ini", "GENERAL", CStr(current_portrait)), CStr(ChrCreation.Race_Get), vbTextCompare)
                                    'Check we don´t go past the last portrait
                                    current_portrait = current_portrait + 1
                                    If current_portrait > last_portrait Then
                                        'Start over
                                        current_portrait = first_portrait
                                    End If
                                Loop
                                'Show the pic
                                GUI.Button_Set_Unpressed_Grh winCharCreation, 0, current_portrait
                                GUI.Button_Set_Pressed_Grh winCharCreation, 0, current_portrait
                            End If
                            
                        Case 8
                            'Align (-)
                            GUI.Label_Set_Text winCharCreation, 1, ChrCreation.Alignment_Change(True)
                        Case 9
                            'Align (+)
                            GUI.Label_Set_Text winCharCreation, 1, ChrCreation.Alignment_Change(False)
                        Case 10
                            'STR (-)
                            GUI.Label_Set_Text winCharCreation, 2, CStr(ChrCreation.Char_STR_Change(True))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 11
                            'STR (+)
                            GUI.Label_Set_Text winCharCreation, 2, CStr(ChrCreation.Char_STR_Change(False))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 12
                            'DEX (-)
                            GUI.Label_Set_Text winCharCreation, 3, CStr(ChrCreation.Char_DEX_Change(True))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 13
                            'DEX (+)
                            GUI.Label_Set_Text winCharCreation, 3, CStr(ChrCreation.Char_DEX_Change(False))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 14
                            'CON (-)
                            GUI.Label_Set_Text winCharCreation, 4, CStr(ChrCreation.Char_CON_Change(True))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 15
                            'CON (+)
                            GUI.Label_Set_Text winCharCreation, 4, CStr(ChrCreation.Char_CON_Change(False))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 16
                            'INT (-)
                            GUI.Label_Set_Text winCharCreation, 5, CStr(ChrCreation.Char_INT_Change(True))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 17
                            'INT (+)
                            GUI.Label_Set_Text winCharCreation, 5, CStr(ChrCreation.Char_INT_Change(False))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 18
                            'WIS (-)
                            GUI.Label_Set_Text winCharCreation, 6, CStr(ChrCreation.Char_WIS_Change(True))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 19
                            'WIS (+)
                            GUI.Label_Set_Text winCharCreation, 6, CStr(ChrCreation.Char_WIS_Change(False))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 20
                            'CHR (-)
                            GUI.Label_Set_Text winCharCreation, 7, CStr(ChrCreation.Char_CHR_Change(True))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 21
                            'CHR (+)
                            GUI.Label_Set_Text winCharCreation, 7, CStr(ChrCreation.Char_CHR_Change(False))
                            GUI.Label_Set_Text winCharCreation, 8, CStr(ChrCreation.FreePoints_Get)
                        Case 22
                            'Roll dices
                            frmMain.dp_client.Client_Request_Stats_Roll
                        Case 23
                            'Cancel
                            'Hide char creation window
                            GUI.Window_Show winCharCreation, False
                            GUI.Window_Show winAccount, True
                            'Return to Game_Login
                            Exit Sub
                        Case 24
                            'Create char
'TODO: Complete this with char_data_index !!!!
                            frmMain.dp_client.Client_Account_Char_Create GUI.Text_Get_Text(winCharCreation, 0), ChrCreation.Race_Get, ChrCreation.Class_Get, _
                                                                ChrCreation.Alignment_Get, ChrCreation.Sphere_Get, _
                                                                ChrCreation.Psionic_Power_Get, ChrCreation.Char_STR_Get, _
                                                                ChrCreation.Char_DEX_Get, ChrCreation.Char_CON_Get, ChrCreation.Char_INT_Get, _
                                                                ChrCreation.Char_WIS_Get, ChrCreation.Char_CHR_Get, current_portrait, 1, slot
                            'Store the name so we can ask for the char´s stats when ESC is hitted later on
                            chars(slot).name = GUI.Text_Get_Text(winCharCreation, 0)
                    End Select
                End If
                
                'Check Key input.
                KeysCheck False
            End If
            
            'Check if the song need looping.
            Sound.Music_MP3_Get_Loop
            
            'Let window think...
            DoEvents
        Loop
    End If
End Sub

Public Sub Game_Main()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2004
'
'**************************************************************
    Dim NPC_Response As Long
    Dim response_index As Long
    
    'Remove focus from the account window
    GUI.Window_Set_Active -1
    
    If engine_initialized Then
        Run = True
        'Check if run = True, else quit loop.
        Do While Run
            If frmMain.WindowState <> vbMinimized Then
                'Start rendering
                Engine.Engine_Render_Start
                
                'Render the GUI.
                GUI.Render_Visible
                
                'Render chats
                Chats.Render
                
'TODO: Fill this up!!!
                'Render the chat text being written
                If Writing_Chat Then
                    'Render text and background
                    Engine.GUI_Text_Render "you say: " & Chat_Text, 1, 120, ViewHeight * 8 / 9, 400, 16, &HFFFFFFFF, fa_left
                    GUI.Window_Show GUMP(2), True
                Else
                    GUI.Window_Show GUMP(2), False
                End If
                
                'Render Mouse
                ShowCursor False
                Engine.GUI_Grh_Render 30000, MouseX, MouseY
                
                'End Rendering
                Engine.Engine_Render_End
                
                'Check mouse input
                CheckGUIControls
                
                'Check NPC speeches
                If player_talking_to_NPC Then
                    Select Case GUI.Window_Get_Active
                        Case winNPCSpeech1
                            Select Case MouseHitButton
                                Case 1
                                    NPCSpeechOffset1 = NPCSpeechOffset1 - 1
                                    NPC_Speech_Render
                                    MouseHitButton = -1
                                
                                Case 2
                                    NPCSpeechOffset1 = NPCSpeechOffset1 + 1
                                    NPC_Speech_Render
                                    MouseHitButton = -1
                            End Select
                        
                        Case winNPCSpeech2
                            'A response was chosen
                            If MouseHitLabel > -1 Then
                                If GUI.Label_Get_Text(winNPCSpeech2, MouseHitLabel) <> "" Then
                                    'Check which response it belongs to
                                    NPC_Response = MouseHitLabel + 1 + NPCSpeechOffset2
                                    response_index = 0
                                    While NPC_Response > 0
                                        response_index = response_index + 1
                                        NPC_Response = NPC_Response - UBound(Current_Speech.Responses(response_index).text_line) - 1
                                    Wend
                                    frmMain.dp_client.NPC_Respond response_index
                                    MouseHitLabel = -1
                                End If
                            End If
                            
                            'A scroll button was hitted
                            Select Case MouseHitButton
                                Case 0
                                    NPCSpeechOffset2 = NPCSpeechOffset2 - 1
                                    NPC_Speech_Render
                                    MouseHitButton = -1
                                
                                Case 1
                                    NPCSpeechOffset2 = NPCSpeechOffset2 + 1
                                    NPC_Speech_Render
                                    MouseHitButton = -1
                            End Select
                    End Select
                End If
                
                'Check normal mouse input
                CheckMouse
                
                'Check keyboard input
                KeysCheck True
            End If
            
            Sound.Music_MP3_Get_Loop
            
            DoEvents
        Loop
    End If
End Sub

Public Sub Loading_Resource_Files()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/23/2004
'
'**************************************************************
On Error GoTo ErrHandler
    Dim Graphics As Long
    Dim TotalGraphics As Long
    Dim MP3 As Long
    Dim TotalMP3 As Long
    Dim MIDI As Long
    Dim TotalMidis As Long
    Dim WAV As Long
    Dim TotalWavs As Long
    Dim Scripts As Long
    Dim TotalScripts As Long
    Dim TotalPatches As Long
    Dim Patch As Long
    Dim TotalFilesInPatches As Long
    Dim CurPatch As String
    Dim CurFile As Long
    Dim FileHead As FILEHEADER
    Dim ResourceFilesSize As Currency
    Dim InfoHead() As INFOHEADER
    Dim LoopC As Long
    
    'Check out how many files are compressed in the resource files
    '***********
    ' GRAPHICS
    '***********
    If General_File_Exists(App.Path & resource_path & "\Graphics.ORE", vbNormal) Then
        Graphics = FreeFile
        Open App.Path & resource_path & "\Graphics.ORE" For Binary Access Read Lock Read Write As Graphics
        Get Graphics, 1, FileHead
        Encrypt_File_Header FileHead
        TotalGraphics = FileHead.intNumFiles
        ReDim InfoHead(FileHead.intNumFiles - 1)
        Get Graphics, , InfoHead
        For LoopC = 0 To UBound(InfoHead)
            Encrypt_Info_Header InfoHead(LoopC)
            ResourceFilesSize = ResourceFilesSize + InfoHead(LoopC).lngFileSizeUncompressed
        Next LoopC
        Close Graphics
    End If
    
    '***********
    ' MP3
    '***********
    If General_File_Exists(App.Path & resource_path & "\Hi-Def Music.ORE", vbNormal) Then
        MP3 = FreeFile
        Open App.Path & resource_path & "\Hi-Def Music.ORE" For Binary Access Read Lock Read Write As MP3
        Get MP3, 1, FileHead
        Encrypt_File_Header FileHead
        TotalMP3 = FileHead.intNumFiles
        ReDim InfoHead(FileHead.intNumFiles - 1)
        Get MP3, , InfoHead
        For LoopC = 0 To UBound(InfoHead)
            Encrypt_Info_Header InfoHead(LoopC)
            ResourceFilesSize = ResourceFilesSize + InfoHead(LoopC).lngFileSizeUncompressed
        Next LoopC
        Close MP3
    End If
    
    
    '***********
    ' MIDIs
    '***********
    If General_File_Exists(App.Path & resource_path & "\Low-Def Music.ORE", vbNormal) Then
        MIDI = FreeFile
        Open App.Path & resource_path & "\Low-Def Music.ORE" For Binary Access Read Lock Read Write As MIDI
        Get MIDI, 1, FileHead
        Encrypt_File_Header FileHead
        TotalMidis = FileHead.intNumFiles
        ReDim InfoHead(FileHead.intNumFiles - 1)
        Get MIDI, , InfoHead
        For LoopC = 0 To UBound(InfoHead)
            Encrypt_Info_Header InfoHead(LoopC)
            ResourceFilesSize = ResourceFilesSize + InfoHead(LoopC).lngFileSizeUncompressed
        Next LoopC
        Close MIDI
    End If
    
    '***********
    ' WAVs
    '***********
    If General_File_Exists(App.Path & resource_path & "\Sounds.ORE", vbNormal) Then
        WAV = FreeFile
        Open App.Path & resource_path & "\Sounds.ORE" For Binary Access Read Lock Read Write As WAV
        Get WAV, 1, FileHead
        Encrypt_File_Header FileHead
        TotalWavs = FileHead.intNumFiles
        ReDim InfoHead(FileHead.intNumFiles - 1)
        Get WAV, , InfoHead
        For LoopC = 0 To UBound(InfoHead)
            Encrypt_Info_Header InfoHead(LoopC)
            ResourceFilesSize = ResourceFilesSize + InfoHead(LoopC).lngFileSizeUncompressed
        Next LoopC
        Close WAV
    End If
    
    '***********
    ' Scripts
    '***********
    If General_File_Exists(App.Path & resource_path & "\Scripts.ORE", vbNormal) Then
        Scripts = FreeFile
        Open App.Path & resource_path & "\Scripts.ORE" For Binary Access Read Lock Read Write As Scripts
        Get Scripts, 1, FileHead
        Encrypt_File_Header FileHead
        TotalScripts = FileHead.intNumFiles
        ReDim InfoHead(FileHead.intNumFiles - 1)
        Get Scripts, , InfoHead
        For LoopC = 0 To UBound(InfoHead)
            Encrypt_Info_Header InfoHead(LoopC)
            ResourceFilesSize = ResourceFilesSize + InfoHead(LoopC).lngFileSizeUncompressed
        Next LoopC
        Close Scripts
    End If
    
    
    'Patches should be alright (since they are untouched on the FTP), so we don´t check them
    CurPatch = Dir(App.Path & resource_path & "\patch*.ORE", vbNormal)
    Do Until CurPatch = ""
        TotalPatches = TotalPatches + 1
        Patch = FreeFile
        Open CurPatch For Binary Access Read Write Lock Read Write As Patch
        Get Patch, 1, FileHead
        Encrypt_File_Header FileHead
        TotalFilesInPatches = TotalFilesInPatches + FileHead.intNumFiles
        ReDim InfoHead(FileHead.intNumFiles - 1)
        Get Patch, , InfoHead
        For LoopC = 0 To UBound(InfoHead)
            Encrypt_Info_Header InfoHead(LoopC)
            ResourceFilesSize = ResourceFilesSize + InfoHead(LoopC).lngFileSizeUncompressed
        Next LoopC
        Close Patch
        CurPatch = Dir
    Loop
    
    
    'Check there´s enough space for the source files
    If General_Drive_Get_Free_Bytes(left$(App.Path, 3)) <= ResourceFilesSize Then
        MsgBox "There is not enough free space in your hard disk. You need to have at least " & ResourceFilesSize & " free bytes.", , "Error"
        GoTo ErrHandler2
    End If
    
    frmMain.GeneralPrgBar.Max = TotalGraphics + TotalMP3 + TotalMidis + TotalWavs + TotalScripts
    
    'Load Grhs
    If General_File_Exists(App.Path & resource_path & "\Graphics.ORE", vbNormal) Then
        frmMain.ResourceFilePrgBar.Max = TotalGraphics
        frmMain.GeneralLbl.Caption = "Loading file 1 of " & str(TotalGraphics + TotalMP3 + TotalWavs + TotalMidis + TotalScripts + TotalFilesInPatches) & "..."
        If Not Extract_Files(grh, App.Path & resource_path, frmMain.ResourceFilePrgBar, frmMain.GeneralPrgBar, frmMain.GeneralLbl) Then GoTo ErrHandler
    End If
    
    'Load MP3s
    If General_File_Exists(App.Path & resource_path & "\Hi-Def Music.ORE", vbNormal) Then
        frmMain.ResourceFileLbl.Caption = "Loading Hi-Def Music..."
        frmMain.ResourceFilePrgBar.Max = TotalMP3
        frmMain.ResourceFilePrgBar.value = 0
        If Not Extract_Files(MP3, App.Path & resource_path, frmMain.ResourceFilePrgBar, frmMain.GeneralPrgBar, frmMain.GeneralLbl) Then GoTo ErrHandler
    End If
    
    'Load MIDIs
    If General_File_Exists(App.Path & resource_path & "\Low-Def Music.ORE", vbNormal) Then
        frmMain.ResourceFileLbl.Caption = "Loading Low-Def Music..."
        frmMain.ResourceFilePrgBar.Max = TotalMidis
        frmMain.ResourceFilePrgBar.value = 0
        If Not Extract_Files(MIDI, App.Path & resource_path, frmMain.ResourceFilePrgBar, frmMain.GeneralPrgBar, frmMain.GeneralLbl) Then GoTo ErrHandler
    End If
    
    'Load Wavs
    If General_File_Exists(App.Path & resource_path & "\Sounds.ORE", vbNormal) Then
        frmMain.ResourceFileLbl = "Loading Sounds..."
        frmMain.ResourceFilePrgBar.Max = TotalWavs
        frmMain.ResourceFilePrgBar.value = 0
        If Not Extract_Files(WAV, App.Path & resource_path, frmMain.ResourceFilePrgBar, frmMain.GeneralPrgBar, frmMain.GeneralLbl) Then GoTo ErrHandler
    End If
    
    'Load scripts
    If General_File_Exists(App.Path & resource_path & "\Scripts.ORE", vbNormal) Then
        frmMain.ResourceFileLbl.Caption = "Loading Scripts..."
        frmMain.ResourceFilePrgBar.Max = TotalScripts
        frmMain.ResourceFilePrgBar.value = 0
        If Not Extract_Files(Scripts, App.Path & resource_path, frmMain.ResourceFilePrgBar, frmMain.GeneralPrgBar, frmMain.GeneralLbl) Then GoTo ErrHandler
    End If
    
    'Patches
    If TotalFilesInPatches > 0 Then
        frmMain.ResourceFileLbl.Caption = "Loading new resources..."
        frmMain.ResourceFilePrgBar.Max = TotalFilesInPatches
        frmMain.ResourceFilePrgBar.value = 0
        CurPatch = Dir(App.Path & resource_path & "\patch*.ORE", vbNormal)
        Do Until CurPatch = ""
            If Not Extract_Patch(App.Path & resource_path, CurPatch, frmMain.ResourceFilePrgBar, frmMain.GeneralPrgBar, frmMain.GeneralLbl) Then GoTo ErrHandler
            CurPatch = Dir
        Loop
        
        'Now, we must create .ORE files again (if needed), so that the new resources are included
        frmMain.ResourceFileLbl.Caption = "Adding new resources..."
        If GraphicsDSO Then
            If Not Compress_Files(grh, App.Path & resource_path, App.Path & resource_path, frmMain.GeneralPrgBar) Then GoTo ErrHandler
        End If
        
        If MP3DSO Then
            If Not Compress_Files(MP3, App.Path & resource_path, App.Path & resource_path, frmMain.GeneralPrgBar) Then GoTo ErrHandler
        End If
        
        If MIDIDSO Then
            If Not Compress_Files(MIDI, App.Path & resource_path, App.Path & resource_path, frmMain.GeneralPrgBar) Then GoTo ErrHandler
        End If
        
        If WAVDSO Then
            If Not Compress_Files(WAV, App.Path & resource_path, App.Path & resource_path, frmMain.GeneralPrgBar) Then GoTo ErrHandler
        End If
        
        If ScriptsDSO Then
            If Not Compress_Files(Scripts, App.Path & resource_path, App.Path & resource_path, frmMain.GeneralPrgBar) Then GoTo ErrHandler
        End If
    End If
    
    'Start game
    frmMain.LoadingPic.visible = False
Exit Sub

ErrHandler:
    MsgBox "An error ocurred while loading the resources.", , "Error"

ErrHandler2:
    'Couldn´t compress / decompress
    Delete_Resources App.Path & resource_path
    End
End Sub

Public Sub Game_Splashscreen()
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 3/24/2004
'Modified by Juan Martín Sotuyo dodero
'**************************************************************
    Dim Ready As Boolean
    
    Ready = False
    'Create SplashScreen
    Windows_Create_Splashscreen_Window
    GUI.Window_Show winSplashScreen, True
    
    SplashScreenAlphaBlend = 255
    frmMain.tmrSplash.enabled = True
    Do While Ready = False
        Engine.Engine_Render_Start
        Engine.GUI_Box_Filled_Render 0, 0, ViewWidth + 1, ViewHeight + 1, D3DColorARGB(255, 0, 0, 0)
        GUI.Render_Visible
        Engine.GUI_Box_Filled_Render 0, 0, ViewWidth + 1, ViewHeight + 1, D3DColorARGB(SplashScreenAlphaBlend, 0, 0, 0)
        Engine.Engine_Render_End
        
        KeysCheck
        'Exit as soon as we finish with the splashscreen
        If SplashScreenAlphaBlend = 0 Then Ready = True
        DoEvents
    Loop
    
    frmMain.tmrSplash = False
End Sub

Public Sub Play_Song(song_name As String)
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 3/24/2004
'
'**************************************************************
    If General_File_Exists(App.Path & "\..\Resources\Sounds\Music\" & song_name & ".mp3", vbNormal) Then
        Sound.Music_MP3_Load App.Path & "\..\Resources\Sounds\Music\" & song_name & ".mp3"
        Sound.Music_MP3_Play
        Exit Sub
    ElseIf General_File_Exists(App.Path & "\..\Resources\Sounds\Music\" & song_name & ".mid", vbNormal) Then
' TODO: Make it load the midi file
        'Sound.Music_MP3_Load App.Path & "\..\Resources\Sounds\Music\" & song_name & ".mid"
        'Sound.Music_MP3_Play
        Exit Sub
    End If
End Sub

Public Sub Load_Particle_Streams(ByVal StreamFile As String)
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    Dim LoopC As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long

    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).name = General_Var_Get(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = General_Var_Get(StreamFile, Val(LoopC), "NumOfParticles")
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
        StreamData(LoopC).friction = General_Var_Get(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).spin = General_Var_Get(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).AlphaBlend = General_Var_Get(StreamFile, Val(LoopC), "AlphaBlend")
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
        
        StreamData(LoopC).NumGrhs = General_Var_Get(StreamFile, Val(LoopC), "NumGrhs")
        
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(LoopC), "Grh_List")
        
        For i = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(i) = General_Field_Read(str(i), GrhListing, 44)
        Next i
        StreamData(LoopC).grh_list(i - 1) = StreamData(LoopC).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).b = General_Field_Read(3, TempSet, 44)
        Next ColorSet
        
    Next LoopC
End Sub

Public Sub CheckGUIControls()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/07/2004
'
'**************************************************************
    Static LastClick As Long
    
    Engine.Input_Mouse_View_Get MouseX, MouseY
    
    GUI.ComboBox_Mouse_On_Top_Set GUI.Window_Get_Active, MouseX, MouseY
    
    'Right Click
    If Engine.Input_Mouse_Button_Right_Click_Get Then
        'if it's closeable we close the window
        If GUI.Window_Hit_Test(MouseX, MouseY, True) = -1 Then
            GUI.Window_Mouse_Set_Up
            MouseHit = -1
            MouseHitButton = -1
            Exit Sub
        End If
    End If
    
    'Left click
    If Engine.Input_Mouse_Button_Left_Get Then
        If Engine.Input_Mouse_In_View Then
            MouseHit = GUI.Window_Hit_Test(MouseX, MouseY)
            If MouseHit <> -1 Then
            
                'Check if a window was moved
                If GUI.Window_Mouse_Get_Pressed = -1 Then
                    GUI.Window_Mouse_Set_Down (MouseHit)
                    MouseHitX = MouseX - GUI.Window_Get_Left(MouseHit)
                    MouseHitY = MouseY - GUI.Window_Get_Top(MouseHit)
                ElseIf MouseHit = GUI.Window_Mouse_Get_Pressed Then
                    If MouseX <> MouseHitX Or MouseY <> MouseHitY Then
                        GUI.Window_Move MouseHit, MouseX - MouseHitX, MouseY - MouseHitY
                    End If
                End If
                
                'Leave some free time between clicks
                If LastClick + 300 >= GetTickCount Then
                    MouseHitButton = -1
                    MouseHitLabel = -1
                    Exit Sub
                Else
                    LastClick = GetTickCount
                End If
                
                'Check if any button was clicked.
                If GUI.Window_Get_Button_Count(MouseHit) > -1 Then
                    MouseHitButton = GUI.Button_Hit_Test(MouseHit, MouseX, MouseY)
                    'If it was a GUMP generated button send the message
                    Select Case GUI.Button_Get_GUMP_Type(MouseHit, MouseHitButton)
                        Case 0
                            'It's a GUMP page
                            frmMain.dp_client.Client_Request_Gump_Page GUI.Button_Get_Button_ID(MouseHit, MouseHitButton)
                        Case 1
                            'It's a normal button
                            Dim params As String
                            Dim LoopC As Long
                            params = CStr(GUI.Button_Get_Button_ID(MouseHit, MouseHitButton))
                            
                            'Add CheckBoxe's switch ids
                            If GUI.Window_Get_CheckBox_Count(MouseHit) > -1 Then
                                For LoopC = 0 To GUI.Window_Get_CheckBox_Count(MouseHit)
                                    If GUI.CheckBox_Get_Value(MouseHit, LoopC) Then
                                        params = params & "/" & CStr(GUI.CheckBox_Get_Switch_ID(MouseHit, LoopC))
                                    End If
                                Next LoopC
                            End If
                            
                            'Add Option Button's switch ids
                            If GUI.Window_Get_OptionButton_Count(MouseHit) > -1 Then
                                For LoopC = 0 To GUI.Window_Get_OptionButton_Count(MouseHit)
                                    If GUI.OptionButton_Get_Value(MouseHit, LoopC) Then
                                        params = params & "/" & CStr(GUI.OptionButton_Get_Switch_ID(MouseHit, LoopC))
                                    End If
                                Next LoopC
                            End If
                            
                            'Add texts
                            If GUI.Window_Get_Text_Count(MouseHit) > -1 Then
                                For LoopC = 0 To GUI.Window_Get_Text_Count(MouseHit)
                                    params = params & "/" & GUI.Text_Get_Text(MouseHit, LoopC)
                                Next LoopC
                            End If
                            
                            frmMain.dp_client.Client_Send_Gump_Button params
                    End Select
                End If
                
                'Check if any Textbox was clicked.
                If GUI.Window_Get_Text_Count(MouseHit) > -1 Then
                    MouseHitText = GUI.Text_Hit_Test(MouseHit, MouseX, MouseY)
                    If MouseHitText <> -1 Then
                        GUI.Text_Set_Active MouseHit, MouseHitText
                    End If
                End If
                
                'Check if any Combobox was clicked
                If GUI.Window_Get_ComboBox_Count(MouseHit) > -1 Then
                    If GUI.ComboBox_Get_ScrollButton_Clicked(MouseHit) = -1 Then
                        'Just do the hit test. The combobox is designed to do all events by itself.
                        MouseHitComboBox = GUI.ComboBox_Hit_Test(MouseHit, MouseX, MouseY)
                        MouseHitComboBoxY = MouseY - GUI.ComboBox_Get_ScrollButton_Top(MouseHit, MouseHitComboBox)
                    ElseIf MouseY <> MouseHitComboBoxY Then
                        GUI.ComboBox_ScrollButton_Move MouseHit, MouseHitComboBox, MouseY - MouseHitComboBoxY
                    End If
                End If
                
                'Check if any CheckBox was clicked
                If GUI.Window_Get_CheckBox_Count(MouseHit) > -1 Then
                    GUI.CheckBox_Hit_Test MouseHit, MouseX, MouseY
                End If
                
                'Check if any OptionButton was clicked
                If GUI.Window_Get_OptionButton_Count(MouseHit) > -1 Then
                    GUI.OptionButton_Hit_Test MouseHit, MouseX, MouseY
                End If
                
                'Check if any label was clicked
                If GUI.Window_Get_Label_Count(MouseHit) > -1 Then
                    MouseHitLabel = GUI.Label_Hit_Test(MouseHit, MouseX, MouseY)
                End If
            Else
                GUI.Window_Mouse_Reset
            End If
        End If
    Else
        GUI.Window_Mouse_Set_Up
        MouseHit = -1
        MouseHitButton = -1
    End If
End Sub

Public Sub CheckMouse()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    'Check in-game mouse input
    Static LastClick As Long
    
    Engine.Input_Mouse_View_Get MouseX, MouseY
    
    'Both buttons clicked (Non-click movement)
    If Engine.Input_Mouse_Button_Left_Get And Engine.Input_Mouse_Button_Right_Get Then
        NonClickMovement = True
    End If
    
    If NonClickMovement Then
        'Move towards mouse
        frmMain.dp_client.Player_Move Engine.Input_Mouse_Heading_Get
        'If we were talking to a NPC stop doing so
        If player_talking_to_NPC Then
            GUI.Window_Show winNPCSpeech1, False
            GUI.Window_Show winNPCSpeech2, False
            player_talking_to_NPC = False
        End If
    End If
    
    'Right click (move)
    If Engine.Input_Mouse_Button_Right_Get And Not Engine.Input_Mouse_Button_Left_Get Then
        frmMain.dp_client.Player_Move Engine.Input_Mouse_Heading_Get
        'Cancel Non-Click Movement
        NonClickMovement = False
        'If we were talking to a NPC stop doing so
        If player_talking_to_NPC Then
            GUI.Window_Show winNPCSpeech1, False
            GUI.Window_Show winNPCSpeech2, False
            player_talking_to_NPC = False
        End If
    End If
    
    'Separate any none-movements commands
    If LastClick + 300 >= GetTickCount Or GUI.Window_Get_Active <> -1 Then
        Exit Sub
    Else
        LastClick = GetTickCount
    End If
    
    'Left click (talk to NPC)
    If Engine.Input_Mouse_Button_Left_Double_Click_Get Then
        Dim map_x As Long
        Dim map_y As Long
        
        'Get map cordinates
        Engine.Input_Mouse_Map_Get map_x, map_y
        
        'Ignore it if we are already talking to a NPC
        If Not player_talking_to_NPC Then
            'Make sure they are valid
            If Engine.Map_In_Bounds(map_x, map_y) Then
                frmMain.dp_client.NPC_Talk map_x, map_y
            End If
        End If
    End If
End Sub

Public Sub Play_Video(ByVal video_path As String, Optional ByVal fullscreen As Boolean = False, Optional ByVal volume As Long = 0, Optional ByVal balance As Long = 0, Optional ByVal loop_video As Boolean = False, Optional ByVal hide_cursor As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    frmMain.MoviePic.visible = True
    
    If Video.File_Load(video_path, fullscreen, volume, balance, loop_video, hide_cursor) Then
        
        DoEvents
        Do Until Video.File_Check_If_Ended
            DoEvents
        Loop
    End If
    
    frmMain.MoviePic.visible = False
End Sub

Public Sub Save_Ini_File(ByVal account_name As String, ByVal account_password As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    General_Var_Write App.Path & "\client.ini", "ACCOUNT INFO", "name", account_name
    General_Var_Write App.Path & "\client.ini", "ACCOUNT INFO", "password", account_password
End Sub

Public Sub Load_Ini_File(ByRef account_name As String, ByRef account_password As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    account_name = General_Var_Get(App.Path & "\client.ini", "ACCOUNT INFO", "name")
    account_password = General_Var_Get(App.Path & "\client.ini", "ACCOUNT INFO", "password")
    
    'Get server IP and port
    server_ip = General_Var_Get(App.Path & "\client.ini", "GENERAL", "server_ip")
    server_port = Val(General_Var_Get(App.Path & "\client.ini", "GENERAL", "server_port"))
End Sub

Public Sub KeysCheck(Optional ByVal in_game As Boolean)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2004
'
'**************************************************************
    If connected And in_game Then
        
        'Don´t walk around while typing
        If Not Writing_Chat Then
            If Engine.Input_Key_Get(vbKeyW) And Engine.Input_Key_Get(vbKeyD) Then
                frmMain.dp_client.Player_Move 2
            End If
            
            If Engine.Input_Key_Get(vbKeyS) And Engine.Input_Key_Get(vbKeyD) Then
                frmMain.dp_client.Player_Move 4
            End If
            
            If Engine.Input_Key_Get(vbKeyS) And Engine.Input_Key_Get(vbKeyA) Then
                frmMain.dp_client.Player_Move 6
            End If
            
            If Engine.Input_Key_Get(vbKeyW) And Engine.Input_Key_Get(vbKeyA) Then
                frmMain.dp_client.Player_Move 8
            End If
            
            If Engine.Input_Key_Get(vbKeyW) Then
                frmMain.dp_client.Player_Move 1
            End If
            
            If Engine.Input_Key_Get(vbKeyD) Then
                frmMain.dp_client.Player_Move 3
            End If
            
            If Engine.Input_Key_Get(vbKeyS) Then
                frmMain.dp_client.Player_Move 5
            End If
            
            If Engine.Input_Key_Get(vbKeyA) Then
                frmMain.dp_client.Player_Move 7
            End If
            
            
            If Engine.Input_Key_Get(vbKeyR) Then
                Change_Resolution False, 800, 600
            End If
            If Engine.Input_Key_Get(vbKeyQ) Then
                frmMain.dp_client.Client_Logoff_Char
                Run = False
            End If
        End If
        
        
        If Engine.Input_Key_Get(vbKeySpace) Then
            frmMain.dp_client.Player_Attack
        End If
        
        If Engine.Input_Key_Get(vbKeyHome) Then
            frmMain.dp_client.Player_Item_Pickup
        End If
        
        If Engine.Input_Key_Get(vbKeyEscape) Then
            'Hide all in-game specyfic windows to prevent them from being displayed
            GUI.Window_Show winNPCSpeech1, False
            GUI.Window_Show winNPCSpeech2, False
            GUI.Window_Show GUMP(1), False
            frmMain.dp_client.Client_Logoff_Char
            Run = False
        End If
    End If
    
    If Engine.Input_Key_Get(vbKeyControl) And Engine.Input_Key_Get(vbKeyQ) Then
        Unload frmMain
        Exit Sub
    End If
End Sub

Public Sub NPC_Speech_Render()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2004
'Draws the NPC speech windows according to the scroll offsets
'**************************************************************
    Dim ShowScroll As Boolean
    Dim NumLines As Long
    Dim OutputLine As Long
    Dim LoopC As Long
    Dim LoopC2 As Long
    
    '****************************************
    'Arrange NPC chatting system windows
    '****************************************
    'NPC greet
    ShowScroll = (UBound(Current_Speech.NPC_greet()) > 2)
    
    GUI.Button_Set_Visible winNPCSpeech1, 1, ShowScroll
    GUI.Button_Set_Visible winNPCSpeech1, 2, ShowScroll
    
    'Check if we should enable them
    GUI.Button_Set_Enabled winNPCSpeech1, 1, False
    GUI.Button_Set_Enabled winNPCSpeech1, 2, False
    If ShowScroll Then
        If NPCSpeechOffset1 > 0 Then
            GUI.Button_Set_Enabled winNPCSpeech1, 1, True
        End If
        If NPCSpeechOffset1 < UBound(Current_Speech.NPC_greet()) - 2 Then
            GUI.Button_Set_Enabled winNPCSpeech1, 2, True
        End If
    End If
    
    'Responses
    For LoopC = 1 To UBound(Current_Speech.Responses())
        NumLines = NumLines + UBound(Current_Speech.Responses(LoopC).text_line()) + 1
    Next LoopC
    
    ShowScroll = (NumLines > 3)
    
    GUI.Button_Set_Visible winNPCSpeech2, 0, ShowScroll
    GUI.Button_Set_Visible winNPCSpeech2, 1, ShowScroll
    
    'Check if we should enable them
    GUI.Button_Set_Enabled winNPCSpeech2, 0, False
    GUI.Button_Set_Enabled winNPCSpeech2, 1, False
    If ShowScroll Then
        If NPCSpeechOffset2 > 0 Then
            GUI.Button_Set_Enabled winNPCSpeech2, 0, True
        End If
        If NPCSpeechOffset2 < NumLines - 4 Then
            GUI.Button_Set_Enabled winNPCSpeech2, 1, True
        End If
    End If
    
    '*************
    'Display it
    '*************
    'Clear labels
    For LoopC = 1 To 3
        GUI.Label_Set_Text winNPCSpeech1, LoopC, ""
    Next LoopC
    
    For LoopC = 0 To UBound(Current_Speech.NPC_greet())
        GUI.Label_Set_Text winNPCSpeech1, LoopC + 1, Current_Speech.NPC_greet(LoopC + NPCSpeechOffset1)
        If LoopC = 2 Then Exit For
    Next LoopC
    
    'Clear labels
    For LoopC = 0 To 3
        GUI.Label_Set_Text winNPCSpeech2, LoopC, ""
    Next LoopC
    
    NumLines = 0
    For LoopC = 1 To UBound(Current_Speech.Responses())
        For LoopC2 = 0 To UBound(Current_Speech.Responses(LoopC).text_line())
            If NumLines >= NPCSpeechOffset2 Then
                GUI.Label_Set_Text winNPCSpeech2, OutputLine, Current_Speech.Responses(LoopC).text_line(LoopC2)
                OutputLine = OutputLine + 1
            End If
            NumLines = NumLines + 1
            If OutputLine = 4 Then Exit For
        Next LoopC2
        If OutputLine = 4 Then Exit For
    Next LoopC
End Sub

Public Sub Change_Resolution(ByVal windowed As Boolean, ByVal s_width As Long, ByVal s_height As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/31/2004
'Changes the game resolution
'**************************************************************
'TODO: Buttons grhs (portraits) should be set
    Dim LoopC As Long
    
    'Store values
    ViewWidth = s_width
    ViewHeight = s_height
    fullscreen = Not windowed
    
    'Resize form
    frmMain.width = s_width * 15    'Multiply by 15 to convert to twips
    frmMain.height = s_height * 15
    
    'Center the form
    frmMain.left = Screen.width * Screen.TwipsPerPixelX / 2 - frmMain.width / 2
    frmMain.top = Screen.height * Screen.TwipsPerPixelY / 2 - frmMain.height / 2
    
    'Chenge resolution
    If fullscreen Then
        engine_initialized = Engine.Engine_Resolution_Change(False, ViewWidth, ViewHeight, 0, 0, ViewWidth / 32, ViewHeight / 32)
    Else
        engine_initialized = Engine.Engine_Resolution_Change(True, , , , , ViewWidth / 32, ViewHeight / 32)
    End If
    
    '************************************************
    'Recreate all windows using the new resolution
    '(windows should be created in the same order they were created initially)
    '************************************************
    'Make a backup copy of the GUI which wiil be used to copy stuff
    Dim oldGUI As New clsGui
    Set oldGUI = GUI
    
    'Delete all windows from the original object
    Set GUI = New clsGui
    
'TODO: Check how it´s created to add the grh directly.
    If winSplashScreen <> -1 Then
        winSplashScreen = -1
        Windows_Create_Splashscreen_Window
        GUI.Window_Show winSplashScreen, oldGUI.Window_Get_Visible(winSplashScreen)
    End If
    
    If winLogin <> -1 Then
        winLogin = -1
        Windows_Create_Login_Window
        GUI.Window_Show winLogin, oldGUI.Window_Get_Visible(winLogin)
        'Copy old values
        GUI.Text_Add_Text winLogin, 0, oldGUI.Text_Get_Text(winLogin, 0)
        GUI.Text_Add_Text winLogin, 1, oldGUI.Text_Get_Text(winLogin, 1)
        GUI.CheckBox_Set_Value winLogin, 0, oldGUI.CheckBox_Get_Value(winLogin, 0)
    End If
    
    If winAccount <> -1 Then
        winAccount = -1
        Windows_Create_Account_Window
        GUI.Window_Show winAccount, oldGUI.Window_Get_Visible(winAccount)
        
        'Copy button's grhs
        For LoopC = 0 To 5
            GUI.Button_Set_Pressed_Grh winAccount, LoopC, oldGUI.Button_Get_Grh_Pressed(winAccount, LoopC)
            GUI.Button_Set_Unpressed_Grh winAccount, LoopC, oldGUI.Button_Get_Grh_Unpressed(winAccount, LoopC)
        Next LoopC
        
        'Copy labels
        For LoopC = 0 To 10
            GUI.Label_Set_Text winAccount, LoopC, oldGUI.Label_Get_Text(winAccount, LoopC)
        Next LoopC
        
        'Check if we need to enable the accept button
        GUI.Button_Set_Enabled winAccount, 7, oldGUI.Button_Get_Enabled(winAccount, 7)
    End If
    
    If winCharCreation <> -1 Then
        winCharCreation = -1
        Windows_Create_CharCreation_Window
        GUI.Window_Show winCharCreation, oldGUI.Window_Get_Visible(winCharCreation)
        
        'copy button's grhs
        GUI.Button_Set_Pressed_Grh winAccount, 0, oldGUI.Button_Get_Grh_Pressed(winAccount, 0)
        GUI.Button_Set_Unpressed_Grh winAccount, 0, oldGUI.Button_Get_Grh_Unpressed(winAccount, 0)
        GUI.Button_Set_Pressed_Grh winAccount, 1, oldGUI.Button_Get_Grh_Pressed(winAccount, 1)
        GUI.Button_Set_Unpressed_Grh winAccount, 1, oldGUI.Button_Get_Grh_Unpressed(winAccount, 1)
        
        For LoopC = 25 To 27
            GUI.Button_Set_Pressed_Grh winAccount, LoopC, oldGUI.Button_Get_Grh_Pressed(winAccount, LoopC)
            GUI.Button_Set_Unpressed_Grh winAccount, LoopC, oldGUI.Button_Get_Grh_Unpressed(winAccount, LoopC)
        Next LoopC
        
        'Copy option buttons and checkboxes values
        For LoopC = 0 To 2
            GUI.OptionButton_Set_Selected winCharCreation, oldGUI.OptionButton_Get_Selected_From_Group(winCharCreation, LoopC)
            GUI.CheckBox_Set_Value winCharCreation, LoopC, oldGUI.CheckBox_Get_Value(winCharCreation, LoopC)
        Next LoopC
        
        'Enable / disable option buttons
        For LoopC = 0 To 8
            GUI.OptionButton_Set_Enabled winCharCreation, LoopC, oldGUI.OptionButton_Get_Enabled(winCharCreation, LoopC)
        Next LoopC
        
        For LoopC = 13 To 15
            GUI.OptionButton_Set_Enabled winCharCreation, LoopC, oldGUI.OptionButton_Get_Enabled(winCharCreation, LoopC)
        Next LoopC
        
        'Enable / disable checkboxes
        For LoopC = 0 To 2
            GUI.CheckBox_Set_Enabled winCharCreation, LoopC, oldGUI.CheckBox_Get_Enabled(winCharCreation, LoopC)
            GUI.CheckBox_Set_Visible winCharCreation, LoopC, oldGUI.CheckBox_Get_Visible(winCharCreation, LoopC)
        Next LoopC
        
        'Copy labels
        For LoopC = 0 To 8
            GUI.Label_Set_Text winCharCreation, LoopC, oldGUI.Label_Get_Text(winCharCreation, LoopC)
        Next LoopC
        
        'Copy textobox
        GUI.Text_Add_Text winCharCreation, 0, oldGUI.Text_Get_Text(winCharCreation, 0)
    End If
    
    If winNPCSpeech1 <> -1 Then
        winNPCSpeech1 = -1
        winNPCSpeech2 = -1
        Windows_Create_NPC_Speech_Window
        GUI.Window_Show winNPCSpeech1, oldGUI.Window_Get_Visible(winNPCSpeech1)
        GUI.Window_Show winNPCSpeech2, oldGUI.Window_Get_Visible(winNPCSpeech1)
        
        'Redraw the texts
        NPC_Speech_Render
    End If
    
    'Destroy backup copy of GUI
    Set oldGUI = Nothing
    
    'Update the chat object
    Chats.Engine_Resize ViewWidth, ViewHeight
End Sub
