Attribute VB_Name = "modWindows"
'******************************************************
'Coded by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'******************************************************
'Creates all windows and their correspoding controls
'******************************************************

Option Explicit

Public Sub Windows_Create_CharCreation_Window()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 10/24/2003
'
'**************************************************************
    Dim grh_height As Long
    Dim grh_width As Long
    Dim TempLng As Long
    Dim TempStr As String
    
    Engine.Grh_Info_Get 20447, TempStr, TempLng, TempLng, grh_width, grh_height, TempLng
    
    '*************************
    'Character Creation window
    '*************************
    If winCharCreation = -1 Then
        winCharCreation = GUI.Window_Create(20447, ViewHeight / 2 - grh_height / 2, ViewWidth / 2 - grh_width / 2, grh_width, grh_height)
        
        'Buttons
        GUI.Button_Create 20099, winCharCreation, "", 0, 32, 32, 63, 63     '0 : Portrait
        GUI.Button_Create 0, winCharCreation, "", 0, 32, 112, 63, 63, , &H0 '1 : Char
        
        GUI.Button_Create 0, winCharCreation, "", 0, 107, 32, 8, 8          '2 : Portrait (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 107, 87, 8, 8          '3 : Portrait (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 107, 112, 8, 8         '4 : Char (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 107, 167, 8, 8         '5 : Char (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 148, 51, 8, 8          '6 : Race (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 148, 164, 8, 8         '7 : Race (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 164, 51, 8, 8          '8 : Align (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 164, 164, 8, 8         '9 : Align (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 196, 52, 8, 8          '10 : STR (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 196, 100, 8, 8         '11 : STR (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 212, 51, 8, 8          '12 : DEX (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 212, 100, 8, 8         '13 : DEX (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 228, 51, 8, 8          '14 : CON (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 228, 100, 8, 8         '15 : CON (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 244, 51, 8, 8          '16 : INT (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 244, 100, 8, 8         '17 : INT (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 260, 51, 8, 8          '18 : WIS (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 260, 100, 8, 8         '19 : WIS (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 276, 51, 8, 8          '20 : CHR (-)
        GUI.Button_Create 0, winCharCreation, "", 0, 276, 100, 8, 8         '21 : CHR (+)
        
        GUI.Button_Create 0, winCharCreation, "", 0, 214, 266, 75, 20       '22 : Roll
        
        GUI.Button_Create 20451, winCharCreation, "", 0, 263, 210, 87, 31, 20452    '23 : Cancel
        
        GUI.Button_Create 20466, winCharCreation, "", 0, 263, 305, 87, 31, 20467    '24 : Create Char
        
'TODO: Put the correct indexes to the 3 dice buttons
        'The 3 dices (don´t work as buttons really)
        GUI.Button_Create 0, winAccount, "", 0, 179, 243, 26, 26        '25 : Dice 1
        GUI.Button_Create 0, winAccount, "", 0, 179, 291, 26, 26        '26 : Dice 2
        GUI.Button_Create 0, winAccount, "", 0, 179, 339, 26, 26        '27 : Dice 3
        
        'This button just displays the name textbox graphic
        GUI.Button_Create 20453, winCharCreation, "", 0, 127, 56, 115, 19   '28 : Textbox border
        
        'Option Buttons
        GUI.OptionButton_Create winCharCreation, 37, 212, 8, 6, 20084, 20085, 20086, "", 0    '0 : Gladiator
        GUI.OptionButton_Create winCharCreation, 53, 212, 8, 6, 20084, 20085, 20086, "", 0    '1 : Fighter
        GUI.OptionButton_Create winCharCreation, 69, 212, 8, 6, 20084, 20085, 20086, "", 0    '2 : Thief
        GUI.OptionButton_Create winCharCreation, 85, 212, 8, 6, 20084, 20085, 20086, "", 0    '3 : Ranger
        GUI.OptionButton_Create winCharCreation, 101, 212, 8, 6, 20084, 20085, 20086, "", 0   '4 : Psionicist
        GUI.OptionButton_Create winCharCreation, 117, 212, 8, 6, 20084, 20085, 20086, "", 0   '5 : Preserver
        GUI.OptionButton_Create winCharCreation, 133, 212, 8, 6, 20084, 20085, 20086, "", 0   '6 : Defiler
        GUI.OptionButton_Create winCharCreation, 37, 276, 8, 6, 20084, 20085, 20086, "", 0    '7 : Cleric
        GUI.OptionButton_Create winCharCreation, 53, 276, 8, 6, 20084, 20085, 20086, "", 0    '8 : Druid
        
        GUI.OptionButton_Create winCharCreation, 85, 276, 8, 6, 20084, 20085, 20086, "", 1    '9 : Earth
        GUI.OptionButton_Create winCharCreation, 100, 276, 8, 6, 20084, 20085, 20086, "", 1   '10 : Air
        GUI.OptionButton_Create winCharCreation, 117, 276, 8, 6, 20084, 20085, 20086, "", 1   '11 : Fire
        GUI.OptionButton_Create winCharCreation, 133, 276, 8, 6, 20084, 20085, 20086, "", 1   '12 : Water
        
        GUI.OptionButton_Create winCharCreation, 85, 324, 8, 6, 20084, 20085, 20086, "", 2    '13 : Kinetic
        GUI.OptionButton_Create winCharCreation, 101, 324, 8, 6, 20084, 20085, 20086, "", 2   '14 : Telepathic
        GUI.OptionButton_Create winCharCreation, 116, 324, 8, 6, 20084, 20085, 20086, "", 2   '15 : Metabolic
        
        'CheckBoxes
        'The options Buttons are normally used, but are disabled when psionicist is selected
        'so that more than one can be selected
        GUI.CheckBox_Create winCharCreation, 85, 324, 8, 6, 20084, 20085, "", , True    '0 : Kinetic
        GUI.CheckBox_Create winCharCreation, 101, 324, 8, 6, 20084, 20085, "", , True   '1 : Telepathic
        GUI.CheckBox_Create winCharCreation, 116, 324, 8, 6, 20084, 20085, "", , True   '2 : Metabolic
        'Disable the comboboxes
        GUI.CheckBox_Set_Enabled winCharCreation, 0, False
        GUI.CheckBox_Set_Enabled winCharCreation, 1, False
        GUI.CheckBox_Set_Enabled winCharCreation, 2, False
        'Hide the comboboxes
        GUI.CheckBox_Set_Visible winCharCreation, 0, False
        GUI.CheckBox_Set_Visible winCharCreation, 1, False
        GUI.CheckBox_Set_Visible winCharCreation, 2, False
        
        'Labels
        GUI.Label_Create winCharCreation, 2, &HFFFFFFFF, "", 145, 64, 96, 14, fa_center    '0 : Race
        GUI.Label_Create winCharCreation, 2, &HFFFFFFFF, "", 161, 64, 96, 14, fa_center    '1 : Alignment
        GUI.Label_Create winCharCreation, 2, &HFFFFFFFF, "", 193, 64, 33, 14, fa_center    '2 : STR
        GUI.Label_Create winCharCreation, 2, &HFFFFFFFF, "", 209, 64, 33, 14, fa_center    '3 : DEX
        GUI.Label_Create winCharCreation, 2, &HFFFFFFFF, "", 225, 64, 33, 14, fa_center    '4 : CON
        GUI.Label_Create winCharCreation, 2, &HFFFFFFFF, "", 241, 64, 33, 14, fa_center    '5 : INT
        GUI.Label_Create winCharCreation, 2, &HFFFFFFFF, "", 257, 64, 33, 14, fa_center    '6 : WIT
        GUI.Label_Create winCharCreation, 2, &HFFFFFFFF, "", 273, 64, 33, 14, fa_center    '7 : CHR
        GUI.Label_Create winCharCreation, 2, &HFFFFFFFF, "", 289, 66, 33, 14, fa_center    '8 : Points
        
        'Textboxes
        GUI.Text_Create 0, winCharCreation, "", 128, 63, 100, 16, 2, &HFFFFFFFF     '0 : Name
    End If
End Sub

Public Sub Windows_Create_Login_Window()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 10/24/2003
'
'**************************************************************
    Dim grh_height As Long
    Dim grh_width As Long
    Dim TempLng As Long
    Dim TempStr As String
    
    Engine.Grh_Info_Get 20448, TempStr, TempLng, TempLng, grh_width, grh_height, TempLng
    
    '*************************
    'Login window
    '*************************
    If winLogin = -1 Then
        winLogin = GUI.Window_Create(20448, ViewHeight - grh_height - 20, ViewWidth / 2 - grh_width / 2, grh_width, grh_height)
        
        'Create the Text Boxes
        GUI.Text_Create 0, winLogin, "", 20, 142, 100, 16, 2, &HFFFFFFFF, &HFFC0C0C0
        GUI.Text_Create 0, winLogin, "", 44, 142, 100, 16, 2, &HFFFFFFFF, &HFFC0C0C0, , True
        
        'Create the CheckBox
        GUI.CheckBox_Create winLogin, 68, 67, 200, 20, 20431, 20044, "" '0 : Remember Password
        
        'Create the buttons
        GUI.Button_Create 20451, winLogin, "", 0, 88, 22, 87, 31, 20452     '0 : Cancel
        GUI.Button_Create 20449, winLogin, "", 0, 88, 162, 87, 31, 20450    '1 : Login
    End If
End Sub

Public Sub Windows_Create_Account_Window()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 10/24/2003
'
'**************************************************************
    Dim grh_height As Long
    Dim grh_width As Long
    Dim TempLng As Long
    Dim TempStr As String
    Dim LoopC As Long
    
    Engine.Grh_Info_Get 20454, TempStr, TempLng, TempLng, grh_width, grh_height, TempLng
    
    '*************************
    'Account window
    '*************************
    If winAccount = -1 Then
        winAccount = GUI.Window_Create(20454, ViewHeight / 2 - grh_height / 2, ViewWidth / 2 - grh_width / 2, grh_width, grh_height)
        
        'Create buttons
        GUI.Button_Create 0, winAccount, "", 0, 32, 32, 63, 63    '0 : Top-left pic
        GUI.Button_Create 0, winAccount, "", 0, 32, 128, 63, 63   '1 : Top-right pic
        GUI.Button_Create 0, winAccount, "", 0, 176, 32, 63, 63   '2 : Bottom-left pic
        GUI.Button_Create 0, winAccount, "", 0, 176, 128, 63, 63  '3 : Bottom-right pic
        
        'Buttons 4 and 5 change their Grh according buttons 0 - 3, and are initialized to 0
        GUI.Button_Create 0, winAccount, "", 0, 214, 218, 75, 20    '4 : Create Char / Modify Class
        GUI.Button_Create 0, winAccount, "", 0, 214, 305, 75, 20    '5 : Delete Char
        
        GUI.Button_Create 20451, winAccount, "", 0, 263, 210, 87, 31, 20452    '6 : Cancel
        
        GUI.Button_Create 20466, winAccount, "", 0, 263, 305, 87, 31, 20467, , False    '7 : Start Game
        
        'Create labels
        'Char info
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 38, 260, 120, 12, fa_center '0 : Name
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 54, 260, 120, 12, fa_center '1 : Race
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 70, 260, 120, 12, fa_center '2 : Class
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 86, 260, 120, 12, fa_center '3 : Level
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 102, 260, 120, 12, fa_center '4 : Align
        
        'Stats
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 134, 260, 30, 12, fa_center '5 : STR
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 150, 260, 30, 12, fa_center '6 : DEX
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 166, 260, 30, 12, fa_center '7 : CON
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 134, 328, 30, 12, fa_center '8 : INT
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 150, 328, 30, 12, fa_center '9 : WIS
        GUI.Label_Create winAccount, 2, &HFFFFFFFF, "", 166, 328, 30, 12, fa_center '10 : CHR
    End If
    
    'Clear labels
    For LoopC = 0 To GUI.Window_Get_Label_Count(winAccount)
        GUI.Label_Set_Text winAccount, LoopC, ""
    Next LoopC
    
    'Clear buttons
    GUI.Button_Set_Unpressed_Grh winAccount, 4, 0
    GUI.Button_Set_Unpressed_Grh winAccount, 5, 0
    
    'Disable the Accept button
    GUI.Button_Set_Enabled winAccount, 7, False
End Sub

Public Sub Windows_Create_Splashscreen_Window()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 10/24/2003
'
'**************************************************************
    Dim grh_width As Long
    Dim grh_height As Long
    Dim TempStr As String
    Dim TempLng As Long
    
    Engine.Grh_Info_Get 20100, TempStr, TempLng, TempLng, grh_width, grh_height, TempLng
    
    winSplashScreen = GUI.Window_Create(20100, ViewHeight / 2 - grh_height / 2, ViewWidth / 2 - grh_width / 2, grh_width, grh_height)
End Sub

Public Sub Windows_CharCreation_Enable_Psionic_Powers(Optional ByVal JustOne As Boolean = True)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 10/24/2003
'
'**************************************************************
    If winCharCreation = -1 Then Exit Sub
    
    If JustOne Then
        'Hide the checkboxes
        GUI.CheckBox_Set_Visible winCharCreation, 0, False
        GUI.CheckBox_Set_Visible winCharCreation, 1, False
        GUI.CheckBox_Set_Visible winCharCreation, 2, False
        'Show the optionbuttons
        GUI.OptionButton_Set_Visible winCharCreation, 13, True
        GUI.OptionButton_Set_Visible winCharCreation, 14, True
        GUI.OptionButton_Set_Visible winCharCreation, 15, True
    Else
        'Show the checkboxes (they are already set to true and disabled, so there´s nothing left to do)
        GUI.CheckBox_Set_Visible winCharCreation, 0, True
        GUI.CheckBox_Set_Visible winCharCreation, 1, True
        GUI.CheckBox_Set_Visible winCharCreation, 2, True
        'Hide the optionbuttons
        GUI.OptionButton_Set_Visible winCharCreation, 13, False
        GUI.OptionButton_Set_Visible winCharCreation, 14, False
        GUI.OptionButton_Set_Visible winCharCreation, 15, False
    End If
End Sub

Public Sub Windows_CharCreation_Arrange_Classes()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 10/24/2003
'
'**************************************************************
    If winCharCreation = -1 Then Exit Sub
    
    GUI.OptionButton_Group_Enable winCharCreation, 0
    
    Select Case GUI.Label_Get_Text(winCharCreation, 0)
        Case Is = "Half-Giant"
            GUI.OptionButton_Set_Enabled winCharCreation, 2, False
            GUI.OptionButton_Set_Enabled winCharCreation, 5, False
            GUI.OptionButton_Set_Enabled winCharCreation, 6, False
        Case Is = "Mul"
            GUI.OptionButton_Set_Enabled winCharCreation, 3, False
            GUI.OptionButton_Set_Enabled winCharCreation, 5, False
            GUI.OptionButton_Set_Enabled winCharCreation, 6, False
        Case Is = "Thri-Kreen"
            GUI.OptionButton_Set_Enabled winCharCreation, 2, False
            GUI.OptionButton_Set_Enabled winCharCreation, 5, False
            GUI.OptionButton_Set_Enabled winCharCreation, 6, False
        Case Is = "Elf"
            GUI.OptionButton_Set_Enabled winCharCreation, 8, False
        Case Is = "Halfling"
            GUI.OptionButton_Set_Enabled winCharCreation, 5, False
            GUI.OptionButton_Set_Enabled winCharCreation, 6, False
        Case Is = "Dwarf"
            GUI.OptionButton_Set_Enabled winCharCreation, 3, False
            GUI.OptionButton_Set_Enabled winCharCreation, 5, False
            GUI.OptionButton_Set_Enabled winCharCreation, 6, False
            GUI.OptionButton_Set_Enabled winCharCreation, 8, False
    End Select
    
    'Set gladiator as selected
    GUI.OptionButton_Set_Selected winCharCreation, 0
End Sub

Public Sub Windows_Create_NPC_Speech_Window()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2004
'
'**************************************************************
    Dim grh_height As Long
    Dim grh_width As Long
    Dim TempLng As Long
    Dim TempStr As String
    Dim stretching_factor As Single
    
    Engine.Grh_Info_Get 20001, TempStr, TempLng, TempLng, grh_width, grh_height, TempLng
    
    '*************************
    'NPC Speech windows
    '*************************
    If winNPCSpeech1 = -1 Then
        stretching_factor = ViewWidth / grh_width
        
        MAX_CHARS_PER_LINE_NPC1 = 75 * ViewWidth / 640
        MAX_CHARS_PER_LINE_NPC2 = 120 * ViewHeight / 640
        
        winNPCSpeech1 = GUI.Window_Create(20002, 0, 0, ViewWidth, grh_height * stretching_factor)
        winNPCSpeech2 = GUI.Window_Create(20001, ViewHeight - grh_height * stretching_factor, 0, ViewWidth, grh_height * stretching_factor)
        
        'Buttons
        GUI.Button_Create 0, winNPCSpeech1, "", 0, 25 * stretching_factor, 7 * stretching_factor, 64 * stretching_factor, 64 * stretching_factor                '0 : NPC portrait
        GUI.Button_Create 20435, winNPCSpeech1, "", 0, 25 * stretching_factor, 490 * stretching_factor, 15 * stretching_factor, 15 * stretching_factor, 20436    '1 : Scroll up
        GUI.Button_Create 20433, winNPCSpeech1, "", 0, 80 * stretching_factor, 490 * stretching_factor, 15 * stretching_factor, 15 * stretching_factor, 20434   '2 : Scroll down
        
        GUI.Button_Create 20435, winNPCSpeech2, "", 0, 25 * stretching_factor, 490 * stretching_factor, 15 * stretching_factor, 15 * stretching_factor, 20436   '0 : Scroll up
        GUI.Button_Create 20433, winNPCSpeech2, "", 0, 80 * stretching_factor, 490 * stretching_factor, 15 * stretching_factor, 15 * stretching_factor, 20434   '1 : Scroll down
        
        'Labels
        GUI.Label_Create winNPCSpeech1, 4, &HFFFFFF00, "", 20 * stretching_factor, 85 * stretching_factor, 200, 16, fa_topleft           '0 : NPC name
        GUI.Label_Create winNPCSpeech1, 3, &HFFFFFF00, "", 45 * stretching_factor, 85 * stretching_factor, ViewWidth - 110 * stretching_factor, 16, fa_topleft  '1 : greet line 1
        GUI.Label_Create winNPCSpeech1, 3, &HFFFFFF00, "", 60 * stretching_factor, 85 * stretching_factor, ViewWidth - 110 * stretching_factor, 16, fa_topleft  '2 : greet line 2
        GUI.Label_Create winNPCSpeech1, 3, &HFFFFFF00, "", 75 * stretching_factor, 85 * stretching_factor, ViewWidth - 110 * stretching_factor, 16, fa_topleft  '3 : greet line 3
        
        GUI.Label_Create winNPCSpeech2, 3, &HFFFFFF00, "", 20 * stretching_factor, 25 * stretching_factor, ViewWidth - 125 * stretching_factor, 16, fa_topleft '0 : response line 1
        GUI.Label_Create winNPCSpeech2, 3, &HFFFFFF00, "", 35 * stretching_factor, 25 * stretching_factor, ViewWidth - 125 * stretching_factor, 16, fa_topleft '1 : response line 2
        GUI.Label_Create winNPCSpeech2, 3, &HFFFFFF00, "", 50 * stretching_factor, 25 * stretching_factor, ViewWidth - 125 * stretching_factor, 16, fa_topleft '2 : response line 3
        GUI.Label_Create winNPCSpeech2, 3, &HFFFFFF00, "", 65 * stretching_factor, 25 * stretching_factor, ViewWidth - 125 * stretching_factor, 16, fa_topleft '3 : response line 4
    End If
End Sub
