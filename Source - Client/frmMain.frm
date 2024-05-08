VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORE Client"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MoviePic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox LoadingPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   2400
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   0
      Top             =   2760
      Width           =   6855
      Begin MSComctlLib.ProgressBar ResourceFilePrgBar 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   2280
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar GeneralPrgBar 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label GeneralLbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   " Loading File 1 of ###..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   1710
      End
      Begin VB.Label ResourceFileLbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   " Loading Graphics..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   1425
      End
   End
   Begin VB.Timer tmrSplash 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin OREClient.ctlDirectPlayClient dp_client 
      Left            =   120
      Top             =   600
      _ExtentX        =   873
      _ExtentY        =   873
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dp_client_CharAuthenticated()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    'Start game (we hide the Login, Account cnd CharCreation windows)
    GUI.Window_Show winLogin, False
    GUI.Window_Show winAccount, False
    GUI.Window_Show winCharCreation, False
    Run = False
End Sub

Private Sub dp_client_ClientAuthenticated(ByVal first_char_name As String, ByVal second_char_name As String, ByVal third_char_name As String, ByVal fourth_char_name As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    'Store char names
    chars(1).name = first_char_name
    chars(2).name = second_char_name
    chars(3).name = third_char_name
    chars(4).name = fourth_char_name
    
    'Request data on all existing chars to be displayed in the account window
    Dim LoopC As Long
    For LoopC = 1 To 4
        If chars(LoopC).name <> "" Then
            dp_client.Client_Request_Char_Stats chars(LoopC).name
        End If
    Next LoopC
    
    GUI.Window_Show winLogin, False
    GUI.Window_Show winAccount, True
    
    attempting_to_connect = False
End Sub

Private Sub dp_client_ClientConnected()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'Try to login using the given account info
'**************************************************************
    dp_client.Client_Authenticate GUI.Text_Get_Text(winLogin, 1)
    attempting_to_connect = False
    connected = True
End Sub

Private Sub dp_client_ClientConnectionFailed()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    attempting_to_connect = False
    'Send Message
    MsgBox "Error connecting to server."
    'Set Status to Disconnect
    Call dp_client_ClientDisconnected
End Sub

Private Sub dp_client_ClientDisconnected()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/25/2004
'
'**************************************************************
    'Connection ended, go to the login window
    connected = False
    
    dp_client.Client_Initialize Engine, "{12345678-1234-1234-1234-123456789ABC}"
    
    'Hide unused windows and show those needed
    If winAccount > -1 Then
        GUI.Window_Show winAccount, False
    End If
    If winNPCSpeech1 > -1 Then
        GUI.Window_Show winNPCSpeech1, False
        GUI.Window_Show winNPCSpeech2, False
    End If
    Game_Login
End Sub

Private Sub dp_client_EngineStart()
'We do nothing here, since the Engine was already started
End Sub

Private Sub dp_client_EngineStop()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    'We only get disconnected when shutting down
    connected = False
End Sub

Private Sub dp_client_GumpLoaded(ByVal cstring As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/07/2004
'
'**************************************************************
    Dim s As String
    Dim i As Integer
    Dim nomove As Boolean
    Dim noclose As Boolean
    Dim noresize As Boolean
    Dim str(1 To 10) As String
    
    'Used for getting grh's data
    Dim grh_width As Long
    Dim grh_height As Long
    Dim temp_string As String
    Dim temp As Long
    
    i = 1
    gumpcount = gumpcount + 1
    Do
        'Get string
        s = General_Field_Read_GUMP(i, cstring, Asc("}"))
        If s = "" Then Exit Do
        'clean of {
        s = Mid(s, 2)
        
        If s = " nomove " Then nomove = True
        If s = " noclose " Then noclose = True
        If s = " noresize " Then noresize = True
        
        'Window background
        If Mid(s, 1, 9) = "resizepic" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'height
            str(4) = General_Field_Read_GUMP(5, s, Asc("/")) 'width
            str(5) = General_Field_Read_GUMP(6, s, Asc("/")) 'grh id
            GUMP(gumpcount) = GUI.Window_Create(CLng(str(5)), CLng(str(1)), CLng(str(2)), CLng(str(3)), CLng(str(4)), Not nomove, Not noresize, Not noclose)
            GUI.Window_Show GUMP(gumpcount), True
        End If
        
        'Alpharegion
        If Mid(s, 1, 12) = "checkertrans" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'height
            str(4) = General_Field_Read_GUMP(5, s, Asc("/")) 'width
            
            GUI.Window_Change_Alphablending GUMP(gumpcount), CLng(str(1)), CLng(str(2)), CLng(str(4)), CLng(str(3))
        End If
        
        'button
        If Mid(s, 1, 6) = "button" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'grh id1
            str(4) = General_Field_Read_GUMP(5, s, Asc("/")) 'grh id2
            str(5) = General_Field_Read_GUMP(5, s, Asc("/")) 'button id
            str(6) = General_Field_Read_GUMP(5, s, Asc("/")) 'button type
            
            Engine.Grh_Info_Get CLng(str(3)), temp_string, temp, temp, grh_width, grh_height, temp
            
            GUI.Button_Create CLng(str(3)), GUMP(gumpcount), "", &H0, CLng(str(2)), CLng(str(1)), grh_width, grh_height, CLng(str(4)), , , CLng(str(5)), CLng(str(6))
        End If
        
        'checkbox
        If Mid(s, 1, 8) = "checkbox" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'id1
            str(4) = General_Field_Read_GUMP(5, s, Asc("/")) 'id2
            str(5) = General_Field_Read_GUMP(5, s, Asc("/")) 'init
            str(6) = General_Field_Read_GUMP(5, s, Asc("/")) 'switch id
            
            Engine.Grh_Info_Get CLng(str(3)), temp_string, temp, temp, grh_width, grh_height, temp
            
            GUI.CheckBox_Create GUMP(gumpcount), CLng(str(1)), CLng(str(2)), grh_width, grh_height, CLng(str(3)), CLng(str(4)), "", , CBool(str(5)), , , CLng(str(6))
        End If
        
        'image
        If Mid(s, 1, 7) = "gumppic" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'gump_id
            
            Engine.Grh_Info_Get CLng(str(3)), temp_string, temp, temp, grh_width, grh_height, temp
            
            GUI.Image_Create GUMP(gumpcount), CLng(str(1)), CLng(str(2)), CLng(str(3)), grh_height, grh_width
        End If
        
        'picture tiled
        If Mid(s, 1, 12) = "gumppictiled" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'height
            str(4) = General_Field_Read_GUMP(5, s, Asc("/")) 'width
            str(5) = General_Field_Read_GUMP(6, s, Asc("/")) 'grh id
            
            GUI.Image_Create GUMP(gumpcount), CLng(str(1)), CLng(str(2)), CLng(str(5)), CLng(str(4)), CLng(str(3))
        End If
        
'TODO: Get this done!!!!
        'item
        If Mid(s, 1, 7) = "tilepic" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'item_id
            
            
        End If
        
        'label
        If Mid(s, 1, 4) = "text" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'hue
            str(4) = General_Field_Read_GUMP(5, s, Asc("/")) 'text
            
            'It's as big as the screen so it can allways fit in
            GUI.Label_Create GUMP(gumpcount), 1, CLng(str(3)), str(4), CLng(str(2)), CLng(str(1)), ViewWidth, ViewHeight, fa_topleft
        End If
        
        'labelC
        If Mid(s, 1, 11) = "textcropped" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'hue
            str(4) = General_Field_Read_GUMP(5, s, Asc("/")) 'height
            str(5) = General_Field_Read_GUMP(5, s, Asc("/")) 'width
            str(6) = General_Field_Read_GUMP(5, s, Asc("/")) 'text
            
            GUI.Label_Create GUMP(gumpcount), 1, CLng(str(3)), str(6), CLng(str(2)), CLng(str(1)), CLng(str(5)), CLng(str(4)), fa_topleft
        End If
        
        'radio
        If Mid(s, 1, 5) = "radio" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'id1
            str(4) = General_Field_Read_GUMP(5, s, Asc("/")) 'id2
            str(5) = General_Field_Read_GUMP(5, s, Asc("/")) 'init
            str(6) = General_Field_Read_GUMP(5, s, Asc("/")) 'switch id
            
            Engine.Grh_Info_Get CLng(str(3)), temp_string, temp, temp, grh_width, grh_height, temp
            
'TODO: Add group index!!!
            GUI.OptionButton_Create GUMP(gumpcount), CLng(str(2)), CLng(str(1)), grh_width, grh_height, CLng(str(4)), CLng(str(3)), CLng(str(3)), "", , CLng(str(6))
        End If
        
        'text
        If Mid(s, 1, 9) = "textentry" Then
            str(1) = General_Field_Read_GUMP(2, s, Asc("/")) 'x
            str(2) = General_Field_Read_GUMP(3, s, Asc("/")) 'y
            str(3) = General_Field_Read_GUMP(4, s, Asc("/")) 'hue
            str(4) = General_Field_Read_GUMP(5, s, Asc("/")) 'height
            str(5) = General_Field_Read_GUMP(5, s, Asc("/")) 'width
            str(6) = General_Field_Read_GUMP(5, s, Asc("/")) 'text
            
            GUI.Text_Create 0, GUMP(gumpcount), str(6), CLng(str(2)), CLng(str(1)), CLng(str(5)), CLng(str(4)), 1, CLng(str(3))
        End If
    
        i = i + 1
    Loop
End Sub

Private Sub dp_client_InventorySlotChanged(ByVal slot As Long, ByVal Item_Index As Long, ByVal amount As Long, ByVal equiped As Boolean)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/20/2004
'
'**************************************************************
    With User_Inventory(slot)
        .item_data_index = Item_Index
        .amount = amount
        .equiped = equiped
    End With
End Sub

Private Sub dp_client_NPCChatReceive(ByVal NPC_greet As String, ByRef Responses() As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/27/2004
'
'**************************************************************
    Dim LoopC As Long
    
    'Store it in memory
    ReDim Current_Speech.Responses(1 To UBound(Responses))
    
    Chats.Dialog_Format NPC_greet, MAX_CHARS_PER_LINE_NPC1, Current_Speech.NPC_greet()
    
    For LoopC = 1 To UBound(Responses)
        Chats.Dialog_Format CStr(LoopC) & ") " & Responses(LoopC), MAX_CHARS_PER_LINE_NPC2, Current_Speech.Responses(LoopC).text_line()
    Next LoopC
    
    'Reset scroll offset
    NPCSpeechOffset1 = 0
    NPCSpeechOffset2 = 0
    
    NPC_Speech_Render
    
    GUI.Window_Show winNPCSpeech1, True
    GUI.Window_Show winNPCSpeech2, True
    
    player_talking_to_NPC = True
End Sub

Private Sub dp_client_NPCChatStart(ByVal NPC_greet As String, ByRef Responses() As String, ByVal NPC_name As String, ByVal NPC_portrait As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/26/2004
'
'**************************************************************
    Current_Speech.NPC_name = NPC_name
    Current_Speech.NPC_portrait_pic = NPC_portrait
    
    'Display chat
    Windows_Create_NPC_Speech_Window
    
    'Display name and pic
    GUI.Label_Set_Text winNPCSpeech1, 0, NPC_name
    GUI.Button_Set_Unpressed_Grh winNPCSpeech1, 0, NPC_portrait
    GUI.Button_Set_Pressed_Grh winNPCSpeech1, 0, NPC_portrait
    
    'Set the rest of the data
    dp_client_NPCChatReceive NPC_greet, Responses()
End Sub

Private Sub dp_client_NPCChatTerminated()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    GUI.Window_Show winNPCSpeech1, False
    GUI.Window_Show winNPCSpeech2, False
    player_talking_to_NPC = False
    
    'Allow normal mouse unput again
    GUI.Window_Set_Active -1
End Sub

Private Sub dp_client_ReceiveCharStats(ByVal char_name As String, ByVal race As races, ByVal Class As classes, ByVal Alignment As Alignment, ByVal sphere As spheres, ByVal psionic_power As psionic_powers, ByVal level As Long, ByVal char_STR As Long, ByVal char_DEX As Long, ByVal char_CON As Long, ByVal char_INT As Long, ByVal char_WIS As Long, ByVal char_CHR As Long, ByVal portrait As Long, ByVal char_data_index As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    Dim LoopC As Long
    
    LoopC = 1
    
    'Look which char index corresponds to it
    Do Until chars(LoopC).name = char_name Or LoopC > 4
        LoopC = LoopC + 1
    Loop
    
    If LoopC > 4 Then Exit Sub
    
    'Store stats to be displayed when needed
    chars(LoopC).race = race
    chars(LoopC).Class = Class
    chars(LoopC).align = Alignment
    chars(LoopC).sphere = sphere
    chars(LoopC).psionic_power = psionic_power
    chars(LoopC).level = level
    chars(LoopC).char_STR = char_STR
    chars(LoopC).char_DEX = char_DEX
    chars(LoopC).char_CON = char_CON
    chars(LoopC).char_INT = char_INT
    chars(LoopC).char_WIS = char_WIS
    chars(LoopC).char_CHR = char_CHR
    chars(LoopC).portrait = portrait
    chars(LoopC).char_data_index = char_data_index
    
    'Display portriat in the corresponding button
    GUI.Button_Set_Unpressed_Grh winAccount, LoopC - 1, portrait
    GUI.Button_Set_Pressed_Grh winAccount, LoopC - 1, portrait
End Sub

Private Sub dp_client_ReceiveChatCritical(ByVal Chat_Text As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    'Display message
    MsgBox Chat_Text
End Sub

Private Sub dp_client_ReceiveChatText(ByVal Chat_Text As String, ByVal Chat As chat_type, ByVal sender_id As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    'Check which kind of chat we received and display it accordingly.
    Select Case Chat
        Case CT_Private
            'Display the text in the topleft corner
            Chats.Chat_Add Chat_Text, 1, &HFF00FFFF
        
        Case CT_Normal
            'Assign it to a char and display it over it´s head
            Engine.Char_Chat_Set Engine.Char_Find(sender_id), Chat_Text, &HFFFFFFFF, 1
            Chats.Chat_Add Chat_Text, 1, &HFF00FFFF
        
        Case CT_Map
            'The char shouted. Show it above the char's head and in the topleft corner
            Chats.Chat_Add "You here someone shouting: " & Chat_Text, 1, &HFFFF00FF
            Engine.Char_Chat_Set Engine.Char_Find(sender_id), Chat_Text, &HFFFF00FF, 1
        
        Case CT_Global
            'Important message sent by the server or a GM. Display it on the top-left corner.
            Chats.Chat_Add Chat_Text, 1, &HFFFFFF00, True
    End Select
End Sub

Private Sub dp_client_StatsRolled(ByVal points As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    Dim Dice1 As Long
    Dim Dice2 As Long
    Dim Dice3 As Long
    Dim char_STR As Long
    Dim char_DEX As Long
    Dim char_CON As Long
    Dim char_INT As Long
    Dim char_WIS As Long
    Dim char_CHR As Long
    Dim temp_points As Long
    
    temp_points = points / 6
RollDice1:
    Dice1 = Round(General_Random_Number(1, 6), 0)
    
    'Check value
    If (temp_points - Dice1) / 2 > 6 Or temp_points - Dice1 < 0 Then
        GoTo RollDice1
    End If
    
RollDice2:
    Dice2 = Round(General_Random_Number(1, 6), 0)
    
    'Check value
    If temp_points - Dice1 - Dice2 > 6 Or temp_points - Dice1 - Dice2 < 0 Then
        GoTo RollDice2
    End If
    
    Dice3 = temp_points - Dice1 - Dice2
    
'TODO: Display the 3 graphics
    
    
    points = points - ChrCreation.Char_STR_Get_Min(ChrCreation.Race_Get) - ChrCreation.Char_DEX_Get_Min(ChrCreation.Race_Get) - ChrCreation.Char_CON_Get_Min(ChrCreation.Race_Get) - ChrCreation.Char_INT_Get_Min(ChrCreation.Race_Get) - ChrCreation.Char_WIS_Get_Min(ChrCreation.Race_Get) - ChrCreation.Char_CHR_Get_Min(ChrCreation.Race_Get)
    
    'Divide the points randomly between the 6 stats
RollSTR:
    char_STR = Round(General_Random_Number(char_STR, ChrCreation.Char_STR_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_STR_Get_Min(ChrCreation.Race_Get)), 0)
    
    'Check it´s a valid value
    If (points - char_STR) > (ChrCreation.Char_CHR_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_CHR_Get_Min(ChrCreation.Race_Get) + ChrCreation.Char_CON_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_CON_Get_Min(ChrCreation.Race_Get) + _
      ChrCreation.Char_DEX_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_DEX_Get_Min(ChrCreation.Race_Get) + ChrCreation.Char_INT_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_INT_Get_Min(ChrCreation.Race_Get) + ChrCreation.Char_WIS_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_WIS_Get_Min(ChrCreation.Race_Get)) Then
        GoTo RollSTR:
    End If
    
    points = points - char_STR
    
RollDEX:
    If points > ChrCreation.Char_DEX_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_DEX_Get_Min(ChrCreation.Race_Get) Then
        char_DEX = Round(General_Random_Number(char_DEX, ChrCreation.Char_DEX_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_DEX_Get_Min(ChrCreation.Race_Get)), 0)
    Else
        char_DEX = Round(General_Random_Number(0, points), 0)
    End If
    
    'Check it´s a valid value
    If (points - char_DEX) > (ChrCreation.Char_CHR_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_CHR_Get_Min(ChrCreation.Race_Get) + ChrCreation.Char_CON_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_CON_Get_Min(ChrCreation.Race_Get) + _
      ChrCreation.Char_INT_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_INT_Get_Min(ChrCreation.Race_Get) + ChrCreation.Char_WIS_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_WIS_Get_Min(ChrCreation.Race_Get)) Then
        GoTo RollDEX:
    End If
    
    points = points - char_DEX
    
RollCON:
    If points > (ChrCreation.Char_CON_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_CON_Get_Min(ChrCreation.Race_Get)) Then
        char_CON = Round(General_Random_Number(char_CON, ChrCreation.Char_CON_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_CON_Get_Min(ChrCreation.Race_Get)), 0)
    Else
        char_CON = Round(General_Random_Number(0, points), 0)
    End If
    
    'Check it´s a valid value
    If (points - char_CON) > (ChrCreation.Char_CHR_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_CHR_Get_Min(ChrCreation.Race_Get) + _
      ChrCreation.Char_INT_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_INT_Get_Min(ChrCreation.Race_Get) + ChrCreation.Char_WIS_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_WIS_Get_Min(ChrCreation.Race_Get)) Then
        GoTo RollCON:
    End If
    
    points = points - char_CON
    
RollINT:
    If points > ChrCreation.Char_INT_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_INT_Get_Min(ChrCreation.Race_Get) Then
        char_INT = Round(General_Random_Number(char_INT, ChrCreation.Char_INT_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_INT_Get_Min(ChrCreation.Race_Get)), 0)
    Else
        char_INT = Round(General_Random_Number(0, points), 0)
    End If
    
    'Check it´s a valid value
    If (points - char_INT) > (ChrCreation.Char_CHR_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_CHR_Get_Min(ChrCreation.Race_Get) + _
      ChrCreation.Char_WIS_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_WIS_Get_Min(ChrCreation.Race_Get)) Then
        GoTo RollINT
    End If
    
    points = points - char_INT
    
RollWIS:
    If points > ChrCreation.Char_WIS_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_WIS_Get_Min(ChrCreation.Race_Get) Then
        char_WIS = Round(General_Random_Number(char_WIS, ChrCreation.Char_WIS_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_WIS_Get_Min(ChrCreation.Race_Get)), 0)
    Else
        char_WIS = Round(General_Random_Number(0, points), 0)
    End If
    
    'Check it´s a valid value
    If (points - char_WIS) > (ChrCreation.Char_CHR_Get_Max(ChrCreation.Race_Get) - ChrCreation.Char_CHR_Get_Min(ChrCreation.Race_Get)) Then
        GoTo RollWIS:
    End If
    
    points = points - char_WIS
    
    'Finish setting up char stats
    Call ChrCreation.Char_STR_Set(ChrCreation.Char_STR_Get_Min(ChrCreation.Race_Get) + char_STR)
    Call ChrCreation.Char_DEX_Set(ChrCreation.Char_DEX_Get_Min(ChrCreation.Race_Get) + char_DEX)
    Call ChrCreation.Char_CON_Set(ChrCreation.Char_CON_Get_Min(ChrCreation.Race_Get) + char_CON)
    Call ChrCreation.Char_INT_Set(ChrCreation.Char_INT_Get_Min(ChrCreation.Race_Get) + char_INT)
    Call ChrCreation.Char_WIS_Set(ChrCreation.Char_WIS_Get_Min(ChrCreation.Race_Get) + char_WIS)
    Call ChrCreation.Char_CHR_Set(ChrCreation.Char_CHR_Get_Min(ChrCreation.Race_Get) + points)
    
    GUI.Label_Set_Text winCharCreation, 2, CStr(ChrCreation.Char_STR_Get)
    GUI.Label_Set_Text winCharCreation, 3, CStr(ChrCreation.Char_DEX_Get)
    GUI.Label_Set_Text winCharCreation, 4, CStr(ChrCreation.Char_CON_Get)
    GUI.Label_Set_Text winCharCreation, 5, CStr(ChrCreation.Char_INT_Get)
    GUI.Label_Set_Text winCharCreation, 6, CStr(ChrCreation.Char_WIS_Get)
    GUI.Label_Set_Text winCharCreation, 7, CStr(ChrCreation.Char_CHR_Get)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2004
'
'**************************************************************
    'We are filling a text box
    If GUI.Window_Get_Visible(GUI.Text_Active_Get_Window) Then
        If GUI.Text_Get_Active <> -1 Then
            If KeyAscii = 8 Then
                GUI.Text_Remove_Char GUI.Text_Active_Get_Window, GUI.Text_Get_Active, 1
            ElseIf KeyAscii >= 32 And KeyAscii <= 126 Then
                GUI.Text_Add_Text GUI.Text_Active_Get_Window, GUI.Text_Get_Active, Chr(KeyAscii)
            ElseIf KeyAscii = vbKeyTab Then
                'Move focus to the next text box in window
                GUI.Text_Move_Focus
            End If
            Exit Sub
        End If
    End If
    
    'We are writing a chat
    If Writing_Chat Then
        If KeyAscii = 8 Then
            Chat_Text = left$(Chat_Text, Len(Chat_Text) - 1)
        ElseIf KeyAscii >= 32 And KeyAscii <= 126 Then
            Chat_Text = Chat_Text & Chr(KeyAscii)
        ElseIf KeyAscii = vbKeyReturn Then
            'Check if we are sending a Sysop command
            If UCase(left$(Chat_Text, 10)) = "/WORLDSAVE" Then
                dp_client.Client_Sysop_WorldSave
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 6)) = "/RESET" Then
                dp_client.Client_Sysop_Reset
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 9)) = "/SHUTDOWN" Then
                dp_client.Client_Sysop_Shutdown
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 11)) = "/PLAYERLIST" Then
                dp_client.Client_Sysop_PlayerList
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 5)) = "/GOTO" Then
                Chat_Text = Right$(Chat_Text, Len(Chat_Text) - 6)
                dp_client.Client_Sysop_GoTo General_Field_Read(1, Chat_Text, Asc(" ")), _
                                            CLng(General_Field_Read(2, Chat_Text, Asc(" "))), _
                                            CLng(General_Field_Read(3, Chat_Text, Asc(" ")))
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 7)) = "/SUMMON" Then
                dp_client.Client_Sysop_Summon Right$(Chat_Text, Len(Chat_Text) - 8)
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 7)) = "/FREEZE" Then
                dp_client.Client_Sysop_Freeze
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 5)) = "/JAIL" Then
                dp_client.Client_Sysop_Jail
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 4)) = "/BAN" Then
                Chat_Text = Right$(Chat_Text, Len(Chat_Text) - 5)
                dp_client.Client_Sysop_Ban (UCase(General_Field_Read(1, Chat_Text, Asc(" "))) = "Y"), General_Field_Read(2, Chat_Text, Asc(" "))
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 9)) = "/WORLDMSG" Then
                dp_client.Client_Sysop_WorldMessage Right$(Chat_Text, Len(Chat_Text) - 9)
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 5)) = "/HIDE" Then
                dp_client.Client_Sysop_Hide
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 6)) = "/SHIDE" Then
                dp_client.Client_Sysop_Shide
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 8)) = "/ACCOUNT" Then
                dp_client.Client_Sysop_Account
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 6)) = "/QUEST" Then
                dp_client.Client_Sysop_Quest
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            If UCase(left$(Chat_Text, 8)) = "/NEWITEM" Then
                dp_client.Client_Sysop_NewItem CLng(General_Field_Read(1, Right$(Chat_Text, Len(Chat_Text) - 9), Asc(" ")))
                Chat_Text = ""
                Writing_Chat = False
                Exit Sub
            End If
            
            'Send message and clear it from memory
            dp_client.Chat_Send Chat_Text
            Chat_Text = ""
            Writing_Chat = False
            GUI.Window_Show GUMP(1), False
            Exit Sub
        End If
    Else
        'We start writing a chat
        If KeyAscii = vbKeyReturn Then Writing_Chat = True
        GUI.Window_Show GUMP(1), True
    End If
End Sub

Private Sub Form_Load()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    'Launch Auto-Patcher
    'Shell App.Path & "\Autopatcher.exe", vbNormalFocus
    
    'Resize according to the resolution setted
    Me.height = Val(General_Var_Get(App.Path & "\Client.ini", "GRAPHICS", "Height")) * 15   'We multiply by 15 to convert to twips
    Me.width = Val(General_Var_Get(App.Path & "\Client.ini", "GRAPHICS", "Width")) * 15
    
    'Center the pic
    LoadingPic.top = Me.ScaleHeight / 2 - LoadingPic.ScaleHeight / 2
    LoadingPic.left = Me.ScaleWidth / 2 - LoadingPic.ScaleWidth / 2
    
    'Show the form
    Me.show
    DoEvents
    
    'Set the resource path
    resource_path = General_Var_Get(App.Path & "\client.ini", "GENERAL", "resource_path")
    
    'Load resources
    Loading_Resource_Files
    
    'Set all window handlers to -1 (they weren´r created)
    winSplashScreen = -1
    winLogin = -1
    winAccount = -1
    winCharCreation = -1
    winNPCSpeech1 = -1
    winNPCSpeech2 = -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    Run = False
    game_running = False
    Cancel = 1
End Sub

Private Sub tmrSplash_Timer()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    If Inc Then
        SplashScreenAlphaBlend = SplashScreenAlphaBlend + 1
    Else
        SplashScreenAlphaBlend = SplashScreenAlphaBlend - 1
    End If
    
    If SplashScreenAlphaBlend = 255 Then
        Inc = False
    End If
    If SplashScreenAlphaBlend = 0 Then
        Inc = True
    End If
End Sub
