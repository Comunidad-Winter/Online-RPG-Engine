Attribute VB_Name = "modCommunicationCommands"
'*****************************************************************
'modCommunicationCommands.bas - ORE Communication Command Constants - v0.5.0
'
'Specifies the communication protocol between the server and
'client.
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
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 8/20/2004
'   - Add: s_Packet_Inventory_Item packet
'   - Add: s_Char_Set_Inventory_Slot command
'   - Add: Generic_Packet_Map_Pos packet (other packets now call this one)
'   - Add: c_Packet_NPC_Respond packet
'   - Add: s_Packet_NPC_Speech packet
'   - Add: c_Action_NPC_Chat command
'   - Add: s_NPC header
'   - Add: s_NPC_Chat command
'   - Add: s_Packet_Item_Create packet
'   - Add: c_Authenticate_Char command
'   - Add: s_Char_Authenticated command
'   - Add: c_Request_Roll_Stats command
'   - Add: c_Request_Char_Stats command
'   - Add: s_Char_Stats_Rolled command
'   - Add: s_Packet_Stats_Rolled packet
'   - Add: s_Char_Stats_Get command
'   - Add: s_Packet_Character_Stats packet
'   - Add: s_Packet_Character_Names packet
'   - Add: c_Account header
'   - Add: c_Account_Add_Char command
'   - Add: c_Account_Remove_Char command
'
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************

Option Explicit

'***************************
'Constants and enumerations
'***************************
Public Const P_DELIMITER_CODE As Byte = 31
Public P_DELIMITER As String

'***************************
'Player stats declare
'***************************
Public Const c_Player_Item_Slots As Long = 99

'***************
'Server commands
'***************
'Header
Public Enum ServerPacketHeader
    s_Player = 1 'Start at 1
    s_Chat
    s_map
    s_Char
    s_NPC
    S_Gump
End Enum

'Command
Public Enum ServerPacketCommand
    s_Player_Authenticated = 1
    s_Player_Engine_Start
    s_Player_Engine_Stop
    s_Chat_Text
    s_Chat_Critical
    s_Chat_Global
    s_Chat_Whisper
    s_Chat_Normal
    s_Chat_Map
    s_Map_Load
    s_Map_Item_Add
    s_Map_Item_Remove
    s_Char_Authenticated
    s_Char_ID_Set
    s_Char_Create
    s_Char_Label_Set
    s_Char_Data_Set
    s_Char_Data_Body_Set
    s_Char_Pos_Set
    s_Char_Heading_Set
    s_Char_Move
    s_Char_Remove
    s_Char_Stats_Get
    s_Char_Stats_Rolled
    s_Char_Target
    s_Char_Recive_EXP
    s_Char_Hit
    s_Char_Set_Inventory_Slot
    s_NPC_Chat
    S_Gump_page
    s_Gump_Button
End Enum

'***************
'Client commands
'***************
'Header
Public Enum ClientPacketHeader
    c_Authenticate = 10001 'Start at 10001
    c_Chat
    c_Request
    c_Move
    c_Action
    c_Logoff
    c_Account
    c_Gump
    c_Sysop
End Enum

'Command
'Edited by Juan Martín Sotuyo Dodero
Public Enum ClientPacketCommand
    c_Authenticate_Login = 10001
    c_Authenticate_New
    c_Authenticate_Char
    c_Chat_Global
    c_Chat_Map
    c_Chat_Normal
    c_Chat_Whisper
    c_Request_Char_Stats
    c_Request_Item_List
    c_Request_Pos_Update
    c_Request_Roll_Stats
    c_Move_Moved
    c_Action_Fight_Escape
    c_Action_Fight_Item
    c_Action_Fight_Spell
    c_Action_Fight_Start
    c_Action_Fight_Weapon
    c_Action_Fight_Die
    c_Action_Item_Drop
    c_Action_Item_Move
    c_Action_Item_Pickup
    c_Action_Item_Equip
    c_Action_Attack
    c_Action_NPC_Chat
    c_Logoff_Account
    c_Logoff_Char
    c_Logoff_Session_Terminate
    c_Account_Add_Char
    c_Account_Remove_Char
    c_Gump_Page
    c_Gump_button
    c_Sysop_Saveworld
    c_Sysop_Reset
    c_Sysop_Shutdown
    c_Sysop_Playerlist
    c_Sysop_Goto
    c_Sysop_Summon
    c_Sysop_Freeze
    c_Sysop_Jail
    c_Sysop_Ban
    c_Sysop_Hide
    c_Sysop_sHide
    c_Sysop_Account
    c_Sysop_Quest
    c_Sysop_NewItem
End Enum

Public Function s_Packet_Char_Create(ByVal char_id As Long, ByVal map_x As Long, ByVal map_y As Long, ByVal Heading As Long, _
                                    ByVal char_data_index As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Create = CStr(char_id) _
                    & P_DELIMITER & Generic_Packet_Map_Pos(map_x, map_y) _
                    & P_DELIMITER & CStr(Heading) _
                    & P_DELIMITER & CStr(char_data_index)
End Function

Public Function s_Packet_Char_Label_Set(ByVal char_id As Long, ByVal label As String, ByVal font_index As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Label_Set = CStr(char_id) _
                        & P_DELIMITER & label _
                        & P_DELIMITER & CStr(font_index)
End Function

Public Function s_Packet_Char_Heading_Set(ByVal char_id As Long, ByVal Heading As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Heading_Set = CStr(char_id) _
                          & P_DELIMITER & CStr(Heading)
End Function

Public Function s_Packet_Char_Move(ByVal char_id As Long, ByVal Heading As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Move = CStr(char_id) _
                    & P_DELIMITER & CStr(Heading)
End Function

Public Function s_Packet_Char_Pos_Set(ByVal char_id As Long, ByVal map_x As Long, ByVal map_y As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Pos_Set = CStr(char_id) _
                    & P_DELIMITER & Generic_Packet_Map_Pos(map_x, map_y)
End Function

Public Function s_Packet_Char_Data_Body_Set(ByVal char_id As Long, ByVal body_index As Long, ByVal noloop As Boolean) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Data_Body_Set = CStr(char_id) _
                            & P_DELIMITER & CStr(body_index) _
                            & P_DELIMITER & CStr(CByte(noloop))
End Function

Public Function s_Packet_Char_Data_Set(ByVal char_id As Long, ByVal char_data_index As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Data_Set = CStr(char_id) _
                        & P_DELIMITER & CStr(char_data_index)
End Function

Public Function c_Packet_Player_New(ByVal password As String, ByVal profile_name As String, ByVal first_name As String, ByVal last_name As String, ByVal email As String) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/7/2004
'Take parameters and creates string for inserting into a packet
'**************************************************************
c_Packet_Player_New = password _
                        & P_DELIMITER & profile_name _
                        & P_DELIMITER & first_name _
                        & P_DELIMITER & last_name _
                        & P_DELIMITER & email
End Function

Public Function s_Packet_Stats_Rolled(ByVal aSTR As Long, ByVal aDEX As Long, ByVal aCON As Long, ByVal aINT As Long, ByVal aWIS As Long, ByVal aCHR As Long) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Stats_Rolled = CStr(aSTR) _
                        & P_DELIMITER & CStr(aDEX) _
                        & P_DELIMITER & CStr(aCON) _
                        & P_DELIMITER & CStr(aINT) _
                        & P_DELIMITER & CStr(aWIS) _
                        & P_DELIMITER & CStr(aCHR)
End Function

Public Function s_Packet_Character_Names(ByVal first_char As String, ByVal second_char As String, ByVal third_char As String, ByVal fourth_char As String) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Character_Names = first_char _
                            & P_DELIMITER & second_char _
                            & P_DELIMITER & third_char _
                            & P_DELIMITER & fourth_char
End Function

Public Function s_Packet_Character_Stats(ByVal char_name As String, ByVal race As Long, ByVal Class As Long, ByVal align As Long, ByVal sphere As Long, ByVal psionic_power As Long, ByVal level As Long, ByVal char_STR As Long, _
                                            ByVal char_DEX As Long, ByVal char_CON As Long, ByVal char_INT As Long, ByVal char_WIS As Long, ByVal char_CHR As Long, ByVal char_portrait As Long, ByVal char_data_index As Long) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/6/2004
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Character_Stats = char_name _
                            & P_DELIMITER & CStr(race) _
                            & P_DELIMITER & CStr(Class) _
                            & P_DELIMITER & CStr(align) _
                            & P_DELIMITER & CStr(sphere) _
                            & P_DELIMITER & CStr(psionic_power) _
                            & P_DELIMITER & CStr(level) _
                            & P_DELIMITER & CStr(char_STR) _
                            & P_DELIMITER & CStr(char_DEX) _
                            & P_DELIMITER & CStr(char_CON) _
                            & P_DELIMITER & CStr(char_INT) _
                            & P_DELIMITER & CStr(char_WIS) _
                            & P_DELIMITER & CStr(char_CHR) _
                            & P_DELIMITER & CStr(char_portrait) _
                            & P_DELIMITER & CStr(char_data_index)
End Function

Public Function s_Packet_Item_Create(ByVal map_x As Long, ByVal map_y As Long, ByVal Amount As Long, ByVal ItemDataIndex As Long) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/11/2004
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Item_Create = Generic_Packet_Map_Pos(map_x, map_y) _
                        & P_DELIMITER & CStr(Amount) _
                        & P_DELIMITER & CStr(ItemDataIndex)
End Function

Public Function c_Packet_NPC_Respond(ByVal npc_speech_index As Long, ByVal response_index As Long) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/11/2004
'Take parameters and creates string for inserting into a packet
'**************************************************************
c_Packet_NPC_Respond = CStr(npc_speech_index) _
                    & P_DELIMITER & CStr(response_index)
End Function

Public Function s_Packet_NPC_Speech(ByVal NPC_greet As String, ByVal response1 As String, ByVal response2 As String, ByVal response3 As String, ByVal response4 As String) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/11/2004
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_NPC_Speech = NPC_greet _
                        & P_DELIMITER & response1 _
                        & P_DELIMITER & response2 _
                        & P_DELIMITER & response3 _
                        & P_DELIMITER & response4
End Function

Public Function Generic_Packet_Map_Pos(ByVal map_x As Long, ByVal map_y As Long) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/11/2004
'Take parameters and creates string for inserting into a packet
'**************************************************************
Generic_Packet_Map_Pos = CStr(map_x) _
                        & P_DELIMITER & CStr(map_y)
End Function

Public Function s_Packet_Inventory_Item(ByVal slot As Long, ByVal item_index As Long, ByVal Amount As Long, ByVal equiped As Boolean) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/20/2004
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Inventory_Item = CStr(slot) _
                        & P_DELIMITER & CStr(item_index) _
                        & P_DELIMITER & CStr(Amount) _
                        & P_DELIMITER & CStr(equiped)
End Function

