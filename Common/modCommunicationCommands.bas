Attribute VB_Name = "modCommunicationCommands"
'*****************************************************************
'modORECommands.bas - ORE Communication Command Constants - v0.5.0
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
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************

Option Explicit

'***************************
'Constants and enumerations
'***************************
Public Const P_DELIMITER_CODE As Byte = 31
Public P_DELIMITER As String

'***************
'Server commands
'***************
'Header
Public Enum ServerPacketHeader
    s_Player = 1 'Start at 1
    s_Chat
    s_Map
    s_Char
End Enum

'Command
Public Enum ServerPacketCommand
    s_Player_Authenticated = 1
    s_Player_Engine_Start
    s_Player_Engine_Stop
    s_Chat_Text
    s_Chat_Critical
    s_Map_Load
    s_Char_ID_Set
    s_Char_Create
    s_Char_Label_Set
    s_Char_Data_Set
    s_Char_Data_Body_Set
    s_Char_Pos_Set
    s_Char_Heading_Set
    s_Char_Move
    s_Char_Remove
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
End Enum

'Command
Public Enum ClientPacketCommand
    c_Authenticate_Login = 10001
    c_Authenticate_New
    c_Chat_Global
    c_Request_Pos_Update
    c_Move_Moved
    c_Action_Attack
End Enum

Public Function s_Packet_Char_Create(ByVal char_id As Long, ByVal map_x As Long, ByVal map_y As Long, ByVal heading As Long, _
                                    ByVal char_data_index As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Create = CStr(char_id) _
                    & P_DELIMITER & CStr(map_x) _
                    & P_DELIMITER & CStr(map_y) _
                    & P_DELIMITER & CStr(heading) _
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

Public Function s_Packet_Char_Heading_Set(ByVal char_id As Long, ByVal heading As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Heading_Set = CStr(char_id) _
                          & P_DELIMITER & CStr(heading)
End Function

Public Function s_Packet_Char_Move(ByVal char_id As Long, ByVal heading As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Move = CStr(char_id) _
                    & P_DELIMITER & CStr(heading)
End Function

Public Function s_Packet_Char_Pos_Set(ByVal char_id As Long, ByVal map_x As Long, ByVal map_y As Long) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/25/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
s_Packet_Char_Pos_Set = CStr(char_id) _
                    & P_DELIMITER & CStr(map_x) _
                    & P_DELIMITER & CStr(map_y)
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

Public Function c_Packet_Player_New(ByVal password As String, ByVal profile_name As String) As String
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'Take parameters and creates string for inserting into a packet
'**************************************************************
c_Packet_Player_New = password _
                      & P_DELIMITER & profile_name
End Function

