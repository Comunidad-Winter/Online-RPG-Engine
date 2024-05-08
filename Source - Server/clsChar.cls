VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsChar.cls - ORE Server Character Object - v0.5.0
'
'Ties a player or NPC to a map and is used to refer to characters
'client side.
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
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************
Option Explicit

'***************************
'Variables
'***************************
Private char_player_id As Long
Private char_npc_id As Long

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
End Sub

Public Property Get PlayerID() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'*****************************************************************
        PlayerID = char_player_id
End Property

Public Property Let PlayerID(vNewValue As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'*****************************************************************
        char_player_id = vNewValue
End Property

Public Property Get NPCID() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'*****************************************************************
        NPCID = char_npc_id
End Property

Public Property Let NPCID(vNewValue As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/5/2003
'
'*****************************************************************
        char_npc_id = vNewValue
End Property
