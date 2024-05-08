VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsSpeech.cls - Speech Class - v1.0.0
'
'Holds all NPC speeches along with all possible answers
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
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 6/16/2004
'   -First Release
'*****************************************************************
Option Explicit

Private Type response
    response_text() As String       'The text of the player´s response
    leads_to() As Long              'The NPC´s dialog to be shown with that response
    ends_conversation As Boolean    'Weather if it ends the conversation or not
End Type

Private npc_speech() As String
Private responses() As response

Private speech_path As String
Public speech_name As String

Public Function Initialize(ByVal s_speech_path As String, ByVal s_speech_file As String) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/16/2004
'
'*****************************************************************
    Dim LoopC As Long
    Dim LoopD As Long
    Dim temp_text As String
    Dim NumResponses As Long
    
    'Check speech file is correct
    If Not General_File_Exists(s_speech_path & "\" & s_speech_file, vbNormal) Then
        Initialize = False
        Exit Function
    End If
    
    speech_path = s_speech_path & "\"
    speech_name = s_speech_file
    
    '****************
    'Load speech
    '****************
    
    'Resize array
    ReDim npc_speech(0)
    ReDim responses(0)
    
    'Set number of responses
    NumResponses = Val(General_Var_Get(speech_path & speech_name, "SPEECH", "NumResponses"))
    ReDim responses(0).leads_to(1 To NumResponses)
    ReDim responses(0).response_text(1 To NumResponses)
    
    'Load greeting
    npc_speech(0) = General_Var_Get(speech_path & speech_name, "SPEECH", "greet")
    For LoopC = 1 To NumResponses
        temp_text = General_Var_Get(speech_path & speech_name, "SPEECH", CStr(LoopC))
        responses(0).leads_to(LoopC) = Val(General_Field_Read(1, temp_text, 44))
        If temp_text <> "" Then
            responses(0).response_text(LoopC) = Right$(temp_text, Len(temp_text) - 1 - Len(CStr(responses(0).leads_to(LoopC))))
        End If
    Next LoopC
    
    'Now do the same with all responses
    LoopD = 1
    Do Until General_Var_Get(speech_path & speech_name, CStr(LoopD), "1") = "" And General_Var_Get(speech_path & speech_name, CStr(LoopD), "end") = ""
            'Resize arrays
            ReDim Preserve npc_speech(0 To LoopD)
            ReDim Preserve responses(0 To LoopD)
            
            npc_speech(LoopD) = General_Var_Get(speech_path & speech_name, CStr(LoopD), "greet")
            
            NumResponses = Val(General_Var_Get(speech_path & speech_name, CStr(LoopD), "NumResponses"))
            
            If NumResponses Then
                ReDim Preserve responses(LoopD).response_text(1 To NumResponses)
                ReDim Preserve responses(LoopD).leads_to(1 To NumResponses)
            Else
                ReDim Preserve responses(LoopD).response_text(0)
                ReDim Preserve responses(LoopD).leads_to(0)
            End If
            
            For LoopC = 1 To NumResponses
                temp_text = General_Var_Get(speech_path & speech_name, CStr(LoopD), CStr(LoopC))
                responses(LoopD).leads_to(LoopC) = Val(General_Field_Read(1, temp_text, 44))
                responses(LoopD).response_text(LoopC) = General_Field_Read(2, temp_text, 44)
            Next LoopC
            
            If General_Var_Get(speech_path & speech_name, CStr(LoopD), "end") = "true" Then
                responses(LoopD).ends_conversation = True
            End If
        LoopD = LoopD + 1
    Loop
    
    Initialize = True
End Function

Public Function Speech_Get_Speech(ByVal NPC_greet_index As Long) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/16/2004
'Creates a packet that returns the NPC´s greet and all responses available
'*****************************************************************
    Dim temp_text As String
    Dim LoopC As Long
    
    For LoopC = 1 To UBound(responses(NPC_greet_index).response_text())
        temp_text = temp_text & P_DELIMITER & responses(NPC_greet_index).response_text(LoopC)
    Next LoopC
    
    Speech_Get_Speech = npc_speech(NPC_greet_index) & P_DELIMITER & UBound(responses(NPC_greet_index).response_text()) & temp_text
End Function

Public Function Speech_Respond(ByVal NPC_greet_index As Long, ByVal response_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/16/2004
'Returns the NPC´s speech index corresponding to the given response and -1 if the conversation ended
'*****************************************************************
    If responses(responses(NPC_greet_index).leads_to(response_index)).ends_conversation Then
        Speech_Respond = -1
        Exit Function
    End If
    
    Speech_Respond = responses(NPC_greet_index).leads_to(response_index)
End Function
