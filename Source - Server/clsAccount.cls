VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsAccount.cls - ORE Server Account Handling Object - v0.5.0
'
'Handles user´s accounts
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
'David Justus - 8/14/2004
'   - Add: Account_Ban, Account_Character_Get_Name_Web and Delete_Account methods
'
'Juan Martín Sotuyo Dodero(juansotuyo@hotmail.com) - 2/02/2004
'   - First Release
'*****************************************************************
Option Explicit

Dim Logged_In As Boolean
Dim acc_name As String
Dim acc_path As String

Public Function Account_Create(ByVal Account_name As String, ByVal password As String, ByVal firstname As String, ByVal lastname As String, ByVal email As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'Since it´s called from a website, we don´t need to log in
'**************************************************************
    Dim Account_path As String

    Account_path = acc_path & "\" & Account_name & ".acc"

    'Check if Account already exists
    If General_File_Exists(Account_path, vbNormal) Then
        Account_Create = False
        Exit Function
    End If

    'Write Account INFO
    General_Var_Write Account_path, "INFO", "password", password
    General_Var_Write Account_path, "INFO", "firstname", firstname
    General_Var_Write Account_path, "INFO", "lastname", lastname
    General_Var_Write Account_path, "INFO", "email", email
    General_Var_Write Account_path, "INFO", "Ban", "false"

    'Create empty holders for the chars
    General_Var_Write Account_path, "CHARACTERS", "1", ""
    General_Var_Write Account_path, "CHARACTERS", "2", ""
    General_Var_Write Account_path, "CHARACTERS", "3", ""
    General_Var_Write Account_path, "CHARACTERS", "4", ""

    Account_Create = True
End Function

Public Function Delete_Account(ByVal Account_name As String)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'Since it´s called from a website, we don´t need to log in
'**************************************************************
    Kill acc_path & "\" & Account_name & ".acc"
End Function

Public Function Account_Ban(ByVal Ban As Boolean, Account_name As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'Since it´s called from a website, we don´t need to log in
'**************************************************************
    Dim Account_path As String

    Account_path = App.Path & "\accounts\" & Account_name & ".acc"

    'Make sure Account exists
    If Not General_File_Exists(Account_path, vbNormal) Then
        Account_Ban = False
        Exit Function
    End If
    
    
    'Change password
    If Ban = True Then General_Var_Write Account_path, "INFO", "Ban", "true"
    If Ban = False Then General_Var_Write Account_path, "INFO", "Ban", "false"
    
    Account_Ban = True
End Function
Public Function Account_Password_Change(ByVal Account_name As String, ByVal cur_password As String, ByVal new_password As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'Since it´s called from a website, we don´t need to log in
'**************************************************************
    Dim Account_path As String

    Account_path = App.Path & "\accounts\" & Account_name & ".acc"

    'Make sure Account exists
    If Not General_File_Exists(Account_path, vbNormal) Then
        Account_Password_Change = False
        Exit Function
    End If
    
    'Check if the password suplyed is correct
    If Not cur_password = General_Var_Get(Account_path, "INFO", "password") Then
        Account_Password_Change = False
        Exit Function
    End If
    
    'Change password
    General_Var_Write Account_path, "INFO", "password", new_password
    
    Account_Password_Change = True
End Function

Public Function Account_Login(ByVal Account_name As String, ByVal password As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'**************************************************************
    acc_name = "\" & Account_name & ".acc"

    'Make sure Account exists
    If Not General_File_Exists(acc_path & acc_name, vbNormal) Then
        Logged_In = False
        Account_Login = False
        Exit Function
    End If
    
    If General_Var_Get(acc_path & acc_name, "INFO", "Ban") = "true" Then
        Logged_In = False
        Account_Login = False
        Exit Function
    End If
    'Check if the password suplyed is correct
    If Not password = General_Var_Get(acc_path & acc_name, "INFO", "password") Then
        Logged_In = False
        Account_Login = False
        Exit Function
    End If
    
    Logged_In = True
    Account_Login = True
End Function

Public Function Account_Character_Add(ByVal char_name As String, ByVal slot As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'**************************************************************
    If Not Logged_In Then
        Account_Character_Add = False
        Exit Function
    End If
    
    If slot > 4 Or slot < 0 Then
        Account_Character_Add = False
        Exit Function
    End If
    
    'Make sure the slot is free
    If General_Var_Get(acc_path & acc_name, "CHARACTERS", CStr(slot)) <> "" Then
        Account_Character_Add = False
        Exit Function
    End If
    
    General_Var_Write acc_path & acc_name, "CHARACTERS", CStr(slot), char_name
    
    Account_Character_Add = True
End Function

Public Function Account_Character_Remove(ByVal slot As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'**************************************************************
    If Not Logged_In Then
        Account_Character_Remove = False
        Exit Function
    End If
    
    If slot > 4 Or slot < 0 Then
        Account_Character_Remove = False
        Exit Function
    End If
    
    General_Var_Write acc_path & acc_name, "CHARACTERS", CStr(slot), ""
    
    Account_Character_Remove = True
End Function

Public Function Account_Login_Using_Char(ByVal slot As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'**************************************************************
    If Not Logged_In Then
        Account_Login_Using_Char = False
        Exit Function
    End If
    
    If slot > 4 Or slot < 0 Then
        Account_Login_Using_Char = False
        Exit Function
    End If
    
    'Load char from file
End Function

Public Function Account_Character_Get_Name(ByVal slot As Long) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'**************************************************************
    Dim LoopC As Long
    
    If Not Logged_In Then
        Account_Character_Get_Name = ""
        Exit Function
    End If
    
    If slot < 0 Or slot > 4 Then
        Account_Character_Get_Name = ""
        Exit Function
    End If
    
    Account_Character_Get_Name = General_Var_Get(acc_path & acc_name, "CHARACTERS", CStr(slot))
End Function

Public Function Account_Character_Get_Name_Web(ByVal slot As Long) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'**************************************************************
    On Error Resume Next
    
    Account_Character_Get_Name_Web = General_Var_Get(acc_path & acc_name, "CHARACTERS", CStr(slot))
End Function

Public Function Account_Character_Check(ByVal char_name As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/4/2004
'Checks a char belong to the current Account
'**************************************************************
    Dim LoopC As Long
    
    If Not Logged_In Then
        Account_Character_Check = False
        Exit Function
    End If
    
    LoopC = 1
    
    Do Until LoopC > 4
        If char_name = General_Var_Get(acc_path & acc_name, "CHARACTERS", CStr(LoopC)) Then
            Account_Character_Check = True
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    Account_Character_Check = False
End Function

Public Function Account_Initialize(ByVal accounts_path As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'**************************************************************
    acc_path = accounts_path
    
    Account_Initialize = True
End Function

Public Sub Account_Logoff()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'**************************************************************
    Logged_In = False
    acc_name = ""
End Sub

Public Property Get LoggedIn() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'**************************************************************
    LoggedIn = Logged_In
End Property
