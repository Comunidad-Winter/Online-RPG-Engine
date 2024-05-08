Attribute VB_Name = "modGeneral"
'*****************************************************************
'modGeneral.bas - Generic Functions - v0.5.0
'
'Generic helper functions
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
'Constants
'***************************
Public Const PI As Single = 3.14159265358979

'***************************
'External Functions
'***************************
'General_Var_Get and General_Var_Write
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'For making a form always on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Public Function General_File_Exists(ByVal file_path As String, ByVal file_type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************
    If Dir(file_path, file_type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, value, file
End Sub

Public Function General_Field_Read(ByVal field_pos As Long, ByVal text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(text)
        If delimiter = CByte(Asc(Mid$(text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = Mid$(text, LastPos + 1, (InStr(LastPos + 1, text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = Mid$(text, LastPos + 1)
    End If
End Function

Public Function General_Field_Count(ByVal text As String, ByVal delimiter As Byte) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Count the number of fields in a delimited string
'*****************************************************************
    'If string is empty there aren't any fields
    If Len(text) = 0 Then
        Exit Function
    End If

    Dim i As Long
    Dim FieldNum As Long
    FieldNum = 0
    For i = 1 To Len(text)
        If delimiter = CByte(Asc(Mid$(text, i, 1))) Then
            FieldNum = FieldNum + 1
        End If
    Next i
    General_Field_Count = FieldNum + 1
End Function

Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
'*****************************************************************
'Author: Aaron Perkins
'Find a Random number between a range
'*****************************************************************
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound + 1) * Rnd + LowerBound
End Function

Public Sub General_Sleep(ByVal length As Double)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sleep for a given number a seconds
'*****************************************************************
    Dim curFreq As Currency
    Dim curStart As Currency
    Dim curEnd As Currency
    Dim dblResult As Double
    
    QueryPerformanceFrequency curFreq 'Get the timer frequency
    QueryPerformanceCounter curStart 'Get the start time
    
    Do Until dblResult >= length
        QueryPerformanceCounter curEnd 'Get the end time
        dblResult = (curEnd - curStart) / curFreq 'Calculate the duration (in seconds)
    Loop
End Sub

Public Function General_Get_Elapsed_Time() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If

    'Get current time
    QueryPerformanceCounter start_time
    
    'Calculate elapsed time
    General_Get_Elapsed_Time = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    QueryPerformanceCounter end_time
End Function

Public Function General_String_Is_Alphanumeric(ByVal test_string As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/09/2003
'Alows no special characters. Only letters, numbers, and spaces
'**************************************************************
    'Check Name
    Dim loopc As Long
    Dim ts As Byte
    For loopc = 1 To Len(test_string)
        ts = Asc(Mid(test_string, loopc))
        If ts = 32 Or (ts >= 48 And ts <= 57) Or (ts >= 65 And ts <= 90) Or (ts >= 96 And ts <= 122) Then
            General_String_Is_Alphanumeric = True
        Else
            General_String_Is_Alphanumeric = False
            Exit Function
        End If
    Next loopc
End Function

Public Sub General_Form_On_Top_Set(ByRef s_form As Form, Optional ByVal s_on_top As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/09/2003
'Set a form on always on top or not
'**************************************************************
    Dim hwnd_flag As Long
    If s_on_top Then
        hwnd_flag = HWND_TOPMOST
    Else
        hwnd_flag = HWND_NOTOPMOST
    End If
    SetWindowPos s_form.hWnd, hwnd_flag, s_form.left / Screen.TwipsPerPixelX, s_form.top / Screen.TwipsPerPixelY, s_form.width / Screen.TwipsPerPixelX, s_form.height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Function General_Covert_Degrees_To_Radians(s_degree As Single) As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/12/2003
'Converts a degree to a radian
'**************************************************************
    General_Covert_Degrees_To_Radians = (s_degree * PI) / 180
End Function

Public Function General_Distance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Finds the distance between two points
'**************************************************************
    General_Distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function
