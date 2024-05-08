Attribute VB_Name = "modGrh"
'*****************************************************************
'Minimap Color Finder - v1.0.0
'
'Finds the avarage color of a grh so it can be shown on the minimap
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
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 10/13/2004
'   -Change: Transparent pixels are not taken into account
'   -Change: The file is now binary, saving space
'   -Fix: Memory-leaks
'   -Fix: Several small bugs
'
'David Justus (big.david@txun.net) - 10/7/2004
'   - First Release
'*****************************************************************
Option Explicit

Public Const COLOR_KEY As Long = &H0

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'How big it is and animation info
Private Type Grh_Data
    active As Boolean
    texture_index As Long
    src_x As Long
    src_y As Long
    src_width As Long
    src_height As Long
    
    frame_count As Long
    frame_list(1 To 16) As Long
    frame_speed As Single
End Type

Public grh_list() As Grh_Data
Public grh_count As Long

Public Sub loadgrh()
On Error GoTo ErrorHandler
    Dim grh As Long
    Dim frame As Long
    Dim TempInt As Long
    
    Dim inipath As String
    inipath = App.Path & "\..\Resources\Graphics\"
    
    'Get number of grhs
    grh_count = Val(General_Var_Get(inipath & "grh.ini", "INIT", "grh_count"))
    
    'Resize arrays
    ReDim grh_list(1 To grh_count) As Grh_Data
    
    'Open files
    Open inipath & "grh.dat" For Binary As #1
    Seek #1, 1
    
    'Get Header
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    
    'Fill Grh List
    
    'Get first Grh Number
    Get #1, , grh
    
    Do Until grh <= 0
        
        grh_list(grh).active = True
        
        'Get number of frames
        Get #1, , grh_list(grh).frame_count
        If grh_list(grh).frame_count <= 0 Then GoTo ErrorHandler
        
        If grh_list(grh).frame_count > 1 Then
        
            'Read a animation GRH set
            For frame = 1 To grh_list(grh).frame_count
            
                Get #1, , grh_list(grh).frame_list(frame)
                If grh_list(grh).frame_list(frame) <= 0 Or grh_list(grh).frame_list(frame) > grh_count Then GoTo ErrorHandler
            
            Next frame
        
            Get #1, , grh_list(grh).frame_speed
            If grh_list(grh).frame_speed = 0 Then GoTo ErrorHandler
            
            'Compute width and height
            grh_list(grh).src_height = grh_list(grh_list(grh).frame_list(1)).src_height
            If grh_list(grh).src_height <= 0 Then GoTo ErrorHandler
            
            grh_list(grh).src_width = grh_list(grh_list(grh).frame_list(1)).src_width
            If grh_list(grh).src_width <= 0 Then GoTo ErrorHandler
        
        Else
        
            'Read in normal GRH data
            Get #1, , grh_list(grh).texture_index
            If grh_list(grh).texture_index <= 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_x
            If grh_list(grh).src_x < 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_y
            If grh_list(grh).src_y < 0 Then GoTo ErrorHandler
                
            Get #1, , grh_list(grh).src_width
            If grh_list(grh).src_width <= 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_height
            If grh_list(grh).src_height <= 0 Then GoTo ErrorHandler
            
            grh_list(grh).frame_list(1) = grh
                
        End If
    
        'Get Next Grh Number
        Get #1, , grh
    
    Loop
    '************************************************
    
    Close #1
Exit Sub
ErrorHandler:
    Close #1
    MsgBox "Error while loading the grh.dat! Stopped at GRH number: " & grh
End Sub

Public Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grh_count Then
        If grh_list(grh_index).active Then
            Grh_Check = True
        End If
    End If
End Function

Function Grh_get_value(ByVal grh_index As Long, ByVal destHDC As Long, ByVal screen_x As Long, ByVal screen_y As Long, Optional transparent As Boolean = False) As Long
'**************************************************************
'Author: David Justus
'Last Modify Date: 10/09/2004
'Modified by Juan Martín Sotuyo Dodero
'*************************************************************
    Dim X As Long
    Dim Y As Long
    Dim file_path As String
    Dim src_x As Long
    Dim src_y As Long
    Dim src_width As Long
    Dim src_height As Long
    Dim hdcsrc As Long
    Dim OldObj As Long
    Dim value As Currency
    
    'If it's animated switch grh_index to first frame
    If grh_list(grh_index).frame_count <> 1 Then
        grh_index = grh_list(grh_index).frame_list(1)
    End If
    
    file_path = App.Path & "\..\resources\graphics\grh" & grh_list(grh_index).texture_index & ".bmp"
    
    If Not General_File_Exists(file_path, vbNormal) Then Exit Function
    
    src_x = grh_list(grh_index).src_x
    src_y = grh_list(grh_index).src_y
    src_width = grh_list(grh_index).src_width
    src_height = grh_list(grh_index).src_height
    
    hdcsrc = CreateCompatibleDC(destHDC)
    OldObj = SelectObject(hdcsrc, LoadPicture(file_path))
    
    BitBlt destHDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
    
    DeleteObject SelectObject(hdcsrc, OldObj)
    DeleteDC hdcsrc
    
    DoEvents
    
    Dim R As Currency
    Dim B As Currency
    Dim G As Currency
    Dim TempR As Integer
    Dim TempG As Integer
    Dim TempB As Integer
    Dim InvalidPixels As Long
    
    For X = 0 To grh_list(grh_index).src_height - 1
        For Y = 0 To grh_list(grh_index).src_width - 1
            'Color is not taken into account if the color is transparent
            If GetPixel(destHDC, X, Y) = COLOR_KEY Then
                InvalidPixels = InvalidPixels + 1
            Else
                General_Long_Color_to_RGB GetPixel(destHDC, X, Y), TempR, TempG, TempB
                R = R + TempR
                G = G + TempG
                B = B + TempB
            End If
            DoEvents
        Next Y
    Next X
    
    Dim size As Long
    
    size = grh_list(grh_index).src_height * grh_list(grh_index).src_width - InvalidPixels
    
    If size = 0 Then size = 1
    Grh_get_value = RGB(R / size, G / size, B / size)
End Function
