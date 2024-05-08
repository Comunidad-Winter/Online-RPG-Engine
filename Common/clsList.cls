VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsList.cls - Generic Object List Class - v1.0.0
'
'A generic list data structure that takes object of any kind.
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
'Aaron Perkins(aaron@baronsoft.com) - 5/12/2003
'   -Add: UpperBound, LowerBound
'   -Add: Remove_Index, Update_Index
'
'Aaron Perkins(aaron@baronsoft.com) - 2/28/2003
'   -First Release
'*****************************************************************
Option Explicit

'***************************
'Variables
'***************************
Private list() As Object
Private list_count As Long
Private list_last As Long

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
    'Destory all objects in list
    Dim LoopC As Long
    For LoopC = LowerBound To UpperBound
        Set list(LoopC) = Nothing
    Next LoopC
End Sub

Private Function List_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Returns the next open index
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
    
    LoopC = 1
    Do Until list(LoopC) Is Nothing
        If LoopC = list_last Then
            List_Next_Open = list_last + 1
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    List_Next_Open = LoopC
Exit Function
ErrorHandler:
    List_Next_Open = 1
End Function

Public Function Add(ByRef Item As Object) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Add an object to the list
'*****************************************************************
    Dim list_index As Long
    
    list_index = List_Next_Open
    
    'Update array size
    If list_index > list_last Then
        list_last = list_index
        ReDim Preserve list(1 To list_last)
    End If
    list_count = list_count + 1
    
    Set list(list_index) = Item
    
    Add = list_index
End Function

Public Function Find(ByVal item_property As String, ByVal property_value As Variant) As Object
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Find and return an object by it's property
'*****************************************************************
    Dim Index As Long
    Index = Find_Index(item_property, property_value)
    
    If Index <> 0 Then
        Set Find = list(Index)
    Else
        Set Find = Nothing
    End If
End Function

Public Function Find_Index(ByVal property As String, ByVal value As Variant) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Find and return an index to a object by it's property
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
    Dim tempvalue As Variant
    
    LoopC = 1
    If Not (list(LoopC) Is Nothing) Then
        tempvalue = CallByName(list(LoopC), property, VbGet)
    End If
    Do Until tempvalue = value
        If LoopC = list_last Then
            Find_Index = 0
            Exit Function
        End If
        LoopC = LoopC + 1
        If Not (list(LoopC) Is Nothing) Then
             tempvalue = CallByName(list(LoopC), property, VbGet)
        End If
    Loop
    
    Find_Index = LoopC
Exit Function
ErrorHandler:
    Find_Index = 0
End Function

Public Sub Update(ByVal item_property As String, ByVal property_value As Variant, ByRef obj As Object)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Replace one object with another
'*****************************************************************
    Dim Index As Long
    Index = Find_Index(item_property, property_value)
    
    If Index <> 0 Then
        Update_Index Index, obj
    End If
End Sub

Public Sub Update_Index(ByVal Index As Long, ByRef obj As Object)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'Replace one object with another
'*****************************************************************
    If Index <> 0 Then
        Set list(Index) = obj
    End If
End Sub

Public Function Item(ByVal Index As Long) As Object
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Return and object by its index
'*****************************************************************
    If Index <> 0 Then
        Set Item = list(Index)
    Else
        Set Item = Nothing
    End If
End Function

Public Sub Remove(ByVal item_property As String, ByVal property_value As Variant)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Remove an object
'*****************************************************************
    Dim Index As Long
    Index = Find_Index(item_property, property_value)

    If Index <> 0 Then
        Remove_Index Index
    End If
End Sub

Public Sub Remove_Index(ByVal Index As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'Remove an object
'*****************************************************************
    If Index <> 0 Then
        'Destroy object
        Set list(Index) = Nothing
        'Update array size
        If Index = list_last Then
            Do Until Not (list(list_last) Is Nothing)
                list_last = list_last - 1
                If list_last = 0 Then
                    list_count = 0
                    Exit Sub
                End If
            Loop
            ReDim Preserve list(1 To list_last)
        End If
        list_count = list_count - 1
    Else
        
    End If
End Sub

Public Function Count() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'Return the total number of objects in the list
'*****************************************************************
    Count = list_count
End Function

Public Function UpperBound() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'Return the upper bounds of the list
'*****************************************************************
    UpperBound = list_last
End Function

Public Function LowerBound() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'Return the lower bounds of the list
'*****************************************************************
    LowerBound = 1
End Function

