VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORE Scripter"
   ClientHeight    =   3450
   ClientLeft      =   1935
   ClientTop       =   1905
   ClientWidth     =   7560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7560
   Begin VB.Frame Frame2 
      Caption         =   "Layer"
      Height          =   1335
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
      Begin VB.CommandButton SetDefaultLayerCmd 
         Caption         =   "Set as default layer"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox LayerLst 
         Height          =   315
         ItemData        =   "frmMain.frx":08CA
         Left            =   120
         List            =   "frmMain.frx":08DD
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nodes"
      Height          =   1575
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      Begin VB.TextBox NodeNameTxt 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton NodCreateCmd 
         Caption         =   "Create Node"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.ListBox GrhList 
      Height          =   3180
      Left            =   120
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin MSComctlLib.TreeView ScriptTree 
      Height          =   3255
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5741
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&File"
      Begin VB.Menu LoadGrhScriptMnu 
         Caption         =   "Load &Grh Script"
      End
      Begin VB.Menu LoadTileScriptMnu 
         Caption         =   "Load &Tile Script"
      End
      Begin VB.Menu SaveMnu 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveGrhScriptMnu 
         Caption         =   "S&ave as Grh Script"
         Shortcut        =   ^G
      End
      Begin VB.Menu SaveTileScriptMnu 
         Caption         =   "Sa&ve as Tile Script"
         Shortcut        =   ^T
      End
      Begin VB.Menu SeparatorMnu 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu ViewMnu 
      Caption         =   "&View"
      Begin VB.Menu GrhViewerChkMnu 
         Caption         =   "&Grh Viewer"
         Checked         =   -1  'True
      End
      Begin VB.Menu HideGrhsMnu 
         Caption         =   "&Hide already placed Grhs"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu ConfigMnu 
      Caption         =   "&Configuration"
   End
   Begin VB.Menu AboutMnu 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'ORE scriptor - v1.0
'
'Creates script files for ORE.
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

'*****************************************************************
' 2/12/2004 - Juan Martín Sotuyo Dodero  (juansotuyo@hotmail.com)
'   - First Release
'*****************************************************************
Option Explicit

Dim loading_tree As Boolean

Private Sub ConfigMnu_Click()
    If GrhViewerChkMnu.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    
    frmConfiguration.Show vbModal
    
    If GrhViewerChkMnu.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
End Sub

Private Sub ExitMnu_Click()
    Dim Cancel As Integer
    Dim UnloadMode As Integer
    Form_QueryUnload Cancel, UnloadMode
End Sub

Private Sub Form_Load()
    Dim resource_path As String
    Dim LoopC As Long
    
    Me.Show
    frmGrhViewer.Show
    Load_User_Defined
    
Check:
    'Make sure graphics path is valid
    If Dir(graphics_path & "\grh.dat") = "" Then
        MsgBox "Couldn´t find grh.dat file at the given graphics path."
        frmConfiguration.Show vbModal
        GoTo Check
    End If
    
    General_Form_On_Top_Set frmGrhViewer, True
    
    'Create resource path
    For LoopC = 1 To General_Field_Count(graphics_path, 92) - 1
        resource_path = resource_path & General_Field_Read(LoopC, graphics_path, 92) & "\"
    Next LoopC
    'Remove last backslash
    resource_path = left$(resource_path, Len(resource_path) - 1)
    engine.Engine_Initialize Me.hWnd, Me.hWnd, True, resource_path
    engine.Grh_Add_GrhList_To_ListBox Me.GrhList
    LayerLst.ListIndex = 0
End Sub

Public Function Load_Grh_Tree(TreeName As TreeView, file_path As String) As Boolean
'*************************************************
'Coded by Juan Martín Sotuyo Dodero
'Last Modified: 12/10/03
'Loads the Tile Script file to the given tree view
'*************************************************
    Dim ScriptLine As String
    Dim NodeLevel As Long
    Dim NodeCount As Long
    Dim GrhCount As Long
    Dim ParentNodeKey As String
    Dim LoopC As Long
    Dim fso As FileSystemObject
    Dim strm As TextStream
    Dim NodeX As Node
    
    'Set flag
    loading_tree = True
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Check script file exists
    If Not General_File_Exists(file_path, vbNormal) Then
        MsgBox "File " & file_path & " doesn´t exist!", , "Error"
        Exit Function
    End If
    
    Set strm = fso.OpenTextFile(file_path, ForReading)
     
    With strm
        Do Until .AtEndOfStream
            ScriptLine = .ReadLine
            
            If left$(ScriptLine, 1) = "#" And left$(ScriptLine, 4) <> "#EOF" Then
                NodeLevel = NodeLevel + 1
                NodeCount = NodeCount + 1
                'Check if it´s a child or a root
                If NodeLevel = 1 Then
                    Set NodeX = TreeName.Nodes.Add(, , Str(NodeLevel) & "-" & Str(NodeCount), Right$(ScriptLine, Len(ScriptLine) - 1))
                Else
                    Do Until Val(General_Field_Read(1, NodeX.Key, 45)) = NodeLevel - 1
                        Set NodeX = NodeX.Parent
                    Loop
                    ParentNodeKey = NodeX.Key
                    Set NodeX = TreeName.Nodes.Add(ParentNodeKey, tvwChild, Str(NodeLevel) & "-" & Str(NodeCount), Right$(ScriptLine, Len(ScriptLine) - 1))
                    NodeX.EnsureVisible
                End If
            End If
            
            If left$(ScriptLine, 4) = "#EOF" Then
                NodeLevel = NodeLevel - 1
                'Check if file has ended
                If NodeLevel = -1 Then Exit Do
                NodeCount = Val(General_Field_Read(2, NodeX.Parent.Key, 45))
            End If
            
            If left$(ScriptLine, 1) = "$" Then
                GrhCount = GrhCount + 1
                Set NodeX = TreeName.Nodes.Add(Str(NodeLevel) & "-" & Str(NodeCount), tvwChild, "grh" & Str(GrhCount), "< On " & LayerLst.List(Val(Right$(ScriptLine, Len(ScriptLine) - 1)) - 1) & " Layer >")
                NodeX.EnsureVisible
            End If
            
            If left$(ScriptLine, 1) = ">" Then
                GrhCount = GrhCount + 1
                If left$(ScriptLine, 2) = ">$" Then
                    'Parent is a Grh
                    Set NodeX = TreeName.Nodes.Add("grh" & "-" & Str(GrhCount - 1), tvwChild, "grh" & Str(GrhCount), "< On " & LayerLst.List(Val(Right$(ScriptLine, Len(ScriptLine) - 1)) - 1) & " Layer >")
                Else
                    Set NodeX = TreeName.Nodes.Add(Str(NodeLevel) & "-" & Str(NodeCount), tvwChild, "grh" & Str(GrhCount), "Grh " & Right$(ScriptLine, Len(ScriptLine) - 1))
                End If
                NodeX.EnsureVisible
            End If
        Loop
        
        For LoopC = 1 To TreeName.Nodes.Count
            TreeName.Nodes(LoopC).Expanded = False
        Next LoopC
        
    End With
    
    'Reset flag
    loading_tree = False
    
    Load_Grh_Tree = True
Exit Function

err:
    loading_tree = False
    'Hide loading message
    MsgBox "Incorrect Script found."
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If GrhViewerChkMnu.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    
    'Check if file was edited
    If Modified Then
        'Check if a file is loaded
        If LoadGrhScriptMnu.Enabled = False Or LoadTileScriptMnu.Enabled = False Then
            If MsgBox("Changes have been made since the file was last saved. If you don´t save changes will be lost. Do you want to save now?", vbYesNo) = vbYes Then
                Call SaveMnu_Click
            End If
        End If
    End If
    
    'Close program
    If GrhViewerChkMnu.Checked Then
        Unload frmGrhViewer
    End If
    
    Unload Me
End Sub

Private Sub GrhList_Click()
    'Dsiplay grh in the grh viewer
    If GrhViewerChkMnu.Checked Then
        frmGrhViewer.Cls
        engine.Grh_Render_To_Hdc Val(GrhList.text), frmGrhViewer.hdc, 0, 0
    End If
End Sub

Private Sub GrhList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DataArray As String
    Dim NodeKey As String
    Dim index As Long
    Dim GrhIndex As Long
    Dim LoopC As Long
    
    'Check used Grh are being hidden, otherwise it´s useless
    If Not HideGrhsMnu.Checked Then Exit Sub
    
    'If a Grh is being moved back allow it
    DataArray = Data.GetData(1)
    
    'Get the text of the first node. If it´s a Grh it´s everything we need, and otherwise we won´t do anything
    NodeKey = General_Field_Read(1, DataArray, 92)
    DataArray = General_Field_Read(2, DataArray, 92)
    
    If left$(DataArray, 4) = "Grh " Then
        GrhIndex = Val(Right(DataArray, Len(DataArray) - 4))
        'Find the correct index
        Do Until Val(GrhList.List(LoopC)) > GrhIndex
            LoopC = LoopC + 1
        Loop
        GrhList.AddItem GrhIndex, LoopC
        
        'Remove the node from the tree
        For LoopC = 1 To ScriptTree.Nodes.Count
            If ScriptTree.Nodes(LoopC).Key = NodeKey Then
                ScriptTree.Nodes.Remove LoopC
                Exit For
            End If
        Next LoopC
    End If
End Sub

Private Sub GrhList_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    'Data is set here, since the OLESetData event is called from the Drop place,
    'and has no way to know it came from here (besides it´s too short and isn´t worth doing another method)
    Data.SetData "Grh " & GrhList.text, 1
End Sub

Private Sub GrhViewerChkMnu_Click()
    If GrhViewerChkMnu.Checked Then
        Unload frmGrhViewer
    Else
        GrhViewerChkMnu.Checked = True
        frmGrhViewer.Show
    End If
End Sub

Private Sub AboutMnu_Click()
    frmAbout.Show vbModal
End Sub

Private Sub HideGrhsMnu_Click()
    HideGrhsMnu.Checked = Not HideGrhsMnu.Checked
    
    If HideGrhsMnu.Checked Then
        Dim LoopC As Long
        Dim LoopC2 As Long
        For LoopC = ScriptTree.Nodes.Count To 1 Step -1
            'Check if it´s a Grh
            If left$(ScriptTree.Nodes(LoopC).text, 4) = "Grh " Then
                'Remove it from the Grh List
                For LoopC2 = 0 To GrhList.ListCount
                    If GrhList.List(LoopC2) = Right(ScriptTree.Nodes(LoopC).text, Len(ScriptTree.Nodes(LoopC).text) - 4) Then
                        GrhList.RemoveItem LoopC2
                        'Continue
                        Exit For
                    End If
                Next LoopC2
            End If
        Next LoopC
    Else
        GrhList.Clear
        engine.Grh_Add_GrhList_To_ListBox GrhList
    End If
End Sub

Private Sub LoadGrhScriptMnu_Click()
    'Check previous file was not modified
    If Modified Then
        'Check if it´s a new file
        If frmMain.Caption = "Dark Sun Online Scripter" Then
            If MsgBox("Changes have been made since the file was last saved. If you proceed changes will be lost. Continue?", vbYesNo) = vbNo Then Exit Sub
        Else
            Select Case MsgBox("Changes have been made since the file was last saved. If you don´t save changes will be lost. Do you want to save now?", vbYesNoCancel)
                Case vbYes
                    Call SaveMnu_Click
                Case vbCancel
                    Exit Sub
            End Select
        End If
    End If
    Clear_Tree
    Modified = False
    If Not Load_Grh_Tree(ScriptTree, scripts_path & "\Grh Script.dat") Then
        MsgBox "File couldn´t be loaded. Make sure file exists or path is correct."
        LoadGrhScriptMnu.Enabled = True
        LoadTileScriptMnu.Enabled = True
        Caption = "Dark Sun Online Scripter"
        Exit Sub
    End If
    
    LoadGrhScriptMnu.Enabled = False
    LoadTileScriptMnu.Enabled = True
    
    Caption = "Dark Sun Online Scripter - Grh Script"
End Sub

Private Sub LoadTileScriptMnu_Click()
    'Check previous file was not modified
    If Modified Then
        'Check if it´s a new file
        If frmMain.Caption = "Dark Sun Online Scripter" Then
            If MsgBox("Changes have been made since the file was last saved. If you proceed changes will be lost. Continue?", vbYesNo) = vbNo Then Exit Sub
        Else
            Select Case MsgBox("Changes have been made since the file was last saved. If you don´t save changes will be lost. Do you want to save now?", vbYesNoCancel)
                Case vbYes
                    Call SaveMnu_Click
                Case vbCancel
                    Exit Sub
            End Select
        End If
    End If
    Clear_Tree
    Modified = False
    If Not Load_Grh_Tree(ScriptTree, scripts_path & "\Tile Script.dat") Then
        MsgBox "File couldn´t be loaded. Make sure file exists or path is correct."
        LoadGrhScriptMnu.Enabled = True
        LoadTileScriptMnu.Enabled = True
        Caption = "Dark Sun Online Scripter"
        Exit Sub
    End If
    
    LoadGrhScriptMnu.Enabled = True
    LoadTileScriptMnu.Enabled = False
    
    Caption = "Dark Sun Online Scripter - Tile Script"
End Sub

Private Sub Clear_Tree()
    Dim LoopC As Long
    
    For LoopC = ScriptTree.Nodes.Count To 1 Step -1
        ScriptTree.Nodes.Remove (LoopC)
    Next
End Sub

Private Sub Load_User_Defined()
    graphics_path = General_Var_Get(App.Path & "\ORE Scripter.ini", "INIT", "graphics_path")
    scripts_path = General_Var_Get(App.Path & "\ORE Scripter.ini", "INIT", "scripts_path")
End Sub

Private Sub NodCreateCmd_Click()
    'Creates a new node (not grhs)
    'Nodes are created at the bottom as first level nodes. Then they can be moved around
    If NodeNameTxt.text = "" Then
        MsgBox "Node must have a name"
        Exit Sub
    End If
    
    Dim NodeX As Node
    
    NodeCount = NodeCount + 1
    Set NodeX = ScriptTree.Nodes.Add(, , "Temp" & NodeCount, NodeNameTxt.text)
    NodeX.EnsureVisible
    
    Modified = True
    SaveMnu.Enabled = True
End Sub

Private Sub NodeNameTxt_GotFocus()
    NodeNameTxt.SelStart = 0
    NodeNameTxt.SelLength = Len(NodeNameTxt.text)
End Sub

Private Sub SaveGrhScriptMnu_Click()
    If Not Save_Script_File(scripts_path & "\Grh Script.dat") Then GoTo Errhandler
    
    Modified = False
    SaveMnu.Enabled = False
    
    MsgBox "Script file saved."
Exit Sub

Errhandler:
    MsgBox "Script file couldn´t be saved."
End Sub

Private Sub SaveMnu_Click()
    'Check which file we are editing
    If LoadGrhScriptMnu.Enabled = False Then
        If Not Save_Script_File(scripts_path & "\Grh Script.dat") Then GoTo Errhandler
    ElseIf LoadTileScriptMnu.Enabled = False Then
        If Not Save_Script_File(scripts_path & "\Tile Script.dat") Then GoTo Errhandler
    Else
        Exit Sub
    End If
    
    Modified = False
    SaveMnu.Enabled = False
    
    MsgBox "Script file saved."
Exit Sub

Errhandler:
    MsgBox "Script file couldn´t be saved."
End Sub

Private Sub SaveTileScriptMnu_Click()
    If Not Save_Script_File(scripts_path & "\Tile Script.dat") Then GoTo Errhandler
    
    Modified = False
    SaveMnu.Enabled = False
    
    MsgBox "Script file saved."
Exit Sub

Errhandler:
    MsgBox "Script file couldn´t be saved."
End Sub

Private Sub ScriptTree_KeyDown(KeyCode As Integer, Shift As Integer)
    'Check if the Supr key was hitted, and if so delete selected node and subnodes
    Dim LoopC As Long
    Dim SelectedNode As Node
    
    'Ignore all commands while loading
    If loading_tree Then Exit Sub
    
    If KeyCode = 46 And ScriptTree.Nodes.Count > 0 Then
        For LoopC = 1 To ScriptTree.Nodes.Count
            If ScriptTree.Nodes(LoopC).Selected Then
                Set SelectedNode = ScriptTree.Nodes(LoopC)
            End If
        Next LoopC
        If HideGrhsMnu.Checked Then
            Delete_Node SelectedNode
        Else
            ScriptTree.Nodes.Remove SelectedNode.index
        End If
    End If
End Sub

Private Sub ScriptTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Errhandler:
    'Ignore all comands while loading
    If loading_tree Then Exit Sub
    
    'When clicking and draging the node is not selected automatically, so we do it here
    Dim TempNode As Node
    Dim LoopC As Long
    
    'Values are already in twips, so we don´t edit them
    Set TempNode = ScriptTree.HitTest(X, Y)
    
    For LoopC = ScriptTree.Nodes.Count To 1 Step -1
        If ScriptTree.Nodes(LoopC).Key = TempNode.Key Then
            ScriptTree.Nodes(LoopC).Selected = True
        Else
            ScriptTree.Nodes(LoopC).Selected = False
        End If
    Next LoopC
    
Errhandler:
End Sub

Private Sub ScriptTree_NodeClick(ByVal Node As MSComctlLib.Node)
    'Check if it´s a Grh
    Dim grh_index As Long
    
    'Ignore all commands while loading
    If loading_tree Then Exit Sub
    
    If left$(Node.text, 4) = "Grh " Then
        grh_index = Val(Right(Node.text, Len(Node.text) - 4))
        If GrhViewerChkMnu.Checked Then
            frmGrhViewer.Cls
            engine.Grh_Render_To_Hdc grh_index, frmGrhViewer.hdc, 0, 0
        End If
    End If
End Sub

Private Sub ScriptTree_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Errhandler
    Dim DataArray As String
    Dim NodeName As String
    Dim NodeParent As Node
    Dim NewNode As Node
    Dim LoopC As Long
    
    'Ignore all comands while loading
    If loading_tree Then Exit Sub
    
    'Get Data
    DataArray = Data.GetData(1)
    
    'Get parent node (multiply by 15 to convert to twips)
    'Here an error may occur if no node is found. Otherwise the sub runs without problem.
    Set NodeParent = ScriptTree.HitTest(X * 15, Y * 15)
    If left$(NodeParent.text, 3) = "Grh" Or left$(NodeParent.text, 5) = "< On " Then
        'Make sure parent is valid, and not a Grh or a layer node
        Set NodeParent = NodeParent.Parent
    End If
    
    'Check that NodeParent is not nested in nodes to be passed
    For LoopC = 1 To ScriptTree.Nodes.Count
        If ScriptTree.Nodes(LoopC).Key = General_Field_Read(1, DataArray, 92) Then
            Set NewNode = ScriptTree.Nodes(LoopC)
            Exit For
        End If
    Next LoopC
    If left$(DataArray, 4) <> "Grh " Then
        If Check_Node_Is_Children(NewNode, NodeParent) Then
            MsgBox "You can´t move a node into another nested beneath.", , "Error"
            Exit Sub
        End If
    End If
    
    'Delete source nodes
    For LoopC = 1 To ScriptTree.Nodes.Count
        If ScriptTree.Nodes(LoopC).Key = General_Field_Read(1, DataArray, 92) Then
            ScriptTree.Nodes.Remove LoopC
            DataArray = Right(DataArray, Len(DataArray) - Len(General_Field_Read(1, DataArray, 92)) - 1)
            Exit For
        End If
    Next LoopC
    
    'Create nodes
    Do While Len(DataArray) > 0
        NodeName = General_Field_Read(1, DataArray, 92)
        If NodeName <> "" Then
            'Update data array
            DataArray = Right(DataArray, Len(DataArray) - Len(NodeName))
            'Create node
            NodeCount = NodeCount + 1
            Set NewNode = ScriptTree.Nodes.Add(NodeParent.Key, tvwChild, "Temp" & NodeCount, NodeName)
            NewNode.EnsureVisible
            
            'If the created node is a grh, and the Hide Grhs option is checked, remove it from Grh list
            If left$(NodeName, 4) = "Grh " And HideGrhsMnu.Checked Then
                For LoopC = 0 To GrhList.ListCount
                    If GrhList.List(LoopC) = Right(NodeName, Len(NodeName) - 4) Then
                        GrhList.RemoveItem LoopC
                        Exit For
                    End If
                Next LoopC
            End If
        End If
        
        'Check next separator
        If left$(DataArray, 2) = Child Then
            Set NodeParent = NewNode
            DataArray = Right(DataArray, Len(DataArray) - 2)
        ElseIf left$(DataArray, 4) = ChildrenEnd Then
            Set NodeParent = NodeParent.Parent
            DataArray = Right(DataArray, Len(DataArray) - 4)
        ElseIf Len(DataArray) > 0 Then
            'Parent remains the same
            DataArray = Right(DataArray, Len(DataArray) - 1)
        End If
    Loop
    
    Modified = True
    SaveMnu.Enabled = True

Exit Sub

Errhandler:
End Sub

Private Sub ScriptTree_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    Dim LoopC As Long
    Dim SelectedNode As Node
    Dim DataArray As String
    
    'Ignore all comands while loading
    If loading_tree Then Exit Sub
    
    'Check the name of the clicked node
    For LoopC = ScriptTree.Nodes.Count To 1 Step -1
        If ScriptTree.Nodes(LoopC).Selected Then
            Set SelectedNode = ScriptTree.Nodes(LoopC)
        End If
    Next LoopC
    
    'Create the data array.
    DataArray = SelectedNode.Key & Separator & SelectedNode.text
    Add_Children_To_Data_Array DataArray, SelectedNode
    
    Data.SetData DataArray, 1
End Sub

Private Sub ScriptTree_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    'Ignore all comands while loading
    If loading_tree Then Exit Sub
    
    AllowedEffects = 1  'Copy items, so they aren´t deleted. We will do so if they were correctly copied
End Sub

Private Sub Add_Children_To_Data_Array(ByRef DataArray As String, SelectedNode As Node)
    Dim LastSibling As Node
    Dim CurrentSibling As Node
    
    'Set the first separator, so when the second one is added the Child separator is formed
    DataArray = DataArray & Separator
    If SelectedNode.Children > 0 Then
        'Load all children
        
        'Set the first child as the selected node
        Set SelectedNode = SelectedNode.Child
        'Start storing data on all it´s siblings
        Set LastSibling = SelectedNode.LastSibling
        Set CurrentSibling = SelectedNode.FirstSibling
        Do While CurrentSibling.Key <> LastSibling.Key
            DataArray = DataArray & Separator & CurrentSibling.text
            If CurrentSibling.Children > 0 Then
                Add_Children_To_Data_Array DataArray, CurrentSibling
            End If
            Set CurrentSibling = SelectedNode.Next
        Loop
        
        'Now we store the last sibling
        DataArray = DataArray & Separator & LastSibling.text
        If LastSibling.Children > 0 Then
            Add_Children_To_Data_Array DataArray, LastSibling
        End If
        'Add the Children End separator
        DataArray = DataArray & ChildrenEnd
    End If
End Sub

Private Function Save_Script_File(ByVal file_path As String) As Boolean
'On Error GoTo ErrHandler
    Dim LoopC As Long
    Dim fso As FileSystemObject
    Dim strm As TextStream
    Dim TempStr As String
    Dim TempLayer As Long
    Dim NodeLevel As Long
    
    'Script keywords
    Const Node = "#"
    Const EOF = "#EOF"
    Const Grh = ">"
    Const layer = "$"
    'The only existing combination is >$, which specify the layer is just for that grh
    
    'If file already exists delete it
    If General_File_Exists(file_path, vbNormal) Then
        Kill file_path
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set strm = fso.CreateTextFile(file_path)
    
    For LoopC = 1 To ScriptTree.Nodes.Count
        'Check what the node is
        If left$(ScriptTree.Nodes(LoopC).text, 4) = "Grh " Then
            strm.WriteLine Grh & Right(ScriptTree.Nodes(LoopC).text, Len(ScriptTree.Nodes(LoopC).text) - 4)
            GoTo CheckSiblings
        ElseIf left$(ScriptTree.Nodes(LoopC).text, 5) = "< On " Then
            TempStr = left$(ScriptTree.Nodes(LoopC).text, Len(ScriptTree.Nodes(LoopC).text) - 8)
            TempStr = Right(TempStr, Len(TempStr) - 5)
            'Check the layer index corresponding to it
            For TempLayer = 0 To LayerLst.ListCount
                If LayerLst.List(TempLayer) = TempStr Then
                    TempLayer = TempLayer + 1
                    Exit For
                End If
            Next TempLayer
            'Check if parent is a Grh
            If left$(ScriptTree.Nodes(LoopC).Parent.text, 4) = "Grh " Then
                strm.WriteLine Grh & layer & Str(TempLayer)
            Else
                strm.WriteLine layer & Str(TempLayer)
                GoTo CheckSiblings
            End If
        Else
            strm.WriteLine Node & ScriptTree.Nodes(LoopC).text
            NodeLevel = General_Field_Count(ScriptTree.Nodes(LoopC).FullPath, 92)
            
            'Check if it has children
            If ScriptTree.Nodes(LoopC).Children = 0 Then
CheckSiblings:
                'If it has no more brothers, we move back one node level (therefore we print #EOF)
                If ScriptTree.Nodes(LoopC).Key = ScriptTree.Nodes(LoopC).LastSibling.Key Then
                    'Node Level doesn´t decrease here, since it´s just the end of children (grhs and layers).
                    'Nevertheless, there might be an empty node, so we allow nodes to check
                    strm.WriteLine EOF
                End If
                
                'Returns as many levels as necessary
                If LoopC < ScriptTree.Nodes.Count Then
                    Dim NextNodeLevel As Long
                    NextNodeLevel = General_Field_Count(ScriptTree.Nodes(LoopC + 1).FullPath, 92)
                    If NextNodeLevel < NodeLevel Then
                        Do Until NextNodeLevel = NodeLevel
                            NodeLevel = NodeLevel - 1
                            strm.WriteLine EOF
                        Loop
                    End If
                End If
            End If
        End If
    Next LoopC
    
    Do Until NodeLevel = 0
        strm.WriteLine EOF
        NodeLevel = NodeLevel - 1
    Loop
    
    strm.Close
    
    Save_Script_File = True
Exit Function

Errhandler:
    Save_Script_File = False
End Function

Public Function Check_Node_Is_Children(Parent As Node, Child As Node) As Boolean
    Dim LoopC As Long
    
    If Parent.Children = 0 Then
        Check_Node_Is_Children = False
        Exit Function
    End If
    
    For LoopC = Parent.Child.index To Parent.index + Parent.Children
        If ScriptTree.Nodes(LoopC).Key = Child.Key Then
            Check_Node_Is_Children = True
            Exit Function
        End If
        If ScriptTree.Nodes(LoopC).Children > 0 Then
            Check_Node_Is_Children = Check_Node_Is_Children(ScriptTree.Nodes(LoopC), Child)
            If Check_Node_Is_Children Then Exit Function
        End If
    Next LoopC
End Function

Private Sub SetDefaultLayerCmd_Click()
    'Sets the default layer for a Grh or a group
    Dim LoopC As Long
    Dim SelectedNode As Node
    Dim SelectedNodeIndex As Long
    Dim NewNode As Node
    
    'Make sure we have a node selected
    If ScriptTree.Nodes.Count = 0 Then
        MsgBox "There are no nodes in the tree. You need a node to which to set the default layer.", , "Error"
        Exit Sub
    End If
    
    For LoopC = 1 To ScriptTree.Nodes.Count
        If ScriptTree.Nodes(LoopC).Selected Then
            SelectedNodeIndex = LoopC
            Exit For
        End If
    Next LoopC
    
    If SelectedNodeIndex = 0 Then
        MsgBox "There is no node selected. Choose the node to which you want to set the default layer.", , "Error"
        Exit Sub
    End If
    
    Set SelectedNode = ScriptTree.Nodes(SelectedNodeIndex)
    
    'Make sure selected node is not a layer node
    If left$(SelectedNode.Key, 5) = "Layer" Then Exit Sub
    
    'Make sure the selected node doen´t already have a default layer
    If SelectedNode.Children > 0 Then
        For LoopC = SelectedNode.index + 1 To SelectedNode.index + SelectedNode.Children
            If left$(ScriptTree.Nodes(LoopC).Key, 5) = "Layer" Then Exit Sub
        Next LoopC
    End If
    
    NodeCount = NodeCount + 1
    Set NewNode = ScriptTree.Nodes.Add(SelectedNode.Key, tvwChild, "Layer" & NodeCount, "< On " & LayerLst.text & " Layer >")
    NewNode.EnsureVisible
    
    Modified = True
    SaveMnu.Enabled = True
End Sub

Private Sub Delete_Node(Node As Node)
    Dim grh_index As Long
    Dim list_index As Long
    Dim LastSibling As Node
    
    'Check if it has children
    If Node.Children > 0 Then
        Delete_Node Node.Child
    End If
    
    'Check if it is a grh
    If left$(Node.text, 4) = "Grh " Then
        'Get it´s index
        grh_index = Val(Right(Node.text, Len(Node.text) - 4))
        'Find it´s place in the list
        For list_index = 0 To GrhList.ListCount
            If Val(GrhList.List(list_index)) > grh_index Then
                GrhList.AddItem grh_index, list_index
                Exit For
            End If
        Next list_index
    End If
    
    'Check for siblings
    Do Until Node.Key = Node.LastSibling.Key
        Delete_Node Node.Next
    Loop
    
    'Remove node
    ScriptTree.Nodes.Remove Node.index
End Sub
