Attribute VB_Name = "General"
'*****************************************************************
'ORE 1.0 Map Editor - v0.7.0
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
' 10/12/2004 - Juan Martín Sotuyo Dodero (juansotuyo@hotmailcom)
'   -Add: Map Editor now can use resource files
'   -Add: a few extra controls
'   -Fix: some minor bugs
'       David Justus (big.david@txun.net)
'          - Add: Minimap
'       Fredrik Alexandersson (fredrik_alexandersson@hotmail.com)
'           - Add: frmNewNPC and frmNewItem (still not functional)
'
' 2/15/2004 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Change: GUI was changed a bit
'   -Fix: several bugs
'
' 12/9/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: controls for decorations
'   -Add: Map Go To.. And Map Resize controls
'   -Add: Speed and Life to particle groups
'   -Change: Tile Groups.dat uses the new script system
'   -Change: Tile Groups GUI a bit.
'   -Change: There are 3 layers now, and a decorative layer which can be rendered at 2 different levels
'   -Change: Grhs can be aligned either vertically, horizontally or both
'   -Change: FireStarter's scripting system to allow unlimited numbers of nested nodes
'   -Fix: Several bugs in the Particle Editor
'   Sub Release Contributors:
'       12/2/2003 - Murat Sütunç (Firestarter)
'           -Add: Scripting system for grh script.dat
'
' 11/16/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: the view pos is reset after loading / creating a map
'   -Change: Center Grh option alligns both vertically and horizontally
'
' 10/28/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: Tile Groups now has a grid
'   -Add: the ability to center or not Grhs on layers 2-4.
'   -Add: Undo - Redo buttons to the toolbar.
'   -Add: Tile Groups (works as old Grh Groups, but only takes grhs the same size the tiles are).
'   -Change: Tile Groups GUI a bit.
'   -Change: map scroll now uses the keyboard. Now keys are A, S, D and W (combine them to get diagonals).
'   -Change: Grh Groups. Now grhs can be different sizes. Nevertheless, only 1 grh can be displayed at a time.
'   -Fix: Lists can now be scrolled using the arrow keys with no problem.
'
' 8/28/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: undo - redo
'   -Change: the way to select Light/shadow/base_lights
'   -Change: the way to select either to erase 1 or all layers
'   -Change: Adjacent map now allows to set no map in any side
'   -Fix: several minor bugs
'
' 7/4/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   -Add: Included the Particle Editor itself.
'   -Change: Exits can now lead to the same map without having to save and restart the Map Editor.
'   -Change: the GUI a little bit.
'
'
' 6/27/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   First Release

Option Explicit

Sub Main()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last modified: 10/12/2004
'Main
'*************************************************
    frmLoading.Show
    
    'Init user-defined stuff
    Load_User_Defined_Data
    
    'Are we using resource files?
    If use_resource_files Then
        'Load graphic resources
        Extract_Files grh, resource_path, Nothing, Nothing, frmLoading.LoadingLbl
    End If
    
    'Start up Engine
    frmLoading.LoadingLbl.Caption = "Initializing Engine..."
    'windowed
    prgRun = Engine.Engine_Initialize(frmMain.hwnd, frmMain.MainView.hwnd, 1, resource_path, , , , , 17, 13, tile_size, use_resource_files)
    
    Engine.Engine_Base_Speed_Set base_speed
    Engine.Engine_View_Pos_Set 5, 5
    Engine.Engine_Special_Tiles_Show_Toggle
    Engine.Engine_Blocked_Tiles_Show_Toggle
    
    'Initialize the Light_Color (white)
    Light_Color = &HFFFFFF
    
    'Load .dat and .ini files
    frmLoading.LoadingLbl.Caption = "Loading items..."
    Load_Items_Data frmMain.OBJList
    
    frmLoading.LoadingLbl.Caption = "Loading NPCs..."
    Load_NPC_Data frmMain.NPCList
    
    frmLoading.LoadingLbl.Caption = "Loading Grh Script..."
    Load_Grh_Tree frmMain.tree, App.Path & "\Grh Script.dat"
    
    'Load Tile tree
    frmLoading.LoadingLbl.Caption = "Loading Tile Script..."
    Load_Grh_Tree frmTileGroups.tree, App.Path & "\Tile Script.dat"
    
    frmLoading.LoadingLbl.Caption = "Loading Triggers..."
    Load_Triggers_Data_To_List frmMain.TriggerList
    
    frmLoading.LoadingLbl.Caption = "Loading Particle Streams..."
    Load_Particle_Streams_To_ComboBox frmMain.ParticleType
    
    'Load all maps names to the exit´s map list
    Load_Maps_To_ComboBox frmMain.ExitMapsList
    
    Modified = False
    
    'Set the Grh tool as default
    Dim Button As Button
    Set Button = frmMain.Toolbar1.Buttons(8)
    frmMain.toolbar1_ButtonClick Button
    
    'Hide frmLoading
    frmLoading.Hide
    Unload frmLoading
    
    'Show the GrhViewer
    frmGrhViewer.Show
    
    'Show the Mini Map
    frmMap.Show
    
    'Show Main frame
    frmMain.Show
    
    'Show Grh1
    Engine.Grh_Render_To_Hdc 1, frmGrhViewer.hdc, 0, 0
    
    'Let window think
    DoEvents
    
    'Start AutoSaver timer
    frmMain.AutoSaveTimer.Enabled = True
    
MainLoop:
    Do While prgRun
    
        If frmMain.WindowState <> 1 Then
            '********* Render **********
            prgRun = Engine.Engine_Render_Start
            prgRun = Engine.Engine_Render_End
            
            '********* Check Keys *********
            Check_Keys
        
            '********* Walk with mouse *********
            If Walk_Mode Then
                Mouse_Walk
            End If
        End If
        
        '********* Go do other events *********
        DoEvents
    Loop
    
    '*******************
    'Close Down
    '*******************
    'Unload engine
    Engine.Engine_DeInitialize
    
    'Unlock and delete all graphics
    Close_Handler_List GRH_Handles, True
    Kill resource_path & "\graphics\*.bmp"
    
    'Destroy the Engine object
    Set Engine = Nothing
    
    '********* Unload forms and end *********
    Unload frmMap
    Unload frmGrhViewer
    Unload frmTileGroups
    Unload frmMain
    End
End Sub

Sub ArrangeDialog(ByRef target As Object, ByVal Action As Single)
'*************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last modified: 2/18/2003
'Arranges the data for the common dialog control
'*************************************************
    
    With target
        Select Case Action
            'Load
            Case 1
                .Filter = "Maps|*.map"
                .DialogTitle = "Load"
                .FileName = ""
                .InitDir = resource_path & "\maps\"
                .flags = cdlOFNFileMustExist
                .ShowOpen
            
            'Save
            Case 2
                .Filter = "Maps|*.map"
                .DialogTitle = "Save"
                .DefaultExt = ".map"
                .FileName = ""
                .InitDir = resource_path & "\maps\"
                .flags = cdlOFNPathMustExist
                .ShowSave
            'Color, complete (unfolded)
            Case 3
                .flags = cdlCCRGBInit Or cdlCCFullOpen
                .ShowColor
        End Select
    End With
    
End Sub

Sub Check_Keys()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last modified: 10/18/2003
'Checks keys
'*************************************************
    Dim X As Long, Y As Long
    Engine.Engine_View_Pos_Get X, Y
    X = X - 9
    Y = Y - 7
    If Engine.Input_Key_Get(vbKeyW) And Engine.Input_Key_Get(vbKeyD) Then
            frmMap.shparea.top = Y
            frmMap.shparea.left = X
       If Walk_Mode Then
            If Engine.Map_Legal_Char_Pos_By_Heading(User_Char_Index, 2) Then
                If Engine.Engine_View_Move(2) Then
                    Engine.Char_Move User_Char_Index, 2
                    Engine.Light_Move_By_Head Cursor_Light_Index, 2
                End If
            Else
                Engine.Char_Heading_Set User_Char_Index, 2
            End If
        Else
            Engine.Engine_View_Move 2
        End If
        frmMain.Statusbar_Update
    End If
    
    If Engine.Input_Key_Get(vbKeyD) And Engine.Input_Key_Get(vbKeyS) Then
                    frmMap.shparea.top = Y
            frmMap.shparea.left = X
        If Walk_Mode Then
            If Engine.Map_Legal_Char_Pos_By_Heading(User_Char_Index, 4) Then
                If Engine.Engine_View_Move(4) Then
                    Engine.Char_Move User_Char_Index, 4
                    Engine.Light_Move_By_Head Cursor_Light_Index, 4
                End If
            Else
                Engine.Char_Heading_Set User_Char_Index, 4
            End If
        Else
            Engine.Engine_View_Move 4
            
        End If
        frmMain.Statusbar_Update
    End If
    
    If Engine.Input_Key_Get(vbKeyS) And Engine.Input_Key_Get(vbKeyA) Then
                    frmMap.shparea.top = Y
            frmMap.shparea.left = X
        If Walk_Mode Then
            If Engine.Map_Legal_Char_Pos_By_Heading(User_Char_Index, 6) Then
                If Engine.Engine_View_Move(6) Then
                    Engine.Char_Move User_Char_Index, 6
                    Engine.Light_Move_By_Head Cursor_Light_Index, 6
                End If
            Else
                Engine.Char_Heading_Set User_Char_Index, 6
            End If
        Else
            Engine.Engine_View_Move 6
           
        End If
        frmMain.Statusbar_Update
    End If
    
    If Engine.Input_Key_Get(vbKeyA) And Engine.Input_Key_Get(vbKeyW) Then
                    frmMap.shparea.top = Y
            frmMap.shparea.left = X
        If Walk_Mode Then
            If Engine.Map_Legal_Char_Pos_By_Heading(User_Char_Index, 8) Then
                If Engine.Engine_View_Move(8) Then
                    Engine.Char_Move User_Char_Index, 8
                    Engine.Light_Move_By_Head Cursor_Light_Index, 8
                End If
            Else
                Engine.Char_Heading_Set User_Char_Index, 8
            End If
        Else
            Engine.Engine_View_Move 8
            frmMap.shparea.top = frmMap.shparea.top + 1
            frmMap.shparea.left = frmMap.shparea.left - 1
        End If
        frmMain.Statusbar_Update
    End If
    
    If Engine.Input_Key_Get(vbKeyW) Then
                    frmMap.shparea.top = Y
            frmMap.shparea.left = X
        If Walk_Mode Then
            If Engine.Map_Legal_Char_Pos_By_Heading(User_Char_Index, 1) Then
                If Engine.Engine_View_Move(1) Then
                    Engine.Char_Move User_Char_Index, 1
                    Engine.Light_Move_By_Head Cursor_Light_Index, 1
                End If
            Else
                Engine.Char_Heading_Set User_Char_Index, 1
            End If
        Else
            Engine.Engine_View_Move 1
            
        End If
        frmMain.Statusbar_Update
    End If
    
    If Engine.Input_Key_Get(vbKeyD) Then
            frmMap.shparea.top = Y
            frmMap.shparea.left = X
        If Walk_Mode Then
            If Engine.Map_Legal_Char_Pos_By_Heading(User_Char_Index, 3) Then
                If Engine.Engine_View_Move(3) Then
                    Engine.Char_Move User_Char_Index, 3
                    Engine.Light_Move_By_Head Cursor_Light_Index, 3
                End If
            Else
                Engine.Char_Heading_Set User_Char_Index, 3
               
            End If
        Else
            Engine.Engine_View_Move 3
        End If
        frmMain.Statusbar_Update
    End If
    
    If Engine.Input_Key_Get(vbKeyS) Then
            frmMap.shparea.top = Y
            frmMap.shparea.left = X
        If Walk_Mode Then
            If Engine.Map_Legal_Char_Pos_By_Heading(User_Char_Index, 5) Then
                If Engine.Engine_View_Move(5) Then
                    Engine.Char_Move User_Char_Index, 5
                    Engine.Light_Move_By_Head Cursor_Light_Index, 5
                End If
            Else
                Engine.Char_Heading_Set User_Char_Index, 5
                
            End If
        Else
            Engine.Engine_View_Move 5
        End If
        frmMain.Statusbar_Update
    End If
    
    If Engine.Input_Key_Get(vbKeyA) Then
        frmMap.shparea.top = Y
        frmMap.shparea.left = X
        If Walk_Mode Then
            If Engine.Map_Legal_Char_Pos_By_Heading(User_Char_Index, 7) Then
                If Engine.Engine_View_Move(7) Then
                    Engine.Char_Move User_Char_Index, 7
                    Engine.Light_Move_By_Head Cursor_Light_Index, 7
                End If
            Else
                Engine.Char_Heading_Set User_Char_Index, 7
            End If
        Else
            Engine.Engine_View_Move 7
            
        End If
        frmMain.Statusbar_Update
    End If
    
    'Set tool
    Dim Button As Button
    'Set grh tool
    If Engine.Input_Key_Get(vbKeyF1) Then
        Set Button = frmMain.Toolbar1.Buttons(8)
        frmMain.toolbar1_ButtonClick Button
    End If
    
    'Set tiles tool
    If Engine.Input_Key_Get(vbKeyF2) Then
        Set Button = frmMain.Toolbar1.Buttons(9)
        frmMain.toolbar1_ButtonClick Button
    End If
    
    'Set lights tool
    If Engine.Input_Key_Get(vbKeyF3) Then
        Set Button = frmMain.Toolbar1.Buttons(10)
        frmMain.toolbar1_ButtonClick Button
    End If
    
    'Set particle groups tool
    If Engine.Input_Key_Get(vbKeyF4) Then
        Set Button = frmMain.Toolbar1.Buttons(11)
        frmMain.toolbar1_ButtonClick Button
    End If
    
    'Set exits tool
    If Engine.Input_Key_Get(vbKeyF5) Then
        Set Button = frmMain.Toolbar1.Buttons(12)
        frmMain.toolbar1_ButtonClick Button
    End If
    
    'Set OBJs tool
    If Engine.Input_Key_Get(vbKeyF6) Then
        Set Button = frmMain.Toolbar1.Buttons(13)
        frmMain.toolbar1_ButtonClick Button
    End If
    
    'Set NPCs tool
    If Engine.Input_Key_Get(vbKeyF7) Then
        Set Button = frmMain.Toolbar1.Buttons(14)
        frmMain.toolbar1_ButtonClick Button
    End If

End Sub

Sub Mouse_Walk()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last modified: 5/18/2003
'Moves char using mouse
'*************************************************
    
    'Make sure the mouse is in the view area
    If Engine.Input_Mouse_In_View Then
    
        'Get position in the map
        Dim temp_x As Long
        Dim temp_y As Long
        Engine.Input_Mouse_Map_Get temp_x, temp_y
    
        'Check the mouse for movement
        If Engine.Input_Mouse_Moved_Get Then
            'Move the light over the cursor
            Engine.Light_Move Cursor_Light_Index, temp_x, temp_y
        End If
        
        'Check left button
        If Engine.Input_Mouse_Button_Left_Get Then
            'Check it´s a legal pos
            If Engine.Map_Legal_Char_Pos_By_Heading(User_Char_Index, Engine.Input_Mouse_Heading_Get) Then
                'Move the view position, the user_char and the cursor_light
                If Engine.Engine_View_Move(Engine.Input_Mouse_Heading_Get) Then
                    Engine.Char_Move User_Char_Index, Engine.Input_Mouse_Heading_Get
                    Engine.Light_Move_By_Head Cursor_Light_Index, Engine.Input_Mouse_Heading_Get
                End If
            Else
                'Turn in the place
                Engine.Char_Heading_Set User_Char_Index, Engine.Input_Mouse_Heading_Get
            End If
            frmMain.Statusbar_Update
        End If
    End If
    
End Sub

Sub Mouse_React_to_Click()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last modified: 10/13/2004
'Reacts to mouse click
'*************************************************
    Dim X As Long
    Dim Y As Long
    Dim LoopC As Long
    
    'Get map pos
    Engine.Input_Mouse_Map_Get X, Y
    
    
    ''************ Left click
    If Engine.Input_Mouse_Button_Left_Get Then
        Select Case tool
            '********** Grh Tool **********
            Case Is = "grh"
                'Pick Grh
                If frmMain.GrhPickCmd.Enabled = False Then
                    'Find the Grh index
                    Dim index As Long
                    If frmMain.GrhLayerList.ListIndex = 0 Or frmMain.GrhLayerList.ListIndex = 2 Or frmMain.GrhLayerList.ListIndex = 4 Then
                        index = Engine.Map_Grh_Get(X, Y, frmMain.GrhLayerList.ListIndex / 2 + 1)
                    ElseIf frmMain.GrhLayerList.ListIndex = 1 Or frmMain.GrhLayerList.ListIndex = 3 Then
                        index = Engine.Map_Decoration_Get(X, Y, Engine.Input_Mouse_Subtile_Get)
                    End If
                    If index = 0 Then
                        Exit Sub
                    End If
                    current_grh = index
                    If frmMain.GrhViewerMnuChk.Checked Then
                        frmGrhViewer.Cls
                        Engine.Grh_Render_To_Hdc current_grh, frmGrhViewer.hdc, 0, 0
                    End If
                    For LoopC = 1 To frmMain.tree.Nodes.count
                        If frmMain.tree.Nodes(LoopC).text = "Grh " & current_grh Then
                            frmMain.tree.Nodes(LoopC).Selected = True
                        Else
                            frmMain.tree.Nodes(LoopC).Selected = False
                        End If
                    Next LoopC
                    Exit Sub
                End If
                'Place a Grh
                If frmMain.GrhLayerList.ListIndex = 0 Or frmMain.GrhLayerList.ListIndex = 2 Or frmMain.GrhLayerList.ListIndex = 4 Then
                    If Engine.Map_Grh_Set(X, Y, current_grh, (frmMain.GrhLayerList.ListIndex + 2) / 2, frmMain.GrhAlphaBlendingChk.value, Val(frmMain.GrhAngleTxt.text) * PI / 180, frmMain.GrhHCenteredChk.value, frmMain.GrhVCenteredChk.value) Then
                        store_action grh_map, place, X, Y, , frmMain.GrhLayerList.ListIndex + 1
                        'Update mini map
                        If (frmMain.GrhLayerList.ListIndex = 0 Or frmMain.GrhLayerList.ListIndex = 2) And frmMap.Visible Then
                            frmMap.Cls
                            Engine.Engine_Render_Mini_Map_To_hDC frmMap.hdc
                        End If
                        Modified = True
                        Exit Sub
                    End If
                'Place a decoration
                ElseIf frmMain.GrhLayerList.ListIndex = 1 Or frmMain.GrhLayerList.ListIndex = 3 Then
                    Dim Render_On_Top As Boolean
                    If frmMain.GrhLayerList.ListIndex = 1 Then
                        Render_On_Top = False
                    Else
                        Render_On_Top = True
                    End If
                    If Engine.Map_Decoration_Add(X, Y, frmMain.DecorationPositionLst.ListIndex, current_grh, Render_On_Top, frmMain.GrhAlphaBlendingChk.value, Val(frmMain.GrhAngleTxt.text), frmMain.GrhHCenteredChk.value, frmMain.GrhVCenteredChk.value) Then
                        store_action grh_map, place, X, Y, , frmMain.GrhLayerList.ListIndex + 1, , , , , , , , , , , frmMain.DecorationPositionLst.ListIndex
                        Modified = True
                        Exit Sub
                    End If
                End If
            '********** Blocking Tool **********
            Case Is = "tiles"
                'Block a tile
                If frmMain.TileToolChk(0).value = True Then
                    'Check if it wasn´t already blocked
                    If Engine.Map_Blocked_Get(X, Y) Then
                        Exit Sub
                    End If
                    store_action blocking, place, X, Y
                    If Engine.Map_Blocked_Set(X, Y, True) Then
                        Modified = True
                        Exit Sub
                    End If
                Else
                    store_action trigger, place, X, Y
                    If Engine.Map_Trigger_Set(X, Y, frmMain.TriggerList.ListIndex) Then
                        Modified = True
                        Exit Sub
                    End If
                End If
            '********** Lights Tool **********
            Case Is = "lights"
                'Pick Light Color
                If frmMain.PickColorCmd.Enabled = False Then
                    Engine.Light_Color_Value_Get Engine.Map_Light_Get(X, Y), Light_Color
                    frmMain.arrange_light_color Light_Color
                    Exit Sub
                End If
                'Pick Base Light Color
                If frmMain.PickBaseLightColorCmd.Enabled = False Then
                    Light_Color = Engine.Map_Base_Light_Get(X, Y)
                    frmMain.arrange_light_color Light_Color
                    Exit Sub
                End If
                'Place a Base Light
                If frmMain.LightToolChk(1).value = True Then
                    store_action lights, place, X, Y, , , Light_Color, , , base_light
                    Engine.Map_Base_Light_Set X, Y, Light_Color
                    Modified = True
                    Exit Sub
                End If
                'Place a Shadow
                If frmMain.LightToolChk(2).value = True Then
                    Dim counter As Long
                    For LoopC = 0 To 3
                        If frmMain.CornerChk(LoopC).value = vbChecked Then
                            counter = counter + 1
                            store_action lights, place, X, Y, counter, , , , LoopC, shadow
                            Engine.Map_Base_Light_Set X, Y, Light_Color, LoopC
                        End If
                    Next LoopC
                    Modified = True
                    Exit Sub
                End If
                'Place a Light
                Engine.Light_Create X, Y, Light_Color, Val(frmMain.Rangetxt.text)
                store_action lights, place, X, Y, , , Light_Color, Val(frmMain.Rangetxt.text), , Light
                Modified = True
                Exit Sub
            '********** Particle Groups Tools **********
            Case Is = "particle_groups"
                Dim PIndex As Long
                If Not Engine.Map_In_Bounds(X, Y) Then Exit Sub
                
                PIndex = frmMain.ParticleType.ListIndex + 1
                
                If frmMain.ParticleType.ListIndex >= 0 Then
                    
                    Dim rgb_list(0 To 3) As Long
                    rgb_list(0) = RGB(StreamData(PIndex).colortint(0).r, StreamData(PIndex).colortint(0).g, StreamData(PIndex).colortint(0).b)
                    rgb_list(1) = RGB(StreamData(PIndex).colortint(1).r, StreamData(PIndex).colortint(1).g, StreamData(PIndex).colortint(1).b)
                    rgb_list(2) = RGB(StreamData(PIndex).colortint(2).r, StreamData(PIndex).colortint(2).g, StreamData(PIndex).colortint(2).b)
                    rgb_list(3) = RGB(StreamData(PIndex).colortint(3).r, StreamData(PIndex).colortint(3).g, StreamData(PIndex).colortint(3).b)
                    
                    Engine.Particle_Group_Create X, Y, StreamData(PIndex).grh_list, rgb_list(), StreamData(PIndex).NumOfParticles, frmMain.ParticleType.ListIndex + 1, _
                                StreamData(PIndex).AlphaBlend, StreamData(PIndex).life_counter, StreamData(PIndex).speed, , StreamData(PIndex).x1, StreamData(PIndex).y1, StreamData(PIndex).angle, _
                                StreamData(PIndex).vecx1, StreamData(PIndex).vecx2, StreamData(PIndex).vecy1, StreamData(PIndex).vecy2, _
                                StreamData(PIndex).life1, StreamData(PIndex).life2, StreamData(PIndex).friction, StreamData(PIndex).spin_speedL, _
                                StreamData(PIndex).gravity, StreamData(PIndex).grav_strength, StreamData(PIndex).bounce_strength, StreamData(PIndex).x2, _
                                StreamData(PIndex).y2, StreamData(PIndex).XMove, StreamData(PIndex).move_x1, StreamData(PIndex).move_x2, StreamData(PIndex).move_y1, _
                                StreamData(PIndex).move_y2, StreamData(PIndex).YMove, StreamData(PIndex).spin_speedH, StreamData(PIndex).spin
                    
                    store_action particle_stream, place, X, Y
                    
                    Modified = True
                    
                    Exit Sub
                End If
            '********* Exit Tools **********
            Case Is = "exits"
                'Pick data from an existing exit
                If Not Engine.Map_In_Bounds(X, Y) Then Exit Sub
                
                If frmMain.ExitPickCmd.Enabled = False Then
                    Dim map_name As String
                    Dim map_x As Long
                    Dim map_y As Long
                    Engine.Map_Exit_Get X, Y, map_name, map_x, map_y
                    
                    'Make sure there is an exit in the tile
                    If map_name = "" Then
                        frmMain.ExitMapsList.ListIndex = 0
                        frmMain.ExitXCoordTxt.text = "0"
                        frmMain.ExitYCoordTxt.text = "0"
                    Else
                        Do Until frmMain.ExitMapsList.List(LoopC) = map_name
                            LoopC = LoopC + 1
                        Loop
                        
                        frmMain.ExitMapsList.ListIndex = LoopC
                        frmMain.ExitXCoordTxt.text = map_x
                        frmMain.ExitYCoordTxt.text = map_y
                    End If
                    Exit Sub
                End If
                'Place exit
                If frmMain.ExitMapsList.text = "None" Then Exit Sub
                store_action exits, place, X, Y
                Engine.Map_Exit_Add X, Y, frmMain.ExitMapsList.text, Val(frmMain.ExitXCoordTxt.text), Val(frmMain.ExitYCoordTxt.text)
                Modified = True
                Exit Sub
            '********* Object Tools **********
            Case Is = "OBJs"
                'Place OBJ
                If Not Engine.Map_In_Bounds(X, Y) Then Exit Sub
                
                store_action object, place, X, Y
                Engine.Map_Item_Add X, Y, frmMain.OBJList.ListIndex + 1, Val(frmMain.OBJAmountTxt.text)
                Modified = True
                Exit Sub
            '********* NPCs Tools **********
            Case Is = "NPCs"
                'Place NPC
                If Not Engine.Map_In_Bounds(X, Y) Then Exit Sub
                
                store_action NPC, place, X, Y
                Engine.Map_NPC_Add X, Y, frmMain.NPCList.ListIndex + 1
                Modified = True
                Exit Sub
        End Select
    End If
    
    
    '************ Right click
    If Engine.Input_Mouse_Button_Right_Get Then
    Engine.Engine_Render_Mini_Map_To_hDC frmMap.picmain.hdc
        Select Case tool
            '********** Grh Tool **********
            Case Is = "grh"
                'Erase Layers
                If frmMain.GrhErase(1).value = True Then
                    'Erase layers 1 - 3 and both decoration layers
                    store_action grh_map, Remove_all, X, Y
                    For LoopC = 1 To 3
                        If Engine.Map_Grh_UnSet(X, Y, LoopC) Then
                            Modified = True
                        End If
                    Next LoopC
                    For LoopC = 0 To 8
                        If Engine.Map_Decoration_Remove(X, Y, LoopC) Then
                            Modified = True
                        End If
                    Next LoopC
                    Exit Sub
                Else
                    'Erase 1 layer
                    If frmMain.DecorationPositionLst.Enabled And frmMain.DecorationPositionLst.ListIndex > -1 Then
                        store_action grh_map, Remove, X, Y, , frmMain.GrhLayerList.ListIndex + 1, , , , , , , , , , , frmMain.DecorationPositionLst.ListIndex
                    Else
                        store_action grh_map, Remove, X, Y, , frmMain.GrhLayerList.ListIndex + 1
                    End If
                    If frmMain.GrhLayerList.ListIndex = 0 Or frmMain.GrhLayerList.ListIndex = 2 Or frmMain.GrhLayerList.ListIndex = 4 Then
                        If Engine.Map_Grh_UnSet(X, Y, frmMain.GrhLayerList.ListIndex / 2 + 1) Then
                            Modified = True
                        End If
                    Else
                        If Engine.Map_Decoration_Remove(X, Y, Engine.Input_Mouse_Subtile_Get) Then
                            Modified = True
                        End If
                    End If
                    Exit Sub
                End If
            '********* Blocking Tool **********
            Case Is = "tiles"
                If frmMain.TileToolChk(0).value = True Then
                    If Not Engine.Map_Blocked_Get(X, Y) Then
                        Exit Sub
                    End If
                    store_action blocking, Remove, X, Y
                    If Engine.Map_Blocked_Set(X, Y, False) Then
                        Modified = True
                        Exit Sub
                    End If
                Else
                    If Engine.Map_Trigger_Get(X, Y) = 0 Then
                        Exit Sub
                    End If
                    store_action trigger, Remove, X, Y
                    If Engine.Map_Trigger_Unset(X, Y) Then
                        Modified = True
                        Exit Sub
                    End If
                End If
            '********* Lights Tool **********
            Case Is = "lights"
                If Engine.Map_Light_Get(X, Y) = 0 Then
                    Exit Sub
                End If
                store_action lights, Remove, X, Y, , , , , , Light
                If Engine.light_remove(Engine.Map_Light_Get(X, Y)) Then
                    Modified = True
                    Exit Sub
                End If
            '********* Particle Groups Tools **********
            Case Is = "particle_groups"
                store_action particle_stream, Remove, X, Y
                If Engine.Particle_Group_Remove(Engine.Map_Particle_Group_Get(X, Y)) Then
                    Modified = True
                    Exit Sub
                End If
            '********* Exits **********
            Case Is = "exits"
                store_action exits, Remove, X, Y
                If Engine.Map_Exit_Remove(X, Y) Then
                    Modified = True
                    Exit Sub
                End If
            '********* OBJs **********
            Case Is = "OBJs"
                store_action object, Remove, X, Y
                If Engine.Map_Item_Remove(X, Y) Then
                    Modified = True
                    Exit Sub
                End If
            '********* NPCs **********
            Case Is = "NPCs"
                store_action NPC, Remove, X, Y
                If Engine.Map_NPC_Remove(X, Y) Then
                    Modified = True
                    Exit Sub
                End If
        End Select
    End If

End Sub

Public Sub Toggle_Walk_Mode()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last modified: 2/23/2003
'Toggles Wal_Mode on / off
'*************************************************
    If Walk_Mode = False Then
        Walk_Mode = True
    Else
        Walk_Mode = False
    End If
    
    If Not Walk_Mode Then
        'Erase character
        Engine.Char_Remove User_Char_Index
        'Erase light
        Engine.light_remove Cursor_Light_Index
    Else
        Dim X As Long
        Dim Y As Long
        
        Engine.Engine_View_Pos_Get X, Y
    
        'Make Character and cursor light
        If Not Engine.Map_Blocked_Get(X, Y) Then
            User_Char_Index = Engine.Char_Create(X, Y, 5, 1)
            Cursor_Light_Index = Engine.Light_Create(X, Y, RGB(255, 255, 255))
            'Char label
            Engine.Char_Label_Set User_Char_Index, char_label, 1
        Else
            MsgBox "Error: Must move to a free tile first."
            frmMain.WalkModeChk.value = 0
            Walk_Mode = False
        End If
    End If
    
End Sub

Public Sub Load_User_Defined_Data()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last modified: 10/12/2004
'Loads all user-defined stuff from the Map Editor.ini file
'*************************************************
    
    resource_path = App.Path & General_Var_Get(App.Path & "\Map Editor.ini", "SYSTEM", "ResourcePath")
    autosave_delay = Val(General_Var_Get(App.Path & "\Map Editor.ini", "SYSTEM", "AutoSaveDelay"))
    use_ini_files = CBool(General_Var_Get(App.Path & "\Map Editor.ini", "SYSTEM", "UseIniFiles"))
    use_resource_files = CBool(General_Var_Get(App.Path & "\Map Editor.ini", "SYSTEM", "UseResFiles"))
    tile_size = Val(General_Var_Get(App.Path & "\Map Editor.ini", "PREFERENCES", "TileSize"))
    map_height = Val(General_Var_Get(App.Path & "\Map Editor.ini", "PREFERENCES", "MapHeight"))
    map_width = Val(General_Var_Get(App.Path & "\Map Editor.ini", "PREFERENCES", "MapWidth"))
    x_border = Val(General_Var_Get(App.Path & "\Map Editor.ini", "PREFERENCES", "MapBorderX"))
    y_border = Val(General_Var_Get(App.Path & "\Map Editor.ini", "PREFERENCES", "MapBorderY"))
    base_speed = Val(General_Var_Get(App.Path & "\Map Editor.ini", "PREFERENCES", "EngineSpeed"))
    char_label = General_Var_Get(App.Path & "\Map Editor.ini", "PREFERENCES", "CharName")

End Sub

Public Sub Save_User_Defined_Data()
'*************************************************
'Author: Juan Martín Sotuyo Dodero
'Last modified: 10/12/2004
'Saves all user-defined stuff to the Map Editor.ini file
'*************************************************
    
    General_Var_Write App.Path & "\Map Editor.ini", "SYSTEM", "ResourcePath", resource_path
    General_Var_Write App.Path & "\Map Editor.ini", "SYSTEM", "AutoSaveDelay", autosave_delay
    General_Var_Write App.Path & "\Map Editor.ini", "SYSTEM", "UseIniFiles", CLng(use_ini_files)
    General_Var_Write App.Path & "\Map Editor.ini", "SYSTEM", "UseResFiles", CLng(use_resource_files)
    General_Var_Write App.Path & "\Map Editor.ini", "PREFERENCES", "TileSize", tile_size
    General_Var_Write App.Path & "\Map Editor.ini", "PREFERENCES", "MapHeight", map_height
    General_Var_Write App.Path & "\Map Editor.ini", "PREFERENCES", "MapWidth", map_width
    General_Var_Write App.Path & "\Map Editor.ini", "PREFERENCES", "MapBorderX", x_border
    General_Var_Write App.Path & "\Map Editor.ini", "PREFERENCES", "MapBorderY", y_border
    General_Var_Write App.Path & "\Map Editor.ini", "PREFERENCES", "EngineSpeed", Str(base_speed)
    General_Var_Write App.Path & "\Map Editor.ini", "PREFERENCES", "CharName", char_label

End Sub

Public Sub Load_Maps_To_ComboBox(ByRef ComboBoxName As ComboBox)
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 6/24/2003
'Loads all map names to the given ComboBox
'*************************************
    Dim path_name As String
    Dim map_name As String
    
    ComboBoxName.Clear
    
    path_name = resource_path & "\maps\"
    
    ComboBoxName.AddItem "None"
    
    'check there are maps
    map_name = Dir(resource_path & "\maps\*.map", vbNormal)
    If map_name = "" Then
        Exit Sub
    End If
    
    'Load all maps
    Do While map_name <> ""
        'Add it without the .map extension
        ComboBoxName.AddItem left$(map_name, Len(map_name) - 4)
        'Load next map
        map_name = Dir
    Loop
    
    'See if current map already exists, or add it to the map
    If Current_Map = "" Then
        ComboBoxName.AddItem "Current Map"
    End If
    
    'Select first map
    ComboBoxName.ListIndex = 0
End Sub

Public Sub Load_Items_Data(ByRef ListBoxName As ListBox)
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/20/2003
'Loads items.ini to the given ListBox
'*************************************
    Dim NumItems As Long
    Dim Item As Long
    
    NumItems = Val(General_Var_Get(App.Path & resource_path & "\scripts\item.ini", "GENERAL", "item_count"))
    
    For Item = 1 To NumItems
        ListBoxName.AddItem General_Var_Get(App.Path & resource_path & "\scripts\item.ini", "ITEM" & Item, "item_name")
    Next Item
    
End Sub

Public Sub Load_NPC_Data(ByRef ListBoxName As ListBox)
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/20/2003
'Loads NPC.ini to the given ListBox
'*************************************
    Dim NumNPCs As Long
    Dim NPC As Long
    
    NumNPCs = Val(General_Var_Get(resource_path & "\scripts\npc.ini", "GENERAL", "npc_count"))
    For NPC = 1 To NumNPCs
        ListBoxName.AddItem General_Var_Get(resource_path & "\scripts\npc.ini", "NPC" & NPC, "npc_name")
    Next NPC
    
End Sub

Public Function Item_Get_Index_From_Name(ByVal item_name As String) As Long
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/22/2003
'Gets the item index based on the given name
'*************************************
    Dim NumItems As Long
    Dim LoopC As Long
    
    NumItems = Val(General_Var_Get(resource_path & "\scripts\item.ini", "GENERAL", "item_count"))
    
    For LoopC = 1 To NumItems
        If item_name = General_Var_Get(resource_path & "\scripts\item.ini", "ITEM" & LoopC, "item_name") Then
            Item_Get_Index_From_Name = LoopC
            Exit Function
        End If
    Next LoopC
    
End Function

Public Function NPC_Get_Index_From_Name(ByVal npc_name As String) As Long
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/22/2003
'Gets the npc index based on the given name
'*************************************
    Dim NumNPCs As Long
    Dim LoopC As Long
    
    NumNPCs = Val(General_Var_Get(resource_path & "\scripts\npc.ini", "GENERAL", "npc_count"))
    
    For LoopC = 1 To NumNPCs
        If npc_name = General_Var_Get(resource_path & "\scripts\npc.ini", "NPC" & LoopC, "npc_name") Then
            NPC_Get_Index_From_Name = LoopC
            Exit Function
        End If
    Next LoopC
    
End Function

Public Sub Draw_Tile_Group(ByVal Draw_Grid As Boolean, ByVal Grid_Color As Long)
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 10/18/2003
'Draws a Grh Group
'*************************************
    Dim Grhs_per_Row As Byte
    Dim Row As Long
    Dim Column As Long
    Dim LoopC As Long
    
    'Check there´s is a selected group
    If Current_Group.Name = "" Then
        Exit Sub
    End If
    
    'Calculate how many Grhs fit in a row
    Grhs_per_Row = frmTileGroups.TileGroupViewer.ScaleWidth / tile_size
    
    'Erase viewer
    frmTileGroups.TileGroupViewer.Cls
    
    For LoopC = Grhs_per_Row * TileGroupOffset + 1 To UBound(Current_Group.GrhIndexes)
        Engine.Grh_Render_To_Hdc Current_Group.GrhIndexes(LoopC), frmTileGroups.TileGroupViewer.hdc, Column * tile_size, Row * tile_size
        
        Column = Column + 1
        If Column = Grhs_per_Row Then
            Row = Row + 1
            Column = 0
        End If
    Next LoopC
    
    'Draw grid
    If Draw_Grid Then
        For LoopC = 1 To Grhs_per_Row - 1
            frmTileGroups.Grid(LoopC).x1 = LoopC * tile_size - 1
            frmTileGroups.Grid(LoopC).x2 = frmTileGroups.Grid(LoopC).x1
            frmTileGroups.Grid(LoopC).Visible = True
            frmTileGroups.Grid(LoopC).BorderColor = Grid_Color
        Next LoopC
        For LoopC = 7 To 9
            frmTileGroups.Grid(LoopC).y1 = (LoopC - 6) * tile_size - 1
            frmTileGroups.Grid(LoopC).y2 = frmTileGroups.Grid(LoopC).y1
            frmTileGroups.Grid(LoopC).Visible = True
            frmTileGroups.Grid(LoopC).BorderColor = Grid_Color
        Next LoopC
    Else
        For LoopC = 1 To 9
            frmTileGroups.Grid(LoopC).Visible = False
        Next LoopC
    End If
    
End Sub

Public Function Tile_Group_Index_Get(ByVal X As Long, ByVal Y As Long) As Long
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 10/18/2003
'Gets the grh index, based on the click coords
'*************************************
    Dim temp_x As Long
    Dim temp_y As Long
    Dim Grhs_per_Row As Long
    Dim temp_grh As Long
    
    'Check there´s is a selected group
    If Current_Group.Name = "" Then
        Exit Function
    End If
    
    temp_x = CLng(X \ tile_size)
    temp_y = CLng(Y \ tile_size)
    
    'Calculate how many Grhs fit in a row
    Grhs_per_Row = frmTileGroups.TileGroupViewer.ScaleWidth / tile_size
    
    temp_grh = temp_x + (temp_y * Grhs_per_Row + 1) + (TileGroupOffset * Grhs_per_Row)
    
    If temp_grh = 0 Then temp_grh = 1
    
    If temp_grh > UBound(Current_Group.GrhIndexes) Then temp_grh = UBound(Current_Group.GrhIndexes)
    
    Tile_Group_Index_Get = Current_Group.GrhIndexes(temp_grh)
    
End Function

Public Sub Load_Triggers_Data_To_List(ByRef ListBoxName As ListBox)
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/27/2003
'Loads the Triggers.dat file to the given list
'*************************************
    Dim Numtriggers As Long
    Dim trigger As Long
    
    Numtriggers = Val(General_Var_Get(App.Path & "\Triggers.dat", "INIT", "NumTriggers"))
    
    For trigger = 1 To Numtriggers
        ListBoxName.AddItem General_Var_Get(App.Path & "\Triggers.dat", "TRIG" & trigger, "Name")
    Next trigger
    
End Sub

Public Sub Load_Particle_Streams_To_ComboBox(ByRef ComboBoxName As ComboBox)
'*************************************
'Coded by Onezero (onezero_ss@hotmail.com)
'Last Modified: 6/4/2003
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    Dim LoopC As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim StreamFile As String
    
    StreamFile = resource_path & "\Particles.ini"
    
    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'clear combo box
    ComboBoxName.Clear
    
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).Name = General_Var_Get(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = General_Var_Get(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).x1 = General_Var_Get(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).y1 = General_Var_Get(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).x2 = General_Var_Get(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).y2 = General_Var_Get(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).angle = General_Var_Get(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = General_Var_Get(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = General_Var_Get(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = General_Var_Get(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = General_Var_Get(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = General_Var_Get(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = General_Var_Get(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).friction = General_Var_Get(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).spin = General_Var_Get(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).AlphaBlend = General_Var_Get(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = General_Var_Get(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = General_Var_Get(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = General_Var_Get(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = General_Var_Get(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = General_Var_Get(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = General_Var_Get(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = General_Var_Get(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = General_Var_Get(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = General_Var_Get(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = General_Var_Get(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).speed = Val(General_Var_Get(StreamFile, Val(LoopC), "Speed"))
        
        StreamData(LoopC).NumGrhs = General_Var_Get(StreamFile, Val(LoopC), "NumGrhs")
        
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(LoopC), "Grh_List")
        
        For i = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(i) = General_Field_Read(Str(i), GrhListing, 44)
        Next i
        StreamData(LoopC).grh_list(i - 1) = StreamData(LoopC).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).b = General_Field_Read(3, TempSet, 44)
            DoEvents
        Next ColorSet
        
        'fill stream type combo box
        ComboBoxName.AddItem LoopC & " - " & StreamData(LoopC).Name
    Next LoopC
    
    'set list box index to 1st item
    ComboBoxName.ListIndex = 0

End Sub

Public Function Load_Grh_Tree(ByRef TreeName As TreeView, ByVal file_path As String) As Boolean
'*************************************************
'Coded by Juan Martín Sotuyo Dodero
'Last Modified: 6/13/2004
'Loads the Tile Script file to the given tree view
'*************************************************
On Local Error GoTo ErrHandler
    Dim ScriptLine As String
    Dim NodeLevel As Long
    Dim NodeCount As Long
    Dim GrhCount As Long
    Dim ParentNodeKey As String
    Dim LoopC As Long
    Dim NodeX As Node
    Dim FirstChar As String * 1
    
    'Check script file exists
    If Not General_File_Exists(file_path, vbNormal) Then
        MsgBox "File " & file_path & " doesn´t exist!", , "Error"
        Exit Function
    End If
    
    Open file_path For Input As #1
    
    Do Until EOF(1)
        Line Input #1, ScriptLine
        
        'Get First Char
        FirstChar = ScriptLine
        
        If FirstChar = "#" And left$(ScriptLine, 4) <> "#EOF" Then
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
            End If
        End If
        
        If left$(ScriptLine, 4) = "#EOF" Then
            NodeLevel = NodeLevel - 1
            'Check if file has ended
            If NodeLevel = -1 Then Exit Do
            NodeCount = Val(General_Field_Read(2, NodeX.Parent.Key, 45))
        End If
        
        If FirstChar = "$" Then
            GrhCount = GrhCount + 1
            Set NodeX = TreeName.Nodes.Add(Str(NodeLevel) & "-" & Str(NodeCount), tvwChild, "grh" & Str(GrhCount), "< On " & frmMain.GrhLayerList.List(Val(Right$(ScriptLine, Len(ScriptLine) - 1)) - 1) & " Layer >")
        End If
        
        If FirstChar = ">" Then
            GrhCount = GrhCount + 1
            If left$(ScriptLine, 2) = ">$" Then
                'Parent is a Grh
                Set NodeX = TreeName.Nodes.Add("grh" & "-" & Str(GrhCount - 1), tvwChild, "grh" & Str(GrhCount), "< On " & frmMain.GrhLayerList.List(Val(Right$(ScriptLine, Len(ScriptLine) - 1)) - 1) & " Layer >")
            Else
                Set NodeX = TreeName.Nodes.Add(Str(NodeLevel) & "-" & Str(NodeCount), tvwChild, "grh" & Str(GrhCount), "Grh " & Right$(ScriptLine, Len(ScriptLine) - 1))
            End If
        End If
        
        DoEvents
    Loop
    
    Set NodeX = Nothing
    Close #1
    Load_Grh_Tree = True
Exit Function

ErrHandler:
    MsgBox "Incorrect Script found. Aborting."
    Set NodeX = Nothing
    prgRun = False
    Close #1
End Function
