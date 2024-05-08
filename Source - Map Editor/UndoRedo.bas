Attribute VB_Name = "UndoRedo"
'**************************************
'Undo-Redo  Coded to work with ORE 0.5
'**************************************
'Coded by Juan Martín Sotuyo Dodero
'(juansotuyo@hotmail.com)
'For exclusive use of DSO development team
'or others allowed by the coder
'**************************************
'Store action must be called BEFORE doing
'a modification on the map, EXCEPT when
'placing a light, a particle stream or
'setting an adjacent map
'**************************************
Option Explicit

Public Enum tools
    grh_map = 1
    lights
    blocking
    trigger
    object
    NPC
    exits
    particle_stream
End Enum

Public Enum action_type
    place = 1
    Remove
    fill    'Also for block borders and adjacent maps
    Remove_all
End Enum

Public Enum light_type
    Light = 1
    base_light
    shadow
End Enum

Private Type Action
    tool As tools
    action_type As action_type

    grh_list_index As Long
    light_list_index As Long
    particle_streams_list_index As Long
    blocked_list_index As Long
    triggers_list_index As Long
    exits_list_index As Long
    object_list_index As Long
    npc_list_index As Long
End Type

Private Type grh
    map_x As Long
    map_y As Long
    
    index As Long
    layer As Byte
    alpha_blending As Boolean
    angle As Single
    subtile As Byte
    h_centered As Boolean
    v_centered As Boolean
End Type

Private Type Light
    map_x As Long
    map_y As Long
    
    color As Long
    range As Byte
    
    index As Long
End Type

Private Type base_light
    'Used with base_lights
    map_x As Long
    map_y As Long
    
    color(0 To 3) As Long
End Type

Private Type shadow
    map_x As Long
    map_y As Long
    
    corner As Long
    color As Long
End Type

Private Type tile_exit
    map_x As Long
    map_y As Long
    
    dest_map As String
    dest_map_x As Long
    dest_map_y As Long
End Type

Private Type objects
    map_x As Long
    map_y As Long

    amount As Long
    index As Long
End Type

Private Type generic
    map_x As Long
    map_y As Long

    index As Long
End Type

Private Type grh_tool
    grh() As grh
End Type

Private Type light_tool
    Light() As Light
    base_light() As base_light
    shadow() As shadow
    light_type As light_type
    fill_color As Long
End Type

Private Type blocking_tool
    map_x() As Long
    map_y() As Long
    
    block_bounds As Boolean 'TRUE = borders have been blocked
    map_border_x As Byte
    map_border_y As Byte
End Type

Private Type trigger_tool
    trigger() As generic
End Type

Private Type exit_tool
    tile_exit() As tile_exit
End Type

Private Type npc_tool
    NPC() As generic
End Type

Private Type obj_tool
    object() As objects
End Type

Private Type particle_group_tool
    particle_stream() As generic
End Type

'Lists
Dim grh_list() As grh_tool
Dim light_list() As light_tool
Dim block_list() As blocking_tool
Dim trigger_list() As trigger_tool
Dim exits_list() As exit_tool
Dim objects_list() As obj_tool
Dim npc_list() As npc_tool
Dim particle_stream_list() As particle_group_tool

Dim last_grh As Long
Dim last_light As Long
Dim last_block As Long
Dim last_trigger As Long
Dim last_exit As Long
Dim last_object As Long
Dim last_npc As Long
Dim last_particle_stream As Long

'Action´s list
Const MAX_ACTIONS = 50
Dim action_list(1 To MAX_ACTIONS) As Action
Dim last_action As Byte

Public Sub store_action(ByRef tool As tools, ByVal action_type As action_type, Optional ByVal map_x As Long = 1, Optional ByVal map_y As Long = 1, Optional subindex As Long = 1, _
                        Optional ByVal layer As Byte, Optional ByVal color As Long, Optional ByVal range As Byte, Optional ByVal corner As Byte, Optional light_type As light_type, _
                        Optional ByVal fill_color As Long, Optional ByVal block_bounds As Boolean = False, Optional ByVal x_border As Long, Optional ByVal y_border As Long, _
                        Optional redo As Boolean = False, Optional ByVal same_index As Boolean, Optional ByVal subtile As Byte = 9)
'The subindex indicates (if using the same index) how many have been stored before (used especially for Remove All)
    If Not Engine.Map_In_Bounds(map_x, map_y) Then
        Exit Sub
    End If
    
    If subindex = 1 Then
        'Increase action
        Increase_Action
    End If
    
    If Not redo Then
        'Erase all others
        Dim a As Long
        Dim correction As Integer
        If same_index Then
            correction = 1
        End If
        If last_action < MAX_ACTIONS Then
            For a = last_action + correction To MAX_ACTIONS
                'Check what kind of action it was and resize that list
                Select Case action_list(a).tool
                    Case blocking
                        destroy_blocking action_list(a).blocked_list_index
                    Case exits
                        destroy_exits action_list(a).exits_list_index
                    Case grh
                        destroy_grh action_list(a).grh_list_index
                    Case lights
                        destroy_light action_list(a).light_list_index
                    Case NPC
                        destroy_npc action_list(a).npc_list_index
                    Case object
                        destroy_object action_list(a).object_list_index
                    Case particle_stream
                        destroy_particle_stream action_list(a).particle_streams_list_index
                    Case trigger
                        destroy_trigger action_list(a).triggers_list_index
                End Select
                
                action_list(a).action_type = 0
                action_list(a).blocked_list_index = 0
                action_list(a).exits_list_index = 0
                action_list(a).grh_list_index = 0
                action_list(a).light_list_index = 0
                action_list(a).npc_list_index = 0
                action_list(a).object_list_index = 0
                action_list(a).particle_streams_list_index = 0
                action_list(a).tool = 0
                action_list(a).triggers_list_index = 0
            Next a
            'If so, redo should be disabled
            frmMain.MnuRedo.Enabled = False
            frmMain.Toolbar1.Buttons(6).Enabled = False
        End If
    End If
    
    'Copy data
    action_list(last_action).tool = tool
    action_list(last_action).action_type = action_type
    
    Select Case tool
        Case grh
            Select Case action_type
                Case Is <= Remove   '(place or remove)
                    action_list(last_action).grh_list_index = store_grh(map_x, map_y, layer, , subtile)
                Case Remove_all
                    action_list(last_action).grh_list_index = store_grh_remove_all(map_x, map_y)
                Case fill
                    action_list(last_action).grh_list_index = store_grh_map_fill(layer)
            End Select
            
        Case lights
            If light_type = Light Then
                Select Case action_type
                    Case place
                        action_list(last_action).light_list_index = store_light_place(map_x, map_y, color, range)
                    Case Remove
                        action_list(last_action).light_list_index = store_light_remove(map_x, map_y)
                    Case Remove_all
                        action_list(last_action).light_list_index = store_light_remove_all
                End Select
            ElseIf light_type = base_light Then
                Select Case action_type
                    Case place
                        action_list(last_action).light_list_index = store_base_light_place(map_x, map_y)
                    Case fill
                        action_list(last_action).light_list_index = store_base_light_map_fill(fill_color)
                End Select
            Else
                action_list(last_action).light_list_index = store_shadow_place(map_x, map_y, corner, subindex)
            End If
            
        Case blocking
            Select Case action_type
                Case Is <= Remove    '(place or remove)
                    action_list(last_action).blocked_list_index = store_block(map_x, map_y)
                Case Remove_all
                    action_list(last_action).blocked_list_index = store_block_remove_all
                Case fill
                    action_list(last_action).blocked_list_index = store_block_borders(x_border, y_border, block_bounds)
            End Select
            
        Case trigger
            Select Case action_type
                Case Is <= Remove   '(place or remove)
                    action_list(last_action).triggers_list_index = store_trigger(map_x, map_y)
                Case Is >= fill     '(fill or remove all)
                    action_list(last_action).triggers_list_index = store_trigger_remove_all
            End Select
            
        Case object
            Select Case action_type
                Case Is <= Remove   '(place or remove)
                    action_list(last_action).object_list_index = store_object(map_x, map_y)
                Case Remove_all
                    action_list(last_action).object_list_index = store_object_remove_all
            End Select
            
        Case NPC
            Select Case action_type
                Case Is <= Remove   '(place or remove)
                    action_list(last_action).npc_list_index = store_npc(map_x, map_y)
                Case Remove_all
                    action_list(last_action).npc_list_index = store_npc_remove_all
            End Select
            
        Case exits
            Select Case action_type
                Case Is <= fill '(place, remove or fill)
                    action_list(last_action).exits_list_index = store_exit(map_x, map_y, subindex)
                Case Remove_all
                    action_list(last_action).exits_list_index = store_exit_remove_all
            End Select
            
        Case particle_stream
            Select Case action_type
                Case Is <= Remove   '(place or remove)
                    action_list(last_action).particle_streams_list_index = store_particle_stream(map_x, map_y)
                Case Remove_all
                    action_list(last_action).particle_streams_list_index = store_particle_stream_remove_all
            End Select
    End Select
    
    If last_action > 0 Then
        frmMain.MnuUndo.Enabled = True
        frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
End Sub

Private Function store_grh(ByVal map_x As Long, ByVal map_y As Long, ByVal layer As Long, Optional index As Long = 1, Optional ByVal subtile As Byte = 9) As Long
'**************************************************************
'Gets the existing Grh at the given pos and layer and stores it
'**************************************************************
    If index = 1 Then last_grh = last_grh + 1
    
    ReDim Preserve grh_list(1 To last_grh)
    ReDim Preserve grh_list(last_grh).grh(1 To index)
    
    grh_list(last_grh).grh(index).map_x = map_x
    grh_list(last_grh).grh(index).map_y = map_y
    grh_list(last_grh).grh(index).layer = layer
    
    If layer = 1 Or layer = 3 Or layer = 5 Then
        grh_list(last_grh).grh(index).index = Engine.Map_Grh_Info_Get(map_x, map_y, (layer + 1) / 2, _
                                            grh_list(last_grh).grh(index).alpha_blending, _
                                            grh_list(last_grh).grh(index).angle, _
                                            grh_list(last_grh).grh(index).h_centered, grh_list(last_grh).grh(index).v_centered)

    Else
        grh_list(last_grh).grh(index).subtile = subtile
        grh_list(last_grh).grh(index).index = Engine.Map_Decoration_Info_Get(map_x, map_y, subtile, grh_list(last_grh).grh(index).alpha_blending, _
                                                                            grh_list(last_grh).grh(index).angle, grh_list(last_grh).grh(index).h_centered, _
                                                                            grh_list(last_grh).grh(index).v_centered)
        
    End If
    
    store_grh = last_grh
End Function

Private Function store_grh_remove_all(ByVal map_x As Long, ByVal map_y As Long) As Long
'**********************************************************
'Gets the existing Grhs at the given pos and stores them
'**********************************************************
    Dim LoopC As Long
    Dim LoopC2 As Long
    Dim index As Long
    
    For LoopC = 1 To 5 Step 2
        index = index + 1
        LoopC2 = LoopC2 + 1
        store_grh_remove_all = store_grh(map_x, map_y, LoopC, index)
    Next LoopC
    
    For LoopC2 = 2 To 4 Step 2
        For LoopC = 0 To 8
            index = index + 1
            store_grh_remove_all = store_grh(map_x, map_y, LoopC2, index, LoopC)
        Next LoopC
    Next LoopC2
    
End Function

Private Function store_grh_map_fill(ByVal layer As Byte) As Long
'**********************************************************
'Gets the existing Grhs of all tiles at the given layer and stores them
'**********************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    'Get map dimensions
    Engine.Map_Bounds_Get max_x, max_y
    
    'Store indexes
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            LoopC = LoopC + 1
            store_grh_map_fill = store_grh(map_x, map_y, layer, LoopC)
        Next map_y
    Next map_x
    
End Function

Private Function store_light_place(ByVal map_x As Long, ByVal map_y As Long, ByVal color As Long, ByVal range As Byte) As Long
    last_light = last_light + 1
    ReDim Preserve light_list(1 To last_light)
    ReDim light_list(last_light).Light(1 To 1)
    
    light_list(last_light).light_type = Light
    
    'Store index to destroy it (undo)
    light_list(last_light).Light(1).index = Engine.Map_Light_Get(map_x, map_y)
    
    'Store data to rebiuld it (redo)
    light_list(last_light).Light(1).color = color
    light_list(last_light).Light(1).map_x = map_x
    light_list(last_light).Light(1).map_y = map_y
    light_list(last_light).Light(1).range = range
    
    store_light_place = last_light
End Function

Private Function store_light_remove(ByVal map_x As Long, ByVal map_y As Long) As Long
    last_light = last_light + 1
    
    ReDim Preserve light_list(1 To last_light)
    ReDim light_list(last_light).Light(1 To 1)
    
    light_list(last_light).light_type = Light
    
    'Get the light´s data to rebiuld it
    Engine.Light_Info_Get Engine.Map_Light_Get(map_x, map_y), light_list(last_light).Light(1).map_x, light_list(last_light).Light(1).map_y, _
            light_list(last_light).Light(1).color, light_list(last_light).Light(1).range
    
    store_light_remove = last_light
End Function

Private Function store_light_remove_all() As Long
    Dim LoopC As Long
    Dim light_counter As Long
    
    last_light = last_light + 1
    ReDim Preserve light_list(1 To last_light)
    
    light_list(last_light).light_type = Light
    
    For LoopC = 1 To Engine.Light_Count_Get
        light_counter = light_counter + 1
        ReDim Preserve light_list(last_light).Light(1 To light_counter)
        If Not Engine.Light_Info_Get(LoopC, light_list(last_light).Light(light_counter).map_x, light_list(last_light).Light(light_counter).map_y, _
                    light_list(last_light).Light(light_counter).color, light_list(last_light).Light(light_counter).range) Then
            light_counter = light_counter - 1
        End If
    Next LoopC
    
    ReDim Preserve light_list(last_light).Light(1 To light_counter)
    
    store_light_remove_all = last_light
End Function

Private Function store_base_light_place(ByVal map_x As Long, ByVal map_y As Long) As Long
    Dim LoopC As Long
    
    last_light = last_light + 1
    ReDim Preserve light_list(1 To last_light)
    ReDim light_list(last_light).base_light(1 To 1)
    
    light_list(last_light).light_type = base_light
    
    light_list(last_light).base_light(1).map_x = map_x
    light_list(last_light).base_light(1).map_y = map_y
    
    For LoopC = 0 To 3
        light_list(last_light).base_light(1).color(LoopC) = Engine.Map_Base_Light_Get(map_x, map_y, LoopC)
    Next LoopC
    
    store_base_light_place = last_light
End Function

Private Function store_base_light_map_fill(ByVal color As Long) As Long
    Dim LoopC As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    
    'Get map dimensions
    Engine.Map_Bounds_Get max_x, max_y
    
    last_light = last_light + 1
    ReDim Preserve light_list(1 To last_light)
    ReDim light_list(last_light).base_light(1 To max_x * max_y)
    
    light_list(last_light).light_type = base_light
    light_list(last_light).fill_color = color
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            light_list(last_light).base_light(max_x * (map_y - 1) + map_x).map_x = map_x
            light_list(last_light).base_light(max_x * (map_y - 1) + map_x).map_y = map_y
    
            For LoopC = 0 To 3
                light_list(last_light).base_light(max_x * (map_y - 1) + map_x).color(LoopC) = Engine.Map_Base_Light_Get(map_x, map_y, LoopC)
            Next LoopC
        Next map_y
    Next map_x
    
    store_base_light_map_fill = last_light
End Function

Private Function store_shadow_place(ByVal map_x As Long, ByVal map_y As Long, ByVal corner As Long, ByVal subindex As Long) As Long
    If subindex = 1 Then last_light = last_light + 1
    
    ReDim Preserve light_list(1 To last_light)
    ReDim Preserve light_list(last_light).shadow(1 To subindex)
    
    light_list(last_light).shadow(subindex).map_x = map_x
    light_list(last_light).shadow(subindex).map_y = map_y
    light_list(last_light).shadow(subindex).corner = corner
    light_list(last_light).shadow(subindex).color = Engine.Map_Base_Light_Get(map_x, map_y, corner)
    
    store_shadow_place = last_light
End Function

Private Function store_block(ByVal map_x As Long, ByVal map_y As Long) As Long
    last_block = last_block + 1
    
    ReDim Preserve block_list(1 To last_block)
    ReDim block_list(last_block).map_x(1 To 1)
    ReDim block_list(last_block).map_y(1 To 1)
    
    block_list(last_block).map_x(1) = map_x
    block_list(last_block).map_y(1) = map_y
    
    store_block = last_block
End Function

Private Function store_block_remove_all() As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim counter As Long
    
    'Get map size
    Engine.Map_Bounds_Get max_x, max_y
    
    last_block = last_block + 1
    
    ReDim Preserve block_list(1 To last_block)
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            counter = counter + 1
            If Engine.Map_Blocked_Get(map_x, map_y) Then
                ReDim Preserve block_list(last_block).map_x(1 To counter)
                ReDim Preserve block_list(last_block).map_y(1 To counter)
                
                block_list(last_block).map_x(counter) = map_x
                block_list(last_block).map_y(counter) = map_y
            Else
                counter = counter - 1
            End If
        Next map_y
    Next map_x
    
    store_block_remove_all = last_block
End Function

Private Function store_block_borders(ByVal x_border As Byte, ByVal y_border As Byte, ByVal block_state As Boolean) As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    last_block = last_block + 1
    
    ReDim Preserve block_list(1 To last_block)
    ReDim Preserve block_list(last_block).map_x(1 To 1)
    ReDim Preserve block_list(last_block).map_y(1 To 1)
    
    block_list(last_block).block_bounds = True
    block_list(last_block).map_border_x = x_border
    block_list(last_block).map_border_y = y_border
    block_list(last_block).block_bounds = block_state
    
    Engine.Map_Bounds_Get max_x, max_y
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            If map_x <= x_border Or map_y <= y_border Then
                If block_state Then
                    If Engine.Map_Blocked_Get(map_x, map_y) Then
                        LoopC = LoopC + 1
                        ReDim Preserve block_list(last_block).map_x(1 To LoopC)
                        ReDim Preserve block_list(last_block).map_y(1 To LoopC)
                        block_list(last_block).map_x(LoopC) = map_x
                        block_list(last_block).map_y(LoopC) = map_y
                    End If
                Else
                    If Not Engine.Map_Blocked_Get(map_x, map_y) Then
                        LoopC = LoopC + 1
                        ReDim Preserve block_list(last_block).map_x(1 To LoopC)
                        ReDim Preserve block_list(last_block).map_y(1 To LoopC)
                        block_list(last_block).map_x(LoopC) = map_x
                        block_list(last_block).map_y(LoopC) = map_y
                    End If
                End If
            End If

            If map_x > max_x - x_border Or map_y > max_y - y_border Then
                If Engine.Map_Blocked_Get(map_x, map_y) Then
                    LoopC = LoopC + 1
                    ReDim Preserve block_list(last_block).map_x(1 To LoopC)
                    ReDim Preserve block_list(last_block).map_y(1 To LoopC)
                    block_list(last_block).map_x(LoopC) = map_x
                    block_list(last_block).map_y(LoopC) = map_y
                End If
            End If
        Next map_y
    Next map_x
    
    store_block_borders = last_block
End Function

Private Function store_trigger(ByVal map_x As Long, ByVal map_y As Long) As Long
    last_trigger = last_trigger + 1
    
    ReDim Preserve trigger_list(1 To last_trigger)
    ReDim trigger_list(last_trigger).trigger(1 To 1)
    
    trigger_list(last_trigger).trigger(1).index = Engine.Map_Trigger_Get(map_x, map_y)
    trigger_list(last_trigger).trigger(1).map_x = map_x
    trigger_list(last_trigger).trigger(1).map_y = map_y
    
    store_trigger = last_trigger
End Function

Private Function store_trigger_remove_all() As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    'Get map size
    Engine.Map_Bounds_Get max_x, max_y
    
    last_trigger = last_trigger + 1
    ReDim Preserve trigger_list(1 To last_trigger)
    ReDim trigger_list(last_trigger).trigger(1 To max_x * max_y)
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            LoopC = LoopC + 1
            trigger_list(last_trigger).trigger(LoopC).index = Engine.Map_Trigger_Get(map_x, map_y)
            trigger_list(last_trigger).trigger(LoopC).map_x = map_x
            trigger_list(last_trigger).trigger(LoopC).map_y = map_y
        Next map_y
    Next map_x
    
    store_trigger_remove_all = last_trigger
End Function

Private Function store_object(ByVal map_x As Long, ByVal map_y As Long) As Long
    last_object = last_object + 1
    
    ReDim Preserve objects_list(1 To last_object)
    ReDim objects_list(last_object).object(1 To 1)
    
    objects_list(last_object).object(1).map_x = map_x
    objects_list(last_object).object(1).map_y = map_y
    
    Engine.Map_Item_Get map_x, map_y, objects_list(last_object).object(1).index, _
                        objects_list(last_object).object(1).amount
    
    store_object = last_object
End Function

Private Function store_object_remove_all() As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    'Get map size
    Engine.Map_Bounds_Get max_x, max_y
    
    last_object = last_object + 1
    
    ReDim Preserve objects_list(1 To last_object)
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            LoopC = LoopC + 1
            ReDim Preserve objects_list(last_object).object(1 To LoopC)
            If Not Engine.Map_Item_Get(map_x, map_y, objects_list(last_object).object(LoopC).index, _
                                objects_list(last_object).object(LoopC).amount) Then
                LoopC = LoopC - 1
            Else
                objects_list(last_object).object(LoopC).map_x = map_x
                objects_list(last_object).object(LoopC).map_y = map_y
            End If
        Next map_y
    Next map_x
    
    ReDim Preserve objects_list(last_object).object(1 To LoopC)
    
    store_object_remove_all = last_object
End Function

Private Function store_npc(ByVal map_x As Long, ByVal map_y As Long) As Long
    On Error Resume Next
    last_npc = last_npc + 1
    
    ReDim Preserve npc_list(1 To last_npc)
    ReDim npc_list(last_npc).NPC(1 To 1)
    
    npc_list(last_npc).NPC(1).map_x = map_x
    npc_list(last_npc).NPC(1).map_y = map_y
    Engine.Map_NPC_Get map_x, map_y, npc_list(last_npc).NPC(1).index
    
    store_npc = last_npc
End Function

Private Function store_npc_remove_all() As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    'Get map size
    Engine.Map_Bounds_Get max_x, max_y
    
    last_npc = last_npc + 1
    ReDim Preserve npc_list(1 To last_npc)
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            LoopC = LoopC + 1
            ReDim Preserve npc_list(last_npc).NPC(1 To LoopC)
            If Not Engine.Map_NPC_Get(map_x, map_y, npc_list(last_npc).NPC(LoopC).index) Then
                LoopC = LoopC - 1
            Else
                npc_list(last_npc).NPC(LoopC).map_x = map_x
                npc_list(last_npc).NPC(LoopC).map_y = map_y
            End If
        Next map_y
    Next map_x
    
    ReDim Preserve npc_list(last_npc).NPC(1 To LoopC)
    
    store_npc_remove_all = last_npc
End Function

Private Function store_exit(ByVal map_x As Long, ByVal map_y As Long, Optional ByVal subindex As Long = 1) As Long
    If subindex = 1 Then last_exit = last_exit + 1
    
    ReDim Preserve exits_list(1 To last_exit)
    ReDim Preserve exits_list(last_exit).tile_exit(1 To subindex)
    
    exits_list(last_exit).tile_exit(subindex).map_x = map_x
    exits_list(last_exit).tile_exit(subindex).map_y = map_y
    
    Engine.Map_Exit_Get map_x, map_y, exits_list(last_exit).tile_exit(subindex).dest_map, _
                        exits_list(last_exit).tile_exit(subindex).dest_map_x, _
                        exits_list(last_exit).tile_exit(subindex).dest_map_y
    
    store_exit = last_exit
End Function

Private Function store_exit_remove_all() As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    'Get map size
    Engine.Map_Bounds_Get max_x, max_y
    
    last_exit = last_exit + 1
    ReDim Preserve exits_list(1 To last_exit)
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            LoopC = LoopC + 1
            ReDim Preserve exits_list(last_exit).tile_exit(1 To LoopC)
            If Not Engine.Map_Exit_Get(map_x, map_y, exits_list(last_exit).tile_exit(LoopC).dest_map, _
                                    exits_list(last_exit).tile_exit(LoopC).dest_map_x, _
                                    exits_list(last_exit).tile_exit(LoopC).dest_map_y) Then
                LoopC = LoopC - 1
            Else
                exits_list(last_exit).tile_exit(LoopC).map_x = map_x
                exits_list(last_exit).tile_exit(LoopC).map_y = map_y
            End If
        Next map_y
    Next map_x
    
    ReDim Preserve exits_list(last_exit).tile_exit(1 To LoopC)
    
    store_exit_remove_all = last_exit
End Function

Private Function store_particle_stream(ByVal map_x As Long, ByVal map_y As Long) As Long
    last_particle_stream = last_particle_stream + 1
    ReDim Preserve particle_stream_list(1 To last_particle_stream)
    ReDim particle_stream_list(last_particle_stream).particle_stream(1 To 1)
    
    particle_stream_list(last_particle_stream).particle_stream(1).index = Engine.Map_Particle_Group_Get(map_x, map_y)
    particle_stream_list(last_particle_stream).particle_stream(1).map_x = map_x
    particle_stream_list(last_particle_stream).particle_stream(1).map_y = map_y
    
    store_particle_stream = last_particle_stream
End Function

Private Function store_particle_stream_remove_all() As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    Dim temp_index As Long
    
    'Get map size
    Engine.Map_Bounds_Get max_x, max_y
    
    last_particle_stream = last_particle_stream + 1
    ReDim Preserve particle_stream_list(1 To last_particle_stream)
    
    For map_x = 1 To max_x
        For map_y = 1 To max_y
            temp_index = Engine.Particle_Type_Get(Engine.Map_Particle_Group_Get(map_x, map_y))
            If temp_index Then
                LoopC = LoopC + 1
                ReDim Preserve particle_stream_list(last_particle_stream).particle_stream(1 To LoopC)
                particle_stream_list(last_particle_stream).particle_stream(LoopC).map_x = map_x
                particle_stream_list(last_particle_stream).particle_stream(LoopC).map_y = map_y
                particle_stream_list(last_particle_stream).particle_stream(LoopC).index = temp_index
            End If
        Next map_y
    Next map_x
    
    store_particle_stream_remove_all = last_particle_stream
End Function

Private Sub Increase_Action()
    Dim LoopC As Long
    
    'Increase action pos
    If last_action > MAX_ACTIONS - 1 Then
        'Destroy that action index
        destroy_index 1
        'Move list back
        For LoopC = 1 To MAX_ACTIONS - 1
            action_list(LoopC) = action_list(LoopC + 1)
        Next LoopC
    Else
        last_action = last_action + 1
    End If
    
    'Clear all data
    action_list(last_action).action_type = 0
    action_list(last_action).blocked_list_index = 0
    action_list(last_action).exits_list_index = 0
    action_list(last_action).grh_list_index = 0
    action_list(last_action).light_list_index = 0
    action_list(last_action).npc_list_index = 0
    action_list(last_action).object_list_index = 0
    action_list(last_action).particle_streams_list_index = 0
    action_list(last_action).tool = 0
    action_list(last_action).triggers_list_index = 0
End Sub

Private Sub destroy_index(ByVal index As Long)
    'Remove data from action lists
    Select Case action_list(index).tool
        Case grh
            destroy_grh action_list(index).grh_list_index
        Case lights
            destroy_light action_list(index).light_list_index
        Case particle_stream
            destroy_particle_stream action_list(index).particle_streams_list_index
        Case object
            destroy_object action_list(index).object_list_index
        Case NPC
            destroy_npc action_list(index).npc_list_index
        Case exits
            destroy_exits action_list(index).exits_list_index
        Case trigger
            destroy_trigger action_list(index).triggers_list_index
        Case blocking
            destroy_blocking action_list(index).blocked_list_index
    End Select
    
    'Clear all data from the list itself
    action_list(index).action_type = 0
    action_list(index).blocked_list_index = 0
    action_list(index).exits_list_index = 0
    action_list(index).grh_list_index = 0
    action_list(index).light_list_index = 0
    action_list(index).npc_list_index = 0
    action_list(index).object_list_index = 0
    action_list(index).particle_streams_list_index = 0
    action_list(index).tool = 0
    action_list(index).triggers_list_index = 0
End Sub

Public Sub Clear_Action_List()
    Dim LoopC As Long
    
    'Destroy all indexes
    For LoopC = 1 To MAX_ACTIONS
        destroy_index LoopC
    Next LoopC
    
    'Reset last_action
    last_action = 0
    
    'Disable undo-redo buttons / commands
    frmMain.MnuRedo.Enabled = False
    frmMain.Toolbar1.Buttons(6).Enabled = False
    frmMain.MnuUndo.Enabled = False
    frmMain.Toolbar1.Buttons(5).Enabled = False
End Sub

Public Sub undo_redo(ByVal undo As Boolean)
'If undo = TRUE we undo last action, otherwise, we redo it
    If undo Then
        last_action = last_action - 1
    End If
    
    'Check which tool was used
    Select Case action_list(last_action + 1).tool
        Case blocking
            'block
            Select Case action_list(last_action + 1).action_type
                Case place
                    block_place action_list(last_action + 1).blocked_list_index, undo
                Case Remove
                    block_remove action_list(last_action + 1).blocked_list_index, undo
                Case Remove_all
                    block_remove_all action_list(last_action + 1).blocked_list_index, undo
                Case fill
                    block_fill action_list(last_action + 1).blocked_list_index, undo
            End Select
        
        Case exits
            'exits
            Select Case action_list(last_action + 1).action_type
                Case place
                    exit_place action_list(last_action + 1).exits_list_index, undo
                Case Remove
                    exit_remove action_list(last_action + 1).exits_list_index, undo
                Case Remove_all
                    exit_remove_all action_list(last_action + 1).exits_list_index, undo
                Case fill
                    exit_fill action_list(last_action + 1).exits_list_index, undo
            End Select
        
        Case grh
            'grh
            Select Case action_list(last_action + 1).action_type
                Case place
                    grh_place action_list(last_action + 1).grh_list_index
                Case Remove
                    grh_remove action_list(last_action + 1).grh_list_index, undo
                Case Remove_all
                    grh_remove_all action_list(last_action + 1).grh_list_index, undo
                Case fill
                    grh_fill action_list(last_action + 1).grh_list_index, undo
            End Select
        
        Case lights
            'light
            If light_list(action_list(last_action + 1).light_list_index).light_type = Light Then
                Select Case action_list(last_action + 1).action_type
                    Case place
                        light_place action_list(last_action + 1).light_list_index, undo
                    Case Remove
                        light_remove action_list(last_action + 1).light_list_index, undo
                    Case Remove_all
                        light_remove_all action_list(last_action + 1).light_list_index, undo
                    End Select
            ElseIf light_list(action_list(last_action + 1).light_list_index).light_type = base_light Then
                Select Case action_list(last_action + 1).action_type
                    Case place
                        base_light_place action_list(last_action + 1).light_list_index
                    Case fill
                        base_light_fill action_list(last_action + 1).light_list_index, undo
                    End Select
            Else
                shadow_place action_list(last_action + 1).light_list_index
            End If
    
        Case NPC
            'npc
            Select Case action_list(last_action + 1).action_type
                Case place
                    npc_place action_list(last_action + 1).npc_list_index, undo
                Case Remove
                    npc_remove action_list(last_action + 1).npc_list_index, undo
                Case Remove_all
                    npc_remove_all action_list(last_action + 1).npc_list_index, undo
            End Select
    
        Case object
            'object
            Select Case action_list(last_action + 1).action_type
                Case place
                    object_place action_list(last_action + 1).object_list_index, undo
                Case Remove
                    object_remove action_list(last_action + 1).object_list_index, undo
                Case Remove_all
                    object_remove_all action_list(last_action + 1).object_list_index, undo
            End Select
    
        Case particle_stream
            'particle stream
            Select Case action_list(last_action + 1).action_type
                Case place
                    particle_stream_place action_list(last_action + 1).particle_streams_list_index, undo
                Case Remove
                    particle_stream_remove action_list(last_action + 1).particle_streams_list_index, undo
                Case Remove_all
                    particle_stream_remove_all action_list(last_action + 1).particle_streams_list_index, undo
            End Select
    
        Case trigger
            'trigger
            Select Case action_list(last_action + 1).action_type
                Case place
                    trigger_place action_list(last_action + 1).triggers_list_index
                Case Remove
                    trigger_remove action_list(last_action + 1).triggers_list_index, undo
                Case Remove_all
                    trigger_remove_all action_list(last_action + 1).triggers_list_index, undo
                Case fill
                    trigger_fill action_list(last_action + 1).triggers_list_index, undo
            End Select
    End Select
    
    If undo Then
        last_action = last_action - 1
    End If
    
    If last_action = 0 Then
        frmMain.MnuUndo.Enabled = False
        frmMain.Toolbar1.Buttons(5).Enabled = False
    End If
    
    If last_action = MAX_ACTIONS Then
        frmMain.MnuRedo.Enabled = False
        frmMain.Toolbar1.Buttons(6).Enabled = False
    End If
    
    If last_action + 1 <= MAX_ACTIONS Then
        frmMain.MnuRedo.Enabled = (action_list(last_action + 1).action_type > 0)
        frmMain.Toolbar1.Buttons(6).Enabled = frmMain.MnuRedo.Enabled
    End If
    
    If last_action > 0 Then
        frmMain.MnuUndo.Enabled = True
        frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
End Sub

Private Sub block_place(ByVal list_index As Long, ByVal undo As Boolean)
    Dim block_state As Boolean
    
    store_action blocking, place, block_list(list_index).map_x(1), block_list(list_index).map_y(1), , , , , , , , , , , True
    
    If undo Then
        block_state = False
    Else
        block_state = True
    End If
    
    Engine.Map_Blocked_Set block_list(list_index).map_x(1), block_list(list_index).map_y(1), block_state
End Sub

Private Sub block_remove(ByVal list_index As Long, ByVal undo As Boolean)
    Dim block_state As Boolean
    
    store_action blocking, Remove, block_list(list_index).map_x(1), block_list(list_index).map_y(1), , , , , , , , , , , True
    
    If undo Then
        block_state = True
    Else
        block_state = False
    End If
    
    Engine.Map_Blocked_Set block_list(list_index).map_x(1), block_list(list_index).map_y(1), block_state
End Sub

Private Sub block_remove_all(ByVal list_index As Long, ByVal undo As Boolean)
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    If undo Then
        For LoopC = 1 To UBound(block_list(list_index).map_x)
            Engine.Map_Blocked_Set block_list(list_index).map_x(LoopC), block_list(list_index).map_y(LoopC), True
        Next LoopC
    Else
        Engine.Map_Bounds_Get max_x, max_y
    
        For map_x = 1 To max_x
            For map_y = 1 To max_y
                Engine.Map_Blocked_Set map_x, map_y, False
            Next map_y
        Next map_x
    End If
    
    'Since nothing is stored, we increase last action to avoid errors
    last_action = last_action + 1
End Sub

Private Sub block_fill(ByVal list_index As Long, ByVal undo As Boolean)
    Dim LoopC As Long
    Dim block_state As Boolean
    
    store_action blocking, fill, , , , , , , , , , block_list(list_index).block_bounds, block_list(list_index).map_border_x, block_list(list_index).map_border_y, True
    
    If undo Then
        If block_list(list_index).block_bounds Then
            block_state = False
        Else
            block_state = True
        End If
        Engine.Map_Edges_Blocked_Set block_list(list_index).map_border_x, block_list(list_index).map_border_y, block_state
        'Block those tiles which were blocked before we blocked bounds (we invert block_state once more)
        If block_state Then
            block_state = False
        Else
            block_state = True
        End If
        For LoopC = 1 To UBound(block_list(list_index).map_x)
            Engine.Map_Blocked_Set block_list(list_index).map_x(LoopC), block_list(list_index).map_y(LoopC), block_state
        Next LoopC
        'Set the button´s caption
        frmMain.BlockBordersCmd.Caption = "Block Borders ON"
    Else
        Engine.Map_Edges_Blocked_Set block_list(list_index).map_border_x, block_list(list_index).map_border_y, block_list(list_index).block_bounds
        'Set the button´s caption
        frmMain.BlockBordersCmd.Caption = "Block Borders OFF"
    End If
End Sub

Private Sub exit_place(ByVal list_index As Long, ByVal undo As Boolean)
    store_action exits, place, exits_list(list_index).tile_exit(1).map_x, exits_list(list_index).tile_exit(1).map_y, , , , , , , , , , , True
    
    If undo Then
        Engine.Map_Exit_Remove exits_list(list_index).tile_exit(1).map_x, exits_list(list_index).tile_exit(1).map_y
    Else
        Engine.Map_Exit_Add exits_list(list_index).tile_exit(1).map_x, exits_list(list_index).tile_exit(1).map_y, exits_list(list_index).tile_exit(1).dest_map, _
                            exits_list(list_index).tile_exit(1).dest_map_x, exits_list(list_index).tile_exit(1).dest_map_y
    End If
End Sub

Private Sub exit_remove(ByVal list_index As Long, ByVal undo As Boolean)
    If undo Then
        Engine.Map_Exit_Add exits_list(list_index).tile_exit(1).map_x, exits_list(list_index).tile_exit(1).map_y, _
                            exits_list(list_index).tile_exit(1).dest_map, exits_list(list_index).tile_exit(1).map_x, _
                            exits_list(list_index).tile_exit(1).map_y
    Else
        Engine.Map_Exit_Remove exits_list(list_index).tile_exit(1).map_x, exits_list(list_index).tile_exit(1).map_y
    End If
    
    last_action = last_action + 1
End Sub

Private Sub exit_remove_all(ByVal list_index As Long, ByVal undo As Boolean)
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    If undo Then
        For LoopC = 1 To UBound(exits_list(list_index).tile_exit)
            Engine.Map_Exit_Add exits_list(list_index).tile_exit(LoopC).map_x, exits_list(list_index).tile_exit(LoopC).map_y, _
                                exits_list(list_index).tile_exit(LoopC).dest_map, exits_list(list_index).tile_exit(LoopC).dest_map_x, _
                                exits_list(list_index).tile_exit(LoopC).dest_map_y
        Next LoopC
    Else
        Engine.Map_Bounds_Get max_x, max_y
        
        For map_x = 1 To max_x
            For map_y = 1 To max_y
                Engine.Map_Exit_Remove map_x, map_y
            Next map_y
        Next map_x
    End If
    
    last_action = last_action + 1
End Sub

Private Sub exit_fill(ByVal list_index As Long, ByVal undo As Boolean)
    Dim LoopC As Long
    
    If undo Then
        For LoopC = 1 To UBound(exits_list(list_index).tile_exit)
            Engine.Map_Exit_Remove exits_list(list_index).tile_exit(LoopC).map_x, exits_list(list_index).tile_exit(LoopC).map_y
        Next LoopC
    Else
        For LoopC = 1 To UBound(exits_list(list_index).tile_exit)
            Engine.Map_Exit_Add exits_list(list_index).tile_exit(LoopC).map_x, exits_list(list_index).tile_exit(LoopC).map_y, _
                                exits_list(list_index).tile_exit(LoopC).dest_map, exits_list(list_index).tile_exit(LoopC).dest_map_x, _
                                exits_list(list_index).tile_exit(LoopC).dest_map_y
        Next LoopC
    End If
    
    'Since nothing was stored, we increase last_action to avoid errors
    last_action = last_action + 1
End Sub

Private Sub grh_place(ByVal list_index As Long)
    store_action grh, place, grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, , grh_list(list_index).grh(1).layer, , , , , , , , , , True, _
                grh_list(list_index).grh(1).subtile
    
    If grh_list(list_index).grh(1).layer = 1 Or grh_list(list_index).grh(1).layer = 3 Or grh_list(list_index).grh(1).layer = 5 Then
        If grh_list(list_index).grh(1).index > 0 Then
            Engine.Map_Grh_Set grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, _
                                grh_list(list_index).grh(1).index, (grh_list(list_index).grh(1).layer + 1) / 2, _
                                grh_list(list_index).grh(1).alpha_blending, grh_list(list_index).grh(1).angle, _
                                grh_list(list_index).grh(1).h_centered, grh_list(list_index).grh(1).v_centered
        Else
            Engine.Map_Grh_UnSet grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, (grh_list(list_index).grh(1).layer + 1) / 2
        End If
    Else
        Dim Render_On_Top As Boolean
        If grh_list(list_index).grh(1).layer = 2 Then
            Render_On_Top = False
        Else
            Render_On_Top = True
        End If
        If grh_list(list_index).grh(1).index > 0 Then
            Engine.Map_Decoration_Add grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, grh_list(list_index).grh(1).subtile, _
                                        grh_list(list_index).grh(1).index, Render_On_Top, grh_list(list_index).grh(1).alpha_blending, _
                                        grh_list(list_index).grh(1).angle, grh_list(list_index).grh(1).h_centered, grh_list(list_index).grh(1).v_centered
        Else
            Engine.Map_Decoration_Remove grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, grh_list(list_index).grh(1).subtile
        End If
    End If
End Sub

Private Sub grh_remove(ByVal list_index As Long, ByVal undo As Boolean)
    store_action grh, Remove, grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, , grh_list(list_index).grh(1).layer, , , , , , , , , , True, _
                grh_list(list_index).grh(1).subtile
    
    Dim Render_On_Top As Boolean
    If grh_list(list_index).grh(1).layer = 2 Then
        Render_On_Top = False
    Else
        Render_On_Top = True
    End If
    
    If undo Then
        If grh_list(list_index).grh(1).layer = 1 Or grh_list(list_index).grh(1).layer = 3 Or grh_list(list_index).grh(1).layer = 5 Then
            Engine.Map_Grh_Set grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, _
                                grh_list(list_index).grh(1).index, (grh_list(list_index).grh(1).layer + 1) / 2, _
                                grh_list(list_index).grh(1).alpha_blending, grh_list(list_index).grh(1).angle
        Else
            Engine.Map_Decoration_Add grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, grh_list(list_index).grh(1).subtile, grh_list(list_index).grh(1).index, _
                                        Render_On_Top, grh_list(list_index).grh(1).alpha_blending, grh_list(list_index).grh(1).angle, grh_list(list_index).grh(1).h_centered, _
                                        grh_list(list_index).grh(1).v_centered
            
        End If
    Else
        If grh_list(list_index).grh(1).layer = 1 Or grh_list(list_index).grh(1).layer = 3 Or grh_list(list_index).grh(1).layer = 5 Then
            Engine.Map_Grh_UnSet grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, _
                                grh_list(list_index).grh(1).layer
        Else
            Engine.Map_Decoration_Remove grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, grh_list(list_index).grh(1).subtile
        End If
    End If
End Sub

Private Sub grh_remove_all(ByVal list_index As Long, ByVal undo As Boolean)
    Dim index As Long
    Dim LoopC As Long
    Dim LoopC2 As Long
    Dim Render_On_Top As Boolean
    
    store_action grh, Remove_all, grh_list(list_index).grh(1).map_x, grh_list(list_index).grh(1).map_y, , , , , , , , , , , True, , grh_list(list_index).grh(1).subtile
    
    For LoopC = 1 To 5 Step 2
        index = index + 1
        If undo Then
            Engine.Map_Grh_Set grh_list(list_index).grh(index).map_x, grh_list(list_index).grh(index).map_y, grh_list(list_index).grh(index).index, _
                                index, grh_list(list_index).grh(index).alpha_blending, grh_list(list_index).grh(index).angle, grh_list(list_index).grh(index).h_centered, _
                                grh_list(list_index).grh(index).v_centered
        Else
            Engine.Map_Grh_UnSet grh_list(list_index).grh(index).map_x, grh_list(list_index).grh(index).map_y, index
        End If
    Next LoopC
    
    For LoopC = 1 To 2
        For LoopC2 = 0 To 8
            index = index + 1
            If undo Then
                Engine.Map_Decoration_Add grh_list(list_index).grh(index).map_x, grh_list(list_index).grh(index).map_y, grh_list(list_index).grh(index).subtile, _
                                            grh_list(list_index).grh(index).index, Render_On_Top, grh_list(list_index).grh(index).alpha_blending, grh_list(list_index).grh(index).angle, _
                                            grh_list(list_index).grh(index).h_centered, grh_list(list_index).grh(index).v_centered
            Else
                Engine.Map_Decoration_Remove grh_list(list_index).grh(index).map_x, grh_list(list_index).grh(index).map_y, grh_list(list_index).grh(index).subtile
            End If
        Next LoopC2
        Render_On_Top = True
    Next LoopC
End Sub

Private Sub grh_fill(ByVal list_index As Long, ByVal undo As Boolean)
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim layer As Byte
    
    store_action grh, fill, , , , grh_list(list_index).grh(1).layer, , , , , , , , , True
    
    layer = grh_list(list_index).grh(1).layer
    
    If undo Then
        Engine.Map_Bounds_Get max_x, max_y
        
        For map_x = 1 To max_x
            For map_y = 1 To max_y
                'If index doesn´t exist (e.g. 0), we place nothing on the place
                If Not Engine.Map_Grh_Set(map_x, map_y, grh_list(list_index).grh(max_x * (map_y - 1) + map_x).index, layer, grh_list(list_index).grh(max_x * (map_y - 1) + map_x).alpha_blending, _
                                            grh_list(list_index).grh(max_x * (map_y - 1) + map_x).angle) Then
                    Engine.Map_Grh_UnSet map_x, map_y, layer
                End If
            Next map_y
        Next map_x
    Else
        Engine.Map_Fill grh_list(list_index).grh(1).index, layer, , grh_list(list_index).grh(1).alpha_blending, grh_list(list_index).grh(1).angle
    End If
End Sub

Private Sub light_place(ByVal list_index As Long, ByVal undo As Boolean)
    If undo Then
        Engine.light_remove light_list(list_index).Light(1).index
    Else
        Engine.Light_Create light_list(list_index).Light(1).map_x, light_list(list_index).Light(1).map_y, light_list(list_index).Light(1).color, _
                                    light_list(list_index).Light(1).range
    End If
    
    'Since nothing was stored, we increase last_action to avoid errors
    last_action = last_action + 1
End Sub

Private Sub light_remove(ByVal list_index As Long, ByVal undo As Boolean)
    If undo Then
        Engine.Light_Create light_list(list_index).Light(1).map_x, light_list(list_index).Light(1).map_y, _
                light_list(list_index).Light(1).color, light_list(list_index).Light(1).range
    Else
        Engine.light_remove Engine.Map_Light_Get(light_list(list_index).Light(1).map_x, light_list(list_index).Light(1).map_y)
    End If
    
    'Since nothing was stored, we increase last_action to avoid errors
    last_action = last_action + 1
End Sub

Private Sub light_remove_all(ByVal list_index As Long, ByVal undo As Boolean)
    Dim LoopC As Long

    If undo Then
        For LoopC = 1 To UBound(light_list(list_index).Light)
            Engine.Light_Create light_list(list_index).Light(LoopC).map_x, light_list(list_index).Light(LoopC).map_y, _
                    light_list(list_index).Light(LoopC).color, light_list(list_index).Light(LoopC).range
        Next LoopC
    Else
        Engine.light_remove_all
    End If
    
    'Since nothing was stored, we increase last_action to avoid errors
    last_action = last_action + 1
End Sub

Private Sub base_light_place(ByVal list_index As Long)
    Dim LoopC As Long
    
    store_action lights, place, light_list(list_index).base_light(1).map_x, light_list(list_index).base_light(1).map_y, , , , , , base_light, , , , , True
    
    For LoopC = 0 To 3
        Engine.Map_Base_Light_Set light_list(list_index).base_light(1).map_x, light_list(list_index).base_light(1).map_y, light_list(list_index).base_light(1).color(LoopC), LoopC
    Next LoopC
End Sub

Private Sub base_light_fill(ByVal list_index As Long, ByVal undo As Boolean)
    Dim LoopC As Long
    Dim corner As Long
    
    If undo Then
        For LoopC = 1 To UBound(light_list(list_index).base_light)
            For corner = 0 To 3
                Engine.Map_Base_Light_Set light_list(list_index).base_light(LoopC).map_x, light_list(list_index).base_light(LoopC).map_y, _
                                        light_list(list_index).base_light(LoopC).color(corner), corner
            Next corner
        Next LoopC
    Else
        Engine.Map_Base_Light_Fill light_list(list_index).fill_color
    End If
    
    'Since nothing was stored, we increase last_action to avoid errors
    last_action = last_action + 1
End Sub

Private Sub shadow_place(ByVal list_index As Long)
    Dim LoopC As Long
    
    For LoopC = 1 To UBound(light_list(list_index).shadow)
        store_action lights, place, light_list(list_index).shadow(LoopC).map_x, light_list(list_index).shadow(LoopC).map_y, LoopC, , , , , , light_list(list_index).shadow(LoopC).corner, shadow, , , , , True
        Engine.Map_Base_Light_Set light_list(list_index).shadow(LoopC).map_x, light_list(list_index).shadow(LoopC).map_y, light_list(list_index).shadow(LoopC).color, light_list(list_index).shadow(LoopC).corner
    Next LoopC
End Sub

Private Sub npc_place(ByVal list_index As Long, ByVal undo As Boolean)
    store_action NPC, place, npc_list(list_index).NPC(1).map_x, npc_list(list_index).NPC(1).map_y, , , , , , , , , , , True
    
    If undo Then
        Engine.Map_NPC_Remove npc_list(list_index).NPC(1).map_x, npc_list(list_index).NPC(1).map_y
    Else
        Engine.Map_NPC_Add npc_list(list_index).NPC(1).map_x, npc_list(list_index).NPC(1).map_y, npc_list(list_index).NPC(1).index
    End If
End Sub

Private Sub npc_remove(ByVal list_index As Long, ByVal undo As Boolean)
    If undo Then
        Engine.Map_NPC_Add npc_list(list_index).NPC(1).map_x, npc_list(list_index).NPC(1).map_y, _
                            npc_list(list_index).NPC(1).index
    Else
        Engine.Map_NPC_Remove npc_list(list_index).NPC(1).map_x, npc_list(list_index).NPC(1).map_y
    End If
    
    last_action = last_action + 1
End Sub

Private Sub npc_remove_all(ByVal list_index As Long, ByVal undo As Boolean)
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    If undo Then
        For LoopC = 1 To UBound(npc_list(list_index).NPC)
            Engine.Map_NPC_Add npc_list(list_index).NPC(LoopC).map_x, npc_list(list_index).NPC(LoopC).map_y, _
                                npc_list(list_index).NPC(LoopC).index
        Next LoopC
    Else
        Engine.Map_Bounds_Get max_x, max_y
    
        For map_x = 1 To max_x
            For map_y = 1 To max_y
                Engine.Map_NPC_Remove map_x, map_y
            Next map_y
        Next map_x
    End If
    
    last_action = last_action + 1
End Sub

Private Sub object_place(ByVal list_index As Long, ByVal undo As Boolean)
    store_action object, place, objects_list(list_index).object(1).map_x, objects_list(list_index).object(1).map_y, , , , , , , , , , , True
    
    If undo Then
        Engine.Map_Item_Remove objects_list(list_index).object(1).map_x, objects_list(list_index).object(1).map_y
    Else
        Engine.Map_Item_Add objects_list(list_index).object(1).map_x, objects_list(list_index).object(1).map_y, objects_list(list_index).object(1).index, objects_list(list_index).object(1).amount
    End If
End Sub

Private Sub object_remove(ByVal list_index As Long, ByVal undo As Boolean)
    If undo Then
        Engine.Map_Item_Add objects_list(list_index).object(1).map_x, objects_list(list_index).object(1).map_y, _
                            objects_list(list_index).object(1).index, objects_list(list_index).object(1).amount
    Else
        Engine.Map_Item_Remove objects_list(list_index).object(1).map_x, objects_list(list_index).object(1).map_y
    End If
    
    last_action = last_action + 1
End Sub

Private Sub object_remove_all(ByVal list_index As Long, ByVal undo As Boolean)
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    If undo Then
        For LoopC = 1 To UBound(objects_list(list_index).object)
            Engine.Map_Item_Add objects_list(list_index).object(LoopC).map_x, objects_list(list_index).object(LoopC).map_y, _
                                objects_list(list_index).object(LoopC).index, objects_list(list_index).object(LoopC).amount
        Next LoopC
    Else
        Engine.Map_Bounds_Get max_x, max_y
        For map_x = 1 To max_x
            For map_y = 1 To max_y
                Engine.Map_Item_Remove map_x, map_y
            Next map_y
        Next map_x
    End If
    
    last_action = last_action + 1
End Sub

Private Sub particle_stream_place(ByVal list_index As Long, ByVal undo As Boolean)
    Dim rgb_list(0 To 3) As Long
    Dim PIndex As Long
    
    store_action particle_stream, place, particle_stream_list(list_index).particle_stream(1).map_x, particle_stream_list(list_index).particle_stream(1).map_y, , , , , , , , , , , True
    
    If undo Then
        Engine.Particle_Group_Remove Engine.Map_Particle_Group_Get(particle_stream_list(list_index).particle_stream(1).map_x, particle_stream_list(list_index).particle_stream(1).map_y)
    Else
        PIndex = particle_stream_list(list_index).particle_stream(1).index
        
        rgb_list(0) = RGB(StreamData(PIndex).colortint(0).r, StreamData(PIndex).colortint(0).g, StreamData(PIndex).colortint(0).b)
        rgb_list(1) = RGB(StreamData(PIndex).colortint(1).r, StreamData(PIndex).colortint(1).g, StreamData(PIndex).colortint(1).b)
        rgb_list(2) = RGB(StreamData(PIndex).colortint(2).r, StreamData(PIndex).colortint(2).g, StreamData(PIndex).colortint(2).b)
        rgb_list(3) = RGB(StreamData(PIndex).colortint(3).r, StreamData(PIndex).colortint(3).g, StreamData(PIndex).colortint(3).b)
        
        Engine.Particle_Group_Create particle_stream_list(list_index).particle_stream(1).map_x, particle_stream_list(list_index).particle_stream(1).map_y, _
                    StreamData(PIndex).grh_list, rgb_list(), StreamData(PIndex).NumOfParticles, 3, StreamData(PIndex).AlphaBlend, StreamData(PIndex).life_counter, StreamData(PIndex).speed, , StreamData(PIndex).x1, _
                    StreamData(PIndex).y1, StreamData(PIndex).angle, StreamData(PIndex).vecx1, StreamData(PIndex).vecx2, StreamData(PIndex).vecy1, _
                    StreamData(PIndex).vecy2, StreamData(PIndex).life1, StreamData(PIndex).life2, StreamData(PIndex).friction, StreamData(PIndex).spin_speedL, _
                    StreamData(PIndex).gravity, StreamData(PIndex).grav_strength, StreamData(PIndex).bounce_strength, StreamData(PIndex).x2, _
                    StreamData(PIndex).y2, StreamData(PIndex).XMove, StreamData(PIndex).move_x1, StreamData(PIndex).move_x2, StreamData(PIndex).move_y1, _
                    StreamData(PIndex).move_y2, StreamData(PIndex).YMove, StreamData(PIndex).spin_speedH, StreamData(PIndex).spin
    End If
End Sub

Private Sub particle_stream_remove(ByVal list_index As Long, ByVal undo As Boolean)
    Dim rgb_list(0 To 3) As Long
    Dim PIndex As Long
    
    store_action particle_stream, Remove, particle_stream_list(list_index).particle_stream(1).map_x, particle_stream_list(list_index).particle_stream(1).map_y, , , , , , , , , , , True
    
    If undo Then
        PIndex = particle_stream_list(list_index).particle_stream(1).index
        
        rgb_list(0) = RGB(StreamData(PIndex).colortint(0).r, StreamData(PIndex).colortint(0).g, StreamData(PIndex).colortint(0).b)
        rgb_list(1) = RGB(StreamData(PIndex).colortint(1).r, StreamData(PIndex).colortint(1).g, StreamData(PIndex).colortint(1).b)
        rgb_list(2) = RGB(StreamData(PIndex).colortint(2).r, StreamData(PIndex).colortint(2).g, StreamData(PIndex).colortint(2).b)
        rgb_list(3) = RGB(StreamData(PIndex).colortint(3).r, StreamData(PIndex).colortint(3).g, StreamData(PIndex).colortint(3).b)
        
        Engine.Particle_Group_Create particle_stream_list(list_index).particle_stream(1).map_x, particle_stream_list(list_index).particle_stream(1).map_y, _
                    StreamData(PIndex).grh_list, rgb_list(), StreamData(PIndex).NumOfParticles, 3, StreamData(PIndex).AlphaBlend, , , , StreamData(PIndex).x1, _
                    StreamData(PIndex).y1, StreamData(PIndex).angle, StreamData(PIndex).vecx1, StreamData(PIndex).vecx2, StreamData(PIndex).vecy1, _
                    StreamData(PIndex).vecy2, StreamData(PIndex).life1, StreamData(PIndex).life2, StreamData(PIndex).friction, StreamData(PIndex).spin_speedL, _
                    StreamData(PIndex).gravity, StreamData(PIndex).grav_strength, StreamData(PIndex).bounce_strength, StreamData(PIndex).x2, _
                    StreamData(PIndex).y2, StreamData(PIndex).XMove, StreamData(PIndex).move_x1, StreamData(PIndex).move_x2, StreamData(PIndex).move_y1, _
                    StreamData(PIndex).move_y2, StreamData(PIndex).YMove, StreamData(PIndex).spin_speedH, StreamData(PIndex).spin
    Else
        Engine.Particle_Group_Remove Engine.Map_Particle_Group_Get(particle_stream_list(list_index).particle_stream(1).map_x, particle_stream_list(list_index).particle_stream(1).map_y)
    End If
End Sub

Private Sub particle_stream_remove_all(ByVal list_index As Long, ByVal undo As Boolean)
    Dim LoopC As Long
    Dim PIndex As Long
    
    store_action particle_stream, Remove_all, , , , , , , , , , , , , True
    
    If undo Then
        For LoopC = 1 To UBound(particle_stream_list(list_index).particle_stream)
            PIndex = particle_stream_list(list_index).particle_stream(LoopC).index
            
            Dim rgb_list(0 To 3) As Long
            rgb_list(0) = RGB(StreamData(PIndex).colortint(0).r, StreamData(PIndex).colortint(0).g, StreamData(PIndex).colortint(0).b)
            rgb_list(1) = RGB(StreamData(PIndex).colortint(1).r, StreamData(PIndex).colortint(1).g, StreamData(PIndex).colortint(1).b)
            rgb_list(2) = RGB(StreamData(PIndex).colortint(2).r, StreamData(PIndex).colortint(2).g, StreamData(PIndex).colortint(2).b)
            rgb_list(3) = RGB(StreamData(PIndex).colortint(3).r, StreamData(PIndex).colortint(3).g, StreamData(PIndex).colortint(3).b)
                
            Engine.Particle_Group_Create particle_stream_list(list_index).particle_stream(LoopC).map_x, particle_stream_list(list_index).particle_stream(LoopC).map_y, _
                                        StreamData(PIndex).grh_list, rgb_list, StreamData(PIndex).NumOfParticles, 3, StreamData(PIndex).AlphaBlend, , , , StreamData(PIndex).x1, _
                                        StreamData(PIndex).y1, StreamData(PIndex).angle, StreamData(PIndex).vecx1, StreamData(PIndex).vecx2, StreamData(PIndex).vecy1, StreamData(PIndex).vecy2, StreamData(PIndex).life1, _
                                        StreamData(PIndex).life2, StreamData(PIndex).friction, StreamData(PIndex).spin_speedL, StreamData(PIndex).gravity, StreamData(PIndex).grav_strength, StreamData(PIndex).bounce_strength, _
                                        StreamData(PIndex).x2, StreamData(PIndex).y2, StreamData(PIndex).XMove, StreamData(PIndex).move_x1, StreamData(PIndex).move_x2, StreamData(PIndex).move_y1, StreamData(PIndex).move_y2, _
                                        StreamData(PIndex).YMove, StreamData(PIndex).spin_speedH, StreamData(PIndex).spin
        Next LoopC
    Else
        Engine.Particle_Group_Remove_All
    End If
End Sub

Private Sub trigger_place(ByVal list_index As Long)
    store_action trigger, place, trigger_list(list_index).trigger(1).map_x, trigger_list(list_index).trigger(1).map_y, , , , , , , , , , , , , True
    
    Engine.Map_Trigger_Set trigger_list(list_index).trigger(1).map_x, trigger_list(list_index).trigger(1).map_y, trigger_list(list_index).trigger(1).index
End Sub

Private Sub trigger_remove(ByVal list_index As Long, ByVal undo As Boolean)
    store_action trigger, Remove, trigger_list(list_index).trigger(1).map_x, trigger_list(list_index).trigger(1).map_y, , , , , , , , , , , , , True
    
    If undo Then
        Engine.Map_Trigger_Set trigger_list(list_index).trigger(1).map_x, trigger_list(list_index).trigger(1).map_y, trigger_list(list_index).trigger(1).index
    Else
        Engine.Map_Trigger_Unset trigger_list(list_index).trigger(1).map_x, trigger_list(list_index).trigger(1).map_y
    End If
End Sub

Private Sub trigger_remove_all(ByVal list_index As Long, ByVal undo As Boolean)
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    store_action trigger, Remove_all, , , , , , , , , , , , , , , True
    
    If undo Then
        For LoopC = 1 To UBound(trigger_list(list_index).trigger)
            Engine.Map_Trigger_Set trigger_list(list_index).trigger(LoopC).map_x, _
                    trigger_list(list_index).trigger(LoopC).map_y, trigger_list(list_index).trigger(LoopC).index
        Next LoopC
    Else
        Engine.Map_Bounds_Get max_x, max_y
    
        For map_x = 1 To max_x
            For map_y = 1 To max_y
                Engine.Map_Trigger_Unset map_x, map_y
            Next map_y
        Next map_x
    End If
End Sub

Private Sub trigger_fill(ByVal list_index As Long, ByVal undo As Boolean)
    Dim map_x As Long
    Dim map_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim LoopC As Long
    
    store_action trigger, fill, , , , , , , , , , , , , , , True
    
    If undo Then
        For LoopC = 1 To UBound(trigger_list(list_index).trigger)
            Engine.Map_Trigger_Set trigger_list(list_index).trigger(LoopC).map_x, _
                    trigger_list(list_index).trigger(LoopC).map_y, trigger_list(list_index).trigger(LoopC).index
        Next LoopC
    Else
        Engine.Map_Bounds_Get max_x, max_y
        
        For map_x = 1 To max_x
            For map_y = 1 To max_y
                Engine.Map_Trigger_Set map_x, map_y, trigger_list(list_index).trigger(1).index
            Next map_y
        Next map_x
    End If
End Sub

Private Sub destroy_grh(ByVal index As Long)
    Dim LoopC As Long

    If Not index = last_grh Then
        'Move backwards the list to save space
        For LoopC = index + 1 To last_grh
            grh_list(LoopC - 1) = grh_list(LoopC)
        Next LoopC
    End If
    
    last_grh = last_grh - 1
    If last_grh > 0 Then
        ReDim Preserve grh_list(1 To last_grh)
    Else
        ReDim grh_list(0)
    End If
    
    'Correct all other indexes in the action list
    For LoopC = 1 To MAX_ACTIONS
        If action_list(LoopC).tool = grh And action_list(LoopC).grh_list_index > index Then
            action_list(LoopC).grh_list_index = action_list(LoopC).grh_list_index - 1
        End If
    Next LoopC
End Sub

Private Sub destroy_light(ByVal index As Long)
    Dim LoopC As Long

    If Not index = last_light Then
        'Move backwards the list to save space
        For LoopC = index + 1 To last_light
            light_list(LoopC - 1) = light_list(LoopC)
        Next LoopC
    End If
    
    last_light = last_light - 1
    If last_light > 0 Then
        ReDim Preserve light_list(1 To last_light)
    Else
        ReDim light_list(0)
    End If
    
    'Correct all other indexes in the action list
    For LoopC = 1 To MAX_ACTIONS
        If action_list(LoopC).tool = lights And action_list(LoopC).light_list_index > index Then
            action_list(LoopC).light_list_index = action_list(LoopC).light_list_index - 1
        End If
    Next LoopC
End Sub

Private Sub destroy_particle_stream(ByVal index As Long)
    Dim LoopC As Long
    
    If Not index = last_particle_stream Then
        'Move backwards the list to save space
        For LoopC = index + 1 To last_particle_stream
            particle_stream_list(LoopC - 1) = particle_stream_list(LoopC)
        Next LoopC
    End If
    
    last_particle_stream = last_particle_stream - 1
    If last_particle_stream > 0 Then
        ReDim Preserve particle_stream_list(1 To last_particle_stream)
    Else
        ReDim particle_stream_list(0)
    End If
    
    'Correct all other indexes in the action list
    For LoopC = 1 To MAX_ACTIONS
        If action_list(LoopC).tool = particle_stream And action_list(LoopC).particle_streams_list_index > index Then
            action_list(LoopC).particle_streams_list_index = action_list(LoopC).particle_streams_list_index - 1
        End If
    Next LoopC
End Sub

Private Sub destroy_npc(ByVal index As Long)
    Dim LoopC As Long

    If Not index = last_npc Then
        'Move backwards the list to save space
        For LoopC = index + 1 To last_npc
            npc_list(LoopC - 1) = npc_list(LoopC)
        Next LoopC
    End If
    
    last_npc = last_npc - 1
    If last_npc > 0 Then
        ReDim Preserve npc_list(1 To last_npc)
    Else
        ReDim npc_list(0)
    End If
    
    'Correct all other indexes in the action list
    For LoopC = 1 To MAX_ACTIONS
        If action_list(LoopC).tool = NPC And action_list(LoopC).npc_list_index > index Then
            action_list(LoopC).npc_list_index = action_list(LoopC).npc_list_index - 1
        End If
    Next LoopC
End Sub

Private Sub destroy_object(ByVal index As Long)
    Dim LoopC As Long

    If Not index = last_object Then
        'Move backwards the list to save space
        For LoopC = index + 1 To last_object
            objects_list(LoopC - 1) = objects_list(LoopC)
        Next LoopC
    End If
    
    last_object = last_object - 1
    If last_object > 0 Then
        ReDim Preserve objects_list(1 To last_object)
    Else
        ReDim objects_list(0)
    End If
    
    'Correct all other indexes in the action list
    For LoopC = 1 To MAX_ACTIONS
        If action_list(LoopC).tool = object And action_list(LoopC).object_list_index > index Then
            action_list(LoopC).object_list_index = action_list(LoopC).object_list_index - 1
        End If
    Next LoopC
End Sub

Private Sub destroy_exits(ByVal index As Long)
    Dim LoopC As Long
    
    If Not index = last_exit Then
        'Move backwards the list to save space
        For LoopC = index + 1 To last_exit
            exits_list(LoopC - 1) = exits_list(LoopC)
        Next LoopC
    End If
    
    last_exit = last_exit - 1
    If last_exit > 0 Then
        ReDim Preserve exits_list(1 To last_exit)
    Else
        ReDim exits_list(0)
    End If
    
    'Correct all other indexes in the action list
    For LoopC = 1 To MAX_ACTIONS
        If action_list(LoopC).tool = exits And action_list(LoopC).exits_list_index > index Then
            action_list(LoopC).exits_list_index = action_list(LoopC).exits_list_index - 1
        End If
    Next LoopC
End Sub

Private Sub destroy_trigger(ByVal index As Long)
    Dim LoopC As Long
    
    If Not index = last_trigger Then
        'Move backwards the list to save space
        For LoopC = index + 1 To last_trigger
            trigger_list(LoopC - 1) = trigger_list(LoopC)
        Next LoopC
    End If
    
    last_trigger = last_trigger - 1
    If last_trigger > 0 Then
        ReDim Preserve trigger_list(1 To last_trigger)
    Else
        ReDim exits_list(0)
    End If
    
    'Correct all other indexes in the action list
    For LoopC = 1 To MAX_ACTIONS
        If action_list(LoopC).tool = trigger And action_list(LoopC).triggers_list_index > index Then
            action_list(LoopC).triggers_list_index = action_list(LoopC).triggers_list_index - 1
        End If
    Next LoopC
End Sub

Private Sub destroy_blocking(ByVal index As Long)
    Dim LoopC As Long

    If Not index = last_block Then
        'Move backwards the list to save space
        For LoopC = index + 1 To last_block
            block_list(LoopC - 1) = block_list(LoopC)
        Next LoopC
    End If
    
    last_block = last_block - 1
    If last_block > 0 Then
        ReDim Preserve block_list(1 To last_block)
    Else
        ReDim block_list(0)
    End If
    
    'Correct all other indexes in the action list
    For LoopC = 1 To MAX_ACTIONS
        If action_list(LoopC).tool = blocking And action_list(LoopC).blocked_list_index > index Then
            action_list(LoopC).blocked_list_index = action_list(LoopC).blocked_list_index - 1
        End If
    Next LoopC
End Sub
