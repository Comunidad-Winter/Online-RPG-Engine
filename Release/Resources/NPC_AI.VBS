'***********************************************************
'NPC AI script file
'
'Events:
'***********************************************************

Sub NPC_AI_Random(npc_id)
	'Make movement timer
	If ORE.NPC_Session_Check(npc_id,"move_timer") = False Then
		ORE.NPC_Session_Create npc_id,"move_timer",0
	End If

	'Get move timer
	move_timer = ORE.NPC_Session_Get(npc_id,"move_timer")

	'Move when timer reaches 20 ticks
	If move_timer = 20 Then
		'Reset timer
		move_timer = 0
		
		'Move NPC
		ORE.NPC_Move_By_Heading npc_id, ORE.Random_Number(1,8)
		
		'Look for player
		p_id = ORE.Map_Find_Nearest_Player(ORE.NPC_Map_Name_Get(npc_id), ORE.NPC_Map_X_Get(npc_id), ORE.NPC_Map_Y_Get(npc_id), 5)
		If p_id Then
			'Send that player a message
			ORE.Chat_To_Player p_id, ORE.NPC_Name_Get(npc_id) & ": Hey " & ORE.Player_Name_Get(p_id) & ", I see you ...", ORE.Player_Name_Get(p_id)
		End If
	Else
		'Increment timer
		move_timer = move_timer + 1
	End If

	'Set move timer
	ORE.NPC_Session_Set npc_id,"move_timer", move_timer
End Sub