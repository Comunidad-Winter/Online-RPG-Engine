'***********************************************************
'Main script file
'
'Events:
'Server_Start_Up()
'Server_Tick(server_ticks)
'Map_Tile_Exit_Activated(parms(4))
'Player_Move(parms(2))
'Player_Login(player_id)
'Player_Logout(player_id)
'Player_New(player_id)
'***********************************************************

Sub Server_Start_Up()
'******************************************
'Called when server start up
'******************************************
	'Create a NPC
	ORE.NPC_Create 1, 1, 11, 11
End Sub

Sub Server_Tick(server_ticks)
'******************************************
'Called every server tick
'******************************************
End Sub

Sub Map_Tile_Exit_Activated(parms())
'******************************************
'Called when a player activates a tile exit
'1 - player_id
'2 - exit_map_name
'3 - exit_map_x
'4 - exit_map_y
'******************************************
	ORE.Player_Position_Set parms(1), parms(2), parms(3), parms(4) 'Leave this line. This makes tile exits work!
End Sub

Sub Player_Move(parms())
'******************************************
'Called every time a player moves
'1 - player_id
'2 - heading
'******************************************
	'Save the players heading in a session variable
	If ORE.Player_Session_Create(parms(1),"LastHeading", parms(2), True) = False Then
		ORE.Player_Session_Set parms(1),"LastHeading", parms(2)
	End If
End Sub

Sub Player_Login(player_id)
'******************************************
'Called when a player logins
'******************************************
	'Tell the player what his last heading was
	'ORE.Chat_To_Player player_id, "Your last heading was: " & ORE.Player_Session_Get(player_id, "LastHeading")
End Sub

