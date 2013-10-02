Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, X As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr25 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If GetTickCount > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.Target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.Target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For X = 1 To MAX_DOTS
                        HandleDoT_Player i, X
                        HandleHoT_Player i, X
                    Next
                End If
            Next
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr25 = GetTickCount + 25
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            For i = 1 To MAX_EDITORS
                If frmServer.EditorSocket(i).State > sckConnected Then
                    Call CloseEditorSocket(i)
                End If
            Next
            UpdateMapLogic
            tmr500 = GetTickCount + 500
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            ' If scripting is enabled, run our little OnServerTime script.
            ' It'll let people run things on specific seconds, minutes and hours.
            If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnServerTime " & Trim$(Hour(Now)) & "," & Trim$(Minute(Now)) & "," & Trim$(Second(Now))
            
            tmr1000 = GetTickCount + 1000
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 300000
        End If
        
        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub UpdateMapSpawnItems()
    Dim X As Long
    Dim Y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For Y = 1 To MAX_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(Y) Then

            ' Clear out unnecessary junk
            For X = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(X, Y)
            Next

            ' Spawn the items
            Call SpawnMapItems(Y)
            Call SendMapItemsToAll(Y)
        End If

        DoEvents
    Next

End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, X As Long, MapNum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, NPCNum As Long
    Dim Target As Long, TargetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean

    For MapNum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(MapNum, i).Num > 0 Then
                If MapItem(MapNum, i).playerName <> vbNullString Then
                    ' make item public?
                    If MapItem(MapNum, i).playerTimer < GetTickCount Then
                        ' make it public
                        MapItem(MapNum, i).playerName = vbNullString
                        MapItem(MapNum, i).playerTimer = 0
                        ' send updates to everyone
                        SendMapItemsToAll MapNum
                    End If
                    ' despawn item?
                    If MapItem(MapNum, i).canDespawn Then
                        If MapItem(MapNum, i).despawnTimer < GetTickCount Then
                            ' despawn it
                            ClearMapItem i, MapNum
                            ' send updates to everyone
                            SendMapItemsToAll MapNum
                        End If
                    End If
                End If
            End If
        Next
        
        '  Close the doors
        If TickCount > TempTile(MapNum).DoorTimer + 5000 Then
            For x1 = 0 To Map(MapNum).MaxX
                For y1 = 0 To Map(MapNum).MaxY
                    If Map(MapNum).Tile(x1, y1).Type = TileTypeKey And TempTile(MapNum).DoorOpen(x1, y1) = YES Then
                        TempTile(MapNum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap MapNum, x1, y1, 0
                    End If
                Next
            Next
        End If
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(MapNum).Npc(i).Num > 0 Then
                For X = 1 To MAX_DOTS
                    HandleDoT_Npc MapNum, i, X
                    HandleHoT_Npc MapNum, i, X
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(MapNum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(MapNum).Resource_Count
                Resource_index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(i).X, ResourceCache(MapNum).ResourceData(i).Y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(MapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(MapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(MapNum).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(MapNum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(MapNum).ResourceData(i).cur_health = Resource(Resource_index).health
                            SendResourceCacheToMap MapNum, i
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(MapNum) = YES Then
            TickCount = GetTickCount
            
            For X = 1 To MAX_MAP_NPCS
                NPCNum = MapNpc(MapNum).Npc(X).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NPCNum).Behaviour = NPCTypeAggressive Or Npc(NPCNum).Behaviour = NPCTypeProtectAllies Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(MapNum).Npc(X).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And MapNpc(MapNum).Npc(X).Target = 0 And GetPlayerAccess(i) <= RankModerator Then
                                        n = Npc(NPCNum).Range
                                        DistanceX = MapNpc(MapNum).Npc(X).X - GetPlayerX(i)
                                        DistanceY = MapNpc(MapNum).Npc(X).Y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If Npc(NPCNum).Behaviour = NPCTypeAggressive Or GetPlayerPK(i) = YES Then
                                                If Len(Trim$(Npc(NPCNum).AttackSay)) > 0 Then
                                                    Call PlayerMsg(i, Trim$(Npc(NPCNum).Name) & " says: " & Trim$(Npc(NPCNum).AttackSay), SayColor)
                                                End If
                                                MapNpc(MapNum).Npc(X).TargetType = 1 ' player
                                                MapNpc(MapNum).Npc(X).Target = i
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then
                    If MapNpc(MapNum).Npc(X).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(MapNum).Npc(X).StunTimer + (MapNpc(MapNum).Npc(X).StunDuration * 1000) Then
                            MapNpc(MapNum).Npc(X).StunDuration = 0
                            MapNpc(MapNum).Npc(X).StunTimer = 0
                        End If
                    Else
                            
                        Target = MapNpc(MapNum).Npc(X).Target
                        TargetType = MapNpc(MapNum).Npc(X).TargetType
    
                        ' Check to see if its time for the npc to walk
                        If Npc(NPCNum).Behaviour <> NPCTypeStationary Then
                        
                            If TargetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If Target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = GetPlayerY(Target)
                                        TargetX = GetPlayerX(Target)
                                    Else
                                        MapNpc(MapNum).Npc(X).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(X).Target = 0
                                    End If
                                End If
                            
                            ElseIf TargetType = 2 Then 'npc
                                
                                If Target > 0 Then
                                    
                                    If MapNpc(MapNum).Npc(Target).Num > 0 Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = MapNpc(MapNum).Npc(Target).Y
                                        TargetX = MapNpc(MapNum).Npc(Target).X
                                    Else
                                        MapNpc(MapNum).Npc(X).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(X).Target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                
                                i = Int(Rnd * 5)
    
                                ' Lets move the npc
                                Select Case i
                                    Case 0
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, North) Then
                                                Call NpcMove(MapNum, X, North, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, South) Then
                                                Call NpcMove(MapNum, X, South, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, West) Then
                                                Call NpcMove(MapNum, X, West, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, East) Then
                                                Call NpcMove(MapNum, X, East, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 1
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, East) Then
                                                Call NpcMove(MapNum, X, East, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, West) Then
                                                Call NpcMove(MapNum, X, West, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, South) Then
                                                Call NpcMove(MapNum, X, South, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, North) Then
                                                Call NpcMove(MapNum, X, North, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 2
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, South) Then
                                                Call NpcMove(MapNum, X, South, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, North) Then
                                                Call NpcMove(MapNum, X, North, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, East) Then
                                                Call NpcMove(MapNum, X, East, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, West) Then
                                                Call NpcMove(MapNum, X, West, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 3
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, West) Then
                                                Call NpcMove(MapNum, X, West, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, East) Then
                                                Call NpcMove(MapNum, X, East, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, North) Then
                                                Call NpcMove(MapNum, X, North, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, South) Then
                                                Call NpcMove(MapNum, X, South, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                End Select
    
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(MapNum).Npc(X).X - 1 = TargetX And MapNpc(MapNum).Npc(X).Y = TargetY Then
                                        If MapNpc(MapNum).Npc(X).Dir <> West Then
                                            Call NpcDir(MapNum, X, West)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(MapNum).Npc(X).X + 1 = TargetX And MapNpc(MapNum).Npc(X).Y = TargetY Then
                                        If MapNpc(MapNum).Npc(X).Dir <> East Then
                                            Call NpcDir(MapNum, X, East)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(MapNum).Npc(X).X = TargetX And MapNpc(MapNum).Npc(X).Y - 1 = TargetY Then
                                        If MapNpc(MapNum).Npc(X).Dir <> North Then
                                            Call NpcDir(MapNum, X, North)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(MapNum).Npc(X).X = TargetX And MapNpc(MapNum).Npc(X).Y + 1 = TargetY Then
                                        If MapNpc(MapNum).Npc(X).Dir <> South Then
                                            Call NpcDir(MapNum, X, South)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    ' We could not move so Target must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
    
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
    
                                            If CanNpcMove(MapNum, X, i) Then
                                                Call NpcMove(MapNum, X, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
    
                            Else
                                i = Int(Rnd * 4)
    
                                If i = 1 Then
                                    i = Int(Rnd * 4)
    
                                    If CanNpcMove(MapNum, X, i) Then
                                        Call NpcMove(MapNum, X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then
                    Target = MapNpc(MapNum).Npc(X).Target
                    TargetType = MapNpc(MapNum).Npc(X).TargetType

                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                    
                        If TargetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                TryNpcAttackPlayer X, Target
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(MapNum).Npc(X).Target = 0
                                MapNpc(MapNum).Npc(X).TargetType = 0 ' clear
                            End If
                        Else
                            ' lol no npc combat :(
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(MapNum).Npc(X).stopRegen Then
                    If MapNpc(MapNum).Npc(X).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(MapNum).Npc(X).Vital(Vitals.HP) > 0 Then
                            MapNpc(MapNum).Npc(X).Vital(Vitals.HP) = MapNpc(MapNum).Npc(X).Vital(Vitals.HP) + GetNpcVitalRegen(NPCNum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(MapNum).Npc(X).Vital(Vitals.HP) > GetNPCMaxVital(NPCNum, Vitals.HP) Then
                                MapNpc(MapNum).Npc(X).Vital(Vitals.HP) = GetNPCMaxVital(NPCNum, Vitals.HP)
                            End If
                        End If
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(MapNum).Npc(X).Num = 0 And Map(MapNum).Npc(X) > 0 Then
                    If TickCount > MapNpc(MapNum).Npc(X).SpawnWait + (Npc(Map(MapNum).Npc(X)).SpawnSecs * 1000) Then
                        Call SpawnNpc(X, MapNum)
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
    
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.MP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call SavePlayer(i)
                Call SaveBank(i)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub
