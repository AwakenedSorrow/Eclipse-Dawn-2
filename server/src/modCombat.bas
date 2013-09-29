Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(Index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 15 + 150
                Case 2 ' Mage
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 5 + 65
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 15 + 150
            End Select
        Case MP
            Select Case GetPlayerClass(Index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 5 + 25
                Case 2 ' Mage
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 30 + 85
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 5 + 25
            End Select
    End Select
End Function

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (GetPlayerStat(Index, Stats.Willpower) * 0.8) + 6
        Case MP
            i = (GetPlayerStat(Index, Stats.Willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(Index, Weapon)
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) * Item(weaponNum).Data2 + (GetPlayerLevel(Index) / 5)
    Else
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) + (GetPlayerLevel(Index) / 5)
    End If

End Function

Function GetNpcMaxVital(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = Npc(NPCNum).HP
        Case MP
            GetNpcMaxVital = 30 + (Npc(NPCNum).Stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (Npc(NPCNum).Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (Npc(NPCNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal NPCNum As Long) As Long
    GetNpcDamage = 0.085 * 5 * Npc(NPCNum).Stat(Stats.Strength) * Npc(NPCNum).Damage + (Npc(NPCNum).Level / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(Index, Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(Index, Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False

    rate = GetPlayerStat(Index, Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcBlock = False

    rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = Npc(NPCNum).Stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = Npc(NPCNum).Stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = Npc(NPCNum).Stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal MapNPCNum As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, MapNPCNum) Then
    
        MapNum = GetPlayerMap(Index)
        NPCNum = MapNpc(MapNum).Npc(MapNPCNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).Npc(MapNPCNum).X * 32), (MapNpc(MapNum).Npc(MapNPCNum).Y * 32)
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (MapNpc(MapNum).Npc(MapNPCNum).X * 32), (MapNpc(MapNum).Npc(MapNPCNum).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(MapNPCNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (Npc(NPCNum).Stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, MapNPCNum, Damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal MapNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).Npc(MapNPCNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(attacker)
    NPCNum = MapNpc(MapNum).Npc(MapNPCNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
        ' exit out early
        If IsSpell Then
             If NPCNum > 0 Then
                If Npc(NPCNum).Behaviour <> NPCTypeFriendly And Npc(NPCNum).Behaviour <> NPCTypeStationary And Npc(NPCNum).Behaviour <> NPCTypeScripted Then
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If NPCNum > 0 And GetTickCount > TempPlayer(attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case North
                    NpcX = MapNpc(MapNum).Npc(MapNPCNum).X
                    NpcY = MapNpc(MapNum).Npc(MapNPCNum).Y + 1
                Case South
                    NpcX = MapNpc(MapNum).Npc(MapNPCNum).X
                    NpcY = MapNpc(MapNum).Npc(MapNPCNum).Y - 1
                Case West
                    NpcX = MapNpc(MapNum).Npc(MapNPCNum).X + 1
                    NpcY = MapNpc(MapNum).Npc(MapNPCNum).Y
                Case East
                    NpcX = MapNpc(MapNum).Npc(MapNPCNum).X - 1
                    NpcY = MapNpc(MapNum).Npc(MapNPCNum).Y
            End Select

            If NpcX = GetPlayerX(attacker) Then
                If NpcY = GetPlayerY(attacker) Then
                    If Npc(NPCNum).Behaviour <> NPCTypeFriendly And Npc(NPCNum).Behaviour <> NPCTypeStationary And Npc(NPCNum).Behaviour <> NPCTypeScripted Then
                        CanPlayerAttackNpc = True
                    Else
                        If Len(Trim$(Npc(NPCNum).AttackSay)) > 0 Then
                            PlayerMsg attacker, Trim$(Npc(NPCNum).Name) & ": " & Trim$(Npc(NPCNum).AttackSay), White
                        End If
                        
                        If Options.Scripting = 1 And Npc(NPCNum).Behaviour = NPCTypeScripted Then MyScript.ExecuteStatement "main.eds", "OnUseNPC " & Trim$(STR$(attacker)) & "," & Trim$(STR$(MapNum)) & "," & Trim$(STR$(MapNPCNum)) & "," & Trim$(STR$(NPCNum))
                        
                        ' Reset attack timer
                        TempPlayer(attacker).AttackTimer = GetTickCount
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal MapNPCNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim MapNum As Long
    Dim NPCNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(attacker)
    NPCNum = MapNpc(MapNum).Npc(MapNPCNum).Num
    Name = Trim$(Npc(NPCNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If Damage >= MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(MapNum).Npc(MapNPCNum).X * 32), (MapNpc(MapNum).Npc(MapNPCNum).Y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y)
            End If
        End If

        ' Calculate exp to give attacker
        exp = Npc(NPCNum).exp

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, exp, attacker
        Else
            ' no party - keep exp for self
            GivePlayerEXP attacker, exp
        End If
        
        'Drop the goods if they get it
        n = Int(Rnd * Npc(NPCNum).DropChance) + 1

        If n = 1 Then
            Call SpawnItem(Npc(NPCNum).DropItem, Npc(NPCNum).DropItemValue, MapNum, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y)
        End If
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(MapNPCNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).Npc(MapNPCNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong MapNPCNum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = MapNum Then
                    If TempPlayer(i).TargetType = TargetTypeNPC Then
                        If TempPlayer(i).Target = MapNPCNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TargetTypeNone
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
        
        ' OnPlayerKill
        If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnPlayerKill " & Trim$(attacker) & "," & Trim$(MapNPCNum) & "," & Trim$(TargetTypeNPC) & "," & Trim$(SpellNum) & "," & Trim$(Damage)
        
        ' OnNPCDeath
        If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnNPCDeath " & Trim$(MapNum) & "," & Trim$(MapNPCNum) & "," & Trim$(attacker) & "," & Trim$(TargetTypePlayer)
    
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(MapNPCNum).Num = 0
        MapNpc(MapNum).Npc(MapNPCNum).SpawnWait = GetTickCount
        MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) = 0
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) = MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).Npc(MapNPCNum).X * 32), (MapNpc(MapNum).Npc(MapNPCNum).Y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TargetTypeNPC, MapNPCNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = 1 ' player
        MapNpc(MapNum).Npc(MapNPCNum).Target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Behaviour = NPCTypeProtectAllies Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(MapNPCNum).Num Then
                    MapNpc(MapNum).Npc(i).Target = attacker
                    MapNpc(MapNum).Npc(i).TargetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).Npc(MapNPCNum).stopRegen = True
        MapNpc(MapNum).Npc(MapNPCNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC MapNPCNum, MapNum, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Npc MapNum, MapNPCNum, SpellNum, attacker
            End If
        End If
        
        SendMapNpcVitals MapNum, MapNPCNum
        
        ' OnPlayerHit
        If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnPlayerHit " & Trim$(attacker) & "," & Trim$(MapNPCNum) & "," & Trim$(TargetTypeNPC) & "," & Trim$(SpellNum) & "," & Trim$(Damage)
    
        ' OnNPCHurt
        If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnNPCHurt " & Trim$(MapNum) & "," & Trim$(MapNPCNum) & "," & Trim$(attacker) & "," & Trim$(TargetTypePlayer)
    End If

    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = GetTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal Index As Long)
Dim MapNum As Long, NPCNum As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNPCNum, Index) Then
        MapNum = GetPlayerMap(Index)
        NPCNum = MapNpc(MapNum).Npc(MapNPCNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(Index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(Index, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(NPCNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(MapNPCNum).X * 32), (MapNpc(MapNum).Npc(MapNPCNum).Y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(MapNPCNum, Index, Damage)
        Else
            SendActionMsg MapNum, "Block!", Cyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long
Dim NPCNum As Long
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).Npc(MapNPCNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Index)
    NPCNum = MapNpc(MapNum).Npc(MapNPCNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum).Npc(MapNPCNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If
    
    MapNpc(MapNum).Npc(MapNPCNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NPCNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum).Npc(MapNPCNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNPCNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum).Npc(MapNPCNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNPCNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNPCNum).Y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum).Npc(MapNPCNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNPCNum).Y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum).Npc(MapNPCNum).X) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If CanNpcAttackPlayer = True Then
        ' Send this packet so they can see the npc attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong ServerPackets.SNpcAttack
        Buffer.WriteLong MapNPCNum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
End Function

Sub NpcAttackPlayer(ByVal MapNPCNum As Long, ByVal victim As Long, ByVal Damage As Long)
Dim Name As String
Dim exp As Long
Dim MapNum As Long
Dim i As Long

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).Npc(MapNPCNum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(victim)
    Name = Trim$(Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Name)
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(MapNum).Npc(MapNPCNum).stopRegen = True
    MapNpc(MapNum).Npc(MapNPCNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNPCNum).Num
        
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & Name, BrightRed)
        
        ' OnNPCHit
        If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnNPCKill " & Trim$(MapNum) & "," & Trim$(MapNPCNum) & "," & Trim$(victim) & "," & Trim$(TargetTypePlayer)
        
        ' Set NPC target to 0
        MapNpc(MapNum).Npc(MapNPCNum).Target = 0
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        Call SendAnimation(MapNum, Npc(MapNpc(GetPlayerMap(victim)).Npc(MapNPCNum).Num).Animation, 0, 0, TargetTypePlayer, victim)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNPCNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' OnPlayerHurt
        If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnPlayerHurt " & Trim$(victim) & "," & Trim$(MapNPCNum) & "," & Trim$(TargetTypeNPC) & "," & Trim$(Damage)
        
        ' OnNPCHit
        If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnNPCHit " & Trim$(MapNum) & "," & Trim$(MapNPCNum) & "," & Trim$(victim) & "," & Trim$(TargetTypePlayer)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
    End If

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        MapNum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(victim, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case North
    
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case South
    
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case West
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case East
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MapMoralNone Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > RankModerator Then
        Call PlayerMsg(attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > RankModerator Then
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - exp)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(i).Target = TargetTypePlayer Then
                        If TempPlayer(i).Target = victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TargetTypeNone
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(victim) = NO Then
            If GetPlayerPK(attacker) = NO Then
                Call SetPlayerPK(attacker, YES)
                Call SendPlayerData(attacker)
                Call GlobalMsg(GetPlayerName(attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If

        Call OnDeath(victim)
        
        ' OnPlayerKill
        MyScript.ExecuteStatement "main.eds", "OnPlayerKill " & Trim$(attacker) & "," & Trim$(victim) & "," & Trim$(TargetTypePlayer) & "," & Trim$(SpellNum) & "," & Trim$(Damage)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        ' Run our lovely little script.
        If Options.Scripting = 1 Then
            ' OnPlayerHit
            MyScript.ExecuteStatement "main.eds", "OnPlayerHit " & Trim$(attacker) & "," & Trim$(victim) & "," & Trim$(TargetTypePlayer) & "," & Trim$(SpellNum) & "," & Trim$(Damage)
            
            ' OnPlayerHurt
            MyScript.ExecuteStatement "main.eds", "OnPlayerHurt " & Trim$(victim) & "," & Trim$(attacker) & "," & Trim$(TargetTypePlayer) & "," & Trim$(Damage)
        End If
        
        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer victim, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Player victim, SpellNum, attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal Index As Long, ByVal spellslot As Long)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim TargetType As Byte
    Dim Target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    SpellNum = GetPlayerSpell(Index, spellslot)
    MapNum = GetPlayerMap(Index)
    
    If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    TargetType = TempPlayer(Index).TargetType
    Target = TempPlayer(Index).Target
    Range = Spell(SpellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                PlayerMsg Index, "You do not have a target.", BrightRed
            End If
            If TargetType = TargetTypePlayer Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SpellTypeDamageHP And Spell(SpellNum).Type <> SpellTypeDamageMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TargetTypeNPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(MapNum).Npc(Target).X, MapNpc(MapNum).Npc(Target).Y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SpellTypeDamageHP And Spell(SpellNum).Type <> SpellTypeDamageMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TargetTypePlayer, Index
        SendActionMsg MapNum, "Casting " & Trim$(Spell(SpellNum).Name) & "!", BrightRed, ActionMsgScroll, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        TempPlayer(Index).spellBuffer.Spell = spellslot
        TempPlayer(Index).spellBuffer.Timer = GetTickCount
        TempPlayer(Index).spellBuffer.Target = TempPlayer(Index).Target
        TempPlayer(Index).spellBuffer.tType = TempPlayer(Index).TargetType
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal TargetType As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim X As Long, Y As Long
    
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    SpellNum = GetPlayerSpell(Index, spellslot)
    MapNum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    ' set the vital
    Vital = Spell(SpellNum).Vital
    AoE = Spell(SpellNum).AoE
    Range = Spell(SpellNum).Range
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(SpellNum).Type
                Case SpellTypeHealHP
                    SpellPlayer_Effect Vitals.HP, True, Index, Vital, SpellNum
                    DidCast = True
                Case SpellTypeHealMP
                    SpellPlayer_Effect Vitals.MP, True, Index, Vital, SpellNum
                    DidCast = True
                Case SpellTypeWarp
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TargetTypePlayer, Index
                    PlayerWarp Index, Spell(SpellNum).Map, Spell(SpellNum).X, Spell(SpellNum).Y
                    SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, 0, 0, TargetTypePlayer, Index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                X = GetPlayerX(Index)
                Y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                If TargetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
                
                If TargetType = TargetTypePlayer Then
                    X = GetPlayerX(Target)
                    Y = GetPlayerY(Target)
                Else
                    X = MapNpc(MapNum).Npc(Target).X
                    Y = MapNpc(MapNum).Npc(Target).Y
                End If
                
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    SendClearSpellBuffer Index
                End If
            End If
            Select Case Spell(SpellNum).Type
                Case SpellTypeDamageHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> Index Then
                                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(Index, i, True) Then
                                            SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TargetTypePlayer, i
                                            PlayerAttackPlayer Index, i, Vital, SpellNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, MapNpc(MapNum).Npc(i).X, MapNpc(MapNum).Npc(i).Y) Then
                                    If CanPlayerAttackNpc(Index, i, True) Then
                                        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TargetTypeNPC, i
                                        PlayerAttackNpc Index, i, Vital, SpellNum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SpellTypeHealHP, SpellTypeHealMP, SpellTypeDamageMP
                    If Spell(SpellNum).Type = SpellTypeHealHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SpellTypeHealMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SpellTypeDamageMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                    
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, Vital, SpellNum
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, MapNpc(MapNum).Npc(i).X, MapNpc(MapNum).Npc(i).Y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, SpellNum, MapNum
                                End If
                            End If
                        End If
                    Next
                Case SpellTypeScripted
                    If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnUseSpell " & Trim$(STR$(Index)) & "," & Trim$(STR$(SpellNum)) & "," & Trim$(STR$(Target)) & "," & Trim$(STR$(TargetType))
                    DidCast = True
            End Select
        Case 2 ' targetted
            If TargetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
            
            If TargetType = TargetTypePlayer Then
                X = GetPlayerX(Target)
                Y = GetPlayerY(Target)
            Else
                X = MapNpc(MapNum).Npc(Target).X
                Y = MapNpc(MapNum).Npc(Target).Y
            End If
                
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
            
            Select Case Spell(SpellNum).Type
                Case SpellTypeDamageHP
                    If TargetType = TargetTypePlayer Then
                        If CanPlayerAttackPlayer(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TargetTypePlayer, Target
                                PlayerAttackPlayer Index, Target, Vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TargetTypeNPC, Target
                                PlayerAttackNpc Index, Target, Vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    End If
                    
                Case SpellTypeDamageMP, SpellTypeHealMP, SpellTypeHealHP
                    If Spell(SpellNum).Type = SpellTypeDamageMP Then
                        VitalType = Vitals.MP
                        increment = False
                    ElseIf Spell(SpellNum).Type = SpellTypeHealMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SpellTypeHealHP Then
                        VitalType = Vitals.HP
                        increment = True
                    End If
                    
                    If TargetType = TargetTypePlayer Then
                        If Spell(SpellNum).Type = SpellTypeDamageMP Then
                            If CanPlayerAttackPlayer(Index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, SpellNum
                                DidCast = True
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, SpellNum
                            DidCast = True
                        End If
                    Else
                        If Spell(SpellNum).Type = SpellTypeDamageMP Then
                            If CanPlayerAttackNpc(Index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, SpellNum, MapNum
                                DidCast = True
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, SpellNum, MapNum
                            DidCast = True
                        End If
                    End If
                Case SpellTypeScripted
                    If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnUseSpell " & Trim$(STR$(Index)) & "," & Trim$(STR$(SpellNum)) & "," & Trim$(STR$(Target)) & "," & Trim$(STR$(TargetType))
                    DidCast = True
            End Select
    End Select
    
    If DidCast Then
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPCost)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
        
        TempPlayer(Index).SpellCD(spellslot) = GetTickCount + (Spell(SpellNum).CDTime * 1000)
        Call SendCooldown(Index, spellslot)
        SendActionMsg MapNum, Trim$(Spell(SpellNum).Name) & "!", BrightRed, ActionMsgScroll, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, 0, 0, TargetTypePlayer, Index
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
        
        If increment Then
            ' If it's a HoT we don't want to give them an initial heal, we just want to give them the HoT.
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Player Index, SpellNum
            Else
                SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ActionMsgScroll, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
                DoEvents
                SendVital Index, HP
            End If
        ElseIf Not increment Then
            SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ActionMsgScroll, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
            DoEvents
            SendVital Index, HP
        End If
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal MapNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TargetTypeNPC, Index
        SendActionMsg MapNum, sSymbol & Damage, Colour, ActionMsgScroll, MapNpc(MapNum).Npc(Index).X * 32, MapNpc(MapNum).Npc(Index).Y * 32
        
        ' send the sound
        SendMapSound Index, MapNpc(MapNum).Npc(Index).X, MapNpc(MapNum).Npc(Index).Y, SoundEntity.seSpell, SpellNum
        
        If increment Then
            MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Npc MapNum, Index, SpellNum
            End If
        ElseIf Not increment Then
            MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) - Damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal Index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(Index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(Index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)
    With TempPlayer(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, Index, True) Then
                    PlayerAttackPlayer .Caster, Index, Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)
    With TempPlayer(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg Player(Index).Map, "+" & Spell(.Spell).Vital, BrightGreen, ActionMsgScroll, Player(Index).X * 32, Player(Index).Y * 32
                Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + Spell(.Spell).Vital
                SendVital Index, HP
                .Timer = GetTickCount
                ' check if HoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime > (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal dotNum As Long)
    With MapNpc(MapNum).Npc(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, Index, True) Then
                    PlayerAttackNpc .Caster, Index, Spell(.Spell).Vital, , True
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal hotNum As Long)
    With MapNpc(MapNum).Npc(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightGreen, ActionMsgScroll, MapNpc(MapNum).Npc(Index).X * 32, MapNpc(MapNum).Npc(Index).Y * 32
                MapNpc(MapNum).Npc(Index).Vital(Vitals.HP) = MapNpc(MapNum).Npc(Index).Vital(Vitals.HP) + Spell(.Spell).Vital
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(Index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal MapNum As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(MapNum).Npc(Index).StunDuration = Spell(SpellNum).StunDuration
        MapNpc(MapNum).Npc(Index).StunTimer = GetTickCount
    End If
End Sub


