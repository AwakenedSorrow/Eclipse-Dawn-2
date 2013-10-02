Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenEditorSlot() As Long
    Dim i As Long
    FindOpenEditorSlot = 0

    For i = 1 To MAX_EDITORS

        If Not IsEditorConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, MapNum, X, Y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
            MapItem(MapNum, i).playerName = playerName
            MapItem(MapNum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(MapNum, i).canDespawn = canDespawn
            MapItem(MapNum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(MapNum, i).Num = ItemNum
            MapItem(MapNum, i).Value = ItemVal
            MapItem(MapNum, i).X = X
            MapItem(MapNum, i).Y = Y
            ' send to map
            SendSpawnItemToMap MapNum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim X As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(X, Y).Type = TileTypeItem) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(X, Y).Data1).Type = ItemTypeCurrency And Map(MapNum).Tile(X, Y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, 1, MapNum, X, Y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, Map(MapNum).Tile(X, Y).Data2, MapNum, X, Y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal MapNPCNum As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Dim NPCNum As Long
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    NPCNum = Map(MapNum).Npc(MapNPCNum)

    If NPCNum > 0 Then
    
        MapNpc(MapNum).Npc(MapNPCNum).Num = NPCNum
        MapNpc(MapNum).Npc(MapNPCNum).Target = 0
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = 0 ' clear
        
        MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) = GetNPCMaxVital(NPCNum, Vitals.HP)
        MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.MP) = GetNPCMaxVital(NPCNum, Vitals.MP)
        
        MapNpc(MapNum).Npc(MapNPCNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For X = 0 To Map(MapNum).MaxX
            For Y = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(X, Y).Type = TileTypeNPCSpawn Then
                    If Map(MapNum).Tile(X, Y).Data1 = MapNPCNum Then
                        MapNpc(MapNum).Npc(MapNPCNum).X = X
                        MapNpc(MapNum).Npc(MapNPCNum).Y = Y
                        MapNpc(MapNum).Npc(MapNPCNum).Dir = Map(MapNum).Tile(X, Y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next Y
        Next X
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                X = Random(0, Map(MapNum).MaxX)
                Y = Random(0, Map(MapNum).MaxY)
    
                If X > Map(MapNum).MaxX Then X = Map(MapNum).MaxX
                If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, X, Y) Then
                    MapNpc(MapNum).Npc(MapNPCNum).X = X
                    MapNpc(MapNum).Npc(MapNPCNum).Y = Y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn 't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
        
            For X = 0 To Map(MapNum).MaxX
                For Y = 0 To Map(MapNum).MaxY

                    If NpcTileIsOpen(MapNum, X, Y) Then
                        MapNpc(MapNum).Npc(MapNPCNum).X = X
                        MapNpc(MapNum).Npc(MapNPCNum).Y = Y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Num
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
            
            ' And run our NPC Spawn Script!
            If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnNPCSPawn " & Trim$(MapNum) & "," & Trim$(MapNPCNum)
        End If
        
        SendMapNpcVitals MapNum, MapNPCNum
    End If

End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = X Then
                    If GetPlayerY(LoopI) = Y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(MapNum).Npc(LoopI).X = X Then
                If MapNpc(MapNum).Npc(LoopI).Y = Y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(MapNum).Tile(X, Y).Type <> TileTypeWalkable Then
        If Map(MapNum).Tile(X, Y).Type <> TileTypeNPCSpawn Then
            If Map(MapNum).Tile(X, Y).Type <> TileTypeItem Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim X As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Dir < North Or Dir > East Then
        Exit Function
    End If

    X = MapNpc(MapNum).Npc(MapNPCNum).X
    Y = MapNpc(MapNum).Npc(MapNPCNum).Y
    CanNpcMove = True

    Select Case Dir
        Case North

            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = Map(MapNum).Tile(X, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TileTypeWalkable And n <> TileTypeItem And n <> TileTypeNPCSpawn Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, North + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case South

            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(X, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TileTypeWalkable And n <> TileTypeItem And n <> TileTypeNPCSpawn Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, South + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case West

            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TileTypeWalkable And n <> TileTypeItem And n <> TileTypeNPCSpawn Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X - 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, West + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case East

            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(X + 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TileTypeWalkable And n <> TileTypeItem And n <> TileTypeNPCSpawn Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X + 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, East + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Dir < North Or Dir > East Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(MapNPCNum).Dir = Dir

    Select Case Dir
        Case North
            MapNpc(MapNum).Npc(MapNPCNum).Y = MapNpc(MapNum).Npc(MapNPCNum).Y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case South
            MapNpc(MapNum).Npc(MapNPCNum).Y = MapNpc(MapNum).Npc(MapNPCNum).Y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case West
            MapNpc(MapNum).Npc(MapNPCNum).X = MapNpc(MapNum).Npc(MapNPCNum).X - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case East
            MapNpc(MapNum).Npc(MapNPCNum).X = MapNpc(MapNum).Npc(MapNPCNum).X + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Dir < North Or Dir > East Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(MapNPCNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong MapNPCNum
    Buffer.WriteLong Dir
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal MapNum As Long)
    Dim Y As Long
    Dim X As Long
    TempTile(MapNum).DoorTimer = 0
    ReDim TempTile(MapNum).DoorOpen(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            TempTile(MapNum).DoorOpen(X, Y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal MapNum As Long)
    Dim X As Long, Y As Long, Resource_Count As Long
    Resource_Count = 0

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            If Map(MapNum).Tile(X, Y).Type = TileTypeResource Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).X = X
                ResourceCache(MapNum).ResourceData(Resource_Count).Y = Y
                ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).Tile(X, Y).Data1).health
            End If

        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(Index, oldSlot)
    OldValue = GetPlayerBankItemValue(Index, oldSlot)
    NewNum = GetPlayerBankItemNum(Index, newSlot)
    NewValue = GetPlayerBankItemValue(Index, newSlot)
    
    SetPlayerBankItemNum Index, newSlot, OldNum
    SetPlayerBankItemValue Index, newSlot, OldValue
    
    SetPlayerBankItemNum Index, oldSlot, NewNum
    SetPlayerBankItemValue Index, oldSlot, NewValue
        
    SendBank Index
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(Index, oldSlot)
    OldValue = GetPlayerInvItemValue(Index, oldSlot)
    NewNum = GetPlayerInvItemNum(Index, newSlot)
    NewValue = GetPlayerInvItemValue(Index, newSlot)
    SetPlayerInvItemNum Index, newSlot, OldNum
    SetPlayerInvItemValue Index, newSlot, OldValue
    SetPlayerInvItemNum Index, oldSlot, NewNum
    SetPlayerInvItemValue Index, oldSlot, NewValue
    SendInventory Index
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(Index, oldSlot)
    NewNum = GetPlayerSpell(Index, newSlot)
    SetPlayerSpell Index, oldSlot, NewNum
    SetPlayerSpell Index, newSlot, OldNum
    SendPlayerSpells Index
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(Index, GetPlayerEquipment(Index, EqSlot)) > 0 Then
        GiveInvItem Index, GetPlayerEquipment(Index, EqSlot), 0
        PlayerMsg Index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(Index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, EqSlot)
        ' remove equipment
        SetPlayerEquipment Index, 0, EqSlot
        SendWornEquipment Index
        SendMapEquipment Index
        SendStats Index
        ' send vitals
        Call SendVital(Index, Vitals.HP)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    Else
        PlayerMsg Index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
Dim partyNum As Long, i As Long

    partyNum = TempPlayer(Index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
            ' check if leader
            If Party(partyNum).Leader = Index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> Index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = Index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo Index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = Index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo Index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                Index = Party(partyNum).Member(i)
                ' player exist?
                If Index > 0 Then
                    ' remove them
                    TempPlayer(Index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo Index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg Index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg Index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(Index).inParty > 0 Then
        partyNum = TempPlayer(Index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = Index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, Index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = Index
                    ' let them know
                    PlayerMsg Index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg Index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg Index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, Index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = Index
        ' let them know
        PlayerMsg Index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if already in a party
    If TempPlayer(Index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(Index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                'add to the party
                Party(partyNum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                SendPartyVitals partyNum, targetPlayer
                ' let everyone know they've joined
                PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg Index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = Index
        Party(partyNum).Member(1) = Index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, Index
        SendPartyVitals partyNum, targetPlayer
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(Index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(Index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal targetPlayer As Long)
    PlayerMsg Index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, X As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For X = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(X) = Party(partyNum).Member(X + 1)
                    Party(partyNum).Member(X + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal exp As Long, ByVal Index As Long)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

    ' check if it's worth sharing
    If Not exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP Index, exp
        Exit Sub
    End If
    
    ' find out the equal share
    expShare = exp \ Party(partyNum).MemberCount
    leftOver = exp Mod Party(partyNum).MemberCount
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(RAND(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal exp As Long)
    ' give the exp
    Call SetPlayerExp(Index, GetPlayerExp(Index) + exp)
    SendEXP Index
    SendActionMsg GetPlayerMap(Index), "+" & exp & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp Index
End Sub
