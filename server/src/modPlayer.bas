Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal Index As Long)
    If Not IsPlaying(Index) Then
        Call JoinGame(Index)
        Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal Index As Long)
    Dim i As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    
    ' send the login ok
    SendLoginOk Index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendAnimations(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendResources(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendPlayerSpells(Index)
    Call SendHotbar(Index)
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
    Next
    SendEXP Index
    Call SendStats(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) <= RankModerator Then
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Send welcome messages
    Call SendWelcome(Index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next
    
    ' Send the flag so they know they can start doing stuff
    SendInGame Index
End Sub

Sub LeftGame(ByVal Index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False
        
        ' Loop through entire map and purge Player from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    If TempPlayer(i).TargetType = TargetTypePlayer Then
                        If TempPlayer(i).Target = Index Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TargetTypeNone
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
        
        'Loop through the mapnpcs to remove the player from their targets
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(Index)).Npc(i).TargetType = TargetTypePlayer Then
                If MapNpc(GetPlayerMap(Index)).Npc(i).Target = Index Then
                    MapNpc(GetPlayerMap(Index)).Npc(i).Target = 0
                    MapNpc(GetPlayerMap(Index)).Npc(i).TargetType = TargetTypeNone
                End If
            End If
        Next
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(Index).InTrade > 0 Then
            tradeTarget = TempPlayer(Index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave Index

        ' save and clear data.
        Call SavePlayer(Index)
        Call SaveBank(Index)
        Call ClearBank(Index)

        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= RankModerator Then
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(Index)
End Sub

Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(Index, Armor)
    Helm = GetPlayerEquipment(Index, Helmet)
    GetPlayerProtection = (GetPlayerStat(Index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(Index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(Index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Endurance) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim shopNum As Long
    Dim OldMap As Long, OldX As Long, OldY As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If X > Map(MapNum).MaxX Then X = Map(MapNum).MaxX
    If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    
    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(Index) Then
        SendPlayerXYToMap Index
    End If
    
    ' clear target
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = TargetTypeNone
    SendTarget Index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If OldMap <> MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If
    
    OldX = GetPlayerX(Index)
    OldY = GetPlayerY(Index)
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)
    
    ' send player's equipment to new map
    SendMapEquipment Index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(MapNum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    SendMapEquipmentTo i, Index
                End If
            End If
        Next
    End If

    ' Run the script if needed.
     If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnPlayerWarp " & Trim$(STR$(Index)) & "," & Trim$(STR$(OldMap)) & "," & Trim$(STR$(OldX)) & "," & Trim$(STR$(OldY)) & "," & Trim$(STR$(MapNum)) & "," & Trim$(STR$(X)) & "," & Trim$(STR$(Y))
    
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If MapNpc(OldMap).Npc(i).Num > 0 Then
                MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNPCMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    TempPlayer(Index).GettingMap = YES
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong MapNum
    Buffer.WriteLong Map(MapNum).Revision
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, MapNum As Long
    Dim X As Long, Y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, amount As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < North Or Dir > East Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Moved = NO
    MapNum = GetPlayerMap(Index)
    
    Select Case Dir
        Case North

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, North + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TileTypeBlocked Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TileTypeResource Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TileTypeKey Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TileTypeKey And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(Index)).Up).MaxY
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TargetTypeNone
                    SendTarget Index
                End If
            End If

        Case South

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < Map(MapNum).MaxY Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, South + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TileTypeBlocked Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TileTypeResource Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TileTypeKey Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TileTypeKey And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Down > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TargetTypeNone
                    SendTarget Index
                End If
            End If

        Case West

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, West + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TileTypeBlocked Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TileTypeResource Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TileTypeKey Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TileTypeKey And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(Index)).Left).MaxX
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, NewMapX, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TargetTypeNone
                    SendTarget Index
                End If
            End If

        Case East

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < Map(MapNum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, East + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TileTypeBlocked Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TileTypeResource Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TileTypeKey Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TileTypeKey And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Right > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TargetTypeNone
                    SendTarget Index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TileTypeWarp Then
            MapNum = .Data1
            X = .Data2
            Y = .Data3
            Call PlayerWarp(Index, MapNum, X, Y)
            Moved = YES
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TileTypeDoor Then
            MapNum = .Data1
            X = .Data2
            Y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
            Call PlayerWarp(Index, MapNum, X, Y)
            Moved = YES
        End If
    
        ' Check for key trigger open
        If .Type = TileTypeKeyOpen Then
            X = .Data1
            Y = .Data2
    
            If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TileTypeKey And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                SendMapKey Index, X, Y, 1
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TileTypeShop Then
            X = .Data1
            If X > 0 Then ' shop exists?
                If Len(Trim$(Shop(X).Name)) > 0 Then ' name exists?
                    SendOpenShop Index, X
                    TempPlayer(Index).InShop = X ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TileTypeBank Then
            SendBank Index
            TempPlayer(Index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TileTypeHeal Then
            VitalType = .Data1
            amount = .Data2
            If Not GetPlayerVital(Index, VitalType) = GetPlayerMaxVital(Index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(Index), "+" & amount, Colour, ActionMsgScroll, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
                SetPlayerVital Index, VitalType, GetPlayerVital(Index, VitalType) + amount
                PlayerMsg Index, "You feel rejuvinating forces flowing through your boy.", BrightGreen
                Call SendVital(Index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TileTypeTrap Then
            amount = .Data1
            SendActionMsg GetPlayerMap(Index), "-" & amount, BrightRed, ActionMsgScroll, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
            If GetPlayerVital(Index, HP) - amount <= 0 Then
                KillPlayer Index
                PlayerMsg Index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital Index, HP, GetPlayerVital(Index, HP) - amount
                PlayerMsg Index, "You're injured by a trap.", BrightRed
                Call SendVital(Index, HP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            Moved = YES
        End If
        
        ' Slide
        If .Type = TileTypeSlide Then
             Select Case .Data1
                 Case North
                     If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TileTypeResource Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TileTypeBlocked Then Exit Sub
                Case West
                     If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TileTypeResource Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TileTypeBlocked Then Exit Sub
                Case South
                     If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TileTypeResource Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TileTypeBlocked Then Exit Sub
                Case East
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TileTypeResource Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TileTypeBlocked Then Exit Sub
            End Select
            ForcePlayerMove Index, MOVING_WALKING, .Data1
            Moved = YES
        End If
        
         ' Scripted
        If .Type = TileTypeScripted Then
            If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnUseTile " & Trim$(STR$(Index)) & "," & Trim$(STR$(GetPlayerMap(Index))) & "," & Trim$(STR$(.Data1))
        End If
    End With

    
    ' They tried to hack
    If Moved = NO Then
        PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
    End If

End Sub

Sub ForcePlayerMove(ByVal Index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < North Or Direction > East Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case North
            If GetPlayerY(Index) = 0 Then Exit Sub
        Case West
            If GetPlayerX(Index) = 0 Then Exit Sub
        Case South
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
        Case East
            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove Index, Direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal Index As Long)
    Dim Slot As Long
    Dim itemnum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(Index, i)

        If itemnum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(itemnum).Type <> ItemTypeWeapon Then SetPlayerEquipment Index, 0, i
                Case Equipment.Armor

                    If Item(itemnum).Type <> ItemTypeArmor Then SetPlayerEquipment Index, 0, i
                Case Equipment.Helmet

                    If Item(itemnum).Type <> ItemTypeHelmet Then SetPlayerEquipment Index, 0, i
                Case Equipment.Shield

                    If Item(itemnum).Type <> ItemTypeShield Then SetPlayerEquipment Index, 0, i
            End Select

        Else
            SetPlayerEquipment Index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(itemnum).Type = ItemTypeCurrency Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(Index, i) = itemnum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Function
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, i) = itemnum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal Index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = itemnum Then
            If Item(itemnum).Type = ItemTypeCurrency Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal Index As Long, ByVal itemnum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = itemnum Then
            If Item(itemnum).Type = ItemTypeCurrency Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim itemnum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvSlot <= 0 Or InvSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    itemnum = GetPlayerInvItemNum(Index, InvSlot)

    If Item(itemnum).Type = ItemTypeCurrency Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(Index, InvSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(Index, InvSlot, GetPlayerInvItemValue(Index, InvSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(Index, InvSlot, 0)
        Call SetPlayerInvItemValue(Index, InvSlot, 0)
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal Index As Long, ByVal itemnum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If
    
    If Item(itemnum).Type = ItemTypeCurrency Or ItemVal <= 1 Then
        ' A single item, or a currency item.
        i = FindOpenInvSlot(Index, itemnum)

        ' Check to see if inventory is full
        If i <> 0 Then
            Call SetPlayerInvItemNum(Index, i, itemnum)
            Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
            If sendUpdate Then Call SendInventoryUpdate(Index, i)
            GiveInvItem = True
        Else
            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
            GiveInvItem = False
        End If
    
    Else
        ' Multiple Items.
        If GetOpenInvSlots(Index) < ItemVal Then
            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
            GiveInvItem = False
        Else
            For i = 1 To ItemVal
                i = FindOpenInvSlot(Index, itemnum)
                
                Call SetPlayerInvItemNum(Index, i, itemnum)
                Call SetPlayerInvItemValue(Index, i, 0)
                If sendUpdate Then Call SendInventoryUpdate(Index, i)
            Next i
            
            GiveInvItem = True
        End If
        
    End If

End Function

Public Function GetOpenInvSlots(ByVal Index As Long)
Dim i As Long
    
    GetOpenInvSlots = 0
    For i = 1 To MAX_INV
        If Player(Index).Inv(i).Num < 1 Then
            GetOpenInvSlots = GetOpenInvSlots + 1
        End If
    Next

End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim i As Long
    Dim n As Long
    Dim MapNum As Long
    Dim Msg As String

    If Not IsPlaying(Index) Then Exit Sub
    MapNum = GetPlayerMap(Index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(Index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, i).X = GetPlayerX(Index)) Then
                    If (MapItem(MapNum, i).Y = GetPlayerY(Index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
    
                            If Item(GetPlayerInvItemNum(Index, n)).Type = ItemTypeCurrency Then
                                Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, i).Value)
                                Msg = MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(Index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                            End If
    
                            ' Erase item from the map
                            ClearMapItem i, MapNum
                            
                            Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0)
                            SendActionMsg GetPlayerMap(Index), Msg, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                            Exit For
                        Else
                            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal Index As Long, ByVal mapItemNum As Long)
Dim MapNum As Long

    MapNum = GetPlayerMap(Index)
    
    ' no lock or locked to player?
    If MapItem(MapNum, mapItemNum).playerName = vbNullString Or MapItem(MapNum, mapItemNum).playerName = Trim$(GetPlayerName(Index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal Index As Long, ByVal invNum As Long, ByVal amount As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Or TempPlayer(Index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(Index, invNum) > 0) Then
        If (GetPlayerInvItemNum(Index, invNum) <= MAX_ITEMS) Then
            i = FindOpenMapItemSlot(GetPlayerMap(Index))

            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, invNum)
                MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).Y = GetPlayerY(Index)
                MapItem(GetPlayerMap(Index), i).playerName = Trim$(GetPlayerName(Index))
                MapItem(GetPlayerMap(Index), i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(Index), i).canDespawn = True
                MapItem(GetPlayerMap(Index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME

                If Item(GetPlayerInvItemNum(Index, invNum)).Type = ItemTypeCurrency Then

                    ' Check if its more then they have and if so drop it all
                    If amount >= GetPlayerInvItemValue(Index, invNum) Then
                        MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, invNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, invNum, 0)
                        Call SetPlayerInvItemValue(Index, invNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), i).Value = amount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, invNum, GetPlayerInvItemValue(Index, invNum) - amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), i).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, invNum, 0)
                    Call SetPlayerInvItemValue(Index, invNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(Index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, amount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), Trim$(GetPlayerName(Index)), MapItem(GetPlayerMap(Index), i).canDespawn)
            Else
                Call PlayerMsg(Index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        expRollover = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
        
        ' can level up?
        If Not SetPlayerLevel(Index, GetPlayerLevel(Index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPoints(Index, GetPlayerPoints(Index) + 3)
        Call SetPlayerExp(Index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        ' Scripting or not?
        If Options.Scripting <> 1 Then
            If level_count = 1 Then
                'singular
                GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " level!", Brown
            Else
                'plural
                GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " levels!", Brown
            End If
        Else
            MyScript.ExecuteStatement "main.eds", "OnPlayerLevelUp " & Trim$(STR$(Index)) & "," & Trim$(STR$(GetPlayerLevel(Index))) & "," & Trim$(STR$(level_count))
        End If
        
        SendEXP Index
        SendPlayerData Index
        For i = 1 To Vitals.Vital_Count - 1
            SendVital Index, i
        Next
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
Dim F As Long
    
    If Index < 1 Or Index > Player_HighIndex Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub
    If Len(Trim$(Name)) < 3 Then Exit Sub
    
    ' Remove old name from file
    Call DeleteName(Trim$(Player(Index).Name))
    
    Player(Index).Name = Trim$(Name)
    
    ' Append name to file
    F = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Append As #F
    Print #F, Name
    Close #F
    Call SavePlayer(Index)
    Call SendPlayerData(Index)
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
End Function

Function SetPlayerLevel(ByVal Index As Long, ByVal Level As Long) As Boolean
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    Player(Index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(Index) + 1) ^ 3 - (6 * (GetPlayerLevel(Index) + 1) ^ 2) + 17 * (GetPlayerLevel(Index) + 1) - 12)
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal exp As Long)
    Player(Index).exp = exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Dim X As Long, i As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    X = Player(Index).Stat(Stat) + GetClassStat(Player(Index).Class, Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(Index).Equipment(i) > 0 Then
            If Item(Player(Index).Equipment(i)).Add_Stat(Stat) > 0 Then
                X = X + Item(Player(Index).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = X
End Function

Public Function GetPlayerRawStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(Index).Stat(Stat) + Class(Player(Index).Class).Stat(Stat)
End Function

Public Function GetPlayerRawStatNoClass(ByVal Index As Long, ByVal Stat As Stats) As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStatNoClass = Player(Index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).Stat(Stat) = Value
End Sub

Function GetPlayerPoints(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPoints = Player(Index).Points
End Function

Sub SetPlayerPoints(ByVal Index As Long, ByVal Points As Long)
    If Points <= 0 Then Points = 0
    Player(Index).Points = Points
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Map = MapNum
    End If

End Sub

Function GetPlayerX(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If InvSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal itemnum As Long)
    Player(Index).Inv(InvSlot).Num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(Index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal SpellNum As Long)
    Player(Index).Spell(spellslot) = SpellNum
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot) = invNum
End Sub

' ToDo
Sub OnDeath(ByVal Index As Long)
    Dim i As Long, n As Long
    Dim Item As Long
    Dim Slot As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(Index, Vitals.HP, 0)
    
    ' Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                If TempPlayer(i).TargetType = TargetTypePlayer Then
                    If TempPlayer(i).Target = Index Then
                        TempPlayer(i).Target = 0
                        TempPlayer(i).TargetType = TargetTypeNone
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next
    
    ' Is scripting enabled?
    If Options.Scripting = 1 Then
        MyScript.ExecuteStatement "main.eds", "OnPlayerDeath " & Trim$(STR$(Index))
    Else
        ' Drop all worn items
        For i = 1 To Equipment.Equipment_Count - 1
            If GetPlayerEquipment(Index, i) > 0 Then
                Item = GetPlayerEquipment(Index, i)
                
                Slot = 0
                For n = 1 To MAX_INV
                    If Player(Index).Inv(i).Num = Item Then Slot = i
                Next
                
                If Slot > 0 Then PlayerMapDropItem Index, Slot, 0
            End If
        Next

        ' Warp player away
        Call SetPlayerDir(Index, South)
    
        With Map(GetPlayerMap(Index))
            ' to the bootmap if it is set
            If .BootMap > 0 Then
                PlayerWarp Index, .BootMap, .BootX, .BootY
            Else
                Call PlayerWarp(Index, START_MAP, START_X, START_Y)
            End If
        End With
    End If
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(Index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(Index).spellBuffer.Spell = 0
    TempPlayer(Index).spellBuffer.Timer = 0
    TempPlayer(Index).spellBuffer.Target = 0
    TempPlayer(Index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(Index)
    
    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
    End If

End Sub

Sub CheckResource(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    
    ' Check attack timer
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        If GetTickCount < TempPlayer(Index).AttackTimer + Item(GetPlayerEquipment(Index, Weapon)).Speed Then Exit Sub
    Else
        If GetTickCount < TempPlayer(Index).AttackTimer + 1000 Then Exit Sub
    End If
    
    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TileTypeResource Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(Index)).Tile(X, Y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count

            If ResourceCache(GetPlayerMap(Index)).ResourceData(i).X = X Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(i).Y = Y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(Index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(Index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Then

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(Index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).X
                        rY = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).Y
                        
                        Damage = Item(GetPlayerEquipment(Index, Weapon)).Data2
                    
                        ' check if damage is more than health
                        If Damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                SendActionMsg GetPlayerMap(Index), "-" & ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                End If
                                ' carry on
                                GiveInvItem Index, Resource(Resource_index).ItemReward, 1
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage
                                SendActionMsg GetPlayerMap(Index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            End If
                            ' send the sound
                            SendMapSound Index, rX, rY, SoundEntity.seResource, Resource_index
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(Index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                        End If
                    End If
                    ' Reset attack timer
                    TempPlayer(Index).AttackTimer = GetTickCount
                Else
                    PlayerMsg Index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg Index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(Index).Item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal itemnum As Long)
    Bank(Index).Item(BankSlot).Num = itemnum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(Index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(Index).Item(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal InvSlot As Long, ByVal amount As Long)
Dim BankSlot

    If InvSlot < 0 Or InvSlot > MAX_INV Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerInvItemValue(Index, InvSlot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(Index, GetPlayerInvItemNum(Index, InvSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, InvSlot)).Type = ItemTypeCurrency Then
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, InvSlot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, InvSlot), amount)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, InvSlot))
                Call SetPlayerBankItemValue(Index, BankSlot, amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, InvSlot), amount)
            End If
        Else
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, InvSlot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, InvSlot), 0)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, InvSlot))
                Call SetPlayerBankItemValue(Index, BankSlot, 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, InvSlot), 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal BankSlot As Long, ByVal amount As Long)
Dim InvSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerBankItemValue(Index, BankSlot) Then
        Exit Sub
    End If
    
    InvSlot = FindOpenInvSlot(Index, GetPlayerBankItemNum(Index, BankSlot))
        
    If InvSlot > 0 Then
        If Item(GetPlayerBankItemNum(Index, BankSlot)).Type = ItemTypeCurrency Then
            Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), amount)
            Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - amount)
            If GetPlayerBankItemValue(Index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(Index, BankSlot) > 1 Then
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - 1)
            Else
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

End Sub

Public Sub KillPlayer(ByVal Index As Long)
Dim exp As Long

    ' Calculate exp to give attacker
    exp = GetPlayerExp(Index) \ 3

    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    If exp = 0 Then
        Call PlayerMsg(Index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(Index, GetPlayerExp(Index) - exp)
        SendEXP Index
        Call PlayerMsg(Index, "You lost " & exp & " exp.", BrightRed)
    End If
    
    Call OnDeath(Index)
End Sub

Public Sub UseItem(ByVal Index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, X As Long, Y As Long, itemnum As Long

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(Index, invNum) > 0) And (GetPlayerInvItemNum(Index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(Index, invNum)).Data2
        itemnum = GetPlayerInvItemNum(Index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(itemnum).Type
            Case ItemTypeArmor
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(itemnum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Armor)
                End If

                SetPlayerEquipment Index, itemnum, Armor
                PlayerMsg Index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem Index, itemnum, 0

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ItemTypeWeapon
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(itemnum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Weapon) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Weapon)
                End If

                SetPlayerEquipment Index, itemnum, Weapon
                PlayerMsg Index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem Index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ItemTypeHelmet
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(itemnum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Helmet)
                End If

                SetPlayerEquipment Index, itemnum, Helmet
                PlayerMsg Index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem Index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ItemTypeShield
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(itemnum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Shield)
                End If

                SetPlayerEquipment Index, itemnum, Shield
                PlayerMsg Index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem Index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            ' consumable
            Case ItemTypeConsume
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(itemnum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(itemnum).AddHP > 0 Then
                    Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + Item(itemnum).AddHP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(itemnum).AddHP, BrightGreen, ActionMsgScroll, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, HP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add mp
                If Item(itemnum).AddMP > 0 Then
                    Player(Index).Vital(Vitals.MP) = Player(Index).Vital(Vitals.MP) + Item(itemnum).AddMP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(itemnum).AddMP, BrightBlue, ActionMsgScroll, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, MP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add exp
                If Item(itemnum).AddEXP > 0 Then
                    SetPlayerExp Index, GetPlayerExp(Index) + Item(itemnum).AddEXP
                    CheckPlayerLevelUp Index
                    SendActionMsg GetPlayerMap(Index), "+" & Item(itemnum).AddEXP & " EXP", White, ActionMsgScroll, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendEXP Index
                End If
                Call SendAnimation(GetPlayerMap(Index), Item(itemnum).Animation, 0, 0, TargetTypePlayer, Index)
                Call TakeInvItem(Index, Player(Index).Inv(invNum).Num, 0)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ItemTypeKey
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(itemnum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(Index)
                    Case North

                        If GetPlayerY(Index) > 0 Then
                            X = GetPlayerX(Index)
                            Y = GetPlayerY(Index) - 1
                        Else
                            Exit Sub
                        End If

                    Case South

                        If GetPlayerY(Index) < Map(GetPlayerMap(Index)).MaxY Then
                            X = GetPlayerX(Index)
                            Y = GetPlayerY(Index) + 1
                        Else
                            Exit Sub
                        End If

                    Case West

                        If GetPlayerX(Index) > 0 Then
                            X = GetPlayerX(Index) - 1
                            Y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                    Case East

                        If GetPlayerX(Index) < Map(GetPlayerMap(Index)).MaxX Then
                            X = GetPlayerX(Index) + 1
                            Y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TileTypeKey Then

                    ' Check if the key they are using matches the map key
                    If itemnum = Map(GetPlayerMap(Index)).Tile(X, Y).Data1 Then
                        TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        SendMapKey Index, X, Y, 1
                        Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(Index), Item(itemnum).Animation, X, Y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(Index)).Tile(X, Y).Data2 = 1 Then
                            Call TakeInvItem(Index, itemnum, 0)
                            Call PlayerMsg(Index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ItemTypeSpell
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(itemnum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(itemnum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(Index) Then
                            i = FindOpenSpellSlot(Index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(Index, n) Then
                                    Call SetPlayerSpell(Index, i, n)
                                    Call SendAnimation(GetPlayerMap(Index), Item(itemnum).Animation, 0, 0, TargetTypePlayer, Index)
                                    Call TakeInvItem(Index, itemnum, 0)
                                    Call PlayerMsg(Index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                    Call SendPlayerSpells(Index)
                                Else
                                    Call PlayerMsg(Index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(Index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(Index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(Index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
                
            Case ItemTypeScripted
                If Options.Scripting = 1 Then MyScript.ExecuteStatement "main.eds", "OnUseItem " & Trim$(STR$(Index)) & "," & Trim$(STR$(itemnum)) & "," & Trim$(STR$(invNum))
        
        End Select
    End If
End Sub
