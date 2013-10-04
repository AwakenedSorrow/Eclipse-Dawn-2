Attribute VB_Name = "modServerClientTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next

End Sub

Function IsConnected(ByVal Index As Long) As Boolean

    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean

    If IsConnected(Index) Then
        If LenB(Trim$(Player(Index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsBanned(ByVal IP As String) As Boolean
    Dim filename As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long
    filename = App.Path & "\data\banlist.txt"

    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(Index) Then
        Set buffer = New clsBuffer
        
        buffer.PreAllocate 4 + (UBound(Data) - LBound(Data)) + 1
        buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        buffer.WriteBytes Data()
              
        frmServer.Socket(Index).SendData buffer.ToArray()
        
        '  Experimental
        DoEvents
        
        Set buffer = Nothing
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> Index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next

End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(i) > 0 Then
            Call SendDataTo(Party(partyNum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SGlobalMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToAll buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SAdminMsg
    buffer.WriteString Msg
    buffer.WriteLong Color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, buffer.ToArray
        End If
    Next
    
    Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataTo Index, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SMapMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToMap MapNum, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SAlertMsg
    buffer.WriteString Msg
    SendDataTo Index, buffer.ToArray
    DoEvents
    Call CloseSocket(Index)
    
    Set buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
    ' send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
                PlayerMsg Party(partyNum).Member(i), Msg, Color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)

    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(Index, "You have lost your connection with " & Options.Game_Name & ".")
    End If

End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (Index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal Index As Long)
Dim i As Long

    If Index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(Index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".")
        Else
            Call AlertMsg(Index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
    End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    If GetPlayerAccess(Index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then
            If GetTickCount < TempPlayer(Index).DataTimer Then
                Exit Sub
            End If
        End If

        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 25 Then
            If GetTickCount < TempPlayer(Index).DataTimer Then
                Exit Sub
            End If
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = GetTickCount + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(Index).GetData buffer(), vbUnicode, DataLength
    TempPlayer(Index).buffer.WriteBytes buffer()
    
    If TempPlayer(Index).buffer.Length >= 4 Then
        pLength = TempPlayer(Index).buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).buffer.Length - 4
        If pLength <= TempPlayer(Index).buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).buffer.ReadLong
            HandleData Index, TempPlayer(Index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(Index).buffer.Length >= 4 Then
            pLength = TempPlayer(Index).buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(Index).buffer.Trim
End Sub

Sub CloseSocket(ByVal Index As Long)

    If Index > 0 Then
        Call LeftGame(Index)
        Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        frmServer.Socket(Index).Close
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If

End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong MapNum
    buffer.WriteString Trim$(Map(MapNum).Name)
    buffer.WriteString Trim$(Map(MapNum).Music)
    buffer.WriteLong Map(MapNum).Revision
    buffer.WriteByte Map(MapNum).Moral
    buffer.WriteLong Map(MapNum).Up
    buffer.WriteLong Map(MapNum).Down
    buffer.WriteLong Map(MapNum).Left
    buffer.WriteLong Map(MapNum).Right
    buffer.WriteLong Map(MapNum).BootMap
    buffer.WriteByte Map(MapNum).BootX
    buffer.WriteByte Map(MapNum).BootY
    buffer.WriteByte Map(MapNum).MaxX
    buffer.WriteByte Map(MapNum).MaxY

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            With Map(MapNum).Tile(X, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).X
                    buffer.WriteLong .Layer(i).Y
                    buffer.WriteLong .Layer(i).Tileset
                Next
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        buffer.WriteLong Map(MapNum).Npc(X)
    Next

    MapCache(MapNum).Data = buffer.ToArray()
    
    Set buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(Index, s, WhoColor)
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
    Dim buffer As clsBuffer, i As Long

    If Index > MAX_PLAYERS Then Exit Function
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerData
    buffer.WriteLong Index
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerLevel(Index)
    buffer.WriteLong GetPlayerPoints(Index)
    buffer.WriteLong GetPlayerSprite(Index)
    buffer.WriteLong GetPlayerMap(Index)
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(Index, i)
    Next
    
    PlayerData = buffer.ToArray()
    Set buffer = Nothing
End Function

Sub SendJoinMap(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    SendDataTo Index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
    
    Set buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeft
    buffer.WriteLong Index
    SendDataToMapBut Index, MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    buffer.WriteLong SMapData
    buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(MapNum, i).playerName
        buffer.WriteLong MapItem(MapNum, i).Num
        buffer.WriteLong MapItem(MapNum, i).Value
        buffer.WriteLong MapItem(MapNum, i).X
        buffer.WriteLong MapItem(MapNum, i).Y
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(MapNum, i).playerName
        buffer.WriteLong MapItem(MapNum, i).Num
        buffer.WriteLong MapItem(MapNum, i).Value
        buffer.WriteLong MapItem(MapNum, i).X
        buffer.WriteLong MapItem(MapNum, i).Y
    Next

    SendDataToMap MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal MapNum As Long, ByVal MapNPCNum As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcVitals
    buffer.WriteLong MapNPCNum
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Vital(i)
    Next

    SendDataToMap MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        buffer.WriteLong MapNpc(MapNum).Npc(i).X
        buffer.WriteLong MapNpc(MapNum).Npc(i).Y
        buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        buffer.WriteLong MapNpc(MapNum).Npc(i).X
        buffer.WriteLong MapNpc(MapNum).Npc(i).Y
        buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
    Next

    SendDataToMap MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If

    Next

End Sub

Sub SendAnimations(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(Index, i)
        End If

    Next

End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If

    Next

End Sub

Sub SendResources(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(Index, i)
        End If

    Next

End Sub

Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(Index, i)
        buffer.WriteLong GetPlayerInvItemValue(Index, i)
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInvUpdate
    buffer.WriteLong InvSlot
    buffer.WriteLong GetPlayerInvItemNum(Index, InvSlot)
    buffer.WriteLong GetPlayerInvItemValue(Index, InvSlot)
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerWornEq
    buffer.WriteLong GetPlayerEquipment(Index, Armor)
    buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    buffer.WriteLong GetPlayerEquipment(Index, Helmet)
    buffer.WriteLong GetPlayerEquipment(Index, Shield)
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerEquipment(Index, Armor)
    buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    buffer.WriteLong GetPlayerEquipment(Index, Helmet)
    buffer.WriteLong GetPlayerEquipment(Index, Shield)
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong PlayerNum
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Helmet)
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHp
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMp
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendEXP(ByVal Index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerEXP
    buffer.WriteLong GetPlayerExp(Index)
    buffer.WriteLong GetPlayerNextLevel(Index)
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendStats(ByVal Index As Long)
Dim i As Long
Dim packet As String
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(Index, i)
    Next
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendWelcome(ByVal Index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(Index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SClassesData
    buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        buffer.WriteString GetClassName(i)
        buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        
        ' send array size
        buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        
        ' send array size
        buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SNewCharClasses
    buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        buffer.WriteString GetClassName(i)
        buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerData
    buffer.WriteLong Index
    buffer.WriteString vbNullString
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    SendDataToAllBut Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXY
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXYMap
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    buffer.WriteLong SUpdateItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    buffer.WriteLong SUpdateItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal NPCNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set buffer = New clsBuffer
    NPCSize = LenB(Npc(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NPCNum)), NPCSize
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong NPCNum
    buffer.WriteBytes NPCData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NPCNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set buffer = New clsBuffer
    NPCSize = LenB(Npc(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NPCNum)), NPCSize
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong NPCNum
    buffer.WriteBytes NPCData
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If

    Next

End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong GetPlayerSpell(Index, i)
    Next

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal Index As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(GetPlayerMap(Index)).Resource_Count

    If ResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceState
            buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).X
            buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).Y
        Next

    End If

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(MapNum).Resource_Count

    If ResourceCache(MapNum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(MapNum).Resource_Count
            buffer.WriteByte ResourceCache(MapNum).ResourceData(i).ResourceState
            buffer.WriteLong ResourceCache(MapNum).ResourceData(i).X
            buffer.WriteLong ResourceCache(MapNum).ResourceData(i).Y
        Next

    End If

    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SDoorAnimation
    buffer.WriteLong X
    buffer.WriteLong Y
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendActionMsg(ByVal MapNum As Long, ByVal Message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal X As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SActionMsg
    buffer.WriteString Message
    buffer.WriteLong Color
    buffer.WriteLong MsgType
    buffer.WriteLong X
    buffer.WriteLong Y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, buffer.ToArray()
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendBlood(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBlood
    buffer.WriteLong X
    buffer.WriteLong Y
    
    SendDataToMap MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendAnimation(ByVal MapNum As Long, ByVal Anim As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAnimation
    buffer.WriteLong Anim
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteByte LockType
    buffer.WriteLong LockIndex
    
    SendDataToMap MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendCooldown(ByVal Index As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCooldown
    buffer.WriteLong Slot
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClearSpellBuffer
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Long, ByVal Index As Long, ByVal Message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteString Message
    buffer.WriteString "[Map] "
    buffer.WriteLong saycolour
    
    SendDataToMap MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal Message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteString Message
    buffer.WriteString "[Global] "
    buffer.WriteLong saycolour
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub ResetShopAction(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SResetShopAction
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendStunned(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStunned
    buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendBank(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        buffer.WriteLong Bank(Index).Item(i).Num
        buffer.WriteLong Bank(Index).Item(i).Value
    Next
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapKey(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteByte Value
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteByte Value
    SendDataToMap MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendOpenShop(ByVal Index As Long, ByVal shopNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    buffer.WriteLong shopNum
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut Index, GetPlayerMap(Index), buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendTrade(ByVal Index As Long, ByVal tradeTarget As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    buffer.WriteLong tradeTarget
    buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal dataType As Byte)
Dim buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
    tradeTarget = TempPlayer(Index).InTrade
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdate
    buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            buffer.WriteLong TempPlayer(Index).TradeOffer(i).Num
            buffer.WriteLong TempPlayer(Index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(Index).TradeOffer(i).Num).Type = ItemTypeCurrency Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).price * TempPlayer(Index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Type = ItemTypeCurrency Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    buffer.WriteLong totalWorth
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeStatus
    buffer.WriteByte Status
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTarget(ByVal Index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STarget
    buffer.WriteLong TempPlayer(Index).Target
    buffer.WriteLong TempPlayer(Index).TargetType
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHotbar(ByVal Index As Long)
Dim i As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        buffer.WriteLong Player(Index).Hotbar(i).Slot
        buffer.WriteByte Player(Index).Hotbar(i).sType
    Next
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendLoginOk(ByVal Index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SLoginOk
    buffer.WriteLong Index
    buffer.WriteLong Player_HighIndex
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInGame(ByVal Index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHighIndex()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    buffer.WriteLong Player_HighIndex
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal Index As Long, ByVal targetPlayer As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyInvite
    buffer.WriteString Trim$(Player(targetPlayer).Name)
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    buffer.WriteByte 1
    buffer.WriteLong Party(partyNum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        buffer.WriteLong Party(partyNum).Member(i)
    Next
    buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal Index As Long)
Dim buffer As clsBuffer, i As Long, partyNum As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(Index).inParty
    If partyNum > 0 Then
        ' send party data
        buffer.WriteByte 1
        buffer.WriteLong Party(partyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            buffer.WriteLong Party(partyNum).Member(i)
        Next
        buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        buffer.WriteByte 0
    End If
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal Index As Long)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyVitals
    buffer.WriteLong Index
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong GetPlayerMaxVital(Index, i)
        buffer.WriteLong Player(Index).Vital(i)
    Next
    SendDataToParty partyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal MapNum As Long, ByVal Index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSpawnItem
    buffer.WriteLong Index
    buffer.WriteString MapItem(MapNum, Index).playerName
    buffer.WriteLong MapItem(MapNum, Index).Num
    buffer.WriteLong MapItem(MapNum, Index).Value
    buffer.WriteLong MapItem(MapNum, Index).X
    buffer.WriteLong MapItem(MapNum, Index).Y
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

