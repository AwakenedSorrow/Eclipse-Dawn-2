Attribute VB_Name = "modHandleEditorData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitEditorMessages()
    HandleEditorDataSub(CE_LoginUser) = GetAddress(AddressOf HandleEditorLogin)
    HandleEditorDataSub(CE_VersionCheck) = GetAddress(AddressOf HandleEditorVersionCheck)
    HandleEditorDataSub(CE_SaveDeveloper) = GetAddress(AddressOf HandleEditorSaveDeveloper)
    HandleEditorDataSub(CE_RequestMap) = GetAddress(AddressOf HandleEditorRequestMap)
    HandleEditorDataSub(CE_SaveMap) = GetAddress(AddressOf HandleEditorSaveMap)
End Sub

Public Sub HandleEditorData(ByVal Index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long
        
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CE_MSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleEditorDataSub(MsgType), Index, buffer.ReadBytes(buffer.Length), 0, 0
End Sub

Private Sub HandleEditorLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, Username As String, Password As String, i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Username = buffer.ReadString()
    
    ' Already logged in?
    For i = 1 To MAX_EDITORS
        If Trim$(Editor(i).Username) = Trim$(Username) Then
            SendEditorAlertMsg Index, "Multiple logins on the same account are not permitted.", True
            Set buffer = Nothing
            Exit Sub
        End If
    Next
    
    If FileExist("data\developers\" & Trim$(Username) & ".bin") Then
        Password = buffer.ReadString
        
        LoadEditor Index, Trim$(Username)
        
        If Trim$(Editor(Index).Password) = Trim$(Password) Then
            SendEditorLoginOK Index
            SendEditorMaxAmounts Index
            SendMapEditorNames Index
            SendEditorResources Index
            SendEditorAnimations Index
            SendEditorSpells Index
            SendEditorShops Index
        Else
            ClearEditor (Index)
            SendEditorAlertMsg Index, "This password is incorrect, you are not authorized to access this editor.", True
        End If
    Else
        SendEditorAlertMsg Index, "This account does not exist, you are not authorized to access this editor.", True
    End If
    
    Set buffer = Nothing
End Sub

Private Sub HandleEditorSaveDeveloper(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, FilePath As String, Name As String

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Editor(0).Username = buffer.ReadString()
    Editor(0).Password = buffer.ReadString()
    
    For i = 1 To Editor_MaxRights - 1
        Editor(0).HasRight(i) = buffer.ReadByte()
    Next
    
    SaveEditor 0
    
    Set buffer = Nothing
    
End Sub

Private Sub HandleEditorVersionCheck(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TempVer As String
    
    '  Make sure our editor is actually connected.
    If IsEditorConnected(Index) Then
        '  Write the data to the local buffer so we can extract the required data.
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        
        '  Retrieve the version of the editor client.
        TempVer = buffer.ReadString()
        
        ' Does it match? If not, disconnect them with a warning.
        If Trim$(TempVer) <> EDITOR_VERSION Then
            SendEditorAlertMsg Index, "Your editor appears to be outdated, please request the latest version from your server administrator!", True
        Else
            SendEditorVersionOK Index
        End If
    End If
End Sub

Private Sub HandleEditorRequestMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, MapNum As Long, i As Long

    If Editor(Index).HasRight(CanEditMap) = 0 Then
        SendEditorAlertMsg Index, "Insufficient permissions, you are now allowed to edit any maps.", False
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MapNum = buffer.ReadLong()
    
    ' Is any other dev editing said map?
    For i = 1 To MAX_EDITORS
        If TempEditor(i).InEditor = EditorMap And i <> Index Then
            If TempEditor(i).OnIndex = MapNum Then
                SendEditorAlertMsg Index, "This map is locked, someone else is currently editing it.", False
                Set buffer = Nothing
                Exit Sub
            End If
        End If
    Next
    
    ' Send the map data to the editor.
    SendEditorMap Index, MapNum
    
    ' Set editor nonsense.
    TempEditor(Index).InEditor = EditorMap
    TempEditor(Index).OnIndex = MapNum
    
    Set buffer = Nothing
End Sub

Private Sub HandleEditorSaveMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim MapNum As Long
    Dim X As Long
    Dim Y As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Prevent hacking
     If Editor(Index).HasRight(CanEditMap) = 0 Then
        SendEditorAlertMsg Index, "Insufficient permissions, you are now allowed to edit any maps.", False
        Exit Sub
    End If

    MapNum = buffer.ReadLong
    i = Map(MapNum).Revision + 1
    Call ClearMap(MapNum)
    
    Map(MapNum).Name = buffer.ReadString
    Map(MapNum).Music = buffer.ReadString
    Map(MapNum).Revision = i
    Map(MapNum).Moral = buffer.ReadByte
    Map(MapNum).Up = buffer.ReadLong
    Map(MapNum).Down = buffer.ReadLong
    Map(MapNum).Left = buffer.ReadLong
    Map(MapNum).Right = buffer.ReadLong
    Map(MapNum).BootMap = buffer.ReadLong
    Map(MapNum).BootX = buffer.ReadByte
    Map(MapNum).BootY = buffer.ReadByte
    Map(MapNum).MaxX = buffer.ReadByte
    Map(MapNum).MaxY = buffer.ReadByte
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map(MapNum).Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map(MapNum).Tile(X, Y).Layer(i).Tileset = buffer.ReadLong
            Next
            Map(MapNum).Tile(X, Y).Type = buffer.ReadByte
            Map(MapNum).Tile(X, Y).Data1 = buffer.ReadLong
            Map(MapNum).Tile(X, Y).Data2 = buffer.ReadLong
            Map(MapNum).Tile(X, Y).Data3 = buffer.ReadLong
            Map(MapNum).Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(X) = buffer.ReadLong
        Call ClearMapNpc(X, MapNum)
    Next

    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, MapNum, MapItem(MapNum, i).X, MapItem(MapNum, i).Y)
        Call ClearMapItem(i, MapNum)
    Next

    ' Respawn
    Call SpawnMapItems(MapNum)
    ' Save the map
    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    Call ClearTempTile(MapNum)
    Call CacheResources(MapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i
    
    ' Make sure all the editors are made aware of the change.
    Call SendEditorMapSaved(MapNum)
    
    Set buffer = Nothing
    
End Sub
