Attribute VB_Name = "modHandleData"
Option Explicit

Public Sub InitMessages()
    HandleDataSub(SE_AlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SE_VersionOK) = GetAddress(AddressOf HandleVersionOK)
    HandleDataSub(SE_LoginOK) = GetAddress(AddressOf HandleLoginOK)
    HandleDataSub(SE_MapNames) = GetAddress(AddressOf HandleMapNames)
    HandleDataSub(SE_MapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SE_ResourceData) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SE_MaxAmounts) = GetAddress(AddressOf Handlemaxamounts)
End Sub

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Sub HandleData(ByRef data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyEditor
        Exit Sub
    End If
    
    If MsgType >= SE_MSG_COUNT Then
        DestroyEditor
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, buffer.ReadBytes(buffer.length), 0, 0
    
End Sub

Private Sub HandleAlertMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String, Disc As Byte
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
        
    Msg = buffer.ReadString
    Disc = buffer.ReadByte
    
    Set buffer = Nothing
    MsgBox Msg, vbOKOnly
    
    If Disc = 1 Then DestroyEditor
    
End Sub

Private Sub HandleVersionOK(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' We've got our confirmation, time to stop checking versions.
    CheckingVersion = False
End Sub

Private Sub HandleMapNames(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, TempName As String, TempRev As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    frmEditor.lstMapList.Clear
    For i = 1 To MAX_MAPS
        TempName = buffer.ReadString()
        TempRev = buffer.ReadLong()
        frmEditor.lstMapList.AddItem CStr(i) & ": " & Trim$(TempName) & " | Rev." & Trim(CStr(TempRev))
    Next
    
    SetStatus "Received Map Names and applied them to the list."
    
    Set buffer = Nothing
    
    If Editor.HasRight(CanEditMap) = 1 Then
        SendRequestMap 1
    End If
End Sub

Private Sub HandleLoginOK(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long
    ' We've got our confirmation, time to stop the timeout loop.
    LoggingIn = False
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Editor.Username = buffer.ReadString
    
    For i = 1 To Editor_MaxRights - 1
        Editor.HasRight(i) = buffer.ReadByte()
    Next
    
    Set buffer = Nothing
    
End Sub

Private Sub HandleMapData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim buffer As clsBuffer
Dim MapNum As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()

    MapNum = buffer.ReadLong
    Map.name = buffer.ReadString
    Map.Music = buffer.ReadString
    Map.Revision = buffer.ReadLong
    Map.Moral = buffer.ReadByte
    Map.Up = buffer.ReadLong
    Map.Down = buffer.ReadLong
    Map.Left = buffer.ReadLong
    Map.Right = buffer.ReadLong
    Map.BootMap = buffer.ReadLong
    Map.BootX = buffer.ReadByte
    Map.BootY = buffer.ReadByte
    Map.MaxX = buffer.ReadByte
    Map.MaxY = buffer.ReadByte
    
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Tileset = buffer.ReadLong
            Next
            Map.Tile(X, Y).Type = buffer.ReadByte
            Map.Tile(X, Y).Data1 = buffer.ReadLong
            Map.Tile(X, Y).Data2 = buffer.ReadLong
            Map.Tile(X, Y).Data3 = buffer.ReadLong
            Map.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.Npc(X) = buffer.ReadLong
        n = n + 1
    Next
        
    Set buffer = Nothing
    
    SetStatus "Received map data for Map " & Trim$(CStr(MapNum))
    MapViewTileOffSetX = 0
    MapViewTileOffSetY = 0
    EditorTileWidth = 1
    EditorTileHeight = 1
    EditorTileX = 0
    EditorTileY = 0
    HasMapChanged = False
    
    If CurrentMap > 0 Then ClearAttributeFrames
    CurrentMap = MapNum
    frmEditor.optBlocked.value = True
End Sub

Private Sub HandleUpdateResource(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
        
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ResourceNum = buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set buffer = Nothing
    
    SetStatus "Received Resource Data."
    
End Sub

Private Sub Handlemaxamounts(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
        
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
        
    MAX_MAPS = buffer.ReadLong()
    MAX_ITEMS = buffer.ReadLong()
    MAX_NPCS = buffer.ReadLong()
    MAX_ANIMATIONS = buffer.ReadLong()
    MAX_INV = buffer.ReadLong()
    MAX_MAP_ITEMS = buffer.ReadLong()
    MAX_MAP_NPCS = buffer.ReadLong()
    MAX_SHOPS = buffer.ReadLong()
    MAX_PLAYER_SPELLS = buffer.ReadLong()
    MAX_SPELLS = buffer.ReadLong()
    MAX_RESOURCES = buffer.ReadLong()
    MAX_LEVELS = buffer.ReadLong()
    MAX_BANK = buffer.ReadLong()
    MAX_HOTBAR = buffer.ReadLong()
    
    ReDim Resource(1 To MAX_RESOURCES)
    ReDim Map.Npc(1 To MAX_MAP_NPCS)
    
    Set buffer = Nothing
End Sub
