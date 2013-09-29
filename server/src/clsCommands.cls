VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This file contains commands that're accessible by the scripting language.
' Please do not put ANYTHING in here that should be used elsewhere, or not
' accessible by scripts.
Option Explicit

' Declaring some lovely things we need for our scripting commands.
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space(5000)
    File = App.Path & "\" & File
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Dim fpath As String
    fpath = App.Path & "\" & File
    Call WritePrivateProfileString(Header, Var, Value, fpath)

    On Error GoTo 0
    Exit Sub
End Sub

Public Sub SetGlobalSpawnPoint(ByVal Map As Long, ByVal X As Long, ByVal Y As Long)
    ' Check if the map is valid.
    If Map < 0 Or Map > MAX_MAPS Then
        ' It's not, we're using the default values and notify the server.
        START_MAP = 1
        START_X = 7
        START_Y = 6
        
        Call SetStatus("[SCRIPT ERROR] SetSpawnPoint() references a false map.")
        Call SetStatus("Using default values of Map: 1 X: 7 Y: 6")
    Else
        ' We're using a correct value. Let's set it!
        START_MAP = Map
        START_X = X
        START_Y = Y
        
        Call SetStatus("[SCRIPT] Global Spawnpoint has been set to Map: " & Trim$(STR$(Map)) & " X: " & Trim$(STR$(X)) & " Y: " & Trim$(STR$(Y)))
    End If
End Sub

Public Sub ShowServerText(ByVal Text As String)
    Call TextAdd(Text)
End Sub

Public Sub GlobalMsg(ByVal Text As String, ByVal Color As Byte)
    Call modServerTCP.GlobalMsg(Text, Color)
    Call TextAdd("[SCRIPT] Global Message: " & Trim$(Text))
End Sub

Public Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(modPlayer.GetPlayerName(Index))
End Function

Public Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    modPlayer.SetPlayerName Index, Name
End Sub

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Byte) As Long
    GetPlayerStat = modPlayer.GetPlayerStat(Index, Stat)
End Function

Public Function GetPlayerRawStat(ByVal Index As Long, ByVal Stat As Byte) As Long
    GetPlayerRawStat = modPlayer.GetPlayerRawStat(Index, Stat)
End Function

Public Function GetPlayerRawStatNoClass(ByVal Index As Long, ByVal Stat As Byte) As Long
    GetPlayerRawStatNoClass = modPlayer.GetPlayerRawStatNoClass(Index, Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Byte, ByVal Value As Long)
    modPlayer.SetPlayerStat Index, Stat, Value
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Text As String, ByVal Color As Byte)
    modServerTCP.PlayerMsg Index, Text, Color
End Sub

Public Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = modPlayer.GetPlayerLevel(Index)
End Function

Public Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    modPlayer.SetPlayerLevel Index, Level
End Sub

Public Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = modPlayer.GetPlayerClass(Index)
End Function

Public Sub SetPlayerClass(ByVal Index As Long, ByVal Class As Long)
    modPlayer.SetPlayerClass Index, Class
End Sub

Public Function GetPlayerPoints(ByVal Index As Long) As Byte
    GetPlayerPoints = modPlayer.GetPlayerPoints(Index)
End Function

Public Sub SetPlayerPoints(ByVal Index As Long, ByVal Points As Byte)
    modPlayer.SetPlayerPoints Index, Points
End Sub

Public Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = modPlayer.GetPlayerMap(Index)
End Function

Public Function GetPlayerSex(ByVal Index As Long) As Byte
    GetPlayerSex = Player(Index).Sex
End Function

Public Sub WarpPlayer(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Call modPlayer.PlayerWarp(Index, MapNum, X, Y)
End Sub

Public Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = modPlayer.GetPlayerSprite(Index)
End Function

Public Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    modPlayer.SetPlayerSprite Index, Sprite
End Sub

Public Function PlayerHasItem(ByVal Index As Long, ByVal Num As Long, ByVal Value As Long) As Boolean
Dim i As Byte

    PlayerHasItem = False
    
    For i = 1 To MAX_INV
        If Player(Index).Inv(i).Num = Num Then
            If Item(Num).Type = ItemTypeCurrency Then
                If Player(Index).Inv(i).Value >= Value Then
                    PlayerHasItem = True
                    Exit Function
                End If
            Else
                PlayerHasItem = True
                Exit Function
            End If
        End If
    Next
    
End Function

Public Function GetInvItemSlot(ByVal Index As Long, ByVal Item As Long) As Byte
Dim i As Byte

    GetInvItemSlot = 0
    
    For i = 1 To MAX_INV
        If Player(Index).Inv(i).Num = Item Then
            GetInvItemSlot = i
            Exit Function
        End If
    Next

End Function

Public Sub SetInvItemSlot(ByVal Index As Long, ByVal Slot As Long, ByVal Item As Long)
    If Index < 1 Or Index > Player_HighIndex Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub
    If Item < 1 Or Item > MAX_ITEMS Then Exit Sub
    If Slot < 1 Or Slot > MAX_INV Then Exit Sub
    
    Player(Index).Inv(Slot).Num = Item
End Sub

Public Function GetInvItemVal(ByVal Index As Long, ByVal Slot As Long) As Long
    GetInvItemVal = Player(Index).Inv(Slot).Value
End Function

Public Sub SetInvItemVal(ByVal Index As Long, ByVal Slot As Long, ByVal Value As Long)
    If Index < 1 Or Index > Player_HighIndex Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub
    If Slot < 1 Or Slot > MAX_INV Then Exit Sub
    
    Player(Index).Inv(Slot).Value = Value
End Sub

Public Function TakePlayerItem(ByVal Index As Long, ByVal Item As Long, ByVal Value As Long) As Boolean
    TakePlayerItem = TakeInvItem(Index, Item, Value)
End Function

Public Function GivePlayerItem(ByVal Index As Long, ByVal Item As Long, ByVal Value As Long) As Boolean
    GivePlayerItem = GiveInvItem(Index, Item, Value, True)
End Function

Public Sub MapMsg(ByVal MapNum As Long, ByVal Text As String, ByVal Color As Byte)
    modServerTCP.MapMsg MapNum, Text, Color
End Sub

Public Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = modPlayer.GetPlayerX(Index)
End Function

Public Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = modPlayer.GetPlayerY(Index)
End Function

Public Function GetPlayerDir(ByVal Index As Long) As Byte
    GetPlayerDir = modPlayer.GetPlayerDir(Index)
End Function

Public Sub SetPlayerDir(ByVal Index As Long, Dir As Byte)
    modPlayer.SetPlayerDir Index, Dir
End Sub

Public Function GetPlayerTarget(ByVal Index As Long) As Long
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerTarget = TempPlayer(Index).Target
End Function

Public Function GetPlayerTargetType(ByVal Index As Long) As Byte
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerTargetType = TempPlayer(Index).TargetType
End Function

Public Function GetNPCX(ByVal Map As Long, ByVal MapNPCNum As Long) As Long
    If Map < 1 Or Map > MAX_MAPS Then Exit Function
    If MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Then Exit Function
    
    GetNPCX = MapNpc(Map).Npc(MapNPCNum).X
End Function

Public Function GetNPCY(ByVal Map As Long, ByVal MapNPCNum As Long) As Long
    If Map < 1 Or Map > MAX_MAPS Then Exit Function
    If MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Then Exit Function
    
    GetNPCY = MapNpc(Map).Npc(MapNPCNum).Y
End Function

Public Function GetNPCName(ByVal MapNum As Long, ByVal MapNPCNum As Long) As String
    GetNPCName = Trim$(Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Name)
End Function

Public Sub SendAnimation(ByVal Map As Long, ByVal X As Long, ByVal Y As Long, ByVal Animation As Long)
    modServerTCP.SendAnimation Map, Animation, X, Y
End Sub

Public Function GetOriginNPCNum(ByVal MapNum As Long, ByVal Num As Long) As Long
    GetOriginNPCNum = MapNpc(MapNum).Npc(Num).Num
End Function

Public Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Byte) As Long
    GetPlayerVital = modPlayer.GetPlayerVital(Index, Vital)
End Function

Public Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Byte) As Long
    GetPlayerMaxVital = modCombat.GetPlayerMaxVital(Index, Vital)
End Function

Public Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Byte, ByVal Value As Long)
    modPlayer.SetPlayerVital Index, Vital, Value
End Sub

Public Sub SendPlayerData(ByVal Index As Long)
    modServerTCP.SendPlayerData Index
End Sub

Public Sub SendActionMsg(ByVal Map As Long, ByVal Text As String, ByVal Color As Byte, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
    modServerTCP.SendActionMsg Map, Text, Color, MsgType, X, Y
End Sub

Public Function GetPlayerEquipment(ByVal Index As Long, ByVal Slot As Byte) As Long
    GetPlayerEquipment = modPlayer.GetPlayerEquipment(Index, Slot)
End Function

Public Sub SetPlayerEquipment(ByVal Index As Long, ByVal Slot As Byte, ByVal InvSlot As Long)
    modPlayer.SetPlayerEquipment Index, InvSlot, Slot
End Sub

Public Sub PlayerMapDropItem(ByVal Index As Long, InvSlot As Long, Value As Long)
    modPlayer.PlayerMapDropItem Index, InvSlot, Value
End Sub

Public Function GetMapBootMap(ByVal MapNum As Long) As Long
    GetMapBootMap = Map(MapNum).BootMap
End Function

Public Function GetMapBootX(ByVal MapNum As Long) As Long
    GetMapBootX = Map(MapNum).BootX
End Function

Public Function GetMapBootY(ByVal MapNum As Long) As Long
    GetMapBootY = Map(MapNum).BootY
End Function

Public Function RAND(ByVal Num1 As Long, ByVal Num2 As Long) As Long
    RAND = modGameLogic.RAND(Num1, Num2)
End Function

Public Function GetNPCTarget(ByVal MapNum As Long, ByVal NPCNum As Long) As Long
    GetNPCTarget = MapNpc(MapNum).Npc(NPCNum).Target
End Function

Public Function GetNPCTargetType(ByVal MapNum As Long, ByVal NPCNum As Long) As Long
    GetNPCTargetType = MapNpc(MapNum).Npc(NPCNum).TargetType
End Function

Public Function GetNPCMaxVital(ByVal MapNum As Long, ByVal NPCNum As Long, ByVal Vital As Byte) As Long
    GetNPCMaxVital = modCombat.GetNPCMaxVital(Map(MapNum).Npc(NPCNum), Vital)
End Function

Public Function GetNPCVital(ByVal MapNum As Long, ByVal NPCNum As Long, ByVal Vital As Byte) As Long
    GetNPCVital = MapNpc(MapNum).Npc(NPCNum).Vital(Vital)
End Function

Public Sub SetNPCVital(ByVal MapNum As Long, ByVal NPCNum As Long, ByVal Vital As Byte, ByVal Num As Long)
    MapNpc(MapNum).Npc(NPCNum).Vital(Vital) = Num
End Sub
