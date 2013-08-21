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
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Sub SetGlobalSpawnPoint(ByVal Map As Long, ByVal X As Long, ByVal Y As Long)
    ' Check if the map is valid.
    If Map < 0 Or Map > MAX_MAPS Then
        ' It's not, we're using the default values and notify the server.
        START_MAP = 1
        START_X = 5
        START_Y = 5
        
        Call SetStatus("[SCRIPT ERROR] SetSpawnPoint() references a false map.")
        Call SetStatus("Using default values of Map: 1 X: 5 Y: 5")
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

Public Sub GlobalMessage(ByVal Text As String, ByVal Color As Byte)
    Call GlobalMsg(Text, Color)
    Call TextAdd("[SCRIPT] Global Message: " & Trim$(Text))
End Sub

Public Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Public Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
Dim F As Long
    
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

Public Function GetPlayerStr(ByVal Index As Long) As Byte
    GetPlayerStr = GetPlayerStat(Index, Strength)
End Function

Public Function GetPlayerRawStr(ByVal Index As Long) As Byte
    GetPlayerRawStr = GetPlayerRawStatNoClass(Index, Strength)
End Function

Public Function GetPlayerEnd(ByVal Index As Long) As Byte
    GetPlayerEnd = GetPlayerStat(Index, Endurance)
End Function

Public Function GetPlayerRawEnd(ByVal Index As Long) As Byte
    GetPlayerRawEnd = GetPlayerRawStatNoClass(Index, Endurance)
End Function

Public Function GetPlayerAgi(ByVal Index As Long) As Byte
    GetPlayerAgi = GetPlayerStat(Index, Agility)
End Function

Public Function GetPlayerRawAgi(ByVal Index As Long) As Byte
    GetPlayerRawAgi = GetPlayerRawStatNoClass(Index, Agility)
End Function

Public Function GetPlayerInt(ByVal Index As Long) As Byte
    GetPlayerInt = GetPlayerStat(Index, Intelligence)
End Function

Public Function GetPlayerRawInt(ByVal Index As Long) As Byte
    GetPlayerRawInt = GetPlayerRawStatNoClass(Index, Intelligence)
End Function

Public Function GetPlayerWill(ByVal Index As Long) As Byte
    GetPlayerWill = GetPlayerStat(Index, Willpower)
End Function

Public Function GetPlayerRawWill(ByVal Index As Long) As Byte
    GetPlayerRawWill = GetPlayerRawStatNoClass(Index, Willpower)
End Function

Public Sub SetPlayerStr(ByVal Index As Long, ByVal Value As Byte)
    Player(Index).Stat(Strength) = Value
End Sub

Public Sub SetPlayerAgi(ByVal Index As Long, ByVal Value As Byte)
    Player(Index).Stat(Agility) = Value
End Sub

Public Sub SetPlayerEnd(ByVal Index As Long, ByVal Value As Byte)
    Player(Index).Stat(Endurance) = Value
End Sub

Public Sub SetPlayerInt(ByVal Index As Long, ByVal Value As Byte)
    Player(Index).Stat(Intelligence) = Value
End Sub

Public Sub SetPlayerWill(ByVal Index As Long, ByVal Value As Byte)
    Player(Index).Stat(Willpower) = Value
End Sub

Public Sub PlayerMessage(ByVal Index As Long, ByVal Text As String, ByVal Color As Byte)
    Call PlayerMsg(Index, Text, Color)
End Sub

Public Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Public Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Public Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Public Sub SetPlayerClass(ByVal Index As Long, ByVal Class As Long)
    Player(Index).Class = Class
    Call SendPlayerData(Index)
End Sub
