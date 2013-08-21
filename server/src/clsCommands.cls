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
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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
        
        Call SetStatus("[SCRIPT FAULT] SetSpawnPoint() references a false map.")
        Call SetStatus("Using default values of Map: 1 X: 5 Y: 5")
    Else
        ' We're using a correct value. Let's set it!
        START_MAP = Map
        START_X = X
        START_Y = Y
        
        Call SetStatus("[SCRIPT] Global Spawnpoint has been set to Map: " & Trim$(STR$(Map)) & " X: " & Trim$(STR$(X)) & " Y: " & Trim$(STR$(Y)))
    End If
End Sub
