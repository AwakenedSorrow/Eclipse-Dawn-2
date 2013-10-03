Attribute VB_Name = "modTypes"
Option Explicit

Public Options As OptionsRec
Public Editor As EditorRec
Public TempEditor As EditorRec

'  **************************************
'  **************************************
'  **************************************

Type OptionsRec
    '  Server Related
    ServerIP As String
    ServerPort As Long
    
    ' Account
    RememberUser As Byte
    Username As String
    
    ' Debug
    device As Byte
End Type

Private Type EditorRec
    Username As String
    Password As String
    
    HasRight(Editor_MaxRights - 1) As Byte
End Type
