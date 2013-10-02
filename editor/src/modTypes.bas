Attribute VB_Name = "modTypes"
Option Explicit

Public Options As OptionsRec

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
    Device As Byte
End Type
