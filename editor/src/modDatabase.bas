Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)

End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If
    
End Function

' gets a string from a text file
Public Function GetVar(file As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

' writes a variable to a text file
Public Sub PutVar(file As String, Header As String, Var As String, value As String)
    
    Call WritePrivateProfileString$(Header, Var, value, file)
    
End Sub

Public Sub LoadOptions(ByVal FileName As String)
Dim i As Long

    ' Reset the option loading percentage to 0.
    SetLoadStatus LoadStateOptions, 0
    
    ' Check if said file exists before we start using it.
    If FileExist(FileName, True) Then
        ' The file exists, so we can simply load the settings from it.
        
        '  Load Server Settings
        Options.ServerIP = GetVar(FileName, "SERVER", "ServerIP")
        Options.ServerPort = CLng(GetVar(FileName, "SERVER", "ServerPort"))
        SetLoadStatus LoadStateOptions, 33
        
        ' Load Account Settings
        Options.RememberUser = CByte(GetVar(FileName, "ACCOUNT", "RememberUser"))
        Options.Username = GetVar(FileName, "ACCOUNT", "Username")
        SetLoadStatus LoadStateOptions, 66
        
        ' Load Tileset Settings
        If NumTileSets > 0 Then
            ReDim Preserve Options.TileSetName(1 To NumTileSets)
            For i = 1 To NumTileSets
                Options.TileSetName(i) = GetVar(FileName, "TILESET", "Name" & CStr(i))
            Next i
        End If
        
        '  Load Debug Settings
        Options.device = CByte(GetVar(FileName, "DEBUG", "Device"))
        SetLoadStatus LoadStateOptions, 100
    Else
        '  The file does not exist, so we need to manually set the options and save the file to make sure it does next time.
        Options.ServerIP = "localhost"
        Options.ServerPort = 8001
        
        Options.RememberUser = 0
        Options.Username = vbNullString
        
        If NumTileSets > 0 Then
            ReDim Preserve Options.TileSetName(1 To NumTileSets)
            For i = 1 To NumTileSets
                Options.TileSetName(i) = ""
            Next i
        End If
        
        Options.device = 2
        SetLoadStatus LoadStateOptions, 50
        
        SaveOptions FileName
        SetLoadStatus LoadStateOptions, 100
    End If
End Sub

Public Sub SaveOptions(ByVal FileName As String)
    '  Save Server Settings
    PutVar FileName, "SERVER", "ServerIP", Options.ServerIP
    PutVar FileName, "SERVER", "ServerPort", Trim$(CStr(Options.ServerPort))
        
    ' Load Account Settings
    PutVar FileName, "ACCOUNT", "RememberUser", Trim$(CStr(Options.RememberUser))
    PutVar FileName, "ACCOUNT", "Username", Trim$(Options.Username)
    
    If NumTileSets > 0 Then
        ReDim Preserve Options.TileSetName(1 To NumTileSets)
        For i = 1 To NumTileSets
            PutVar FileName, "TILESET", "Name" & CStr(i), Options.TileSetName(i)
        Next i
    End If
        
    '  Load Debug Settings
    PutVar FileName, "DEBUG", "Device", Trim$(CStr(Options.device))
End Sub

Public Sub ClearEditor()
Dim i As Long

    Editor.Username = vbNullString
    For i = 1 To Editor_MaxRights - 1
        Editor.HasRight(i) = 0
    Next

End Sub
