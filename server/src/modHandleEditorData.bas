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
            SendMapEditorNames Index
            SendEditorResources Index
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
