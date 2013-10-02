Attribute VB_Name = "modHandleEditorData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitEditorMessages()
    HandleEditorDataSub(CE_LoginUser) = GetAddress(AddressOf HandleEditorLogin)
    HandleEditorDataSub(CE_VersionCheck) = GetAddress(AddressOf HandleEditorVersionCheck)
End Sub

Public Sub HandleEditorData(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CE_MSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleEditorDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleEditorLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

Private Sub HandleEditorVersionCheck(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim TempVer As String
    
    '  Make sure our editor is actually connected.
    If IsEditorConnected(Index) Then
        '  Write the data to the local buffer so we can extract the required data.
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        '  Retrieve the version of the editor client.
        TempVer = Buffer.ReadString()
        
        ' Does it match? If not, disconnect them with a warning.
        If Trim$(TempVer) <> EDITOR_VERSION Then
            SendEditorAlertMsg Index, "Your editor appears to be outdated, please request the latest version from your server administrator!"
        Else
            SendEditorVersionOK Index
        End If
    End If
End Sub
