Attribute VB_Name = "modHandleData"
Option Explicit

Public Sub InitMessages()
    HandleDataSub(SE_AlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SE_VersionOK) = GetAddress(AddressOf HandleVersionOK)
    HandleDataSub(SE_LoginOK) = GetAddress(AddressOf HandleLoginOK)
    HandleDataSub(SE_MapNames) = GetAddress(AddressOf HandleMapNames)
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
Dim Msg As String
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
        
    Msg = buffer.ReadString
    
    Set buffer = Nothing
    MsgBox Msg, vbOKOnly, "Error"
    
    '  An Alert Message means something went horribly wrong and we can't continue.
    DestroyEditor
    
End Sub

Private Sub HandleVersionOK(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' We've got our confirmation, time to stop checking versions.
    CheckingVersion = False
End Sub

Private Sub HandleMapNames(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, TempName As String, TempRev As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    MAX_MAPS = buffer.ReadLong()
    frmEditor.lstMapList.Clear
    For i = 1 To MAX_MAPS
        TempName = buffer.ReadString()
        TempRev = buffer.ReadLong()
        frmEditor.lstMapList.AddItem CStr(i) & ": " & Trim$(TempName) & " | Rev." & Trim(CStr(TempRev))
    Next
    
    Set buffer = Nothing
    
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
