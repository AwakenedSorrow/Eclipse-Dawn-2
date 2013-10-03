Attribute VB_Name = "modHandleData"
Option Explicit

Public Sub InitMessages()
    HandleDataSub(SE_AlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SE_VersionOK) = GetAddress(AddressOf HandleVersionOK)
    HandleDataSub(SE_LoginOK) = GetAddress(AddressOf HandleLoginOK)
End Sub

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Sub HandleData(ByRef data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyEditor
        Exit Sub
    End If
    
    If MsgType >= SE_MSG_COUNT Then
        DestroyEditor
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.length), 0, 0
    
End Sub

Private Sub HandleAlertMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
        
    Msg = Buffer.ReadString
    
    Set Buffer = Nothing
    MsgBox Msg, vbOKOnly, "Error"
    
    '  An Alert Message means something went horribly wrong and we can't continue.
    DestroyEditor
    
End Sub

Private Sub HandleVersionOK(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' We've got our confirmation, time to stop checking versions.
    CheckingVersion = False
End Sub

Private Sub HandleLoginOK(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Long
    ' We've got our confirmation, time to stop the timeout loop.
    LoggingIn = False
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Editor.Username = Buffer.ReadString
    
    For i = 1 To Editor_MaxRights - 1
        Editor.HasRight(i) = Buffer.ReadByte()
    Next
    
    Set Buffer = Nothing
    
End Sub
