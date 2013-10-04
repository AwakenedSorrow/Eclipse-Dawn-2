Attribute VB_Name = "modEditorTCP"
Option Explicit

Public Sub TcpInit()
    SetLoadStatus LoadStateTCP, 0
    
    '  Set up a fresh buffer object to use.
    Set EditorBuffer = New clsBuffer
    SetLoadStatus LoadStateTCP, 33
    
    ' Set up connection data.
    frmLoad.Socket.RemoteHost = Options.ServerIP
    frmLoad.Socket.RemotePort = Options.ServerPort
    SetLoadStatus LoadStateTCP, 66
    
    ' Initialize messages.
    InitMessages
    SetLoadStatus LoadStateTCP, 100
End Sub

Public Sub DestroyTCP()
    frmLoad.Socket.close
End Sub

Function IsConnected() As Boolean
    
    If frmLoad.Socket.State = sckConnected Then
        IsConnected = True
    End If
End Function

Public Function ConnectToServer() As Boolean
Dim Wait As Long, TempPerc As Long
    
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmLoad.Socket.close
    frmLoad.Socket.Connect
    
    SetLoadStatus LoadStateConnecting, 0
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        TempPerc = LoadBarPerc * (((Wait + 3000) - GetTickCount) / 1500)
        SetLoadStatus LoadStateConnecting, TempPerc
        DoEvents
    Loop
    
    ConnectToServer = IsConnected
End Function

Sub SendData(ByRef data() As Byte)
Dim buffer As clsBuffer
    
    If IsConnected Then
        Set buffer = New clsBuffer
                
        buffer.WriteLong (UBound(data) - LBound(data)) + 1
        buffer.WriteBytes data()
        frmLoad.Socket.SendData buffer.ToArray()
    End If
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    frmLoad.Socket.GetData buffer, vbUnicode, DataLength
    
    EditorBuffer.WriteBytes buffer()
    
    If EditorBuffer.length >= 4 Then pLength = EditorBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= EditorBuffer.length - 4
        If pLength <= EditorBuffer.length - 4 Then
            EditorBuffer.ReadLong
            HandleData EditorBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If EditorBuffer.length >= 4 Then pLength = EditorBuffer.ReadLong(False)
    Loop
    EditorBuffer.Trim
    DoEvents
End Sub

Public Sub SendVersionCheck()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CE_VersionCheck
    
    buffer.WriteString EDITOR_VERSION
    
    SendData buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub SendSaveDeveloper()
Dim buffer As clsBuffer, i As Long, Hash As String
    
    ' Do we have permissions to do this?
    If LCase(Trim$(frmDatabase.txtUsername.Text)) = Trim$(Editor.Username) Then
        If Editor.HasRight(CanEditOwnDetails) <> 1 Then
            MsgBox "Insufficient Permissions, can not change own details.", vbInformation
            Exit Sub
        End If
    Else
        If Editor.HasRight(CanEditDeveloper) <> 1 Then
            MsgBox "Insufficient Permissions, can not edit other developers.", vbInformation
            Exit Sub
        End If
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong CE_SaveDeveloper
    
    buffer.WriteString LCase(Trim$(frmDatabase.txtUsername.Text))
    
    i = InitCrc32()
    i = AddCrc32(frmDatabase.txtPassword.Text, i)
    Hash = CStr(i)
    frmDatabase.txtPassword.Text = Hash
    
    buffer.WriteString Hash
    
    For i = 1 To Editor_MaxRights - 1
        buffer.WriteByte 1
    Next
    
    SendData buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Public Sub SendUserLogin()
Dim buffer As clsBuffer, i As Long, Hash As String
    
    Set buffer = New clsBuffer
    buffer.WriteLong CE_LoginUser
    
    buffer.WriteString LCase(Trim$(frmLogin.txtUsername.Text))
    
    i = InitCrc32()
    i = AddCrc32(frmLogin.txtPassword.Text, i)
    Hash = CStr(i)
    frmLogin.txtPassword.Text = vbNullString
    
    buffer.WriteString Hash
    
    SendData buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub SendRequestMap(ByVal MapNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CE_RequestMap
    
    buffer.WriteLong MapNum
    
    SendData buffer.ToArray
    
    SetStatus "Sent out a request for data on Map " & Trim$(CStr(MapNum))
    Set buffer = Nothing

End Sub
