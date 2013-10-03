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
Dim Buffer As clsBuffer
    
    If IsConnected Then
        Set Buffer = New clsBuffer
                
        Buffer.WriteLong (UBound(data) - LBound(data)) + 1
        Buffer.WriteBytes data()
        frmLoad.Socket.SendData Buffer.ToArray()
    End If
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    frmLoad.Socket.GetData Buffer, vbUnicode, DataLength
    
    EditorBuffer.WriteBytes Buffer()
    
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
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CE_VersionCheck
    
    Buffer.WriteString EDITOR_VERSION
    
    SendData Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub SendSaveDeveloper()
Dim Buffer As clsBuffer, i As Long, Hash As String

    Set Buffer = New clsBuffer
    Buffer.WriteLong CE_SaveDeveloper
    
    Buffer.WriteString LCase(Trim$(frmDatabase.txtUsername.Text))
    
    i = InitCrc32()
    i = AddCrc32(frmDatabase.txtPassword.Text, i)
    Hash = CStr(i)
    frmDatabase.txtPassword.Text = Hash
    
    Buffer.WriteString Hash
    
    For i = 1 To Editor_MaxRights - 1
        Buffer.WriteByte 1
    Next
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Public Sub SendUserLogin()
Dim Buffer As clsBuffer, i As Long, Hash As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CE_LoginUser
    
    Buffer.WriteString LCase(Trim$(frmLogin.txtUsername.Text))
    
    i = InitCrc32()
    i = AddCrc32(frmLogin.txtPassword.Text, i)
    Hash = CStr(i)
    frmLogin.txtPassword.Text = Hash
    
    Buffer.WriteString Hash
    
    SendData Buffer.ToArray
    
    Set Buffer = Nothing
End Sub
