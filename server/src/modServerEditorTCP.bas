Attribute VB_Name = "modServerEditorTCP"
Option Explicit

Sub IncomingEditorData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long
            
    ' Check if elapsed time has passed
    TempEditor(Index).DataBytes = TempEditor(Index).DataBytes + DataLength
    If GetTickCount >= TempEditor(Index).DataTimer Then
        TempEditor(Index).DataTimer = GetTickCount + 1000
        TempEditor(Index).DataBytes = 0
        TempEditor(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.EditorSocket(Index).GetData Buffer(), vbUnicode, DataLength
    TempEditor(Index).Buffer.WriteBytes Buffer()
    
    If TempEditor(Index).Buffer.Length >= 4 Then
        pLength = TempEditor(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempEditor(Index).Buffer.Length - 4
        If pLength <= TempEditor(Index).Buffer.Length - 4 Then
            TempEditor(Index).DataPackets = TempEditor(Index).DataPackets + 1
            TempEditor(Index).Buffer.ReadLong
            HandleEditorData Index, TempEditor(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempEditor(Index).Buffer.Length >= 4 Then
            pLength = TempEditor(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempEditor(Index).Buffer.Trim
End Sub

Sub AcceptEditorConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (Index = 0) Then
        i = FindOpenEditorSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.EditorSocket(i).Close
            frmServer.EditorSocket(i).Accept SocketId
            Call TextAdd("Received connection from " & GetEditorIP(Index) & ".")
        Else
            SendEditorAlertMsg Index, "The server appears to be full, try again later."
        End If
    End If

End Sub

Sub CloseEditorSocket(ByVal Index As Long)

    If Index > 0 Then
        Call TextAdd("Connection from " & GetEditorIP(Index) & " has been terminated.")
        frmServer.EditorSocket(Index).Close
        Call ClearEditor(Index)
    End If

End Sub

Function IsEditorConnected(ByVal Index As Long) As Boolean

    If frmServer.EditorSocket(Index).State = sckConnected Then
        IsEditorConnected = True
    End If

End Function

Sub SendEditorDataTo(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsEditorConnected(Index) Then
        Set Buffer = New clsBuffer
        
        Buffer.PreAllocate 4 + (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data()
              
        frmServer.EditorSocket(Index).SendData Buffer.ToArray()
        
        '  Experimental
        DoEvents
        
        Set Buffer = Nothing
    End If
End Sub

Public Sub SendEditorAlertMsg(ByVal Index As Byte, ByVal Message As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SE_AlertMsg
    Buffer.WriteString Message
    
    SendEditorDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Public Sub SendEditorVersionOK(ByVal Index As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SE_VersionOK
    
    SendEditorDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub
