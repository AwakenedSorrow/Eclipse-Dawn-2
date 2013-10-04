Attribute VB_Name = "modServerEditorTCP"
Option Explicit

Sub IncomingEditorData(ByVal Index As Long, ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long
            
    ' Check if elapsed time has passed
    TempEditor(Index).DataBytes = TempEditor(Index).DataBytes + DataLength
    If GetTickCount >= TempEditor(Index).DataTimer Then
        TempEditor(Index).DataTimer = GetTickCount + 1000
        TempEditor(Index).DataBytes = 0
        TempEditor(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.EditorSocket(Index).GetData buffer(), vbUnicode, DataLength
    TempEditor(Index).buffer.WriteBytes buffer()
    
    If TempEditor(Index).buffer.Length >= 4 Then
        pLength = TempEditor(Index).buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempEditor(Index).buffer.Length - 4
        If pLength <= TempEditor(Index).buffer.Length - 4 Then
            TempEditor(Index).DataPackets = TempEditor(Index).DataPackets + 1
            TempEditor(Index).buffer.ReadLong
            HandleEditorData Index, TempEditor(Index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempEditor(Index).buffer.Length >= 4 Then
            pLength = TempEditor(Index).buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempEditor(Index).buffer.Trim
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
            SendEditorAlertMsg Index, "The server appears to be full, try again later.", True
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
Dim buffer As clsBuffer
Dim TempData() As Byte

    If IsEditorConnected(Index) Then
        Set buffer = New clsBuffer
        
        buffer.PreAllocate 4 + (UBound(Data) - LBound(Data)) + 1
        buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        buffer.WriteBytes Data()
              
        frmServer.EditorSocket(Index).SendData buffer.ToArray()
        
        '  Experimental
        DoEvents
        
        Set buffer = Nothing
    End If
End Sub

Public Sub SendEditorAlertMsg(ByVal Index As Byte, ByVal Message As String, Disconnect As Boolean)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SE_AlertMsg
    buffer.WriteString Message
    If Disconnect Then
        buffer.WriteByte 1
    Else
        buffer.WriteByte 0
    End If
    
    SendEditorDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

End Sub

Sub SendEditorMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    buffer.WriteLong SE_MapData
    buffer.WriteBytes MapCache(MapNum).Data()
    SendEditorDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Public Sub SendEditorVersionOK(ByVal Index As Byte)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SE_VersionOK
    
    SendEditorDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Public Sub SendEditorLoginOK(ByVal Index As Byte)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    
    buffer.WriteLong SE_LoginOK
    
    buffer.WriteString Editor(Index).Username
    For i = 1 To Editor_MaxRights - 1
        buffer.WriteByte Editor(Index).HasRight(i)
    Next
    
    SendEditorDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
    
End Sub

Public Sub SendMapEditorNames(ByVal Index As Byte)
Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SE_MapNames
    
    buffer.WriteLong MAX_MAPS
    For i = 1 To MAX_MAPS
        buffer.WriteString Map(i).Name
        buffer.WriteLong Map(i).Revision
    Next
    
    SendEditorDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
    
End Sub

Sub SendEditorResources(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendEditorUpdateResourceTo(Index, i)
        End If

    Next

End Sub

Sub SendEditorUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SE_ResourceData
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    
    SendEditorDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

