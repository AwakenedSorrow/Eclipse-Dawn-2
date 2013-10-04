Attribute VB_Name = "modEditorLogic"
Public Sub EditorLoop()
Dim Tick As Long
    Do While EditorLooping = True
        Tick = GetTickCount
        
        ' Render the graphics on our displays.
        RenderGraphics
        
        ' Handle the forms and whatnot
        DoEvents
        
        ' Lock the FPS. I mean it's just an editor for christ's sake.
        Do While GetTickCount < Tick + 15
            DoEvents
            Sleep 1
        Loop
    Loop
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
        
    End If
    
End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > TEXTURE_WIDTH / PIC_X Then X = texture_wifth / PIC_X
        If Y < 0 Then Y = 0
        If Y > TileSelectHeight / PIC_Y Then Y = TileSelectHeight / PIC_Y
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        End If
    End If

End Sub
