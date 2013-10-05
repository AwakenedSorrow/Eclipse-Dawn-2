Attribute VB_Name = "modEditorLogic"
Option Explicit

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
        
        SetStatus "Selected a tile at location X" & Trim$(CStr(EditorTileX)) & " Y" & Trim$(CStr(EditorTileY))
    End If
    
End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > TEXTURE_WIDTH / PIC_X Then X = TEXTURE_WIDTH / PIC_X
        If Y < 0 Then Y = 0
        If Y > TileSelectHeight / PIC_Y Then Y = TileSelectHeight / PIC_Y
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        End If
        SetStatus "Selected a tile at location X" & Trim$(CStr(EditorTileX)) & " Y" & Trim$(CStr(EditorTileY))
    End If

End Sub

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
        
End Function

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
Dim i As Long
Dim CurLayer As Long
Dim tmpDir As Byte

    ' find which layer we're on
    CurLayer = frmEditor.cmbLayerSelect.ListIndex + 1
    
    If Not IsValidMapPoint(CurX, CurY) Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor.cmbLayerSelect.ListIndex + 1 < Layer_Count Then
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer
                HasMapChanged = True
            Else ' multi tile!
                MapEditorSetTile CurX, CurY, CurLayer, True
                HasMapChanged = True
            End If
        ElseIf frmEditor.cmbLayerSelect.ListIndex + 1 = Layer_Count Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor.optBlocked.value Then .Type = TileTypeBlocked: HasMapChanged = True
                ' warp tile
                If frmEditor.optWarp.value Then
                    .Type = TileTypeWarp
                    .Data1 = frmEditor.cmbWarpMap.ListIndex + 1
                    .Data2 = Val(frmEditor.txtWarpX.text)
                    .Data3 = Val(frmEditor.txtWarpY.text)
                    HasMapChanged = True
                End If
                ' heal
                If frmEditor.optHeal.value Then
                    .Type = TileTypeHeal
                    .Data1 = frmEditor.cmbHeal.ListIndex + 1
                    .Data2 = Val(frmEditor.txtHealAmount.text)
                    .Data3 = 0
                    HasMapChanged = True
                End If
                ' trap
                If frmEditor.optTrap.value Then
                    .Type = TileTypeTrap
                    .Data1 = Val(frmEditor.txtDamageAmount.text)
                    .Data2 = 0
                    .Data3 = 0
                    HasMapChanged = True
                End If
                ' slide
                If frmEditor.optSlide.value Then
                    .Type = TileTypeSlide
                    .Data1 = frmEditor.cmbSlide.ListIndex
                    .Data2 = 0
                    .Data3 = 0
                    HasMapChanged = True
                End If
            End With
        ElseIf frmEditor.optDirBlock.value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)
            ' see if it hits an arrow
            'For i = 1 To 4
            '    If X >= DirArrowX(i) And X <= DirArrowX(i) + 8 Then
            '        If Y >= DirArrowY(i) And Y <= DirArrowY(i) + 8 Then
            '            ' flip the value.
            '            setDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(i))
            '            Exit Sub
            '        End If
            '    End If
            'Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor.cmbLayerSelect.ListIndex + 1 < Layer_Count Then
            With Map.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).X = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).Tileset = 0
                HasMapChanged = True
            End With
        ElseIf frmEditor.cmbLayerSelect.ListIndex + 1 = Layer_Count Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
                HasMapChanged = True
            End With

        End If
    End If

End Sub

Public Sub ClearAttributeFrames()
Dim i As Long
    '  Clear Block
    frmEditor.fraBlock.Visible = False
    
    '  Clear Warp
    frmEditor.cmbWarpMap.Clear
    For i = 1 To MAX_MAPS
        frmEditor.cmbWarpMap.AddItem frmEditor.lstMapList.List(i - 1)
    Next i
    frmEditor.cmbWarpMap.ListIndex = 0
    frmEditor.txtWarpX.text = "0"
    frmEditor.txtWarpY.text = "0"
    frmEditor.fraWarp.Visible = False
    
    '  CLear Slide
    frmEditor.cmbSlide.ListIndex = 0
    frmEditor.fraSlide.Visible = False
    
    ' Clear Heal
    frmEditor.cmbHeal.ListIndex = 0
    frmEditor.txtHealAmount.text = "0"
    frmEditor.fraHeal.Visible = False
    
    ' Clear Damage
    frmEditor.cmbDamage.ListIndex = 0
    frmEditor.txtDamageAmount.text = "0"
    frmEditor.fraDamage.Visible = False
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False)
Dim X2 As Long, Y2 As Long, CurX As Long, CurY As Long
    If Not multitile Then ' single
        With Map.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor.cmbTileSet.ListIndex + 1
        End With
    Else ' multitile
        CurY = Y
        CurX = X
        Y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            .Layer(CurLayer).X = EditorTileX + X2
                            .Layer(CurLayer).Y = EditorTileY + Y2
                            .Layer(CurLayer).Tileset = frmEditor.cmbTileSet.ListIndex + 1
                        End With
                    End If
                End If
                X2 = X2 + 1
            Next
            Y2 = Y2 + 1
        Next
    End If
End Sub
