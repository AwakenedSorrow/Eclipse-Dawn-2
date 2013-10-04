Attribute VB_Name = "modRendering"
Option Explicit



Public Sub RenderGraphics()
    ' Render Map Editor Components
    MapEditor_DrawTileSet
    MapEditor_DrawMapView
    
End Sub

Public Sub MapEditor_DrawMapView()
Dim X As Long, Y As Long, OY As Long, EY As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
     
    ' Render the backdrop.
    For X = 0 To (MapViewWidth / PIC_X)
        For Y = 0 To (MapViewHeight / PIC_Y)
            Call RenderGraphic(Tex_EditorBackDrop, X * PIC_X, Y * PIC_Y, PIC_X, PIC_Y, 0, 0, 0, 0)
        Next
    Next
    
    If Map.Revision <> 0 Then
        '  Shows us where we can edit on empty maps.
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Call RenderGraphic(Tex_Blank, (X * PIC_X) + (MapViewTileOffSetX * PIC_X), (Y * PIC_Y) + (MapViewTileOffSetY * PIC_Y), PIC_X, PIC_Y, 0, 0, 0, 0)
            Next
        Next
        
        '  Tiles
        If NumTileSets > 0 Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    If IsValidMapPoint(X, Y) Then
                        Call RenderMapTile(X, Y)
                    End If
                Next
            Next
        End If
        
        '  Now for the resources
        For Y = 0 To Map.MaxY
            For X = 0 To Map.MaxX
                If IsValidMapPoint(X, Y) Then
                    ' Check if there's a resource on this tile.
                    If Map.Tile(X, Y).Type = TileTypeResource Then
                        DrawResource Map.Tile(X, Y).Data1, X, Y
                    End If
                End If
            Next
        Next
        
        '  Attribute Labels
        If frmEditor.cmbLayerSelect.ListIndex + 1 = Layer_Count Then DrawMapAttributes
    End If
    
    ' Render the mouse outline.
    If ShowMouse = True Then
        X = (MouseX / PIC_X)
        Y = (MouseY / PIC_Y)
        RenderTileOutline X, Y
    End If
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = MapViewWidth
        .Y1 = 0
        .Y2 = MapViewHeight
    End With
    
    With destRect
        .X1 = MapViewWindow.X1
        .X2 = MapViewWindow.X2
        .Y1 = MapViewWindow.Y1
        .Y2 = MapViewWindow.Y2
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor.hWnd, ByVal 0)
    
End Sub

Public Sub DrawResource(ByVal ResourceNum As Long, ByVal X As Long, ByVal Y As Long)
Dim ResourcePic As Long
    
    ResourcePic = Resource(ResourceNum).ResourceImage
    X = ((X * PIC_X) + (MapViewTileOffSetX * PIC_X)) - (D3DT_TEXTURE(Tex_Resource(ResourcePic)).Width / 2) + 16
    Y = ((Y * PIC_Y) + (MapViewTileOffSetY * PIC_Y) + PIC_Y) - D3DT_TEXTURE(Tex_Resource(ResourcePic)).Height
    Call RenderGraphic(Tex_Resource(ResourcePic), X, Y, D3DT_TEXTURE(Tex_Resource(ResourcePic)).Width, D3DT_TEXTURE(Tex_Resource(ResourcePic)).Height, 0, 0, 0, 0, Resource(ResourceNum).Red(1), Resource(ResourceNum).Green(1), Resource(ResourceNum).Blue(1), Resource(ResourceNum).Alpha(1))
End Sub

Public Sub MapEditor_DrawTileSet()
Dim Tileset As Long, X As Long, Y As Long, OY As Long, EY As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long

    Tileset = frmEditor.cmbTileSet.ListIndex + 1
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
     
    ' Render the backdrop.
    For X = 0 To (TEXTURE_WIDTH / PIC_X) - 1
        For Y = 0 To (TileSelectHeight / PIC_Y)
            Call RenderGraphic(Tex_EditorBackDrop, X * PIC_X, Y * PIC_Y, PIC_X, PIC_Y, 0, 0, 0, 0)
        Next
    Next
    
    ' Does the tileset exist? If so, render it.
    If Tileset >= 1 Or Tileset <= NumTileSets Then
        If D3DT_TEXTURE(Tex_TileSet(Tileset)).Height > TileSelectHeight Then
            frmEditor.scrlTileSelect.max = (D3DT_TEXTURE(Tex_TileSet(Tileset)).Height \ 32) - (TileSelectHeight \ 32) - 1
        End If
        
        OY = frmEditor.scrlTileSelect.value * 32
        EY = OY + TileSelectHeight
        
        Call RenderGraphic(Tex_TileSet(Tileset), 0, 0, TEXTURE_WIDTH, EY, 0, 0, 0, OY)
        
        ' Render the tile selection square.
    X1 = (EditorTileX * PIC_X)
    X2 = (EditorTileWidth * PIC_X) + X1
    Y1 = (EditorTileY * PIC_Y) - frmEditor.scrlTileSelect.value * 32
    Y2 = (EditorTileHeight * PIC_Y) + Y1
    Call DrawSelectionBox(X1, X2, Y1, Y2)
    End If
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = TEXTURE_WIDTH
        .Y1 = 0
        .Y2 = TileSelectHeight
    End With
    
    With destRect
        .X1 = RecTileSelectWindow.X1
        .X2 = RecTileSelectWindow.X2
        .Y1 = RecTileSelectWindow.Y1
        .Y2 = RecTileSelectWindow.Y2
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor.hWnd, ByVal 0)
    
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long
    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If
End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long
    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If
End Function

Public Sub DrawSelectionBox(ByVal X1 As Long, ByVal X2 As Long, ByVal Y1 As Long, ByVal Y2 As Long)
Dim Width As Long, Height As Long, X As Long, Y As Long
    Width = X2 - X1
    Height = Y2 - Y1
    X = X1
    Y = Y1
    If Width > 6 And Height > 6 Then
        Call RenderGraphic(Tex_Select, X, Y, Width, 3, 0, 0, 0, 0, 255, 0, 255)                 'Top Bar
        Call RenderGraphic(Tex_Select, X, Y, 3, Height, 0, 0, 0, 0, 255, 0, 255)                'Left bar
        Call RenderGraphic(Tex_Select, X, Y + Height - 3, Width, 3, 0, 0, 0, 0, 255, 0, 255)    'Bottom Bar
        Call RenderGraphic(Tex_Select, X + Width - 3, Y, 3, Height, 0, 0, 0, 0, 255, 0, 255)    'Right Bar
    End If
End Sub

Public Sub RenderMapTile(ByVal X As Long, ByVal Y As Long)
Dim i As Long

    With Map.Tile(X, Y)
        ' Time to loop through our layers for this tile.
        For i = MapLayer.Ground To MapLayer.Fringe2
            ' Should we skip the tile?
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                Call RenderGraphic(Tex_TileSet(.Layer(i).Tileset), (X * PIC_X) + (MapViewTileOffSetX * PIC_X), (Y * PIC_Y) + (MapViewTileOffSetY * PIC_Y), PIC_X, PIC_Y, 0, 0, .Layer(i).X * PIC_X, .Layer(i).Y * PIC_Y)
            End If
        Next
    End With

End Sub

Public Sub RenderTileOutline(ByVal X As Long, ByVal Y As Long)
    If ShowMouse = False Then Exit Sub
    
    Call RenderGraphic(Tex_Outline, X * PIC_X, Y * PIC_Y, PIC_X, PIC_Y, 0, 0, 0, 0)
End Sub
