Attribute VB_Name = "modRendering"
Option Explicit



Public Sub RenderGraphics()
    ' Render Map Components
    MapEditor_DrawTileSet
End Sub

Public Function MapEditor_DrawTileSet()
Dim Tileset

    Tileset = frmEditor.cmbTileSet.ListIndex + 1
    
    If D3DT_TEXTURE(Tex_TileSet(Tileset)).Height > frmEditor.picTileSelect.ScaleHeight Then
            frmEditor.scrlTileSelect.max = (D3DT_TEXTURE(Tex_TileSet(Tileset)).Height \ 16) - (frmEditor.picTileSelect.ScaleHeight \ 16)
    End If
    
End Function

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
