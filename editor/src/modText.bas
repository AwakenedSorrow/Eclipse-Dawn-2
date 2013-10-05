Attribute VB_Name = "modText"
Option Explicit

Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Private Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type CustomFont
    HeaderInfo As VFH
    Texture As Direct3DTexture8
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
    TextureSize As POINTAPI
End Type

Public MainFont As CustomFont

Public Sub RenderText(ByRef UseFont As CustomFont, text As String, X As Long, Y As Long, color As Long, Optional Alpha As Byte = 255)
Dim TempVA(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim Count As Integer
Dim Ascii() As Byte
Dim i As Long, j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim srcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim yOffset As Single

    ' Set the color
    color = DX8Colour(color, Alpha)

    ' Check for valid text to render
    If LenB(text) = 0 Then Exit Sub
    
    ' Get the text into arrays (split by vbCrLf)
    TempStr = Split(text, vbCrLf)
    
    ' Set the temp color (or else the first character has no color)
    TempColor = color
    
    ' Set the texture
    D3DDevice8.SetTexture 0, UseFont.Texture
    CurrentTexture = -1
    
    ' Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0
            ' Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            
            ' Loop through the characters
            For j = 1 To Len(TempStr(i))
                ' Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4)
                
                ' Set up the verticies
                TempVA(0).X = X + Count
                TempVA(0).Y = Y + yOffset
                TempVA(1).X = TempVA(1).X + X + Count
                TempVA(1).Y = TempVA(0).Y
                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y
                
                ' Set the colors
                TempVA(0).color = TempColor
                TempVA(1).color = TempColor
                TempVA(2).color = TempColor
                TempVA(3).color = TempColor
                
                ' Draw the verticies
                Call D3DDevice8.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size)
                
                ' Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                ' Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = color
                End If
            Next
        End If
    Next
End Sub

Public Sub EngineInitFontTextures()

    ' Check if we have the device
    If D3DDevice8.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    ' silkscreen
    Set MainFont.Texture = D3DX8.CreateTextureFromFileEx(D3DDevice8, App.Path & GFX_PATH & "fonts\" & FONT_NAME & ".png", 512, 512, 0, 0, _
    D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, RGB(255, 0, 255), ByVal 0, ByVal 0)
    
    MainFont.TextureSize.X = 512
    MainFont.TextureSize.Y = 512
    
    ' Init the fonts
    InitFonts
End Sub

Public Sub LoadFontHeader(ByRef UseFont As CustomFont, FileName As String)
Dim F As Long, i As Long
Dim Row As Single, u As Single, v As Single

    ' Load the header information
    F = FreeFile
    Open App.Path & GFX_PATH & "fonts\" & FileName For Binary As #F
        Get #F, , UseFont.HeaderInfo
    Close #F
    
    ' Calculate some common values
    UseFont.CharHeight = UseFont.HeaderInfo.CellHeight - 4
    UseFont.RowPitch = UseFont.HeaderInfo.BitmapWidth \ UseFont.HeaderInfo.CellWidth
    UseFont.ColFactor = UseFont.HeaderInfo.CellWidth / UseFont.HeaderInfo.BitmapWidth
    UseFont.RowFactor = UseFont.HeaderInfo.CellHeight / UseFont.HeaderInfo.BitmapHeight
    
    ' Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For i = 0 To 255
        ' tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (i - UseFont.HeaderInfo.BaseCharOffset) \ UseFont.RowPitch
        u = ((i - UseFont.HeaderInfo.BaseCharOffset) - (Row * UseFont.RowPitch)) * UseFont.ColFactor
        v = Row * UseFont.RowFactor
        
        ' Set the verticies
        With UseFont.HeaderInfo.CharVA(i)
            .Vertex(0).color = D3DColorARGB(255, 0, 0, 0) ' Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).z = 0
            .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).tu = u + UseFont.ColFactor
            .Vertex(1).tv = v
            .Vertex(1).X = UseFont.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).z = 0
            .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + UseFont.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = UseFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).tu = u + UseFont.ColFactor
            .Vertex(3).tv = v + UseFont.RowFactor
            .Vertex(3).X = UseFont.HeaderInfo.CellWidth
            .Vertex(3).Y = UseFont.HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
    Next
End Sub

Public Sub InitFonts()
    LoadFontHeader MainFont, FONT_NAME & ".dat"
End Sub

Public Function DX8Colour(ByVal ColourNum As Long, ByVal Alpha As Long) As Long

    Select Case ColourNum
        Case 0 ' Black
            DX8Colour = D3DColorARGB(Alpha, 0, 0, 0)

        Case 1 ' Blue
            DX8Colour = D3DColorARGB(Alpha, 16, 104, 237)

        Case 2 ' Green
            DX8Colour = D3DColorARGB(Alpha, 119, 188, 84)

        Case 3 ' Cyan
            DX8Colour = D3DColorARGB(Alpha, 16, 224, 237)

        Case 4 ' Red
            DX8Colour = D3DColorARGB(Alpha, 201, 0, 0)

        Case 5 ' Magenta
            DX8Colour = D3DColorARGB(Alpha, 255, 0, 255)

        Case 6 ' Brown
            DX8Colour = D3DColorARGB(Alpha, 175, 149, 92)

        Case 7 ' Grey
            DX8Colour = D3DColorARGB(Alpha, 192, 192, 192)

        Case 8 ' DarkGrey
            DX8Colour = D3DColorARGB(Alpha, 128, 128, 128)

        Case 9 ' BrightBlue
            DX8Colour = D3DColorARGB(Alpha, 126, 182, 240)

        Case 10 ' BrightGreen
            DX8Colour = D3DColorARGB(Alpha, 0, 193, 22)

        Case 11 ' BrightCyan
            DX8Colour = D3DColorARGB(Alpha, 157, 242, 242)

        Case 12 ' BrightRed
            DX8Colour = D3DColorARGB(Alpha, 255, 0, 0)

        Case 13 ' Pink
            DX8Colour = D3DColorARGB(Alpha, 255, 118, 221)

        Case 14 ' Yellow
            DX8Colour = D3DColorARGB(Alpha, 255, 255, 0)

        Case 15 ' White
            DX8Colour = D3DColorARGB(Alpha, 255, 255, 255)

        Case 16 ' DarkBrown
            DX8Colour = D3DColorARGB(Alpha, 98, 84, 52)
        
        Case 17 ' Orange
            DX8Colour = D3DColorARGB(Alpha, 255, 180, 0)
    End Select
End Function

Public Function GetTextWidth(ByRef UseFont As CustomFont, ByVal text As String) As Integer
Dim i As Long

    ' Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    ' Loop through the text
    For i = 1 To Len(text)
        GetTextWidth = GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(text, i, 1)))
    Next
End Function

Public Sub DrawMapAttributes()
    Dim X As Long
    Dim Y As Long
    Dim tx As Long
    Dim ty As Long

    For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                If IsValidMapPoint(X, Y) Then
                With Map.Tile(X, Y)
                    tx = ((MapViewTileOffSetX + X) * 32) + 10
                    ty = ((MapViewTileOffSetY + Y) * 32) + 10
                        Select Case .Type
                            Case TileTypeBlocked
                                RenderText MainFont, "Bl", tx, ty, BrightRed, 200
                            Case TileTypeWarp
                                RenderText MainFont, "Wa", tx, ty, BrightBlue, 200
                            Case TileTypeItem
                                RenderText MainFont, "It", tx, ty, White, 200
                            Case TileTypeNPCAvoid
                                RenderText MainFont, "Na", tx, ty, White, 200
                            Case TileTypeKey
                                RenderText MainFont, "Ke", tx, ty, White, 200
                            Case TileTypeKeyOpen
                                RenderText MainFont, "Op", tx, ty, White, 200
                            Case TileTypeResource
                                RenderText MainFont, "Re", tx, ty, Green, 200
                            Case TileTypeDoor
                                RenderText MainFont, "Do", tx, ty, Brown, 200
                            Case TileTypeNPCSpawn
                                RenderText MainFont, "Ns", tx, ty, yellow, 200
                            Case TileTypeShop
                                RenderText MainFont, "Sh", tx, ty, BrightBlue, 200
                            Case TileTypeBank
                                RenderText MainFont, "Ba", tx, ty, Blue, 200
                            Case TileTypeHeal
                                RenderText MainFont, "He", tx, ty, brightgreen, 200
                            Case TileTypeTrap
                                RenderText MainFont, "Da", tx, ty, BrightRed, 200
                            Case TileTypeSlide
                                RenderText MainFont, "Sl", tx, ty, BrightCyan, 200
                            Case TileTypeScripted
                                RenderText MainFont, "Sc", tx, ty, Pink, 200
                        End Select
                    End With
                End If
            Next
        Next
End Sub


