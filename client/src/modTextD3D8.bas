Attribute VB_Name = "modText"
Option Explicit

' Text color pointers
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan  As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const DarkBrown As Byte = 16
Public Const Orange As Byte = 17

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
    x As Long
    y As Long
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

Public Sub RenderText(ByRef UseFont As CustomFont, text As String, x As Long, y As Long, color As Long, Optional Alpha As Byte = 255)
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
                TempVA(0).x = x + Count
                TempVA(0).y = y + yOffset
                TempVA(1).x = TempVA(1).x + x + Count
                TempVA(1).y = TempVA(0).y
                TempVA(2).x = TempVA(0).x
                TempVA(2).y = TempVA(2).y + TempVA(0).y
                TempVA(3).x = TempVA(1).x
                TempVA(3).y = TempVA(2).y
                
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
    
    MainFont.TextureSize.x = 512
    MainFont.TextureSize.y = 512
    
    ' Init the fonts
    InitFonts
End Sub

Public Sub LoadFontHeader(ByRef UseFont As CustomFont, fileName As String)
Dim F As Long, i As Long
Dim Row As Single, u As Single, v As Single

    ' Load the header information
    F = FreeFile
    Open App.Path & GFX_PATH & "fonts\" & fileName For Binary As #F
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
            .Vertex(0).x = 0
            .Vertex(0).y = 0
            .Vertex(0).z = 0
            .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).tu = u + UseFont.ColFactor
            .Vertex(1).tv = v
            .Vertex(1).x = UseFont.HeaderInfo.CellWidth
            .Vertex(1).y = 0
            .Vertex(1).z = 0
            .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + UseFont.RowFactor
            .Vertex(2).x = 0
            .Vertex(2).y = UseFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).tu = u + UseFont.ColFactor
            .Vertex(3).tv = v + UseFont.RowFactor
            .Vertex(3).x = UseFont.HeaderInfo.CellWidth
            .Vertex(3).y = UseFont.HeaderInfo.CellHeight
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

Public Sub AddText(ByVal Msg As String, ByVal color As Integer)
Dim S As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    S = vbNewLine & Msg
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = QBColor(color)
    frmMain.txtChat.SelText = S
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function DrawMapAttributes()
    Dim x As Long
    Dim y As Long
    Dim tx As Long
    Dim ty As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optAttribs.Value Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    With Map.Tile(x, y)
                        tx = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText MainFont, "B", tx, ty, BrightRed, 200
                            Case TILE_TYPE_WARP
                                RenderText MainFont, "W", tx, ty, BrightBlue, 200
                            Case TILE_TYPE_ITEM
                                RenderText MainFont, "I", tx, ty, White, 200
                            Case TILE_TYPE_NPCAVOID
                                RenderText MainFont, "Na", tx, ty, White, 200
                            Case TILE_TYPE_KEY
                                RenderText MainFont, "K", tx, ty, White, 200
                            Case TILE_TYPE_KEYOPEN
                                RenderText MainFont, "O", tx, ty, White, 200
                            Case TILE_TYPE_RESOURCE
                                RenderText MainFont, "R", tx, ty, Green, 200
                            Case TILE_TYPE_DOOR
                                RenderText MainFont, "D", tx, ty, Brown, 200
                            Case TILE_TYPE_NPCSPAWN
                                RenderText MainFont, "Ns", tx, ty, Yellow, 200
                            Case TILE_TYPE_SHOP
                                RenderText MainFont, "Sh", tx, ty, BrightBlue, 200
                            Case TILE_TYPE_BANK
                                RenderText MainFont, "Ba", tx, ty, Blue, 200
                            Case TILE_TYPE_HEAL
                                RenderText MainFont, "H", tx, ty, BrightGreen, 200
                            Case TILE_TYPE_TRAP
                                RenderText MainFont, "T", tx, ty, BrightRed, 200
                            Case TILE_TYPE_SLIDE
                                RenderText MainFont, "Sl", tx, ty, BrightCyan, 200
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "DrawMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawActionMsg(ByVal Index As Long)
    Dim x As Long, y As Long, i As Long, Time As Long, color As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If ActionMsg(Index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            Time = 1500

            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            Time = 1500
        
            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If
            Next
            x = (frmMain.picScreen.Width \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            y = 425

    End Select
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    If GetTickCount < ActionMsg(Index).Created + Time Then
        color = ActionMsg(Index).color
        RenderText MainFont, ActionMsg(Index).message, x, y, color
    Else
        ClearActionMsg Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check access level
    If GetPlayerPK(Index) = NO Then

        Select Case GetPlayerAccess(Index)
            Case 0
                color = Orange
            Case 1
                color = White
            Case 2
                color = Cyan
            Case 3
                color = BrightGreen
            Case 4
                color = Yellow
        End Select

    Else
        color = BrightRed
    End If

    Name = Trim$(Player(Index).Name)
    ' calc pos
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - (GetTextWidth(MainFont, (Trim$(Name))) / 2)
    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - (D3DT_TEXTURE(Tex_Character(GetPlayerSprite(Index))).Height / 4) + 16
    End If

    ' Draw name
    RenderText MainFont, Name, TextX, TextY, color
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String
Dim npcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    npcNum = MapNpc(Index).num

    Select Case Npc(npcNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            color = BrightRed
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            color = Yellow
        Case NPC_BEHAVIOUR_GUARD
            color = Grey
        Case Else
            color = BrightGreen
    End Select

    Name = Trim$(Npc(npcNum).Name)
    TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).XOffset + (PIC_X \ 2) - (GetTextWidth(MainFont, (Trim$(Name))) / 2)
    If Npc(npcNum).Sprite < 1 Or Npc(npcNum).Sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - (D3DT_TEXTURE(Tex_Character(Npc(npcNum).Sprite)).Height / 4) + 16
    End If

    ' Draw name
    RenderText MainFont, Name, TextX, TextY, color
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
