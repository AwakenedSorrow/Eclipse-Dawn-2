Attribute VB_Name = "modRendering"
Public Sub Render_Game()
Dim x As Long, y As Long, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Some checks at the start of the sub, if the client's minimized or we're still loading a map
    ' We do not want the client to actually render the map, it's a waste of resources or might cause
    ' some very odd glitches.
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub
    
    ' Update the viewpoint of the Camera to make sure it sticks to the player on Scrolling Maps.
    ' Not quite sure why this isn't built into the player movement system, but we'll go with it
    ' for the time being.
    UpdateCamera
    
    ' Clear the device surface and open the scene up for rendering.
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Render the tiles that will be under the player, in this case Ground, Mask1 and Mask2.
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call RenderMapTile(x, y)
                End If
            Next
        Next
    End If
    
    ' Blood Decals. These need to be under the player and everything else for sure.
    ' Imagine that, floating blood?
    For i = 1 To MAX_BYTE
        Call RenderBlood(i)
    Next
    
    ' Render the Map Items, these are also under the players and such because they're on the ground
    ' and our lovely little players will walk over them to pick them up.
    If NumItems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call RenderMapItem(i)
            End If
        Next
    End If
    
    ' Y-Based Rendering time! Stuff that's "further" away from the front of the screen
    ' (Y-0 being the furthest and the highest being whatever)
    ' Will be rendered first, so it's behind everything else regardless of what it is.
    For y = 0 To Map.MaxY
        
        ' Player Characters
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If Player(i).y = y Then
                    Call RenderPlayer(i)
                End If
            End If
        Next
        
    ' The end of the Y-Based rendering loop!
    Next
    
    ' Render the tiles that will be above the player, in this case Fringe1 and Fringe 2.
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call RenderUpperMapTile(x, y)
                End If
            Next
        Next
    End If
    
    ' End the rendering scene and present it to the player.
    ' This makes sure we can actually SEE what we rendered onto the device above.
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "Render_Game", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderMapTile(ByVal x As Long, ByVal y As Long)
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With Map.Tile(x, y)
        ' Time to loop through our layers for this tile.
        For i = MapLayer.Ground To MapLayer.Mask2
            ' Should we skip the tile?
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).x > 0 Or .Layer(i).y > 0) Then
                Call RenderGraphic(Tex_TileSet(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), PIC_X, PIC_Y, 0, 0, .Layer(i).x * PIC_X, .Layer(i).y * PIC_Y)
            End If
        Next
    End With
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderMapTile", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderUpperMapTile(ByVal x As Long, ByVal y As Long)
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With Map.Tile(x, y)
        ' Time to loop through our layers for this tile.
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' Should we skip the tile?
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).x > 0 Or .Layer(i).y > 0) Then
                Call RenderGraphic(Tex_TileSet(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), PIC_X, PIC_Y, 0, 0, .Layer(i).x * PIC_X, .Layer(i).y * PIC_Y)
            End If
        Next
    End With
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderUpperMapTile", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderPlayer(ByVal Index As Long)
Dim SpriteFrame As Byte, i As Long, x As Long, y As Long
Dim Sprite As Long, SpriteDir As Long
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Get the sprite we're using for this player.
    Sprite = GetPlayerSprite(Index)
    
    ' Check if the sprite's valid, if not exit the sub so we don't cause any issues.
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' Check if the texture for the sprite is loaded, if not load it.
    ' We need it loaded for later, since we're using the height/width to calculate certain things.
    If D3DT_TEXTURE(Tex_Character(Sprite)).Loaded = False Then
        LoadTexture Tex_Character(Sprite)
    End If

    ' Retrieve the weapon speed if the player has one equipped.
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    ' Reset the movement frame if the player;s at the end of beginning of a movement.
    If Player(Index).Step = 3 Then
        SpriteFrame = 0
    ElseIf Player(Index).Step = 1 Then
        SpriteFrame = 2
    End If
    
    ' Check if we should be using the attack Frame
    If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If Player(Index).Attacking = 1 Then
            SpriteFrame = 3
        End If
    Else
        ' Apparently not, so we'll be using the regular movement animations!
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).yOffset > 8) Then SpriteFrame = Player(Index).Step
            Case DIR_DOWN
                If (Player(Index).yOffset < -8) Then SpriteFrame = Player(Index).Step
            Case DIR_LEFT
                If (Player(Index).XOffset > 8) Then SpriteFrame = Player(Index).Step
            Case DIR_RIGHT
                If (Player(Index).XOffset < -8) Then SpriteFrame = Player(Index).Step
        End Select
    End If

    ' Are we allowed to still attack the next frame? Probably not!
    With Player(Index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Decide which direction our little sprite should be facing, 0 being the top row and 3 being the bottom
    ' On your spritesheet. Please change these if your spritesheet's in a different order from the standard
    ' RMXP format.
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            SpriteDir = 3
        Case DIR_RIGHT
            SpriteDir = 2
        Case DIR_DOWN
            SpriteDir = 0
        Case DIR_LEFT
            SpriteDir = 1
    End Select

    ' Calculate the X position to render the player sprite at. Along with offset, of course!
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((D3DT_TEXTURE(Tex_Character(Sprite)).Width / 4 - 32) / 2)

    ' Time to work on the Y position.
    ' But first, let's check if the sprite is more than 32 pixels high.
    If (D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4) - 32)
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
    End If

    ' We're done here, let's render the sprite!
    Call RenderSprite(Sprite, x, y, SpriteFrame, SpriteDir)
    
    ' Let's not do paperdolling just yet shall we? Would like to get the rest to work first.
    'For i = 1 To UBound(PaperdollOrder)
        'If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            'If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                'Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
            'End If
        'End If
    'Next
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderPlayer", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderSprite(ByVal Sprite As Long, ByVal x2 As Long, y2 As Long, ByVal SpriteFrame As Long, ByVal SpriteDir As Long)
Dim x As Long
Dim y As Long
Dim Width As Long
Dim Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Is the sprite valid? If not, exit the sub to prevent issues from occuring.
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    ' Convert the provided values to values we can use on the map.
    x = ConvertMapX(x2)
    y = ConvertMapY(y2)
    
    ' Pre-Calculate these values, it makes the render line look a lot cleaner.
    Width = D3DT_TEXTURE(Tex_Character(Sprite)).Width / 4
    Height = D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4
    
    ' Render the sprite itself! Please do -NOT- touch this line unless you know what you're doing.
    Call RenderGraphic(Tex_Character(Sprite), x, y, Width, Height, 0, 0, SpriteFrame * Width, SpriteDir * Height)
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderBlood(ByVal Index As Long)
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With Blood(Index)
        ' Before we continue, we may want to see if the Blood Decal is still "valid".
        ' If not, we just exit the sub and continue on to the next one.
        If .Timer + 20000 < GetTickCount Then Exit Sub
        
            ' Right, a little addition of my own. The longer blood's been on the map the less visible it will become.
            ' It will fade a bit every two seconds. It's nothing fancy but I prefer it this way. :)
            If .LastTimer + 2000 < GetTickCount Then
                .Alpha = .Alpha - 25
                .LastTimer = .LastTimer + 2000
            End If
            
            ' Now that we've got all that sorted, let's get to rendering this bugger!
            Call RenderGraphic(Tex_Blood, ConvertMapX(.x * PIC_X), ConvertMapY(.y * PIC_Y), PIC_X, PIC_Y, 0, 0, (.Sprite - 1) * PIC_X, 0, 0, 0, 0, .Alpha)
   
    End With
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderBlood", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderMapItem(ByVal Index As Long)
Dim AnimFrame

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if the itemdrop belongs to anyone, if it does and it is not us we skip rendering it.
    ' Wouldn't want to be able to steal someone else's items right?
    If MapItem(Index).playerName <> vbNullString Then
        If MapItem(Index).playerName <> Trim$(GetPlayerName(MyIndex)) Then Exit Sub
    End If
    
    With Item(MapItem(Index).num)
        
        ' Let's make sure we're using a valid picture for the item, if not we skip rendering it to avoid issues.
        If .Pic < 1 Or .Pic > NumItems Then Exit Sub
        
        ' Make sure the texture is loaded.
        ' We need it to see if it is Animated.
        If D3DT_TEXTURE(Tex_Item(.Pic)).Loaded = False Then
            Call LoadTexture(Tex_Item(.Pic))
        End If
        
        ' Check if the Texture has multiple frames on it for Animation.
        ' If it is, we'll need to use the current animation frame to display on the map, if it isn't then.. Well, it won't matter much then.
        If D3DT_TEXTURE(Tex_Item(.Pic)).Width > 64 Then
            AnimFrame = MapItem(Index).Frame
        Else
            AnimFrame = 0
        End If
        
        ' We've done all the fancy stuff, now let's get to rendering this item!
        Call RenderGraphic(Tex_Item(.Pic), ConvertMapX(MapItem(Index).x * PIC_X), ConvertMapY(MapItem(Index).y * PIC_Y), PIC_X, PIC_Y, 0, 0, AnimFrame * PIC_X, 0)
        
    End With

' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderMapItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
