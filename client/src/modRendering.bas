Attribute VB_Name = "modRendering"
Public Sub Render_Game()
Dim x As Long, y As Long, i As Long
Dim rec As RECT
Dim srcRect As D3DRECT

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
            For y = TileView.Top To TileView.bottom
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
    
    ' Render the BOTTOM layers of animations, this is to make sure that the lower part is always below
    ' the player, your target or whatever really.
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                RenderAnimation i, 0
            End If
        Next
    End If
    
    ' Y-Based Rendering time! Stuff that's "further" away from the front of the screen
    ' (Y-0 being the furthest and the highest being whatever)
    ' Will be rendered first, so it's behind everything else regardless of what it is.
    For y = 0 To Map.MaxY
        
        ' Check if we have any sprites loaded, if so we can start rendering players and NPCs!
        If NumCharacters > 0 Then
            ' Player Characters
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).y = y Then
                        Call RenderPlayer(i)
                    End If
                End If
            Next
            
            ' Non-Player Characters
            For i = 1 To Npc_HighIndex
                If MapNpc(i).y = y Then
                    Call RenderNPC(i)
                End If
            Next
        End If
        
        ' Time to start cracking on rendering Resources!
        If NumResources > 0 And Resources_Init And Resource_Index > 0 Then
            For i = 1 To Resource_Index
                If MapResource(i).y = y Then
                    Call RenderMapResource(i)
                End If
            Next
        End If
        
    ' The end of the Y-Based rendering loop!
    Next
    
    ' Render the TOP layers of animations, this is to make sure that the upper part is always above
    ' the player, your target or whatever really.
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                RenderAnimation i, 1
            End If
        Next
    End If
    
    ' Render the tiles that will be above the player, in this case Fringe1 and Fringe 2.
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call RenderUpperMapTile(x, y)
                End If
            Next
        Next
    End If
    
    ' Some Map Editor specific stuff, such as Directional Blocking and the Mouse Tile Outline.
    ' Fairly important stuff if you want to get it all to work. :)
    If InMapEditor Then
        ' Directional Blocking
        If frmEditor_Map.optBlock.Value = True Then
            For x = TileView.Left To TileView.Right
                For y = TileView.Top To TileView.bottom
                    If IsValidMapPoint(x, y) Then
                        Call RenderDirBlock(x, y)
                    End If
                Next
            Next
        End If
        ' Mouse Cursor Tile Outline.
        Call RenderTileOutline
    End If
    
    ' Render Health, Mana and Cast Bars.
    Call RenderBars
    
    ' Render all the hover and target textures.
    Call RenderHoverAndTarget
    
    ' We've got all the graphics we need rendered, but no test to understand what's going on yet!
    ' Should probably get that sorted out below.
            
    ' Displays the FPS.
    If BFPS Then
        RenderText MainFont, "FPS: " & CStr(GameFPS), 2, 39, Yellow
    End If
            
    ' draw cursor, player X and Y locations
    If BLoc Then
        RenderText MainFont, Trim$("cur x: " & CurX & " y: " & CurY), 2, 1, Yellow
        RenderText MainFont, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 2, 15, Yellow
        RenderText MainFont, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 2, 27, Yellow
    End If
    
    ' Draw the on-screen action messages.
    For i = 1 To Action_HighIndex
        Call DrawActionMsg(i)
    Next i
    
    ' Render Player Names
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(i)
        End If
    Next
    
    ' Render NPC Names
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            Call DrawNpcName(i)
        End If
    Next
    
    ' Draw map name
    RenderText MainFont, Map.Name, DrawMapNameX, DrawMapNameY, Yellow
            
    ' If we're in the map editor, draw the attributes.
    If InMapEditor Then Call DrawMapAttributes
    
    ' End the rendering scene and present it to the player.
    ' This makes sure we can actually SEE what we rendered onto the device above.
    With srcRect
        .X1 = 0
        .x2 = frmMain.picScreen.ScaleWidth
        .Y1 = 0
        .y2 = frmMain.picScreen.ScaleHeight
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, ByVal 0, 0, ByVal 0)
    
    ' Now that we've done all the gamescreen stuff, we can start rendering the graphics on our GDI components!
    Call DrawGDI
    
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
    
    ' Paperdoll time!
    ' This will render any equipped items that the player has on which have been assigned a
    ' paperdoll.
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call RenderPaperdoll(x, y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, SpriteFrame, SpriteDir)
            End If
        End If
    Next
    
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
        If .timer + 60000 < GetTickCount Then Exit Sub
        
            ' Right, a little addition of my own. The longer blood's been on the map the less visible it will become.
            ' It will fade a bit every 1.5 seconds. It's nothing fancy but I prefer it this way. :)
            If .LastTimer + 1500 < GetTickCount Then
                If .Alpha >= 7 Then
                    .Alpha = .Alpha - 7
                Else
                    .Alpha = 0
                End If
                .LastTimer = .LastTimer + 1500
            End If
            
            ' Now that we've got all that sorted, let's get to rendering this bugger!
            Call RenderGraphic(Tex_Blood, ConvertMapX(.x * PIC_X), ConvertMapY(.y * PIC_Y), PIC_X, PIC_Y, 0, 0, (.Sprite - 1) * PIC_X, 0, 255, 255, 255, .Alpha)
   
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

Sub RenderAnimation(ByVal Index As Long, ByVal Layer As Byte)
Dim Sprite As Long
Dim i As Long
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim x As Long, y As Long
Dim lockindex As Long
Dim AnimFrame As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' If the animation instance doesn't hold anything for whatever reason, clear it and exit the sub.
    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    ' Retrieve the texture we're using for this Animation and check if it is a valid number.
    ' If it isn't, exit out of this sub and continue on to the next.
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    ' Retrieve the framecount set in the editor.
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    ' Set the timer.
    AnimationTimer(Sprite) = GetTickCount + SurfaceTimerMax
    
    ' Get and set the Height and Width of the sprite frame we'll be using.
    Width = D3DT_TEXTURE(Tex_Animation(Sprite)).Width / FrameCount
    Height = D3DT_TEXTURE(Tex_Animation(Sprite)).Height
    
    ' Set the Animation Frame we'll be using.
    ' Note that unlike most other render subs, this frame already includes the full location of the texture.
    ' This does not need to be multiplied again further on to get the proper frame texture.
    AnimFrame = (AnimInstance(Index).FrameIndex(Layer) - 1) * Width

    ' Let's change the X/Y offset if the Animation is locked onto a target so it actually displays on
    ' top of it.
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' Locked On to a Player.
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    x = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).XOffset
                    y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    x = (MapNpc(lockindex).x * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).XOffset
                    y = (MapNpc(lockindex).y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).yOffset
                Else
                    ' The NPC isn't alive anymore, sadly with the way the system works this means we need to destroy the animation as well.
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' The NPC isn't alive anymore, sadly with the way the system works this means we need to destroy the animation as well.
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        x = (AnimInstance(Index).x * 32) + 16 - (Width / 2)
        y = (AnimInstance(Index).y * 32) + 16 - (Height / 2)
    End If
    
    ' Convert these values beforehand. Saves some space down there.
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    ' Render the actual texture. Should be some fancy animation now!
    Call RenderGraphic(Tex_Animation(Sprite), x, y, Width, Height, 0, 0, AnimFrame, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderAnimation", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderNPC(ByVal Index As Long)
Dim SpriteAnim As Byte, i As Long, x As Long, y As Long, Sprite As Long, SpriteDir As Long
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check if we actually set an NPC here, if not exit the sub and carry on to the next one.
    If MapNpc(Index).num = 0 Then Exit Sub
    
    ' Retrieve the sprite we'll be using, and check if it is valid.
    ' If not, exit the sub and carry on.
    Sprite = Npc(MapNpc(Index).num).Sprite
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    attackspeed = 1000

    ' Reset the animation frame
    SpriteAnim = 0
    ' Check for attacking animation
    If MapNpc(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If MapNpc(Index).Attacking = 1 Then
            SpriteAnim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(Index).Dir
            Case DIR_UP
                If (MapNpc(Index).yOffset > 8) Then SpriteAnim = MapNpc(Index).Step
            Case DIR_DOWN
                If (MapNpc(Index).yOffset < -8) Then SpriteAnim = MapNpc(Index).Step
            Case DIR_LEFT
                If (MapNpc(Index).XOffset > 8) Then SpriteAnim = MapNpc(Index).Step
            Case DIR_RIGHT
                If (MapNpc(Index).XOffset < -8) Then SpriteAnim = MapNpc(Index).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(Index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the Sprite Direction
    Select Case MapNpc(Index).Dir
        Case DIR_UP
            SpriteDir = 3
        Case DIR_RIGHT
            SpriteDir = 2
        Case DIR_DOWN
            SpriteDir = 0
        Case DIR_LEFT
            SpriteDir = 1
    End Select

    ' Calculate the X
    x = MapNpc(Index).x * PIC_X + MapNpc(Index).XOffset - ((D3DT_TEXTURE(Tex_Character(Sprite)).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = MapNpc(Index).y * PIC_Y + MapNpc(Index).yOffset - ((D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4) - 32)
    Else
        ' Proceed as normal
        y = MapNpc(Index).y * PIC_Y + MapNpc(Index).yOffset
    End If

    Call RenderSprite(Sprite, x, y, SpriteAnim, SpriteDir)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderNpc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderMapResource(ByVal Resource_num As Long)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).x > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).y > Map.MaxY Then Exit Sub
        
    ' Get the Original Resource.
    Resource_master = Map.Tile(MapResource(Resource_num).x, MapResource(Resource_num).y).Data1
    
    ' Check if it's a valid resource, if it is not then we skip rendering it and carry on.
    ' Same thing with the image.
    If Resource_master < 1 Or Resource_master > MAX_RESOURCES Then Exit Sub
    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    
    ' Get the Resource state (e.g. Full or Empty)
    Resource_state = MapResource(Resource_num).ResourceState
    
    ' If we're in the map editor we're reverting all Resources to their Empty state to make it easier to see the entire map.
    ' If this is now the case however, we'll need to use the image that represents the current state of it.
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    Else
        If Resource_state = 0 Then ' Full
            Resource_sprite = Resource(Resource_master).ResourceImage
        ElseIf Resource_state = 1 Then ' Empty
            Resource_sprite = Resource(Resource_master).ExhaustedImage
        End If
    End If
    
    ' Set base x + y, then the offset due to size
    x = (MapResource(Resource_num).x * PIC_X) - (D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Width / 2) + 16
    y = (MapResource(Resource_num).y * PIC_Y) - D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Height + 32
    
    ' render it
    Call RenderResource(Resource_sprite, x, y, D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Width, D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Height)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderMapResource", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderResource(ByVal Resource As Long, ByVal DX As Long, ByVal DY As Long, ByVal Width As Long, ByVal Height As Long)
Dim x As Long
Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check if the resource is a valid number. If not, exit out.
    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ' Convert the provided values to something we can use on the map.
    x = ConvertMapX(DX)
    y = ConvertMapY(DY)
    
    ' Render the actual resource on the map!
    Call RenderGraphic(Tex_Resource(Resource), x, y, Width, Height, 0, 0, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderResource", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderDirBlock(ByVal x As Long, ByVal y As Long)
Dim i As Long, Left As Long, Top As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Render the Grid Texture.
    Call RenderGraphic(Tex_DirBlock, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), PIC_X, PIC_Y, 0, 0, 0, 24)
    
    ' render dir blobs
    For i = 1 To 4
        Left = (i - 1) * 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(x, y).DirBlock, CByte(i)) Then
            Top = 8
        Else
            Top = 16
        End If
        'render the actual thing!
        Call RenderGraphic(Tex_DirBlock, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), 8, 8, 0, 0, Left, Top)
        
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderDirBlock", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub RenderTileOutline()
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' If we're editing Directional Blocks, disable this outline. It just looks silly.
    If frmEditor_Map.optBlock.Value Then Exit Sub
    
    ' Render the outline to the screen!
    Call RenderGraphic(Tex_Outline, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), PIC_X, PIC_Y, 0, 0, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTileOutline", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim Top As Long, Right As Long
Dim barWidth As Long
Dim i As Long, npcNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' dynamic bar calculations
    sWidth = D3DT_TEXTURE(Tex_Bars).Width
    sHeight = D3DT_TEXTURE(Tex_Bars).Height / 4
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(npcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).x * PIC_X + MapNpc(i).XOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).y * PIC_Y + MapNpc(i).yOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (Npc(npcNum).HP / sWidth)) * sWidth
                
                ' draw bar background
                Top = sHeight * 1 ' HP bar background
                Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sWidth, sHeight, 0, 0, 0, Top)
                
                ' draw the content of the bar.
                Top = 0 ' HP bar
                Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), barWidth, sHeight, 0, 0, 0, Top)
            End If
        End If
    Next
    
    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 35 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' draw bar background
            Top = sHeight * 3 ' Spell bar background
            Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sWidth, sHeight, 0, 0, 0, Top)
            
            ' draw the bar proper
            Top = sHeight * 2 ' Spell bar
            Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), barWidth, sHeight, 0, 0, 0, Top)
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).yOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' draw bar background
        Top = sHeight * 1 ' HP bar background
        Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sWidth, sHeight, 0, 0, 0, Top)
       
        ' draw the bar proper
        Top = 0 ' HP bar
        Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), barWidth, sHeight, 0, 0, 0, Top)
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).XOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).yOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    Top = sHeight * 1 ' HP bar background
                    Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sWidth, sHeight, 0, 0, 0, Top)
                    
                    ' draw the bar's content.
                    Top = 0 ' HP bar
                    Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), barWidth, sHeight, 0, 0, 0, Top)
                End If
            End If
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderBars", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderHoverAndTarget()
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Handle the rendering of the target texture if one exists.
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then ' If the target is a player.
            Call RenderTarget((Player(myTarget).x * 32) + Player(myTarget).XOffset, (Player(myTarget).y * 32) + Player(myTarget).yOffset)
        ElseIf myTargetType = TARGET_TYPE_NPC Then ' If the target is an NPC.
            Call RenderTarget((MapNpc(myTarget).x * 32) + MapNpc(myTarget).XOffset, (MapNpc(myTarget).y * 32) + MapNpc(myTarget).yOffset)
        End If
    End If
    
    ' Now it's time to figure out if we're mousing over a Player or NPC on the map.
    ' If we are, we'll render a hover texture over their characters so we know it's a valid character we can target to cast spells on.
    For i = 1 To Player_HighIndex ' Players
        If IsPlaying(i) Then
            If Player(i).Map = Player(MyIndex).Map Then
                If CurX = Player(i).x And CurY = Player(i).y Then ' Is our cursor over something?
                    If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                        'We're already targetting this player, so no point in rendering this as well.
                    Else
                        Call RenderHover(TARGET_TYPE_PLAYER, i, (Player(i).x * 32) + Player(i).XOffset, (Player(i).y * 32) + Player(i).yOffset)
                    End If
                End If
            End If
        End If
    Next
    For i = 1 To Npc_HighIndex 'NPCs
        If MapNpc(i).num > 0 Then
            If CurX = MapNpc(i).x And CurY = MapNpc(i).y Then ' Is our cursor over something?
                If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                    'We're already targetting this NPC, so no point in rendering this as well.
                Else
                    Call RenderHover(TARGET_TYPE_NPC, i, (MapNpc(i).x * 32) + MapNpc(i).XOffset, (MapNpc(i).y * 32) + MapNpc(i).yOffset)
                End If
            End If
        End If
    Next
    
' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderHoverAndTarget", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderTarget(ByVal x As Long, ByVal y As Long)
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Let's do some magic and figure out where we'll need to render this texture!
    Width = D3DT_TEXTURE(Tex_Target).Width / 2
    Height = D3DT_TEXTURE(Tex_Target).Height
    
    ' Center it on the Target.
    x = x - ((Width - 32) / 2)
    y = y - (Height / 2)
    
    ' And convert it to be useful on our map!
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    ' Now render the texture to the screen.
    Call RenderGraphic(Tex_Target, x, y, Width, Height, 0, 0, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTarget", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RenderHover(ByVal tType As Long, ByVal target As Long, ByVal x As Long, ByVal y As Long)
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Let's do some magic and figure out where we'll need to render this texture!
    Width = D3DT_TEXTURE(Tex_Target).Width / 2
    Height = D3DT_TEXTURE(Tex_Target).Height
    
    ' Center it on the Target.
    x = x - ((Width - 32) / 2)
    y = y - (Height / 2)
    
    ' And convert it to be useful on our map!
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    ' Now render the texture to the screen.
    Call RenderGraphic(Tex_Target, x, y, Width, Height, 0, 0, Width, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderHover", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal x As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
 
    ConvertMapX = x - (TileView.Left * PIC_X) - Camera.Left
   
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
 
Public Function ConvertMapY(ByVal y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
 
    ConvertMapY = y - (TileView.Top * PIC_Y) - Camera.Top
   
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub DrawInventory()
Dim i As Long, x As Long, y As Long, itemnum As Long, itempic As Long
Dim Amount As Long
Dim Top As Long, Left As Long
Dim colour As Long
Dim tmpItem As Long, amountModifier As Long
Dim srcRect As D3DRECT, destRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' If we're not in-game we should probably not be here to begin with.
    ' So let's exit out before things go awry!
    If Not InGame Then Exit Sub

    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene

    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            
            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For x = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(x).num)
                    If TradeYourOffer(x).num = i Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(x).Value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(x).Value
                            End If
                        End If
                    End If
                Next
            End If

            If itempic > 0 And itempic <= NumItems Then
                If D3DT_TEXTURE(Tex_Item(itempic)).Width <= 64 Then ' This checks if it's an animated item, if it is we'll need to render it elsewhere.
                    
                    ' Calculate where we need to render the item.
                    Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    
                    ' Render the item Icon.
                    Call RenderGraphic(Tex_Item(itempic), Left, Top, PIC_X, PIC_Y, 0, 0, 32, 0)
                    
                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        y = Top + 22
                        x = Left - 4
                        
                        Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            colour = White
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            colour = Yellow
                        ElseIf Amount > 10000000 Then
                            colour = BrightGreen
                        End If
                        
                        Call RenderText(MainFont, Format$(ConvertCurrency(Str(Amount)), "#,###,###,###"), x, y, colour)

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If
NextLoop:
    Next
    
    'update animated items
    DrawAnimatedInvItems
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .x2 = frmMain.picInventory.Width
        .Y1 = 28
        .y2 = frmMain.picInventory.Height + .Y1
    End With
    
    With destRect
        .X1 = 0
        .x2 = frmMain.picInventory.Width
        .Y1 = 32
        .y2 = frmMain.picInventory.Height + .Y1
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picInventory.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventory", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawGDI()
    'Cycle Through in-game stuff before cycling through editors
    If frmMenu.Visible Then
        'If frmMenu.picCharacter.Visible Then NewCharacterDrawSprite
    End If
    
    If frmMain.Visible Then
        If frmMain.picTempInv.Visible Then DrawDraggedItem frmMain.picTempInv.Left, frmMain.picTempInv.Top
        If frmMain.picTempSpell.Visible Then DrawDraggedSpell frmMain.picTempSpell.Left, frmMain.picTempSpell.Top
        If frmMain.picSpellDesc.Visible Then DrawSpellDesc LastSpellDesc
        If frmMain.picItemDesc.Visible Then DrawItemDesc LastItemDesc
        If frmMain.picHotbar.Visible Then DrawHotbar
        If frmMain.picInventory.Visible Then DrawInventory
        If frmMain.picSpells.Visible Then DrawPlayerSpells
        'If frmMain.picCharacter.Visible Then DrawFace: DrawEquipment
        'If frmMain.picShop.Visible Then DrawShop
        'If frmMain.picTempBank.Visible Then DrawBankItem frmMain.picTempBank.Left, frmMain.picTempBank.Top
        'If frmMain.picBank.Visible Then DrawBank
        'If frmMain.picTrade.Visible Then DrawTrade
    End If
    
    
    If frmEditor_Animation.Visible Then
        'EditorAnim_DrawAnim
    End If
    
    If frmEditor_Item.Visible Then
        'EditorItem_DrawItem
        'EditorItem_DrawPaperdoll
    End If
    
    If frmEditor_Map.Visible Then
        'EditorMap_DrawTileset
        'If frmEditor_Map.fraMapItem.Visible Then EditorMap_DrawMapItem
        'If frmEditor_Map.fraMapKey.Visible Then EditorMap_DrawKey
    End If
    
    If frmEditor_NPC.Visible Then
        'EditorNpc_DrawSprite
    End If
    
    If frmEditor_Resource.Visible Then
        'EditorResource_DrawSprite
    End If
    
    If frmEditor_Spell.Visible Then
        'EditorSpell_DrawIcon
    End If

End Sub

Public Sub DrawAnimatedInvItems()
Dim i As Long, colour As Long
Dim itemnum As Long, itempic As Long
Dim x As Long, y As Long
Dim MaxFrames As Byte
Dim Amount As Long, AnimLeft As Long, Top As Long, Left As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check if we're in the game or not, if we aren't we really shouldn't be here.
    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS

        If MapItem(i).num > 0 Then
            itempic = Item(MapItem(i).num).Pic

            If itempic < 1 Or itempic > NumItems Then Exit Sub
            MaxFrames = (D3DT_TEXTURE(Tex_Item(itempic)).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < MaxFrames - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 1
            End If
        End If

    Next

    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= NumItems Then
                If DDSD_Item(itempic).lWidth > 64 Then
                    MaxFrames = (D3DT_TEXTURE(Tex_Item(itempic)).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(i) < MaxFrames - 1 Then
                        InvItemFrame(i) = InvItemFrame(i) + 1
                    Else
                        InvItemFrame(i) = 1
                    End If

                    
                    AnimLeft = (D3DT_TEXTURE(Tex_Item(itempic)).Width / 2) + (InvItemFrame(i) * 32) ' middle to get the start of inv gfx, then +32 for each frame
                        
                    Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))

                    ' We'll now re-blt the item, and place the currency value over it again :P
                    Call RenderGraphic(Tex_Item(itempic), Left, Top, PIC_X, PIC_Y, 0, 0, AnimLeft, 0)

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        y = Top + 22
                        x = Left - 4
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            colour = White
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            colour = Yellow
                        ElseIf Amount > 10000000 Then
                            colour = BrightGreen
                        End If
                        
                        Call RenderText(MainFont, Format$(ConvertCurrency(Str(Amount)), "#,###,###,###"), x, y, colour)

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If

    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawAnimatedInvItems", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderPaperdoll(ByVal x2 As Long, ByVal y2 As Long, ByVal Sprite As Long, ByVal SpriteFrame As Long, ByVal SpriteDir As Long)
Dim Top As Long, Left As Long
Dim x As Long, y As Long
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if the sprite is a valid one. If it isn't then exit out of the sub.
    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    Top = SpriteDir * (D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Height / 4)
    Left = SpriteFrame * (D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Width / 4)
    
    ' Caclculate a few things we might need.
    x = ConvertMapX(x2)
    y = ConvertMapY(y2)
    Width = (D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Width / 4)
    Height = (D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Height / 4)
    
    ' Rendering time!
    Call RenderGraphic(Tex_Paperdoll(Sprite), x, y, Width, Height, 0, 0, Left, Top)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderPaperdoll", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawItemDesc(ByVal itemnum As Long)
Dim itempic As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim Top As Long, Left As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure the item number is valid before we continue, if it isn't we'll simply skip rendering the icon.
    If itemnum > 0 And itemnum <= MAX_ITEMS Then
        
        ' Retrieve the item image, and check if it is valid.
        itempic = Item(itemnum).Pic
        If itempic < 1 Or itempic > NumItems Then Exit Sub
        
        ' Let's open clear ourselves a nice clean slate to render on shall we?
        Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
        Call D3DDevice8.BeginScene
        
        ' Calculate what image we need to use to render here.
        ' Note that the tooltips do not support animations.
        ' It simply shows the first icon of the inventory row.
        Top = 0
        Left = D3DT_TEXTURE(Tex_Item(itempic)).Width / 2
        
        ' Render it on the surface.
        Call RenderGraphic(Tex_Item(itempic), 0, 0, PIC_X, PIC_Y, 0, 0, Left, Top)
        
        ' We're done for now, so we can close the lovely little rendering device and present it to our user!
        ' Of course, we also need to do a few calculations to make sure it appears where it should.
        With srcRect
            .X1 = 0
            .x2 = PIC_X
            .Y1 = 0
            .y2 = PIC_Y
        End With
    
        With destRect
            .X1 = 0
            .x2 = frmMain.picItemDescPic.Width
            .Y1 = 0
            .y2 = frmMain.picItemDescPic.Height
        End With
    
        Call D3DDevice8.EndScene
        Call D3DDevice8.Present(srcRect, destRect, frmMain.picItemDescPic.hWnd, ByVal 0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawItemDesc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDraggedItem(ByVal x As Long, ByVal y As Long)
Dim Top As Long, Left As Long
Dim itemnum As Long, itempic As Long
Dim srcRect As D3DRECT, destRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Retrieve the item number we're trying to drag around.
    itemnum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)
    
    ' If the item number is valid then make sure we do something with it, wouldn't like to have an invisible icon right?
    If itemnum > 0 And itemnum <= MAX_ITEMS Then
    
        ' Retrieve the item texture and make sure it is valid before we continue.
        itempic = Item(itemnum).Pic
        If itempic < 1 Or itempic > NumItems Then Exit Sub
        
        ' Let's open clear ourselves a nice clean slate to render on shall we?
        Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
        Call D3DDevice8.BeginScene
        
        ' Calculate what image we need to grab from the texture.
        Top = 0
        Left = D3DT_TEXTURE(Tex_Item(itempic)).Width / 2
        
        ' Render the texture to the screen, we're using a 2pixel offset to make sure it's centered and doesn't clip
        ' with the picturebox. It's an original design choice in Mirage4, lord knows why.
        Call RenderGraphic(Tex_Item(itempic), 2, 2, PIC_X, PIC_Y, 0, 0, Left, Top)
        
        ' We're done for now, so we can close the lovely little rendering device and present it to our user!
        ' Of course, we also need to do a few calculations to make sure it appears where it should.
        With srcRect
            .X1 = 2
            .x2 = .X1 + PIC_X
            .Y1 = 2
            .y2 = .Y1 + PIC_Y
        End With
    
        With destRect
            .X1 = 0
            .x2 = frmMain.picTempInv.Width
            .Y1 = 0
            .y2 = frmMain.picTempInv.Height
        End With
    
        Call D3DDevice8.EndScene
        Call D3DDevice8.Present(srcRect, destRect, frmMain.picTempInv.hWnd, ByVal 0)

        With frmMain.picTempInv
            .Top = y
            .Left = x
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDraggedItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDraggedSpell(ByVal x As Long, ByVal y As Long)
Dim Top As Long, Left As Long
Dim spellnum As Long, spellpic As Long
Dim srcRect As D3DRECT, destRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Retrieve the item number we're trying to drag around.
    spellnum = PlayerSpells(DragSpell)
    
    ' If the spell number is valid then make sure we do something with it, wouldn't like to have an invisible icon right?
    If spellnum > 0 And spellnum <= MAX_SPELLS Then
    
        ' Retrieve the spell texture and make sure it is valid before we continue.
        spellpic = Spell(spellnum).Icon
        If spellpic < 1 Or spellpic > NumSpellIcons Then Exit Sub
        
        ' Let's open clear ourselves a nice clean slate to render on shall we?
        Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
        Call D3DDevice8.BeginScene
        
        ' Calculate what image we need to grab from the texture.
        Top = 0
        Left = 0
        
        ' Render the texture to the screen, we're using a 2pixel offset to make sure it's centered and doesn't clip
        ' with the picturebox. It's an original design choice in Mirage4, lord knows why.
        Call RenderGraphic(Tex_SpellIcon(spellpic), 2, 2, PIC_X, PIC_Y, 0, 0, Left, Top)
        
        ' We're done for now, so we can close the lovely little rendering device and present it to our user!
        ' Of course, we also need to do a few calculations to make sure it appears where it should.
        With srcRect
            .X1 = 2
            .x2 = .X1 + PIC_X
            .Y1 = 2
            .y2 = .Y1 + PIC_Y
        End With
    
        With destRect
            .X1 = 0
            .x2 = frmMain.picTempSpell.Width
            .Y1 = 0
            .y2 = frmMain.picTempSpell.Height
        End With
    
        Call D3DDevice8.EndScene
        Call D3DDevice8.Present(srcRect, destRect, frmMain.picTempSpell.hWnd, ByVal 0)

        With frmMain.picTempSpell
            .Top = y
            .Left = x
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDraggedSpell", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawSpellDesc(ByVal spellnum As Long)
Dim spellpic As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim Top As Long, Left As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure the spell number is valid before we continue, if it isn't we'll simply skip rendering the icon.
    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        
        ' Retrieve the spell image, and check if it is valid.
        spellpic = Spell(spellnum).Icon
        If spellpic < 1 Or spellpic > NumSpellIcons Then Exit Sub

        ' Let's open clear ourselves a nice clean slate to render on shall we?
        Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
        Call D3DDevice8.BeginScene
        
        ' Calculate what image we need to use to render here.
        ' Note that the tooltips do not support animations.
        ' It simply shows the first icon of the inventory row.
        Top = 0
        Left = 0
        
        ' Render it on the surface.
        Call RenderGraphic(Tex_SpellIcon(spellpic), 0, 0, PIC_X, PIC_Y, 0, 0, Left, Top)
        
        ' We're done for now, so we can close the lovely little rendering device and present it to our user!
        ' Of course, we also need to do a few calculations to make sure it appears where it should.
        With srcRect
            .X1 = 0
            .x2 = PIC_X
            .Y1 = 0
            .y2 = PIC_Y
        End With
    
        With destRect
            .X1 = 0
            .x2 = frmMain.picSpellDescPic.Width
            .Y1 = 0
            .y2 = frmMain.picSpellDescPic.Height
        End With
    
        Call D3DDevice8.EndScene
        Call D3DDevice8.Present(srcRect, destRect, frmMain.picSpellDescPic.hWnd, ByVal 0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawspellDesc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHotbar()
Dim i As Long, num As Long, n As Long, text As String
Dim IconTop As Long, IconLeft As Long, Top As Long, Left As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Draw the background of the hotbar in the screen.
    Call RenderGraphic(Tex_Hotbar, 0, 0, D3DT_TEXTURE(Tex_Hotbar).Width, D3DT_TEXTURE(Tex_Hotbar).Height, 0, 0, 0, 0)
    
    ' Loop through the hotbar slots and render the appropriate items / spells.
    For i = 1 To MAX_HOTBAR
    
        ' Some positioning calculations
        Top = HotbarTop
        Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
        
        IconTop = 0
        IconLeft = 0
        
        Select Case Hotbar(i).sType
            Case 1 ' The hotbar slot contains an item!
                ' Retrieve the item number we're using, and check if it is valid.
                ' If it is not valid, we're not even going to bother rendering the icon.
                num = Hotbar(i).Slot
                If num >= 1 And num <= MAX_ITEMS Then
                    ' Well then, the item is valid! Let's check if the icon is valid as well.
                    If Item(num).Pic >= 1 And Item(num).Pic <= NumItems Then
                        ' Everything checks out, we can render it!
                        ' Of course we need to know what item image to render there as well.
                        IconLeft = D3DT_TEXTURE(Tex_Item(Item(num).Pic)).Width / 2
                        
                        ' Now let's actually render it. :)
                        Call RenderGraphic(Tex_Item(Item(num).Pic), Left, Top, PIC_X, PIC_Y, 0, 0, IconLeft, IconTop)
                    End If
                End If
            Case 2 ' The hotbar slot contains a spell!
                ' Let's check if the spell  we're trying to render actually exists.
                num = Hotbar(i).Slot
                If num >= 1 And num <= MAX_SPELLS Then
                    ' It exists, so let's check if the icon is valid!
                        If Spell(num).Icon > 0 Then
                            ' Check if the spell is on a cooldown, if it is we need to make a slight adjustment to the
                            ' position of the graphic we're grabbing to render.
                            For n = 1 To MAX_PLAYER_SPELLS
                                ' Is this the spell we're trying to figure out?
                                If PlayerSpells(n) = Hotbar(i).Slot Then
                                    ' Let's check if this spell is on a cooldown or not.
                                    If Not SpellCD(n) = 0 Then
                                        IconLeft = 32
                                    End If
                                End If
                            Next
                            
                            ' Now let's actually render it. :)
                        Call RenderGraphic(Tex_SpellIcon(Spell(num).Icon), Left, Top, PIC_X, PIC_Y, 0, 0, IconLeft, IconTop)
                        End If
                End If
        End Select
        
        ' Render the hotbar letters on top of the icons.
        text = "F" & Str(i)
        Call RenderText(MainFont, text, Left + 2, Top + 16, White)
    Next
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .x2 = frmMain.picHotbar.Width
        .Y1 = 0
        .y2 = frmMain.picHotbar.Height
    End With
    
    With destRect
        .X1 = 0
        .x2 = frmMain.picHotbar.Width
        .Y1 = 0
        .y2 = frmMain.picHotbar.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picHotbar.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHotbar", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawPlayerSpells()
Dim i As Long, x As Long, y As Long, spellnum As Long, spellicon As Long
Dim Amount As String
Dim colour As Long, Left As Long, Top As Long, RenderLeft As Long, RenderTop As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' If we're not in the game, exit it. Although it should be impossible to get here without being in-game.
    If Not InGame Then Exit Sub
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Time to start looping through the spells!
    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i)
        
        ' Check if the spellnumber and icon are valid.
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellicon = Spell(spellnum).Icon
            If spellicon > 0 And spellicon <= NumSpellIcons Then
                ' They are, let's set the location to grab the image from.
                Top = 0
                Left = 0
                ' If the spell's on a cooldown we need to grab the second image.
                If Not SpellCD(i) = 0 Then
                    Left = 32
                End If

                RenderTop = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                RenderLeft = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                
                ' Render the icon to the display!
                Call RenderGraphic(Tex_SpellIcon(spellicon), RenderLeft, RenderTop, PIC_X, PIC_Y, 0, 0, Left, Top)
            End If
        End If
    Next
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .x2 = frmMain.picSpells.Width
        .Y1 = 28
        .y2 = frmMain.picSpells.Height
    End With
    
    With destRect
        .X1 = 0
        .x2 = frmMain.picSpells.Width
        .Y1 = 32
        .y2 = frmMain.picSpells.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picSpells.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerSpells", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
