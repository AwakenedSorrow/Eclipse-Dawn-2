Attribute VB_Name = "modRendering"
Public Sub Render_Game()
Dim X As Long, Y As Long, i As Long
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
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                    Call RenderMapTile(X, Y)
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
    For Y = 0 To Map.MaxY
        
        ' Check if we have any sprites loaded, if so we can start rendering players and NPCs!
        If NumCharacters > 0 Then
            ' Player Characters
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).Y = Y Then
                        Call RenderPlayer(i)
                    End If
                End If
            Next
            
            ' Non-Player Characters
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Y = Y Then
                    Call RenderNPC(i)
                End If
            Next
        End If
        
        ' Time to start cracking on rendering Resources!
        If NumResources > 0 And Resources_Init And Resource_Index > 0 Then
            For i = 1 To Resource_Index
                If MapResource(i).Y = Y Then
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
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                    Call RenderUpperMapTile(X, Y)
                End If
            Next
        Next
    End If
    
    ' Some Map Editor specific stuff, such as Directional Blocking and the Mouse Tile Outline.
    ' Fairly important stuff if you want to get it all to work. :)
    If InMapEditor Then
        ' Directional Blocking
        If frmEditor_Map.optBlock.value = True Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.top To TileView.bottom
                    If IsValidMapPoint(X, Y) Then
                        Call RenderDirBlock(X, Y)
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
        RenderText MainFont, "FPS: " & CStr(GameFPS), 2, 39, yellow
    End If
            
    ' draw cursor, player X and Y locations
    If BLoc Then
        RenderText MainFont, Trim$("cur x: " & CurX & " y: " & CurY), 2, 1, yellow
        RenderText MainFont, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 2, 15, yellow
        RenderText MainFont, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 2, 27, yellow
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
    RenderText MainFont, Map.name, DrawMapNameX, DrawMapNameY, yellow
            
    ' If we're in the map editor, draw the attributes.
    If InMapEditor Then Call DrawMapAttributes
    
    ' End the rendering scene and present it to the player.
    ' This makes sure we can actually SEE what we rendered onto the device above.
    With srcRect
        .X1 = 0
        .X2 = frmMain.picScreen.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picScreen.ScaleHeight
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

Public Sub RenderMapTile(ByVal X As Long, ByVal Y As Long)
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With Map.Tile(X, Y)
        ' Time to loop through our layers for this tile.
        For i = MapLayer.Ground To MapLayer.Mask2
            ' Should we skip the tile?
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                Call RenderGraphic(Tex_TileSet(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), PIC_X, PIC_Y, 0, 0, .Layer(i).X * PIC_X, .Layer(i).Y * PIC_Y)
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

Public Sub RenderUpperMapTile(ByVal X As Long, ByVal Y As Long)
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With Map.Tile(X, Y)
        ' Time to loop through our layers for this tile.
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' Should we skip the tile?
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                Call RenderGraphic(Tex_TileSet(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), PIC_X, PIC_Y, 0, 0, .Layer(i).X * PIC_X, .Layer(i).Y * PIC_Y)
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

Public Sub RenderPlayer(ByVal Index As Long)
Dim SpriteFrame As Byte, i As Long, X As Long, Y As Long
Dim Sprite As Long, SpriteDir As Long
Dim attackspeed As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
    
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
            Case North
                If (Player(Index).yOffset > 8) Then SpriteFrame = Player(Index).Step
            Case South
                If (Player(Index).yOffset < -8) Then SpriteFrame = Player(Index).Step
            Case West
                If (Player(Index).XOffset > 8) Then SpriteFrame = Player(Index).Step
            Case East
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
        Case North
            SpriteDir = 3
        Case East
            SpriteDir = 2
        Case South
            SpriteDir = 0
        Case West
            SpriteDir = 1
    End Select

    ' Calculate the X position to render the player sprite at. Along with offset, of course!
    X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((D3DT_TEXTURE(Tex_Character(Sprite)).Width / 4 - 32) / 2)

    ' Time to work on the Y position.
    ' But first, let's check if the sprite is more than 32 pixels high.
    If (D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
    End If

    ' We're done here, let's render the sprite!
    Call RenderSprite(Sprite, X, Y, SpriteFrame, SpriteDir)
    
    ' Paperdoll time!
    ' This will render any equipped items that the player has on which have been assigned a
    ' paperdoll.
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Red = Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Red
                Green = Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Green
                Blue = Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Blue
                Alpha = Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Alpha
                Call RenderPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, SpriteFrame, SpriteDir, Red, Green, Blue, Alpha)
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

Public Sub RenderSprite(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, ByVal SpriteFrame As Long, ByVal SpriteDir As Long)
Dim X As Long, Y As Long, Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Is the sprite valid? If not, exit the sub to prevent issues from occuring.
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    ' Convert the provided values to values we can use on the map.
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    
    ' Pre-Calculate these values, it makes the render line look a lot cleaner.
    Width = D3DT_TEXTURE(Tex_Character(Sprite)).Width / 4
    Height = D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4
    SpriteFrame = SpriteFrame * Width
    SpriteDir = SpriteDir * Height
    
    ' Render the sprite itself! Please do -NOT- touch this line unless you know what you're doing.
    Call RenderGraphic(Tex_Character(Sprite), X, Y, Width, Height, 0, 0, SpriteFrame, SpriteDir)
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderBlood(ByVal Index As Long)
Dim X As Long, Y As Long, Sprite As Long

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
            
        ' Calculate the locations we'll use for the blood.
        X = ConvertMapX(.X * PIC_X)
        Y = ConvertMapY(.Y * PIC_Y)
        Sprite = (.Sprite - 1) * PIC_X
            
        ' Now that we've got all that sorted, let's get to rendering this bugger!
        Call RenderGraphic(Tex_Blood, X, Y, PIC_X, PIC_Y, 0, 0, Sprite, 0, 255, 255, 255, .Alpha)
   
    End With
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderBlood", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderMapItem(ByVal Index As Long)
Dim AnimFrame As Long, X As Long, Y As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte

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
        
        ' Calculate the locations we'll be using for the item.
        X = ConvertMapX(MapItem(Index).X * PIC_X)
        Y = ConvertMapY(MapItem(Index).Y * PIC_Y)
        AnimFrame = ItemAnimFrame(MapItem(Index).num) * PIC_X
        
        ' The Colors maaan!
        Red = Item(MapItem(Index).num).Red
        Green = Item(MapItem(Index).num).Green
        Blue = Item(MapItem(Index).num).Blue
        Alpha = Item(MapItem(Index).num).Alpha
        
        ' We've done all the fancy stuff, now let's get to rendering this item!
        Call RenderGraphic(Tex_Item(.Pic), X, Y, PIC_X, PIC_Y, 0, 0, AnimFrame, 0, Red, Green, Blue, Alpha)
        
    End With

' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderMapItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderAnimation(ByVal Index As Long, ByVal Layer As Byte)
Dim Sprite As Long
Dim i As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim X As Long, Y As Long
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
    
    ' Get and set the Height and Width of the sprite frame we'll be using.
    Width = D3DT_TEXTURE(Tex_Animation(Sprite)).Width / FrameCount
    Height = D3DT_TEXTURE(Tex_Animation(Sprite)).Height
    
    ' Set the Animation Frame we'll be using.
    ' Note that unlike most other render subs, this frame already includes the full location of the texture.
    ' This does not need to be multiplied again further on to get the proper frame texture.
    AnimFrame = (AnimInstance(Index).FrameIndex(Layer) - 1) * Width

    ' Let's change the X/Y offset if the Animation is locked onto a target so it actually displays on
    ' top of it.
    If AnimInstance(Index).LockType > TargetTypeNone Then ' if <> none
        ' Locked On to a Player.
        If AnimInstance(Index).LockType = TargetTypePlayer Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).XOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TargetTypeNPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).XOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).yOffset
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
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If
    
    ' Convert these values beforehand. Saves some space down there.
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' Get the colors.
    Red = Animation(AnimInstance(Index).Animation).Red(Layer)
    Green = Animation(AnimInstance(Index).Animation).Green(Layer)
    Blue = Animation(AnimInstance(Index).Animation).Blue(Layer)
    Alpha = Animation(AnimInstance(Index).Animation).Alpha(Layer)
    
    ' Render the actual texture. Should be some fancy animation now!
    Call RenderGraphic(Tex_Animation(Sprite), X, Y, Width, Height, 0, 0, AnimFrame, 0, Red, Green, Blue, Alpha)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderAnimation", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderNPC(ByVal Index As Long)
Dim SpriteAnim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, SpriteDir As Long
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
            Case North
                If (MapNpc(Index).yOffset > 8) Then SpriteAnim = MapNpc(Index).Step
            Case South
                If (MapNpc(Index).yOffset < -8) Then SpriteAnim = MapNpc(Index).Step
            Case West
                If (MapNpc(Index).XOffset > 8) Then SpriteAnim = MapNpc(Index).Step
            Case East
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
        Case North
            SpriteDir = 3
        Case East
            SpriteDir = 2
        Case South
            SpriteDir = 0
        Case West
            SpriteDir = 1
    End Select

    ' Calculate the X
    X = MapNpc(Index).X * PIC_X + MapNpc(Index).XOffset - ((D3DT_TEXTURE(Tex_Character(Sprite)).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(Index).Y * PIC_Y + MapNpc(Index).yOffset - ((D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(Index).Y * PIC_Y + MapNpc(Index).yOffset
    End If

    Call RenderSprite(Sprite, X, Y, SpriteAnim, SpriteDir)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderNpc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderMapResource(ByVal Resource_num As Long)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub
        
    ' Get the Original Resource.
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).Data1
    
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
    X = (MapResource(Resource_num).X * PIC_X) - (D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Width / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Height + 32
    
    ' render it
    Call RenderResource(Resource_sprite, X, Y, D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Width, D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Height)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderMapResource", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderResource(ByVal Resource As Long, ByVal DX As Long, ByVal DY As Long, ByVal Width As Long, ByVal Height As Long)
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check if the resource is a valid number. If not, exit out.
    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ' Convert the provided values to something we can use on the map.
    X = ConvertMapX(DX)
    Y = ConvertMapY(DY)
    
    ' Render the actual resource on the map!
    Call RenderGraphic(Tex_Resource(Resource), X, Y, Width, Height, 0, 0, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderResource", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderDirBlock(ByVal X As Long, ByVal Y As Long)
Dim i As Long, Left As Long, top As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Render the Grid Texture.
    Call RenderGraphic(Tex_DirBlock, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), PIC_X, PIC_Y, 0, 0, 0, 24)
    
    ' render dir blobs
    For i = 1 To 4
        Left = (i - 1) * 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(i)) Then
            top = 8
        Else
            top = 16
        End If
        'render the actual thing!
        Call RenderGraphic(Tex_DirBlock, ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), 8, 8, 0, 0, Left, top)
        
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderDirBlock", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderTileOutline()
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' If we're editing Directional Blocks, disable this outline. It just looks silly.
    If frmEditor_Map.optBlock.value Then Exit Sub
    
    ' Render the outline to the screen!
    Call RenderGraphic(Tex_Outline, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), PIC_X, PIC_Y, 0, 0, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTileOutline", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim top As Long, Right As Long
Dim barWidth As Long
Dim i As Long, npcNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' dynamic bar calculations
    sWidth = D3DT_TEXTURE(Tex_Bars).Width
    sHeight = D3DT_TEXTURE(Tex_Bars).Height / 2
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(npcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).XOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).yOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (Npc(npcNum).HP / sWidth)) * sWidth
                
                ' draw the content of the bar.
                top = 0 ' The Bar itself
                Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), barWidth, sHeight, 0, 0, 0, top, 255, 0, 0, 255)
                
                ' draw bar overlay
                top = sHeight * 1 ' Bar overlay
                Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sWidth, sHeight, 0, 0, 0, top, 128, 128, 128, 255)
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
            
            ' draw the bar proper
            top = 0 ' The Bar Itself
            Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), barWidth, sHeight, 0, 0, 0, top, 0, 0, 255, 255)
            
            ' draw bar Overlay
            top = sHeight * 1 ' The bar overlay
            Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sWidth, sHeight, 0, 0, 0, top, 128, 128, 128, 255)
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).yOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' draw the content of the bar.
        top = 0 ' The Bar itself
        Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), barWidth, sHeight, 0, 0, 0, top, 255, 0, 0, 255)
                
        ' draw bar overlay
        top = sHeight * 1 ' Bar overlay
        Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sWidth, sHeight, 0, 0, 0, top, 128, 128, 128, 255)
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
                    
                    ' draw the content of the bar.
                    top = 0 ' The Bar itself
                    Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), barWidth, sHeight, 0, 0, 0, top, 255, 0, 0, 255)
                
                    ' draw bar overlay
                    top = sHeight * 1 ' Bar overlay
                    Call RenderGraphic(Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sWidth, sHeight, 0, 0, 0, top, 128, 128, 128, 255)
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

Public Sub RenderHoverAndTarget()
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Handle the rendering of the target texture if one exists.
    If myTarget > 0 Then
        If myTargetType = TargetTypePlayer Then ' If the target is a player.
            Call RenderTarget((Player(myTarget).X * 32) + Player(myTarget).XOffset, (Player(myTarget).Y * 32) + Player(myTarget).yOffset)
        ElseIf myTargetType = TargetTypeNPC Then ' If the target is an NPC.
            Call RenderTarget((MapNpc(myTarget).X * 32) + MapNpc(myTarget).XOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).yOffset)
        End If
    End If
    
    ' Now it's time to figure out if we're mousing over a Player or NPC on the map.
    ' If we are, we'll render a hover texture over their characters so we know it's a valid character we can target to cast spells on.
    For i = 1 To Player_HighIndex ' Players
        If IsPlaying(i) Then
            If Player(i).Map = Player(MyIndex).Map Then
                If CurX = Player(i).X And CurY = Player(i).Y Then ' Is our cursor over something?
                    If myTargetType = TargetTypePlayer And myTarget = i Then
                        'We're already targetting this player, so no point in rendering this as well.
                    Else
                        Call RenderHover(TargetTypePlayer, i, (Player(i).X * 32) + Player(i).XOffset, (Player(i).Y * 32) + Player(i).yOffset)
                    End If
                End If
            End If
        End If
    Next
    For i = 1 To Npc_HighIndex 'NPCs
        If MapNpc(i).num > 0 Then
            If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then ' Is our cursor over something?
                If myTargetType = TargetTypeNPC And myTarget = i Then
                    'We're already targetting this NPC, so no point in rendering this as well.
                Else
                    Call RenderHover(TargetTypeNPC, i, (MapNpc(i).X * 32) + MapNpc(i).XOffset, (MapNpc(i).Y * 32) + MapNpc(i).yOffset)
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

Public Sub RenderTarget(ByVal X As Long, ByVal Y As Long)
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Let's do some magic and figure out where we'll need to render this texture!
    Width = D3DT_TEXTURE(Tex_Target).Width / 2
    Height = D3DT_TEXTURE(Tex_Target).Height
    
    ' Center it on the Target.
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)
    
    ' And convert it to be useful on our map!
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' Now render the texture to the screen.
    Call RenderGraphic(Tex_Target, X, Y, Width, Height, 0, 0, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTarget", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderHover(ByVal tType As Long, ByVal target As Long, ByVal X As Long, ByVal Y As Long)
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Let's do some magic and figure out where we'll need to render this texture!
    Width = D3DT_TEXTURE(Tex_Target).Width / 2
    Height = D3DT_TEXTURE(Tex_Target).Height
    
    ' Center it on the Target.
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)
    
    ' And convert it to be useful on our map!
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' Now render the texture to the screen.
    Call RenderGraphic(Tex_Target, X, Y, Width, Height, 0, 0, Width, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderHover", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
 
    ConvertMapX = X - (TileView.Left * PIC_X) - Camera.Left
   
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
 
Public Function ConvertMapY(ByVal Y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
 
    ConvertMapY = Y - (TileView.top * PIC_Y) - Camera.top
   
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub RenderPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal Sprite As Long, ByVal SpriteFrame As Long, ByVal SpriteDir As Long, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, ByVal Alpha As Byte)
Dim top As Long, Left As Long
Dim X As Long, Y As Long
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if the sprite is a valid one. If it isn't then exit out of the sub.
    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    top = SpriteDir * (D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Height / 4)
    Left = SpriteFrame * (D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Width / 4)
    
    ' Caclculate a few things we might need.
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Width / 4)
    Height = (D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Height / 4)
    
    ' Rendering time!
    Call RenderGraphic(Tex_Paperdoll(Sprite), X, Y, Width, Height, 0, 0, Left, top, Red, Green, Blue, Alpha)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderPaperdoll", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    offsetX = Player(MyIndex).XOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).XOffset > 0 Then
                offsetX = Player(MyIndex).XOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).yOffset > 0 Then
                offsetY = Player(MyIndex).yOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        offsetX = 32
        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).XOffset < 0 Then
                offsetX = Player(MyIndex).XOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        offsetY = 32
        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).yOffset < 0 Then
                offsetY = Player(MyIndex).yOffset + PIC_Y
            End If
        End If
        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .top = StartY
        .bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .top = offsetY
        .bottom = .top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
