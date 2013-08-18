Attribute VB_Name = "modGUI"
Option Explicit

Public Sub DrawGDI()
    'Cycle Through in-game stuff before cycling through editors
    If frmMenu.Visible Then
        If frmMenu.picCharacter.Visible Then DrawNewCharacterSprite
    End If
    
    If frmMain.Visible Then
        If frmMain.picTempInv.Visible Then DrawDraggedItem frmMain.picTempInv.Left, frmMain.picTempInv.top
        If frmMain.picTempSpell.Visible Then DrawDraggedSpell frmMain.picTempSpell.Left, frmMain.picTempSpell.top
        If frmMain.picSpellDesc.Visible Then DrawSpellDesc LastSpellDesc
        If frmMain.picItemDesc.Visible Then DrawItemDesc LastItemDesc
        If frmMain.picHotbar.Visible Then DrawHotbar
        If frmMain.picInventory.Visible Then DrawInventory
        If frmMain.picSpells.Visible Then DrawPlayerSpells
        If frmMain.picCharacter.Visible Then DrawCharacterScreen
        If frmMain.picShop.Visible Then DrawShop
        If frmMain.picTempBank.Visible Then DrawDraggedBank frmMain.picTempBank.Left, frmMain.picTempBank.top
        If frmMain.picBank.Visible Then DrawBank
        If frmMain.picTrade.Visible Then DrawTradeOffer: DrawTradeReceive
    End If
    
    
    If frmEditor_Animation.Visible Then
        EditorAnim_DrawAnim
    End If
    
    If frmEditor_Item.Visible Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
    End If
    
    If frmEditor_Map.Visible Then
        EditorMap_DrawTileset
        If frmEditor_Map.fraMapItem.Visible Then EditorMap_DrawMapItem
        If frmEditor_Map.fraMapKey.Visible Then EditorMap_DrawKey
    End If
    
    If frmEditor_NPC.Visible Then
        EditorNpc_DrawSprite
    End If
    
    If frmEditor_Resource.Visible Then
        EditorResource_DrawNormalSprite
        EditorResource_DrawExhaustedSprite
    End If
    
    If frmEditor_Spell.Visible Then
        EditorSpell_DrawIcon
    End If

End Sub

Public Sub DrawHotbar()
Dim i As Long, num As Long, n As Long, text As String, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim IconTop As Long, IconLeft As Long, top As Long, Left As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Draw the background of the hotbar in the screen.
    Call RenderGraphic(Tex_GUI(HotBarE), 0, 0, D3DT_TEXTURE(Tex_GUI(HotBarE)).Width, D3DT_TEXTURE(Tex_GUI(HotBarE)).Height, 0, 0, 0, 0)
    
    ' Loop through the hotbar slots and render the appropriate items / spells.
    For i = 1 To MAX_HOTBAR
    
        ' Some positioning calculations
        top = HotbarTop
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
                        IconLeft = ItemAnimFrame(num) * PIC_X
                        
                        ' The Colors maaan!
                        Red = Item(num).Red
                        Green = Item(num).Green
                        Blue = Item(num).Blue
                        Alpha = Item(num).Alpha
                        
                        ' Now let's actually render it. :)
                        Call RenderGraphic(Tex_Item(Item(num).Pic), Left, top, PIC_X, PIC_Y, 0, 0, IconLeft, IconTop, Red, Green, Blue, Alpha)
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
                        Call RenderGraphic(Tex_SpellIcon(Spell(num).Icon), Left, top, PIC_X, PIC_Y, 0, 0, IconLeft, IconTop)
                        End If
                End If
        End Select
        
        ' Render the hotbar letters on top of the icons.
        text = "F" & Str(i)
        Call RenderText(MainFont, text, Left + 2, top + 16, White)
    Next
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmMain.picHotbar.Width
        .Y1 = 0
        .Y2 = frmMain.picHotbar.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmMain.picHotbar.Width
        .Y1 = 0
        .Y2 = frmMain.picHotbar.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picHotbar.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHotbar", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayerSpells()
Dim i As Long, X As Long, Y As Long, spellnum As Long, spellicon As Long
Dim Amount As String
Dim Colour As Long, Left As Long, top As Long, RenderLeft As Long, RenderTop As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' If we're not in the game, exit it. Although it should be impossible to get here without being in-game.
    If Not InGame Then Exit Sub

    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Render the backdrop
    Call RenderGraphic(Tex_GUI(SpellsE), 0, 0, D3DT_TEXTURE(Tex_GUI(SpellsE)).Width, D3DT_TEXTURE(Tex_GUI(SpellsE)).Height, 0, 0, 0, 0)
    
    ' Time to start looping through the spells!
    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i)
        
        ' Check if the spellnumber and icon are valid.
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellicon = Spell(spellnum).Icon
            If spellicon > 0 And spellicon <= NumSpellIcons Then
                ' They are, let's set the location to grab the image from.
                top = 0
                Left = 0
                ' If the spell's on a cooldown we need to grab the second image.
                If Not SpellCD(i) = 0 Then
                    Left = 32
                End If

                RenderTop = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                RenderLeft = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                
                ' Render the icon to the display!
                Call RenderGraphic(Tex_SpellIcon(spellicon), RenderLeft, RenderTop, PIC_X, PIC_Y, 0, 0, Left, top)
            End If
        End If
    Next
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmMain.picSpells.Width
        .Y1 = 0
        .Y2 = frmMain.picSpells.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmMain.picSpells.Width
        .Y1 = 0
        .Y2 = frmMain.picSpells.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picSpells.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerSpells", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawInventory()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim top As Long, Left As Long
Dim Colour As Long
Dim tmpItem As Long, amountModifier As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim AnimLeft As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' If we're not in-game we should probably not be here to begin with.
    ' So let's exit out before things go awry!
    If Not InGame Then Exit Sub
    
    ' Reset the gold label every ~0.5 seconds.
    ' Not doing this every frame as it spazzes out pretty badly if you do.
    ' But not doing this means the counter never resets to 0 if you drop or lose all your gold.
    If GoldTimer + 500 <= GetTickCount Then
        frmMain.lblGold.Caption = "0g"
        GoldTimer = GetTickCount + 500
    End If
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Render the backdrop
    Call RenderGraphic(Tex_GUI(InventoryE), 0, 0, D3DT_TEXTURE(Tex_GUI(InventoryE)).Width, D3DT_TEXTURE(Tex_GUI(InventoryE)).Height, 0, 0, 0, 0)
    
    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            
            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                    If TradeYourOffer(X).num = i Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ItemTypeCurrency Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(X).value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(X).value
                            End If
                        End If
                    End If
                Next
            End If

            If itempic > 0 And itempic <= NumItems Then
                    
                ' Calculate where we need to render the item.
                top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    
                ' Calculate the Animation Frame
                AnimLeft = ItemAnimFrame(itemnum) * PIC_X
                
                ' The Colors maaan!
                Red = Item(itemnum).Red
                Green = Item(itemnum).Green
                Blue = Item(itemnum).Blue
                Alpha = Item(itemnum).Alpha
                
                ' Render the item Icon.
                Call RenderGraphic(Tex_Item(itempic), Left, top, PIC_X, PIC_Y, 0, 0, AnimLeft, 0, Red, Green, Blue, Alpha)
                    
                ' If item is a stack - draw the amount you have
                If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                    Y = top + 22
                    X = Left - 4
                        
                    Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                        
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        Colour = White
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        Colour = Yellow
                    ElseIf Amount > 10000000 Then
                        Colour = BrightGreen
                    End If
                        
                    Call RenderText(MainFont, Format$(ConvertCurrency(Str(Amount)), "#,###,###,###"), X, Y, Colour)

                    ' Check if it's gold, and update the label
                    If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                        frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                    End If
                End If
            End If
        End If
NextLoop:
    Next
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmMain.picInventory.Width
        .Y1 = 0
        .Y2 = frmMain.picInventory.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmMain.picInventory.Width
        .Y1 = 0
        .Y2 = frmMain.picInventory.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picInventory.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventory", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDraggedItem(ByVal X As Long, ByVal Y As Long)
Dim top As Long, Left As Long
Dim itemnum As Long, itempic As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim AnimLeft As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte

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
        
         ' Render the backdrop
        Call RenderGraphic(Tex_GUI(DragBoxE), 0, 0, D3DT_TEXTURE(Tex_GUI(DragBoxE)).Width, D3DT_TEXTURE(Tex_GUI(DragBoxE)).Height, 0, 0, 0, 0)
        
        ' Calculate what image we need to grab from the texture.
        top = 0
        
        ' Calculate the Animation Frame
        AnimLeft = ItemAnimFrame(itemnum) * PIC_X
        
        ' The Colors maaan!
        Red = Item(itemnum).Red
        Green = Item(itemnum).Green
        Blue = Item(itemnum).Blue
        Alpha = Item(itemnum).Alpha
        
        ' Render the texture to the screen, we're using a 2pixel offset to make sure it's centered and doesn't clip
        ' with the picturebox. It's an original design choice in Mirage4, lord knows why.
        Call RenderGraphic(Tex_Item(itempic), 2, 2, PIC_X, PIC_Y, 0, 0, AnimLeft, top, Red, Green, Blue, Alpha)
        
        ' We're done for now, so we can close the lovely little rendering device and present it to our user!
        ' Of course, we also need to do a few calculations to make sure it appears where it should.
        With srcRect
            .X1 = 0
            .X2 = frmMain.picTempInv.Width
            .Y1 = 0
            .Y2 = frmMain.picTempInv.Height
        End With
    
        With destRect
            .X1 = 0
            .X2 = frmMain.picTempInv.Width
            .Y1 = 0
            .Y2 = frmMain.picTempInv.Height
        End With
    
        Call D3DDevice8.EndScene
        Call D3DDevice8.Present(srcRect, destRect, frmMain.picTempInv.hWnd, ByVal 0)

        With frmMain.picTempInv
            .top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDraggedItem", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDraggedBank(ByVal X As Long, ByVal Y As Long)
Dim top As Long, Left As Long
Dim itemnum As Long, itempic As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim AnimLeft As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Retrieve the item number we're trying to drag around.
    itemnum = GetBankItemNum(DragBankSlotNum)
    
    ' If the item number is valid then make sure we do something with it, wouldn't like to have an invisible icon right?
    If itemnum > 0 And itemnum <= MAX_ITEMS Then
    
        ' Retrieve the item texture and make sure it is valid before we continue.
        itempic = Item(itemnum).Pic
        If itempic < 1 Or itempic > NumItems Then Exit Sub
        
        ' Let's open clear ourselves a nice clean slate to render on shall we?
        Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
        Call D3DDevice8.BeginScene
        
         ' Render the backdrop
        Call RenderGraphic(Tex_GUI(DragBoxE), 0, 0, D3DT_TEXTURE(Tex_GUI(DragBoxE)).Width, D3DT_TEXTURE(Tex_GUI(DragBoxE)).Height, 0, 0, 0, 0)
        
        ' Calculate what image we need to grab from the texture.
        top = 0
        
        ' Calculate the Animation Frame
        AnimLeft = ItemAnimFrame(itemnum) * PIC_X
        
        ' The Colors maaan!
        Red = Item(itemnum).Red
        Green = Item(itemnum).Green
        Blue = Item(itemnum).Blue
        Alpha = Item(itemnum).Alpha
        
        ' Render the texture to the screen, we're using a 2pixel offset to make sure it's centered and doesn't clip
        ' with the picturebox. It's an original design choice in Mirage4, lord knows why.
        Call RenderGraphic(Tex_Item(itempic), 2, 2, PIC_X, PIC_Y, 0, 0, AnimLeft, top, Red, Green, Blue, Alpha)
        
        ' We're done for now, so we can close the lovely little rendering device and present it to our user!
        ' Of course, we also need to do a few calculations to make sure it appears where it should.
        With srcRect
            .X1 = 0
            .X2 = frmMain.picTempBank.Width
            .Y1 = 0
            .Y2 = frmMain.picTempBank.Height
        End With
    
        With destRect
            .X1 = 0
            .X2 = frmMain.picTempBank.Width
            .Y1 = 0
            .Y2 = frmMain.picTempBank.Height
        End With
    
        Call D3DDevice8.EndScene
        Call D3DDevice8.Present(srcRect, destRect, frmMain.picTempBank.hWnd, ByVal 0)

        With frmMain.picTempBank
            .top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDraggedBank", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawItemDesc(ByVal itemnum As Long)
Dim itempic As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim top As Long, Left As Long
Dim name As String, Firstletter As String * 1, Colour As Long, X As Long, Y As Long, desc As String

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
        
        ' Render the backdrop
        Call RenderGraphic(Tex_GUI(ItemDescE), 0, 0, D3DT_TEXTURE(Tex_GUI(ItemDescE)).Width, D3DT_TEXTURE(Tex_GUI(ItemDescE)).Height, 0, 0, 0, 0)
        
        ' Change the name a bit if it isn't valid.
        Firstletter = LCase$(Mid$(Trim$(Item(itemnum).name), 1, 1))
        If Firstletter = "$" Then
            name = (Mid$(Trim$(Item(itemnum).name), 2, Len(Trim$(Item(itemnum).name)) - 1))
        Else
            name = Trim$(Item(itemnum).name)
        End If
        
        ' Get the color of the item name.
        Select Case Item(itemnum).Rarity
            Case 0 ' grey
                Colour = Grey
            Case 1 ' white
                Colour = White
            Case 2 ' green
                Colour = BrightGreen
            Case 3 ' blue
                Colour = BrightBlue
            Case 4 ' purple
                Colour = Magenta
            Case 5 ' orange
                Colour = Orange
        End Select
        
        ' Render the Item Name.
        X = (frmMain.picItemDesc.Width / 2) - (GetTextWidth(MainFont, name) / 2)
        Y = 14
        Call RenderText(MainFont, name, X, Y, Colour)
        
        ' Render the Item Description
        X = 16
        Y = 120
        desc = WordWrap(Item(itemnum).desc, 27)
        Call RenderText(MainFont, desc, X, Y, White)
        
        ' Calculate what image we need to use to render here.
        top = 0
        Left = ItemAnimFrame(itemnum) * PIC_X
        
        ' The position where we need to render it.
        X = (frmMain.picItemDesc.Width / 2) - PIC_X
        Y = 40
        
        ' The Colors maaan!
        Red = Item(itemnum).Red
        Green = Item(itemnum).Green
        Blue = Item(itemnum).Blue
        Alpha = Item(itemnum).Alpha
        
        ' Render it on the surface.
        ' Note that we're using PIC_X and PIC_Y here, although they're multiplied.
        ' This is because we're scaling the image up a bit, to fill a 64x64 slot
        ' on the item description panel.
        Call RenderGraphic(Tex_Item(itempic), X, Y, PIC_X * 2, PIC_Y * 2, PIC_X, PIC_Y, Left, top, Red, Green, Blue, Alpha)
        
        ' We're done for now, so we can close the lovely little rendering device and present it to our user!
        ' Of course, we also need to do a few calculations to make sure it appears where it should.
        With srcRect
            .X1 = 0
            .X2 = frmMain.picItemDesc.Width
            .Y1 = 0
            .Y2 = frmMain.picItemDesc.Height
        End With
    
        With destRect
            .X1 = 0
            .X2 = frmMain.picItemDesc.Width
            .Y1 = 0
            .Y2 = frmMain.picItemDesc.Height
        End With
    
        Call D3DDevice8.EndScene
        Call D3DDevice8.Present(srcRect, destRect, frmMain.picItemDesc.hWnd, ByVal 0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawItemDesc", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawSpellDesc(ByVal spellnum As Long)
Dim spellpic As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim top As Long, Left As Long, X As Long, Y As Long, desc As String

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
        
        ' Render the backdrop
        Call RenderGraphic(Tex_GUI(SpellDescE), 0, 0, D3DT_TEXTURE(Tex_GUI(SpellDescE)).Width, D3DT_TEXTURE(Tex_GUI(SpellDescE)).Height, 0, 0, 0, 0)
        
        ' Render the Spell Name.
        X = (frmMain.picSpellDesc.Width / 2) - (GetTextWidth(MainFont, Spell(spellnum).name) / 2)
        Y = 14
        Call RenderText(MainFont, Spell(spellnum).name, X, Y, White)
        
        ' Render the Item Description
        X = 16
        Y = 120
        desc = WordWrap(Spell(spellnum).desc, 27)
        Call RenderText(MainFont, desc, X, Y, White)
        
        ' Calculate what image we need to use to render here.
        ' Note that the tooltips do not support animations.
        ' It simply shows the first icon of the inventory row.
        top = 0
        Left = 0
        
        ' Render it on the surface.
        Call RenderGraphic(Tex_SpellIcon(spellpic), (frmMain.picSpellDesc.Width / 2) - PIC_X, 40, PIC_X * 2, PIC_Y * 2, PIC_X, PIC_Y, Left, top)
        
        ' We're done for now, so we can close the lovely little rendering device and present it to our user!
        ' Of course, we also need to do a few calculations to make sure it appears where it should.
        With srcRect
            .X1 = 0
            .X2 = frmMain.picSpellDesc.Width
            .Y1 = 0
            .Y2 = frmMain.picSpellDesc.Height
        End With
    
        With destRect
            .X1 = 0
            .X2 = frmMain.picSpellDesc.Width
            .Y1 = 0
            .Y2 = frmMain.picSpellDesc.Height
        End With
    
        Call D3DDevice8.EndScene
        Call D3DDevice8.Present(srcRect, destRect, frmMain.picSpellDesc.hWnd, ByVal 0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawspellDesc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDraggedSpell(ByVal X As Long, ByVal Y As Long)
Dim top As Long, Left As Long
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
        
        ' Render the backdrop.
        Call RenderGraphic(Tex_GUI(DragBoxE), 0, 0, D3DT_TEXTURE(Tex_GUI(DragBoxE)).Width, D3DT_TEXTURE(Tex_GUI(DragBoxE)).Height, 0, 0, 0, 0)
        
        ' Calculate what image we need to grab from the texture.
        top = 0
        Left = 0
        
        ' Render the texture to the screen, we're using a 2pixel offset to make sure it's centered and doesn't clip
        ' with the picturebox. It's an original design choice in Mirage4, lord knows why.
        Call RenderGraphic(Tex_SpellIcon(spellpic), 2, 2, PIC_X, PIC_Y, 0, 0, Left, top)
        
        ' We're done for now, so we can close the lovely little rendering device and present it to our user!
        ' Of course, we also need to do a few calculations to make sure it appears where it should.
        With srcRect
            .X1 = 0
            .X2 = frmMain.picTempSpell.Width
            .Y1 = 0
            .Y2 = frmMain.picTempSpell.Height
        End With
    
        With destRect
            .X1 = 0
            .X2 = frmMain.picTempSpell.Width
            .Y1 = 0
            .Y2 = frmMain.picTempSpell.Height
        End With
    
        Call D3DDevice8.EndScene
        Call D3DDevice8.Present(srcRect, destRect, frmMain.picTempSpell.hWnd, ByVal 0)

        With frmMain.picTempSpell
            .top = Y
            .Left = X
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

Public Sub DrawCharacterScreen()
Dim faceNum As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim top As Long, Left As Long, AnimFrame As Long, i As Long, itemnum As Long, itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Time to render the background.
    Call RenderGraphic(Tex_GUI(CharacterE), 0, 0, D3DT_TEXTURE(Tex_GUI(CharacterE)).Width, D3DT_TEXTURE(Tex_GUI(CharacterE)).Height, 0, 0, 0, 0)
    
    ' Carry on to rendering the face
    ' Check if we even have any faces loaded to begin with.
    If NumFaces <> 0 Then
        ' Check if we have the face loaded for this particular sprite.
        faceNum = GetPlayerSprite(MyIndex)
        If faceNum >= 1 Or faceNum <= NumFaces Then
            ' Render the actual face.
            Call RenderGraphic(Tex_Face(faceNum), 49, 60, D3DT_TEXTURE(Tex_Face(faceNum)).Width, D3DT_TEXTURE(Tex_Face(faceNum)).Height, 0, 0, 0, 0)
        End If
    End If
    
    ' Now time to start rendering the equipped items.
    ' Check if we have any item textures loaded.
    If NumItems <> 0 Then
        ' Loop through the equipment slots.
        For i = 1 To Equipment.Equipment_Count - 1
            ' Make sure there's a valid item equipped.
            itemnum = GetPlayerEquipment(MyIndex, i)
            If itemnum >= 1 And itemnum <= MAX_ITEMS Then
                ' retrieve the item image and see if it is valid.
                itempic = Item(itemnum).Pic
                If itempic >= 1 And itempic <= NumItems Then
                    ' We can start calculating where we need to render the texture now!
                    top = EqTop
                    Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))

                    ' Get the item animation frame
                    AnimFrame = ItemAnimFrame(itemnum) * PIC_X
                    
                    ' And now to render it.
                    Call RenderGraphic(Tex_Item(itempic), Left, top, PIC_X, PIC_Y, 0, 0, AnimFrame, 0)
                End If
            End If
        Next
    End If

    ' Time to start rendering all the text on the Character Screen!
    ' First, let's start with the name.
    Call RenderText(MainFont, Trim$(Player(MyIndex).name), (frmMain.picCharacter.Width / 2) - (GetTextWidth(MainFont, Trim$(Player(MyIndex).name)) / 2), 33, White)
    
    ' Render the Stat Counts
    Call RenderText(MainFont, Trim$(Str$(Player(MyIndex).Stat(Strength))), 70, 169, White)
    Call RenderText(MainFont, Trim$(Str$(Player(MyIndex).Stat(Endurance))), 70, 184, White)
    Call RenderText(MainFont, Trim$(Str$(Player(MyIndex).Stat(Intelligence))), 70, 198, White)
    Call RenderText(MainFont, Trim$(Str$(Player(MyIndex).Stat(Agility))), 144, 169, White)
    Call RenderText(MainFont, Trim$(Str$(Player(MyIndex).Stat(Willpower))), 144, 184, White)
    Call RenderText(MainFont, Trim$(Str$(Player(MyIndex).POINTS)), 144, 198, White)
    
    ' Render the Point +, this is a bit tricky.
    ' Note that the rendered ones are ON TOP of the actual buttons, they're fakes.
    ' What we render here is just a visual representatation so that we can see them.
    ' They're not functional, and moving these will require you to move the + signs
    ' on the form.
    
    ' Make sure the player has points before we display these.
    If Player(MyIndex).POINTS > 0 Then
        Call RenderText(MainFont, "+", 96, 169, White)
        Call RenderText(MainFont, "+", 96, 184, White)
        Call RenderText(MainFont, "+", 96, 198, White)
        Call RenderText(MainFont, "+", 170, 169, White)
        Call RenderText(MainFont, "+", 170, 184, White)
    End If
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmMain.picCharacter.Width
        .Y1 = 0
        .Y2 = frmMain.picCharacter.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmMain.picCharacter.Width
        .Y1 = 0
        .Y2 = frmMain.picCharacter.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picCharacter.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawCharacterScreen", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNewCharacterSprite()
Dim Sprite As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Do we have a valic class selected? If not, exit out.
    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    ' Should we pick a male or female sprite?
    If frmMenu.optMale.value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    ' Is the sprite we're planning to render valid? If it isn't exit out of the sub.
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Calculate the width and height of a single sprite on the sheet.
    Width = D3DT_TEXTURE(Tex_Character(Sprite)).Width / 4
    Height = D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    Call RenderGraphic(Tex_Character(Sprite), 0, 0, Width, Height, 0, 0, 0, 0)
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmMenu.picSprite.Width
        .Y1 = 0
        .Y2 = frmMenu.picSprite.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmMenu.picSprite.Width
        .Y1 = 0
        .Y2 = frmMenu.picSprite.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMenu.picSprite.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNewCharacterSprite", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawTileset()
Dim Height As Long, Width As Long, Tileset As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim SrcTop As Long, SrcBottom As Long, SrcLeft As Long, SrcRight As Long, scrlX As Long, scrlY As Long
Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
     
    ' Get the position we need to render it all on.
    scrlX = frmEditor_Map.scrlPictureX.value * PIC_X
    scrlY = frmEditor_Map.scrlPictureY.value * PIC_Y
    
    Height = D3DT_TEXTURE(Tex_TileSet(Tileset)).Height - scrlY
    Width = D3DT_TEXTURE(Tex_TileSet(Tileset)).Width - scrlX
    
    SrcTop = frmEditor_Map.scrlPictureY.value * PIC_Y
    SrcLeft = frmEditor_Map.scrlPictureX.value * PIC_X
    SrcBottom = SrcTop + Height
    SrcRight = SrcLeft + Width
    
    ' Change the background we're rendering on to the right size.
    frmEditor_Map.picBackSelect.Height = Height
    frmEditor_Map.picBackSelect.Width = Width
    
    ' Render the tileset on the background.
    Call RenderGraphic(Tex_TileSet(Tileset), 0, 0, SrcRight, SrcBottom, 0, 0, SrcLeft, SrcTop)
    
    ' Render the tile selection square.
    X1 = (EditorTileX * PIC_X) - SrcLeft
    X2 = (EditorTileWidth * PIC_X) + X1
    Y1 = (EditorTileY * PIC_Y) - SrcTop
    Y2 = (EditorTileHeight * PIC_Y) + Y1
    Call DrawSelectionBox(X1, X2, Y1, Y2)
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Map.picBackSelect.Width
        .Y1 = 0
        .Y2 = frmEditor_Map.picBackSelect.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Map.picBackSelect.Width
        .Y1 = 0
        .Y2 = frmEditor_Map.picBackSelect.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor_Map.picBackSelect.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawTileset", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

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

Public Sub EditorMap_DrawMapItem()
Dim itemnum As Long, AnimFrame As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim srcRect As D3DRECT, destRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' What image are we using?
    itemnum = Item(frmEditor_Map.scrlMapItem.value).Pic
    ' Is it a valid one?
    If itemnum < 1 Or itemnum > NumItems Then
        Exit Sub
    End If
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' See what frame we need to use.
    AnimFrame = ItemAnimFrame(frmEditor_Map.scrlMapItem.value) * PIC_X
    
    ' The Colors maaan!
    Red = Item(frmEditor_Map.scrlMapItem.value).Red
    Green = Item(frmEditor_Map.scrlMapItem.value).Green
    Blue = Item(frmEditor_Map.scrlMapItem.value).Blue
    Alpha = Item(frmEditor_Map.scrlMapItem.value).Alpha
    
    ' Render the actual graphic to the screen.
    Call RenderGraphic(Tex_Item(itemnum), 0, 0, PIC_X, PIC_Y, 0, 0, AnimFrame, 0, Red, Green, Blue, Alpha)
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Map.picMapItem.Width
        .Y1 = 0
        .Y2 = frmEditor_Map.picMapItem.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Map.picMapItem.Width
        .Y1 = 0
        .Y2 = frmEditor_Map.picMapItem.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor_Map.picMapItem.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawMapItem", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawKey()
Dim itempic As Long, AnimFrame As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Retrieve what picture we'll be using.
    itempic = Item(frmEditor_Map.scrlMapKey.value).Pic

    ' Is it valid?
    If itempic < 1 Or itempic > NumItems Then
        Exit Sub
    End If

    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' See what frame we need to use.
    AnimFrame = ItemAnimFrame(frmEditor_Map.scrlMapKey.value) * PIC_X
    
    ' Render the actual graphic to the screen.
    Call RenderGraphic(Tex_Item(itempic), 0, 0, PIC_X, PIC_Y, 0, 0, AnimFrame, 0)
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Map.picMapKey.Width
        .Y1 = 0
        .Y2 = frmEditor_Map.picMapKey.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Map.picMapKey.Width
        .Y1 = 0
        .Y2 = frmEditor_Map.picMapKey.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor_Map.picMapKey.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawKey", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_DrawAnim()
Dim Animationnum As Long
Dim i As Long, Left As Long
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim ShouldRender As Boolean
Dim srcRect As D3DRECT, destRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Loop through the two animation screens.
    For i = 0 To 1
        ' Retrieve the animation we'll be rendering.
        Animationnum = frmEditor_Animation.scrlSprite(i).value
        ' Is it a valid one?
        If Animationnum >= 1 Or Animationnum <= NumAnimations Then
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                ' Let's open clear ourselves a nice clean slate to render on shall we?
                Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
                Call D3DDevice8.BeginScene
            
                If frmEditor_Animation.scrlFrameCount(i).value > 0 Then
                    ' total width divided by frame count
                    Width = D3DT_TEXTURE(Tex_Animation(Animationnum)).Width / frmEditor_Animation.scrlFrameCount(i).value
                    Height = D3DT_TEXTURE(Tex_Animation(Animationnum)).Height
                    
                    Left = (AnimEditorFrame(i) - 1) * Width
                    
                    Call RenderGraphic(Tex_Animation(Animationnum), 0, 0, Width, Height, 0, 0, Left, 0)
                End If
                
                ' We're done for now, so we can close the lovely little rendering device and present it to our user!
                ' Of course, we also need to do a few calculations to make sure it appears where it should.
                With srcRect
                    .X1 = 0
                    .X2 = frmEditor_Animation.picSprite(i).Width
                    .Y1 = 0
                    .Y2 = frmEditor_Animation.picSprite(i).Height
                End With
    
                With destRect
                    .X1 = 0
                    .X2 = frmEditor_Animation.picSprite(i).Width
                    .Y1 = 0
                    .Y2 = frmEditor_Animation.picSprite(i).Height
                End With
    
                Call D3DDevice8.EndScene
                Call D3DDevice8.Present(srcRect, destRect, frmEditor_Animation.picSprite(i).hWnd, ByVal 0)
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_DrawAnim", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawItem()
Dim itemnum As Long, AnimFrame As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim srcRect As D3DRECT, destRect As D3DRECT
    
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Retrieve the item picture we'll be using.
    itemnum = frmEditor_Item.scrlPic.value

    ' Check if it's a valid image.
    If itemnum < 1 Or itemnum > NumItems Then
        ' Clear the picturebox to make sure it doesn't display anything anymore.
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Figure out our animation frame
    AnimFrame = ItemAnimFrame(EditorIndex) * PIC_X
    
    ' The Colors maaan!
    Red = Item(EditorIndex).Red
    Green = Item(EditorIndex).Green
    Blue = Item(EditorIndex).Blue
    Alpha = Item(EditorIndex).Alpha
    
    ' Render the texture to the surface.
    Call RenderGraphic(Tex_Item(itemnum), 0, 0, PIC_X, PIC_Y, 0, 0, AnimFrame, 0, Red, Green, Blue, Alpha)
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Item.picItem.Width
        .Y1 = 0
        .Y2 = frmEditor_Item.picItem.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Item.picItem.Width
        .Y1 = 0
        .Y2 = frmEditor_Item.picItem.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor_Item.picItem.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawItem", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawPaperdoll()
Dim Sprite As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Let's retrieve the paperdoll image and check if it's valid.
    Sprite = frmEditor_Item.scrlPaperdoll.value
    If Sprite < 1 Or Sprite > NumPaperdolls Then
        ' Clear the picturebox to make sure it doesn't display anything anymore.
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Calculate the locations and Render it.
    Width = D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Width
    Height = D3DT_TEXTURE(Tex_Paperdoll(Sprite)).Height / 4
    Call RenderGraphic(Tex_Paperdoll(Sprite), 0, 0, Width, Height, 0, 0, 0, 0)
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Item.picPaperdoll.Width
        .Y1 = 0
        .Y2 = frmEditor_Item.picPaperdoll.Height
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Item.picPaperdoll.Width
        .Y1 = 0
        .Y2 = frmEditor_Item.picPaperdoll.Height
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor_Item.picPaperdoll.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawPaperdoll", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_DrawSprite()
Dim Sprite As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim X As Long, Y As Long, Width As Long, Height As Long
    

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' retrieve the sprite we'll be rendering.
    Sprite = frmEditor_NPC.scrlSprite.value

    ' Check if it's valid.
    If Sprite < 1 Or Sprite > NumCharacters Then
        ' Clear the screen so it doesn't leave lingering images.
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Calculate the locations and Render the graphic
    X = (frmEditor_NPC.picSprite.ScaleWidth / 2) - (D3DT_TEXTURE(Tex_Character(Sprite)).Width / 4) / 2
    Y = (frmEditor_NPC.picSprite.ScaleHeight / 2) - (D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4) / 2
    Width = D3DT_TEXTURE(Tex_Character(Sprite)).Width / 4
    Height = D3DT_TEXTURE(Tex_Character(Sprite)).Height / 4
    Call RenderGraphic(Tex_Character(Sprite), X, Y, Width, Height, 0, 0, 0, Height * 4)
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmEditor_NPC.picSprite.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_NPC.picSprite.ScaleHeight
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmEditor_NPC.picSprite.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_NPC.picSprite.ScaleHeight
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor_NPC.picSprite.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_DrawSprite", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_DrawNormalSprite()
Dim Sprite As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim X As Long, Y As Long, Height As Long, Width As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.value

    ' If the sprite is valid, let's calculate the locations and get it rendered.
    If Sprite >= 1 Or Sprite <= NumResources Then
        X = (frmEditor_Resource.picNormalPic.ScaleWidth / 2) - (D3DT_TEXTURE(Tex_Resource(Sprite)).Width / 2)
        Y = (frmEditor_Resource.picNormalPic.ScaleHeight / 2) - (D3DT_TEXTURE(Tex_Resource(Sprite)).Height / 2)
        Width = D3DT_TEXTURE(Tex_Resource(Sprite)).Width
        Height = D3DT_TEXTURE(Tex_Resource(Sprite)).Height
        Call RenderGraphic(Tex_Resource(Sprite), X, Y, Width, Height, 0, 0, 0, 0)
    End If

    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor_Resource.picNormalPic.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_DrawNormalSprite", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_DrawExhaustedSprite()
Dim Sprite As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim X As Long, Y As Long, Height As Long, Width As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.value
    
    ' If the sprite is valid, let's calculate the locations and get it rendered.
    If Sprite >= 1 Or Sprite <= NumResources Then
        X = (frmEditor_Resource.picExhaustedPic.ScaleWidth / 2) - (D3DT_TEXTURE(Tex_Resource(Sprite)).Width / 2)
        Y = (frmEditor_Resource.picExhaustedPic.ScaleHeight / 2) - (D3DT_TEXTURE(Tex_Resource(Sprite)).Height / 2)
        Width = D3DT_TEXTURE(Tex_Resource(Sprite)).Width
        Height = D3DT_TEXTURE(Tex_Resource(Sprite)).Height
        Call RenderGraphic(Tex_Resource(Sprite), X, Y, Width, Height, 0, 0, 0, 0)
    End If

    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor_Resource.picExhaustedPic.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_DrawExhaustedSprite", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawIcon()
Dim iconnum As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' retrieve the item icon.
    iconnum = frmEditor_Spell.scrlIcon.value
    
    ' Is it valid?
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        ' Nope, let's clear the picturebox.
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    Call RenderGraphic(Tex_SpellIcon(iconnum), 0, 0, PIC_X, PIC_Y, 0, 0, 0, 0)
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Spell.picSprite.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Spell.picSprite.ScaleHeight
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Spell.picSprite.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Spell.picSprite.ScaleHeight
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmEditor_Spell.picSprite.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_DrawIcon", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawShop()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As String
Dim Colour As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim top As Long, Left As Long, AnimFrame As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check if we're not here before we should be.
    If Not InGame Then Exit Sub
    
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Loop through the shop's inventory.
    For i = 1 To MAX_TRADES
        itemnum = Shop(InShop).TradeItem(i).Item
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            If itempic > 0 And itempic <= NumItems Then
                
                ' Work out the position
                top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                
                ' The frame we need to use.
                AnimFrame = ItemAnimFrame(Shop(InShop).TradeItem(i).Item) * PIC_X
                
                ' The Colors maaan!
                Red = Item(Shop(InShop).TradeItem(i).Item).Red
                Green = Item(Shop(InShop).TradeItem(i).Item).Green
                Blue = Item(Shop(InShop).TradeItem(i).Item).Blue
                Alpha = Item(Shop(InShop).TradeItem(i).Item).Alpha
                
                ' Render the icon.
                Call RenderGraphic(Tex_Item(itempic), Left, top, PIC_X, PIC_Y, 0, 0, AnimFrame, 0, Red, Green, Blue, Alpha)
                
                ' If item is a stack - draw the amount the shop has
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    Y = top + 22
                    X = Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    Call RenderText(MainFont, ConvertCurrency(Amount), X, Y, Colour)
                End If
            End If
        End If
    Next
    
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmMain.picShopItems.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picShopItems.ScaleHeight
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmMain.picShopItems.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picShopItems.ScaleHeight
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picShopItems.hWnd, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawShop", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawBank()
Dim i As Long, X As Long, Y As Long, itemnum As Long
Dim Amount As String, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim Sprite As Long, Colour As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim top As Long, Left As Long, AnimFrame As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Exit the sub if we're not in-game.
    If InGame = False Then Exit Sub
                
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
                
    ' Render Backdrop.
    Call RenderGraphic(Tex_GUI(BankE), 0, 0, D3DT_TEXTURE(Tex_GUI(BankE)).Width, D3DT_TEXTURE(Tex_GUI(BankE)).Height, 0, 0, 0, 0)
                
    ' Cycle through all the bank items.
    For i = 1 To MAX_BANK
        ' retrieve the item number.
        itemnum = GetBankItemNum(i)
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            ' And the icon.
            Sprite = Item(itemnum).Pic
            ' If the sprite is valid, continue.
            If Sprite > 0 And Sprite <= NumItems Then
                
                ' Calculate the position of the item icon and the animation frame we'll be using.
                top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                AnimFrame = ItemAnimFrame(itemnum) * PIC_X
                
                ' The Colors maaan!
                Red = Item(itemnum).Red
                Green = Item(itemnum).Green
                Blue = Item(itemnum).Blue
                Alpha = Item(itemnum).Alpha
                
                ' Render the item to the screen
                Call RenderGraphic(Tex_Item(Sprite), Left, top, PIC_X, PIC_Y, 0, 0, AnimFrame, 0, Red, Green, Blue, Alpha)

                ' If item is a stack - draw the amount you have
                If GetBankItemValue(i) > 1 Then
                    Y = top + 22
                    X = Left - 4
                
                    Amount = CStr(GetBankItemValue(i))
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    Call RenderText(MainFont, ConvertCurrency(Amount), X, Y, Colour)
                End If
            End If
            End If
        Next
        
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmMain.picBank.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picBank.ScaleHeight
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmMain.picBank.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picBank.ScaleHeight
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picBank.hWnd, ByVal 0)
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBank", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTradeOffer()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim Colour As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim top As Long, Left As Long, AnimFrame As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Exit the sub if we're not in-game.
    If InGame = False Then Exit Sub
                
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Cycle through the offered items.
    For i = 1 To MAX_INV
        ' Let's retrieve what item's offered in this slot.
        itemnum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        
        ' If the item number's valid we'll continue on.
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
        
            ' Retrieve the item picture and check if it's valid, if it is we can finaly get to rendering stuff!
            ' Exciting.
            itempic = Item(itemnum).Pic
            If itempic > 0 And itempic <= NumItems Then
                
                ' Calculate the location to render the texture on.
                top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                AnimFrame = ItemAnimFrame(itemnum) * PIC_X
                
                ' The Colors maaan!
                Red = Item(itemnum).Red
                Green = Item(itemnum).Green
                Blue = Item(itemnum).Blue
                Alpha = Item(itemnum).Alpha
                
                ' Render the actual icon.
                Call RenderGraphic(Tex_Item(itempic), Left, top, PIC_X, PIC_Y, 0, 0, AnimFrame, 0, Red, Green, Blue, Alpha)

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).value > 1 Then
                    Y = top + 22
                    X = Left - 4
                    
                    Amount = TradeYourOffer(i).value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        Colour = White
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        Colour = Yellow
                    ElseIf Amount > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    Call RenderText(MainFont, ConvertCurrency(Str(Amount)), X, Y, Colour)

                End If
            End If
        End If
    Next
        
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmMain.picYourTrade.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picYourTrade.ScaleHeight
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmMain.picYourTrade.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picYourTrade.ScaleHeight
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picYourTrade.hWnd, ByVal 0)
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTradeOffer", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTradeReceive()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As Long, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte
Dim Colour As Long
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim top As Long, Left As Long, AnimFrame As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Exit the sub if we're not in-game.
    If InGame = False Then Exit Sub
                
    ' Let's open clear ourselves a nice clean slate to render on shall we?
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0)
    Call D3DDevice8.BeginScene
    
    ' Cycle through the offered items.
    For i = 1 To MAX_INV
        ' Let's retrieve what item's offered in this slot.
        itemnum = TradeTheirOffer(i).num
        
        ' If the item number's valid we'll continue on.
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
        
            ' Retrieve the item picture and check if it's valid, if it is we can finaly get to rendering stuff!
            ' Exciting.
            itempic = Item(itemnum).Pic
            If itempic > 0 And itempic <= NumItems Then
                
                ' Calculate the location to render the texture on.
                top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                AnimFrame = ItemAnimFrame(itemnum) * PIC_X
                
                ' The Colors maaan!
                Red = Item(itemnum).Red
                Green = Item(itemnum).Green
                Blue = Item(itemnum).Blue
                Alpha = Item(itemnum).Alpha
                
                ' Render the actual icon.
                Call RenderGraphic(Tex_Item(itempic), Left, top, PIC_X, PIC_Y, 0, 0, AnimFrame, 0, Red, Green, Blue, Alpha)

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).value > 1 Then
                    Y = top + 22
                    X = Left - 4
                    
                    Amount = TradeTheirOffer(i).value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        Colour = White
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        Colour = Yellow
                    ElseIf Amount > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    Call RenderText(MainFont, ConvertCurrency(Str(Amount)), X, Y, Colour)

                End If
            End If
        End If
    Next
        
    ' We're done for now, so we can close the lovely little rendering device and present it to our user!
    ' Of course, we also need to do a few calculations to make sure it appears where it should.
    With srcRect
        .X1 = 0
        .X2 = frmMain.picTheirTrade.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picTheirTrade.ScaleHeight
    End With
    
    With destRect
        .X1 = 0
        .X2 = frmMain.picTheirTrade.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picTheirTrade.ScaleHeight
    End With
    
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(srcRect, destRect, frmMain.picTheirTrade.hWnd, ByVal 0)
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTradeReceive", "modGUI", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
