VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmEditor 
   AutoRedraw      =   -1  'True
   Caption         =   "Eclipse Dawn - Editor"
   ClientHeight    =   9000
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAttributes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5295
      ScaleWidth      =   4095
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Frame fraHeal 
         Caption         =   "Heal Player"
         Height          =   1575
         Left            =   120
         TabIndex        =   56
         Top             =   3720
         Width           =   3855
         Begin VB.ComboBox cmbHeal 
            Height          =   315
            ItemData        =   "frmEditor.frx":0000
            Left            =   1560
            List            =   "frmEditor.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtHealAmount 
            Height          =   285
            Left            =   1560
            TabIndex        =   57
            Text            =   "0"
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Restore:"
            Height          =   255
            Left            =   720
            TabIndex        =   60
            Top             =   530
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   720
            TabIndex        =   59
            Top             =   870
            Width           =   615
         End
      End
      Begin VB.Frame fraDamage 
         Caption         =   "Damage Player"
         Height          =   1575
         Left            =   120
         TabIndex        =   51
         Top             =   3720
         Width           =   3855
         Begin VB.ComboBox cmbDamage 
            Height          =   315
            ItemData        =   "frmEditor.frx":001C
            Left            =   1560
            List            =   "frmEditor.frx":0026
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtDamageAmount 
            Height          =   285
            Left            =   1560
            TabIndex        =   52
            Text            =   "0"
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Restore:"
            Height          =   255
            Left            =   720
            TabIndex        =   55
            Top             =   530
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   720
            TabIndex        =   54
            Top             =   870
            Width           =   615
         End
      End
      Begin VB.Frame fraSlide 
         Caption         =   "Slide Player"
         Height          =   1575
         Left            =   120
         TabIndex        =   48
         Top             =   3720
         Width           =   3855
         Begin VB.ComboBox cmbSlide 
            Height          =   315
            ItemData        =   "frmEditor.frx":0038
            Left            =   1320
            List            =   "frmEditor.frx":0048
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "Direction:"
            Height          =   255
            Left            =   360
            TabIndex        =   49
            Top             =   640
            Width           =   1095
         End
      End
      Begin VB.Frame fraWarp 
         Caption         =   "Warp Player"
         Height          =   1575
         Left            =   120
         TabIndex        =   41
         Top             =   3720
         Width           =   3855
         Begin VB.TextBox txtWarpY 
            Height          =   285
            Left            =   3120
            TabIndex        =   47
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtWarpX 
            Height          =   285
            Left            =   1200
            TabIndex        =   46
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cmbWarpMap 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label4 
            Caption         =   "Location Y:"
            Height          =   255
            Left            =   1920
            TabIndex        =   45
            Top             =   870
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Location X:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   870
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Warp to Map:"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   525
            Width           =   1095
         End
      End
      Begin VB.Frame fraBlock 
         Caption         =   "Block Player"
         Height          =   1575
         Left            =   120
         TabIndex        =   39
         Top             =   3720
         Width           =   3855
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "No additional information required."
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   3615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Attributes"
         Height          =   3735
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   3855
         Begin VB.OptionButton optShop 
            Caption         =   "Open Shop"
            Height          =   270
            Left            =   2160
            TabIndex        =   38
            Top             =   2520
            Width           =   1215
         End
         Begin VB.OptionButton optBank 
            Caption         =   "Open Bank"
            Height          =   270
            Left            =   2160
            TabIndex        =   37
            Top             =   2760
            Width           =   1215
         End
         Begin VB.OptionButton optItem 
            Caption         =   "Spawn Item"
            Height          =   255
            Left            =   2160
            TabIndex        =   36
            Top             =   1440
            Width           =   1550
         End
         Begin VB.OptionButton optResource 
            Caption         =   "Spawn Resource"
            Height          =   230
            Left            =   2160
            TabIndex        =   35
            Top             =   1680
            Width           =   1550
         End
         Begin VB.OptionButton optNpcSpawn 
            Caption         =   "Spawn NPC"
            Height          =   195
            Left            =   2160
            TabIndex        =   34
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optNpcAvoid 
            Caption         =   "Avoid Tile"
            Height          =   195
            Left            =   2160
            TabIndex        =   33
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optScript 
            Caption         =   "Script"
            Height          =   270
            Left            =   240
            TabIndex        =   32
            Top             =   3240
            Width           =   1215
         End
         Begin VB.OptionButton optKey 
            Caption         =   "Key"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   2760
            Width           =   1575
         End
         Begin VB.OptionButton optKeyOpen 
            Caption         =   "Key Open"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   3000
            Width           =   1575
         End
         Begin VB.OptionButton optDoor 
            Caption         =   "Door"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   2520
            Width           =   1575
         End
         Begin VB.OptionButton optTrap 
            Caption         =   "Damage Player"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton optBlocked 
            Caption         =   "Block Player"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optWarp 
            Caption         =   "Warp Player"
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optSlide 
            Caption         =   "Slide Player"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optHeal 
            Caption         =   "Heal Player"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton optDirBlock 
            Caption         =   "Directional Block"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Frame Frame6 
            Caption         =   "Player Functions"
            Height          =   1335
            Left            =   2040
            TabIndex        =   22
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Frame Frame5 
            Caption         =   "Map Control"
            Height          =   1335
            Left            =   120
            TabIndex        =   21
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Frame Frame4 
            Caption         =   "Spawn Objects"
            Height          =   975
            Left            =   2040
            TabIndex        =   20
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            Caption         =   "NPC Control"
            Height          =   855
            Left            =   2040
            TabIndex        =   19
            Top             =   240
            Width           =   1695
         End
         Begin VB.Frame Frame1 
            Caption         =   "Player Control"
            Height          =   1935
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1815
         End
      End
   End
   Begin VB.CommandButton cmdOpenChat 
      Caption         =   "Open Chat"
      Height          =   735
      Left            =   8160
      Picture         =   "frmEditor.frx":0066
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdEditDatabase 
      Caption         =   "Edit Database"
      Height          =   735
      Left            =   10200
      Picture         =   "frmEditor.frx":05D9
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1695
   End
   Begin VB.ListBox lstMapList 
      Columns         =   1
      Height          =   2400
      ItemData        =   "frmEditor.frx":0A73
      Left            =   120
      List            =   "frmEditor.frx":0A75
      TabIndex        =   13
      Top             =   6240
      Width           =   4095
   End
   Begin VB.PictureBox picRenameTile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   4335
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdRenameOK 
         Caption         =   "Rename"
         Height          =   315
         Left            =   2880
         TabIndex        =   12
         Top             =   30
         Width           =   1335
      End
      Begin VB.TextBox txtTileName 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "MapName"
         Top             =   50
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdReloadMap 
      Caption         =   "Reload Map"
      Height          =   735
      Left            =   1560
      Picture         =   "frmEditor.frx":0A77
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.ComboBox cmbTileSet 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename"
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cmbLayerSelect 
      Height          =   315
      ItemData        =   "frmEditor.frx":0EB1
      Left            =   5040
      List            =   "frmEditor.frx":0EC7
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox picLayerSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4440
      Picture         =   "frmEditor.frx":0F4E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdClearMap 
      Caption         =   "Clear Map"
      Height          =   735
      Left            =   3000
      Picture         =   "frmEditor.frx":1402
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveMap 
      Caption         =   "Save Map"
      Height          =   735
      Left            =   120
      Picture         =   "frmEditor.frx":18AC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.VScrollBar scrlTileSelect 
      Height          =   4695
      Left            =   3960
      Max             =   32000
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   8700
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   "Connected."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Connected."
            TextSave        =   "Connected."
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMapEditor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   16200
      Left            =   4440
      ScaleHeight     =   1080
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   28800
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmbLayerSelect_Change()
    If cmbLayerSelect.ListIndex + 1 < Layer_Count Then
        frmEditor.picAttributes.Visible = False
    Else
        frmEditor.picAttributes.Visible = True
    End If
End Sub

Private Sub cmbLayerSelect_Click()
    If cmbLayerSelect.ListIndex + 1 < Layer_Count Then
        frmEditor.picAttributes.Visible = False
    Else
        frmEditor.picAttributes.Visible = True
    End If
End Sub

Private Sub cmbTileSet_Change()
    ' Set the scrollbar Max.
    If D3DT_TEXTURE(Tex_TileSet(cmbTileSet.ListIndex + 1)).Height > TileSelectHeight Then
            frmEditor.scrlTileSelect.max = (D3DT_TEXTURE(Tex_TileSet(cmbTileSet.ListIndex + 1)).Height \ 16) - (TileSelectHeight \ 16)
    End If
End Sub

Private Sub cmbTileSet_Click()
' Set the scrollbar Max.
    If D3DT_TEXTURE(Tex_TileSet(cmbTileSet.ListIndex + 1)).Height > TileSelectHeight Then
            frmEditor.scrlTileSelect.max = (D3DT_TEXTURE(Tex_TileSet(cmbTileSet.ListIndex + 1)).Height \ 16) - (TileSelectHeight \ 16)
    End If
End Sub

Private Sub cmdClearMap_Click()
    If Editor.HasRight(CanEditMap) <> 1 Then
        ' No permissions
        MsgBox "Insufficient permissions, you are not allowed to edit maps.", vbInformation
        Exit Sub
    Else
        If Map.Revision < 1 Then Exit Sub
        If HasMapChanged = True Then
            If MsgBox("You've made changes to this map that have not been saved, are you sure you want to clear this map?", vbYesNo) = vbNo Then Exit Sub
        End If
        
        ReDim Map.Tile(Map.MaxX, Map.MaxY)
        HasMapChanged = True
    End If
End Sub

Private Sub cmdEditDatabase_Click()
    If Editor.HasRight(CanOpenDatabase) <> 1 Then
        ' No rights
        MsgBox "Insuficient permissions to access the database editor.", vbInformation
        Exit Sub
    Else
        
    End If
End Sub

Private Sub cmdReloadMap_Click()
    If Editor.HasRight(CanEditMap) <> 1 Then
        ' No permissions
        MsgBox "Insufficient permissions, you are not allowed to edit maps.", vbInformation
        Exit Sub
    Else
        ' Did we edit our current map? If so, we should prompt if the user really wants to load another before saving.
        If HasMapChanged = True Then
            If MsgBox("You've made changes to this map that have not been saved, are you sure you want to reload this map?", vbYesNo) = vbNo Then Exit Sub
        End If
            
        ' Send out a request for the map we want to edit.
        SendRequestMap CurrentMap
            
    End If
End Sub

Private Sub cmdRename_Click()
    picRenameTile.Visible = True
    txtTileName.text = Trim$(Options.TileSetName(cmbTileSet.ListIndex + 1))
End Sub

Private Sub cmdRenameOK_Click()
Dim TempIndex As Long
    TempIndex = cmbTileSet.ListIndex
    cmbTileSet.RemoveItem TempIndex
    cmbTileSet.AddItem Trim$(txtTileName.text), TempIndex
    cmbTileSet.Refresh
    cmbTileSet.ListIndex = TempIndex
    
    Options.TileSetName(TempIndex + 1) = Trim$(txtTileName.text)
    SaveOptions App.Path & "\" & OPTIONS_FILE
    picRenameTile.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Check if we're clicking on the tile selection portion.
    If X > RecTileSelectWindow.X1 And X < RecTileSelectWindow.X2 And Y > RecTileSelectWindow.Y1 And Y < RecTileSelectWindow.Y2 Then
        X = X - TileSetWindowOffSetX
        Y = (Y - TileSetWindowOffSetY) + (scrlTileSelect.value * 32)
        MapEditorChooseTile Button, X, Y
    End If
    
    ' Inside the map editor view.
    If X > MapViewWindow.X1 And X < MapViewWindow.X2 And Y > MapViewWindow.Y1 And Y < MapViewWindow.Y2 Then
        X = X - 16
        Y = Y - 16
        CurX = ((X - MapViewWindowOffSetX) / PIC_X) - (0 + MapViewTileOffSetX)
        CurY = ((Y - MapViewWindowOffSetY) / PIC_Y) - (0 + MapViewTileOffSetY)

        MapEditorMouseDown Button, X, Y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim OldMouseX As Long, OldMouseY As Long
    ' Check if we're moving on the tile selection portion.
    If X > RecTileSelectWindow.X1 And X < RecTileSelectWindow.X2 And Y > RecTileSelectWindow.Y1 And Y < RecTileSelectWindow.Y2 Then
        X = X - TileSetWindowOffSetX
        Y = (Y - TileSetWindowOffSetY) + (scrlTileSelect.value * 32)
        Call MapEditorDrag(Button, X, Y)
    End If
    
    ' Inside the map editor view.
    If X > MapViewWindow.X1 And X < MapViewWindow.X2 And Y > MapViewWindow.Y1 And Y < MapViewWindow.Y2 Then
        ShowMouse = True
        
        OldMouseX = MouseX
        OldMouseY = MouseY
        MouseX = X - MapViewWindow.X1 - 16
        MouseY = Y - MapViewWindow.Y1 - 16
        ' Loads of code to try and slow down the incredibly fast tile movement.
        If MouseX > OldMouseX Then MouseXMove = MouseXMove + 1
        If MouseY > OldMouseY Then MouseYMove = MouseYMove + 1
        If MouseX < OldMouseX Then MouseXMove = MouseXMove - 1
        If MouseY < OldMouseY Then MouseYMove = MouseYMove - 1
        
        If Button = vbMiddleButton Then
            ' Set the offsets if the mouse has moved.
            If MouseXMove >= 2 Then MapViewTileOffSetX = MapViewTileOffSetX + 1: MouseXMove = 0
            If MouseYMove >= 2 Then MapViewTileOffSetY = MapViewTileOffSetY + 1: MouseYMove = 0
            If MouseXMove <= (-2) Then MapViewTileOffSetX = MapViewTileOffSetX - 1: MouseXMove = 0
            If MouseYMove <= (-2) Then MapViewTileOffSetY = MapViewTileOffSetY - 1: MouseYMove = 0
            '  Adjust the view window so we render the correct stuff.
            ' UpdateCamera
        End If
        
        ' Tile dragging
        X = X - 16
        Y = Y - 16
        CurX = ((X - MapViewWindowOffSetX) / PIC_X) - (0 + MapViewTileOffSetX)
        CurY = ((Y - MapViewWindowOffSetY) / PIC_Y) - (0 + MapViewTileOffSetY)

        MapEditorMouseDown Button, X, Y
    Else
        ShowMouse = False
    End If
End Sub

Private Sub Form_Resize()
    If frmEditor.Width < 12120 Then frmEditor.Width = 12120
    If frmEditor.Height < 9465 Then frmEditor.Height = 9465
    
    ' Resize the map selection list.
    lstMapList.Height = 140
    lstMapList.top = frmEditor.ScaleHeight - 155
    
    ' Resize the tile selection screen.
    With RecTileSelectWindow
        .X1 = TileSetWindowOffSetX
        .X2 = .X1 + TEXTURE_WIDTH
        .Y1 = TileSetWindowOffSetY
        .Y2 = .Y1 + frmEditor.ScaleHeight - lstMapList.Height - 122
        TileSelectHeight = (.Y2 - .Y1)
    End With
    scrlTileSelect.Height = TileSelectHeight + 1
    frmEditor.picAttributes.Height = TileSelectHeight + 43
    
    ' Resize the map editor view
    MapViewHeight = frmEditor.ScaleHeight - MapViewWindowOffSetY - 23
    MapViewWidth = frmEditor.ScaleWidth - MapViewWindowOffSetX - 6
    With MapViewWindow
        .X1 = MapViewWindowOffSetX
        .X2 = .X1 + MapViewWidth
        .Y1 = MapViewWindowOffSetY
        .Y2 = .Y1 + MapViewHeight
    End With
    
    ' Relocate a few buttons
    cmdEditDatabase.Left = frmEditor.ScaleWidth - cmdEditDatabase.Width - 8
    cmdOpenChat.Left = cmdEditDatabase.Left - 136
    
    '  update render view
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyEditor
End Sub

Private Sub lstMapList_Click()
    If Editor.HasRight(CanEditMap) <> 1 Then
        ' No permissions
        MsgBox "Insufficient permissions, you are not allowed to edit maps.", vbInformation
        Exit Sub
    Else
        ' Make sure we're not trying to load the same map.
        If (lstMapList.ListIndex + 1) <> CurrentMap Then
            ' Did we edit our current map? If so, we should prompt if the user really wants to load another before saving.
            If HasMapChanged = True Then
                If MsgBox("You've made changes to this map that have not been saved, are you sure you want to load a different one?", vbYesNo) = vbNo Then Exit Sub
            End If
            
            ' Set our current map
            CurrentMap = lstMapList.ListIndex + 1
            
            ' Send out a request for the map we want to edit.
            SendRequestMap CurrentMap
            
        End If
    End If
End Sub

Private Sub optBlocked_Click()
    ClearAttributeFrames
    fraBlock.Visible = True
End Sub

Private Sub optHeal_Click()
    ClearAttributeFrames
    fraHeal.Visible = True
End Sub

Private Sub optSlide_Click()
    ClearAttributeFrames
    fraSlide.Visible = True
End Sub

Private Sub optTrap_Click()
    ClearAttributeFrames
    fraDamage.Visible = True
End Sub

Private Sub optWarp_Click()
    ClearAttributeFrames
    fraWarp.Visible = True
End Sub

Private Sub scrlTileSelect_Change()
    MapEditor_DrawTileSet
End Sub

Private Sub scrlTileSelect_Scroll()
    MapEditor_DrawTileSet
End Sub

Private Sub txtHealAmount_Change()
    If Val(txtHealAmount.text) < 0 Then txtHealAmount.text = "0"
End Sub

Private Sub txtWarpX_Change()
    If Val(txtWarpX.text) < 0 Then txtWarpX.text = "0"
End Sub

Private Sub txtWarpY_Change()
    If CLng(txtWarpY.text) < 0 Then txtWarpY.text = "0"
End Sub
