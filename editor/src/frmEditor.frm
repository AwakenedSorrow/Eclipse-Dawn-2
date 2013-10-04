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
   Begin VB.CommandButton cmdOpenChat 
      Caption         =   "Open Chat"
      Height          =   735
      Left            =   8160
      Picture         =   "frmEditor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdEditDatabase 
      Caption         =   "Edit Database"
      Height          =   735
      Left            =   10200
      Picture         =   "frmEditor.frx":0573
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1695
   End
   Begin VB.ListBox lstMapList 
      Columns         =   1
      Height          =   2400
      ItemData        =   "frmEditor.frx":0A0D
      Left            =   120
      List            =   "frmEditor.frx":0A0F
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
      Picture         =   "frmEditor.frx":0A11
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
      ItemData        =   "frmEditor.frx":0E4B
      Left            =   5040
      List            =   "frmEditor.frx":0E5E
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
      Picture         =   "frmEditor.frx":0ED9
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
      Picture         =   "frmEditor.frx":138D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveMap 
      Caption         =   "Save Map"
      Height          =   735
      Left            =   120
      Picture         =   "frmEditor.frx":1837
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
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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

Private Sub cmdEditDatabase_Click()
    If Editor.HasRight(CanOpenDatabase) <> 1 Then
        ' No rights
        MsgBox "Insuficient permissions to access the database editor.", vbInformation
        Exit Sub
    Else
        
    End If
End Sub

Private Sub cmdRename_Click()
    picRenameTile.Visible = True
    txtTileName.Text = Trim$(Options.TileSetName(cmbTileSet.ListIndex + 1))
End Sub

Private Sub cmdRenameOK_Click()
Dim TempIndex As Long
    TempIndex = cmbTileSet.ListIndex
    cmbTileSet.RemoveItem TempIndex
    cmbTileSet.AddItem Trim$(txtTileName.Text), TempIndex
    cmbTileSet.Refresh
    cmbTileSet.ListIndex = TempIndex
    
    Options.TileSetName(TempIndex + 1) = Trim$(txtTileName.Text)
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
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Check if we're moving on the tile selection portion.
    If X > RecTileSelectWindow.X1 And X < RecTileSelectWindow.X2 And Y > RecTileSelectWindow.Y1 And Y < RecTileSelectWindow.Y2 Then
        X = X - TileSetWindowOffSetX
        Y = (Y - TileSetWindowOffSetY) + (scrlTileSelect.value * 32)
        Call MapEditorDrag(Button, X, Y)
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyEditor
End Sub

Private Sub scrlTileSelect_Change()
    MapEditor_DrawTileSet
End Sub

Private Sub scrlTileSelect_Scroll()
    MapEditor_DrawTileSet
End Sub
