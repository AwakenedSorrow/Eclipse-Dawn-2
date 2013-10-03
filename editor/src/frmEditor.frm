VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmEditor 
   Caption         =   "Eclipse Dawn - Editor"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTileSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   240
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   12
      Top             =   1440
      Width           =   3855
   End
   Begin VB.ComboBox cmbLayerSelect 
      Height          =   315
      ItemData        =   "frmEditor.frx":0000
      Left            =   5040
      List            =   "frmEditor.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   11
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
      Picture         =   "frmEditor.frx":008E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdClearMap 
      Caption         =   "Clear Changes"
      Height          =   735
      Left            =   1320
      Picture         =   "frmEditor.frx":0542
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveMap 
      Caption         =   "Save Map"
      Height          =   735
      Left            =   120
      Picture         =   "frmEditor.frx":09EC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox lstMapList 
      Columns         =   2
      Height          =   1815
      ItemData        =   "frmEditor.frx":0F19
      Left            =   240
      List            =   "frmEditor.frx":0F1B
      TabIndex        =   7
      Top             =   6480
      Width           =   4070
   End
   Begin VB.Frame frmMapList 
      Caption         =   "Map List"
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   4305
   End
   Begin VB.VScrollBar scrlTileSelect 
      Height          =   4695
      Left            =   4080
      Max             =   2
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.Frame frmTileSelect 
      Caption         =   "Tile Selector"
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4305
      Begin VB.CommandButton cmdRename 
         Caption         =   "Rename"
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbTileSet 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
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
      Height          =   5535
      Left            =   4440
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   0
      Top             =   720
      Width           =   7575
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Resize()
    If frmEditor.Width < 12120 Then frmEditor.Width = 12120
    If frmEditor.Height < 9465 Then frmEditor.Height = 9465
    
    ' Resize the map editor screen.
    picMapEditor.Width = frmEditor.ScaleWidth - frmTileSelect.Width - 10
    picMapEditor.Height = frmEditor.ScaleHeight - 47 - stBar.Height - 3
    
    ' Resize the tile selection screen.
    frmTileSelect.Height = (frmEditor.ScaleHeight * 0.75) - 47 - stBar.Height - 3
    scrlTileSelect.Height = (frmEditor.ScaleHeight * 0.75) - 47 - stBar.Height - 58
    picTileSelect.Height = (frmEditor.ScaleHeight * 0.75) - 47 - stBar.Height - 58
    
    ' Resize the map selection list.
    frmMapList.top = frmTileSelect.top + frmTileSelect.Height
    frmMapList.Height = (frmEditor.ScaleHeight * 0.25)
    lstMapList.top = frmMapList.top + (frmEditor.ScaleHeight * 0.25) / 14
    lstMapList.Height = (frmEditor.ScaleHeight * 0.25) - (frmEditor.ScaleHeight * 0.25) / 9.9
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyEditor
End Sub

Private Sub scrlTileSelect_Change()
    ' MapEditor_DrawTileset
End Sub
