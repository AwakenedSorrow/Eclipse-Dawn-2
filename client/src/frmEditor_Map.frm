VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14655
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   977
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picAttributes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   7440
      ScaleHeight     =   7215
      ScaleWidth      =   7095
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame fraScript 
         Caption         =   "Script"
         Height          =   1455
         Left            =   1800
         TabIndex        =   89
         Top             =   3000
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtScript 
            Height          =   270
            Left            =   360
            TabIndex        =   91
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmdScript 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   90
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraSlide 
         Caption         =   "Slide"
         Height          =   1455
         Left            =   1800
         TabIndex        =   81
         Top             =   2880
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbSlide 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3332
            Left            =   240
            List            =   "frmEditor_Map.frx":3342
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdSlide 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   82
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraTrap 
         Caption         =   "Trap"
         Height          =   1575
         Left            =   1800
         TabIndex        =   77
         Top             =   2880
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlTrap 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   79
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdTrap 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   78
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblTrap 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraHeal 
         Caption         =   "Heal"
         Height          =   1815
         Left            =   1800
         TabIndex        =   72
         Top             =   2640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbHeal 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":335D
            Left            =   240
            List            =   "frmEditor_Map.frx":3367
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdHeal 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   74
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlHeal 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   73
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblHeal 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   2655
         Left            =   1800
         TabIndex        =   36
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNpc 
            Height          =   780
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   38
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   37
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblNpcDir 
            Caption         =   "Direction: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraResource 
         Caption         =   "Object"
         Height          =   1695
         Left            =   1800
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   33
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   32
            Top             =   600
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblResource 
            Caption         =   "Object:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Map Warp"
         Height          =   2775
         Left            =   1800
         TabIndex        =   54
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   61
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   56
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label lblMapWarpY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblMapWarpX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblMapWarp 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         Height          =   1335
         Left            =   1920
         TabIndex        =   62
         Top             =   3000
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdShop 
            Caption         =   "Accept"
            Height          =   375
            Left            =   960
            TabIndex        =   64
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraKeyOpen 
         Caption         =   "Key Open"
         Height          =   2055
         Left            =   1800
         TabIndex        =   48
         Top             =   2760
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdKeyOpen 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   53
            Top             =   1440
            Width           =   1215
         End
         Begin VB.HScrollBar scrlKeyY 
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlKeyX 
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblKeyY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblKeyX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame fraMapKey 
         Caption         =   "Map Key"
         Height          =   1815
         Left            =   1800
         TabIndex        =   44
         Top             =   2880
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbMapKey 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   600
            Width           =   2175
         End
         Begin VB.PictureBox picMapKey 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   47
            Top             =   600
            Width           =   480
         End
         Begin VB.CommandButton cmdMapKey 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   46
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkMapKey 
            Caption         =   "Take key away upon use."
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Value           =   1  'Checked
            Width           =   2535
         End
      End
      Begin VB.Frame fraMapItem 
         Caption         =   "Map Item"
         Height          =   1815
         Left            =   1800
         TabIndex        =   41
         Top             =   2640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtMapItemVal 
            Height          =   270
            Left            =   960
            TabIndex        =   87
            Text            =   "0"
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox cmbMapItem 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   480
            Width           =   2535
         End
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1200
            TabIndex        =   43
            Top             =   1200
            Width           =   1215
         End
         Begin VB.PictureBox picMapItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   42
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label2 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   840
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type"
      Height          =   1095
      Left            =   5760
      TabIndex        =   25
      Top             =   5760
      Width           =   1455
      Begin VB.OptionButton optBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "Block"
         Height          =   255
         Left            =   480
         TabIndex        =   67
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optAttribs 
         Alignment       =   1  'Right Justify
         Caption         =   "Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optLayers 
         Alignment       =   1  'Right Justify
         Caption         =   "Layers"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   5295
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   120
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   14
      Top             =   120
      Width           =   5280
      Begin VB.PictureBox picBackSelect 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   0
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   15
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   5295
      Left            =   5400
      Max             =   255
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame fraTileSet 
      Caption         =   "Tileset: 0"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   5535
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      Height          =   5535
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optScript 
         Caption         =   "Script"
         Height          =   270
         Left            =   120
         TabIndex        =   88
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optSlide 
         Caption         =   "Slide"
         Height          =   270
         Left            =   120
         TabIndex        =   71
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optTrap 
         Caption         =   "Trap"
         Height          =   270
         Left            =   120
         TabIndex        =   70
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Heal"
         Height          =   270
         Left            =   120
         TabIndex        =   69
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Bank"
         Height          =   270
         Left            =   120
         TabIndex        =   68
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Shop"
         Height          =   270
         Left            =   120
         TabIndex        =   65
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   270
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optDoor 
         Caption         =   "Door"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         Caption         =   "Resource"
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optKeyOpen 
         Caption         =   "Key Open"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Blocked"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Warp"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Npc Avoid"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optKey 
         Caption         =   "Key"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      Height          =   5535
      Left            =   5760
      TabIndex        =   16
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   390
         Left            =   120
         TabIndex        =   20
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4560
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Drag mouse to select multiple tiles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   5760
      Width           =   5535
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbMapItem_click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Item(cmbMapItem.ListIndex + 1).Type <> ItemTypeCurrency Then
        txtMapItemVal.Enabled = False
        txtMapItemVal.text = "0"
    Else
        txtMapItemVal.Enabled = True
        txtMapItemVal.text = "0"
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdMapItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdHeal_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorHealType = cmbHeal.ListIndex + 1
    MapEditorHealAmount = scrlHeal.value
    picAttributes.Visible = False
    fraHeal.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdKeyOpen_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    KeyOpenEditorX = scrlKeyX.value
    KeyOpenEditorY = scrlKeyY.value
    picAttributes.Visible = False
    fraKeyOpen.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdKeyOpen_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdMapItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorNum = cmbMapItem.ListIndex + 1
    ItemEditorValue = Val(txtMapItemVal.text)
    picAttributes.Visible = False
    fraMapItem.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdMapItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdMapKey_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    KeyEditorNum = cmbMapKey.ListIndex + 1
    KeyEditorTake = chkMapKey.value
    picAttributes.Visible = False
    fraMapKey.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdMapKey_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdMapWarp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EditorWarpMap = scrlMapWarp.value
    EditorWarpX = scrlMapWarpX.value
    EditorWarpY = scrlMapWarpY.value
    picAttributes.Visible = False
    fraMapWarp.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdMapWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdNpcSpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SpawnNpcNum = lstNpc.ListIndex + 1
    SpawnNpcDir = scrlNpcDir.value
    picAttributes.Visible = False
    fraNpcSpawn.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdResourceOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ResourceEditorNum = scrlResource.value
    picAttributes.Visible = False
    fraResource.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdResourceOk_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EditorShopNum = cmbShop.ListIndex
    picAttributes.Visible = False
    fraShop.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSlide_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorSlideDir = cmbSlide.ListIndex
    picAttributes.Visible = False
    fraSlide.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdTrap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorHealAmount = scrlTrap.value
    picAttributes.Visible = False
    fraTrap.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdScript_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorScript = Val(txtScript.text)
    picAttributes.Visible = False
    fraScript.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdScript_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' move the entire attributes box on screen
    picAttributes.Left = 8
    picAttributes.top = 8
    frmEditor_Map.scrlTileSet.max = NumTileSets
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optDoor_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    
    scrlMapWarp.max = MAX_MAPS
    scrlMapWarp.value = 1
    scrlMapWarpX.max = MAX_BYTE
    scrlMapWarpY.max = MAX_BYTE
    scrlMapWarpX.value = 0
    scrlMapWarpY.value = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optDoor_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optHeal_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraHeal.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optLayers_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If optLayers.value Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optLayers_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optAttribs_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If optAttribs.value Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optAttribs_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optNpcSpawn_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lstNpc.Clear
    
    For n = 1 To MAX_MAP_NPCS
        If Map.Npc(n) > 0 Then
            lstNpc.AddItem n & ": " & Npc(Map.Npc(n)).name
        Else
            lstNpc.AddItem n & ": No Npc"
        End If
    Next n
    
    scrlNpcDir.value = 0
    lstNpc.ListIndex = 0
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraNpcSpawn.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraResource.Visible = True
    
    lblResource.Caption = "Resource #" & Trim(Str$(scrlResource.value)) & ": " & Resource(scrlResource.value).name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optResource_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optScript_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    txtScript.text = "0"
    picAttributes.Visible = True
    fraScript.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optScript_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraShop.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSlide_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraSlide.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optTrap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraTrap.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBack_Click()

End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    X = X + (frmEditor_Map.scrlPictureX.value * 32)
    Y = Y + (frmEditor_Map.scrlPictureY.value * 32)
    Call MapEditorChooseTile(Button, X, Y)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBackSelect_MouseDown", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
 
Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorDrag(Button, X, Y)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBackSelect_MouseMove", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSend_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorSend
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSend_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdProperties_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Load frmEditor_MapProperties
    MapEditorProperties
    frmEditor_MapProperties.Show vbModal
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdProperties_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optWarp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    
    scrlMapWarp.max = MAX_MAPS
    scrlMapWarp.value = 1
    scrlMapWarpX.max = MAX_BYTE
    scrlMapWarpY.max = MAX_BYTE
    scrlMapWarpX.value = 0
    scrlMapWarpY.value = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapItem.Visible = True
    
    cmbMapItem.ListIndex = 0
    txtMapItemVal.text = "0"
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optKey_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapKey.Visible = True
    
    cmbMapKey.ListIndex = 0
    chkMapKey.value = 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optKey_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optKeyOpen_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    fraKeyOpen.Visible = True
    picAttributes.Visible = True
    
    scrlKeyX.max = Map.MaxX
    scrlKeyY.max = Map.MaxY
    scrlKeyX.value = 0
    scrlKeyY.value = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optKeyOpen_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdFill_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorFillLayer
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdFill_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdClear_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorClearLayer
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdClear_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdClear2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorClearAttribs
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdClear2_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHeal_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblHeal.Caption = "Amount: " & scrlHeal.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHeal_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlKeyX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblKeyX.Caption = "X: " & scrlKeyX.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlKeyX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlKeyX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlKeyX_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlKeyX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlKeyY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblKeyY.Caption = "Y: " & scrlKeyY.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlKeyY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlKeyY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlKeyY_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlKeyY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTrap_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblTrap.Caption = "Amount: " & scrlTrap.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTrap_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMapWarp.Caption = "Map: " & scrlMapWarp.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarp_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarp_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlMapWarp_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarp_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMapWarpX.Caption = "X: " & scrlMapWarpX.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlMapWarpX_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMapWarpY.Caption = "Y: " & scrlMapWarpY.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlMapWarpY_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNpcDir_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case scrlNpcDir.value
        Case South
            lblNpcDir = "Direction: Down"
        Case North
            lblNpcDir = "Direction: Up"
        Case West
            lblNpcDir = "Direction: Left"
        Case East
            lblNpcDir = "Direction: Right"
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNpcDir_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNpcDir_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlNpcDir_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNpcDir_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResource_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblResource.Caption = "Resource #" & Trim(Str$(scrlResource.value)) & ": " & Resource(scrlResource.value).name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlResource_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResource_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlResource_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlResource_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorTileScroll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPictureX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorTileScroll
    ' Make sure the tileset is drawn as it is being scrolled through.
    Call EditorMap_DrawTileset
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPictureY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPictureY_Change
    ' Make sure the tileset is drawn as it is being scrolled through.
    Call EditorMap_DrawTileset
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPictureX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPictureY_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPictureY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTileSet_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    fraTileSet.Caption = "Tileset: " & scrlTileSet.value
    
    frmEditor_Map.scrlPictureY.max = (frmEditor_Map.picBackSelect.Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (frmEditor_Map.picBackSelect.Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    
    MapEditorTileScroll
    
    frmEditor_Map.picBackSelect.Left = 0
    frmEditor_Map.picBackSelect.top = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTileSet_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTileSet_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlTileSet_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTileSet_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMapItemVal_Change()

    If Val(txtMapItemVal.text) < 0 Then txtMapItemVal.text = "0"

End Sub

