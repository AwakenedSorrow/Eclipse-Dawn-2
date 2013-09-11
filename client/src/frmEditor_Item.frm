VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13095
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtStatReq 
      Height          =   270
      Index           =   2
      Left            =   7920
      TabIndex        =   62
      Text            =   "0"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtStatReq 
      Height          =   270
      Index           =   1
      Left            =   7920
      TabIndex        =   61
      Text            =   "0"
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Graphics"
      Height          =   2535
      Left            =   9720
      TabIndex        =   42
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cmbAnimation 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2040
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAlpha 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   56
         Top             =   1680
         Width           =   1335
      End
      Begin VB.HScrollBar scrlBlue 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   54
         Top             =   1320
         Width           =   1335
      End
      Begin VB.HScrollBar scrlGreen 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   52
         Top             =   960
         Width           =   1335
      End
      Begin VB.HScrollBar scrlRed 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   50
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2650
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   44
         Top             =   840
         Width           =   480
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation:"
         Height          =   180
         Left            =   120
         TabIndex        =   58
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label lblAlpha 
         AutoSize        =   -1  'True
         Caption         =   "Alpha: 255"
         Height          =   180
         Left            =   120
         TabIndex        =   57
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label lblBlue 
         AutoSize        =   -1  'True
         Caption         =   "Blue: 255"
         Height          =   180
         Left            =   120
         TabIndex        =   55
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   750
      End
      Begin VB.Label lblGreen 
         AutoSize        =   -1  'True
         Caption         =   "Green: 255"
         Height          =   180
         Left            =   120
         TabIndex        =   53
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   870
      End
      Begin VB.Label lblRed 
         AutoSize        =   -1  'True
         Caption         =   "Red: 255"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   705
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   45
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   2175
      Left            =   5040
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtSpeed 
         Height          =   270
         Left            =   5400
         TabIndex        =   78
         Text            =   "1000"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtAddStat 
         Height          =   270
         Index           =   5
         Left            =   5400
         TabIndex        =   77
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtAddStat 
         Height          =   270
         Index           =   4
         Left            =   5400
         TabIndex        =   76
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtAddStat 
         Height          =   270
         Index           =   3
         Left            =   2280
         TabIndex        =   75
         Text            =   "0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtAddStat 
         Height          =   270
         Index           =   2
         Left            =   2280
         TabIndex        =   74
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtAddStat 
         Height          =   270
         Index           =   1
         Left            =   2280
         TabIndex        =   73
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtDamage 
         Height          =   270
         Left            =   1320
         TabIndex        =   72
         Text            =   "0"
         Top             =   720
         Width           =   1815
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5450
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   35
         Top             =   240
         Width           =   720
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   1320
         List            =   "frmEditor_Item.frx":3342
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblPaperdoll 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   3240
         TabIndex        =   33
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Weapon Speed (ms):"
         Height          =   180
         Left            =   3240
         TabIndex        =   28
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "Bonus Willpower:"
         Height          =   180
         Index           =   5
         Left            =   3240
         TabIndex        =   27
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "Bonus Agility:"
         Height          =   180
         Index           =   4
         Left            =   3240
         TabIndex        =   26
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "Bonus Inteligence:"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "Bonus Endurance:"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "Bonus Strength: "
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   2535
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtDesc 
         Height          =   2000
         Left            =   3480
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   70
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbRarity 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3363
         Left            =   1440
         List            =   "frmEditor_Item.frx":3379
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtPrice 
         Height          =   270
         Left            =   1440
         TabIndex        =   66
         Text            =   "0"
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":33AC
         Left            =   1440
         List            =   "frmEditor_Item.frx":33B9
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":33E2
         Left            =   1440
         List            =   "frmEditor_Item.frx":3404
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   3480
         TabIndex        =   71
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Item Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Quality:"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Base Value:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   1095
      Left            =   3360
      TabIndex        =   5
      Top             =   2760
      Width           =   9615
      Begin VB.TextBox txtLevelReq 
         Height          =   270
         Left            =   7800
         TabIndex        =   69
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbAccess 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3456
         Left            =   1920
         List            =   "frmEditor_Item.frx":3469
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtStatReq 
         Height          =   270
         Index           =   5
         Left            =   7800
         TabIndex        =   65
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtStatReq 
         Height          =   270
         Index           =   4
         Left            =   6240
         TabIndex        =   64
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtStatReq 
         Height          =   270
         Index           =   3
         Left            =   6240
         TabIndex        =   63
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   7200
         TabIndex        =   49
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   960
         TabIndex        =   48
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req:"
         Height          =   180
         Left            =   960
         TabIndex        =   47
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str:"
         Height          =   180
         Index           =   1
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   285
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End:"
         Height          =   180
         Index           =   2
         Left            =   4080
         TabIndex        =   9
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int:"
         Height          =   180
         Index           =   3
         Left            =   5760
         TabIndex        =   8
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   285
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi:"
         Height          =   180
         Index           =   4
         Left            =   5760
         TabIndex        =   7
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   315
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will:"
         Height          =   180
         Index           =   5
         Left            =   7320
         TabIndex        =   6
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9120
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6180
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   2055
      Left            =   6600
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ComboBox cmbSpellCast 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtAddExp 
         Height          =   270
         Left            =   2040
         TabIndex        =   81
         Text            =   "0"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtAddMP 
         Height          =   270
         Left            =   2040
         TabIndex        =   80
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtAddHP 
         Height          =   270
         Left            =   2040
         TabIndex        =   79
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instant Cast"
         Height          =   255
         Left            =   1920
         TabIndex        =   39
         Top             =   1695
         Width           =   1335
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "Cast Spell:"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Gain EXP:"
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   765
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Recover MP:"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   960
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Recover HP: "
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   990
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   6840
      TabIndex        =   31
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
      Begin VB.ComboBox cmbSpell 
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Teaches Spell:"
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub chkInstant_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).instaCast = chkInstant.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkInstant_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).AccessReq = cmbAccess.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbAccess_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbAnimation_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).Animation = cmbAnimation.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbAnimation_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbRarity_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).Rarity = cmbRarity.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbRarity_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSpell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).Data1 = cmbSpell.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbSpell_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub cmbSpellCast_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).CastSpell = cmbSpellCast.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbSpellCast_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPic.max = NumItems
    scrlPaperdoll.max = NumPaperdolls
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ItemTypeWeapon) And (cmbType.ListIndex <= ItemTypeShield) Then
        fraEquipment.Visible = True
        
        ' Make sure we have a few things set so nothing is out of the ordinary.
        If cmbType.ListIndex = ItemTypeWeapon Then
            frmEditor_Item.lblDamage.Caption = "Damage:"
            frmEditor_Item.cmbTool.Enabled = True
        ElseIf cmbType.ListIndex = ItemTypeHelmet Then
            frmEditor_Item.lblDamage.Caption = "Defense:"
            frmEditor_Item.cmbTool.Enabled = False
        ElseIf cmbType.ListIndex = ItemTypeArmor Then
            frmEditor_Item.lblDamage.Caption = "Defense:"
            frmEditor_Item.cmbTool.Enabled = False
        ElseIf cmbType.ListIndex = ItemTypeShield Then
            frmEditor_Item.lblDamage.Caption = "Block:"
            frmEditor_Item.cmbTool.Enabled = False
        End If
    Else
        fraEquipment.Visible = False
    End If

    If cmbType.ListIndex = ItemTypeConsume Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ItemTypeSpell) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ItemTypeScripted) Then
        fraEquipment.Visible = False
        fraVitals.Visible = False
        fraSpell.Visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAlpha_Change()
    lblAlpha.Caption = "Alpha: " & Trim$(Str$(scrlAlpha.value))
    Item(EditorIndex).Alpha = scrlAlpha.value
End Sub

Private Sub scrlBlue_Change()
    lblBlue.Caption = "Blue: " & Trim$(Str$(scrlBlue.value))
    Item(EditorIndex).Blue = scrlBlue.value
End Sub

Private Sub scrlGreen_Change()
    lblGreen.Caption = "Green: " & Trim$(Str$(scrlGreen.value))
    Item(EditorIndex).Green = scrlGreen.value
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.value
    Call EditorItem_DrawPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.value
    Item(EditorIndex).Pic = scrlPic.value
    Call EditorItem_DrawItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRed_Change()
    lblRed.Caption = "Red: " & Trim$(Str$(scrlRed.value))
    Item(EditorIndex).Red = scrlRed.value
End Sub

Private Sub txtAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).AddEXP = Val(Trim$(txtAddExp.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAddHP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).AddHP = Val(Trim$(txtAddHP.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAddMP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).AddMP = Val(Trim$(txtAddMP.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAddStat_Change(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Val(Trim$(txtAddStat(index).text)) < 0 Then txtAddStat(index).text = "0"
    If Val(Trim$(txtAddStat(index).text)) > 255 Then txtAddStat(index).text = "255"
    
    Item(EditorIndex).Add_Stat(index) = Val(Trim$(txtAddStat(index).text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAddStat_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).Data2 = Val(Trim$(txtDamage.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
   
    ' Check if we're out of bounds, if we are just set it back to the nearest value.
    If Val(Trim$(txtLevelReq.text)) < 0 Then txtLevelReq.text = 0
    If Val(Trim$(txtLevelReq.text)) > MAX_LEVELS Then txtLevelReq.text = MAX_LEVELS
   
    Item(EditorIndex).LevelReq = Val(Trim$(txtLevelReq.text))
   
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).name = txtName.text
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtPrice_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).Price = Val(Trim$(txtPrice.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Item(EditorIndex).Speed = Val(Trim$(txtSpeed.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtStatReq_Change(index As Integer)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check if we're not going out of bounds, and if we are set it back to amounts we can handle.
    If Val(Trim$(txtStatReq(index).text)) < 0 Then txtStatReq(index).text = "0"
    If Val(Trim$(txtStatReq(index).text)) > 255 Then txtStatReq(index).text = "255"
    
    Item(EditorIndex).Stat_Req(index) = Val(Trim$(txtStatReq(index).text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
