VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
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
   Icon            =   "frmEditor_Resource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5760
      TabIndex        =   23
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8040
      TabIndex        =   22
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resource Properties"
      Height          =   6855
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtSpawn 
         Height          =   270
         Left            =   4800
         TabIndex        =   47
         Text            =   "0"
         Top             =   6480
         Width           =   1215
      End
      Begin VB.TextBox txtHealth 
         Height          =   270
         Left            =   1800
         TabIndex        =   46
         Text            =   "0"
         Top             =   6480
         Width           =   1215
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3332
         Left            =   4080
         List            =   "frmEditor_Resource.frx":3342
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   5760
         Width           =   1935
      End
      Begin VB.ComboBox cmbItem 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   6120
         Width           =   1815
      End
      Begin VB.ComboBox cmbAnimation 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   5760
         Width           =   1815
      End
      Begin VB.TextBox txtAlpha 
         Height          =   270
         Index           =   1
         Left            =   5280
         TabIndex        =   42
         Text            =   "255"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox txtBlue 
         Height          =   270
         Index           =   1
         Left            =   5280
         TabIndex        =   40
         Text            =   "255"
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtGreen 
         Height          =   270
         Index           =   1
         Left            =   3840
         TabIndex        =   38
         Text            =   "255"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox txtRed 
         Height          =   270
         Index           =   1
         Left            =   3840
         TabIndex        =   36
         Text            =   "255"
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtAlpha 
         Height          =   270
         Index           =   0
         Left            =   2160
         TabIndex        =   34
         Text            =   "255"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox txtBlue 
         Height          =   270
         Index           =   0
         Left            =   2160
         TabIndex        =   32
         Text            =   "255"
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtGreen 
         Height          =   270
         Index           =   0
         Left            =   720
         TabIndex        =   30
         Text            =   "255"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox txtRed 
         Height          =   270
         Index           =   0
         Left            =   720
         TabIndex        =   28
         Text            =   "255"
         Top             =   4920
         Width           =   735
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   6120
         Width           =   1935
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   1200
         Width           =   2775
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   3240
         Left            =   3240
         ScaleHeight     =   216
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   184
         TabIndex        =   18
         Top             =   1560
         Width           =   2760
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3364
         Left            =   3960
         List            =   "frmEditor_Resource.frx":3374
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   3240
         Left            =   120
         ScaleHeight     =   216
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   184
         TabIndex        =   5
         Top             =   1560
         Width           =   2760
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label13 
         Caption         =   "Alpha:"
         Height          =   255
         Left            =   4680
         TabIndex        =   41
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Blue:"
         Height          =   255
         Left            =   4680
         TabIndex        =   39
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Green:"
         Height          =   255
         Left            =   3240
         TabIndex        =   37
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Red:"
         Height          =   255
         Left            =   3240
         TabIndex        =   35
         Top             =   4920
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   3045
         X2              =   3045
         Y1              =   960
         Y2              =   5520
      End
      Begin VB.Label Label9 
         Caption         =   "Alpha:"
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Blue:"
         Height          =   255
         Left            =   1560
         TabIndex        =   31
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Green:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Red:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   6130
         Width           =   1455
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation:"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   5770
         Width           =   825
      End
      Begin VB.Label lblExhaustedPic 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   20
         Top             =   960
         Width           =   2700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   3120
         TabIndex        =   16
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Item Reward:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   6130
         Width           =   1005
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Req. Tool:"
         Height          =   180
         Left            =   3120
         TabIndex        =   13
         Top             =   5770
         Width           =   765
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         Caption         =   "Health Pool:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   6480
         Width           =   930
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Sec):"
         Height          =   180
         Left            =   3120
         TabIndex        =   11
         Top             =   6480
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   3120
         TabIndex        =   9
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource List"
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6900
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmbAnimation_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).Animation = cmbAnimation.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub cmbItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).ItemReward = cmbItem.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbItem_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).ToolRequired = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearResource EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ResourceEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlExhaustedPic.max = NumResources
    scrlNormalPic.max = NumResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ResourceEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlExhaustedPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblExhaustedPic.Caption = "Exhausted Image: " & scrlExhaustedPic.value
    EditorResource_DrawExhaustedSprite
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlExhaustedPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNormalPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.value
    EditorResource_DrawNormalSprite
    Resource(EditorIndex).ResourceImage = scrlNormalPic.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNormalPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHealth_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Val(Trim$(txtHealth.text)) < 0 Then txtHealth.text = "0"
    
    Resource(EditorIndex).health = Val(Trim$(txtHealth.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txthealthe_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage2_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage2_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Resource(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Resource(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtRed_Change(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Val(Trim$(txtRed(index).text)) < 0 Then txtRed(index).text = "0"
    If Val(Trim$(txtRed(index).text)) > 255 Then txtRed(index).text = "255"
    
    Resource(EditorIndex).Red(index) = Val(Trim$(txtRed(index).text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtRed_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtGreen_Change(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Val(Trim$(txtGreen(index).text)) < 0 Then txtGreen(index).text = "0"
    If Val(Trim$(txtGreen(index).text)) > 255 Then txtGreen(index).text = "255"
    
    Resource(EditorIndex).Green(index) = Val(Trim$(txtGreen(index).text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtGreen_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtBlue_Change(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Val(Trim$(txtBlue(index).text)) < 0 Then txtBlue(index).text = "0"
    If Val(Trim$(txtBlue(index).text)) > 255 Then txtBlue(index).text = "255"
    
    Resource(EditorIndex).Blue(index) = Val(Trim$(txtBlue(index).text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtBlue_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAlpha_Change(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Val(Trim$(txtAlpha(index).text)) < 0 Then txtAlpha(index).text = "0"
    If Val(Trim$(txtAlpha(index).text)) > 255 Then txtAlpha(index).text = "255"
    
    Resource(EditorIndex).Alpha(index) = Val(Trim$(txtAlpha(index).text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAlpha_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpawn_Change()
        ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Val(Trim$(txtSpawn.text)) < 0 Then txtSpawn.text = "0"
    
    Resource(EditorIndex).RespawnTime = Val(Trim$(txtSpawn.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawn_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
