VERSION 5.00
Begin VB.Form frmEditor_Animation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation Editor"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   474
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   674
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framLayer 
      Caption         =   "Layer 1 (Above Player)"
      Height          =   2775
      Index           =   1
      Left            =   3360
      TabIndex        =   29
      Top             =   3720
      Width           =   6615
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1925
         Index           =   1
         Left            =   4440
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   38
         Top             =   600
         Width           =   1925
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   1815
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1815
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   2280
         Width           =   1815
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   1815
      End
      Begin VB.HScrollBar scrlRed 
         Height          =   255
         Index           =   1
         Left            =   2280
         Max             =   255
         TabIndex        =   33
         Top             =   480
         Value           =   255
         Width           =   1815
      End
      Begin VB.HScrollBar scrlGreen 
         Height          =   255
         Index           =   1
         Left            =   2280
         Max             =   255
         TabIndex        =   32
         Top             =   1080
         Value           =   255
         Width           =   1815
      End
      Begin VB.HScrollBar scrlBlue 
         Height          =   255
         Index           =   1
         Left            =   2280
         Max             =   255
         TabIndex        =   31
         Top             =   1680
         Value           =   255
         Width           =   1815
      End
      Begin VB.HScrollBar scrlAlpha 
         Height          =   255
         Index           =   1
         Left            =   2280
         Max             =   255
         TabIndex        =   30
         Top             =   2280
         Value           =   255
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Preview:"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   47
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Label lblRed 
         Caption         =   "Red: 255"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   42
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblGreen 
         Caption         =   "Green: 255"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   41
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblBlue 
         Caption         =   "Blue: 255"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   40
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblAlpha 
         Caption         =   "Alpha: 255"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   39
         Top             =   2040
         Width           =   1695
      End
   End
   Begin VB.Frame framLayer 
      Caption         =   "Layer 0 (Below Player)"
      Height          =   2775
      Index           =   0
      Left            =   3360
      TabIndex        =   10
      Top             =   840
      Width           =   6615
      Begin VB.HScrollBar scrlAlpha 
         Height          =   255
         Index           =   0
         Left            =   2280
         Max             =   255
         TabIndex        =   24
         Top             =   2280
         Value           =   255
         Width           =   1815
      End
      Begin VB.HScrollBar scrlBlue 
         Height          =   255
         Index           =   0
         Left            =   2280
         Max             =   255
         TabIndex        =   23
         Top             =   1680
         Value           =   255
         Width           =   1815
      End
      Begin VB.HScrollBar scrlGreen 
         Height          =   255
         Index           =   0
         Left            =   2280
         Max             =   255
         TabIndex        =   22
         Top             =   1080
         Value           =   255
         Width           =   1815
      End
      Begin VB.HScrollBar scrlRed 
         Height          =   255
         Index           =   0
         Left            =   2280
         Max             =   255
         TabIndex        =   21
         Top             =   480
         Value           =   255
         Width           =   1815
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1815
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1815
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1815
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1925
         Index           =   0
         Left            =   4440
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   11
         Top             =   600
         Width           =   1925
      End
      Begin VB.Label lblAlpha 
         Caption         =   "Alpha: 255"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   28
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblBlue 
         Caption         =   "Blue: 255"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   27
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblGreen 
         Caption         =   "Green: 255"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   26
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblRed 
         Caption         =   "Red: 255"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Preview:"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animation Properties"
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Animation List"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         ItemData        =   "frmEditor_Animation.frx":0000
         Left            =   120
         List            =   "frmEditor_Animation.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Animation(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Animation(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    AnimationEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    
    ClearAnimation EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    AnimationEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    AnimationEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 0 To 1
        scrlSprite(i).max = NumAnimations
        scrlLoopCount(i).max = 100
        scrlFrameCount(i).max = 100
        scrlLoopTime(i).max = 1000
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    AnimationEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAlpha_Change(Index As Integer)
    lblAlpha(Index).Caption = "Alpha: " & Trim(Str(scrlAlpha(Index).value))
    Animation(EditorIndex).Alpha(Index) = scrlAlpha(Index).value
End Sub

Private Sub scrlBlue_Change(Index As Integer)
    lblBlue(Index).Caption = "Blue: " & Trim(Str(scrlBlue(Index).value))
    Animation(EditorIndex).Blue(Index) = scrlBlue(Index).value
End Sub

Private Sub scrlFrameCount_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFrameCount(Index).Caption = "Frame Count: " & scrlFrameCount(Index).value
    Animation(EditorIndex).Frames(Index) = scrlFrameCount(Index).value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFrameCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFrameCount_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlFrameCount_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFrameCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlGreen_Change(Index As Integer)
    lblGreen(Index).Caption = "Green: " & Trim(Str(scrlGreen(Index).value))
    Animation(EditorIndex).Green(Index) = scrlGreen(Index).value
End Sub

Private Sub scrlLoopCount_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblLoopCount(Index).Caption = "Loop Count: " & scrlLoopCount(Index).value
    Animation(EditorIndex).LoopCount(Index) = scrlLoopCount(Index).value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopCount_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlLoopCount_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopTime_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblLoopTime(Index).Caption = "Loop Time: " & scrlLoopTime(Index).value
    Animation(EditorIndex).looptime(Index) = scrlLoopTime(Index).value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopTime_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopTime_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlLoopTime_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopTime_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRed_Change(Index As Integer)
    lblRed(Index).Caption = "Red: " & Trim(Str(scrlRed(Index).value))
    Animation(EditorIndex).Red(Index) = scrlRed(Index).value
End Sub

Private Sub scrlSprite_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite(Index).Caption = "Sprite: " & scrlSprite(Index).value
    Animation(EditorIndex).Sprite(Index) = scrlSprite(Index).value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Animation(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
