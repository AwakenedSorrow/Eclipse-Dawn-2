VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmDatabase 
   Caption         =   "Eclipse Dawn - Database Editor"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12000
   Icon            =   "frmDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDeveloper 
      Caption         =   "Developer - (Name)"
      Height          =   3855
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Remove Developers"
         Height          =   255
         Index           =   14
         Left            =   2640
         TabIndex        =   25
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Add Developers"
         Height          =   255
         Index           =   13
         Left            =   2640
         TabIndex        =   24
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit Developers"
         Height          =   255
         Index           =   12
         Left            =   3240
         TabIndex        =   23
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit Animations"
         Height          =   255
         Index           =   11
         Left            =   3240
         TabIndex        =   22
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit Players"
         Height          =   255
         Index           =   10
         Left            =   3240
         TabIndex        =   21
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit Resources"
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   20
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit Spells"
         Height          =   255
         Index           =   8
         Left            =   3240
         TabIndex        =   19
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit Items"
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   18
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit Shops"
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit Maps"
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   16
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit NPCs"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   15
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Change Own Details"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   14
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Use Chat"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Can Edit Maps"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox chkHasRight 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveDeveloper 
         Caption         =   "Save Changes"
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "#"
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin TabDlg.SSTab TabEditor 
      Height          =   9000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   15875
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   9
      Tab             =   8
      TabsPerRow      =   9
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit NPC"
      TabPicture(0)   =   "frmDatabase.frx":3A0A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Edit Item"
      TabPicture(1)   =   "frmDatabase.frx":3A26
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Edit Spell"
      TabPicture(2)   =   "frmDatabase.frx":3A42
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Edit Animation"
      TabPicture(3)   =   "frmDatabase.frx":3A5E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Edit Resource"
      TabPicture(4)   =   "frmDatabase.frx":3A7A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Edit Shop"
      TabPicture(5)   =   "frmDatabase.frx":3A96
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Edit Script"
      TabPicture(6)   =   "frmDatabase.frx":3AB2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Edit Player"
      TabPicture(7)   =   "frmDatabase.frx":3ACE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Edit Developer"
      TabPicture(8)   =   "frmDatabase.frx":3AEA
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).Control(0)=   "Frame1"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Developer List"
         Height          =   8415
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2655
         Begin VB.CommandButton cmdRemoveDeveloper 
            Caption         =   "Remove Developer"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   8040
            Width           =   2415
         End
         Begin VB.CommandButton cmdNewDeveloper 
            Caption         =   "Add New Developer"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   7680
            Width           =   2415
         End
         Begin VB.ListBox lstDeveloperIndex 
            Height          =   7275
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2415
         End
      End
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSaveDeveloper_Click()
    SendSaveDeveloper
End Sub

Private Sub TabEditor_Click(PreviousTab As Integer)
    ' If the previous tab is anything but the tab we're on right now, let's dump the old data and load the new.
    If PreviousTab <> TabEditor.Tab Then
        MsgBox "switched tabs!"
    End If
End Sub
