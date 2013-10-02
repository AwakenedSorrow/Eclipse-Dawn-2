VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eclipse Dawn Editor - Login"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkRemember 
         Caption         =   "Remember"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkRemember_Click()
    If chkRemember.value = 1 Then
        Options.RememberUser = 1
    Else
        Options.RememberUser = 0
    End If
End Sub

Private Sub cmdLogin_Click()
    ' Check if we need to save our user to the file.
    If Options.RememberUser = 1 And Options.Username <> Trim(txtUsername.Text) Then
        Options.Username = Trim(txtUsername.Text)
        SaveOptions App.Path & "\" & OPTIONS_FILE
    End If
    
    ' Handle the actual log in sequence from here.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Someone closed the login screen, exit the program.
    DestroyEditor
End Sub
