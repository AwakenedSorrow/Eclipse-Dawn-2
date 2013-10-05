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
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "#"
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
Dim Wait As Long
    
    ' Check if we need to save our user to the file.
    If Options.RememberUser = 1 And Options.Username <> Trim(txtUsername.text) Then
        Options.Username = Trim(txtUsername.text)
        SaveOptions App.Path & "\" & OPTIONS_FILE
    End If
    
    ' Handle the actual log in sequence from here.
    If Len(Trim$(txtUsername.text)) > 0 And Len(Trim$(txtPassword.text)) > 0 Then
        SendUserLogin
        frmLogin.Visible = False
        LoggingIn = True
        
        '  Little countdown loop again.
        Wait = GetTickCount
        Do While (GetTickCount <= Wait + 3000) And LoggingIn <> False
            TempPerc = LoadBarPerc * (((Wait + 3000) - GetTickCount) / 1500)
            SetLoadStatus LoadStateLogin, TempPerc
            DoEvents
        Loop
        
        ' If Checking Version hasn't been changed, we haven't received a response. So assuming a timeout.
        If LoggingIn = True Then
            MsgBox "The server could not be reached in time.", vbOKOnly, "Connection Timeout"
            DestroyEditor
        End If
        
        ' Successfully logged in!
        
        ' Load the settings once more, to make sure we get the tileset names.
        LoadOptions App.Path & "\" & OPTIONS_FILE
        
        ' Initialize the map editor.
        InitMapEditor
        
        ' Hide old forms
        frmLoad.Visible = False
        
        '  clear all attribute windows.
        ClearAttributeFrames
        frmEditor.optBlocked.value = True
        
        ' Show the map editor screen.
        frmEditor.Show
        
        ' Start the editor loop
        EditorLooping = True
        EditorLoop
    Else
        MsgBox "Your username/password entry can not be empty!", vbInformation
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Someone closed the login screen, exit the program.
    DestroyEditor
End Sub
