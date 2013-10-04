Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "Kernel32" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Sub Main()
Dim TempPerc As Byte, Wait As Long
    '  We've just entered the program! Oh god!
    ' First of, let's start setting up a few basic settings we need to get done.
    LoadBarWidth = frmLoad.picProgressFor.Width
    LoadBarPerc = LoadBarWidth / 100
    
    '  Load the editor options.
    LoadOptions App.Path & "\" & OPTIONS_FILE
    
    '  Load TCP Settings.
    TcpInit
    
    ' Check if the directories are there, if not make them. This is a failsafe so that if you, or anyone else ever decides to
    ' start deleting these folders the game won't throw any errors about the locations not existing, worst that could happen
    ' now is that it just throws File not Found errors, which is just as silly. But hey, at least the file structure is there
    ' right?
    SetLoadStatus LoadStateCheckDir, 0
     ChkDir App.Path & "\", "data files"
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    SetLoadStatus LoadStateCheckDir, 25
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    SetLoadStatus LoadStateCheckDir, 50
    ChkDir App.Path & "\data files\graphics\", "spellicons"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    SetLoadStatus LoadStateCheckDir, 75
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    SetLoadStatus LoadStateCheckDir, 100
    
    ' Initialize the sound engine.
    InitBASS
    
    ' Initialie the Rendering engine.
    InitDirect3D8
    EngineInitFontTextures
    
    '  We're done loading, we can actually connect to the server and check our version now!
    If ConnectToServer() = True Then
        '  We're connected! Let's move on and send a version check.
        SendVersionCheck
        CheckingVersion = True
        
        '  Little countdown loop again.
        Wait = GetTickCount
        Do While (GetTickCount <= Wait + 3000) And CheckingVersion <> False
            TempPerc = LoadBarPerc * (((Wait + 3000) - GetTickCount) / 1500)
            SetLoadStatus LoadStateConnecting, TempPerc
            DoEvents
        Loop
        
        ' If Checking Version hasn't been changed, we haven't received a response. So assuming a timeout.
        If CheckingVersion = True Then
            MsgBox "The server could not be reached in time.", vbOKOnly, "Connection Timeout"
            DestroyEditor
        End If
        
        ' We passed everything! Time to load the Login menu.
        Load frmLogin
        
        ' Set the userdata if it's set to be remembered.
        If Options.RememberUser = 1 Then
            frmLogin.txtUsername = Trim$(Options.Username)
            frmLogin.chkRemember.value = 1
        End If
        
        frmLoad.Visible = False ' NEVER unload this. It holds our socket.
        frmLogin.Visible = True
    Else
        '  Oh dear, we couldn't seem to connect.
        MsgBox "The server could not be reached in time.", vbOKOnly, "Connection Timeout"
        DestroyEditor
    End If
End Sub

Public Sub SetStatus(ByVal Status As String)
    frmEditor.stBar.SimpleText = Status
End Sub

Public Sub SetLoadStatus(ByVal Status As String, ByVal Percentage As Long)
Dim Width As Long

    '  Make sure the form is visible, so check if it is.
    If frmLoad.Visible = False Then frmLoad.Visible = True
    
    '  Set the status label.
    frmLoad.lblProgress.Caption = Status
    
    '  Calculate the width we need to use for the current progress.
     Width = Percentage * LoadBarPerc
    
    '  Set the width on the form.
    frmLoad.picProgressFor.Width = Width
    
    ' Make sure the form doesn't lock up.
    DoEvents
End Sub

Public Sub DestroyEditor()
    '  Close the TCP connection.
    DestroyTCP
    
    ' Destroy BASS
    DestroyBASS
    
    ' Destroy DirectX
    UnloadDirectX
    
    ' Unload all Forms
    Unload frmLogin
    Unload frmEditor
    Unload frmDatabase
    Unload frmLoad
    End
End Sub

Public Sub InitMapEditor()
Dim i As Long

    ' Populate the tileset list.
    frmEditor.cmbTileSet.Clear
    For i = 1 To NumTileSets
        frmEditor.cmbTileSet.AddItem CStr(i) & ": " & CStr(Options.TileSetName(i))
    Next
    frmEditor.cmbTileSet.ListIndex = 0
    
    ' Select the Ground layer
    frmEditor.cmbLayerSelect.ListIndex = 0
End Sub
