Attribute VB_Name = "modAudio"
Option Explicit

Public SoundIndex As Long
Public MusicIndex As Long

Public CurMusic As String
Public CurSound As String
Public MusicVolume As Double

Public Sub InitBASS()
    SetLoadStatus LoadStateAudio, 0
    
    ' change and set the current path, to prevent from VB not finding BASS.DLL
    Call ChDrive(App.Path)
    Call ChDir(App.Path)
    
    SetLoadStatus LoadStateAudio, 33
    
    ' check the correct BASS was loaded
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("An incorrect version of bass.dll was loaded.", vbCritical)
        End
    End If

    SetLoadStatus LoadStateAudio, 66
    
    ' initialize BASS
    If (BASS_Init(-1, 44100, 0, frmLoad.hWnd, 0) = 0) Then
        MsgBox ("Could not initialise BASS")
        End
    End If
    
    SetLoadStatus LoadStateAudio, 100
    
End Sub

Public Sub DestroyBASS()

    ' Stop everything
    StopMusic
    StopSound
    
    ' Free bass.dll
    Call BASS_Free
End Sub

Public Sub StopSound()
    BASS_ChannelStop (SoundIndex)
    CurSound = vbNullString
End Sub

Public Sub StopMusic()
    BASS_ChannelStop (MusicIndex)
    CurMusic = vbNullString
End Sub

Public Sub PlayMusic(ByVal FileName As String, Optional ByVal NoFade As Boolean = False)
    If CurMusic = FileName Then Exit Sub
    
    If Options.Music = 0 Then Exit Sub
    
    ' Stop and re-start the channel with the new music
    StopMusic
    
    ' Create the music data
    MusicIndex = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & MUSIC_PATH & FileName), 0, 0, BASS_SAMPLE_LOOP)
    
    ' Set the volume
    MusicVolume = Options.Volume / 100
    Call SetVolume(MusicIndex, MusicVolume)
    
    ' Play it
    Call BASS_ChannelPlay(MusicIndex, False)
    
    ' Set the new current music
    CurMusic = FileName
End Sub

Public Sub PlaySound(ByVal FileName As String)
    If Not FileExist(SOUND_PATH & FileName) Then Exit Sub
    
    If Options.Sound = 0 Then Exit Sub
    
    ' Create the sound data
    SoundIndex = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & SOUND_PATH & FileName), 0, 0, 0)
    
    If SoundIndex <> 0 Then
        ' Set the volume
        MusicVolume = Options.Volume / 100
        Call SetVolume(SoundIndex, MusicVolume)
        Call BASS_ChannelPlay(SoundIndex, False)
    End If

    CurSound = FileName
End Sub

Public Sub SetVolume(ByVal channel As Long, ByVal Volume As Double)
    Call BASS_ChannelSetAttribute(channel, BASS_ATTRIB_VOL, Volume)
End Sub

