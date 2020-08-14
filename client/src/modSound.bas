Attribute VB_Name = "modSound"
Option Explicit

' Master object
Public DirectSound As DirectSound8

' Using for playing and stop musics
Public SoundBuffer(0 To Sound_Count - 1) As DirectSoundSecondaryBuffer8
Public Performance As DirectMusicPerformance8
Public Segment As DirectMusicSegment8

' Using for the loading of sound
Public Loader As DirectMusicLoader8
Public SoundDesc As DSBUFFERDESC

' The music volume
Public Const MusicVolume As Integer = 85

' The current music currently loaded
Public CurrentMusic As Byte

Public Function Init_Music() As Boolean
    Dim dmParams As DMUS_AUDIOPARAMS

    ' Handle error then exit out
    On Error GoTo errorhandler

    ' Create the directSound device (with the default device)
    Set DirectSound = DX8.DirectSoundCreate(vbNullString)
    Call DirectSound.SetCooperativeLevel(frmMain.hWnd, DSSCL_PRIORITY)
    
    ' Set up the buffer description for later use
    SoundDesc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    ' Creater music loader
    Set Loader = DX8.DirectMusicLoaderCreate

    ' Create music performance
    Set Performance = DX8.DirectMusicPerformanceCreate

    ' Init performance efects
    Call Performance.InitAudio(frmMain.hWnd, DMUS_AUDIOF_ALL, dmParams, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128)
    Call Performance.SetMasterVolume(MusicVolume)
    Call Performance.SetMasterAutoDownload(True)
    
    ' Error handler
    Exit Function
errorhandler:
    Init_Music = False
End Function

Public Sub Play_Sound(WaveNum As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Exit out early if we have the system turned off
    If Options.Sound = 0 Then Exit Sub
    
    ' Prevent subscript out range
    If WaveNum = 0 Or WaveNum > Sound_Count - 1 Then Exit Sub
    
    ' Create the buffer if needed
    If SoundBuffer(WaveNum) Is Nothing And FileExist(App.Path & "\data files\sound\" & WaveNum & ".wav", True) Then
        Set SoundBuffer(WaveNum) = DirectSound.CreateSoundBufferFromFile(App.Path & "\data files\sound\" & WaveNum & ".wav", SoundDesc)
    End If
    
    ' Don't work if sound buffer not exists
    If SoundBuffer(WaveNum) Is Nothing Then Exit Sub

    ' Play the sound
    Call SoundBuffer(WaveNum).Play(DSBPLAY_DEFAULT)
  
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Play_Sound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Play_Music(MusicNum As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Exit out early if we have the system turned off
    If Options.Music = 0 Then Exit Sub
    
    ' don't re-start currently playing songs
    If CurrentMusic = MusicNum Then Exit Sub
    If Loader Is Nothing Then Exit Sub
    
    ' Load Music Segment
    Set Segment = Loader.LoadSegment(App.Path & "\data files\music\" & MusicNum & ".mid")

    ' Segment efects
    Call Segment.SetStandardMidiFile
    Call Segment.SetRepeats(-1)

    ' Play music
    Call Performance.PlaySegmentEx(Segment, DMUS_SEGF_DEFAULT, 0)

    ' Set current music
    CurrentMusic = MusicNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Play_Music", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Stop_Music()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Stop player current segment
    If Not Segment Is Nothing Then
        Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
    End If
    
    CurrentMusic = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Destroy_Music", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Destroy_Music()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' destroy music engine
    If Not Performance Is Nothing Then Set Performance = Nothing
    If Not Segment Is Nothing Then Set Segment = Nothing
    If Not Loader Is Nothing Then Set Loader = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Destroy_Music", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
