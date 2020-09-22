Attribute VB_Name = "DX7Sound"
Option Explicit

'This module was created by D.R Hall
'For more Information and latest version
'E-mail me, derek.hall@virgin.net

'Heavily adjusted by Danny,--


Private m_dx As New DirectX7
Private m_dxs As DirectSound

Type dxBuffers
  isLoaded As Boolean
  Buffer As DirectSoundBuffer
End Type

Public SndPlayBuffer(MAX_PLAYBACK_BUFFERS) As dxBuffers 'An Array of BUFFERS for playback.
Public SndWavBuffer(MAX_SOUND_BUFFERS) As dxBuffers     'An array for storing wavs in memory.
'Public sndMusicBuffer As dxBuffers

Public Function PlaySoundAnyBuffer2(WavId As Integer, Optional Volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte) As Integer
Static CurrentBuffer As Integer

  ' arguments: Filename As String, Optional Volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte) As Integer
 'Keep looking to find an unused playback buffer, if not, kill the first.
  Do While SndPlayBuffer(CurrentBuffer).Buffer.GetStatus = DSBSTATUS_PLAYING
    CurrentBuffer = CurrentBuffer + 1
    If CurrentBuffer > MAX_PLAYBACK_BUFFERS Then CurrentBuffer = 0
  Loop

 'Ok, we should have an empty playback buffer right now,..
 'Copy the soundbuffer to the playback buffer
  Level.lDebug.Caption = "Buffer: " & CurrentBuffer
 
  SndPlayBuffer(CurrentBuffer) = SndWavBuffer(WavId)
  
' DX7LoadSound CurrentBuffer, Filename
  If PanValue <> 50 Then PanSound SndPlayBuffer(CurrentBuffer), PanValue
  If Volume < 100 Then VolumeLevel SndPlayBuffer(CurrentBuffer), Volume
  If SndPlayBuffer(CurrentBuffer).isLoaded Then SndPlayBuffer(CurrentBuffer).Buffer.Play LoopIt 'dsb_looping=1, dsb_default=0

End Function

'Private CurrentBuffer As Integer 'Holds last assign Random Buffer Number
Public Sub CreateBuffers()
Dim DefaultFile As String
Dim BuffNum As Integer

'## This routine creates & initialises the playback and soundfile-buffers

  DefaultFile = App.Path & WAV_DEFAULT
 
 'Init playblack buffers
  For BuffNum = 0 To MAX_PLAYBACK_BUFFERS
      DX7LoadSound SndPlayBuffer(BuffNum), DefaultFile  'must assign a defualt sound
      VolumeLevel SndPlayBuffer(BuffNum), 50            'set volume to 50% for default
  Next BuffNum

 'Init Sound/wav buffers
  For BuffNum = 0 To MAX_SOUND_BUFFERS
      DX7LoadSound SndWavBuffer(BuffNum), DefaultFile  'must assign a defualt sound
      VolumeLevel SndWavBuffer(BuffNum), 50            'set volume to 50% for default
  Next BuffNum
  
 'Init Music buffer
 'DX7LoadSound sndMusicBuffer, DefaultFile

End Sub

Public Sub SetupDX7Sound(CurrentForm As Form)
  
  On Error Resume Next
  
  Set m_dxs = m_dx.DirectSoundCreate("") 'create a DSound object
  
 'Next you check for any errors, if there are no errors the user has got DX7 and a functional sound card

  If Err.Number <> 0 Then
    MsgBox "Unable to allocate Sound Card, please close other applications and try again! ;)"
    End
  End If
  
  m_dxs.SetCooperativeLevel CurrentForm.hwnd, DSSCL_PRIORITY  'THIS MUST BE SET BEFORE WE CREATE ANY BUFFERS
  
  'associating our DS object with our window is important. This tells windows to stop
  'other sounds from interfering with ours, and ours not to interfere with other apps.
  'The sounds will only be played when the from has got focus.
  'DSSCL_PRIORITY=no cooperation, exclusive access to the sound card, Needed for games
  'DSSCL_NORMAL=cooperates with other apps, shares resources, Good for general windows multimedia apps.
  
End Sub

Public Sub DX7LoadSound(ByRef sBuffer As dxBuffers, ByVal WavFileName As String)
  
  Dim bufferDesc As DSBUFFERDESC  'a new object that when filled in is passed to the DS object to describe
  Dim waveFormat As WAVEFORMATEX  'what sort of buffer to create
  
  bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN _
  Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC 'These settings should do for almost any app....
  
  waveFormat.nFormatTag = WAVE_FORMAT_PCM
  waveFormat.nChannels = 2    '2 channels
  waveFormat.lSamplesPerSec = 44100
  waveFormat.nBitsPerSample = 16  '16 bit rather than 8 bit
  waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
  waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign

  On Error GoTo Continue
  Set sBuffer.Buffer = m_dxs.CreateSoundBufferFromFile(WavFileName, bufferDesc, waveFormat)
  sBuffer.isLoaded = True
  Exit Sub
  
Continue:
  MsgBox "Error can't find sound-file: " & WavFileName

End Sub

Public Function PlaySoundAnyBuffer(WavId As Integer, Optional Volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte) As Integer
Static CurrentBuffer As Integer
Static tel As Integer

  ' arguments: Filename As String, Optional Volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte) As Integer
 'Keep looking to find an unused playback buffer, if not, kill the first.
  For tel = 0 To MAX_PLAYBACK_BUFFERS - 1
     'Search for a free playback buffer
      If Not (SndPlayBuffer(CurrentBuffer).Buffer.GetStatus = DSBSTATUS_PLAYING) Then
        'playback sound
         CurrentBuffer = tel
         SndPlayBuffer(CurrentBuffer) = SndWavBuffer(WavId)
  
         If PanValue <> 50 Then PanSound SndPlayBuffer(CurrentBuffer), PanValue
         If Volume < 100 Then VolumeLevel SndPlayBuffer(CurrentBuffer), Volume
         If SndPlayBuffer(CurrentBuffer).isLoaded Then SndPlayBuffer(CurrentBuffer).Buffer.Play LoopIt 'dsb_looping=1, dsb_default=0
         
        'Quit for-loop
         Exit For
      End If
  Next tel

End Function

Public Function PlayMusicBuffer(WavId As Integer, Optional Volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte) As Integer
Static CurrentBuffer As Integer

 ' arguments: Filename As String, Optional Volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte) As Integer
  
  'sndMusicBuffer = SndWavBuffer(WavId)
  
  'If PanValue <> 50 Then PanSound sndMusicBuffer, PanValue
  'If Volume < 100 Then VolumeLevel sndMusicBuffer, Volume
  'If sndMusicBuffer.isLoaded Then sndMusicBuffer.Buffer.Play 1 'dsb_looping=1, dsb_default=0

End Function

'Public Sub PlaySoundWithPan(Buffer As Integer, Filename As String, Optional Volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte)
'  DX7LoadSound Buffer, Filename
'  If PanValue <> 50 And PanValue < 100 Then PanSound Buffer, PanValue
'  If Volume < 100 Then VolumeLevel Buffer, Volume
'  If SB(Buffer).isLoaded Then SB(Buffer).Buffer.Play LoopIt 'dsb_looping=1, dsb_default=0
'End Sub

Public Sub PanSound(ByRef sBuffer As dxBuffers, PanValue As Byte)
  
  Select Case PanValue
    Case 0
      sBuffer.Buffer.SetPan -10000
    Case 100
      sBuffer.Buffer.SetPan 10000
    Case Else
      sBuffer.Buffer.SetPan (100 * PanValue) - 5000
  End Select
  
End Sub

Public Sub VolumeLevel(ByRef sBuffer As dxBuffers, Volume As Byte)
  If Volume > 0 Then ' stop division by 0
    sBuffer.Buffer.SetVolume (60 * Volume) - 6000
  Else
    sBuffer.Buffer.SetVolume -6000
  End If
End Sub

Public Function IsPlaying(Buffer As Integer) As Long
  IsPlaying = SndPlayBuffer(Buffer).Buffer.GetStatus
End Function
