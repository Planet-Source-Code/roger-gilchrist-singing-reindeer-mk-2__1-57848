Attribute VB_Name = "DX7Sound"
Option Explicit
'***********************************
'Module by D.R Hall, modified by me.
'***********************************
Private m_dx              As New DirectX7
Private m_dxs             As DirectSound 'Then there is the sub object, DirectSound:
'buffers
Type dxBuffers
  isLoaded                As Boolean
  'command button on the form is assigned
  Buffer                  As DirectSoundBuffer
End Type
'
Private SoundFolder       As String 'Holds Path to Sound folder
Private SB()              As dxBuffers 'An Array of BUFFERS,
Private CurrentBuffer     As Long    'Holds last assign Random Buffer Number

Public Sub CreateBuffers(AmountOfBuffer As Long, _
                         DefaultFile As String)

  'create the buffers

  ReDim SB(AmountOfBuffer)
  For AmountOfBuffer = 0 To AmountOfBuffer
    DX7LoadSound AmountOfBuffer, DefaultFile 'must assign a defualt sound
    'set @ 100 in the program, you can change it
    VolumeLevel AmountOfBuffer, 100
  Next AmountOfBuffer

End Sub

Public Sub DX7LoadSound(ByVal intBuffer As Long, _
                        ByVal sfile As String)

  Dim waveFormat As WAVEFORMATEX   'what sort of buffer to create
  Dim fileName   As String
  Dim bufferDesc As DSBUFFERDESC

  'a new object that when filled in is passed to the DS object to describe
  bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC 'These settings should do for almost any app....
  With waveFormat
    .nFormatTag = WAVE_FORMAT_PCM
    .nChannels = 2    '2 channels
    .lSamplesPerSec = 22050
    .nBitsPerSample = 16  '16 bit rather than 8 bit
    .nBlockAlign = .nBitsPerSample / 8 * .nChannels
    .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
  End With 'waveFormat
  fileName = SoundFolder & sfile
  On Error GoTo Continue
  Set SB(intBuffer).Buffer = m_dxs.CreateSoundBufferFromFile(fileName, bufferDesc, waveFormat)
  'success!
  SB(intBuffer).isLoaded = True

Exit Sub

Continue:
  MsgBox "Error can't find file: " & fileName

End Sub

Public Function PlaySoundAnyBuffer(fileName As String, _
                                   Optional ByVal Volume As Byte, _
                                   Optional ByVal PanValue As Byte, _
                                   Optional ByVal LoopIt As Byte) As Long

  'this is what we use, only memory dependent (I think)

  Do While SB(CurrentBuffer).Buffer.GetStatus = DSBSTATUS_PLAYING 'Find an empty buffer
    CurrentBuffer = CurrentBuffer + 1
    If CurrentBuffer > UBound(SB) Then
      CurrentBuffer = 0
    End If
  Loop
  DX7LoadSound CurrentBuffer, fileName
  'loop the sound
  If SB(CurrentBuffer).isLoaded Then
    SB(CurrentBuffer).Buffer.Play LoopIt
  End If

End Function

Public Sub SetupDX7Sound(CurrentForm As Form)

  Set m_dxs = m_dx.DirectSoundCreate("") 'create a DSound object
  'check for any errors, if there are no errors the user has got DX7 and a functional sound card
  If Err.Number <> 0 Then
    MsgBox "Unable to start DirectSound. Check to see that your sound card is properly installed"
    End
  End If
  m_dxs.SetCooperativeLevel CurrentForm.hWnd, DSSCL_PRIORITY 'THIS MUST BE SET BEFORE WE CREATE ANY BUFFERS

End Sub

Public Sub SoundDir(ByVal FolderPath As String)

  'path

  SoundFolder = FolderPath & "\"

End Sub

Public Sub VolumeLevel(ByVal intBuffer As Long, _
                       ByVal Volume As Byte)

  If Volume > 0 Then ' stop division by 0
    SB(intBuffer).Buffer.SetVolume (60 * Volume) - 6000
   Else
    SB(intBuffer).Buffer.SetVolume -6000
  End If

End Sub

':)Code Fixer V2.8.0 (22/12/2004 2:17:49 AM) 16 + 94 = 110 Lines Thanks Ulli for inspiration and lots of code.

