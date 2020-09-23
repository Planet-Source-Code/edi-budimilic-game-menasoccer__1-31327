Attribute VB_Name = "DirectSoundMod"
Option Explicit 'Provjerava varijable kojima nije postavljen Dim,Public i sl.

Public DX As New DirectX7 'Treba paziti da se ovo slucajno ne pojavi u DirectDraw modulu
Public DX_Draw As DirectDraw7
Public DS As DirectSound  'Object Direct Sound
Public dsMissel As DirectSoundBuffer
Public dsArma1 As DirectSoundBuffer
Public dsArma2 As DirectSoundBuffer
Public dsArma3 As DirectSoundBuffer
Public dsTiro As DirectSoundBuffer
Public dsExplosao As DirectSoundBuffer
Public dsExtra As DirectSoundBuffer
Public dsVidaExtra As DirectSoundBuffer
Public dsPausa As DirectSoundBuffer
Public dsKick1 As DirectSoundBuffer
Public BDesc As DSBUFFERDESC      'Direct sound buffer destination
Public DSwaveFormat As WAVEFORMATEX
Public Volume As Integer

Public Type SoundPublic
  dsMissel_ As Integer
  dsArma_ As Integer
  dsArma2_ As Integer
  dsArma3_ As Integer
  dsTiro_ As Integer
  dsExplosao_ As Integer
  dsExtra_ As Integer
  dsVidaExtra_ As Integer
  dsPausa_ As Integer
  dsKick1_ As Integer
End Type
Public Sndp As SoundPublic

Public Sub StopSound_DS()
    Set dsMissel = Nothing
    Set dsArma1 = Nothing
    Set dsArma2 = Nothing
    Set dsArma3 = Nothing
    Set dsTiro = Nothing
    Set dsExplosao = Nothing
    Set dsExtra = Nothing
    Set dsVidaExtra = Nothing
    Set dsPausa = Nothing
    Set DS = Nothing
End Sub

Public Sub InitializeDS()
On Error Resume Next
Set DS = DX.DirectSoundCreate("")
If Err.Number <> 0 Then
MsgBox "Unable to start DirectSound. Close any program that uses SoundCard!"
End If

'DSSCL_PRIORITY=no cooperation, exclusive access to the sound card, Needed for games
'DSSCL_NORMAL=cooperates with other apps, shares resources, Good for general windows multimedia apps.
DS.SetCooperativeLevel Form1.hwnd, DSSCL_EXCLUSIVE
   
BDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
With DSwaveFormat
    .nFormatTag = WAVE_FORMAT_PCM
    .nChannels = 2
    .lSamplesPerSec = 22050
    .nBitsPerSample = 16
    .nBlockAlign = .nBitsPerSample / 8 * .nChannels
End With

'Stvara Buffer unaprijed za zvukeve koji ce se koristiti
Set dsKick1 = DS.CreateSoundBufferFromFile(App.Path & "\sounds\kick1.wav", BDesc, DSwaveFormat)
End Sub

