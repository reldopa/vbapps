Attribute VB_Name = "basPlaySound"
Option Explicit

Private Const SND_FILENAME = &H20000
Private Const SND_ASYNC = &H1

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub PlayMsgSound()
    Call sndPlaySound(gAppPath + "msg.wav", SND_ASYNC Or SND_FILENAME)
End Sub

Public Sub PlayUserSound()
    Call sndPlaySound(gAppPath + "knock.wav", SND_ASYNC Or SND_FILENAME)
End Sub

Public Sub PlayFileSound()
    Call sndPlaySound(gAppPath + "file.wav", SND_ASYNC Or SND_FILENAME)
End Sub

Public Sub PlayRingSound()
    Call sndPlaySound(gAppPath + "ring.wav", SND_ASYNC Or SND_FILENAME)
End Sub
