Attribute VB_Name = "modSound"
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Global Const SND_ASYNC = 0
Global Const SND_NODEFAULT = 1
