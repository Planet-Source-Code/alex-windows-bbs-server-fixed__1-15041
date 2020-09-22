Attribute VB_Name = "Module1"
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    'constants


Public Enum sndConst
    SND_ASYNC = &H1 ' play asynchronously
    SND_LOOP = &H8 ' loop the sound until Next sndPlaySound
    SND_MEMORY = &H4 ' lpszSoundName points To a memory file
    SND_NODEFAULT = &H2 ' silence Not default, If sound not found
    SND_NOSTOP = &H10 ' don't stop any currently playing sound
    SND_SYNC = &H0 ' play synchronously (default), halts prog use till done playing
End Enum
