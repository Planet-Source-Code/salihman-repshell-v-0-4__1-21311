Attribute VB_Name = "modMultimedia"
'Sound API
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As SoundFlags) As Long
    Enum SoundFlags
        SND_SYNC = &H0
        SND_ASYNC = &H1
        SND_NODEFAULT = &H2
        SND_MEMORY = &H4
        SND_LOOP = &H8
        SND_NOSTOP = &H10
    End Enum
    
Public Sub Playsound(sName As String, Optional PlayFlags As SoundFlags)
    sndPlaySound AppResourcePath & "Sound\" & sName & ".wav", SND_ASYNC Or SND_NODEFAULT Or PlayFlags
End Sub
