Attribute VB_Name = "modSoundAPI"
Option Explicit

'-------------------------------------------------------------------------------------
'Min OS Requirements: Windows NT 3.1, Windows 95
'Note: If the played sound is a Resource ID, it must be added with a type of "WAVE"
'      Also, the app must be compiled before you can hear the wave resource plays back
'-------------------------------------------------------------------------------------

Private Const SND_ASYNC As Long = &H1         'Play Asynchronously
Private Const SND_RESOURCE As Long = &H40004  'Name Parameter is a Resource ID
Private Const SND_FILENAME As Long = &H20000  'Name Parameter is a File Name
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Function PlayWave(Sound As String) As Boolean
'The Sound parameter may be a Resource ID of the type "WAVE"
'or simply a wave file path

    Dim Ret As Long
    
    If IsNumeric(Sound) Then
        Ret = PlaySound(CLng(Sound), 0, SND_RESOURCE Or SND_ASYNC)
    Else
        Ret = PlaySound(Sound, 0, SND_FILENAME Or SND_ASYNC)
    End If
    
    If Ret Then PlayWave = True

End Function
