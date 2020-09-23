Attribute VB_Name = "modPlayWav"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     1 September 1999
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Wav module
' -----------------------------------------------
'COMMENTS : This module was made to play a wave
'           sound.
'================================================

'all variables must be declared
Option Explicit

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------
'plays the specified sound file if it exists
Private Declare Function sndPlaySound _
        Lib "winmm" _
        Alias "sndPlaySoundA" _
            (ByVal lpszSoundName As String, _
             ByVal uFlags As Long) _
             As Long

'------------------------------------------------
'                 PROCEDURES
'------------------------------------------------
Public Sub PlaySound(strFile As String)
    'play the wav file specified
    
    Const SYNC      As Long = 1
    
    Dim lngResult   As Long
    
    If (Dir(strFile) <> "") And (Trim(LCase(Right(strFile, 4))) = ".wav") Then
        'if file exists and is a .wav, then play
        lngResult = sndPlaySound(ByVal CStr(strFile), _
                                 SYNC)
    End If
End Sub
