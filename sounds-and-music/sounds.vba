
Option Explicit

''' http://allapi.mentalis.org/apilist/Beep.shtml
''' dwFreq range 37 through 32767
''' dwDuration measured in milliseconds
''' If the function succeeds, the return value is nonzero.
Private Declare Function Tone Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long


Public Function StoMS(ByVal seconds As Single) As Long
On Error GoTo 0
''' Convert seconds to milliseconds
    StoMS = CLng(seconds * 1000)
End Function

Public Sub TestBeep()
''' Test tones
Dim freq As Long
Dim duration As Single
Dim beep_result As Long

freq = 440
duration = 0.25

Tone freq, StoMS(duration)

End Sub
