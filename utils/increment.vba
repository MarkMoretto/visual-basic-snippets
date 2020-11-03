Option Explicit

Public Function Incr(ByRef obj, Optional amount As Variant = 1)
'#################################################
' Increment an object by a given amount, in place
'#################################################

obj = obj + amount
    
End Function


Public Sub test_Incr()
'#################################################
' Test `Utils.Incr()`
' Increment a numeric object by a given amount
'
' Immediate window:
'   Call test_Incr()
'#################################################

Dim i As Long: i = 1

''' Default is increment by 1
Debug.Print "Increment 1 by 1:"
Incr i
Debug.Print i


''' Increment by 2. i should now equal 4.
Debug.Print Chr(13) & "Increment 2 by 2:"
Incr i, 2
Debug.Print i


''' Reset i to zero and increment by 2.
Debug.Print Chr(13) & "Reset variable to 0 and increment by 2:"
i = 0
Incr i, 2
Debug.Print i


''' Reset i and run a loop
Debug.Print Chr(13) & "Reset variable to zero and loop variable by 2 while a second variable increments to 10:"
i = 0
Dim j As Long: j = 0
While j < 10
    Incr i, 2
    Incr j
    Debug.Print "Loop value: " & i
Wend

Debug.Print "Final loop value: " & i

End Sub
