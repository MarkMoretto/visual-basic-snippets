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
Call Utils.Incr(i)

Debug.Print i

''' Increment by 2. i should now equal 4.
Call Utils.Incr(i, 2)
Debug.Print i


''' Reset i to zero and increment by 2.
i = 0
Call Utils.Incr(i, 2)
Debug.Print i


''' Reset i and run a loop
i = 0
Dim j As Long: j = 0
While j < 10
    Call Utils.Incr(i, 2)
    Call Utils.Incr(j)
Wend

Debug.Print i

End Sub
