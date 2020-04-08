Option Explicit

Public Sub TestKeyDict()
' Test key dictionary

Dim k As Long

''' Populate dictionary
Call utils.SetKeys


If KeyDict.Count > 0 Then
    For k = 1 To KeyDict.Count
        'Range("A" & varKey).Value = oDic(varKey)
        Debug.Print k, KeyDict(k)
    Next
End If

End Sub

Public Sub Test_KeyboardClass()

    Dim kb As CKeyboard
    Set kb = New CKeyboard

End Sub

Public Sub Test_ArrayBool()
' Test function to evaluate whether or not a number is in an array

Dim tmp(1 To 3) As Long
Dim test_n As Long
test_n = 1

If utils.IsInArray(test_n, tmp) Then
    Debug.Print "Test value found."
End If

End Sub

Public Sub Test_ArrayCount()

Dim tmp(0 To 1)
Dim i As Integer, n As Integer
n = 5
For i = LBound(tmp) To UBound(tmp)
    n = n + i
    Debug.Print i, n
Next
End Sub
