Option Explicit

Public Function Append(ByRef obj, ByVal value, Optional seperator As Variant = ",")
'#################################################
' Append items to the end of an object.
' obj should be a non-numeric object.
'#################################################

  obj = obj & seperator & value

End Function


Public Sub test_Append()
'#################################################
' Test `Utils.Append()`
' Append a non-numeric object to a target object
'
' Immediate window:
'   Call test_Append()
'#################################################

Dim str As String: str = "Item 1"


''' Append item with default separator
Call Utils.Append(str, "Item 2")

Debug.Print str

''' Append item with specific separator
Call Utils.Append(str, "Item 3", " - ")

Debug.Print str

End Sub
