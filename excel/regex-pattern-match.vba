Option Explicit
Option Compare Binary

' NOTE: Select Tools -> References -> Microsoft VBScript Regular Expressions 5.5
' Alternative method provided in function if this step isn't complete.

Public Function PatternMatch(input_cell As Range, Optional regexp_pattern As String = "") As String
''' Extract value from cell using regular expressions.

If Not regexp_pattern = "" Then
    Dim matches As Object
    
    Dim regex As New RegExp
    
    ' If tool not selected, uncomment and use the following for regex object.
    'Dim regex as Object
    'Set regex = CreateObject("VBScript.RegExp")
    
    Dim match As Variant
    
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = regexp_pattern
    End With
    
    If regex.Test(input_cell.value) Then
        Set matches = regex.Execute(input_cell.value)
        PatternMatch = matches(0).value
    End If
    Set matches = Nothing
Else
    PatternMatch = ""
End If

End Function
