' VBA Function to trim leading and trailing whitespace
' Date: 2020-10-24
' Contributor(s): Mark Moretto




' @summary Trims left and right of string object inplace.
' @param {String} stringObj - A string from which to trim whitespace
' @returns Null

Public Function TrimLR(ByRef stringObj As String)
On Error Goto 0
    Call LTrim(RTrim(stringObj))
End Function
