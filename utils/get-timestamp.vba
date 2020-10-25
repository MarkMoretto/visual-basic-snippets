' VBA Function for formatted current timestamp
' Date: 2020-10-25
' Contributor(s): Mark Moretto


' @summary Return timestamp for current time, formatted to millisecond
' @params Null
' @returns {String} Formatted timestamp
Public Function GetTimestamp() As String
On Error Goto 0
' Note: This function shouldn't really return an error, but it's better to be safe!
    GetTimestamp = Format(Now, "yyyy-MM-dd hh:mm:ss.ns")
End Function
