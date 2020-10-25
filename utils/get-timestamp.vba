' VBA Function for formatted current timestamp
' Date: 2020-10-25
' Contributor(s): Mark Moretto


' @summary Return timestamp for current time, formatted to millisecond
' @params {String} dt_format - Option datetime format string.  Default is `yyyy-MM-dd hh:mm:ss.ns`
' @returns {String} Formatted timestamp
Public Function GetTimestamp(Optional dt_format As String = "yyyy-MM-dd hh:mm:ss.ns") As String
On Error GoTo 0
' Note: This function shouldn't really return an error.
    GetTimestamp = Format(Now, dt_format)
End Function
