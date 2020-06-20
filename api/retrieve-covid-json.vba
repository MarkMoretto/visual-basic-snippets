''' This can go into a general VBA interfact (ALT + F11 in Excel, Access, etc.)
''' Two references are used: `Microsoft Scripting Runtime` and `Microsoft WinHttp Services, version 5.1`

Option Explicit

Private Function FormatParams(params) As String
On Error Resume Next
''' Parse parameters from `params` parameter, which should be a Scripting.Dictionary object.

Dim k
Dim param_str As String

If Not params Is Nothing Then
    
    For Each k In params.Keys()
        param_str = param_str & k & "=" & CStr(params(k)) & "&"
    Next
    
    param_str = Left(param_str, Len(param_str) - 1)

End If

FormatParams = param_str

End Function




Public Function TestApiCall() As String
''' Testing the API call
''' Params are setup manually, but that could be adjusted by adding an interim method
''' or getting data from a range in the workbook.

Dim url As String
Dim resp As String
Dim params As String
Dim param_dict As New Scripting.Dictionary

''' If all references are selected, use the top method
''' otherwise comment that out and use the CreateObject method
Dim oReq As New WinHttp.WinHttpRequest ''' Requires reference to `Microsoft WinHttp Services, version 5.1`

'Dim oReq As Object
'Set oReq = CreateObject("WinHttp.WinHttpRequest.5.1") '' If no reference selected

''' Our base URL
''' This is for CDC.gov resources.
url = "https://data.cdc.gov/resource/r8kw-7aab.json"


'' Add parameters
param_dict.Add "$limit", 10
param_dict.Add "$offset", 0


''' Set parameter string
params = vbNullString
If param_dict.Count > 0 Then
    params = FormatParams(param_dict)
    url = url & "?" & params
End If


With oReq
    .Open "GET", url, False
    .SetRequestHeader "Content-Type", "application/json"
    .SetRequestHeader "Accept", "application/json"
    .Send
    resp = .ResponseText
End With

''' This prints results to the immediate window.
TestREST = resp


Set oReq = Nothing

End Function
