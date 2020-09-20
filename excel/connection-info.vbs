''' Retrieve connection queries and related ModelTables from Excel

'' Set Excel file and output file names.
Dim xl_name, out_name
xl_name = "<name-of-workbook>.xlsx" ''' Name of Excel workbook.
out_name = "workbook-queries.sql" ''' Name of output file.


Public cwd, DebugConfig

''' If DebugConfig = True, then Excel workbook will be visible when getting info.
DebugConfig = False


Private Sub SetCwd()
    ''' Set current working directory variable
    Dim shellObj
    Set shellObj = CreateObject("WScript.Shell")
    cwd = shellObj.CurrentDirectory
    Set shellObj = Nothing  
End Sub

Private Function ParentDir(current_dir)
    ''' Return parent path of directory argument.
    Dim i, arr, tmpStr
    If instr(current_dir, "\") Then
        arr = Split(current_dir, "\")
        For i = 0 to Ubound(arr) - 1
            tmpStr = tmpStr & arr(i) & "\"
        Next
    End If
    ParentDir = tmpStr
End Function

Private Function ChDir(directory)
    ''' Change directory
    ''' Pass path of directory to change into as argument
    Dim shellObj
    Set shellObj = CreateObject("WScript.Shell")
    shellObj.CurrentDirectory = directory
    ' Wscript.Echo "Current Directory (After change): " & shellObj.CurrentDirectory
    Set shellObj = Nothing
End Function

Public Function CreatePath(root, filename)
    CreatePath = root  & "\" & filename
End Function

' Set cwd
Call SetCwd()



'' Set filepaths according to variables values at top of script
Dim xl_path, out_path, parent_path
out_path = CreatePath(cwd, out_name)

parent_path = ParentDir(cwd)
xl_path = CreatePath(parent_path, xl_name)




Dim xlApp, conn, wbConn, connCheck, xlWb
Dim tmpStr

Set xlApp = CreateObject("Excel.Application")
If DebugConfig = True Then
    xlApp.Visible = True
Else
    xlApp.Visible = False
End If


'### Open Workbook
Set xlWb = xlApp.Workbooks.Open(xl_path)
xlApp.ScreenUpdating = False

Dim conn_str, res
Dim output_str
Dim mt

For Each conn in xlWb.Connections
    ''' Is connection xlConnectionTypeOLEDB?
    If conn.Type = 1 Then
        For Each mt In xlWb.Model.ModelTables
            If mt.SourceWorkbookConnection = conn.Name Then
                output_str = output_str & "/***************************************************" & vbCrLf
                output_str = output_str & vbTab & "Table: " & mt.Name & vbCrLf
                output_str = output_str & vbTab & "Connection: " & conn.Name & vbCrLf
                output_str = output_str & "***************************************************/" & vbCrLf & vbCrLf
                output_str = output_str & conn.OLEDBConnection.CommandText & "`"
            End if
        Next
    End If
Next

''' Trim excess comma
output_str = Left(output_str, Len(output_str) - 1)


Dim fso_obj, obj_file, obj
Dim output_arr, i
Set fso_obj = CreateObject("Scripting.FileSystemObject")

' How to write file
output_arr = Split(output_str, "`")

Set obj_file = fso_obj.CreateTextFile(out_path, True)
' For Each obj in output_arr
For i = LBound(output_arr) To UBound(output_arr)
    obj_file.Write output_arr(i) & vbCrLf & vbCrLf & vbCrLf
Next

obj_file.Close
Set fso_obj = Nothing




xlApp.ScreenUpdating = True
Set xlWb = Nothing

xlApp.Quit
Set xlApp = Nothing
WScript.Quit(1)
