Option Explicit

'-------------------------------------------------------------------------
' Set array of processes to cease
' Add in comma-separated values to complete the array with whatever processes you want to end
' Ex. - Array("MSACCESS.EXE", "MSWORD.EXE")
' I have script to look for specific windows, but left that out of this.
Dim proc_array
'proc_array = Array("MSACCESS.EXE") ' Can also using single process in array.
proc_array = Array("<proc1>.exe","<proc2>.exe","<proc3>.exe")

'-------------------------------------------------------------------------
' Add quotes to string
Public Function DoubleQuote(str)
    DoubleQuote = Chr(34) & str & Chr(34)
End Function


'-------------------------------------------------------------------------
' Delay exit of script
' %comspec% points to cmd.exe
Public Function Delay(seconds)
    Delay = False
	Dim w_shell, command_str, temp_exec
	Set w_shell = WScript.CreateObject("WScript.Shell")
	command_str = w_shell.ExpandEnvironmentStrings("%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)")
	temp_exec = w_shell.Run(command_str, 0, True)
    Set w_shell = Nothing
    Delay = True
End Function


'-------------------------------------------------------------------------
' Determine if process is running
' Returns: 0 for False and -1 for True
Public Function Is_Proc_Running(process)

	Dim my_obj
	Dim procs
	Set my_obj = GetObject("winmgmts:")
	Set procs = my_obj.ExecQuery("select * from win32_process where name='" & process & "'")

	If procs.Count > 0 Then
		Is_Proc_Running = True
	Else
		Is_Proc_Running = False
	End If

	Set my_obj = Nothing
	Set procs = Nothing

End Function


'-------------------------------------------------------------------------
' Taskkill Function
' Pass image name (process) to run
Public Sub Kill(process_name)
	Dim my_obj, my_cmd, my_exec
    Set my_obj = WScript.CreateObject("WScript.Shell")
    my_cmd = my_obj.ExpandEnvironmentStrings("%COMSPEC% /C (taskkill /F /IM " & process_name & ")")
	'my_cmd = "cmd /c Taskkill /F /IM " & process_name & ""
	my_exec = my_obj.Run(my_cmd, 0, True)
	Set my_obj = Nothing
End Sub




'-------------------------------------------------------------------------
' Re-loop to ensure that process is killed
Dim out_bool, in_bool, i, zero_count, loop_delay
out_bool = True
in_bool = True
loop_delay = 15 ' Seconds to delay loop check iteration

' Do until out_bool is True
Do While out_bool
    zero_count = 0
    Do While in_bool
        For Each i in proc_array
            If Is_Proc_Running(i) = -1 Then
                Wscript.Echo "Killing: " & i
                zero_count = zero_count + 1
                Kill(i)
            End If
            in_bool = True
        Next
        Delay(loop_delay)
        in_bool = False
    Loop
    If zero_count > 0 Then
        in_bool = True
        out_bool = True
    Else
        in_bool = False
        out_bool = False
    End If
Loop
