' Kill running VBScript task
Option Explicit
Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "taskkill /f /im Cscript.exe", , True
WshShell.Run "taskkill /f /im wscript.exe", , True
Set WshShell = Nothing
