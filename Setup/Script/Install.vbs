Option Explicit
On Error Resume Next

Sub RegistAssembly(wshShell)
	Dim commandLine

    commandLine = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe "'64bit compile
 	commandLine = commandLine & """C:\Program Files\Softcube\DesignTool\DesignTool.dll"""
	commandLine = commandLine & " /Codebase"

	wshShell.Run commandLine, 0, true

End Sub

Dim wshShell
Dim osVersion
Dim obj

Set obj = Wscript.CreateObject("Shell.Application")
if Wscript.Arguments.Count = 0 then
obj.ShellExecute "wscript.exe", WScript.ScriptFullName & " runas", "", "runas", 1
Wscript.Quit
end if

Set wshShell = CreateObject("WScript.Shell")

RegistAssembly wshShell

Set wshShell = Nothing
