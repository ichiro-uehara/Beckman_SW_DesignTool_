Option Explicit
On Error Resume Next

Dim dllName
dllName = "MakeTitleBlock.dll"

'----------------------------------------------------------------------------
' 対象プログラムのパス
'----------------------------------------------------------------------------
' CustomActionData にはプログラムのパスが渡されていることを期待する
'----------------------------------------------------------------------------
'
Dim ProgramPath
ProgramPath = """" & Property("CustomActionData") & dllName & """"

'msgbox ProgramPath

Sub UnRegistAssembly(wshShell, dllName, programPath)
	Dim commandLine
	
	commandLine = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe "'64bit compile
	commandLine = commandLine & "C:\Softcube\DesignTool\DesignTool.dll"
	commandLine = commandLine & " /u"

    'msgbox commandLine

	wshShell.Run commandLine, 0, true

End Sub

Dim wshShell
Dim osVersion

Set wshShell = CreateObject("WScript.Shell")

UnRegistAssembly wshShell, dllName, ProgramPath

Set wshShell = Nothing

