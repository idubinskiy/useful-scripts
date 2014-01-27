'---------------------------------------------------------------------------------
'The MIT License (MIT)
'
'Copyright (c) 2014 Akkifokkusu
'
'Permission is hereby granted, free of charge, to any person obtaining a copy of
'this software and associated documentation files (the "Software"), to deal in
'the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
'FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
'COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
'IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
'CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'---------------------------------------------------------------------------------

Set WshShell = WScript.CreateObject("WScript.Shell") 
If WScript.Arguments.length = 0 Then 
	Set ObjShell = CreateObject("Shell.Application") 
	ObjShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """" & " RunAsAdministrator", , "runas", 1 
Else

Set objFSO = CreateObject("Scripting.FileSystemObject")

ProgramFiles = WshShell.ExpandEnvironmentStrings("%programfiles%")

WorkingDir = objFSO.GetParentFolderName(objFSO.GetFile(Wscript.ScriptFullName))

Set logfile = objFSO.OpenTextFile(WorkingDir & "\filebotautoupdate.log", 8, True)

logfile.WriteLine("---------------Begin Script---------------")
logfile.WriteLine(Date & " " & Time & ":")

WshShell.CurrentDirectory = ProgramFiles & "\Filebot"

If objFSO.FileExists("Filebot.jar.old") Then
	objFSO.DeleteFile "Filebot.jar.old"
End If

logfile.WriteLine("Making backup of installed version of Filebot.")
objFSO.CopyFile "Filebot.jar", "Filebot.jar.old"

logfile.WriteLine("Updating Filebot.")
WshShell.Run "curl -L -O -z Filebot.jar http://sourceforge.net/projects/filebot/files/filebot/HEAD/FileBot.jar", 0, 1
logfile.WriteLine("Updated Filebot.")

logfile.WriteLine("---------------End Script---------------")
logfile.Close

End If