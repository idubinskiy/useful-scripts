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

Set logfile = objFSO.OpenTextFile(WorkingDir & "\piriformautoupdate.log", 8, True)

logfile.WriteLine("---------------Begin Script---------------")
logfile.WriteLine(Date & " " & Time & ":")

WshShell.CurrentDirectory = WorkingDir
logfile.WriteLine("Checking Piriform RSS Feed.")
WshShell.Run "curl -L -o piriform.xml -z piriform.xml http://feeds.feedburner.com/piriform?format=xml", 0, 1

Sub UpdateProgram(ProgramName)

	logfile.WriteLine("----------")
	logfile.WriteLine("Updating " & ProgramName & ".")

	If Not objFSO.FolderExists(ProgramFiles & "\" & ProgramName) Then
		logfile.WriteLine(ProgramName & " not installed!")
		logfile.WriteLine("----------")
		Exit Sub
	End If

	Set PiriformRss = objFSO.OpenTextFile("piriform.xml", 1)

	Set ProgramRegex = CreateObject("VBScript.RegExp")
	ProgramRegex.Pattern = ".title." & ProgramName & "\sv(.*?)(?=./title.)"

	Dim Matched
	Matched = False

	Do Until Matched = True Or PiriformRss.AtEndOfStream
		line = PiriformRss.ReadLine
		Set LineMatches = ProgramRegex.Execute(line)
		If LineMatches.Count > 0 Then
			Matched = True
			ProgramNewVersion = LineMatches.Item(0).SubMatches(0)
			logfile.WriteLine("Latest version of " & ProgramName & " is v" & ProgramNewVersion & ".")
		End If
	Loop
	
	If PiriformRss.AtEndOfStream Then
		logfile.WriteLine("No match for " & ProgramName & " in Piriform RSS feed. Checking release notes.")
		WshShell.Run "curl -L -o " & ProgramName & ".html -z " & ProgramName & ".html http://www.piriform.com/" & ProgramName & "/release-notes", 0, 1
		Set ProgramNotes = objFSO.OpenTextFile(ProgramName & ".html", 1)
		
		ProgramRegex.Pattern = ".strong.v(.*?)(?=./strong.)"
		
		Matched = False
		Do Until Matched = True Or ProgramNotes.AtEndOfStream
			line = ProgramNotes.ReadLine
			Set LineMatches = ProgramRegex.Execute(line)
			If LineMatches.Count > 0 Then
				Matched = True
				ProgramNewVersion = LineMatches.Item(0).SubMatches(0)
				Set VersionRegex = CreateObject("VBScript.RegExp")
				VersionRegex.Pattern = "^([0-9]+?\.[0-9]+?)\..*"
				ProgramNewVersion = VersionRegex.Replace(ProgramNewVersion, "$1")
				logfile.WriteLine("Latest version of " & ProgramName & " is v" & ProgramNewVersion & ".")
			End If
		Loop
		
		If ProgramNotes.AtEndOfStream Then
			logfile.WriteLine("No match for " & ProgramName & " in release notes!")
			logfile.WriteLine("----------")
			PiriformRss.Close
			Exit Sub
		End If
	End If

	ProgramCurrentVersion = objFSO.GetFileVersion(ProgramFiles & "\" & ProgramName & "\" & ProgramName & ".exe")
	logfile.WriteLine("Currently installed version of " & ProgramName & " is v" & ProgramCurrentVersion & ".")
	
	
	Set ProgramVersionRegex = CreateObject("VBScript.RegExp")
	ProgramVersionRegex.Pattern = "^" & ProgramNewVersion

	If  Not ProgramVersionRegex.Test(ProgramCurrentVersion) And ProgramCurrentVersion <> ProgramNewVersion Then
		logfile.WriteLine(ProgramName & " is out of date. Updating.")
		If objFSO.FolderExists(ProgramName) Then
			objFSO.DeleteFolder ProgramName
		End If
		objFSO.CreateFolder ProgramName
		WshShell.CurrentDirectory = WorkingDir & "\" & ProgramName & ""
		logfile.WriteLine("Downloading latest version of " & ProgramName & ".")
		WshShell.Run "curl -L -o " & ProgramName & ".zip http://www.piriform.com/" & ProgramName & "/download/portable/downloadfile", 0, 1
		logfile.WriteLine("Extracting latest version of " & ProgramName & ".")
		WshShell.Run "7z e " & ProgramName & ".zip", 0, 1
		WshShell.CurrentDirectory = ProgramFiles & "\" & ProgramName
		If objFSO.FileExists(ProgramName & ".exe.old") Then
			objFSO.DeleteFile ProgramName & ".exe.old"
		End If
		If objFSO.FileExists(ProgramName & "64.exe.old") Then
			objFSO.DeleteFile ProgramName & "64.exe.old"
		End If
		logfile.WriteLine("Making backup of installed version of " & ProgramName & ".")
		objFSO.CopyFile ProgramName & ".exe", ProgramName & ".exe.old"
		objFSO.CopyFile ProgramName & "64.exe", ProgramName & "64.exe.old"
		logfile.WriteLine("Updating " & ProgramName & ".")
		objFSO.CopyFile WorkingDir & "\" & ProgramName & "\" & ProgramName & ".exe", ProgramName & ".exe", 1
		objFSO.CopyFile WorkingDir & "\" & ProgramName & "\" & ProgramName & "64.exe", ProgramName & "64.exe", 1
		WshShell.CurrentDirectory = WorkingDir
		If objFSO.FolderExists(ProgramName) Then
			objFSO.DeleteFolder ProgramName
		End If
		logfile.WriteLine(ProgramName & " updated successfully!")
	Else
		logfile.WriteLine(ProgramName & " is up to date!")
	End If
	
	logfile.WriteLine("----------")

End Sub

ProgramList = Array("CCleaner","Recuva","Defraggler","Speccy")

For Each Program In ProgramList
	UpdateProgram Program
Next

logfile.WriteLine("---------------End Script---------------")
logfile.Close

End If