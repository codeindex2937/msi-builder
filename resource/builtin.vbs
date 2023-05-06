Function CreateShortcut()
	strDelim = Chr(10)
	strArgs = Session.Property("CustomActionData")
	arrArgs = Split(strArgs, strDelim)

	strPath = arrArgs( 0 )
	strProgramTitle = arrArgs( 1 )
	strProgram = arrArgs( 2 )
	strWorkDir = arrArgs( 3 )

	Set fso = CreateObject("Scripting.FileSystemObject")
	strLinkPath = fso.BuildPath(strPath, strProgramTitle & ".lnk")

	If fso.FileExists(strLinkPath) Then
		Exit Function
	End If

	Dim objShortcut, objShell
	Set objShell = CreateObject("Wscript.Shell")
	Set objShortcut = objShell.CreateShortcut(strLinkPath)
	objShortcut.TargetPath = strProgram
	objShortcut.WorkingDirectory = strWorkDir
	objShortcut.Description = strProgramTitle
	objShortcut.Save
End Function

Function LaunchApp()
	Set oShell = CreateObject("WScript.Shell")
	oShell.Exec Session.Property("LaunchAppPath")
End Function
