Option Explicit

'Variables Definition
Dim wsh, fso, linkName, folderName, fullPath, sc, ans, target

'Instance Generation
Set wsh = CreateObject("WScript.Shell")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

'Get Current Directory
folderName = fso.getParentFolderName(WScript.ScriptFullName)

'Get Target Link Path
'(1)If argument exist, use it
If WScript.Arguments.Count <> 0 Then
	target = WScript.Arguments(0)
'(2)If argument does not exit, use clip board
Else
	Dim ObjHTML
	Set objHTML = CreateObject("htmlfile")
	target = Trim(objHTML.ParentWindow.ClipboardData.GetData("text"))
End If

'Get Shortcut Name
Do
	linkName = InputBox("Please enter the shortcut name", "neutrino")

	'(1)If the shortcut name is empty, end this program
	If linkName = "" Then
		Exit Do
	
	'(2)If the shortcut name is inputted, generate the full path of shortcut
	Else
		fullPath = folderName & "\" & linkName & ".lnk"
	End If

'Create Shortcut File
	'(1)If the shortcut name already exists, confirm the overwrite
	If fso.FileExists(fullPath) = True Then
		ans = MsgBox("The shortcut name already exist. Do you overwite?", vbYesNoCancel, "neutrino")
		'i)If Yes, create shortcut and end this program
		If ans = vbYes Then
			Set sc = wsh.CreateShortcut(fullPath)
			sc.TargetPath = target
			sc.save
			Exit Do
		'ii)If No, show the input box again
		ElseIf ans = vbNo Then

		'iii)If Cancel, end this program
		Else
			Exit Do
		End If
	'(2)If the shortcut name does not exist, create the shortcut fule and end this program
	Else
		Set sc = wsh.CreateShortcut(fullPath)
		sc.TargetPath = target
		sc.save
		Exit Do
	End If
Loop

'Instance Release
Set wsh = Nothing
Set fso = Nothing
Set sc = Nothing