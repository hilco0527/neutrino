Option Explicit

'Variables Definition
Dim wsh, shellApp, sc, fso, folder, fullPath

'Instance Generation
Set wsh = CreateObject("WScript.Shell")
Set shellApp = CreateObject("Shell.Application")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

'Get the folder to put shorcuts.
Set folder = shellApp.BrowseForFolder(0,"Select the folder to put shortcuts."&vbCrLf&"The file path must be in the environment variables.",1)

'Copy the program into the selected folder
fullPath = folder.items.Item.Path & "\neutrino.vbs"
fso.CopyFile ".\neutrino.vbs", fullPath, True

'Generate the shortcut of neutrino.vbs into the selected folder
Set sc = wsh.CreateShortcut(folder.items.Item.Path & "\neu.lnk")
sc.TargetPath = fullPath
sc.save

'Generate the shortcut of neutrino.vbs into the "sendTo" folder
Set sc = wsh.CreateShortcut(wsh.SpecialFolders("sendTo") & "\neutrino.lnk")
sc.TargetPath = fullPath
sc.save

'Show the result message
If Err.Number = 0 Then
	MsgBox("Setup completed!")
Else
	MsgBox("Error:" & Err.Description)
End If

'Instance Release
Set wsh = Nothing
Set shellApp = Nothing
Set sc = Nothing
Set fso = Nothing