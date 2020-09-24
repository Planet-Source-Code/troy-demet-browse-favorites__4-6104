<div align="center">

## Browse Favorites


</div>

### Description

Using the Windows Scripting Host this VBScript retrieves the users favorites folder and loads the url links into an array, then goes to each site for three minutes.
 
### More Info
 
User can input how many sites they wish to browse.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Troy Demet](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/troy-demet.md)
**Level**          |Beginner
**User Rating**    |3.1 (44 globes from 14 users)
**Compatibility**  |VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/troy-demet-browse-favorites__4-6104/archive/master.zip)





### Source Code

```
'==========================================================================
'
' VBScript Source File --
'
' NAME: favoritesURL.vbs
'
' AUTHOR: Troy Allen Demet , TechnoGeek, Inc.
' DATE : 2/25/00
'
' COMMENT: This script will put the url of your favorites into an array
'			and then browse to each web site at 3 minute intervals.
'
'==========================================================================
Option Explicit
	Dim objShell, objWshShell, fso,fld, objFiles
	Dim urlUpper, urlLower, Folder, j, ie, arURL(), fileCount, howMany
	'Dim objFolder, file, count, fileType, holder
	Set objShell = WScript.CreateObject("Shell.Application")
	Set objWshShell = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Folder = objWshShell.SpecialFolders	("Favorites")
	Set fld = fso.GetFolder(Folder)
	set objFiles = fld.Files
	fileCount = objFiles.Count
	ReDim arURL(fileCount)
	howMany = InputBox("Please enter how many sites you wish to browse.","How Many?",10)
	If howMany < 1 Then
		WScript.Quit
	End If
	getFile(Folder)
	urlUpper = UBound(arURL)				' Upper bound of arURL
	urlLower = LBound(arURL)				' Lower bound of arURL
	If urlUpper < 1 Then
		Msgbox "Sorry nothing to show",,"Nothing to Show"
		WScript.Quit
	End IF
	If howMany > urlUpper Then
		howMany = urlUpper - 1
	End If
	' Create the ie object (Internet Explorer)
	Set ie = CreateObject("InternetExplorer.Application")
	' Set the properties of Internet Explorer
	With ie
		.left 		= 100
		.top 		= 100
		.height		= 460
		.width		= 620
		.menubar	= 0						' False
		.toolbar	= 0						' False
		.visible	= 1						' True
	End With
	' Loop through the array
	For j = urlLower to howMany
		if arURL(j) <> "" Then
			goUrl(arURL(j))
		End If
	Next
	MsgBox "Quitting getFiles script"
	' Clean up after yourself
	ie.Quit
	Set ie = Nothing
	WScript.Quit
Function readFile(filePath)
	On Error Resume Next
	Dim fileObject
	Dim link, shellObject, line
	Set fileObject = CreateObject("Scripting.FileSystemObject")
	Set shellObject = CreateObject("Wscript.Shell")
	Set link = shellObject.CreateShortcut(filePath)
	' Use the MsgBox for debugging
	 'MsgBox "temp" & vbCrLf & Link & vbCrLf & link.TargetPath
	' Return the value
	readFile = link.TargetPath
End Function
Function goURL(aURL)
	' go to the web site
	ie.navigate(aURL)
	'Wait 3 minutes
	WSCript.Sleep(180000)
End Function
Sub getFile (dir)
	Dim objFolder, objSubFolder, objFiles, objSubFiles, Folder, subFolder, File, subFileCount, count
	Dim fileType
	Set objFolder = fso.GetFolder(dir)
	Set objSubFolder = objFolder.SubFolders
	Set objFiles = objFolder.Files
	For Each Folder in objSubFolder
		Set subFolder = fso.GetFolder(Folder)
		Set objSubFiles = subFolder.Files
		subFileCount = objSubFiles.Count
		fileCount = fileCount + subFileCount
		ReDim Preserve arURL(fileCount)
		getFile(Folder)
	Next
	File = 0
	count = 0
	For Each File in objFiles
		fileType = File.Type
		' Want only *.url files
		if fileType = "Internet Shortcut" Then
			'MsgBox "fullPath" & vbCrLf & File.Path
			arURL(count) = readFile(File.Path)
		End If
		count = count + 1
	Next
End Sub
```

