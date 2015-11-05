' Lantern.vbs  ( Lantern in Lan )
' @authors Jack Chan (fulicat@qq.com)
' @date    2015-10-31 ‏‎11:54:06
' @update  2015-11-05 22:02:58

' ========== Config ==============================
Dim LanternConfig
LanternConfig = "lantern-2.0.10.yaml" ' eg: lantern-2.0.10.yaml

Dim LanIP
LanIP = "127.0.0.1" ' eg: 172.28.43.3
' ========== End Config ===========================

On Error Resume Next

' Set WScript.Shell
Dim ws, wsPath
Set ws = WScript.CreateObject("WScript.Shell")
wsPath = ws.CurrentDirectory

' Get Computer Name & LAN IP address
Dim strComputer, IPConfigSet
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

'Get Lan IP
Set IPConfigSet = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
For Each objItem In IPConfigSet
	LanIP = objItem.IPAddress(0)
Next
Set objWMIService = Nothing
Set IPConfigSet = Nothing

' Replace 127.0.0.1 to LAN IP address in Lantern' config file
Dim fso, file, content, bakFile
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.fileExists(LanternConfig) Then
	Set file = fso.OpenTextFile(LanternConfig)
	content = file.ReadAll
	file.Close
	content = replace(content, "127.0.0.1:8787", LanIP + ":8787")
	content = replace(content, "127.0.0.1:16823", LanIP + ":16823")
	Set file = fso.CreateTextFile(LanternConfig)
	file.Write content
	file.Close
	bakFile = LanternConfig&".bak"
	If fso.fileExists(bakFile)=False Then
		Set file = fso.CreateTextFile(bakFile)
		file.Write content
		file.Close
	End If
	Set file = Nothing
Else
	MsgBox "Lantern's config file (eg: lantern-2.0.10.yaml) is not exists"& vbNewLine & vbNewLine &" or"& vbNewLine & vbNewLine &" You can edit this file (lantern.vbs) to setting", vbInformation
	ws.Run "notepad.exe "&wsPath &"\lantern.vbs", 1, Ture
	Set ws = Nothing
	WScript.Quit
End If

' Create Shortcut to Desktop
Dim Shortcut, ShortcutName, ShortcutPath, ShortcutLink
ShortcutName = "Lantern in LAN"
ShortcutPath = ws.SpecialFolders("Desktop")
ShortcutLink = ShortcutPath & "\" & ShortcutName & ".lnk"
If fso.fileExists(ShortcutLink)=False Then
	Set Shortcut = ws.CreateShortcut(ShortcutLink)
	Shortcut.TargetPath = wsPath & "\lantern.vbs"
	Shortcut.IconLocation = wsPath & "\lantern.exe,0"  
	Shortcut.WindowStyle = 1
	Shortcut.Hotkey = "CTRL+ALT+L"
	Shortcut.Description = "Lantern in LAN - by Jack.Chan"
	Shortcut.WorkingDirectory = wsPath
	Shortcut.Save
	Set Shortcut = Nothing
End If

' Get IE AutoConfigURL
Dim pacURL
pacURL = ws.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL")
'MsgBox pacURL

' Run Lantern
'ws.Run "lantern.exe"  ' default
'ws.Run "lantern.exe -startup"  ' startup without launch browser
'ws.Run "lantern.exe -headless=true"  ' startup without trayicon
'ws.Run "lantern.exe -clear-proxy-settings"
ws.Run "lantern.exe -startup"

Set ws = Nothing
WScript.Quit