Option Explicit

Dim objWshShell
Set objWshShell = CreateObject( "WScript.Shell" )
Dim RootPath
RootPath = objWshShell.ExpandEnvironmentStrings("%AppData%") & "\SizerFiles"
Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")


' Quit Outlook if it's open
Dim Process, strObject, strProcess, IsProcessRunning
Const strComputer = "." 
strProcess = "OUTLOOK.exe"
IsProcessRunning = False
strObject   = "winmgmts://" & strComputer
For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
If UCase( Process.name ) = UCase( strProcess ) Then
        Dim objOutlook
		Set objOutlook = CreateObject("Outlook.Application")
		objOutlook.Quit
    End If
Next


' Remove the EmailSizer files
Dim D, F
Set D = oFSO.GetFolder(RootPath)
For Each F In D.Files
  oFSO.DeleteFile F.path, True
Next 


' Remove the registry keys
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer", ""
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\Description", "Shows quota usage", "REG_SZ"
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\FriendlyName", "Quota Tool", "REG_SZ"
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\LoadBehavior", 3, "REG_DWORD"
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\Manifest", "file:///" & Replace(RootPath, "\", "/") & "/EmailSizer.vsto|vstolocal", "REG_SZ"


' Open Outlook again
objWshShell.Run "Outlook"