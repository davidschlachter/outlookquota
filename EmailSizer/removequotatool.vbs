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
If oFSO.FolderExists(RootPath) Then
    Dim D, F
    Set D = oFSO.GetFolder(RootPath)
    For Each F In D.Files
        oFSO.DeleteFile F.path, True
    Next 
End If

On Error Resume Next
' Remove the registry keys
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer"
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\Description"
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\FriendlyName"
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\LoadBehavior"
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\Manifest"
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\"
objWshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\VSTA\Solutions\29ec26ae-4c91-439c-b860-80c7cf48fb96\"


' Open Outlook again
objWshShell.Run "Outlook"
