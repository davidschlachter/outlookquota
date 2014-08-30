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


' Check that 'SizerFiles' exists -- if not, create it
If Not oFSO.FolderExists(RootPath) Then
  oFSO.CreateFolder RootPath
End If


' Remove anything in the folder for a clean install
Dim D, F
Set D = oFSO.GetFolder(RootPath)
For Each F In D.Files
  oFSO.DeleteFile F.path, True
Next 


' Download the latest version of the plugin
Dim Url
Url = "https://schlachter.ca/david/EmailSizer.zip"
dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", Url, False
xHttp.Send
with bStrm
    .type = 1
    .open
    .write xHttp.responseBody
    .savetofile RootPath & "\EmailSizer.zip", 2
end with


' Extract the plugin archive files
Extract RootPath & "\EmailSizer.zip", RootPath
Sub Extract( ByVal myZipFile, ByVal myTargetDir )
    Dim intOptions, objShell, objSource, objTarget
    Set objShell = CreateObject( "Shell.Application" )
    Set objSource = objShell.NameSpace( myZipFile ).Items( )
    Set objTarget = objShell.NameSpace( myTargetDir )
    intOptions = 256
    objTarget.CopyHere objSource, intOptions
    Set objSource = Nothing
    Set objTarget = Nothing
    Set objShell  = Nothing
End Sub


' Create the registry keys
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\", ""
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\Description", "Shows quota usage", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\FriendlyName", "Quota Tool", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\LoadBehavior", 3, "REG_DWORD"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EmailSizer\Manifest", "file:///" & Replace(RootPath, "\", "/") & "/QuotaTool.vsto|vstolocal", "REG_SZ"


' Open Outlook again
objWshShell.Run "Outlook"