Set WshShell = CreateObject("WScript.Shell")
WshShell.Run("%windir%\system32\notepad.exe")
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Wscript.sleep 500
WshShell.SendKeys "^w ~  "+scriptdir+"\new.vbs ~"
