With CreateObject("Scripting.FileSystemObject")
   Dim RC6Path: RC6Path = .GetParentFolderName(WScript.ScriptFullName) & "\RC6.dll"
   Dim RegSvr:  RegSvr  = .GetSpecialFolder(0).Path & "\SysWOW64\regsvr32.exe" 
   If Not .FileExists(RC6Path) Then MsgBox "Couldn't find RC6.dll in:" & vbLF & .GetParentFolderName(WScript.ScriptFullName): WScript.Quit
   If Not .FileExists(RegSvr)  Then RegSvr = .GetSpecialFolder(0).Path & "\System32\regsvr32.exe" 
 
   CreateObject("Shell.Application").ShellExecute """" & RegSvr & """", """" & RC6Path & """", "", "runas", 1
End With