'
' This ia script to Run .exe as user + password
'
'

set WshShell = CreateObject("WScript.Shell")
set WshEnv = WshShell.Environment("Process")

WinPath = WshEnv("SystemRoot")&"\System32\runas.exe"
set FSO = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments 

currDir = FSO.GetAbsolutePathName(".")

sUser="Admin"
sPass="password_123"&VBCRLF
 sCmd= "notepad.exe"
  
if(args.Count>=1) then
	 sCmd  =  args(0)
End if

rc=WshShell.Run("runas /user:" & sUser & " " & CHR(34) & sCmd & CHR(34), 2, FALSE)
Wscript.Sleep 200 
WshShell.AppActivate(WinPath) 
WshShell.SendKeys sPass

set WshShell=Nothing
set WshEnv=Nothing
set FSO=Nothing

wscript.quit


