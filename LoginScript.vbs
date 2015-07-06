

' LoginScript
Set WshShell = WScript.CreateObject("WScript.Shell")

' delay to switch to application
WScript.Sleep 2000 

'enter your user name and password

WshShell.SendKeys "root"
WshShell.SendKeys "{TAB}"

WshShell.SendKeys "admin@123"
WshShell.SendKeys "{ENTER}"