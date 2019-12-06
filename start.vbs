Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")

objShell.Run "src\speaker.vbs"
objShell.Run "src\cdroom.vbs"
objShell.run "src\disco.vbs"

Set objShell = Nothing