Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")

objShell.Run "src\speaker.vbs"
objShell.Run "src\cdroom.vbs"
objShell.run "src\disco.vbs"
objShell.run "src\strona.vbs"
objShell.run "src\inba.vbs"
objShell.run "src\juzeph.vbs"

Set objShell = Nothing