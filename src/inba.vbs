Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

' Play audio
oPlayer.URL = "src\inba.mp3"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
    WScript.Sleep 100
Wend

' Release the audio file
oPlayer.close