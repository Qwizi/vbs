Dim sapi

message="Zaczynamy INBE!"
message2="ju. zef. tar. K."

Set sapi = CreateObject("sapi.spvoice")

sapi.Speak message

Do
sapi.Speak message2
Loop