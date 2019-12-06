Dim sapi

message="Zaczynamy INBE!"
message2="Jo zef tar ka ju zef"

Set sapi = CreateObject("sapi.spvoice")

sapi.Speak message

Do
sapi.Speak message2
Loop