Dim message, Mess2, sure, sapi
message= "What up? message 1"
Mess2="Message 2, is cool"
sure="sure"
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message
sapi.Speak sure
sapi.Speak Mess2
