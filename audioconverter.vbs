Dim message,sapi 
message=inputBox("audio converter conversion-Tech-hacks.org""Hover pc Hacks Text-to-audioconveter")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message 