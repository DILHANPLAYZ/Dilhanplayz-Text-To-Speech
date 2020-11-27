Dim message, sapi
message=InputBox("What do you want me to say?","DilhanPlayz Text To Speech")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message
