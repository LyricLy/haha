Dim message, thicc, objShell
message = "fidget spinner"
Set thicc = CreateObject("SAPI.SpVoice")
thicc.Speak message
Set objShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep(1000)
objShell.Run "haha3.vbs"
