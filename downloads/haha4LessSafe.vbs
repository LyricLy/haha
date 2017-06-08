Dim message, thicc
message = "fidget spinner fidget spinner fidget spinner"
Set thicc = CreateObject("SAPI.SpVoice")
thicc.Speak message

WScript.Sleep 15000
Set WshShell = WScript.CreateObject("WScript.Shell")

While True
WshShell.SendKeys "hey guys"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 500
WshShell.SendKeys "my name is FIDGET SPINNER"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 500
thicc.Speak message
WshShell.SendKeys "and I want to spin "
WshShell.SendKeys "{ENTER}"
WScript.Sleep 500
WshShell.SendKeys "my nama jeff"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 5000
thicc.Speak message
WScript.Sleep 5000
Wend