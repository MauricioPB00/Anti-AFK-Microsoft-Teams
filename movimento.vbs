set movimento = CreateObject("WScript.Shell")
Do
 WScript.Sleep(3*60*1000)
 movimento.SendKeys("{F13}")
Loop
