Dim wshShell, cmd

cmd = "...\vbs-testing\ops-process\venv\Scripts\python.exe ""...\vbs-testing\ops-process\test-py.py"""
WScript.Echo cmd

'Set wshShell = CreateObject("WScript.Shell")
CreateObject("WScript.Shell").Run cmd