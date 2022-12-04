Option Explicit

On Error Resume Next

'Wscript.Echo "Running write-to-txt.vbs"

Call TestWriteToFile

Sub TestWriteToFile()
	' https://stackoverflow.com/a/2198973/8213085'

	Const outFile = "...\vbs-testing\ops-process\test-txt.txt"
'	Wscript.Echo outFile

	Dim objFile
	Set objFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(outFile, True)
	objFile.Write "test string" & vbCrLf
	objFile.Close

	Wscript.Echo "Written to " & outFile

	set objFile = Nothing
End Sub