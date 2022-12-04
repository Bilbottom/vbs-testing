Option Explicit

On Error Resume Next

'Wscript.Echo "Running append-to-txt.vbs"

Const sFile = "...\vbs-testing\ops-process\test-txt.csv"

Call AppendToTextFile("Some text 1", sFile, True)
Call AppendToTextFile("Some text 2", sFile, True)
Call AppendToTextFile("Some text 3", sFile, True)


Sub AppendToTextFile(sAppend, sFile, bDatestamp)
	Dim objFile
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(sFile, 8)
	If bDatestamp Then
		objFile.Write Now() & ": " & sAppend & vbCrLf
	Else
		objFile.Write sAppend & vbCrLf
	End If
	objFile.Close

'	Wscript.Echo "Written to " & outFile

	Set objFile = Nothing
End Sub
