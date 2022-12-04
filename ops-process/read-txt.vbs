Option Explicit


' https://stackoverflow.com/a/2198973/8213085

Dim objFSO
Dim objFile
Dim strFile
Dim strLine

Set objFSO = CreateObject("Scripting.FileSystemObject")

'How to read a file
strFile = "...\vbs-testing\ops-process\mylog.csv"
Set objFile = objFSO.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    Wscript.Echo strLine
Loop
objFile.Close
