Option Explicit

On Error Resume Next

Const sFile = "...\vbs-testing\ops-process\auto-updates-log.csv"

Call AppendToTextFile("Started running run-ops.vbs", sFile, True)
Call RunDailyOpsMI
Call AppendToTextFile("Finished running run-ops.vbs", sFile, True)


Sub RunDailyOpsMI()
'	On Error Goto ErrHandle

	Dim xlApp
	Dim xlBook

	Set xlApp = CreateObject("Excel.Application")
'	xlApp.Visible = True

	Set xlBook = xlApp.Workbooks.Open("...\vbs-testing\ops-process\Operations-Daily-MI.xlsm", , False)

'	xlApp.Run "'Operations-Daily-MI.xlsm'!ButtonAllUpdates"

	xlApp.Run "'Operations-Daily-MI.xlsm'!UpdateTMS"
	Call AppendToTextFile("Finished running 'Operations-Daily-MI.xlsm'!UpdateTMS", sFile, True)

	xlApp.Run "'Operations-Daily-MI.xlsm'!UpdateDeloitte"
	Call AppendToTextFile("Finished running 'Operations-Daily-MI.xlsm'!UpdateDeloitte", sFile, True)

	xlApp.Run "'Operations-Daily-MI.xlsm'!UpdateValidationAndHideSheets"
	Call AppendToTextFile("Finished running 'Operations-Daily-MI.xlsm'!UpdateValidationAndHideSheets", sFile, True)

	xlBook.Save
	Call AppendToTextFile("Finished saving Operations-Daily-MI.xlsm", sFile, True)

'	xlApp.Run "'Operations-Daily-MI.xlsm'!WriteNewEmail"
'	Call AppendToTextFile("Finished running 'Operations-Daily-MI.xlsm'!WriteNewEmail", sFile, True)

	xlBook.Close True
	xlApp.Quit

	Set xlBook = Nothing
	Set xlApp = Nothing

'	Exit Sub

'ErrHandle:
'	Call AppendToTextFile("Error in running RunDailyOpsMI", sFile, True)
End Sub


Sub AppendToTextFile(sAppend, sFile, bDatestamp)
	Dim objFile
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(sFile, 8)
	If bDatestamp Then
		objFile.Write WriteNow() & ": " & sAppend & vbCrLf
	Else
		objFile.Write sAppend & vbCrLf
	End If
	objFile.Close

	Set objFile = Nothing
End Sub


Function WriteNow()
	Dim dtTime
	dtTime = Now()
	WriteNow = "" _
		& Year(dtTime) & "-" _
		& LPad(Month(dtTime)) & "-" _
		& LPad(Day(dtTime)) & " " _
		& LPad(Hour(dtTime)) & ":" _
		& LPad(Minute(dtTime)) & ":" _
		& LPad(Second(dtTime))
End Function


Function LPad(str)
	LPad = Right("00" & str, 2)
End Function
