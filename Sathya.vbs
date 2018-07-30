
'setting up variables
Dim inputFilePath
Dim logFilePath
Dim logFile
Dim dateVariable
Dim SplitVariable
Dim fileDate
Dim latestInputFile
Dim isInputFileFound 
Dim inputFileFormat
Dim currDate 
Dim folder
Dim inputFile
Dim scriptPath
Dim xlo
Dim inputExcelFile
Dim EnabledIDColumnNo

'setting up default values
inputFilePath = "C:\Users\Praya\Desktop\Sathya\vbScript\Input\"
logFilePath = "C:\Users\Praya\Desktop\Sathya\vbScript\Logs\"
inputFileFormat = "input file"
scriptPath = "C:\Users\Praya\Desktop\Sathya\vbScript\"
xlCSV = 6

set fso = createObject("Scripting.FileSystemObject")
set xlo = createObject("Excel.Application")
xlo.Visible = False
EnabledIDColumnNo = 6

OpenLogFile()
StartProcess()
EndProcess()

Function StartProcess()
	GetLatestInputFile()
	OpenTempExcelFile()
	StartExcelProcess()
	CloseTempExcelFile()
End Function

Function OpenTempExcelFile()
	WriteToLogFile("Opening Temporary Excel file")
	set inputExcelFile = xlo.Workbooks.open(scriptPath & latestInputFile)
	
	'WriteToLogFile("Save File As CSV format")
	'set myWorkSheet = inputExcelFile.Worksheets("Sheet1")
	'myWorkSheet.SaveAs scriptPath & "CSVFileFormat.csv", 6
	'xlo.DisplayAlerts = False
	'WriteToLogFile("CSV File saved...")
End Function

Function StartExcelProcess()
	deleteDiabledIDs(EnabledIDColumnNo)
	DeleteNonPersonalIds()
	'CheckforUsenameBlanks()
End Function


Function DeleteNonPersonalIds()
	lastrow = inputExcelFile.Worksheets("Sheet1").UsedRange.Rows.Count
	i=2
	Do While  lastrow >= i
		
	Loop
End Function

Function deleteDiabledIDs(EnabledIDColumnNo)
	lastrow = inputExcelFile.Worksheets("Sheet1").UsedRange.Rows.Count
	WriteToLogFile("Row count in Input file is" & lastrow)
	i=2
	Do While lastrow >= i
		WriteToLogFile("Processing row " & i & inputExcelFile.Worksheets("Sheet1").Cells(i,EnabledIDColumnNo).Value)
		IF Not inputExcelFile.Worksheets("Sheet1").Cells(i,EnabledIDColumnNo).Value = "Enabled" Then
			Set SelectedRow = inputExcelFile.Worksheets("Sheet1").Cells(i,EnabledIDColumnNo).EntireRow
			lastrow = lastrow - 1
			WriteToLogFile("Row Deleted" & SelectedRow.Delete)
		Else
			i = i + 1
		End IF
	Loop
End Function

Function CloseTempExcelFile()
	inputExcelFile.Save
	inputExcelFile.close()
End Function

'Create new log file
Function OpenLogFile()
	IF fso.FolderExists(logFilePath) Then
		set logFile = fso.createTextFile(logFilePath &"Log File_"& myDateFormat(now()) &".txt",ForWriting,True)
	Else
		WriteToLogFile("Log File Path Error : "& logFilePath)
	End IF
	'set logFile = fso.OpenTextFile(logFile)
End Function

'Write the Statement to Log File
Function WriteToLogFile(Statement)
	logFile.Write Statement & vbCrlf
End Function

'Closing log File
Function CloseLogFile()
	logFile.close()
End Function

Function myDateFormat(myDate)
    dy = WhatEver(Day(myDate))
    mt = WhatEver(Month(myDate))    
    yr = Year(myDate)
	hr = WhatEver(Hour(myDate))
	mn = WhatEver(Minute(myDate))
	se = WhatEver(Second(myDate))
    myDateFormat= yr & mt & dy & "_5" & hr & mn & se
End Function

Function WhatEver(num)
    If(Len(num)=1) Then
        WhatEver="0"&num
    Else
        WhatEver=num
    End If
End Function

Function EndProcess()
	WriteToLogFile("Process end Closing task..")
	CloseLogFile()
	
	xlo.DisplayAlerts = False
	xlo.Visible = true
	
	set xlo = Nothing
	set fso = Nothing
	WScript.Quit
End Function

Function GetLatestInputFile()
	If fso.FolderExists(inputFilePath) = False Then
		WriteToLogFile("Input Path Error : "& inputFilePath)
		CloseScript()
	End IF
	set folder = fso.GetFolder(inputFilePath)
	set files = folder.Files
	WriteToLogFile("Looking for latest file")
	For each folderIndex  in files
		If InStr(folderIndex.Name,inputFileFormat) Then
			isInputFileFound = True
			splitVariable = Split(folderIndex.Name," ")
			dateVariable = splitVariable(2)
			Dim newDate
			fileDate = CDate("1970-01-01") 
			newDate = CDate(Mid(dateVariable,5,4)&"-"&Mid(dateVariable,1,2)&"-"&Mid(dateVariable,3,2))

			IF DateDiff("s",fileDate,newDate)>1 Then
				fileDate = newDate
				latestInputFile = folderIndex.Name
				End If
			End If
		Next
	IF isInputFileFound = True Then
		WriteToLogFile("Latest Input File is "&latestInputFile)
		CopyToLocalPath()
		Else
			WriteToLogFile("File Not Exist") 
		End If
End Function

Function CopyToLocalPath()
	set inputFile = fso.GetFile(inputFilePath & latestInputFile)
	inputFile.Copy scriptPath & latestInputFile, true
	WriteToLogFile("File Copied to new path")
End Function

