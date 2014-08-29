Set objFSO=CreateObject("Scripting.FileSystemObject")

'Open File / Get Filename
strFileName = CStr(SelectFile( ))
If strFileName = "" Then
    Wscript.Quit
End If

'Parse out Path
strPosition = InStrRev(strFileName,"\")
strPath = Left(strFileName,strPosition)
strFile = replace(strFileName,strPath,"")
'Parse out Filename
strPosition = InStr(strFile,".")
strFileCleaned = Left(strFile,strPosition)

'Alert User to Start
'WScript.Echo "Processing Started of """&strFile&""""&vbCrLf&vbCrLf&"You will be notified when complete."

'Create output Excel file name based on input name and current time
strOutputFileName=CStr(strFileCleaned&"output."&replace(FormatDateTime(now,2),"/","")&replace(FormatDateTime(now,4),":","")&LPad(Second(Now()),2,"0")&".xlsx")
strOutputFile=CStr(strPath&strOutputFileName)

'Create Excel Object
Set objExcel = CreateObject("Excel.Application")
Set objExcelSheet = createObject("Excel.sheet")
objExcel.Visible = False

'Track cell counters
row=1

'headers
objExcelSheet.ActiveSheet.cells(row,1).value="MRN / CPI"
objExcelSheet.ActiveSheet.cells(row,2).value="Name"
objExcelSheet.ActiveSheet.cells(row,3).value="???"
objExcelSheet.ActiveSheet.cells(row,4).value="C"
objExcelSheet.ActiveSheet.cells(row,5).value="G"
objExcelSheet.ActiveSheet.cells(row,6).value="Age"
objExcelSheet.ActiveSheet.cells(row,7).value="BirthDate"
objExcelSheet.ActiveSheet.cells(row,8).value="Case"
objExcelSheet.ActiveSheet.cells(row,9).value="Order"
objExcelSheet.ActiveSheet.cells(row,10).value="???"
objExcelSheet.ActiveSheet.cells(row,11).value="???"
objExcelSheet.ActiveSheet.cells(row,12).value="Diagnosis"

'make em bold!
for bold = 1 to 12
	objExcelSheet.ActiveSheet.cells(row,bold).Font.Bold = True
next

'Open Input File	 
Set objInputFile = objFSO.OpenTextFile(strFileName)
Do Until objInputFile.AtEndOfStream
	strLine = objInputFile.ReadLine
	strCode = trim(left(strLine,7))
	strLine = trim(replace(strLine,strCode,""))
	
	If strCode  = "000001|" Then
		row=row+1
	End If	
	If strCode  = "000002|" Then
		If lastRow = row Then
			row=row+1
		End If
		lastRow=row
	End If	
	If strCode = "000001|" Then
		fields=split(strLine,"|")
		for count = 0 to 6
			position = count+1
			objExcelSheet.ActiveSheet.cells(row,position).value=trim(fields(count))
		next
	ElseIf strCode = "000002|" Then
		fields=split(strLine,"|")
		for count = 0 to 4
			position = count+7
			objExcelSheet.ActiveSheet.cells(row,position).value=fields(count)
		next
	ElseIf strCode = "000003|" Then
		objExcelSheet.ActiveSheet.cells(row,12).value = objExcelSheet.ActiveSheet.cells(row,12).value&strLine
	End IF

	
Loop
'Close Both Files	
objInputFile.Close
objExcelSheet.SaveAs(strOutputFile)
objExcelSheet.Close
objExcel.Quit

'Alert User to End
WScript.Echo "Processing Complete of """&strFile&""""&vbCrLf&vbCrLf&"Output is in same folder as input file."&vbCrLf&vbCrLf&""""&strOutputFileName&""""


'User Defined Functions Below This Point ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Left Pad:  LPad(inputString,length2pad,padCharacter)
Function LPad(s, l, c)
  Dim n : n = 0
  If l > Len(s) Then n = l - Len(s)
  LPad = String(n, c) & s
End Function



Function SelectFile( )
    ' File Browser via HTA
    ' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
    ' Features: Works in Windows Vista and up (Should also work in XP).
    '           Fairly fast.
    '           All native code/controls (No 3rd party DLL/ XP DLL).
    ' Caveats:  Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    ' Source:   http://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15&ælig;-4ba3-bca5-ec349df65ef6

    Dim objExec, strMSHTA, wshShell

    SelectFile = ""

    ' For use in HTAs as well as "plain" VBScript:
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    ' For use in "plain" VBScript only:
    ' strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
    '          & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
    '          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )

    SelectFile = objExec.StdOut.ReadLine( )

    Set objExec = Nothing
    Set wshShell = Nothing
End Function
