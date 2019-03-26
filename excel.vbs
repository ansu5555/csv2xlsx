xlFileName = "new"
nameStartPos = 5
nameLen = 2

Const xlOpenXMLWorkbook = 51

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False
objExcel.Visible = False
Set objFolder = objFSO.GetFolder(objFSO.GetParentFolderName(WScript.ScriptFullName))

'' Create new file
newPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\" & xlFileName & ".xlsx"
Set objWorkbook = objExcel.Workbooks.Add
objWorkbook.SaveAs newPath, xlOpenXMLWorkbook, , , , False

Set Files = objFolder.Files
For Each File in Files
	If UCase(objFSO.GetExtensionName(File.name)) = "CSV" Then
		'' Open the csv
		Set objCSV = objExcel.Workbooks.Open(File.Path, ReadOnly=True)
		Set objCSVSheet = objCSV.Worksheets(1)
		
		'' Copy the sheet from csv
		Set objWorkbookSheet = objWorkbook.Worksheets(objWorkbook.Sheets.Count)
		'on error resume next
		objCSVSheet.Copy objWorkbookSheet
		
		'' Rename the sheet
		Set objWorkbookSheet = objWorkbook.Worksheets(objWorkbook.Sheets.Count-1)
		newSheetName = replace(objCSV.Name, mid(objCSV.Name, nameStartPos, nameLen), "")
		objWorkbookSheet.Name = newSheetName
		
		objWorkbook.Save
		
		Set objCSV = Nothing
		Set objCSVSheet = Nothing
		Set objWorkbookSheet = Nothing
		
    End If
Next

Set objWorkbook = Nothing
Set objFolder = Nothing
objExcel.DisplayAlerts = True
objExcel.Quit
Set objExcel = Nothing
Set objFSO = Nothing
WScript.Echo "Done"
WScript.Quit
