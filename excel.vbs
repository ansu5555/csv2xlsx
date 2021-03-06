xlFileName = "MaxFrequency"
nameStartPos = 1
nameLen = 19

Const xlOpenXMLWorkbook = 51

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False
objExcel.Visible = False
folderPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = objFSO.GetFolder(folderPath)

'' Create new Excel workbook
newPath = folderPath & "\" & xlFileName & ".xlsx"
Set objWorkbook = objExcel.Workbooks.Add
objWorkbook.SaveAs newPath, xlOpenXMLWorkbook, , , , False

''Get all csv file names
fileNames = ""
Set Files = objFolder.Files
For Each File in Files
    If UCase(objFSO.GetExtensionName(File.name)) = "CSV" Then    
        fileNames = fileNames & "*" & File.name
	End If
Next

'' Rename the csvs       
arrFileNames = Split(fileNames, "*")
For Each fileName in arrFileNames
    if fileName <> "" Then
        newFileName = mid(fileName, 1, nameStartPos-1) & mid(fileName, nameStartPos + nameLen)
        objFSO.MoveFile fileName, newFileName
    End if
Next

For Each File in objFolder.Files
    If UCase(objFSO.GetExtensionName(File.name)) = "CSV" Then    
        
        '' Open the csv
        Set objCSV = objExcel.Workbooks.Open(File.Path, ReadOnly=True)
        Set objCSVSheet = objCSV.Worksheets(1)
        
        '' Copy the sheet from csv
        Set objWorkbookSheet = objWorkbook.Worksheets(objWorkbook.Sheets.Count)
        objCSVSheet.Copy objWorkbookSheet
        
        '' Rename the sheet
        'Set objWorkbookSheet = objWorkbook.Worksheets(objWorkbook.Sheets.Count-1)
        'objWorkbookSheet.Name = replace(objCSV.Name, ".csv", "")
        on error resume next
        objWorkbook.Save
        
        objCSV.Close
        
        Set objCSV = Nothing
        Set objCSVSheet = Nothing
        Set objWorkbookSheet = Nothing
        
    End If
Next

'' Rename the Sheet1 to Plot and move it to initial position
Set objWorkbookSheet = objWorkbook.Worksheets(objWorkbook.Sheets.Count)
objWorkbookSheet.Name = "Plot"
objWorkbookSheet.Move objWorkbook.Worksheets(1)

'' Apply filter to the sheets
For i=2 to objWorkbook.Sheets.Count
    objWorkbook.Worksheets(i).Range("A1").AutoFilter 6, "PASS",,,True
Next

on error resume next
objWorkbook.Save
objWorkbook.Close

Set objWorkbook = Nothing
Set objFolder = Nothing
on error resume next
objExcel.DisplayAlerts = True
objExcel.Quit
Set objExcel = Nothing
Set objFSO = Nothing
WScript.Echo "Done"
WScript.Quit
