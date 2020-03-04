Attribute VB_Name = "Module1"
Sub fileCombine()

Dim oFSO As Object
Dim oFile As Object
Dim oFolder As Object
Dim fileList() As Variant
Dim fPath, testPath As String
Dim x, y, z As Long
Dim i, j, k As Long
Dim lastRow, lastColumn As Long
Dim dataLastRow, tableLastRow, storeRepLastRow As Long


Let fPath = Application.ThisWorkbook.Path & "\"
Let dataFilePath = fPath & "data_for_processing\"

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(dataFilePath)

ThisWorkbook.Worksheets("output").Cells.Clear

Let lastDataRow = 2

Application.ScreenUpdating = False
For Each oFile In oFolder.Files

    Filename = oFile.Name
    
    Workbooks.Open Filename:=dataFilePath & Filename
    
    With Workbooks(Filename).Worksheets(1)
        
        .Activate
        .UsedRange 'Refresh UsedRange
        lastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        lastCol = .UsedRange.Columns(.UsedRange.Columns.Count).Column
        colLetter = Split(Cells(1, lastCol).Address, "$")(1)
        Range("A1:" & colLetter & lastRow).Copy
        
    End With
    
    With ThisWorkbook.Worksheets("output")
    
        .Activate
        .Range("A" & lastDataRow).PasteSpecial
        .UsedRange 'Refresh UsedRange
        lastDataRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row + 2
    
    End With
    
    Application.DisplayAlerts = False
    Workbooks(Filename).Close
    Application.DisplayAlerts = True
Next oFile
Application.ScreenUpdating = True




End Sub
