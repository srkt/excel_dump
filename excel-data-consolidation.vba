Option Explicit

' Button click event that calls the main function
Sub ImportDataFromSourceFile()
    Dim sourcePath As String
    
    ' Prompt user to select the source Excel file
    sourcePath = GetSourceFilePath()
    
    If sourcePath = "" Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Call the function to fetch and consolidate data
    FetchDataFromSource sourcePath
End Sub

' Function to get the source file path using a file dialog
Function GetSourceFilePath() As String
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select Source Excel File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        
        If .Show = True Then
            GetSourceFilePath = .SelectedItems(1)
        Else
            GetSourceFilePath = ""
        End If
    End With
End Function

' Main function to fetch and consolidate data
Sub FetchDataFromSource(sourcePath As String)
    Dim sourceWB As Workbook
    Dim targetWB As Workbook
    Dim targetWS As Worksheet
    Dim sourceWS As Worksheet
    Dim tabsToImport As Integer
    Dim i As Integer
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataStartRow As Long
    Dim dataRange As Range
    Dim targetRow As Long
    Dim hasData As Boolean
    
    ' Set target workbook to the active workbook
    Set targetWB = ThisWorkbook
    
    ' Try to open the source workbook
    On Error Resume Next
    Set sourceWB = Workbooks.Open(sourcePath, ReadOnly:=True)
    On Error GoTo 0
    
    If sourceWB Is Nothing Then
        MsgBox "Failed to open the source file.", vbExclamation
        Exit Sub
    End If
    
    ' Ask the user how many tabs to import
    On Error Resume Next
    tabsToImport = InputBox("How many tabs do you want to import from the source file?" & vbCrLf & _
                           "(Source file has " & sourceWB.Worksheets.Count & " tabs)", "Select Number of Tabs", sourceWB.Worksheets.Count)
    On Error GoTo 0
    
    ' Validate user input
    If tabsToImport <= 0 Or tabsToImport > sourceWB.Worksheets.Count Then
        MsgBox "Invalid number of tabs. Operation cancelled.", vbExclamation
        sourceWB.Close False
        Exit Sub
    End If
    
    ' Create a new worksheet in the target workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    targetWB.Worksheets("Consolidated Data").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set targetWS = targetWB.Worksheets.Add
    targetWS.Name = "Consolidated Data"
    
    ' Initialize the target row counter
    targetRow = 1
    
    ' Add headers for identification
    targetWS.Cells(targetRow, 1) = "Source Tab"
    targetRow = targetRow + 1
    
    ' Process each worksheet in the source workbook
    For i = 1 To tabsToImport
        Set sourceWS = sourceWB.Worksheets(i)
        
        ' Find the actual header row (which starts at row 3, but we'll make it dynamic)
        dataStartRow = FindHeaderRow(sourceWS)
        
        ' Check if the worksheet has data
        hasData = CheckWorksheetHasData(sourceWS, dataStartRow)
        
        If hasData Then
            ' Find the last row and column with data
            lastRow = GetLastRow(sourceWS, dataStartRow)
            lastCol = GetLastColumn(sourceWS, dataStartRow)
            
            ' Copy the header row
            If dataStartRow > 0 Then
                sourceWS.Range(sourceWS.Cells(dataStartRow, 1), sourceWS.Cells(dataStartRow, lastCol)).Copy
                targetWS.Cells(targetRow, 1).PasteSpecial xlPasteValues
                targetWS.Cells(targetRow, lastCol + 1) = "Source: " & sourceWS.Name
                targetRow = targetRow + 1
                
                ' Copy the data rows
                If lastRow > dataStartRow Then
                    Set dataRange = sourceWS.Range(sourceWS.Cells(dataStartRow + 1, 1), sourceWS.Cells(lastRow, lastCol))
                    dataRange.Copy
                    targetWS.Cells(targetRow, 1).PasteSpecial xlPasteValues
                    targetRow = targetRow + lastRow - dataStartRow
                End If
                
                ' Add a blank row for separation
                targetRow = targetRow + 1
            End If
        End If
    Next i
    
    ' Auto-fit columns for better readability
    targetWS.UsedRange.Columns.AutoFit
    
    ' Clean up
    Application.CutCopyMode = False
    sourceWB.Close False
    
    MsgBox "Data import completed successfully!", vbInformation
End Sub

' Function to find the header row dynamically
Function FindHeaderRow(ws As Worksheet) As Long
    Dim i As Long
    Dim cellsWithContent As Long
    Dim maxRow As Long
    Dim maxContent As Long
    
    maxRow = 1
    maxContent = 0
    
    ' Check the first 10 rows to find the one with the most content
    ' Usually headers have more cells filled than other rows
    For i = 1 To 10
        If i <= ws.UsedRange.Rows.Count Then
            cellsWithContent = Application.CountA(ws.Rows(i))
            
            If cellsWithContent > maxContent Then
                maxContent = cellsWithContent
                maxRow = i
            End If
        End If
    Next i
    
    ' Default to row 3 if we can't determine a better row
    If maxRow = 1 And maxContent <= 1 Then
        maxRow = 3
    End If
    
    FindHeaderRow = maxRow
End Function

' Function to check if a worksheet has data after the header row
Function CheckWorksheetHasData(ws As Worksheet, headerRow As Long) As Boolean
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' If there are rows after the header, the sheet has data
    CheckWorksheetHasData = (lastRow > headerRow)
End Function

' Function to get the last row with data
Function GetLastRow(ws As Worksheet, headerRow As Long) As Long
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Return at least the header row
    GetLastRow = Application.WorksheetFunction.Max(headerRow, lastRow)
End Function

' Function to get the last column with data
Function GetLastColumn(ws As Worksheet, headerRow As Long) As Long
    Dim lastCol As Long
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    ' Ensure we have at least one column
    GetLastColumn = Application.WorksheetFunction.Max(1, lastCol)
End Function
