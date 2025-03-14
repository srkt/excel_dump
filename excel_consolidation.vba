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
    Dim i As Integer, j As Integer, k As Integer
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataStartRow As Long
    Dim targetRow As Long
    Dim hasData As Boolean
    Dim allColumnHeaders As Collection
    Dim headerMapping As Object
    Dim headerTitle As Variant
    Dim colIndex As Long
    Dim cellValue As Variant
    Dim useAutoDetect As Boolean
    Dim manualHeaderRow As Long
    Dim currentHeader As Variant
    
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
    tabsToImport = CInt(InputBox("How many tabs do you want to import from the source file?" & vbCrLf & _
                                 "(Source file has " & sourceWB.Worksheets.Count & " tabs)", "Select Number of Tabs", sourceWB.Worksheets.Count))
    On Error GoTo 0
    
    ' Validate user input
    If tabsToImport <= 0 Or tabsToImport > sourceWB.Worksheets.Count Then
        MsgBox "Invalid number of tabs. Operation cancelled.", vbExclamation
        sourceWB.Close False
        Exit Sub
    End If
    
    ' Ask if user wants to manually specify header row
    Dim headerChoice As VbMsgBoxResult
    headerChoice = MsgBox("Do you want to manually specify the header row index?" & vbCrLf & _
                          "Click Yes to specify a row number, or No to use automatic detection.", _
                          vbQuestion + vbYesNo, "Header Row Selection")
    
    useAutoDetect = (headerChoice = vbNo)
    
    If Not useAutoDetect Then
        ' Ask for header row index
        On Error Resume Next
        manualHeaderRow = CInt(InputBox("Enter the row number that contains headers (e.g., 3):", "Header Row Index", "3"))
        On Error GoTo 0
        
        If manualHeaderRow <= 0 Then
            MsgBox "Invalid row number. Using automatic detection instead.", vbExclamation
            useAutoDetect = True
        End If
    End If
    
    ' Create a collection to store all unique column headers
    Set allColumnHeaders = New Collection
    
    ' First pass: collect all unique column headers from all worksheets
    For i = 1 To tabsToImport
        Set sourceWS = sourceWB.Worksheets(i)
        
        ' Find the actual header row
        If useAutoDetect Then
            dataStartRow = FindHeaderRow(sourceWS)
        Else
            dataStartRow = manualHeaderRow
        End If
        
        ' Check if the worksheet has data
        hasData = CheckWorksheetHasData(sourceWS, dataStartRow)
        
        If hasData Then
            ' Find the last column with data
            lastCol = GetLastColumn(sourceWS, dataStartRow)
            
            ' Add each header to the collection if it's not already there
            For j = 1 To lastCol
                headerTitle = Trim(CStr(sourceWS.Cells(dataStartRow, j).Value)) ' Explicit conversion to String
                
                ' Only add non-empty headers
                If headerTitle <> "" Then
                    ' Check if this header is already in our collection
                    On Error Resume Next
                    allColumnHeaders.Add headerTitle, CStr(headerTitle)
                    On Error GoTo 0
                End If
            Next j
        End If
    Next i
    
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
    
    ' Add source tab identification header
    targetWS.Cells(targetRow, 1) = "Source Tab"
    targetRow = targetRow + 1
    
    ' Write all the unique headers to the first row of the target worksheet
    colIndex = 2  ' Start at 2 because column 1 is for Source Tab
    For Each currentHeader In allColumnHeaders
        targetWS.Cells(targetRow, colIndex) = currentHeader
        colIndex = colIndex + 1
    Next currentHeader
    
    ' Format the header row
    With targetWS.Rows(targetRow)
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With
    
    targetRow = targetRow + 1
    
    ' Second pass: Import data from each worksheet
    For i = 1 To tabsToImport
        Set sourceWS = sourceWB.Worksheets(i)
        
        ' Find the actual header row
        If useAutoDetect Then
            dataStartRow = FindHeaderRow(sourceWS)
        Else
            dataStartRow = manualHeaderRow
        End If
        
        ' Check if the worksheet has data
        hasData = CheckWorksheetHasData(sourceWS, dataStartRow)
        
        If hasData Then
            ' Find the last row and column with data
            lastRow = GetLastRow(sourceWS, dataStartRow)
            lastCol = GetLastColumn(sourceWS, dataStartRow)
            
            ' Copy data rows
            For j = dataStartRow + 1 To lastRow
                ' Add the sheet name as a source identifier
                targetWS.Cells(targetRow, 1) = sourceWS.Name
                
                ' Copy each cell to the appropriate column in the target worksheet
                For k = 1 To lastCol
                    headerTitle = Trim(CStr(sourceWS.Cells(dataStartRow, k).Value)) ' Explicit conversion to String
                    
                    If headerTitle <> "" Then
                        ' Get the value from the source cell
                        cellValue = sourceWS.Cells(j, k).Value
                        
                        ' Find the corresponding column in the target worksheet
                        colIndex = FindColumnIndex(targetWS, CStr(headerTitle), 2) ' Pass as String
                        
                        If colIndex > 0 Then
                            ' Set the value in the target worksheet
                            targetWS.Cells(targetRow, colIndex) = cellValue
                            
                            ' Copy formatting for numeric values if needed
                            If IsNumeric(cellValue) Then
                                targetWS.Cells(targetRow, colIndex).NumberFormat = sourceWS.Cells(j, k).NumberFormat
                            End If
                        End If
                    End If
                Next k
                
                targetRow = targetRow + 1
            Next j
            
            ' Add a blank row for separation between data from different tabs
            targetRow = targetRow + 1
        End If
    Next i
    
    ' Auto-fit columns for better readability
    targetWS.UsedRange.Columns.AutoFit
    
    ' Freeze the header row
    targetWS.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    ' Clean up
    targetWS.Cells(1, 1).Select
    Application.CutCopyMode = False
    sourceWB.Close False
    
    MsgBox "Data import completed successfully!" & vbCrLf & _
           "Columns with matching headers were consolidated across all tabs.", vbInformation
End Sub

' Function to find the column index with a specific header
Function FindColumnIndex(ws As Worksheet, headerToFind As String, headerRow As Long) As Long
    Dim lastCol As Long
    Dim i As Long
    Dim cellValue As String
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    FindColumnIndex = 0  ' Default to not found
    
    For i = 1 To lastCol
        cellValue = Trim(CStr(ws.Cells(headerRow, i).Value)) ' Explicit conversion to String
        If cellValue = headerToFind Then
            FindColumnIndex = i
            Exit Function
        End If
    Next i
End Function

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
