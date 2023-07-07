Sub CapturarMarcacoes()
    Dim filePath As String
    Dim targetWorkbook As Workbook
    Dim newWorkbook As Workbook
    Dim targetWorksheet As Worksheet
    Dim newWorksheet As Worksheet
    Dim note As comment
    Dim comment As Variant
    Dim note_text As String
    Dim newRow As Long
    Dim id As Long
    
    ' Get path file target
    filePath = ActiveSheet.Range("A2").Value
    ' Check cell A2 file path
    If Len(filePath) = 0 Then
        MsgBox "The target file path was not specified in cell A2.", vbExclamation
        Exit Sub
    End If
    
    Set targetWorkbook = GetWorkbook(filePath)
    
    ' Check if file is closed
    If targetWorkbook Is Nothing Then
        MsgBox "Could not open target file.", vbExclamation
        Exit Sub
    End If
    
    ' Create new excel file
    Set xls = CreateObject("Excel.Application")
    xls.DisplayAlerts = False
    Set newWorkbook = xls.Workbooks.Add
    
    ' Create new workbook
    Set newWorksheet = newWorkbook.Worksheets(1)
    ' Insert headers
    newWorksheet.Range("A1").Value = "ID"
    newWorksheet.Range("B1").Value = "Type"
    newWorksheet.Range("C1").Value = "Author"
    newWorksheet.Range("D1").Value = "Date"
    newWorksheet.Range("E1").Value = "Sheet"
    newWorksheet.Range("F1").Value = "Text"
    
    ' Initilialize counter
    id = 1
    newRow = 1
    WS_Count = targetWorkbook.Worksheets.Count
    
    For I = 1 To WS_Count
    ' Select target worksheet
    Set targetWorksheet = targetWorkbook.Worksheets(I)
    
    ' Loop each note
    For Each note In targetWorksheet.Comments
        
        ' Insert data
        newRow = newRow + 1
        newWorksheet.Cells(newRow, "A").Value = id
        newWorksheet.Cells(newRow, "B").Value = "Note"
        newWorksheet.Cells(newRow, "C").Value = note.Author
        newWorksheet.Cells(newRow, "D").Value = ""
        newWorksheet.Cells(newRow, "E").Value = targetWorksheet.Name
        newWorksheet.Cells(newRow, "F").Value = note.Text

        id = id + 1
    Next note
    
    ' Loop each comment
    For Each comment In targetWorksheet.CommentsThreaded
        
        ' Insert data
        newRow = newRow + 1
        newWorksheet.Cells(newRow, "A").Value = id
        newWorksheet.Cells(newRow, "B").Value = "Comment"
        newWorksheet.Cells(newRow, "C").Value = comment.Author.Name
        newWorksheet.Cells(newRow, "D").Value = comment.Date
        newWorksheet.Cells(newRow, "E").Value = targetWorksheet.Name
        newWorksheet.Cells(newRow, "F").Value = comment.Text
            
        id = id + 1
    Next comment
    Next I
    
    ' Save file
    newWorkbook.SaveAs targetWorkbook.Path & "\" & "Comments and notes.xlsx", AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    
    ' Close files
    newWorkbook.Close
    targetWorkbook.Close
    
    Set newWorksheet = Nothing
    Set newWorkbook = Nothing
    Set targetWorksheet = Nothing
    Set targetWorkbook = Nothing
    
    MsgBox "Tag capture completed successfully! The file with the marks is available in the excel file folder.", vbInformation
End Sub

Function GetWorkbook(WorkbookFullName As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.FullName = WorkbookFullName Then Exit For
    Next

    If wb Is Nothing Then
        If Len(Dir(WorkbookFullName)) > 0 Then
            Set wb = Workbooks.Open(WorkbookFullName)
        End If
    End If
    Set GetWorkbook = wb
End Function
