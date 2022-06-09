Sub Macro1(var)
'
' Macro1 Macro
'

 

'
With ActiveSheet.QueryTables.Add(Connection:= _
"TEXT;" & var _
, Destination:=Range("$A$1"))
.CommandType = 0
.Name = "New Microsoft Excel Worksheet"
.FieldNames = True
.RowNumbers = False
.FillAdjacentFormulas = False
.PreserveFormatting = True
.RefreshOnFileOpen = False
.RefreshStyle = xlInsertDeleteCells
.SavePassword = False
.SaveData = True
.AdjustColumnWidth = True
.RefreshPeriod = 0
.TextFilePromptOnRefresh = False
.TextFilePlatform = 850
.TextFileStartRow = 1
.TextFileParseType = xlDelimited
.TextFileTextQualifier = xlTextQualifierDoubleQuote
.TextFileConsecutiveDelimiter = False
.TextFileTabDelimiter = False
.TextFileSemicolonDelimiter = False
.TextFileCommaDelimiter = True
.TextFileSpaceDelimiter = False
.TextFileColumnDataTypes = Array(2, 2, 2)
.TextFileTrailingMinusNumbers = True
.Refresh BackgroundQuery:=False
End With
Selection.AutoFilter
ActiveWorkbook.Worksheets("Planilha1").AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Planilha1").AutoFilter.Sort.SortFields.Add Key:= _
Range("A1:A553"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
:=xlSortNormal
With ActiveWorkbook.Worksheets("Planilha1").AutoFilter.Sort
.Header = xlYes
.MatchCase = False
.Orientation = xlTopToBottom
.SortMethod = xlPinYin
.Apply
End With
ActiveWorkbook.SaveAs Filename:= _
var & "1", FileFormat:=xlCSV _
, CreateBackup:=False
End Sub

 

Function CSVparaASC(x)
    on error resume next
    Dim xlAPP
        Set xlAPP = CreateObject("EXCEL.APPLICATION")
                               xlAPP.Visible = False
        xlAPP.DisplayAlerts = False
        Dim WBK
        Set WBK = xlAPP.Workbooks.Add
        Dim SHT
        Set SHT = WBK.Sheets(1)

    With SHT.QueryTables.Add("TEXT;" & x, SHT.Range("$A$1"))
    '.CommandType = 0
    .Name = Replace(x, ".csv", "", 1, -1, 1)
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = 1
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 850
    .TextFileStartRow = 1
    .TextFileParseType = 1
    .TextFileTextQualifier = 1
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 2, 2)
    .TextFileTrailingMinusNumbers = True
    .Refresh False
    End With
    
    SHT.UsedRange.AutoFilter
    SHT.AutoFilter.Sort.SortFields.Clear
    Dim LROW1
    LROW1 = SHT.Cells.Find("*", , -4163, 1, 1, 2).Row
    SHT.AutoFilter.Sort.SortFields.Add SHT.Range("A1:A" & CStr(LROW1)), 0, 1, 0
    With SHT.AutoFilter.Sort
        .Header = 1
        .MatchCase = False
        .Orientation = 1
        .SortMethod = 1
        .Apply
    End With

    WBK.SaveAs x, 6, , , , , , 2
    WBK.Close false

    If Err.Number <> 0 Then
        Dim res
        res = "ERRO, número do erro:" & CStr(Err.Number) & ", Descrição do erro:" & CStr(Err.Description)
        CSVparaASC = res
    Else
        CSVparaASC = "1"
    End If
  xlapp.Quit

end Function
