Attribute VB_Name = "CrashReports"
Option Explicit

' Builds the four monthly crash reports from the "Vehicle Crash Data" sheet.
' Each report goes on its own sheet. Existing report sheets with the same name
' are deleted and rebuilt, so run this on a copy of the workbook when testing.
'
' Run everything:  Run_All_Reports
' Run one report:  Report_Crash_Times, Report_Unique_Addresses,
'                  Report_Total_Injured_Dead, Report_Contributing_Factors
'
' Columns are found by their header text in row 1, and the last row is taken
' from the used range, so a longer or shorter month of data needs no edits.

Private Const DATA_SHEET As String = "Vehicle Crash Data"
Private Const SHEET_TIMES As String = "Crash_times"
Private Const SHEET_ADDRESSES As String = "Unique Addresses"
Private Const SHEET_INJURED As String = "Total Injured and Dead"
Private Const SHEET_FACTORS As String = "Count of Contributing_Factors"


Public Sub Run_All_Reports()
    Dim src As Worksheet
    Set src = GetDataSheet()
    If src Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Report_Crash_Times
    Report_Unique_Addresses
    Report_Total_Injured_Dead
    Report_Contributing_Factors
    Application.ScreenUpdating = True

    src.Activate
    MsgBox "Four reports rebuilt.", vbInformation
End Sub


' Report 1: crashes by hour of the day, with a line chart
Public Sub Report_Crash_Times()
    Dim src As Worksheet, ws As Worksheet
    Dim colTime As Long, lastRow As Long, fld As String
    Dim pc As PivotCache, pt As PivotTable
    Dim shp As Shape

    Set src = GetDataSheet()
    If src Is Nothing Then Exit Sub
    colTime = HeaderCol(src, "CRASH TIME")
    If colTime = 0 Then Exit Sub
    lastRow = LastDataRow(src)

    Set ws = NewReportSheet(SHEET_TIMES)
    src.Range(src.Cells(1, colTime), src.Cells(lastRow, colTime)).Copy ws.Range("A1")
    Application.CutCopyMode = False
    ws.Columns("A").ColumnWidth = 15.6
    fld = ws.Range("A1").Value

    Set pc = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ws.Range("A1:A" & lastRow), _
        Version:=6)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("D7"), _
        TableName:="CrashTimesPivot", _
        DefaultVersion:=6)

    With pt.PivotFields(fld)
        .Orientation = xlRowField
        .Position = 1
        .AutoGroup                              ' groups the times by hour
        .Orientation = xlHidden
    End With
    pt.AddDataField pt.PivotFields(fld), "Count of " & fld, xlCount
    pt.ColumnGrand = False                      ' no grand total row in the chart

    ' Hide the "<1/0/1900" bucket if the grouping created one
    On Error Resume Next
    pt.PivotFields("Hours").PivotItems("<1/0/1900").Visible = False
    On Error GoTo 0

    Set shp = ws.Shapes.AddChart2(227, xlLine)
    shp.Left = ws.Range("G2").Left
    shp.Top = ws.Range("G2").Top
    shp.Width = 450
    shp.Height = 260
    With shp.Chart
        .SetSourceData pt.TableRange1
        .HasTitle = True
        .ChartTitle.Text = "Crashes 24Hr Period"
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14
        .ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(89, 89, 89)
    End With

    ws.Range("K33").Select
End Sub


' Report 2: unique on-street names, sorted A to Z
Public Sub Report_Unique_Addresses()
    Dim src As Worksheet, ws As Worksheet
    Dim colStreet As Long, lastRow As Long

    Set src = GetDataSheet()
    If src Is Nothing Then Exit Sub
    colStreet = HeaderCol(src, "ON STREET NAME")
    If colStreet = 0 Then Exit Sub
    lastRow = LastDataRow(src)

    Set ws = NewReportSheet(SHEET_ADDRESSES)
    src.Range(src.Cells(1, colStreet), src.Cells(lastRow, colStreet)).Copy ws.Range("A1")
    Application.CutCopyMode = False
    ws.Columns("A").ColumnWidth = 32.5

    ' Row 1 is the header. Duplicates are compared without regard to case.
    ws.Range("A1:A" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    ' Empty cells sort to the bottom, so the one blank entry that survives
    ' the duplicate removal ends up below the last street name.
    ws.Range("A1:A" & lastRow).Sort Key1:=ws.Range("A1"), _
        Order1:=xlAscending, Header:=xlYes, MatchCase:=False

    ws.Range("A1").Select
End Sub


' Report 3: column totals for the injured and killed counts
Public Sub Report_Total_Injured_Dead()
    Dim src As Worksheet, ws As Worksheet
    Dim colFirst As Long, colLast As Long, lastRow As Long
    Dim i As Long, n As Long

    Set src = GetDataSheet()
    If src Is Nothing Then Exit Sub
    colFirst = HeaderCol(src, "NUMBER OF PERSONS INJURED")
    colLast = HeaderCol(src, "NUMBER OF MOTORIST KILLED")
    If colFirst = 0 Or colLast = 0 Then Exit Sub
    lastRow = LastDataRow(src)
    n = colLast - colFirst + 1

    Set ws = NewReportSheet(SHEET_INJURED)
    ' Row 1 = headers, row 2 = totals of rows 2 to lastRow, stored as values
    src.Range(src.Cells(1, colFirst), src.Cells(1, colLast)).Copy ws.Range("A1")
    Application.CutCopyMode = False
    For i = 1 To n
        ws.Cells(2, i).Value = Application.WorksheetFunction.Sum( _
            src.Range(src.Cells(2, colFirst + i - 1), src.Cells(lastRow, colFirst + i - 1)))
    Next i
    ws.Columns(1).Resize(, n).AutoFit

    ws.Range("A1").Select
End Sub


' Report 4: crashes per contributing factor (vehicle 1), highest to lowest
Public Sub Report_Contributing_Factors()
    Const HEADER As String = "CONTRIBUTING FACTOR VEHICLE 1"

    Dim src As Worksheet, ws As Worksheet
    Dim colFactor As Long, lastRow As Long
    Dim pc As PivotCache, pt As PivotTable

    Set src = GetDataSheet()
    If src Is Nothing Then Exit Sub
    colFactor = HeaderCol(src, HEADER)
    If colFactor = 0 Then Exit Sub
    lastRow = LastDataRow(src)

    Set ws = NewReportSheet(SHEET_FACTORS)
    src.Range(src.Cells(1, colFactor), src.Cells(lastRow, colFactor)).Copy ws.Range("A1")
    Application.CutCopyMode = False

    Set pc = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ws.Range("A1:A" & lastRow), _
        Version:=6)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("D7"), _
        TableName:="ContributingFactorsPivot", _
        DefaultVersion:=6)

    With pt.PivotFields(HEADER)
        .Orientation = xlRowField
        .Position = 1
    End With
    pt.AddDataField pt.PivotFields(HEADER), "Count of " & HEADER, xlCount
    pt.PivotFields(HEADER).AutoSort xlDescending, "Count of " & HEADER

    ws.Range("E15").Select
End Sub


' ---------- helpers ----------

Private Function GetDataSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & DATA_SHEET & "' not found.", vbExclamation
    End If
    Set GetDataSheet = ws
End Function

Private Function HeaderCol(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim c As Variant
    c = Application.Match(header, ws.Rows(1), 0)
    If IsError(c) Then
        MsgBox "Header '" & header & "' not found in row 1 of " & ws.Name & ".", vbExclamation
        HeaderCol = 0
    Else
        HeaderCol = CLng(c)
    End If
End Function

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    LastDataRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
End Function

' Deletes an existing sheet with this name, then adds a fresh one at the end
Private Function NewReportSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ActiveWorkbook.Worksheets.Add( _
        After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = sheetName
    Set NewReportSheet = ws
End Function
