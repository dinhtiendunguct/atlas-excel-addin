Attribute VB_Name = "AtlasTests"
Option Explicit
Private checks As Long
Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise 5, "AtlasAcceptance", message
    checks = checks + 1
End Sub
Private Function ScheduleError(ByVal book As Workbook) As String
    On Error GoTo Failed
    CalculateProject book, False
    Exit Function
Failed:
    ScheduleError = Err.Description
End Function
Public Function AtlasAcceptance(ByVal outputPath As String) As String
    On Error GoTo Failed
    Dim book As Workbook, w As Worksheet, r As Long, oldStart As Variant, previousCost As Double, expected As Double
    Dim relation As Variant, offsets As Variant, i As Long, catalog As Object
    checks = 0
    CreateProject True, True
    Set book = ActiveWorkbook: Set w = ProjectSheet(book)
    Check book.Worksheets.Count = 11, "Eleven project sheets"
    Check w.Shapes.Count > 27, "Demo Gantt shapes"
    Check book.Worksheets("BAO_CAO").ChartObjects.Count = 1, "S curve chart"
    Check book.Worksheets("BIEU_DO_NL").ChartObjects.Count = 3, "Resource charts"
    Check book.Worksheets("BAO_CAO").Cells(16, 6).Value > 0, "Actual snapshot point"
    Check w.Cells(18, 13).Value = 2, "Demo baseline variance"
    book.SaveAs outputPath, xlOpenXMLWorkbook
    ' Use the generated project as a controlled fixture, then close without saving.
    w.Range("A9:Q2008").ClearContents
    w.Range("B4") = DateSerial(2026, 10, 5): w.Range("B5") = DateSerial(2026, 10, 5)
    For r = 9 To 12
        w.Cells(r, 1) = Chr(65 + r - 9): w.Cells(r, 2) = CStr(r - 8): w.Cells(r, 3) = "Test"
        w.Cells(r, 4) = "CV": w.Cells(r, 10) = 0
    Next r
    w.Cells(9, 5) = 3: w.Cells(10, 5) = 6: w.Cells(11, 5) = 4: w.Cells(12, 5) = 10
    w.Cells(10, 8) = "A:FS:0": w.Cells(11, 8) = "A:FS:0": w.Cells(12, 8) = "B:FS:0;C:FS:0"
    CalculateProject book, False
    Check w.Cells(10, 6).Value = DateSerial(2026, 10, 8), "FS working day start"
    Check w.Cells(12, 6).Value = DateSerial(2026, 10, 15), "Merge network start"
    Check w.Cells(11, 12).Value = 2, "Parallel branch float"
    Check w.Range("J4").Value = 19, "Project duration 19"
    oldStart = w.Cells(10, 6).Value
    w.Cells(9, 8) = "D:FS:0"
    Check Len(ScheduleError(book)) > 0, "Cycle rejected"
    Check w.Cells(10, 6).Value = oldStart, "Invalid network preserves outputs"
    w.Cells(9, 8).ClearContents
    w.Cells(10, 1) = "A"
    Check Len(ScheduleError(book)) > 0, "Duplicate ID rejected"
    w.Cells(10, 1) = "B": w.Cells(10, 8) = "UNKNOWN:FS:0"
    Check Len(ScheduleError(book)) > 0, "Unknown predecessor rejected"
    w.Cells(10, 8) = "B:FS:0"
    Check Len(ScheduleError(book)) > 0, "Self reference rejected"
    w.Cells(10, 8) = "A:FS:0": w.Cells(10, 10) = 101
    Check Len(ScheduleError(book)) > 0, "Progress out of range rejected"
    w.Cells(10, 10) = 0: w.Cells(10, 9) = "NO_SUCH_RESOURCE:1"
    Check Len(ScheduleError(book)) > 0, "Unknown resource rejected"
    w.Cells(10, 9).ClearContents
    book.Worksheets("CAU_HINH").Range("A12") = DateSerial(2026, 10, 6)
    CalculateProject book, False
    Check w.Cells(10, 6).Value = DateSerial(2026, 10, 9), "Holiday skipped"
    book.Worksheets("CAU_HINH").Range("A12").ClearContents
    book.Worksheets("CAU_HINH").Range("B6") = "0000011"
    CalculateProject book, False
    Check w.Cells(12, 6).Value = DateSerial(2026, 10, 16), "Five day working week"
    book.Worksheets("CAU_HINH").Range("B6") = "0000001"
    w.Range("A11:Q12").ClearContents: w.Cells(9, 5) = 5: w.Cells(10, 5) = 3
    relation = Array("A:SS:2", "A:FF:1", "A:SF:7", "A:FS:-1")
    offsets = Array(2, 3, 4, 4)
    For i = 0 To 3
        w.Cells(10, 8) = relation(i): CalculateProject book, False
        Check w.Cells(10, 6).Value = DateSerial(2026, 10, 5 + offsets(i)), "Relationship " & relation(i)
    Next i
    w.Cells(10, 4) = "MOC": w.Cells(10, 5) = 0: w.Cells(10, 8) = "A:FS:0"
    CalculateProject book, False
    Check w.Cells(10, 6).Value = w.Cells(10, 7).Value, "Zero day milestone"
    w.Cells(10, 4) = "CV": w.Cells(10, 5) = 3: w.Cells(10, 9) = "NC:2"
    CalculateProject book, False
    w.Range("P9:Q10").Value = w.Range("F9:G10").Value
    w.Cells(10, 5) = 4: CalculateProject book, False
    Check w.Cells(10, 13).Value = 1, "Baseline finish variance"
    BuildReports book, False
    expected = 4 * 2 * 450000
    Check Application.Sum(book.Worksheets("BIEU_DO_NL").Range("J6:J20")) = expected, "Resource cost reconciles"
    Check book.Worksheets("BIEU_DO_NL").Range("K6").Value = expected, "Cumulative cost reconciles"
    w.Range("B6") = "THANG": DrawGantt book
    Check w.Shapes.Count > 0, "Monthly Gantt"
    w.Range("B6") = "NGAY": DrawGantt book
    Check w.Shapes.Count > 0, "Daily Gantt"
    book.Worksheets("CAU_HINH").Range("B8") = "CO": DrawGantt book
    Check w.Shapes("Atlas_Link_1").Connector = msoTrue, "Dependency connector"
    CheckLookahead book
    book.Close False
    AtlasAcceptance = CStr(checks) & " Excel acceptance checks passed"
    Exit Function
Failed:
    AtlasAcceptance = "FAIL after " & checks & " checks: " & Err.Description
End Function

Private Sub CheckLookahead(ByVal book As Workbook)
    Dim w As Worksheet, s As Worksheet, c As Worksheet, v As Worksheet, k As Long, cats As Variant
    Set w = ProjectSheet(book): Set s = book.Worksheets("SAN_SANG")
    Set c = book.Worksheets("RANG_BUOC"): Set v = book.Worksheets("LOOKAHEAD_21N")
    w.Cells(9, 14) = "BCH": w.Cells(9, 15) = "Zone A"
    RefreshLookahead book
    Check v.Range("I6").Value = "HOLD", "Missing readiness must HOLD"
    cats = Array("BAN_VE", "VAT_TU", "MAT_BANG", "QA_QC", "HSE", "NHAN_LUC")
    For k = 0 To 5
        s.Cells(6 + k, 1) = "A": s.Cells(6 + k, 2) = cats(k)
        s.Cells(6 + k, 3) = "BCH": s.Cells(6 + k, 4) = w.Range("B5").Value
        s.Cells(6 + k, 5) = "BB-01": s.Cells(6 + k, 6) = "PASS"
    Next k
    RefreshLookahead book
    Check v.Range("I6").Value = "READY", "Six confirmed conditions READY"
    c.Range("A6") = "C01": c.Range("B6") = "A": c.Range("C6") = "Vat tu"
    c.Range("E6") = CDate(w.Range("B5").Value) - 1: c.Range("F6") = "CLOSED"
    RefreshLookahead book
    Check v.Range("I6").Value = "HOLD" And v.Range("L6").Value = 1, "Unproven CLOSED remains overdue HOLD"
    c.Range("G6") = "BB-02": c.Range("H6") = "BCH": c.Range("I6") = w.Range("B5").Value
    RefreshLookahead book
    Check v.Range("I6").Value = "READY", "Verified constraint closure READY"
    s.Range("D6") = CDate(w.Range("B5").Value) + 1
    RefreshLookahead book
    Check v.Range("I6").Value = "HOLD", "Future confirmation rejected"
    s.Range("D6") = w.Range("B5").Value
    s.Range("A12:F12").Value = s.Range("A6:F6").Value
    RefreshLookahead book
    Check v.Range("I6").Value = "HOLD", "Duplicate confirmation rejected"
    s.Range("A12:F12").ClearContents
    c.Range("B6") = "ORPHAN"
    RefreshLookahead book
    Check v.Range("I6").Value = "HOLD", "Orphan ID blocks release"
    c.Range("B6") = "A"
    w.Range("B5") = CDate(w.Cells(9, 7).Value) + 10
    RefreshLookahead book
    Check v.Range("A6").Value = "A", "Unfinished overdue activity retained"
    w.Cells(9, 10) = 100
    RefreshLookahead book
    Check v.Range("A6").Value <> "A", "Completed activity excluded"
    s.Range("A6:F12005").ClearContents
    PrepareReadiness book
    k = s.Cells(s.Rows.Count, 1).End(xlUp).Row
    PrepareReadiness book
    Check s.Cells(s.Rows.Count, 1).End(xlUp).Row = k, "Prepare confirmations preserves rows without duplicates"
    w.Cells(9, 10) = 0: w.Cells(10, 10) = 100
    w.Range("B5") = DateSerial(2026, 10, 5)
    w.Cells(9, 6) = CDate(w.Range("B5").Value) + 20: w.Cells(9, 7) = w.Cells(9, 6).Value
    RefreshLookahead book
    Check v.Range("A6").Value = "A", "Day 21 included"
    w.Cells(9, 6) = CDate(w.Range("B5").Value) + 21: w.Cells(9, 7) = w.Cells(9, 6).Value
    RefreshLookahead book
    Check v.Range("A6").Value = "", "Day 22 excluded and old rows cleared"
End Sub
