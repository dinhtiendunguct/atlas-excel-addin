Attribute VB_Name = "AtlasReports"
Option Explicit

Public Sub BuildResourceSheet(ByVal book As Workbook)
    Dim w As Worksheet, rows As Variant, row As Variant, i As Long, j As Long
    Set w = book.Worksheets("NGUON_LUC")
    w.Range("A1:F1").Merge: w.Range("A1") = "DANH MỤC NGUỒN LỰC | NHÂN LỰC & THIẾT BỊ"
    TitleStyle w.Range("A1:F1")
    w.Range("A3:F3").Merge: w.Range("A3") = "Đơn giá minh họa — sửa theo dự án. Gán tại TIEN_DO: NC:12;KS:1. Số lượng cố định / ngày làm việc."
    PutHeaders w, 5, Array("Mã", "Tên nguồn lực", "Nhóm", "Đơn vị", "Đơn giá / ngày", "Ghi chú")
    rows = Array(Array("NC", "Công nhân xây dựng", "NL", "người", 450000, "Demo"), Array("KS", "Kỹ sư / cán bộ kỹ thuật", "NL", "người", 800000, "Demo"), Array("CT", "Cẩu tháp", "MAY", "máy", 4500000, "Demo"), Array("VT", "Vận thăng", "MAY", "máy", 2000000, "Demo"), Array("MX", "Máy xúc", "MAY", "máy", 3500000, "Demo"), Array("OT", "Ô tô vận chuyển", "MAY", "xe", 2500000, "Demo"), Array("KN", "Máy khoan cọc nhồi", "MAY", "máy", 12000000, "Demo"), Array("BT", "Bơm bê tông", "MAY", "máy", 4000000, "Demo"))
    For i = 0 To UBound(rows)
        row = rows(i)
        For j = 0 To UBound(row): w.Cells(i + 6, j + 1) = row(j): Next j
    Next i
    w.Columns("A").ColumnWidth = 12: w.Columns("B").ColumnWidth = 32
    w.Columns("C:D").ColumnWidth = 16: w.Columns("E").ColumnWidth = 22: w.Columns("F").ColumnWidth = 35
    w.Range("A6:F100").Interior.Color = RGB(255, 249, 216): w.Range("E6:E100").NumberFormat = "#,##0"
End Sub
Public Function ResourceCatalog(ByVal book As Workbook) As Object
    Dim w As Worksheet, r As Long, id As String, group As String, value As Variant, catalog As Object
    Set w = book.Worksheets("NGUON_LUC"): Set catalog = CreateObject("Scripting.Dictionary"): catalog.CompareMode = vbTextCompare
    For r = 6 To w.Cells(w.Rows.Count, 1).End(xlUp).Row
        id = Trim(CStr(w.Cells(r, 1)))
        If id <> "" Then
            If catalog.Exists(id) Then Err.Raise 5, , "Mã nguồn lực trùng: " & id
            group = UCase(Trim(CStr(w.Cells(r, 3))))
            If group <> "NL" And group <> "MAY" Then Err.Raise 5, , "Nhóm nguồn lực chỉ nhận NL/MAY: " & id
            value = w.Cells(r, 5).Value
            If Not IsNumeric(value) Or IsEmpty(value) Then Err.Raise 5, , "Đơn giá không hợp lệ: " & id
            If CDbl(value) < 0 Then Err.Raise 5, , "Đơn giá phải >= 0: " & id
            catalog.Add id, Array(CStr(w.Cells(r, 2)), group, CDbl(value))
        End If
    Next r
    Set ResourceCatalog = catalog
End Function
Public Sub ValidateResources(ByVal book As Workbook)
    Dim w As Worksheet, catalog As Object, r As Long, parts As Variant, token As Variant, seen As Object
    Set w = ProjectSheet(book): Set catalog = ResourceCatalog(book)
    For r = FIRST_ROW To LastTask(w)
        If Len(Trim(CStr(w.Cells(r, 9)))) > 0 Then
            If w.Cells(r, 4) <> "CV" Then Err.Raise 5, , "Nguồn lực chỉ gán vào CV, dòng " & r
            Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
            For Each token In Split(CStr(w.Cells(r, 9)), ";")
                parts = Split(Trim(CStr(token)), ":")
                If UBound(parts) <> 1 Then Err.Raise 5, , "Nguồn lực dùng NC:12;KS:1 tại dòng " & r
                parts(0) = Trim(parts(0))
                If Not catalog.Exists(parts(0)) Then Err.Raise 5, , "Không tìm thấy nguồn lực " & parts(0) & " tại dòng " & r
                If seen.Exists(parts(0)) Then Err.Raise 5, , "Nguồn lực gán trùng tại dòng " & r
                seen.Add parts(0), True
                If Not IsNumeric(parts(1)) Then Err.Raise 5, , "Số lượng nguồn lực không hợp lệ tại dòng " & r
                If CDbl(parts(1)) < 0 Or CDbl(parts(1)) > 100000 Then Err.Raise 5, , "Số lượng nguồn lực ngoài phạm vi tại dòng " & r
            Next token
        End If
    Next r
End Sub
Public Sub RibbonReports(ByVal control As Object)
    On Error GoTo Failed
    CalculateProject ActiveWorkbook, True
    BuildReports ActiveWorkbook, False
    ShowSheet "BAO_CAO"
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RibbonResourceCharts(ByVal control As Object)
    On Error GoTo Failed
    CalculateProject ActiveWorkbook, False
    BuildReports ActiveWorkbook, False
    ShowSheet "BIEU_DO_NL"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RibbonSnapshot(ByVal control As Object)
    On Error GoTo Failed
    CalculateProject ActiveWorkbook, True
    SaveSnapshot ActiveWorkbook
    BuildReports ActiveWorkbook, False
    MsgBox "Đã lưu kỳ báo cáo. Xem lịch sử ở NHAT_KY và đường cong S ở BAO_CAO.", vbInformation, "Atlas"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub SaveSnapshot(ByVal book As Workbook)
    Dim w As Worksheet, log As Worksheet, last As Long, r As Long, reportDate As Date
    Set w = ProjectSheet(book): Set log = book.Worksheets("NHAT_KY")
    reportDate = DateValue(w.Range("B5").Value)
    last = log.Cells(log.Rows.Count, 1).End(xlUp).Row
    For r = 5 To last
        If IsDate(log.Cells(r, 1)) Then
            If DateValue(log.Cells(r, 1)) = reportDate Then Err.Raise 5, , "Ngày báo cáo đã được ghi nhận. Giữ lịch sử cũ; chọn ngày báo cáo mới."
        End If
    Next r
    r = Application.Max(5, last + 1)
    log.Cells(r, 1) = reportDate: log.Cells(r, 2) = w.Range("J5").Value
    log.Cells(r, 3) = Now: log.Cells(r, 4) = "Atlas " & ATLAS_VERSION
    log.Cells(r, 5) = "Trọng số thời lượng CV; không tính HM / mốc"
    log.Cells(r, 1).NumberFormat = "dd/mm/yyyy": log.Cells(r, 3).NumberFormat = "dd/mm/yyyy hh:mm"
    log.Cells(r, 2).NumberFormat = "0.0%"
End Sub
Private Function PlannedAt(ByVal w As Worksheet, ByVal last As Long, ByVal asOf As Date, ByVal useBaseline As Boolean) As Double
    Dim r As Long, a As Date, b As Date, duration As Double, total As Double, earned As Double, elapsed As Long
    For r = FIRST_ROW To last
        If w.Cells(r, 4) = "CV" Then
            If useBaseline Then a = CDate(w.Cells(r, 16)): b = CDate(w.Cells(r, 17)) Else a = CDate(w.Cells(r, 6)): b = CDate(w.Cells(r, 7))
            duration = WorkCount(a, b): total = total + duration
            If asOf >= a Then
                elapsed = WorkCount(a, Application.Min(CDbl(asOf), CDbl(b)))
                earned = earned + Application.Min(elapsed, duration)
            End If
        End If
    Next r
    If total > 0 Then PlannedAt = earned / total
End Function
Public Sub BuildReports(ByVal book As Workbook, Optional ByVal unused As Boolean = False)
    Dim w As Worksheet, report As Worksheet, resources As Worksheet, log As Worksheet, catalog As Object
    Dim last As Long, r As Long, m As Long, j As Long, k As Long, dateCol As Long, resourceCount As Long
    Dim first As Date, lastDate As Date, monthCursor As Date, monthEnd As Date, a As Date, b As Date, d As Date, reportDate As Date
    Dim base As Boolean, weight As Double, current As Double, previous As Double, planned As Double, cost As Double, running As Double
    Dim keys As Variant, key As Variant, entry As Variant, parts As Variant, token As Variant, qty As Double, days As Long, monthDays As Long
    Dim nl() As Double, machines() As Double, sums As Object, baselineCount As Long, taskCount As Long, actualRow As Long
    Set w = ProjectSheet(book): last = LastTask(w): LoadCalendar book: ValidateResources book
    Set catalog = ResourceCatalog(book): keys = catalog.Keys: resourceCount = catalog.Count
    Set report = book.Worksheets("BAO_CAO"): Set resources = book.Worksheets("BIEU_DO_NL"): Set log = book.Worksheets("NHAT_KY")
    For r = FIRST_ROW To last
        If w.Cells(r, 4) = "CV" Then
            taskCount = taskCount + 1
            If IsDate(w.Cells(r, 16)) And IsDate(w.Cells(r, 17)) Then baselineCount = baselineCount + 1
        End If
    Next r
    base = taskCount > 0 And baselineCount = taskCount
    first = CDate(Application.Min(w.Range("F9:F" & last))): lastDate = CDate(Application.Max(w.Range("G9:G" & last)))
    If base Then
        first = Application.Min(CDbl(first), Application.Min(w.Range("P9:P" & last)))
        lastDate = Application.Max(CDbl(lastDate), Application.Max(w.Range("Q9:Q" & last)))
    End If
    reportDate = DateValue(w.Range("B5"))
    first = DateSerial(Year(first), Month(first), 1)
    ClearReport report: ClearReport resources
    report.Range("A1:N1").Merge: report.Range("A1") = "BÁO CÁO TIẾN ĐỘ & ĐƯỜNG CONG S"
    TitleStyle report.Range("A1:N1")
    report.Range("A2:N2").Merge: report.Range("A2") = w.Range("B2").Value
    report.Range("A4") = "Ngày bắt đầu": report.Range("D4") = w.Range("B4").Value
    report.Range("A5") = "Kết thúc dự kiến": report.Range("D5") = w.Range("F4").Value
    report.Range("A6") = "Ngày báo cáo": report.Range("D6") = reportDate
    report.Range("A7") = "% thực tế hiện tại": report.Range("D7") = w.Range("J5").Value
    report.Range("A8") = "% kế hoạch tại ngày BC": planned = PlannedAt(w, last, reportDate, base): report.Range("D8") = planned
    report.Range("A9") = "Chỉ số tiến độ TT/KH"
    If planned > 0 Then report.Range("D9") = CDbl(w.Range("J5")) / planned Else report.Range("D9") = "—"
    report.Range("A10") = "Số công việc / mốc găng": report.Range("D10") = w.Range("N4").Value
    report.Range("A11") = "Cơ sở kế hoạch": report.Range("D11") = IIf(base, "Baseline", "Lịch hiện tại")
    report.Range("A12:F12").Merge: report.Range("A12") = "Trọng số thời lượng; phân bổ theo ngày LV. Thực tế chỉ dùng các kỳ đã ghi nhận."
    report.Range("A13:F13").Merge: report.Range("A13") = "Đây là chỉ số tiến độ theo thời lượng, không phải SPI theo giá trị EVM."
    report.Range("A12:F13").Font.Size = 9: report.Range("A12:F13").Font.Color = RGB(93, 111, 130)
    report.Range("D4:D6").NumberFormat = "dd/mm/yyyy": report.Range("D7:D8").NumberFormat = "0.0%": report.Range("D9").NumberFormat = "0.00"
    report.Range("A4:D11").Interior.Color = RGB(239, 245, 250)
    PutHeaders report, 15, Array("STT", "Tháng", "% KH trong tháng", "% KH lũy kế", "", "")
    report.Range("E15") = "Ngày thực tế": report.Range("F15") = "% TT ghi nhận"
    resources.Range("A1:K1").Merge: resources.Range("A1") = "BÁO CÁO & BIỂU ĐỒ NGUỒN LỰC KẾ HOẠCH"
    TitleStyle resources.Range("A1:K1")
    resources.Range("A2:K2").Merge: resources.Range("A2") = "Giả định phân bổ đều, quân số / máy bình quân trên ngày làm việc của tháng. Đơn giá theo NGUON_LUC."
    resources.Cells(5, 1) = "Tháng"
    For k = 0 To resourceCount - 1
        entry = catalog(keys(k)): resources.Cells(5, k + 2) = keys(k)
    Next k
    dateCol = resourceCount + 2
    resources.Cells(5, dateCol) = "Chi phí tháng": resources.Cells(5, dateCol + 1) = "Chi phí lũy kế"
    monthCursor = first: r = 16
    Do While monthCursor <= lastDate
        monthEnd = DateSerial(Year(monthCursor), Month(monthCursor) + 1, 0)
        current = PlannedAt(w, last, monthEnd, base)
        report.Cells(r, 1) = r - 15: report.Cells(r, 2) = monthEnd
        report.Cells(r, 3) = current - previous: report.Cells(r, 4) = current: previous = current
        resources.Cells(r - 10, 1) = monthCursor: monthDays = WorkCount(monthCursor, monthEnd)
        Set sums = CreateObject("Scripting.Dictionary"): sums.CompareMode = vbTextCompare
        For Each key In keys: sums(key) = 0#: Next key
        cost = 0
        For j = FIRST_ROW To last
            If w.Cells(j, 4) = "CV" And Len(Trim(CStr(w.Cells(j, 9)))) > 0 Then
                a = Application.Max(CDbl(monthCursor), CDbl(w.Cells(j, 6)))
                b = Application.Min(CDbl(monthEnd), CDbl(w.Cells(j, 7)))
                days = WorkCount(a, b)
                If days > 0 Then
                    For Each token In Split(CStr(w.Cells(j, 9)), ";")
                        parts = Split(Trim(CStr(token)), ":"): key = Trim(parts(0)): qty = CDbl(parts(1))
                        entry = catalog(key): sums(key) = sums(key) + qty * days
                        cost = cost + qty * days * CDbl(entry(2))
                    Next token
                End If
            End If
        Next j
        For k = 0 To resourceCount - 1
            If monthDays > 0 Then resources.Cells(r - 10, k + 2) = sums(keys(k)) / monthDays Else resources.Cells(r - 10, k + 2) = 0
        Next k
        running = running + cost: resources.Cells(r - 10, dateCol) = cost: resources.Cells(r - 10, dateCol + 1) = running
        monthCursor = DateSerial(Year(monthCursor), Month(monthCursor) + 1, 1): r = r + 1
    Loop
    report.Range("B16:B" & r - 1).NumberFormat = "mm/yyyy": report.Range("C16:D" & r - 1).NumberFormat = "0.0%"
    resources.Range(resources.Cells(6, 1), resources.Cells(r - 11, 1)).NumberFormat = "mm/yyyy"
    resources.Range(resources.Cells(6, 2), resources.Cells(r - 11, dateCol - 1)).NumberFormat = "0.0"
    resources.Range(resources.Cells(6, dateCol), resources.Cells(r - 11, dateCol + 1)).NumberFormat = "#,##0"
    actualRow = 16
    For j = 5 To log.Cells(log.Rows.Count, 1).End(xlUp).Row
        If IsDate(log.Cells(j, 1)) Then
            If CDate(log.Cells(j, 1)) <= reportDate Then
                report.Cells(actualRow, 5) = log.Cells(j, 1).Value: report.Cells(actualRow, 6) = log.Cells(j, 2).Value
                actualRow = actualRow + 1
            End If
        End If
    Next j
    If actualRow > 16 Then report.Range("E15:F" & actualRow - 1).Sort Key1:=report.Range("E16"), Order1:=xlAscending, Header:=xlYes
    report.Range("E16:E" & Application.Max(16, actualRow - 1)).NumberFormat = "dd/mm/yyyy": report.Range("F16:F" & Application.Max(16, actualRow - 1)).NumberFormat = "0.0%"
    report.Columns("A").ColumnWidth = 27: report.Columns("B:F").ColumnWidth = 18
    resources.Columns("A:Q").ColumnWidth = 14
    MakeSCurve report, r - 1, actualRow - 1
    MakeResourceCharts resources, catalog, keys, r - 11, dateCol
    report.Rows(1).RowHeight = 32: resources.Rows(1).RowHeight = 32
    report.PageSetup.Orientation = xlLandscape: report.PageSetup.Zoom = False: report.PageSetup.FitToPagesWide = 1
    report.PageSetup.PrintArea = "$A$1:$V$" & Application.Max(32, r)
    resources.PageSetup.Orientation = xlLandscape: resources.PageSetup.Zoom = False: resources.PageSetup.FitToPagesWide = 1
    resources.PageSetup.PrintArea = "$A$1:$Q$" & Application.Max(65, r + 36)
End Sub
Private Sub ClearReport(ByVal w As Worksheet)
    Dim i As Long
    For i = w.ChartObjects.Count To 1 Step -1: w.ChartObjects(i).Delete: Next i
    w.Cells.UnMerge: w.Cells.Clear
End Sub
Private Sub MakeSCurve(ByVal w As Worksheet, ByVal last As Long, ByVal lastActual As Long)
    Dim obj As ChartObject, series As series
    Set obj = w.ChartObjects.Add(w.Range("H4").Left, w.Range("H4").Top, 660, 365)
    With obj.Chart
        .ChartType = xlXYScatterLines
        Set series = .SeriesCollection.NewSeries
        series.Name = "Kế hoạch lũy kế": series.XValues = w.Range("B16:B" & last): series.Values = w.Range("D16:D" & last)
        series.Format.Line.ForeColor.RGB = RGB(46, 106, 194): series.MarkerStyle = xlMarkerStyleDiamond
        If lastActual >= 16 Then
            Set series = .SeriesCollection.NewSeries
            series.Name = "Thực tế đã ghi nhận": series.XValues = w.Range("E16:E" & lastActual): series.Values = w.Range("F16:F" & lastActual)
            series.Format.Line.ForeColor.RGB = RGB(223, 106, 58): series.MarkerStyle = xlMarkerStyleSquare
        End If
        .HasTitle = True: .ChartTitle.Text = "ĐƯỜNG CONG S TIẾN ĐỘ (TRỌNG SỐ THỜI LƯỢNG)"
        .Axes(xlValue).MinimumScale = 0: .Axes(xlValue).MaximumScale = 1: .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .Axes(xlCategory).TickLabels.NumberFormat = "mm/yyyy"
        .Legend.Position = xlLegendPositionBottom
    End With
End Sub
Private Sub MakeResourceCharts(ByVal w As Worksheet, ByVal catalog As Object, ByVal keys As Variant, ByVal last As Long, ByVal costCol As Long)
    Dim obj As ChartObject, series As series, k As Long, entry As Variant, group As Variant, left As Double, i As Long, topRow As Long
    topRow = Application.Max(22, last + 3)
    For Each group In Array("NL", "MAY")
        left = w.Range("A22").Left + IIf(group = "NL", 0, 510)
        Set obj = w.ChartObjects.Add(left, w.Rows(topRow).Top, 490, 265)
        With obj.Chart
            .ChartType = xlColumnStacked
            For k = 0 To catalog.Count - 1
                entry = catalog(keys(k))
                If entry(1) = group Then
                    Set series = .SeriesCollection.NewSeries
                    series.Name = entry(0): series.XValues = w.Range("A6:A" & last): series.Values = w.Range(w.Cells(6, k + 2), w.Cells(last, k + 2))
                End If
            Next k
            .HasTitle = True: .ChartTitle.Text = IIf(group = "NL", "NHÂN LỰC BÌNH QUÂN / THÁNG", "THIẾT BỊ BÌNH QUÂN / THÁNG")
            If .SeriesCollection.Count > 0 Then
                .Legend.Position = xlLegendPositionBottom: .Axes(xlCategory).TickLabels.NumberFormat = "mm/yyyy"
            End If
        End With
    Next group
    Set obj = w.ChartObjects.Add(w.Range("A43").Left, w.Rows(topRow + 21).Top, 1030, 290)
    With obj.Chart
        .ChartType = xlColumnClustered
        Set series = .SeriesCollection.NewSeries
        series.Name = "Chi phí kế hoạch / tháng": series.XValues = w.Range("A6:A" & last): series.Values = w.Range(w.Cells(6, costCol), w.Cells(last, costCol))
        series.Format.Fill.ForeColor.RGB = RGB(47, 111, 209)
        Set series = .SeriesCollection.NewSeries
        series.Name = "Lũy kế chi phí kế hoạch": series.XValues = w.Range("A6:A" & last): series.Values = w.Range(w.Cells(6, costCol + 1), w.Cells(last, costCol + 1))
        series.ChartType = xlLineMarkers: series.AxisGroup = xlSecondary: series.Format.Line.ForeColor.RGB = RGB(225, 99, 64)
        .HasTitle = True: .ChartTitle.Text = "CHI PHÍ NGUỒN LỰC THEO THÁNG & LŨY KẾ"
        .Legend.Position = xlLegendPositionBottom: .Axes(xlCategory).TickLabels.NumberFormat = "mm/yyyy"
    End With
End Sub
