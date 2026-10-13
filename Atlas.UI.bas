Attribute VB_Name = "AtlasUI"
Option Explicit

Public Sub CreateProject(Optional ByVal demo As Boolean = False, Optional ByVal testing As Boolean = False)
    On Error GoTo Failed
    Dim book As Workbook, w As Worksheet, cfg As Worksheet, names As Variant, name As Variant, i As Long
    Application.ScreenUpdating = False
    Set book = Workbooks.Add(xlWBATWorksheet)
    names = Array("TIEN_DO", "NGUON_LUC", "CAU_HINH", "NHAT_KY", "HUONG_DAN", "BIEU_DO_NL", "BAO_CAO")
    book.Worksheets(1).Name = names(0)
    For i = 1 To UBound(names)
        Set w = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count)): w.Name = names(i)
    Next i
    For Each name In names
        Set w = book.Worksheets(CStr(name))
        w.Cells.Font.Name = "Calibri": w.Cells.Font.Size = 10
        w.Tab.Color = RGB(24, 54, 86)
    Next name
    Set cfg = book.Worksheets("CAU_HINH")
    cfg.Range("A1:F1").Merge: cfg.Range("A1") = "ATLAS | CẤU HÌNH & NGÀY NGHỈ"
    cfg.Range("A3") = "Tên dự án": cfg.Range("B3") = "DỰ ÁN THI CÔNG MỚI"
    cfg.Range("A4") = "Gói thầu": cfg.Range("B4") = "Thi công xây dựng"
    cfg.Range("A6") = "Lịch nghỉ (T2 → CN)": cfg.Range("B6").NumberFormat = "@": cfg.Range("B6") = "0000001"
    cfg.Range("A7") = "0 = làm việc; 1 = nghỉ. Ví dụ 0000011: nghỉ thứ Bảy và Chủ nhật."
    cfg.Range("A8") = "Hiện đường liên kết": cfg.Range("B8") = "KHONG"
    cfg.Range("A9") = "Ngày nghỉ lễ bên dưới là ngày nghỉ bổ sung, không tính vào thời lượng công việc."
    cfg.Range("A11") = "Ngày nghỉ": cfg.Range("B11") = "Ghi chú"
    cfg.Range("A12:A500").NumberFormat = "dd/mm/yyyy"
    cfg.Columns("A").ColumnWidth = 29: cfg.Columns("B").ColumnWidth = 55
    cfg.Range("B3:B6").Interior.Color = RGB(255, 249, 216): cfg.Range("A12:B50").Interior.Color = RGB(255, 249, 216)
    TitleStyle cfg.Range("A1:F1")
    BuildResourceSheet book
    Set w = book.Worksheets("NHAT_KY")
    w.Range("A1:F1").Merge: w.Range("A1") = "NHẬT KÝ CẬP NHẬT TIẾN ĐỘ"
    PutHeaders w, 4, Array("Ngày báo cáo", "% thực tế", "Ghi nhận lúc", "Nguồn", "Ghi chú")
    w.Columns("A:E").ColumnWidth = 22: w.Range("B5:B2000").NumberFormat = "0.0%"
    TitleStyle w.Range("A1:F1")
    BuildGuide book
    Set w = book.Worksheets("TIEN_DO")
    w.Range("A1:O1").Merge: w.Range("A1") = "ATLAS | LẬP & KIỂM SOÁT TIẾN ĐỘ DỰ ÁN"
    TitleStyle w.Range("A1:O1"): w.Rows(1).RowHeight = 32
    w.Range("A2") = "Dự án": w.Range("B2:H2").Merge: w.Range("B2") = cfg.Range("B3").Value
    w.Range("A3") = "Gói thầu": w.Range("B3:H3").Merge: w.Range("B3") = cfg.Range("B4").Value
    w.Range("A4") = "Bắt đầu": w.Range("B4") = Date
    w.Range("A5") = "Báo cáo": w.Range("B5") = Date
    w.Range("A6") = "Thang Gantt": w.Range("B6") = "TUAN"
    w.Range("E4") = "Kết thúc": w.Range("I4") = "Ngày làm việc": w.Range("M4") = "Việc găng"
    w.Range("E5") = "Chế độ": w.Range("F5") = "CPM / lịch dự án": w.Range("I5") = "% thực tế": w.Range("M5") = "Phiên bản"
    w.Range("D6:O6").Merge: w.Range("D6") = "■ Kế hoạch    ■ Găng (đỏ)    ■ Hoàn thành (xanh lá)    ━ Baseline (xám)    ◆ Mốc    │ Ngày báo cáo"
    w.Range("A7:O7").Merge: w.Range("A7") = "Ô vàng: nhập liệu · HM: hạng mục · CV: công việc · MOC: mốc · Quan hệ A:FS:0;B:SS:2 · Nguồn lực NC:12;KS:1"
    w.Range("A7:O7").Font.Color = RGB(90, 105, 120): w.Range("A7:O7").Font.Size = 9
    w.Range("B4:B5").NumberFormat = "dd/mm/yyyy": w.Range("F4").NumberFormat = "dd/mm/yyyy": w.Range("J5").NumberFormat = "0.0%"
    w.Range("B2:H3").Interior.Color = RGB(255, 249, 216): w.Range("B4:B6").Interior.Color = RGB(255, 249, 216)
    PutHeaders w, 8, Array("ID", "WBS", "Hạng mục / công việc", "Loại", "Ngày LV", "Bắt đầu", "Kết thúc", "Công việc trước", "Nguồn lực", "% HT", "Mũi thi công / ghi chú", "Dự trữ", "Lệch BL", "Phụ trách", "Khu vực", "BL bắt đầu", "BL kết thúc")
    w.Columns("A").ColumnWidth = 10: w.Columns("B").ColumnWidth = 9: w.Columns("C").ColumnWidth = 34
    w.Columns("D:E").ColumnWidth = 8: w.Columns("F:G").ColumnWidth = 12
    w.Columns("H:I").ColumnWidth = 20: w.Columns("J").ColumnWidth = 8: w.Columns("K").ColumnWidth = 26
    w.Columns("L:M").ColumnWidth = 9: w.Columns("N:O").ColumnWidth = 15
    w.Columns("P:Q").Hidden = True
    w.Rows(8).RowHeight = 32
    w.Range("F9:G2008").NumberFormat = "dd/mm/yyyy"
    w.Range("J9:J2008").NumberFormat = "0.0" & Chr(34) & "%" & Chr(34)
    With w.Range("B6").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "NGAY,TUAN,THANG"
    End With
    With w.Range("D9:D2008").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "HM,CV,MOC"
    End With
    With w.Range("J9:J2008").Validation
        .Delete: .Add xlValidateDecimal, xlValidAlertStop, xlBetween, "0", "100"
    End With
    If demo Then
        FillDemo w
        w.Range("B2") = "CÔNG TRÌNH NHÀ LÀM VIỆC 5 TẦNG + 1 HẦM — DEMO"
        w.Range("B3") = "Gói thầu 01 · Thi công xây dựng và lắp đặt thiết bị"
        w.Range("B4") = DateSerial(2026, 5, 4): w.Range("B5") = DateSerial(2026, 9, 8)
        cfg.Range("B3") = w.Range("B2").Value: cfg.Range("B4") = w.Range("B3").Value
    Else
        AddDemoRow w, 9, "HM01", "1", "Hạng mục 1", "HM", 0, "", "", 0
        AddDemoRow w, 10, "CV01", "1.1", "Công việc đầu tiên", "CV", 5, "", "", 0
    End If
    CalculateProject book, True
    If demo Then
        w.Range("P9:Q35").Value = w.Range("F9:G35").Value
        w.Range("P9:Q35").NumberFormat = "dd/mm/yyyy"
        w.Cells(18, 5) = 22: w.Cells(18, 11) = "Demo: tăng 2 ngày so với baseline"
        CalculateProject book, True
        SaveSnapshot book
    End If
    BuildReports book, False
    book.Worksheets("NGUON_LUC").Tab.Color = RGB(137, 83, 211)
    book.Worksheets("CAU_HINH").Tab.Color = RGB(111, 129, 149)
    book.Worksheets("NHAT_KY").Tab.Color = RGB(91, 112, 127)
    book.Worksheets("BIEU_DO_NL").Tab.Color = RGB(137, 83, 211)
    book.Worksheets("BAO_CAO").Tab.Color = RGB(22, 163, 107)
    w.Activate: ActiveWindow.DisplayGridlines = False: ActiveWindow.Zoom = 80
    If Application.Visible Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.SplitRow = 8: ActiveWindow.SplitColumn = 3: ActiveWindow.FreezePanes = True
    End If
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    If testing Then Err.Raise Err.Number, "CreateProject", Err.Description
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub

Public Sub TitleStyle(ByVal area As Range)
    area.Interior.Color = RGB(23, 49, 78): area.Font.Color = vbWhite
    area.Font.Bold = True: area.Font.Size = 15
End Sub
Public Sub PutHeaders(ByVal w As Worksheet, ByVal r As Long, ByVal items As Variant)
    Dim i As Long
    For i = 0 To UBound(items): w.Cells(r, i + 1) = items(i): Next i
    With w.Range(w.Cells(r, 1), w.Cells(r, UBound(items) + 1))
        .Interior.Color = RGB(23, 49, 78): .Font.Color = vbWhite: .Font.Bold = True
        .WrapText = True: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter
    End With
End Sub
Private Sub AddDemoRow(ByVal w As Worksheet, ByVal r As Long, ByVal id As String, ByVal code As String, ByVal title As String, ByVal kind As String, ByVal days As Long, ByVal predecessors As String, ByVal resources As String, ByVal progress As Double)
    w.Cells(r, 1) = id: w.Cells(r, 2).NumberFormat = "@": w.Cells(r, 2) = code
    w.Cells(r, 3) = title: w.Cells(r, 4) = kind: w.Cells(r, 5) = days
    w.Cells(r, 8) = predecessors: w.Cells(r, 9) = resources: w.Cells(r, 10) = progress
End Sub
Private Sub FillDemo(ByVal w As Worksheet)
    AddDemoRow w, 9, "H1", "1", "CÔNG TÁC CHUẨN BỊ", "HM", 0, "", "", 0
    AddDemoRow w, 10, "A01", "1.1", "Nhận mặt bằng, định vị công trình", "CV", 5, "", "NC:6;KS:1", 100
    AddDemoRow w, 11, "A02", "1.2", "Lán trại, điện nước thi công", "CV", 10, "A01:SS:2", "NC:10;KS:1", 100
    AddDemoRow w, 12, "M01", "1.3", "Mốc: Khởi công", "MOC", 0, "A01:FS:0;A02:FS:0", "", 100
    AddDemoRow w, 13, "H2", "2", "PHẦN NGẦM", "HM", 0, "", "", 0
    AddDemoRow w, 14, "A03", "2.1", "Thi công cọc khoan nhồi D800", "CV", 30, "M01:FS:0", "NC:18;KS:2;KN:1", 100
    AddDemoRow w, 15, "A04", "2.2", "Tường vây, chống hầm và đào đất", "CV", 20, "A03:SS:7", "NC:20;KS:2;MX:1", 100
    AddDemoRow w, 16, "A05", "2.3", "Đào đài móng, vận chuyển đất", "CV", 24, "A03:FS:0;A04:FS:0", "NC:16;MX:2;OT:2", 100
    AddDemoRow w, 17, "A06", "2.4", "Đài móng, giằng móng", "CV", 18, "A05:FS:0", "NC:24;KS:2;BT:1", 100
    AddDemoRow w, 18, "A07", "2.5", "Sàn hầm, tường hầm, chống thấm", "CV", 20, "A06:FS:0", "NC:24;KS:2;BT:1", 80
    AddDemoRow w, 19, "M02", "2.6", "Mốc: Hoàn thành phần ngầm", "MOC", 0, "A07:FS:0", "", 0
    AddDemoRow w, 20, "H3", "3", "PHẦN THÂN", "HM", 0, "", "", 0
    AddDemoRow w, 21, "A08", "3.1", "Khung bê tông cốt thép tầng 1", "CV", 14, "M02:FS:0", "NC:30;KS:2;CT:1", 0
    AddDemoRow w, 22, "A09", "3.2", "Khung bê tông cốt thép tầng 2", "CV", 12, "A08:FS:0", "NC:30;KS:2;CT:1", 0
    AddDemoRow w, 23, "A10", "3.3", "Khung bê tông cốt thép tầng 3", "CV", 12, "A09:FS:0", "NC:30;KS:2;CT:1", 0
    AddDemoRow w, 24, "A11", "3.4", "Khung bê tông cốt thép tầng 4", "CV", 12, "A10:FS:0", "NC:30;KS:2;CT:1", 0
    AddDemoRow w, 25, "A12", "3.5", "Khung bê tông cốt thép tầng 5", "CV", 12, "A11:FS:0", "NC:30;KS:2;CT:1", 0
    AddDemoRow w, 26, "A13", "3.6", "Mái, tum thang", "CV", 10, "A12:FS:0", "NC:22;KS:2;CT:1", 0
    AddDemoRow w, 27, "M03", "3.7", "Mốc: Cất nóc", "MOC", 0, "A13:FS:0", "", 0
    AddDemoRow w, 28, "H4", "4", "HOÀN THIỆN & MEP", "HM", 0, "", "", 0
    AddDemoRow w, 29, "A14", "4.1", "Xây tường, bao che", "CV", 45, "A09:SS:10", "NC:20;KS:1", 0
    AddDemoRow w, 30, "A15", "4.2", "Trát trong, trát ngoài", "CV", 40, "A14:SS:15", "NC:18;KS:1;VT:1", 0
    AddDemoRow w, 31, "A16", "4.3", "Ốp lát, cán nền", "CV", 30, "A15:SS:20", "NC:18;KS:1;VT:1", 0
    AddDemoRow w, 32, "A17", "4.4", "Trần thạch cao, sơn bả", "CV", 35, "A15:FF:5", "NC:16;KS:1", 0
    AddDemoRow w, 33, "A18", "4.5", "Lắp đặt hệ thống MEP", "CV", 45, "A14:SS:10", "NC:12;KS:2", 0
    AddDemoRow w, 34, "A19", "4.6", "Kiểm tra, chạy thử liên động", "CV", 10, "A16:FS:0;A17:FS:0;A18:FS:0;M03:FS:0", "NC:8;KS:3", 0
    AddDemoRow w, 35, "M04", "4.7", "Mốc: Nghiệm thu bàn giao", "MOC", 0, "A19:FS:0", "", 0
    w.Range("K14") = "Mũi cọc: 01 máy khoan + tổ cọc"
    w.Range("K21") = "Mũi kết cấu: 02 tổ cốp pha + thép"
    w.Range("K29") = "Mũi hoàn thiện: bố trí theo tầng"
End Sub
Public Sub StyleTaskRows(ByVal w As Worksheet)
    Dim r As Long, last As Long, depth As Long
    last = LastTask(w)
    w.Rows.ClearOutline
    w.Outline.SummaryRow = xlAbove
    With w.Range("A9:O" & last)
        .RowHeight = 25: .Font.Size = 10: .VerticalAlignment = xlCenter
        .Borders.Color = RGB(225, 232, 239): .Borders.Weight = xlHairline
        .Font.Bold = False
    End With
    For r = FIRST_ROW To last
        w.Range(w.Cells(r, 1), w.Cells(r, 15)).Interior.Color = IIf(r Mod 2 = 0, RGB(246, 249, 252), vbWhite)
        w.Range(w.Cells(r, 4), w.Cells(r, 5)).Interior.Color = RGB(255, 249, 216)
        w.Range(w.Cells(r, 8), w.Cells(r, 11)).Interior.Color = RGB(255, 249, 216)
        w.Range(w.Cells(r, 14), w.Cells(r, 15)).Interior.Color = RGB(255, 249, 216)
        depth = UBound(Split(CStr(w.Cells(r, 2)), "."))
        w.Cells(r, 3).IndentLevel = Application.Min(depth, 10)
        If w.Cells(r, 4) = "HM" Then
            w.Range(w.Cells(r, 1), w.Cells(r, 15)).Interior.Color = RGB(218, 230, 242)
            w.Range(w.Cells(r, 1), w.Cells(r, 15)).Font.Bold = True
        ElseIf w.Cells(r, 4) = "MOC" Then
            w.Cells(r, 3).Font.Italic = True: w.Cells(r, 3).Font.Color = RGB(63, 84, 166)
        End If
        w.Cells(r, 12).Font.Color = IIf(w.Cells(r, 4) <> "HM" And w.Cells(r, 12) = 0, RGB(200, 57, 63), RGB(67, 88, 110))
        w.Cells(r, 13).Font.Color = RGB(200, 57, 63)
    Next r
    GroupRows w, last
    w.Range("F9:G" & last).NumberFormat = "dd/mm/yyyy"
    w.Range("J9:J" & last).NumberFormat = "0.0" & Chr(34) & "%" & Chr(34)
    If w.AutoFilterMode Then w.AutoFilterMode = False
    w.Range("A8:O" & last).AutoFilter
End Sub
Private Sub GroupRows(ByVal w As Worksheet, ByVal last As Long)
    Dim r As Long, j As Long, prefix As String, endRow As Long
    For r = FIRST_ROW To last
        If w.Cells(r, 4) = "HM" Then
            prefix = CStr(w.Cells(r, 2)) & ".": endRow = r
            For j = r + 1 To last
                If Left(CStr(w.Cells(j, 2)), Len(prefix)) <> prefix Then Exit For
                endRow = j
            Next j
            If endRow > r And UBound(Split(prefix, ".")) < 8 Then w.Rows(CStr(r + 1) & ":" & endRow).Group
        End If
    Next r
End Sub

Public Sub DrawGantt(ByVal book As Workbook)
    Dim w As Worksheet, first As Date, lastDate As Date, a As Date, b As Date, cursor As Date, stopDate As Date, reportDate As Date
    Dim r As Long, c As Long, i As Long, last As Long, viewMode As String, cellsWide As Long, stepDays As Long
    Dim shape As Shape, x As Double, y As Double, width As Double, barColor As Long, rowStart As Date, rowEnd As Date, p As Double, d As Date
    Set w = ProjectSheet(book): last = LastTask(w): LoadCalendar book
    first = CDate(Application.Min(w.Range("F9:F" & last))): lastDate = CDate(Application.Max(w.Range("G9:G" & last)))
    If Application.Count(w.Range("P9:Q" & last)) > 0 Then
        first = Application.Min(CDbl(first), Application.Min(w.Range("P9:P" & last)))
        lastDate = Application.Max(CDbl(lastDate), Application.Max(w.Range("Q9:Q" & last)))
    End If
    viewMode = UCase(CStr(w.Range("B6"))): reportDate = CDate(w.Range("B5"))
    Select Case viewMode
        Case "NGAY": stepDays = 1
        Case "TUAN": stepDays = 7: first = first - Weekday(first, vbMonday) + 1
        Case "THANG": stepDays = 0: first = DateSerial(Year(first), Month(first), 1)
        Case Else: Err.Raise 5, , "Thang Gantt chỉ nhận NGAY, TUAN hoặc THANG."
    End Select
    If viewMode = "NGAY" And lastDate - first > 730 Then Err.Raise 5, , "Chế độ ngày tối đa 731 ngày. Chuyển sang TUAN hoặc THANG."
    For i = w.Shapes.Count To 1 Step -1
        If Left(w.Shapes(i).Name, 6) = "Atlas_" Then w.Shapes(i).Delete
    Next i
    w.Range(w.Cells(6, GANTT_COL), w.Cells(Application.Max(last, w.UsedRange.Rows.Count), Application.Max(GANTT_COL, w.UsedRange.Columns.Count))).UnMerge
    w.Range(w.Cells(6, GANTT_COL), w.Cells(Application.Max(last, w.UsedRange.Rows.Count), Application.Max(GANTT_COL, w.UsedRange.Columns.Count))).Clear
    c = GANTT_COL: cursor = first
    Do While cursor <= lastDate
        If stepDays = 0 Then stopDate = DateSerial(Year(cursor), Month(cursor) + 1, 1) Else stopDate = cursor + stepDays
        w.Cells(6, c) = Year(cursor): w.Cells(7, c) = "T" & Format(cursor, "mm/yyyy")
        If viewMode = "THANG" Then w.Cells(8, c) = Format(cursor, "mm/yyyy") Else w.Cells(8, c) = Format(cursor, "dd/mm")
        w.Columns(c).ColumnWidth = IIf(viewMode = "NGAY", 3.8, 9)
        With w.Range(w.Cells(6, c), w.Cells(8, c))
            .Interior.Color = RGB(23, 49, 78): .Font.Color = vbWhite: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .Font.Size = 9
        End With
        For r = FIRST_ROW To last
            w.Cells(r, c).Interior.Color = IIf(r Mod 2 = 0, RGB(246, 249, 252), vbWhite)
            If w.Cells(r, 4) = "HM" Then w.Cells(r, c).Interior.Color = RGB(232, 239, 247)
            w.Cells(r, c).Borders(xlEdgeRight).Color = RGB(224, 232, 240)
        Next r
        If viewMode = "NGAY" And Not IsWorking(cursor) Then w.Range(w.Cells(9, c), w.Cells(last, c)).Interior.Color = RGB(230, 233, 238)
        cursor = stopDate: c = c + 1
    Loop
    cellsWide = c - GANTT_COL
    For r = FIRST_ROW To last
        rowStart = CDate(w.Cells(r, 6)): rowEnd = CDate(w.Cells(r, 7))
        x = DateX(w, first, rowStart, viewMode): y = w.Cells(r, GANTT_COL).Top
        width = DateX(w, first, rowEnd + 1, viewMode) - x
        If w.Cells(r, 4) = "HM" Then
            Set shape = w.Shapes.AddShape(msoShapeRectangle, x, y + 7, Application.Max(width, 2), 4)
            shape.Fill.ForeColor.RGB = RGB(37, 57, 79): shape.Line.Visible = msoFalse
        ElseIf w.Cells(r, 4) = "MOC" Then
            Set shape = w.Shapes.AddShape(msoShapeDiamond, x - 4, y + 7, 9, 9)
            shape.Fill.ForeColor.RGB = RGB(90, 83, 196): shape.Line.Visible = msoFalse
        Else
            barColor = RGB(49, 112, 205)
            If w.Cells(r, 12) = 0 Then barColor = RGB(220, 71, 77)
            Set shape = w.Shapes.AddShape(msoShapeRoundedRectangle, x, y + 5, Application.Max(width, 2), 12)
            shape.Fill.ForeColor.RGB = barColor: shape.Line.Visible = msoFalse
        End If
        shape.Name = "Atlas_Bar_" & r: shape.Placement = xlMove
        If w.Cells(r, 4) = "CV" And CDbl(w.Cells(r, 10)) > 0 Then
            p = CDbl(w.Cells(r, 10)) / 100
            Set shape = w.Shapes.AddShape(msoShapeRectangle, x, y + 9, Application.Max(width * p, 1), 4)
            shape.Name = "Atlas_Actual_" & r: shape.Fill.ForeColor.RGB = RGB(27, 145, 108): shape.Line.Visible = msoFalse: shape.Placement = xlMove
        End If
        If IsDate(w.Cells(r, 16)) And IsDate(w.Cells(r, 17)) Then
            a = CDate(w.Cells(r, 16)): b = CDate(w.Cells(r, 17))
            x = DateX(w, first, a, viewMode): width = DateX(w, first, b + 1, viewMode) - x
            Set shape = w.Shapes.AddShape(msoShapeRectangle, x, y + 20, Application.Max(width, 2), 3)
            shape.Name = "Atlas_Base_" & r: shape.Fill.ForeColor.RGB = RGB(155, 166, 179): shape.Line.Visible = msoFalse: shape.Placement = xlMove
        End If
    Next r
    If UCase(CStr(book.Worksheets("CAU_HINH").Range("B8"))) = "CO" Then DrawLinks w, first, viewMode, last
    If reportDate >= first And reportDate <= lastDate + 1 Then
        x = DateX(w, first, reportDate, viewMode)
        Set shape = w.Shapes.AddLine(x, w.Cells(6, GANTT_COL).Top, x, w.Cells(last, GANTT_COL).Top + w.Rows(last).Height)
        shape.Name = "Atlas_ReportDate": shape.Line.ForeColor.RGB = RGB(231, 142, 43): shape.Line.DashStyle = msoLineDash: shape.Line.Weight = 1.5
    End If
    With w.PageSetup
        .Orientation = xlLandscape: .PaperSize = xlPaperA3: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .PrintTitleRows = "$1:$8": .PrintTitleColumns = "$A:$C"
        .PrintArea = w.Range(w.Cells(1, 1), w.Cells(last, c - 1)).Address
        .CenterFooter = "Atlas " & ATLAS_VERSION & " | &P / &N"
    End With
End Sub
Private Sub DrawLinks(ByVal w As Worksheet, ByVal first As Date, ByVal viewMode As String, ByVal last As Long)
    Dim ids As Object, r As Long, predecessor As Long, token As Variant, parts As Variant, relation As String
    Dim fromDate As Date, toDate As Date, arrow As Shape, x1 As Double, x2 As Double, y1 As Double, y2 As Double, count As Long
    Set ids = CreateObject("Scripting.Dictionary"): ids.CompareMode = vbTextCompare
    For r = FIRST_ROW To last: ids(CStr(w.Cells(r, 1))) = r: Next r
    For r = FIRST_ROW To last
        If Len(CStr(w.Cells(r, 8))) > 0 Then
            For Each token In Split(CStr(w.Cells(r, 8)), ";")
                parts = Split(Trim(CStr(token)), ":"): predecessor = ids(Trim(parts(0))): relation = UCase(Trim(parts(1)))
                If Left(relation, 1) = "F" And w.Cells(predecessor, 4) <> "MOC" Then fromDate = CDate(w.Cells(predecessor, 7)) + 1 Else fromDate = CDate(w.Cells(predecessor, 6))
                If Right(relation, 1) = "F" And w.Cells(r, 4) <> "MOC" Then toDate = CDate(w.Cells(r, 7)) + 1 Else toDate = CDate(w.Cells(r, 6))
                x1 = DateX(w, first, fromDate, viewMode): x2 = DateX(w, first, toDate, viewMode)
                y1 = w.Cells(predecessor, GANTT_COL).Top + 11: y2 = w.Cells(r, GANTT_COL).Top + 11
                Set arrow = w.Shapes.AddConnector(msoConnectorElbow, x1, y1, x2, y2)
                count = count + 1: arrow.Name = "Atlas_Link_" & count
                arrow.Line.ForeColor.RGB = RGB(139, 154, 171): arrow.Line.Weight = 0.75
                arrow.Line.EndArrowheadStyle = msoArrowheadTriangle: arrow.ZOrder msoSendToBack: arrow.Placement = xlMove
            Next token
        End If
    Next r
End Sub
Public Sub RibbonLinks(ByVal control As Object)
    On Error GoTo Failed
    Dim w As Worksheet, cfg As Worksheet
    Set w = ProjectSheet(ActiveWorkbook): Set cfg = w.Parent.Worksheets("CAU_HINH")
    cfg.Range("B8") = IIf(cfg.Range("B8") = "CO", "KHONG", "CO")
    CalculateProject w.Parent, True
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Private Function DateX(ByVal w As Worksheet, ByVal first As Date, ByVal d As Date, ByVal viewMode As String) As Double
    Dim offset As Double, monthStart As Date, days As Long
    Select Case viewMode
        Case "NGAY": offset = CDbl(d - first)
        Case "TUAN": offset = CDbl(d - first) / 7
        Case "THANG"
            monthStart = DateSerial(Year(d), Month(d), 1)
            days = Day(DateSerial(Year(d), Month(d) + 1, 0))
            offset = DateDiff("m", first, monthStart) + CDbl(Day(d) - 1) / days
    End Select
    DateX = w.Cells(9, GANTT_COL).Left + offset * w.Cells(9, GANTT_COL).Width
End Function

Public Sub RibbonDay(ByVal control As Object)
    SetScale "NGAY"
End Sub
Public Sub RibbonWeek(ByVal control As Object)
    SetScale "TUAN"
End Sub
Public Sub RibbonMonth(ByVal control As Object)
    SetScale "THANG"
End Sub
Private Sub SetScale(ByVal viewMode As String)
    On Error GoTo Failed
    Dim w As Worksheet
    Set w = ProjectSheet(ActiveWorkbook): w.Range("B6") = viewMode
    CalculateProject w.Parent, True
    w.Activate
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RibbonConfig(ByVal control As Object)
    ShowSheet "CAU_HINH"
End Sub
Public Sub RibbonResource(ByVal control As Object)
    ShowSheet "NGUON_LUC"
End Sub
Public Sub RibbonGuide(ByVal control As Object)
    ShowSheet "HUONG_DAN"
End Sub
Public Sub ShowSheet(ByVal name As String)
    On Error GoTo Failed
    Dim w As Worksheet
    Set w = ProjectSheet(ActiveWorkbook): w.Parent.Worksheets(name).Activate
    ActiveWindow.DisplayGridlines = False
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RibbonInfo(ByVal control As Object)
    On Error GoTo Failed
    Dim w As Worksheet, value As String
    Set w = ProjectSheet(ActiveWorkbook)
    value = InputBox("Tên dự án", "Atlas", w.Range("B2").Value)
    If Len(value) = 0 Then Exit Sub
    w.Range("B2") = value: w.Parent.Worksheets("CAU_HINH").Range("B3") = value
    w.Activate: w.Range("B5").Select
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RibbonAddTask(ByVal control As Object)
    AddTask "CV"
End Sub
Public Sub RibbonAddMilestone(ByVal control As Object)
    AddTask "MOC"
End Sub
Public Sub RibbonAddGroup(ByVal control As Object)
    AddTask "HM"
End Sub
Private Sub AddTask(ByVal kind As String)
    On Error GoTo Failed
    Dim w As Worksheet, r As Long, last As Long, id As String, seq As Long
    Set w = ProjectSheet(ActiveWorkbook): last = LastTask(w): r = last + 1
    If ActiveSheet.Name = "TIEN_DO" And ActiveCell.Row >= FIRST_ROW And ActiveCell.Row <= last Then r = ActiveCell.Row + 1
    seq = last
    Do
        seq = seq + 1: id = "CV" & Format(seq, "000")
    Loop While Application.CountIf(w.Columns(1), id) > 0
    w.Rows(r).Insert
    w.Range(w.Cells(r, 1), w.Cells(r, 17)).ClearContents
    AddDemoRow w, r, id, CStr(seq), IIf(kind = "HM", "Hạng mục mới", IIf(kind = "MOC", "Mốc mới", "Công việc mới")), kind, IIf(kind = "CV", 5, 0), "", "", 0
    If r > FIRST_ROW Then
        If w.Cells(r - 1, 4) = "HM" Then
            w.Cells(r, 2) = CStr(w.Cells(r - 1, 2)) & ".1"
        Else
            w.Cells(r, 2) = CStr(w.Cells(r - 1, 2))
        End If
    End If
    RenumberWBS w
    StyleTaskRows w
    w.Activate: w.Cells(r, 3).Select
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RibbonIndent(ByVal control As Object)
    ChangeDepth 1
End Sub
Public Sub RibbonOutdent(ByVal control As Object)
    ChangeDepth -1
End Sub
Private Sub ChangeDepth(ByVal delta As Long)
    On Error GoTo Failed
    Dim w As Worksheet, r As Long, last As Long, depth As Long, endRow As Long, j As Long, value As String
    Set w = ProjectSheet(ActiveWorkbook): last = LastTask(w): r = ActiveCell.Row
    If ActiveSheet.Name <> w.Name Or r < FIRST_ROW Or r > last Then Exit Sub
    depth = UBound(Split(CStr(w.Cells(r, 2)), "."))
    If depth + delta < 0 Or depth + delta > 9 Then Exit Sub
    If delta = 1 And r = FIRST_ROW Then Err.Raise 5, , "Dòng đầu không thể thụt cấp."
    If delta = 1 Then
        j = r - 1
        Do While j >= FIRST_ROW
            If UBound(Split(CStr(w.Cells(j, 2)), ".")) = depth Then Exit Do
            j = j - 1
        Loop
        If j < FIRST_ROW Then Err.Raise 5, , "Không tìm được hạng mục cha."
        If w.Cells(j, 4) <> "HM" Then Err.Raise 5, , "Hạng mục cha phải có loại HM. Đổi loại trước khi thụt cấp."
    End If
    endRow = r
    For j = r + 1 To last
        If UBound(Split(CStr(w.Cells(j, 2)), ".")) <= depth Then Exit For
        endRow = j
    Next j
    For j = r To endRow
        value = CStr(w.Cells(j, 2))
        If delta = 1 Then value = value & ".1" Else value = Left(value, InStrRev(value, ".") - 1)
        w.Cells(j, 2) = value
    Next j
    RenumberWBS w
    StyleTaskRows w
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RibbonRenumber(ByVal control As Object)
    On Error GoTo Failed
    RenumberWBS ProjectSheet(ActiveWorkbook)
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RenumberWBS(ByVal w As Worksheet)
    Dim counters(0 To 10) As Long, r As Long, depth As Long, previous As Long, j As Long, value As String
    For r = FIRST_ROW To LastTask(w)
        depth = Application.Min(UBound(Split(CStr(w.Cells(r, 2)), ".")), 9)
        If r = FIRST_ROW Then depth = 0
        If depth > previous + 1 Then depth = previous + 1
        counters(depth) = counters(depth) + 1
        For j = depth + 1 To 10: counters(j) = 0: Next j
        value = CStr(counters(0))
        For j = 1 To depth: value = value & "." & counters(j): Next j
        w.Cells(r, 2).NumberFormat = "@": w.Cells(r, 2) = value
        previous = depth
    Next r
End Sub
Public Sub RibbonDelete(ByVal control As Object)
    On Error GoTo Failed
    Dim w As Worksheet, r As Long, last As Long, id As String, j As Long, token As Variant, parts As Variant
    Set w = ProjectSheet(ActiveWorkbook): r = ActiveCell.Row: last = LastTask(w)
    If ActiveSheet.Name <> "TIEN_DO" Or r < FIRST_ROW Or r > last Then Exit Sub
    If last = FIRST_ROW Then Err.Raise 5, , "Giữ ít nhất một công việc trong dự án."
    If w.Cells(r, 4) = "HM" Then
        For j = FIRST_ROW To last
            If Left(CStr(w.Cells(j, 2)), Len(CStr(w.Cells(r, 2))) + 1) = CStr(w.Cells(r, 2)) & "." Then Err.Raise 5, , "Xóa công việc con trước khi xóa hạng mục."
        Next j
    End If
    id = CStr(w.Cells(r, 1))
    For j = FIRST_ROW To last
        If Len(CStr(w.Cells(j, 8))) > 0 Then
            For Each token In Split(CStr(w.Cells(j, 8)), ";")
                parts = Split(CStr(token), ":")
                If StrComp(Trim(CStr(parts(0))), id, vbTextCompare) = 0 Then Err.Raise 5, , "Công việc đang được tham chiếu bởi " & w.Cells(j, 1)
            Next token
        End If
    Next j
    If MsgBox("Xóa " & id & "? Thao tác VBA không có Undo.", vbYesNo + vbQuestion, "Atlas") <> vbYes Then Exit Sub
    w.Rows(r).Delete
    RenumberWBS w: StyleTaskRows w
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Private Sub BuildGuide(ByVal book As Workbook)
    Dim w As Worksheet, lines As Variant, i As Long
    Set w = book.Worksheets("HUONG_DAN")
    w.Range("A1:H1").Merge: w.Range("A1") = "ATLAS 0.3 | HƯỚNG DẪN NHANH"
    TitleStyle w.Range("A1:H1")
    lines = Array("01  Nhập thông tin dự án, ngày bắt đầu và ngày báo cáo tại TIEN_DO.", "02  HM là hạng mục tổng hợp, CV là công việc, MOC là mốc thời lượng 0.", "03  WBS con cần hạng mục cha HM. Dùng thêm việc / thụt cấp / nâng cấp để tổ chức.", "04  ID giữ ổn định; quan hệ dùng ID:FS:lag, ngăn cách bằng dấu chấm phẩy.", "05  Bốn quan hệ FS/SS/FF/SF; lag là số nguyên ngày làm việc, có thể âm.", "06  Cấu hình lịch: 7 ký tự 0/1 (T2–CN); nhập ngày nghỉ lễ tại CAU_HINH.", "07  % HT nhập 0–100. Cập nhật tiến độ để tính lại lịch và vẽ Gantt.", "08  Lưu baseline trước khi thay đổi kế hoạch. Bản này không ghi đè baseline.", "09  Gán nguồn lực bằng NC:12;KS:1 trong cột Nguồn lực; đơn giá ở NGUON_LUC.", "10  Báo cáo nguồn lực là kế hoạch, giả định số lượng cố định mỗi ngày làm việc.", "11  Ghi nhận kỳ lưu % thực tế có trọng số thời lượng tại ngày báo cáo.", "12  Đường cong S thực tế chỉ thể hiện các kỳ đã ghi nhận; không suy diễn lịch sử.", "13  Đường cong S kế hoạch phân bổ tuyến tính theo ngày làm việc; ưu tiên baseline đầy đủ.", "14  Đây là lập lịch CPM từ ngày bắt đầu, chưa reschedule theo remaining duration / actual dates.", "15  Thang ngày tối đa 731 ngày; tuần / tháng cho dự án dài hơn. Tối đa 2000 dòng.", "16  Dữ liệu demo minh họa, không phải dữ liệu thực tế H2. Workbook dữ liệu lưu riêng dạng .xlsx.", "17  Thay đổi dữ liệu nguồn không tự cập nhật biểu đồ; bấm Cập nhật / Báo cáo.", "18  Nguồn lực chưa được cân bằng tự động. Tổng hợp có thể vượt năng lực công trường.")
    For i = 0 To UBound(lines)
        w.Range(w.Cells(i + 4, 1), w.Cells(i + 4, 8)).Merge: w.Cells(i + 4, 1) = lines(i)
        w.Rows(i + 4).RowHeight = 29
    Next i
    w.Columns("A:H").ColumnWidth = 16
End Sub
