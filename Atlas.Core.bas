Attribute VB_Name = "Atlas"
Option Explicit
Public Const ATLAS_VERSION As String = "0.3.0"
Public Const FIRST_ROW As Long = 9
Public Const GANTT_COL As Long = 18
Private holidays As Object
Private weekendMask As String

Public Function U(ByVal hex As String) As String
    Dim i As Long
    For i = 1 To Len(hex) Step 4
        U = U & ChrW(CLng("&H" & Mid(hex, i, 4)))
    Next i
End Function

Public Sub RibbonUpdate(ByVal control As Object)
    Dim updater As String, sh As Object
    updater = ThisWorkbook.Path & Application.PathSeparator & "Atlas.Update.ps1"
    If Dir(updater) = "" Then
        MsgBox "Chưa có bộ cập nhật. Hãy chạy Install-Atlas.ps1 một lần.", vbInformation, "Atlas"
        Exit Sub
    End If
    Set sh = CreateObject("WScript.Shell")
    sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & updater & Chr(34), 0, False
End Sub
Public Sub RibbonNew(ByVal control As Object)
    CreateProject False
End Sub
Public Sub RibbonDemo(ByVal control As Object)
    CreateProject True
End Sub
Public Sub RibbonCalculate(ByVal control As Object)
    AtlasCalculate
End Sub
Public Sub RibbonBaseline(ByVal control As Object)
    AtlasBaseline
End Sub
Public Sub RibbonCheck(ByVal control As Object)
    On Error GoTo Failed
    CalculateProject ActiveWorkbook, False
    MsgBox "Dữ liệu và quan hệ hợp lệ. Đã tính lại lịch và dự trữ.", vbInformation, "Atlas"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Function ProjectSheet(ByVal book As Workbook) As Worksheet
    On Error Resume Next
    Set ProjectSheet = book.Worksheets("TIEN_DO")
    On Error GoTo 0
    If ProjectSheet Is Nothing Then Err.Raise 5, , "Chọn workbook Atlas 0.3. Dùng Tạo dự án hoặc Mẫu demo để bắt đầu; bảng Atlas cũ được giữ nguyên."
    If ProjectSheet.Cells(8, 1).Value <> "ID" Then Err.Raise 5, , "Cấu trúc TIEN_DO không hợp lệ."
End Function
Public Function LastTask(ByVal w As Worksheet) As Long
    LastTask = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    If LastTask < FIRST_ROW Then Err.Raise 5, , "Chưa có công việc."
End Function
Public Sub LoadCalendar(ByVal book As Workbook)
    Dim cfg As Worksheet, r As Long, i As Long, v As Variant
    Set cfg = book.Worksheets("CAU_HINH")
    weekendMask = CStr(cfg.Range("B6").Value)
    If Len(weekendMask) <> 7 Then Err.Raise 5, , "Lịch nghỉ phải có 7 ký tự 0/1, từ thứ Hai đến Chủ nhật."
    For i = 1 To 7
        If Mid(weekendMask, i, 1) <> "0" And Mid(weekendMask, i, 1) <> "1" Then Err.Raise 5, , "Lịch nghỉ chỉ nhận 0 hoặc 1."
    Next i
    If weekendMask = "1111111" Then Err.Raise 5, , "Lịch cần ít nhất một ngày làm việc mỗi tuần."
    Set holidays = CreateObject("Scripting.Dictionary")
    For r = 12 To cfg.Cells(cfg.Rows.Count, 1).End(xlUp).Row
        v = cfg.Cells(r, 1).Value
        If Len(CStr(v)) > 0 Then
            If Not IsDate(v) Then Err.Raise 5, , "Ngày nghỉ không hợp lệ tại CAU_HINH!A" & r
            holidays(CStr(CLng(CDate(v)))) = True
        End If
    Next r
End Sub
Public Function IsWorking(ByVal d As Date) As Boolean
    IsWorking = Mid(weekendMask, Weekday(d, vbMonday), 1) = "0"
    If holidays.Exists(CStr(CLng(d))) Then IsWorking = False
End Function
Public Function WorkCount(ByVal a As Date, ByVal b As Date) As Long
    Dim d As Date
    For d = a To b
        If IsWorking(d) Then WorkCount = WorkCount + 1
    Next d
End Function
Public Sub AtlasCalculate()
    On Error GoTo Failed
    CalculateProject ActiveWorkbook, True
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub CalculateProject(ByVal book As Workbook, Optional ByVal render As Boolean = True)
    Dim w As Worksheet, n As Long, last As Long, i As Long, j As Long, r As Long, pass As Long, processed As Long
    Dim ids As Object, wbsIds As Object, durations() As Long, early() As Long, late() As Long, done() As Boolean
    Dim isGroup() As Boolean, links() As Collection, item As Variant, parts As Variant, token As Variant
    Dim id As String, code As String, kind As String, relation As String, predecessor As Long, lag As Long, weight As Long
    Dim ready As Boolean, candidate As Long, finish As Long, taskCount As Long, v As Variant, origin As Date
    Dim dates(0 To 3651) As Date, d As Date, k As Long, startDates() As Double, endDates() As Double
    Dim total As Double, actual As Double, critical As Long, groupStart As Double, groupFinish As Double
    Set w = ProjectSheet(book): last = LastTask(w): n = last - FIRST_ROW + 1
    If n > 2000 Then Err.Raise 5, , "Bản 0.3 hỗ trợ tối đa 2.000 dòng."
    If Not IsDate(w.Range("B4")) Or Not IsDate(w.Range("B5")) Then Err.Raise 5, , "Ngày bắt đầu / ngày báo cáo không hợp lệ."
    origin = DateValue(w.Range("B4").Value)
    LoadCalendar book
    d = origin
    Do While Not IsWorking(d)
        d = d + 1
        If d - origin > 3660 Then Err.Raise 5, , "Không tìm được ngày làm việc."
    Loop
    dates(0) = d
    For k = 1 To 3651
        d = d + 1
        Do While Not IsWorking(d): d = d + 1: Loop
        dates(k) = d
    Next k
    ReDim durations(1 To n): ReDim early(1 To n): ReDim late(1 To n): ReDim done(1 To n)
    ReDim links(1 To n): ReDim isGroup(1 To n): ReDim startDates(1 To n): ReDim endDates(1 To n)
    Set ids = CreateObject("Scripting.Dictionary"): ids.CompareMode = vbTextCompare
    Set wbsIds = CreateObject("Scripting.Dictionary"): wbsIds.CompareMode = vbTextCompare
    For i = 1 To n
        r = i + FIRST_ROW - 1
        id = Trim(CStr(w.Cells(r, 1).Value)): code = Trim(CStr(w.Cells(r, 2).Value))
        If id = "" Or ids.Exists(id) Then Err.Raise 5, , "ID trống hoặc trùng tại dòng " & r
        If InStr(id, ":") Or InStr(id, ";") Then Err.Raise 5, , "ID không được chứa : hoặc ; tại dòng " & r
        If code = "" Or wbsIds.Exists(code) Then Err.Raise 5, , "WBS trống hoặc trùng tại dòng " & r
        ids.Add id, i: wbsIds.Add code, i
        kind = UCase(Trim(CStr(w.Cells(r, 4).Value)))
        If kind <> "HM" And kind <> "CV" And kind <> "MOC" Then Err.Raise 5, , "Loại chỉ nhận HM, CV hoặc MOC tại dòng " & r
        isGroup(i) = kind = "HM": Set links(i) = New Collection
        If isGroup(i) Then
            If Len(Trim(CStr(w.Cells(r, 8)))) > 0 Or Len(Trim(CStr(w.Cells(r, 9)))) > 0 Then Err.Raise 5, , "HM không nhận quan hệ / nguồn lực; gán vào công việc con. Dòng " & r
            done(i) = True: processed = processed + 1
        Else
            taskCount = taskCount + 1: v = w.Cells(r, 5).Value
            If Not IsNumeric(v) Or IsEmpty(v) Then Err.Raise 5, , "Thời lượng phải là số nguyên tại dòng " & r
            If CDbl(v) < 0 Or CDbl(v) > 3650 Or CDbl(v) <> Fix(CDbl(v)) Then Err.Raise 5, , "Thời lượng phải từ 0 đến 3650 ngày tại dòng " & r
            durations(i) = CLng(v)
            If (kind = "MOC" And durations(i) <> 0) Or (kind = "CV" And durations(i) = 0) Then Err.Raise 5, , "MOC có thời lượng 0; CV có thời lượng > 0. Dòng " & r
            v = w.Cells(r, 10).Value
            If IsEmpty(v) Or Not IsNumeric(v) Then Err.Raise 5, , "% hoàn thành phải là số 0–100 tại dòng " & r
            If CDbl(v) < 0 Or CDbl(v) > 100 Then Err.Raise 5, , "% hoàn thành ngoài khoảng 0–100 tại dòng " & r
        End If
    Next i
    If taskCount = 0 Then Err.Raise 5, , "Cần ít nhất một công việc hoặc mốc."
    For i = 1 To n
        r = i + FIRST_ROW - 1: code = CStr(w.Cells(r, 2))
        Do While InStrRev(code, ".") > 0
            code = Left(code, InStrRev(code, ".") - 1)
            If Not wbsIds.Exists(code) Then Err.Raise 5, , "Thiếu hạng mục cha WBS " & code & " tại dòng " & r
            If Not isGroup(wbsIds(code)) Then Err.Raise 5, , "WBS cha phải có loại HM: " & code
        Loop
        If Not isGroup(i) And Len(Trim(CStr(w.Cells(r, 8)))) > 0 Then
            For Each token In Split(CStr(w.Cells(r, 8)), ";")
                parts = Split(Trim(CStr(token)), ":")
                If UBound(parts) <> 2 Then Err.Raise 5, , "Quan hệ dùng ID:FS:0;ID:SS:2 tại dòng " & r
                id = Trim(parts(0))
                If Not ids.Exists(id) Then Err.Raise 5, , "Không tìm thấy ID " & id & " tại dòng " & r
                predecessor = ids(id)
                If predecessor = i Or isGroup(predecessor) Then Err.Raise 5, , "Không liên kết chính nó hoặc HM tại dòng " & r
                If Not IsNumeric(parts(2)) Then Err.Raise 5, , "Lag phải là số nguyên tại dòng " & r
                If CDbl(parts(2)) <> Fix(CDbl(parts(2))) Or Abs(CDbl(parts(2))) > 3650 Then Err.Raise 5, , "Lag ngoài phạm vi tại dòng " & r
                lag = CLng(parts(2)): relation = UCase(Trim(parts(1)))
                Select Case relation
                    Case "FS": weight = durations(predecessor) + lag
                    Case "SS": weight = lag
                    Case "FF": weight = durations(predecessor) - durations(i) + lag
                    Case "SF": weight = -durations(i) + lag
                    Case Else: Err.Raise 5, , "Quan hệ chỉ nhận FS/SS/FF/SF tại dòng " & r
                End Select
                links(i).Add Array(predecessor, weight)
            Next token
        End If
    Next i
    For pass = 1 To n
        For i = 1 To n
            If Not done(i) Then
                ready = True: candidate = 0
                For Each item In links(i)
                    If Not done(item(0)) Then ready = False
                    If early(item(0)) + item(1) > candidate Then candidate = early(item(0)) + item(1)
                Next item
                If ready Then
                    early(i) = candidate: done(i) = True: processed = processed + 1
                    If early(i) + durations(i) > finish Then finish = early(i) + durations(i)
                End If
            End If
        Next i
        If processed = n Then Exit For
    Next pass
    If processed <> n Then Err.Raise 5, , "Quan hệ vòng tròn. Kết quả cũ chưa bị ghi đè."
    If finish > 3650 Then Err.Raise 5, , "Tiến độ vượt 3650 ngày làm việc."
    For i = 1 To n: late(i) = finish - durations(i): Next i
    For pass = 1 To n
        For i = n To 1 Step -1
            For Each item In links(i)
                predecessor = item(0)
                If late(i) - item(1) < late(predecessor) Then late(predecessor) = late(i) - item(1)
            Next item
        Next i
    Next pass
    For i = 1 To n
        If Not isGroup(i) Then
            startDates(i) = CDbl(dates(early(i)))
            endDates(i) = CDbl(dates(early(i) + IIf(durations(i) > 0, durations(i) - 1, 0)))
        End If
    Next i
    For i = 1 To n
        If isGroup(i) Then
            groupStart = 0: groupFinish = 0: code = CStr(w.Cells(i + FIRST_ROW - 1, 2)) & "."
            For j = 1 To n
                If Not isGroup(j) And Left(CStr(w.Cells(j + FIRST_ROW - 1, 2)), Len(code)) = code Then
                    If groupStart = 0 Or startDates(j) < groupStart Then groupStart = startDates(j)
                    If endDates(j) > groupFinish Then groupFinish = endDates(j)
                End If
            Next j
            If groupStart = 0 Then Err.Raise 5, , "HM chưa có công việc con: " & w.Cells(i + FIRST_ROW - 1, 1)
            startDates(i) = groupStart: endDates(i) = groupFinish
        End If
    Next i
    ValidateResources book
    Application.ScreenUpdating = False
    For i = 1 To n
        r = i + FIRST_ROW - 1
        w.Cells(r, 6) = CDate(startDates(i)): w.Cells(r, 7) = CDate(endDates(i))
        If Not isGroup(i) Then
            w.Cells(r, 12) = late(i) - early(i)
            If late(i) = early(i) Then critical = critical + 1
            total = total + durations(i): actual = actual + durations(i) * CDbl(w.Cells(r, 10)) / 100
        Else
            w.Cells(r, 5) = WorkCount(CDate(startDates(i)), CDate(endDates(i)))
            GroupProgress w, r, last
            w.Cells(r, 12).ClearContents
        End If
        w.Cells(r, 13).ClearContents
        If IsDate(w.Cells(r, 16)) And IsDate(w.Cells(r, 17)) Then
            d = CDate(w.Cells(r, 17))
            If CDate(endDates(i)) >= d Then
                w.Cells(r, 13) = WorkCount(d + 1, CDate(endDates(i)))
            Else
                w.Cells(r, 13) = -WorkCount(CDate(endDates(i)) + 1, d)
            End If
        End If
    Next i
    w.Range("F4") = CDate(Application.Max(w.Range("G9:G" & last)))
    w.Range("J4") = finish
    If total > 0 Then w.Range("J5") = actual / total Else w.Range("J5") = 0
    w.Range("N4") = critical: w.Range("N5") = "v" & ATLAS_VERSION
    If render Then
        StyleTaskRows w
        DrawGantt book
    End If
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub
Private Sub GroupProgress(ByVal w As Worksheet, ByVal r As Long, ByVal last As Long)
    Dim j As Long, prefix As String, total As Double, actual As Double
    prefix = CStr(w.Cells(r, 2)) & "."
    For j = FIRST_ROW To last
        If w.Cells(j, 4) <> "HM" And Left(CStr(w.Cells(j, 2)), Len(prefix)) = prefix Then
            total = total + CDbl(w.Cells(j, 5)): actual = actual + CDbl(w.Cells(j, 5)) * CDbl(w.Cells(j, 10))
        End If
    Next j
    If total > 0 Then w.Cells(r, 10) = actual / total Else w.Cells(r, 10) = 0
End Sub
Public Sub AtlasBaseline()
    On Error GoTo Failed
    Dim w As Worksheet, last As Long
    Set w = ProjectSheet(ActiveWorkbook): last = LastTask(w)
    If Application.CountA(w.Range("P9:Q" & last)) > 0 Then Err.Raise 5, , "Baseline đã tồn tại. Bản này không ghi đè baseline."
    CalculateProject w.Parent, False
    w.Range("P9:Q" & last).Value = w.Range("F9:G" & last).Value
    w.Range("P9:Q" & last).NumberFormat = "dd/mm/yyyy"
    DrawGantt w.Parent
    MsgBox "Đã lưu baseline. Thanh xám là kế hoạch đã chốt.", vbInformation, "Atlas"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
