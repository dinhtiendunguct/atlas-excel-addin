Attribute VB_Name = "Atlas"
Option Explicit

Public Sub RibbonUpdate(ByVal control As Object)
    Dim updater As String, shell As Object, command As String
    updater = ThisWorkbook.Path & Application.PathSeparator & "Atlas.Update.ps1"
    If Dir(updater) = "" Then
        MsgBox "Chua cai bo cap nhat. Hay chay Install-Atlas.ps1 mot lan.", vbInformation, "Atlas"
        Exit Sub
    End If
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & updater & Chr(34)
    Set shell = CreateObject("WScript.Shell")
    shell.Run command, 0, False
End Sub

Public Sub RibbonNew(ByVal control As Object)
    AtlasNew
End Sub
Public Sub RibbonCalculate(ByVal control As Object)
    AtlasCalculate
End Sub
Public Sub RibbonBaseline(ByVal control As Object)
    AtlasBaseline
End Sub

Private Function WorkingDay(ByVal d As Date) As Boolean
    WorkingDay = Weekday(d, vbMonday) <= 6
End Function

Private Function DayAt(ByVal origin As Date, ByVal offset As Long) As Date
    Dim d As Date, i As Long
    d = origin
    Do While Not WorkingDay(d): d = d + 1: Loop
    For i = 1 To offset
        d = d + 1
        Do While Not WorkingDay(d): d = d + 1: Loop
    Next i
    DayAt = d
End Function

Public Sub AtlasNew()
    Dim w As Worksheet, headers As Variant, i As Long
    Set w = ActiveWorkbook.Worksheets.Add
    w.Name = "Atlas_" & Format(Now, "hhmmss")
    w.Cells(1, 1) = "ATLAS | TIEN DO THI CONG"
    w.Cells(2, 1) = "Ngay bat dau": w.Cells(2, 2) = Date
    w.Cells(3, 1) = "Lich: Thu 2 - Thu 7. Quan he: ID:FS:lag;ID:SS:lag. Lag tinh bang ngay lam viec."
    headers = Array("ID", "WBS", "Cong viec", "So ngay", "Cong viec truoc", "% hoan thanh", "Bat dau", "Ket thuc", "Du tru", "Duong gang", "Baseline start", "Baseline finish")
    For i = 0 To UBound(headers): w.Cells(5, i + 1) = headers(i): Next i
    w.Range("A6:F9").Value = DemoData()
    w.Range("B2").NumberFormat = "dd/mm/yyyy"
    w.Columns("A:B").ColumnWidth = 12
    w.Columns("C").ColumnWidth = 32
    w.Columns("D:F").ColumnWidth = 20
    w.Columns("G:L").ColumnWidth = 15
    w.Range("A5:L5").Interior.Color = RGB(24, 49, 75)
    w.Range("A5:L5").Font.Color = vbWhite
    w.Range("A5:L5").Font.Bold = True
    w.Range("A6:F1000").Font.Color = RGB(0, 85, 170)
    w.Activate
    ActiveWindow.SplitRow = 5: ActiveWindow.SplitColumn = 3: ActiveWindow.FreezePanes = True
    AtlasCalculate
End Sub

Private Function DemoData() As Variant
    Dim a(1 To 4, 1 To 6) As Variant
    a(1, 1) = "A": a(1, 2) = "1": a(1, 3) = "Chuan bi": a(1, 4) = 3: a(1, 6) = 100
    a(2, 1) = "B": a(2, 2) = "2": a(2, 3) = "Thi cong mong": a(2, 4) = 6: a(2, 5) = "A:FS:0": a(2, 6) = 30
    a(3, 1) = "C": a(3, 2) = "3": a(3, 3) = "Chuan bi vat tu": a(3, 4) = 4: a(3, 5) = "A:FS:0": a(3, 6) = 50
    a(4, 1) = "D": a(4, 2) = "4": a(4, 3) = "Ket cau than": a(4, 4) = 10: a(4, 5) = "B:FS:0;C:FS:0": a(4, 6) = 0
    DemoData = a
End Function

Public Sub AtlasCalculate()
    On Error GoTo Failed
    Dim w As Worksheet, n As Long, i As Long, j As Long, pass As Long, count As Long
    Dim ids As Object, duration() As Long, early() As Long, late() As Long, done() As Boolean
    Dim links() As Collection, item As Variant, parts As Variant, token As Variant
    Dim predecessor As Long, lag As Long, weight As Long, relation As String, ready As Boolean, candidate As Long, finish As Long
    Set w = ActiveSheet
    If w.Cells(5, 1).Value <> "ID" Then Err.Raise 5, , "Hay chon sheet Atlas."
    If Not IsDate(w.Cells(2, 2).Value) Then Err.Raise 5, , "Ngay bat dau khong hop le."
    n = w.Cells(w.Rows.Count, 1).End(xlUp).Row - 5
    If n < 1 Then Err.Raise 5, , "Chua co cong viec."
    ReDim duration(1 To n): ReDim early(1 To n): ReDim late(1 To n): ReDim done(1 To n): ReDim links(1 To n)
    Set ids = CreateObject("Scripting.Dictionary")
    ids.CompareMode = vbTextCompare
    For i = 1 To n
        token = Trim(CStr(w.Cells(i + 5, 1).Value))
        If token = "" Or ids.Exists(token) Then Err.Raise 5, , "ID trong hoac trung tai dong " & i + 5
        ids.Add token, i
        If Not IsNumeric(w.Cells(i + 5, 4)) Then Err.Raise 5, , "So ngay khong hop le tai dong " & i + 5
        candidate = CLng(w.Cells(i + 5, 4).Value)
        If candidate < 0 Or candidate <> CDbl(w.Cells(i + 5, 4).Value) Then Err.Raise 5, , "So ngay phai la so nguyen >= 0."
        duration(i) = candidate
        If Not IsNumeric(w.Cells(i + 5, 6)) Then Err.Raise 5, , "% hoan thanh phai tu 0 den 100."
        If w.Cells(i + 5, 6) < 0 Or w.Cells(i + 5, 6) > 100 Then Err.Raise 5, , "% hoan thanh phai tu 0 den 100."
        Set links(i) = New Collection
    Next i
    For i = 1 To n
        If Trim(CStr(w.Cells(i + 5, 5))) <> "" Then
            For Each token In Split(CStr(w.Cells(i + 5, 5)), ";")
                parts = Split(Trim(token), ":")
                If UBound(parts) <> 2 Then Err.Raise 5, , "Quan he can theo mau ID:FS:0 tai dong " & i + 5
                If Not ids.Exists(Trim(parts(0))) Then Err.Raise 5, , "Khong tim thay ID " & parts(0)
                predecessor = ids(Trim(parts(0)))
                If predecessor = i Then Err.Raise 5, , "Cong viec tu phu thuoc."
                If Not IsNumeric(parts(2)) Then Err.Raise 5, , "Lag khong hop le."
                lag = CLng(parts(2))
                If lag <> CDbl(parts(2)) Then Err.Raise 5, , "Lag phai la so nguyen."
                relation = UCase(Trim(parts(1)))
                Select Case relation
                    Case "FS": weight = duration(predecessor) + lag
                    Case "SS": weight = lag
                    Case "FF": weight = duration(predecessor) - duration(i) + lag
                    Case "SF": weight = -duration(i) + lag
                    Case Else: Err.Raise 5, , "Chi ho tro FS, SS, FF, SF."
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
                    early(i) = candidate: done(i) = True: count = count + 1
                    If early(i) + duration(i) > finish Then finish = early(i) + duration(i)
                End If
            End If
        Next i
        If count = n Then Exit For
    Next pass
    If count <> n Then Err.Raise 5, , "Phat hien quan he vong. Khong ghi ket qua."
    For i = 1 To n: late(i) = finish - duration(i): Next i
    For pass = 1 To n
        For i = n To 1 Step -1
            For Each item In links(i)
                predecessor = item(0)
                If late(i) - item(1) < late(predecessor) Then late(predecessor) = late(i) - item(1)
            Next item
        Next i
    Next pass
    For i = 1 To n
        w.Cells(i + 5, 7) = DayAt(CDate(w.Cells(2, 2)), early(i))
        w.Cells(i + 5, 8) = DayAt(CDate(w.Cells(2, 2)), early(i) + IIf(duration(i) > 0, duration(i) - 1, 0))
        w.Cells(i + 5, 9) = late(i) - early(i)
        w.Cells(i + 5, 10) = IIf(late(i) = early(i), "GANG", "")
    Next i
    w.Range("G6:H" & n + 5).NumberFormat = "dd/mm/yyyy"
    DrawGantt w, n
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub

Private Sub DrawGantt(ByVal w As Worksheet, ByVal n As Long)
    Dim first As Date, last As Date, d As Date, c As Long, r As Long, color As Long, total As Long, elapsed As Long
    first = Application.Min(w.Range("G6:G" & n + 5))
    last = Application.Max(w.Range("H6:H" & n + 5))
    If last - first > 730 Then Err.Raise 5, , "Gantt ban dau gioi han 731 ngay."
    w.Range(w.Cells(5, 13), w.Cells(Application.Max(n + 5, w.UsedRange.Rows.Count), Application.Max(13, w.UsedRange.Columns.Count))).Clear
    For d = first To last
        c = 13 + CLng(d - first)
        w.Cells(5, c) = d: w.Cells(5, c).NumberFormat = "dd/mm"
        w.Columns(c).ColumnWidth = 4
        w.Cells(5, c).Orientation = 90
        If Not WorkingDay(d) Then w.Range(w.Cells(5, c), w.Cells(n + 5, c)).Interior.Color = RGB(237, 240, 244)
        For r = 6 To n + 5
            If d >= w.Cells(r, 7) And d <= w.Cells(r, 8) And WorkingDay(d) Then
                color = RGB(49, 130, 190)
                If w.Cells(r, 10) = "GANG" Then color = RGB(216, 70, 70)
                total = CLng(w.Cells(r, 4)): elapsed = Application.WorksheetFunction.NetworkDays_Intl(w.Cells(r, 7), d, 11)
                If total > 0 And elapsed <= Int(total * w.Cells(r, 6) / 100) Then color = RGB(36, 145, 113)
                w.Cells(r, c).Interior.Color = color
                If total = 0 Then w.Cells(r, c) = ChrW(&H25C6)
            End If
            If d = Date Then w.Cells(r, c).Borders(xlEdgeLeft).Color = RGB(245, 160, 0)
        Next r
    Next d
    With w.PageSetup
        .Orientation = xlLandscape: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .PrintTitleRows = "$1:$5": .PrintArea = w.Range(w.Cells(1, 1), w.Cells(n + 5, c)).Address
    End With
End Sub

Public Sub AtlasBaseline()
    Dim w As Worksheet, n As Long
    Set w = ActiveSheet
    If w.Cells(5, 1) <> "ID" Then Exit Sub
    n = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    If Application.CountA(w.Range("K6:L" & n)) > 0 Then
        MsgBox "Baseline da ton tai. Ban dau khong ghi de baseline.", vbInformation: Exit Sub
    End If
    w.Range("K6:L" & n).Value = w.Range("G6:H" & n).Value
    w.Range("K6:L" & n).NumberFormat = "dd/mm/yyyy"
End Sub
