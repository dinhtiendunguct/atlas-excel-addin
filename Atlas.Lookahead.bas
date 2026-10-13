Attribute VB_Name = "AtlasLookahead"
Option Explicit

Private Function ControlSheet(ByVal book As Workbook, ByVal title As String, ByVal headers As Variant) As Worksheet
    On Error Resume Next
    Set ControlSheet = book.Worksheets(title)
    On Error GoTo 0
    If Not ControlSheet Is Nothing Then Exit Function
    Set ControlSheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    With ControlSheet
        .Name = title
        .Cells.Font.Name = "Calibri": .Cells.Font.Size = 10
        .Range("A1:H1").Merge: .Range("A1") = "ATLAS | " & title
        TitleStyle .Range("A1:H1")
        PutHeaders ControlSheet, 5, headers
        .Columns.ColumnWidth = 18: .Columns("C").ColumnWidth = 32
        .Rows(5).RowHeight = 32: .Tab.Color = RGB(0, 133, 128)
    End With
End Function

Public Sub EnsureLookahead(ByVal book As Workbook)
    Dim w As Worksheet, category As Variant, r As Long
    Set w = ProjectSheet(book)
    Set w = ControlSheet(book, "SAN_SANG", Array("Activity ID", "Điều kiện", "Người xác nhận", "Ngày xác nhận", "Bằng chứng / đường dẫn", "Kết quả"))
    w.Range("A3") = "Mỗi ID cần đủ 6 điều kiện PASS, có người, ngày và bằng chứng. Thiếu dữ liệu = HOLD."
    w.Range("A6:F12005").Interior.Color = RGB(255, 249, 216)
    With w.Range("B6:B12005").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "BAN_VE,VAT_TU,MAT_BANG,QA_QC,HSE,NHAN_LUC"
    End With
    With w.Range("F6:F12005").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "PASS,HOLD"
    End With
    w.Range("D6:D12005").NumberFormat = "dd/mm/yyyy"
    Set w = ControlSheet(book, "RANG_BUOC", Array("Constraint ID", "Activity ID", "Nội dung ràng buộc", "Phụ trách", "Hạn xử lý", "Trạng thái", "Bằng chứng đóng", "Người xác nhận", "Ngày đóng", "Cảnh báo"))
    w.Range("A3") = "CLOSED chỉ hợp lệ khi có bằng chứng, người xác nhận và ngày đóng không sau ngày báo cáo."
    w.Range("A6:I2005").Interior.Color = RGB(255, 249, 216)
    w.Range("E6:E2005").NumberFormat = "dd/mm/yyyy": w.Range("I6:I2005").NumberFormat = "dd/mm/yyyy"
    With w.Range("F6:F2005").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "OPEN,CLOSED"
    End With
    Set w = ControlSheet(book, "LOOKAHEAD_21N", Array("Activity ID", "WBS", "Công việc", "Zone", "Phụ trách", "Bắt đầu", "Kết thúc", "% HT", "Trạng thái", "Lý do HOLD", "Ràng buộc mở", "Quá hạn"))
    w.Range("A3") = "Cửa sổ 21 ngày lịch từ ngày báo cáo; gồm việc đang làm và việc tồn quá hạn. Không tự phê duyệt huy động."
    w.Columns("J").ColumnWidth = 65
    Set w = ControlSheet(book, "DIEU_HANH", Array("Chỉ tiêu", "Giá trị"))
    w.Range("A3") = "ATLAS | Từ tiến độ đến điều kiện triển khai tại công trường"
    w.Columns("A").ColumnWidth = 42
End Sub

Private Function ValidConfirmation(ByVal owner As Variant, ByVal day As Variant, ByVal proof As Variant, ByVal reportDay As Date) As Boolean
    If Len(Trim(CStr(owner))) = 0 Or Len(Trim(CStr(proof))) = 0 Then Exit Function
    If Not IsDate(day) Then Exit Function
    ValidConfirmation = CDate(day) <= reportDay And CDate(day) > 0
End Function

Public Sub RefreshLookahead(ByVal book As Workbook)
    Dim tasks As Worksheet, ready As Worksheet, constraints As Worksheet, view As Worksheet, dash As Worksheet
    Dim ids As Object, seen As Object, categories As Variant, id As String, key As String, reason As String
    Dim r As Long, j As Long, n As Long, outputRow As Long, openCount As Long, overdue As Long, readyCount As Long, holdCount As Long
    Dim reportDay As Date, finishDay As Date, closed As Boolean, dataIssues As Long, EachCategoryIndex As Long
    Set tasks = ProjectSheet(book)
    If Not IsDate(tasks.Range("B5").Value) Then Err.Raise 5, , "Ngày báo cáo không hợp lệ."
    reportDay = DateValue(tasks.Range("B5").Value)
    EnsureLookahead book
    Set ready = book.Worksheets("SAN_SANG"): Set constraints = book.Worksheets("RANG_BUOC")
    Set view = book.Worksheets("LOOKAHEAD_21N"): Set dash = book.Worksheets("DIEU_HANH")
    Set ids = CreateObject("Scripting.Dictionary"): ids.CompareMode = vbTextCompare
    For r = FIRST_ROW To LastTask(tasks)
        id = Trim(CStr(tasks.Cells(r, 1).Value))
        If id <> "" Then
            If ids.Exists(id) Then Err.Raise 5, , "Activity ID trùng: " & id
            ids.Add id, r
        End If
    Next r
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    categories = Array("BAN_VE", "VAT_TU", "MAT_BANG", "QA_QC", "HSE", "NHAN_LUC")
    For r = 6 To ready.Cells(ready.Rows.Count, 1).End(xlUp).Row
        id = Trim(CStr(ready.Cells(r, 1).Value))
        If id <> "" Then
            key = id & "|" & UCase(Trim(CStr(ready.Cells(r, 2).Value)))
            If seen.Exists(key) Then
                seen(key) = False
                dataIssues = dataIssues + 1
            Else
                seen.Add key, UCase(Trim(CStr(ready.Cells(r, 6).Value))) = "PASS" And ValidConfirmation(ready.Cells(r, 3).Value, ready.Cells(r, 4).Value, ready.Cells(r, 5).Value, reportDay)
            End If
            If Not ids.Exists(id) Then dataIssues = dataIssues + 1
        End If
    Next r
    constraints.Range("J6:J2005").ClearContents
    Dim constraintIds As Object
    Set constraintIds = CreateObject("Scripting.Dictionary"): constraintIds.CompareMode = vbTextCompare
    For r = 6 To Application.Max(constraints.Cells(constraints.Rows.Count, 1).End(xlUp).Row, constraints.Cells(constraints.Rows.Count, 2).End(xlUp).Row)
        id = Trim(CStr(constraints.Cells(r, 1).Value))
        If id <> "" Then
            If constraintIds.Exists(id) Then dataIssues = dataIssues + 1
            constraintIds(id) = True
            If Not ids.Exists(Trim(CStr(constraints.Cells(r, 2).Value))) Then
                constraints.Cells(r, 10) = "ID KHONG KHOP": dataIssues = dataIssues + 1
            End If
        ElseIf Len(Trim(CStr(constraints.Cells(r, 2).Value))) > 0 Then
            constraints.Cells(r, 10) = "THIEU CONSTRAINT ID": dataIssues = dataIssues + 1
        End If
    Next r
    view.Range("A6:L2005").ClearContents
    outputRow = 6
    For r = FIRST_ROW To LastTask(tasks)
        If UCase(CStr(tasks.Cells(r, 4).Value)) <> "HM" And Val(tasks.Cells(r, 10).Value) < 100 Then
            If Not IsDate(tasks.Cells(r, 6).Value) Or Not IsDate(tasks.Cells(r, 7).Value) Then Err.Raise 5, , "Hãy cập nhật tiến độ trước khi lập Lookahead."
            If CDate(tasks.Cells(r, 6).Value) <= reportDay + 20 Then
                id = Trim(CStr(tasks.Cells(r, 1).Value)): reason = "": openCount = 0: overdue = 0
                For EachCategoryIndex = 0 To 5
                    key = id & "|" & categories(EachCategoryIndex)
                    If Not seen.Exists(key) Then
                        reason = reason & categories(EachCategoryIndex) & "; "
                    ElseIf Not seen(key) Then
                        reason = reason & categories(EachCategoryIndex) & "; "
                    End If
                Next EachCategoryIndex
                For j = 6 To constraints.Cells(constraints.Rows.Count, 2).End(xlUp).Row
                    If StrComp(Trim(CStr(constraints.Cells(j, 2).Value)), id, vbTextCompare) = 0 Then
                        closed = UCase(Trim(CStr(constraints.Cells(j, 6).Value))) = "CLOSED" And Len(Trim(CStr(constraints.Cells(j, 1).Value))) > 0 And ValidConfirmation(constraints.Cells(j, 8).Value, constraints.Cells(j, 9).Value, constraints.Cells(j, 7).Value, reportDay)
                        If Not closed Then
                            openCount = openCount + 1
                            If IsDate(constraints.Cells(j, 5).Value) Then
                                If CDate(constraints.Cells(j, 5).Value) < reportDay Then overdue = overdue + 1
                            End If
                        End If
                    End If
                Next j
                If openCount > 0 Then reason = reason & "Ràng buộc chưa đóng hợp lệ; "
                If Len(Trim(CStr(tasks.Cells(r, 14).Value))) = 0 Then reason = reason & "Thiếu phụ trách; "
                If Len(Trim(CStr(tasks.Cells(r, 15).Value))) = 0 Then reason = reason & "Thiếu zone; "
                If dataIssues > 0 Then reason = reason & "Lỗi liên kết / dữ liệu trùng; "
                view.Cells(outputRow, 1) = id: view.Cells(outputRow, 2) = tasks.Cells(r, 2).Value
                view.Cells(outputRow, 3) = tasks.Cells(r, 3).Value: view.Cells(outputRow, 4) = tasks.Cells(r, 15).Value
                view.Cells(outputRow, 5) = tasks.Cells(r, 14).Value
                view.Cells(outputRow, 6) = tasks.Cells(r, 6).Value: view.Cells(outputRow, 7) = tasks.Cells(r, 7).Value
                view.Cells(outputRow, 8) = tasks.Cells(r, 10).Value
                view.Cells(outputRow, 10) = reason: view.Cells(outputRow, 11) = openCount: view.Cells(outputRow, 12) = overdue
                If reason = "" Then
                    view.Cells(outputRow, 9) = "READY": readyCount = readyCount + 1
                    view.Cells(outputRow, 9).Interior.Color = RGB(192, 235, 211)
                Else
                    view.Cells(outputRow, 9) = "HOLD": holdCount = holdCount + 1
                    view.Cells(outputRow, 9).Interior.Color = RGB(255, 215, 207)
                End If
                outputRow = outputRow + 1
            End If
        End If
    Next r
    view.Range("F6:G2005").NumberFormat = "dd/mm/yyyy"
    view.Range("H6:H2005").NumberFormat = "0.0" & Chr(34) & "%" & Chr(34)
    dash.Range("A6") = "Ngày báo cáo": dash.Range("B6") = reportDay: dash.Range("B6").NumberFormat = "dd/mm/yyyy"
    dash.Range("A7") = "Kết thúc cửa sổ 21 ngày": dash.Range("B7") = reportDay + 20: dash.Range("B7").NumberFormat = "dd/mm/yyyy"
    dash.Range("A8") = "Công việc READY": dash.Range("B8") = readyCount
    dash.Range("A9") = "Công việc HOLD": dash.Range("B9") = holdCount
    dash.Range("A10") = "Lỗi ID / dữ liệu trùng": dash.Range("B10") = dataIssues
    dash.Range("A12") = "READY là xác nhận điều kiện thi công; chưa phải lệnh huy động hoặc phê duyệt chi tiền."
    dash.Range("A13") = "Thiếu dữ liệu tài chính không được hiểu là dòng tiền bằng 0."
End Sub

Public Sub RibbonLookahead(ByVal control As Object)
    On Error GoTo Failed
    RefreshLookahead ActiveWorkbook
    ActiveWorkbook.Worksheets("LOOKAHEAD_21N").Activate
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RibbonConstraints(ByVal control As Object)
    On Error GoTo Failed
    EnsureLookahead ActiveWorkbook: ActiveWorkbook.Worksheets("RANG_BUOC").Activate
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub
Public Sub RibbonReadiness(ByVal control As Object)
    On Error GoTo Failed
    PrepareReadiness ActiveWorkbook
    ActiveWorkbook.Worksheets("SAN_SANG").Activate
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Atlas"
End Sub

Public Sub PrepareReadiness(ByVal book As Workbook)
    Dim w As Worksheet, s As Worksheet, existing As Object, cats As Variant, r As Long, k As Long, dest As Long, id As String, key As String
    EnsureLookahead book
    Set w = ProjectSheet(book): Set s = book.Worksheets("SAN_SANG")
    Set existing = CreateObject("Scripting.Dictionary"): existing.CompareMode = vbTextCompare
    dest = Application.Max(6, s.Cells(s.Rows.Count, 1).End(xlUp).Row + 1)
    For r = 6 To dest - 1
        existing(Trim(CStr(s.Cells(r, 1).Value)) & "|" & UCase(Trim(CStr(s.Cells(r, 2).Value)))) = True
    Next r
    cats = Array("BAN_VE", "VAT_TU", "MAT_BANG", "QA_QC", "HSE", "NHAN_LUC")
    For r = FIRST_ROW To LastTask(w)
        If UCase(CStr(w.Cells(r, 4).Value)) <> "HM" And Val(w.Cells(r, 10).Value) < 100 Then
            id = Trim(CStr(w.Cells(r, 1).Value))
            If id <> "" Then
                For k = 0 To 5
                    key = id & "|" & cats(k)
                    If Not existing.Exists(key) Then
                        If dest > 12005 Then Err.Raise 5, , "Sổ xác nhận vượt 12000 dòng."
                        s.Cells(dest, 1) = id: s.Cells(dest, 2) = cats(k): s.Cells(dest, 6) = "HOLD"
                        existing(key) = True: dest = dest + 1
                    End If
                Next k
            End If
        End If
    Next r
End Sub
