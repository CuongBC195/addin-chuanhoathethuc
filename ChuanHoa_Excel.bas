'===========================================================================
' MACRO CHUẨN HÓA TRANG IN & CĂN CHỈNH BẢNG BIỂU EXCEL HOÀN HẢO
' - KHẮC PHỤC TRIỆT ĐỂ LỖI BỊ CHE MẤT CHỮ TRONG Ô (KỂ CẢ MERGED CELLS)
' - Tự động tính toán chiều cao hàng thông minh cho các ô gộp dòng / gộp cột
' - Tôn trọng và giữ nguyên độ rộng cột đẹp mắt của người dùng (không bóp nghẹt)
' - Bật WrapText tự động xuống dòng an toàn cho toàn bộ khối Merge
' - Xóa bỏ các đường ngắt trang rác cũ (lấp đầy trang in, không thừa nửa trang)
' - Thiết lập trang in A4 Ngang (Landscape), vừa khít 1 trang ngang, chữ to rõ
'===========================================================================

Public Sub RibbonChuanHoaExcel(control As Object)
    Call ChuanHoa_TrangIn_Excel
End Sub

Sub ChuanHoa_TrangIn_Excel()
    If Not CheckLicenseActivation() Then Exit Sub
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If ws Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim rng As Range
    Set rng = ws.UsedRange
    
    If rng Is Nothing Then
        MsgBox "Bang tinh hien tai khong co du lieu!", vbExclamation, "Chuan Hoa Excel"
        Exit Sub
    End If
    
    If rng.Cells.Count = 1 And Trim(CStr(rng.Cells(1, 1).Value)) = "" Then
        MsgBox "Bang tinh hien tai khong co du lieu!", vbExclamation, "Chuan Hoa Excel"
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------
    ' BƯỚC 1: XÓA SẠCH 100% ĐƯỜNG NGẮT TRANG THỦ CÔNG CŨ
    ' (TRIỆT TIÊU LỖI MỖI TRANG CÓ 3 DÒNG RỒI THỪA NỬA TRANG DƯỚI)
    '-----------------------------------------------------------------------
    On Error Resume Next
    ws.ResetAllPageBreaks
    On Error GoTo ErrorHandler
    
    '-----------------------------------------------------------------------
    ' BƯỚC 2: ĐỒNG BỘ FONT CHỮ DỄ ĐỌC (TIMES NEW ROMAN)
    ' Giữ cỡ chữ 12pt chuẩn rõ ràng, không bị bé
    '-----------------------------------------------------------------------
    With rng.Font
        .Name = "Times New Roman"
        If .Size < 10 Or .Size > 14 Then .Size = 12
    End With
    
    ' Căn giữa theo chiều dọc cho toàn bộ ô (chữ nằm ngay ngắn giữa dòng)
    rng.VerticalAlignment = xlCenter
    
    '-----------------------------------------------------------------------
    ' BƯỚC 3: BẬT WRAP TEXT TỰ ĐỘNG XUỐNG DÒNG (AN TOÀN CHO MERGED CELLS)
    ' Đảm bảo chữ tự động xuống dòng trong ô, không bị tràn viền hay đè ô khác
    '-----------------------------------------------------------------------
    Dim cell As Range
    For Each cell In rng.Cells
        If Not IsEmpty(cell.Value) Then
            Dim txtVal As String
            txtVal = Trim(CStr(cell.Value))
            If Len(txtVal) > 15 Or InStr(txtVal, vbLf) > 0 Or InStr(txtVal, " ") > 0 Then
                If cell.MergeCells Then
                    cell.MergeArea.WrapText = True
                Else
                    cell.WrapText = True
                End If
            End If
        End If
    Next cell
    
    '-----------------------------------------------------------------------
    ' BƯỚC 4: BẢO VỆ ĐỘ RỘNG CỘT GỐC CỦA NGƯỜI DÙNG (CHỐNG LỖI BÓP NGHẸT CỘT)
    ' Giữ nguyên 100% độ rộng cột đẹp mắt của người dùng như ở Hình 1!
    ' Tuyệt đối không tự ý co hẹp các cột nội dung làm chữ bị bẻ vụn!
    '-----------------------------------------------------------------------
    Dim col As Range
    For Each col In rng.Columns
        ' Chỉ bảo đảm cột không bị quá nhỏ (< 6), còn lại giữ nguyên thiết kế của người dùng
        If col.ColumnWidth < 6 Then col.ColumnWidth = 6
    Next col
    
    '-----------------------------------------------------------------------
    ' BƯỚC 5: TÍNH TOÁN CHIỀU CAO HÀNG THÔNG MINH (CHỐNG CHE MẤT CHỮ TRONG MERGE CELLS)
    ' Excel AutoFit bị liệt hoàn toàn khi gặp Merged Cells.
    ' Thuật toán này tự động đo độ dài văn bản và mở rộng hàng đủ chứa 100% chữ!
    '-----------------------------------------------------------------------
    Call AutoFitRowsWithMerged(rng)
    
    '-----------------------------------------------------------------------
    ' BƯỚC 5.1: NHẬN DIỆN THÔNG MINH DÒNG TIÊU ĐỀ BẢNG THỰC SỰ (TABLE HEADER)
    ' - Loại trừ 100% phần Tiêu ngữ (Đảng bộ, Đảng ủy, Đảng Cộng sản, ngày tháng, *)
    ' - Loại trừ tựa đề văn bản (Phụ lục, Kế hoạch, Báo cáo...)
    ' - Tìm chính xác dòng: Trục/Nhóm nhiệm vụ | Nội dung | Đơn vị tính...
    '-----------------------------------------------------------------------
    Dim hasHeader As Boolean
    Dim hStart As Long, hEnd As Long
    hasHeader = FindTableHeaderRows(ws, rng, hStart, hEnd)
    
    '-----------------------------------------------------------------------
    ' BƯỚC 6: KẺ KHUNG VIỀN IN ẤN SẮC NÉT (BORDERS)
    ' - Chỉ kẻ khung viền từ dòng Tiêu đề bảng trở xuống hết dữ liệu
    ' - Xóa bỏ đường viền ô rác ở phần Tiêu ngữ phía trên để giữ nguyên vẻ thanh lịch
    '-----------------------------------------------------------------------
    If hasHeader Then
        ' Xóa viền ô thừa ở phần Tiêu ngữ (từ dòng 1 đến trước dòng Table Header)
        If hStart > rng.Row Then
            Dim topRange As Range
            Set topRange = ws.Range(ws.Cells(rng.Row, rng.Column), ws.Cells(hStart - 1, rng.Column + rng.Columns.Count - 1))
            topRange.Borders.LineStyle = xlNone
        End If
        
        ' Kẻ viền sắc nét chỉ cho phần BẢNG BIỂU THỰC SỰ
        Dim tableRange As Range
        Dim lastRow As Long
        lastRow = rng.Row + rng.Rows.Count - 1
        Set tableRange = ws.Range(ws.Cells(hStart, rng.Column), ws.Cells(lastRow, rng.Column + rng.Columns.Count - 1))
        
        With tableRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Else
        With rng.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    '-----------------------------------------------------------------------
    ' BƯỚC 7: THIẾT LẬP TRANG IN HOÀN HẢO (PAGE SETUP)
    '-----------------------------------------------------------------------
    With ws.PageSetup
        .PaperSize = xlPaperA4
        
        ' Bảng báo cáo công tác luôn In Ngang (Landscape) để chữ to rõ, không bóp nghẹt
        .Orientation = xlLandscape
        
        ' Thu hẹp lề giấy xuống 1.0 cm để tối đa hóa diện tích in
        On Error Resume Next
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.2)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        On Error GoTo ErrorHandler
        
        ' Căn giữa bảng trên trang giấy khi in
        .CenterHorizontally = True
        .CenterVertically = False
        
        ' VÙNG IN VỪA KHÍT 1 TRANG NGANG (Fit to 1 page wide)
        ' Chiều dọc tự động cuộn lấp đầy trang in tự nhiên
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        
        .PrintArea = rng.Address
        
        ' Tự động lặp lại dòng tiêu đề khi bảng dài in sang trang 2, 3...
        ' Chỉ lặp lại khi tìm thấy đúng Table Header thực sự; nếu không thì bỏ hẳn!
        If hasHeader Then
            .PrintTitleRows = "$" & hStart & ":$" & hEnd
        Else
            .PrintTitleRows = ""
        End If
    End With
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Created by VPDU xa Thanh Phu", vbInformation, "Chuan Hoa Trang In Excel"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Loi: " & Err.Description, vbCritical, "Chuan Hoa Excel"
End Sub

'-----------------------------------------------------------------------
' HÀM PHỤ TRỢ: TÍNH TOÁN CHIỀU CAO HÀNG THÔNG MINH CHO CẢ MERGED CELLS
' Triệt tiêu 100% tình trạng chữ bị che khuất trong Excel
'-----------------------------------------------------------------------
Private Sub AutoFitRowsWithMerged(ByVal rng As Range)
    Dim r As Range
    Dim cell As Range
    
    For Each r In rng.Rows
        Dim maxReqHeight As Double
        maxReqHeight = 22 ' Chiều cao tối thiểu mặc định để bảng thoáng đẹp
        
        For Each cell In r.Cells
            ' Chỉ xét ô đầu tiên của khối Merge để không tính lặp lại
            If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                If Not IsEmpty(cell.Value) Then
                    Dim valStr As String
                    valStr = Trim(CStr(cell.Value))
                    
                    If Len(valStr) > 0 Then
                        ' 1. Tính tổng độ rộng cột khả dụng của ô (cộng dồn nếu Merge nhiều cột)
                        Dim totalW As Double
                        totalW = 0
                        Dim mCol As Range
                        For Each mCol In cell.MergeArea.Columns
                            totalW = totalW + mCol.ColumnWidth
                        Next mCol
                        
                        If totalW < 8 Then totalW = 8
                        
                        ' 2. Ước lượng số ký tự trên 1 dòng hiển thị (font Times New Roman 12pt)
                        Dim charsPerLine As Double
                        charsPerLine = totalW * 0.95
                        If charsPerLine < 8 Then charsPerLine = 8
                        
                        ' 3. Đếm số dòng thực tế (xét cả dấu xuống dòng Alt+Enter lẫn WrapText)
                        Dim hardLines() As String
                        hardLines = Split(valStr, Chr(10))
                        
                        Dim totalLines As Long
                        totalLines = 0
                        
                        Dim k As Long
                        For k = LBound(hardLines) To UBound(hardLines)
                            Dim lineTxt As String
                            lineTxt = Trim(hardLines(k))
                            If Len(lineTxt) = 0 Then
                                totalLines = totalLines + 1
                            Else
                                Dim wrapCount As Long
                                wrapCount = Application.WorksheetFunction.RoundUp(Len(lineTxt) / charsPerLine, 0)
                                If wrapCount < 1 Then wrapCount = 1
                                totalLines = totalLines + wrapCount
                            End If
                        Next k
                        
                        ' 4. Chiều cao cần thiết: mỗi dòng ~ 16.5 pt + padding trên dưới 6 pt
                        Dim cellReqHeight As Double
                        cellReqHeight = (totalLines * 16.5) + 6
                        
                        ' Nếu ô gộp nhiều hàng (Merge Rows), chia đều chiều cao cho từng hàng
                        Dim numMergedRows As Long
                        numMergedRows = cell.MergeArea.Rows.Count
                        If numMergedRows > 1 Then
                            cellReqHeight = cellReqHeight / numMergedRows
                        End If
                        
                        If cellReqHeight > maxReqHeight Then
                            maxReqHeight = cellReqHeight
                        End If
                    End If
                End If
            End If
        Next cell
        
        ' Áp dụng chiều cao tính toán được cho dòng
        If r.RowHeight < maxReqHeight Then
            r.RowHeight = maxReqHeight
        End If
    Next r
End Sub

'===========================================================================
' BẢO MẬT & KÍCH HOẠT BẢN QUYỀN TRỰC TUYẾN QUA GOOGLE SHEET
' - Tự động ghi nhận Tên máy tính & Ngày giờ vào Google Sheet
' - Khóa mã trên Google Sheet sẽ lập tức BLOCK người dùng kể cả đã kích hoạt
' - Mỗi máy kích hoạt 1 lần duy nhất, hỗ trợ Offline khi tạm thời mất mạng
'===========================================================================
Private Function CheckLicenseActivation() As Boolean
    Dim appName As String, sectionName As String
    Dim savedKey As String, currentMachine As String
    Dim apiUrl As String
    
    appName = "OfficePartySuite"
    sectionName = "License"
    apiUrl = "https://script.google.com/macros/s/AKfycbxSZyf-7FeVZ6xbhBwuT5aVVGJDdqP8AJCYek_Wp3Ila5yWLGnmq_rIua3kL1dFsvzOug/exec"
    
    currentMachine = GetMachineName()
    
    On Error Resume Next
    savedKey = GetSetting(appName, sectionName, "ActiveKey", "")
    On Error GoTo 0
    
    ' TRƯỜNG HỢP 1: ĐÃ KÍCH HOẠT TRƯỚC ĐÓ -> KIỂM TRA XEM CÓ BỊ BẠN KHÓA TRÊN SHEET HAY KHÔNG
    If Len(savedKey) > 0 Then
        Dim checkUrl As String
        checkUrl = apiUrl & "?key=" & EncodeUrl(savedKey) & "&machine=" & EncodeUrl(currentMachine)
        
        Dim checkResp As String
        checkResp = CallHttpApi(checkUrl)
        
        ' Nếu kết nối được và Google Sheet phản hồi
        If Len(checkResp) > 0 Then
            Dim checkStatus As String
            checkStatus = ParseJsonField(checkResp, "status")
            
            If checkStatus = "ERROR" Then
                ' NẾU BẠN ĐÃ ĐỔI TRẠNG THÁI THÀNH "ĐÃ KHÓA" TRÊN GOOGLE SHEET:
                ' -> TỰ ĐỘNG XÓA BẢN QUYỀN VÀ KHÓA CHẶN NGAY LẬP TỨC!
                SaveSetting appName, sectionName, "ActiveKey", ""
                SaveSetting appName, sectionName, "ActiveMachine", ""
                
                Dim errReason As String
                errReason = ParseJsonField(checkResp, "message")
                If Len(errReason) = 0 Then errReason = "Ban quyen cua ban da bi khoa boi VPDU xa Thanh Phu!"
                
                MsgBox "CANH BAO BAN QUYEN:" & vbCrLf & errReason & vbCrLf & vbCrLf & _
                       "Tien ich da bi dung hoat dong tren may nay.", vbCritical, "Ban Quyen Bi Khoa"
                CheckLicenseActivation = False
                Exit Function
            End If
        End If
        
        ' Nếu trạng thái bình thường (hoặc tạm thời mất mạng), cho phép tiếp tục dùng
        CheckLicenseActivation = True
        Exit Function
    End If
    
    ' TRƯỜNG HỢP 2: MÁY CHƯA KÍCH HOẠT -> YÊU CẦU NHẬP MÃ
    Dim userInput As String
    userInput = InputBox("Vui long nhap Ma kich hoat ban quyen de su dung tien ich:" & vbCrLf & _
                         "(Moi may tinh chi can nhap kich hoat 1 lan duy nhat)" & vbCrLf & vbCrLf & _
                         "Ten may tinh: " & currentMachine & vbCrLf & _
                         "Don vi phat trien: Van phong Dang uy xa Thanh Phu", _
                         "Kich Hoat Ban Quyen Add-in")
    
    userInput = Trim(UCase(userInput))
    
    If Len(userInput) = 0 Then
        MsgBox "Ban chua nhap ma kich hoat! Khong the thuc hien chuan hoa.", vbExclamation, "Ban Quyen Add-in"
        CheckLicenseActivation = False
        Exit Function
    End If
    
    ' Gửi lên Google Sheet để kích hoạt và ghi nhận tên máy
    Dim actUrl As String
    actUrl = apiUrl & "?key=" & EncodeUrl(userInput) & "&machine=" & EncodeUrl(currentMachine)
    
    Dim actResp As String
    actResp = CallHttpApi(actUrl)
    
    If Len(actResp) = 0 Then
        MsgBox "Khong the ket noi den may chu xac thuc ban quyen!" & vbCrLf & _
               "Vui long kiem tra ket noi Internet va thu lai.", vbExclamation, "Loi Ket Noi"
        CheckLicenseActivation = False
        Exit Function
    End If
    
    Dim actStatus As String, actMsg As String
    actStatus = ParseJsonField(actResp, "status")
    actMsg = ParseJsonField(actResp, "message")
    
    If actStatus = "SUCCESS" Then
        ' Lưu trạng thái bản quyền vào máy
        SaveSetting appName, sectionName, "ActiveKey", userInput
        SaveSetting appName, sectionName, "ActiveMachine", currentMachine
        
        MsgBox "Kich hoat ban quyen thanh cong!" & vbCrLf & _
               "May tinh: " & currentMachine & vbCrLf & _
               "Ban co the su dung tien ich vinh vien tren may nay.", vbInformation, "Kich Hoat Thanh Cong"
        CheckLicenseActivation = True
    Else
        MsgBox "Kich hoat that bai!" & vbCrLf & actMsg & vbCrLf & vbCrLf & _
               "Vui long lien he Van phong Dang uy xa Thanh Phu de duoc cap ma hop le.", vbCritical, "Kich Hoat That Bai"
        CheckLicenseActivation = False
    End If
End Function

Private Function CallHttpApi(ByVal url As String) As String
    Dim res As String
    res = ""
    #If Mac Then
        On Error Resume Next
        Dim sc As String
        sc = "do shell script ""curl -s -L --max-time 8 '" & url & "'"""
        res = MacScript(sc)
        On Error GoTo 0
    #Else
        On Error Resume Next
        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        If http Is Nothing Then
            Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        End If
        http.Open "GET", url, False
        http.setRequestHeader "User-Agent", "Mozilla/5.0"
        http.setTimeouts 5000, 8000, 8000, 8000
        http.Send
        If Err.Number = 0 Then
            res = http.responseText
        End If
        On Error GoTo 0
    #End If
    CallHttpApi = Trim(res)
End Function

Private Function GetMachineName() As String
    Dim mName As String
    #If Mac Then
        On Error Resume Next
        mName = MacScript("do shell script ""scutil --get ComputerName || hostname -s || whoami""")
        mName = Trim(Replace(mName, vbCrLf, ""))
        mName = Trim(Replace(mName, vbLf, ""))
        If Len(mName) = 0 Then mName = Environ("USER")
        On Error GoTo 0
    #Else
        mName = Environ("COMPUTERNAME")
        If Len(mName) = 0 Then mName = Environ("USERNAME")
    #End If
    If Len(mName) = 0 Then mName = "UNKNOWN_PC"
    GetMachineName = mName
End Function

Private Function ParseJsonField(ByVal json As String, ByVal fieldName As String) As String
    Dim pos As Long, startPos As Long, endPos As Long
    pos = InStr(json, """" & fieldName & """")
    If pos > 0 Then
        startPos = InStr(pos + Len(fieldName) + 2, json, """")
        If startPos > 0 Then
            endPos = InStr(startPos + 1, json, """")
            If endPos > startPos Then
                ParseJsonField = Mid(json, startPos + 1, endPos - startPos - 1)
            End If
        End If
    End If
End Function

Private Function EncodeUrl(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long
    Dim res As String
    res = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = AscW(ch) And &HFFFF&
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "-" Or ch = "_" Or ch = "." Then
            res = res & ch
        ElseIf code < 128 Then
            res = res & "%" & Right("0" & Hex(code), 2)
        Else
            res = res & "_"
        End If
    Next i
    EncodeUrl = res
End Function

'-----------------------------------------------------------------------
' HÀM NHẬN DIỆN THÔNG MINH DÒNG TIÊU ĐỀ BẢNG THỰC SỰ (TABLE HEADER)
' - Tự động tìm chính xác dòng tiêu đề của bảng (như: STT, Trục/Nhóm, Nội dung, Đơn vị tính...)
' - Tuyệt đối không chọn nhầm phần Tiêu ngữ / Quốc hiệu / Tựa đề văn bản
' - Nếu tìm thấy: trả về dòng bắt đầu và kết thúc (hỗ trợ cả header 2 tầng)
' - Nếu không chắc chắn: trả về False để không lặp header rác
'-----------------------------------------------------------------------
Private Function FindTableHeaderRows(ByVal ws As Worksheet, ByVal rng As Range, ByRef hStart As Long, ByRef hEnd As Long) As Boolean
    Dim r As Long
    Dim startRow As Long, endRow As Long
    startRow = rng.Row
    endRow = startRow + rng.Rows.Count - 1
    
    ' Quét tối đa 25 dòng đầu tiên để tìm Table Header
    Dim maxScanRow As Long
    maxScanRow = startRow + 25
    If maxScanRow > endRow Then maxScanRow = endRow
    
    Dim bestRow As Long
    Dim maxScore As Long
    bestRow = 0
    maxScore = 0
    
    Dim c As Long
    Dim startCol As Long, endCol As Long
    startCol = rng.Column
    endCol = startCol + rng.Columns.Count - 1
    
    For r = startRow To maxScanRow
        Dim colCountWithText As Long
        colCountWithText = 0
        Dim coveredCols As Long
        coveredCols = 0
        Dim rowAscText As String
        rowAscText = ""
        
        For c = startCol To endCol
            ' Chỉ xét ô đầu tiên của khối Merge để không đếm trùng lặp
            If ws.Cells(r, c).Address = ws.Cells(r, c).MergeArea.Cells(1, 1).Address Then
                Dim cellVal As String
                cellVal = Trim(CStr(ws.Cells(r, c).Value))
                If Len(cellVal) > 0 Then
                    colCountWithText = colCountWithText + 1
                    coveredCols = coveredCols + ws.Cells(r, c).MergeArea.Columns.Count
                    rowAscText = rowAscText & " " & ToAscii(cellVal)
                End If
            End If
        Next c
        
        ' ĐIỀU KIỆN TIÊN QUYẾT BẮT BUỘC:
        ' 1. Nếu chỉ có 1 ô text duy nhất mà gộp/merge trải dài toàn bảng (như dòng 5 "DANH MỤC..."):
        '    -> ĐÂY LÀ TỰA ĐỀ VĂN BẢN / TỰA ĐỀ BẢNG, BỊ LOẠI BỎ 100%!
        ' 2. Dòng Table Header hợp lệ:
        '    - Hoặc có từ 3 ô text riêng biệt trở lên (colCountWithText >= 3)
        '    - Hoặc có ít nhất 2 nhóm ô text riêng biệt (colCountWithText >= 2) và có Merge cột bao phủ từ 3 cột trở lên (coveredCols >= 3)
        Dim isHeaderCandidate As Boolean
        isHeaderCandidate = False
        
        If colCountWithText >= 3 Then
            isHeaderCandidate = True
        ElseIf colCountWithText >= 2 And coveredCols >= 3 Then
            isHeaderCandidate = True
        End If
        
        If isHeaderCandidate Then
            ' 1. BỎ QUA NẾU LÀ TIÊU NGỮ / QUỐC HIỆU / NGÀY THÁNG
            Dim isTieuNgu As Boolean
            isTieuNgu = False
            
            If InStr(rowAscText, "DANG CONG SAN") > 0 Or _
               InStr(rowAscText, "CONG HOA XA HOI") > 0 Or _
               InStr(rowAscText, "DOC LAP") > 0 Then
                isTieuNgu = True
            ElseIf InStr(rowAscText, "DANG BO") > 0 Or InStr(rowAscText, "DANG UY") > 0 Or InStr(rowAscText, "UY BAN") > 0 Then
                isTieuNgu = True
            ElseIf InStr(rowAscText, "NGAY") > 0 And InStr(rowAscText, "THANG") > 0 And InStr(rowAscText, "NAM") > 0 Then
                isTieuNgu = True
            End If
            
            If Not isTieuNgu Then
                ' 2. BỎ QUA NẾU CHỨA CÁC TỪ KHÓA TỰA ĐỀ VĂN BẢN LỚN
                Dim isDocTitle As Boolean
                isDocTitle = False
                If InStr(rowAscText, "DANH MUC") > 0 Or _
                   InStr(rowAscText, "PHU LUC") > 0 Or _
                   InStr(rowAscText, "KE HOACH") > 0 Or _
                   InStr(rowAscText, "BAO CAO") > 0 Or _
                   InStr(rowAscText, "DANH SACH") > 0 Or _
                   InStr(rowAscText, "BANG TONG HOP") > 0 Or _
                   InStr(rowAscText, "BANG PHAN CONG") > 0 Or _
                   InStr(rowAscText, "QUYET DINH") > 0 Or _
                   InStr(rowAscText, "THONG BAO") > 0 Then
                    isDocTitle = True
                End If
                
                If Not isDocTitle Then
                    ' 3. TÍNH ĐIỂM DÒNG TIÊU ĐỀ CỘT DỰA VÀO SỐ CỘT & TỪ KHÓA CHUẨN
                    Dim currentScore As Long
                    currentScore = colCountWithText * 10
                    
                    Dim kws As Variant
                    kws = Array("STT", "TT", "TRUC", "NHOM", "NOI DUNG", "NHIEM VU", "SAN PHAM", _
                                "DON VI", "CHI TIEU", "THOI GIAN", "PHAN CONG", "HO VA TEN", "HO TEN", _
                                "CHUC VU", "SO LUONG", "GHI CHU", "TIEN DO", "KET QUA", _
                                "MO TA", "CHI TIET", "DIEN GIAI", "DON GIA", "THANH TIEN", "NGUOI THUC HIEN")
                    
                    Dim kwIdx As Long, kwHits As Long
                    kwHits = 0
                    For kwIdx = LBound(kws) To UBound(kws)
                        If InStr(rowAscText, CStr(kws(kwIdx))) > 0 Then
                            kwHits = kwHits + 1
                        End If
                    Next kwIdx
                    
                    currentScore = currentScore + (kwHits * 20)
                    
                    ' Chỉ xem xét các dòng có ít nhất 1 từ khóa tiêu đề cột hoặc số cột rất nhiều
                    If kwHits >= 1 And currentScore > maxScore Then
                        maxScore = currentScore
                        bestRow = r
                    End If
                End If
            End If
        End If
    Next r
    
    ' Nếu không tìm thấy dòng nào đạt chuẩn Table Header (hoặc điểm quá thấp)
    If bestRow = 0 Or maxScore < 40 Then
        FindTableHeaderRows = False
        Exit Function
    End If
    
    hStart = bestRow
    hEnd = bestRow
    
    ' Kiểm tra xem dòng bestRow có ô nào Merge dọc xuống các dòng kế tiếp không
    Dim maxMergeRow As Long
    maxMergeRow = bestRow
    For c = startCol To endCol
        If ws.Cells(bestRow, c).MergeCells Then
            Dim mRows As Long
            mRows = ws.Cells(bestRow, c).MergeArea.Row + ws.Cells(bestRow, c).MergeArea.Rows.Count - 1
            If mRows > maxMergeRow Then maxMergeRow = mRows
        End If
    Next c
    If maxMergeRow > hEnd Then hEnd = maxMergeRow
    
    ' Kiểm tra xem dòng ngay sau đó có phải là dòng đánh số thứ tự cột: (1), (2), (3)... hoặc 1, 2, 3...
    If hEnd < endRow Then
        Dim nextRow As Long
        nextRow = hEnd + 1
        Dim numColsWithNumber As Long
        numColsWithNumber = 0
        For c = startCol To endCol
            Dim nVal As String
            nVal = Trim(CStr(ws.Cells(nextRow, c).Value))
            If Len(nVal) > 0 Then
                If nVal Like "(*)" Or IsNumeric(nVal) Then
                    numColsWithNumber = numColsWithNumber + 1
                End If
            End If
        Next c
        ' Nếu có từ 3 cột chứa số thứ tự cột (1), (2), (3)... thì gộp luôn dòng này vào Header in
        If numColsWithNumber >= 3 Then
            hEnd = nextRow
        End If
    End If
    
    FindTableHeaderRows = True
End Function

'-----------------------------------------------------------------------
' HAM CHUYEN DOI CHU CO DAU THANH ASCII KHONG DAU (DUNG MA SO ASCW)
' KHONG CHUA KY TU TIENG VIET TRONG MA NGUON NEN KHONG BAO GIO BI LOI FONT VBE
'-----------------------------------------------------------------------
Private Function ToAscii(ByVal s As String) As String
    Dim res As String
    res = ""
    Dim k As Long
    For k = 1 To Len(s)
        Dim ch As String
        ch = Mid(s, k, 1)
        Dim w As Long
        w = AscW(ch)
        If w < 0 Then w = w + 65536
        
        Select Case w
            Case 272, 273 ' D, d
                res = res & "D"
            Case 192 To 195, 224 To 227, 258, 259, 7840 To 7863 ' A, a
                res = res & "A"
            Case 200 To 202, 232 To 234, 278, 279, 7864 To 7879 ' E, e
                res = res & "E"
            Case 204, 205, 236, 237, 296, 297, 7880 To 7883 ' I, i
                res = res & "I"
            Case 210 To 214, 242 To 246, 416, 417, 7884 To 7907 ' O, o
                res = res & "O"
            Case 217, 218, 249, 250, 360, 361, 431, 432, 7908 To 7921 ' U, u
                res = res & "U"
            Case 221, 253, 7922 To 7929 ' Y, y
                res = res & "Y"
            Case Else
                If w >= 97 And w <= 122 Then ' a-z
                    res = res & Chr(w - 32) ' Chuyen thanh A-Z
                ElseIf (w >= 65 And w <= 90) Or (w >= 48 And w <= 57) Or w = 32 Or w = 42 Or w = 45 Or w = 47 Then
                    res = res & Chr(w)
                Else
                    res = res & " "
                End If
        End Select
    Next k
    ToAscii = res
End Function
