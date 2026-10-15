'===========================================================================
' MACRO CHUẨN HÓA VĂN BẢN ĐẢNG THEO HD 05-HD/VPTW (Cập nhật thực tế)
'===========================================================================

' Entry point khi bấm nút từ thanh công cụ Ribbon của Word
Public Sub RibbonChuanHoa(control As IRibbonControl)
    Call ChuanHoa_VanBanDang
End Sub

Sub ChuanHoa_VanBanDang()
    If Not CheckLicenseActivation() Then Exit Sub
    Application.ScreenUpdating = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    '-----------------------------------------------------------------------
    ' BƯỚC 0: DỌN RÁC (Shift+Enter, Dòng trắng thừa, Đánh số tự động)
    '-----------------------------------------------------------------------
    
    ' 0.1 Đóng băng toàn bộ Đánh số tự động (Auto Numbering) thành Text tĩnh
    ' Điều này giúp ép lề 1.27cm thành công 100% không bị Word chèn ép
    doc.ConvertNumbersToText
    
    ' 0.2 Đổi Shift+Enter thành Enter chuẩn
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With doc.Content.Find
        .Text = "^p "
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
    
    Dim i As Integer
    For i = 1 To 5
        With doc.Content.Find
            .Text = "^p "
            .Replacement.Text = "^p"
            .Execute Replace:=wdReplaceAll
        End With
    Next i
    
    With doc.Content.Find
        .Text = "^p^t"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
    
    For i = 1 To 3
        With doc.Content.Find
            .Text = "^p^t"
            .Replacement.Text = "^p"
            .Execute Replace:=wdReplaceAll
        End With
    Next i
    
    ' Xóa các dòng trắng thừa ngoài bảng (không xóa dòng trống trong bảng như ô Chữ ký)
    Dim pBlank As Paragraph
    Dim prevBlank As Boolean
    prevBlank = False
    For Each pBlank In doc.Paragraphs
        If Not pBlank.Range.Information(wdWithInTable) Then
            Dim sB As String
            sB = Trim(Replace(Replace(Replace(pBlank.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))
            If Len(sB) = 0 Then
                If prevBlank Then
                    On Error Resume Next
                    pBlank.Range.Delete
                    On Error GoTo 0
                Else
                    prevBlank = True
                End If
            Else
                prevBlank = False
            End If
        Else
            prevBlank = False
        End If
    Next pBlank
    
    ' Xóa nhiều dấu cách liên tiếp
    For i = 1 To 5
        With doc.Content.Find
            .Text = "  "
            .Replacement.Text = " "
            .Execute Replace:=wdReplaceAll
        End With
    Next i
    
    '-----------------------------------------------------------------------
    ' BƯỚC 1: CẤU HÌNH TRANG IN
    '-----------------------------------------------------------------------
    With doc.PageSetup
        .PaperSize = wdPaperA4
        .Orientation = wdOrientPortrait
        .TopMargin = MillimetersToPoints(20)
        .BottomMargin = MillimetersToPoints(20)
        .LeftMargin = MillimetersToPoints(30)
        .RightMargin = MillimetersToPoints(15)
        .HeaderDistance = MillimetersToPoints(10)
        .FooterDistance = MillimetersToPoints(10)
        .MirrorMargins = False
        .DifferentFirstPageHeaderFooter = True
    End With
    
    '-----------------------------------------------------------------------
    ' BƯỚC 2: SỐ TRANG
    '-----------------------------------------------------------------------
    Dim sec As Section
    For Each sec In doc.Sections
        Dim hdr As HeaderFooter
        Set hdr = sec.Headers(wdHeaderFooterPrimary)
        hdr.Range.Delete
        
        hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        hdr.Range.Font.Name = "Times New Roman"
        hdr.Range.Font.Size = 13.5
        hdr.Range.Font.Bold = False
        hdr.Range.Font.Italic = False
        hdr.Range.Font.ColorIndex = wdBlack
        
        hdr.Range.Fields.Add Range:=hdr.Range, Type:=wdFieldPage
        hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        sec.Headers(wdHeaderFooterFirstPage).Range.Delete
    Next sec
    
    '-----------------------------------------------------------------------
    ' BƯỚC 3: DÒ BẮT VÀ CHUẨN HÓA BẢNG HEADER (TIÊU NGỮ / CƠ QUAN)
    '-----------------------------------------------------------------------
    ' Tự động căn vừa cửa sổ trang in (AutoFit Window) cho toàn bộ bảng trong tài liệu
    Dim tFit As Table
    For Each tFit In doc.Tables
        On Error Resume Next
        tFit.AutoFitBehavior 2 ' wdAutoFitWindow = 2
        tFit.AllowAutoFit = True
        On Error GoTo 0
    Next tFit
    
    Dim tbl As Table
    Dim c As Cell
    Dim pCell As Paragraph
    Dim pDCS As Paragraph
    Dim cRight As Cell
    Dim foundHeaderTable As Boolean
    foundHeaderTable = False
    
    For Each tbl In doc.Tables
        Dim tAsc As String
        tAsc = ToAscii(tbl.Range.Text)
        
        ' Bảng Header: Luôn là bảng đầu tiên của tài liệu HOẶC chứa các từ khóa tiêu ngữ/cơ quan
        Dim isHeaderTbl As Boolean
        isHeaderTbl = False
        
        If tbl Is doc.Tables(1) Then
            isHeaderTbl = True
        ElseIf InStr(1, tAsc, "DANG CONG SAN") > 0 Or InStr(1, tAsc, "VIET NAM") > 0 Or InStr(1, tAsc, "DANG BO") > 0 Or InStr(1, tAsc, "DANG UY") > 0 Or InStr(1, tAsc, "CHI BO") > 0 Then
            isHeaderTbl = True
        End If
        
        If isHeaderTbl Then
            foundHeaderTable = True
            Set pDCS = Nothing
            Set cRight = Nothing
            
            ' Tắt viền bảng (Borders None) để chuẩn theo quy định bảng ẩn viền
            On Error Resume Next
            tbl.Borders.Enable = False
            On Error GoTo 0
            
            Dim totalCells As Integer
            totalCells = tbl.Range.Cells.Count
            
            Dim cellIdx As Integer
            cellIdx = 1
            
            Dim leftColCqIdx As Integer
            leftColCqIdx = 1
            Dim leftColInTrichYeu As Boolean
            leftColInTrichYeu = False
            
            For Each c In tbl.Range.Cells
                Dim cellTextRaw As String
                cellTextRaw = Trim(Replace(Replace(Replace(c.Range.Text, vbCr, ""), Chr(7), ""), vbLf, ""))
                
                ' Bỏ qua ô rỗng (ví dụ cột giữa hẹp dùng làm khoảng cách chia 2 bên)
                If Len(cellTextRaw) = 0 Then
                    GoTo NextCellHeader
                End If
                
                Dim cAsc As String
                cAsc = ToAscii(c.Range.Text)
                
                ' Bỏ qua ô Kính gửi và ô phòng ban nếu nằm trong bảng Header (để Bước 3.2 chuyên biệt xử lý)
                If InStr(1, cAsc, "KINH GUI") > 0 Or Left(cellTextRaw, 1) = "-" Then
                    GoTo NextCellHeader
                End If
                
                Dim isRightCell As Boolean
                isRightCell = False
                
                ' Nhận diện ô phải: Chứa tiêu ngữ (DANG CONG SAN / VIET NAM), hoặc ngày tháng, hoặc là ô cuối cùng
                If InStr(1, cAsc, "DANG CONG SAN") > 0 Or InStr(1, cAsc, "VIET NAM") > 0 Then
                    isRightCell = True
                    Set cRight = c
                ElseIf InStr(1, cAsc, "NGAY") > 0 And InStr(1, cAsc, "THANG") > 0 And InStr(1, cAsc, "NAM") > 0 Then
                    isRightCell = True
                    If cRight Is Nothing Then Set cRight = c
                ElseIf cellIdx = totalCells Then
                    isRightCell = True
                    If cRight Is Nothing Then Set cRight = c
                Else
                    On Error Resume Next
                    If c.ColumnIndex > 1 And tbl.Columns.Count >= 2 And c.ColumnIndex = tbl.Columns.Count Then
                        isRightCell = True
                    End If
                    On Error GoTo 0
                End If
                
                ' Reset chuẩn ô: xóa viền ẩn và màu nền
                On Error Resume Next
                c.Borders.Enable = False
                c.Shading.BackgroundPatternColor = wdColorAutomatic
                c.Shading.Texture = wdTextureNone
                On Error GoTo 0
                
                Dim lineIndex As Integer
                lineIndex = 1
                
                For Each pCell In c.Range.Paragraphs
                    ' Xóa khoảng trắng/tab thừa ở đầu dòng để căn lề và lấy tọa độ chính xác tuyệt đối
                    Do While pCell.Range.Characters.Count > 1 And (pCell.Range.Characters(1).Text = " " Or pCell.Range.Characters(1).Text = vbTab)
                        pCell.Range.Characters(1).Delete
                    Loop
                    
                    Dim pAsc As String
                    pAsc = ToAscii(pCell.Range.Text)
                    pAsc = Trim(pAsc)
                    
                    If Len(pAsc) = 0 Then GoTo SkipEmptyLineHeader
                    
                    With pCell.Format
                        .LeftIndent = 0
                        .RightIndent = 0
                        .FirstLineIndent = 0
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .DisableLineHeightGrid = True ' Tắt Snap to Grid
                    End With
                    
                    pCell.Range.Font.Name = "Times New Roman"
                    pCell.Range.Font.ColorIndex = wdBlack
                    pCell.Range.Font.Underline = wdUnderlineNone
                    
                    If isRightCell Then
                        ' --- CỘT BÊN PHẢI (TIÊU NGỮ & NGÀY THÁNG): CĂN SÁT PHẢI ---
                        pCell.Format.Alignment = wdAlignParagraphRight
                        
                        If InStr(1, pAsc, "DANG CONG SAN") > 0 Or InStr(1, pAsc, "VIET NAM") > 0 Then
                            ' Dòng ĐẢNG CỘNG SẢN VIỆT NAM -> In hoa, Đậm, cỡ 15, căn sát phải
                            pCell.Range.Case = wdUpperCase
                            pCell.Range.Font.Size = 15
                            pCell.Range.Font.Bold = True
                            pCell.Range.Font.Italic = False
                            pCell.Format.SpaceAfter = 4
                            Set pDCS = pCell
                        ElseIf InStr(1, pAsc, "NGAY") > 0 And InStr(1, pAsc, "THANG") > 0 Then
                            ' Dòng ngày tháng -> In thường, Nghiêng, không đậm, cỡ 14, căn sát phải
                            pCell.Range.Font.Size = 14
                            pCell.Range.Font.Bold = False
                            pCell.Range.Font.Italic = True
                            pCell.Format.SpaceBefore = 0
                            pCell.Format.SpaceAfter = 0
                        Else
                            pCell.Range.Font.Size = 14
                            pCell.Range.Font.Bold = False
                            pCell.Range.Font.Italic = False
                        End If
                    Else
                        ' --- CỘT BÊN TRÁI (CƠ QUAN BAN HÀNH, SỐ VĂN BẢN, TRÍCH YẾU V/V) ---
                        ' Bắt đầu dòng Trích yếu công văn (V/v...)
                        If InStr(1, pAsc, "V/V") > 0 Or InStr(1, pAsc, "V V") > 0 Or pAsc Like "V/V*" Or pAsc Like "V.V*" Or pAsc Like "VV*" Then
                            leftColInTrichYeu = True
                        End If
                        
                        If leftColInTrichYeu Then
                            ' Dòng Trích yếu V/v... -> Size 12, In nghiêng, không đậm, căn giữa
                            pCell.Format.Alignment = wdAlignParagraphCenter
                            pCell.Range.Font.Size = 12
                            pCell.Range.Font.Bold = False
                            pCell.Range.Font.Italic = True
                            pCell.Format.SpaceBefore = 0
                            pCell.Format.SpaceAfter = 0
                            pCell.Format.LineSpacingRule = wdLineSpaceSingle
                        Else
                            ' Cột bên trái thông thường -> Cỡ chữ 14, căn giữa
                            pCell.Format.Alignment = wdAlignParagraphCenter
                            pCell.Range.Font.Size = 14
                            
                            If pAsc = "*" Or pAsc = "***" Then
                                ' Dấu sao -> Thường, căn giữa
                                pCell.Range.Font.Bold = False
                                pCell.Range.Font.Italic = False
                            ElseIf pAsc Like "SO*" Then
                                ' Dòng Số: ... -> Thường, đứng, căn giữa
                                pCell.Range.Font.Bold = False
                                pCell.Range.Font.Italic = False
                            Else
                                ' Tên cơ quan
                                If leftColCqIdx = 1 Then
                                    ' Cơ quan cấp trên (ĐẢNG BỘ THÀNH PHỐ) -> In hoa, Thường (không đậm)
                                    pCell.Range.Case = wdUpperCase
                                    pCell.Range.Font.Bold = False
                                    pCell.Range.Font.Italic = False
                                ElseIf leftColCqIdx = 2 Then
                                    ' Cơ quan ban hành (ĐẢNG ỦY XÃ / CHI BỘ / VĂN PHÒNG) -> In hoa, Đậm
                                    pCell.Range.Case = wdUpperCase
                                    pCell.Range.Font.Bold = True
                                    pCell.Range.Font.Italic = False
                                Else
                                    pCell.Range.Font.Bold = False
                                    pCell.Range.Font.Italic = False
                                End If
                                leftColCqIdx = leftColCqIdx + 1
                            End If
                        End If
                    End If
                    
                    lineIndex = lineIndex + 1
SkipEmptyLineHeader:
                Next pCell
NextCellHeader:
                cellIdx = cellIdx + 1
            Next c
            
            ' --- TỰ ĐỘNG CHÈN SHAPE LINE DƯỚI "ĐẢNG CỘNG SẢN VIỆT NAM" ---
            If Not (pDCS Is Nothing) Then
                ' 1. Xóa TẤT CẢ đường kẻ / Shape cũ ở khu vực đầu trang (tránh nhân bản khi chạy nhiều lần)
                Dim sIdx As Long
                For sIdx = doc.Shapes.Count To 1 Step -1
                    Dim sItem As Shape
                    Set sItem = Nothing
                    On Error Resume Next
                    Set sItem = doc.Shapes(sIdx)
                    On Error GoTo 0
                    
                    If Not sItem Is Nothing Then
                        Dim shouldDel As Boolean
                        shouldDel = False
                        
                        ' A. Đã từng được đặt tên LineTieuNguDang
                        If StrComp(sItem.Name, "LineTieuNguDang", vbTextCompare) = 0 Then
                            shouldDel = True
                        End If
                        
                        ' B. Nằm trong bảng Header
                        On Error Resume Next
                        If sItem.Anchor.InRange(tbl.Range) Then shouldDel = True
                        On Error GoTo 0
                        
                        ' C. Đường kẻ ngang hoặc Shape có tên liên quan ở nửa trên trang 1 (Top < 250 pt)
                        Dim sNameLower As String
                        sNameLower = LCase(sItem.Name)
                        If InStr(sNameLower, "line") > 0 Or InStr(sNameLower, "thang") > 0 Or _
                           InStr(sNameLower, "thẳng") > 0 Or InStr(sNameLower, "straight") > 0 Or _
                           InStr(sNameLower, "connector") > 0 Or InStr(sNameLower, "shape") > 0 Then
                            On Error Resume Next
                            If sItem.Top < 250 Then shouldDel = True
                            If sItem.Anchor.Information(wdVerticalPositionRelativeToPage) < 250 Then shouldDel = True
                            On Error GoTo 0
                        End If
                        
                        ' D. Bất kỳ Shape nào có chiều cao nét mỏng (< 5 pt) ở đầu trang 1
                        On Error Resume Next
                        If sItem.Height < 5 And sItem.Top < 250 Then shouldDel = True
                        If (sItem.Type = 9 Or sItem.Type = 1) And sItem.Top < 250 Then shouldDel = True
                        On Error GoTo 0
                        
                        If shouldDel Then
                            On Error Resume Next
                            sItem.Delete
                            On Error GoTo 0
                        End If
                    End If
                Next sIdx
                
                ' Xóa các InlineShapes trong bảng Header nếu có
                Dim iShp As InlineShape
                For Each iShp In tbl.Range.InlineShapes
                    On Error Resume Next
                    iShp.Delete
                    On Error GoTo 0
                Next iShp
                
                ' Xóa gạch chân hoặc viền dưới đoạn tiêu ngữ nếu có
                pDCS.Range.Font.Underline = wdUnderlineNone
                On Error Resume Next
                pDCS.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                On Error GoTo 0
                
                ' 2. Tính toán vị trí lề phải chuẩn xác của ô chứa ĐẢNG CỘNG SẢN VIỆT NAM
                Dim cRightCell As Cell
                Set cRightCell = Nothing
                On Error Resume Next
                Set cRightCell = pDCS.Range.Cells(1)
                On Error GoTo 0
                If cRightCell Is Nothing Then Set cRightCell = cRight
                
                Dim rPad As Single
                rPad = 5.4
                Dim cellW As Single
                cellW = 248
                
                If Not (cRightCell Is Nothing) Then
                    cellW = cRightCell.Width
                    On Error Resume Next
                    rPad = cRightCell.RightPadding
                    On Error GoTo 0
                End If
                
                ' Chiều dài đường kẻ = 190 pt
                Dim lineWidth As Single
                lineWidth = 190
                
                ' Đẩy sang trái 5.5 pt để cân đối hoàn hảo dưới chữ ĐẢNG CỘNG SẢN VIỆT NAM (không bị lệch phải)
                Dim lPos As Single
                lPos = cellW - lineWidth - rPad - 5.5
                If lPos < 0 Then lPos = 0
                
                ' 3. Tạo Shape Line mới (Insert > Shapes > Line)
                On Error Resume Next
                Dim newL As Shape
                Set newL = doc.Shapes.AddLine(lPos, 16.5, lPos + lineWidth, 16.5, pDCS.Range)
                With newL
                    .Name = "LineTieuNguDang"
                    .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
                    .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
                    .Left = lPos
                    .Top = 16.5
                    .Line.Weight = 0.75
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .Line.DashStyle = 1 ' 1 = msoLineSolid
                    .WrapFormat.Type = 3 ' 3 = wdWrapNone
                End With
                On Error GoTo 0
            End If
            
            Exit For ' Đã hoàn thành chuẩn hóa bảng Header đầu tiên
        End If
    Next tbl
    
    '-----------------------------------------------------------------------
    ' BƯỚC 3.2: DÒ BẮT VÀ CHUẨN HÓA BẢNG KÍNH GỬI (CÔNG VĂN - HD 36-HD/VPTW)
    ' "kính gửi in nghiêng nằm trong table sát lề phải, liệt kê các phòng ban thì sát lề trái"
    '-----------------------------------------------------------------------
    For Each tbl In doc.Tables
        Dim tAscKG As String
        tAscKG = ToAscii(tbl.Range.Text)
        
        ' Bắt bảng Kính gửi: chứa KINH GUI, không phải Footer (không chứa NOI NHAN, BI THU)
        If InStr(1, tAscKG, "KINH GUI") > 0 And InStr(1, tAscKG, "NOI NHAN") = 0 And InStr(1, tAscKG, "BI THU") = 0 Then
            ' Tắt viền bảng (Borders None)
            On Error Resume Next
            tbl.Borders.Enable = False
            On Error GoTo 0
            
            Dim cKG As Cell
            For Each cKG In tbl.Range.Cells
                Dim cellTxtKG As String
                cellTxtKG = Trim(Replace(Replace(Replace(cKG.Range.Text, vbCr, ""), Chr(7), ""), vbLf, ""))
                
                If Len(cellTxtKG) > 0 Then
                    Dim cAscKG As String
                    cAscKG = ToAscii(cKG.Range.Text)
                    
                    ' Reset chuẩn ô: xóa viền và nền
                    On Error Resume Next
                    cKG.Borders.Enable = False
                    cKG.Shading.BackgroundPatternColor = wdColorAutomatic
                    cKG.Shading.Texture = wdTextureNone
                    On Error GoTo 0
                    
                    Dim isKGOther As Boolean
                    isKGOther = False
                    ' Nếu là ô tiêu ngữ hoặc cơ quan ở bảng 1 gộp thì bỏ qua
                    If InStr(1, cAscKG, "DANG CONG SAN") > 0 Or InStr(1, cAscKG, "DANG UY") > 0 Or InStr(1, cAscKG, "DANG BO") > 0 Then
                        isKGOther = True
                    End If
                    
                    If Not isKGOther Then
                        If InStr(1, cAscKG, "KINH GUI") > 0 Then
                            ' --- Ô KÍNH GỬI: CĂN SÁT LỀ PHẢI, IN NGHIÊNG, SIZE 14 ---
                            Dim pKGLeft As Paragraph
                            For Each pKGLeft In cKG.Range.Paragraphs
                                Do While pKGLeft.Range.Characters.Count > 1 And (pKGLeft.Range.Characters(1).Text = " " Or pKGLeft.Range.Characters(1).Text = vbTab)
                                    pKGLeft.Range.Characters(1).Delete
                                Loop
                                
                                Dim pAscLKG As String
                                pAscLKG = Trim(ToAscii(pKGLeft.Range.Text))
                                If Len(pAscLKG) > 0 Then
                                    With pKGLeft.Format
                                        .Alignment = wdAlignParagraphRight ' SÁT LỀ PHẢI Ô
                                        .LeftIndent = 0
                                        .RightIndent = 0
                                        .FirstLineIndent = 0
                                        .SpaceBefore = 0
                                        .SpaceAfter = 0
                                        .SpaceBeforeAuto = False
                                        .SpaceAfterAuto = False
                                        .LineSpacingRule = wdLineSpaceSingle
                                        .DisableLineHeightGrid = True
                                    End With
                                    pKGLeft.Range.Font.Name = "Times New Roman"
                                    pKGLeft.Range.Font.Size = 14
                                    pKGLeft.Range.Font.Bold = False
                                    pKGLeft.Range.Font.Italic = True ' IN NGHIÊNG
                                    pKGLeft.Range.Font.ColorIndex = wdBlack
                                    pKGLeft.Range.Font.Underline = wdUnderlineNone
                                End If
                            Next pKGLeft
                        Else
                            ' --- Ô LIỆT KÊ PHÒNG BAN: CĂN SÁT LỀ TRÁI, ĐỨNG THƯỜNG, SIZE 14 ---
                            ' 1. Tìm vị trí dòng liệt kê có nội dung cuối cùng
                            Dim lastLineKG As Integer
                            lastLineKG = 0
                            Dim counterKG As Integer
                            counterKG = 1
                            Dim tPkg As Paragraph
                            For Each tPkg In cKG.Range.Paragraphs
                                Dim sPkg As String
                                sPkg = Trim(Replace(Replace(Replace(tPkg.Range.Text, Chr(7), ""), vbCr, ""), vbLf, ""))
                                If Len(sPkg) > 0 Then lastLineKG = counterKG
                                counterKG = counterKG + 1
                            Next tPkg
                            
                            ' 2. Định dạng từng dòng
                            Dim pKGRight As Paragraph
                            Dim lineKGIdx As Integer
                            lineKGIdx = 1
                            For Each pKGRight In cKG.Range.Paragraphs
                                Do While pKGRight.Range.Characters.Count > 1 And (pKGRight.Range.Characters(1).Text = " " Or pKGRight.Range.Characters(1).Text = vbTab)
                                    pKGRight.Range.Characters(1).Delete
                                Loop
                                
                                Dim pAscRKG As String
                                pAscRKG = Trim(ToAscii(pKGRight.Range.Text))
                                If Len(pAscRKG) > 0 Then
                                    With pKGRight.Format
                                        .Alignment = wdAlignParagraphLeft ' SÁT LỀ TRÁI Ô
                                        .LeftIndent = 0
                                        .RightIndent = 0
                                        .FirstLineIndent = 0
                                        .SpaceBefore = 0
                                        .SpaceAfter = 0
                                        .SpaceBeforeAuto = False
                                        .SpaceAfterAuto = False
                                        .LineSpacingRule = wdLineSpaceSingle
                                        .DisableLineHeightGrid = True
                                    End With
                                    pKGRight.Range.Font.Name = "Times New Roman"
                                    pKGRight.Range.Font.Size = 14
                                    pKGRight.Range.Font.Bold = False
                                    pKGRight.Range.Font.Italic = False ' ĐỨNG THƯỜNG
                                    pKGRight.Range.Font.ColorIndex = wdBlack
                                    pKGRight.Range.Font.Underline = wdUnderlineNone
                                    
                                    ' 3. Chuẩn hóa dấu cuối dòng: triệt tiêu 100% lặp ".."
                                    On Error Resume Next
                                    Dim rngLineKG As Range
                                    Set rngLineKG = pKGRight.Range
                                    
                                    ' Thu hẹp bỏ qua các ký tự ngắt đoạn / ngắt ô
                                    Do While rngLineKG.End > rngLineKG.Start
                                        Dim cCodeKG As Long
                                        cCodeKG = AscW(Right(rngLineKG.Text, 1))
                                        If cCodeKG < 0 Then cCodeKG = cCodeKG + 65536
                                        If cCodeKG = 13 Or cCodeKG = 10 Or cCodeKG = 7 Then
                                            rngLineKG.End = rngLineKG.End - 1
                                        Else
                                            Exit Do
                                        End If
                                    Loop
                                    
                                    ' Ghi nhận vị trí cuối chuỗi hiển thị
                                    Dim rngEndPosKG As Long
                                    rngEndPosKG = rngLineKG.End
                                    
                                    ' Thu hẹp tiếp bỏ qua mọi dấu chấm, chấm phẩy, phẩy, khoảng trắng ở đuôi
                                    Do While rngLineKG.End > rngLineKG.Start
                                        Dim endChCodeKG As Long
                                        endChCodeKG = AscW(Right(rngLineKG.Text, 1))
                                        If endChCodeKG < 0 Then endChCodeKG = endChCodeKG + 65536
                                        If endChCodeKG = 32 Or endChCodeKG = 160 Or endChCodeKG = 9 Or endChCodeKG = 46 Or endChCodeKG = 59 Or endChCodeKG = 44 Then
                                            rngLineKG.End = rngLineKG.End - 1
                                        Else
                                            Exit Do
                                        End If
                                    Loop
                                    
                                    ' Vùng dấu ở đuôi
                                    Dim rngTailKG As Range
                                    Set rngTailKG = doc.Range(rngLineKG.End, rngEndPosKG)
                                    
                                    If lineKGIdx = lastLineKG Then
                                        rngTailKG.Text = "."
                                    Else
                                        rngTailKG.Text = ";"
                                    End If
                                    On Error GoTo 0
                                End If
                                lineKGIdx = lineKGIdx + 1
                            Next pKGRight
                        End If
                    End If
                End If
            Next cKG
        End If
    Next tbl
    
    '-----------------------------------------------------------------------
    ' BƯỚC 3.5: DÒ BẮT BẢNG NƠI NHẬN & CHỮ KÝ
    '-----------------------------------------------------------------------
    For Each tbl In doc.Tables
        Dim tAscFooter As String
        tAscFooter = ToAscii(tbl.Range.Text)
        
        If InStr(1, tAscFooter, "NOI NHAN") > 0 Or InStr(1, tAscFooter, "T/M") > 0 Or InStr(1, tAscFooter, "TM.") > 0 Or InStr(1, tAscFooter, "BI THU") > 0 Then
            ' Đã bắt được bảng phần Footer!
            For Each c In tbl.Range.Cells
                Dim cAscFooter As String
                cAscFooter = ToAscii(c.Range.Text)
                
                ' Phân loại xem ô này là Nơi nhận hay Chữ ký
                Dim isSignatureCell As Boolean
                isSignatureCell = False
                If InStr(1, cAscFooter, "T/M") > 0 Or InStr(1, cAscFooter, "TM.") > 0 Or InStr(1, cAscFooter, "BI THU") > 0 Then
                    isSignatureCell = True
                End If
                
                If Not isSignatureCell Then
                    ' --- CỘT NƠI NHẬN ---
                    ' 1. Tìm vị trí dòng liệt kê cuối cùng trong cột Nơi nhận
                    Dim lastListItemIdx As Integer
                    lastListItemIdx = 0
                    Dim tempP As Paragraph
                    Dim tCounter As Integer
                    tCounter = 1
                    For Each tempP In c.Range.Paragraphs
                        Dim tVal As String
                        tVal = Trim(Replace(Replace(Replace(tempP.Range.Text, Chr(7), ""), vbCr, ""), vbLf, ""))
                        If Len(tVal) > 0 Then
                            Dim tValAsc As String
                            tValAsc = ToAscii(tVal)
                            If InStr(1, tValAsc, "NOI NHAN") = 0 Then
                                lastListItemIdx = tCounter
                            End If
                        End If
                        tCounter = tCounter + 1
                    Next tempP
                    
                    ' 2. Định dạng chuẩn từng dòng trong cột Nơi nhận
                    Dim pNoiNhan As Paragraph
                    Dim pIndexNN As Integer
                    pIndexNN = 1
                    For Each pNoiNhan In c.Range.Paragraphs
                        Dim pTextNN As String
                        pTextNN = Trim(Replace(Replace(pNoiNhan.Range.Text, vbCr, ""), Chr(7), ""))
                        If Len(pTextNN) = 0 Then GoTo SkipNNLine
                        
                        On Error Resume Next
                        pNoiNhan.Borders.Enable = False
                        pNoiNhan.Shading.BackgroundPatternColor = wdColorAutomatic
                        pNoiNhan.Shading.Texture = wdTextureNone
                        On Error GoTo 0
                        
                        With pNoiNhan.Format
                            .Alignment = wdAlignParagraphLeft
                            .LeftIndent = 0
                            .FirstLineIndent = 0
                            .SpaceBefore = 0
                            .SpaceAfter = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpaceSingle
                            .DisableLineHeightGrid = True
                        End With
                        
                        pNoiNhan.Range.Font.Name = "Times New Roman"
                        pNoiNhan.Range.Font.ColorIndex = wdBlack
                        
                        Dim pTextAsc As String
                        pTextAsc = ToAscii(pTextNN)
                        
                        If InStr(1, pTextAsc, "NOI NHAN") > 0 Then
                            ' Chữ "Nơi nhận:": Size 14, Gạch chân, Không đậm
                            pNoiNhan.Range.Font.Size = 14
                            pNoiNhan.Range.Font.Underline = wdUnderlineSingle
                            pNoiNhan.Range.Font.Bold = False
                            pNoiNhan.Range.Font.Italic = False
                            pNoiNhan.Format.LeftIndent = 0
                            pNoiNhan.Format.FirstLineIndent = 0
                        Else
                            ' Các dòng liệt kê nơi nhận: Size 12, Không gạch chân, Không đậm
                            pNoiNhan.Range.Font.Name = "Times New Roman"
                            pNoiNhan.Range.Font.Size = 12
                            pNoiNhan.Range.Font.Underline = wdUnderlineNone
                            pNoiNhan.Range.Font.Bold = False
                            pNoiNhan.Range.Font.Italic = False
                            pNoiNhan.Format.LeftIndent = CentimetersToPoints(0.5)
                            pNoiNhan.Format.FirstLineIndent = CentimetersToPoints(-0.5)
                            
                            ' Đảm bảo đầu dòng có gạch ngang "-" nếu thiếu
                            Dim firstChar As String
                            firstChar = Left(pTextNN, 1)
                            If firstChar <> "-" Then
                                If firstChar = ChrW(8211) Or firstChar = ChrW(8212) Or firstChar = "*" Or firstChar = "+" Or firstChar = ChrW(8226) Then
                                    pNoiNhan.Range.Characters(1).Text = "-"
                                Else
                                    pNoiNhan.Range.InsertBefore "- "
                                End If
                            End If
                            
                            ' Chuẩn hóa dấu câu cuối dòng an toàn tuyệt đối, triệt tiêu 100% lặp ".."
                            On Error Resume Next
                            Dim rngLineNN As Range
                            Set rngLineNN = pNoiNhan.Range
                            
                            ' Thu hẹp bỏ qua các ký tự ngắt đoạn / ngắt ô
                            Do While rngLineNN.End > rngLineNN.Start
                                Dim cCodeNN As Long
                                cCodeNN = AscW(Right(rngLineNN.Text, 1))
                                If cCodeNN < 0 Then cCodeNN = cCodeNN + 65536
                                If cCodeNN = 13 Or cCodeNN = 10 Or cCodeNN = 7 Then
                                    rngLineNN.End = rngLineNN.End - 1
                                Else
                                    Exit Do
                                End If
                            Loop
                            
                            ' Ghi nhận vị trí cuối chuỗi hiển thị
                            Dim rngEndPosNN As Long
                            rngEndPosNN = rngLineNN.End
                            
                            ' Thu hẹp tiếp bỏ qua mọi dấu chấm, chấm phẩy, phẩy, khoảng trắng ở đuôi
                            Do While rngLineNN.End > rngLineNN.Start
                                Dim endChCodeNN As Long
                                endChCodeNN = AscW(Right(rngLineNN.Text, 1))
                                If endChCodeNN < 0 Then endChCodeNN = endChCodeNN + 65536
                                If endChCodeNN = 32 Or endChCodeNN = 160 Or endChCodeNN = 9 Or endChCodeNN = 46 Or endChCodeNN = 59 Or endChCodeNN = 44 Then
                                    rngLineNN.End = rngLineNN.End - 1
                                Else
                                    Exit Do
                                End If
                            Loop
                            
                            ' Vùng dấu ở đuôi
                            Dim rngTailNN As Range
                            Set rngTailNN = doc.Range(rngLineNN.End, rngEndPosNN)
                            
                            If pIndexNN = lastListItemIdx Then
                                rngTailNN.Text = "."
                            Else
                                rngTailNN.Text = ";"
                            End If
                            On Error GoTo 0
                        End If
SkipNNLine:
                        pIndexNN = pIndexNN + 1
                    Next pNoiNhan
                Else
                    ' --- CỘT CHỮ KÝ: TỰ ĐỘNG ĐẢM BẢO ĐÚNG 4 DÒNG ENTER ĐỂ KÝ TÊN ---
                    On Error Resume Next
                    Dim textIndices(1 To 20) As Integer
                    Dim textCount As Integer
                    textCount = 0
                    
                    Dim pScan As Paragraph
                    Dim scanIdx As Integer
                    scanIdx = 1
                    For Each pScan In c.Range.Paragraphs
                        Dim chkText As String
                        chkText = Trim(Replace(Replace(Replace(pScan.Range.Text, Chr(7), ""), vbCr, ""), vbLf, ""))
                        If Len(chkText) > 0 Then
                            textCount = textCount + 1
                            textIndices(textCount) = scanIdx
                        End If
                        scanIdx = scanIdx + 1
                    Next pScan
                    
                    If textCount >= 2 Then
                        Dim sigIdx As Integer
                        Dim prevTitleIdx As Integer
                        Dim curEmpty As Integer
                        
                        sigIdx = textIndices(textCount)
                        prevTitleIdx = textIndices(textCount - 1)
                        curEmpty = sigIdx - prevTitleIdx - 1
                        
                        If curEmpty < 4 Then
                            Dim nAdd As Integer
                            For nAdd = 1 To (4 - curEmpty)
                                Dim rIns As Range
                                Set rIns = doc.Range(c.Range.Paragraphs(sigIdx).Range.Start, c.Range.Paragraphs(sigIdx).Range.Start)
                                rIns.Text = vbCr
                            Next nAdd
                        ElseIf curEmpty > 4 Then
                            Dim nSub As Integer
                            For nSub = 1 To (curEmpty - 4)
                                c.Range.Paragraphs(prevTitleIdx + 1).Range.Delete
                            Next nSub
                        End If
                    End If
                    On Error GoTo 0
                    
                    ' Định dạng chuẩn cho toàn bộ các dòng trong ô Chữ ký
                    Dim pSig As Paragraph
                    Dim sigLineIdx As Integer
                    Dim totalSigParas As Integer
                    totalSigParas = c.Range.Paragraphs.Count
                    sigLineIdx = 1
                    
                    For Each pSig In c.Range.Paragraphs
                        On Error Resume Next
                        pSig.Borders.Enable = False
                        pSig.Shading.BackgroundPatternColor = wdColorAutomatic
                        pSig.Shading.Texture = wdTextureNone
                        On Error GoTo 0
                        
                        With pSig.Format
                            .Alignment = wdAlignParagraphCenter
                            .LeftIndent = 0
                            .RightIndent = 0
                            .FirstLineIndent = 0
                            .SpaceBefore = 0
                            .SpaceAfter = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpaceSingle
                            .DisableLineHeightGrid = True
                        End With
                        
                        pSig.Range.Font.Name = "Times New Roman"
                        pSig.Range.Font.Size = 14
                        pSig.Range.Font.ColorIndex = wdBlack
                        pSig.Range.Font.Underline = wdUnderlineNone
                        
                        Dim curText As String
                        curText = Trim(Replace(Replace(Replace(pSig.Range.Text, Chr(7), ""), vbCr, ""), vbLf, ""))
                        
                        If Len(curText) > 0 Then
                            If InStr(1, UCase(curText), "T/M") > 0 Or InStr(1, UCase(curText), "TM.") > 0 Then
                                ' Dòng T/M CHI BỘ: In hoa, Đậm
                                pSig.Range.Case = wdUpperCase
                                pSig.Range.Font.Bold = True
                            ElseIf sigLineIdx = totalSigParas Then
                                ' Dòng cuối cùng (Tên người ký): In đậm
                                pSig.Range.Font.Bold = True
                            ElseIf sigLineIdx = 2 And textCount >= 3 Then
                                ' Dòng chức danh thứ 2 (BÍ THƯ): In hoa, KHÔNG đậm
                                pSig.Range.Case = wdUpperCase
                                pSig.Range.Font.Bold = False
                            ElseIf sigLineIdx = 1 Then
                                pSig.Range.Case = wdUpperCase
                                pSig.Range.Font.Bold = True
                            Else
                                pSig.Range.Font.Bold = True
                            End If
                        Else
                            ' 4 dòng Enter khoảng cách ký tên: Cỡ chữ 14, không đậm
                            pSig.Range.Font.Bold = False
                            pSig.Range.Font.Size = 14
                        End If
                        
                        sigLineIdx = sigLineIdx + 1
                    Next pSig
                End If
            Next c
        End If
    Next tbl
    
    '-----------------------------------------------------------------------
    ' BƯỚC 4: ĐỊNH DẠNG ĐỒNG NHẤT 100% TOÀN BỘ THÂN VĂN BẢN
    ' (MỌI ĐOẠN VĂN ĐỀU CÙNG 1 FORMAT DUY NHẤT: FIRST LINE 1.27CM, AFTER 6PT)
    '-----------------------------------------------------------------------
    Dim p As Paragraph
    Dim s As String
    Dim titleState As Integer
    titleState = 0
    
    For Each p In doc.Paragraphs
        ' Nếu nằm trong bảng thì bỏ qua
        If p.Range.Information(wdWithInTable) Then
            GoTo NextParagraphMain
        End If
        
        ' Lấy text chuẩn và an toàn (Trim không tác dụng nếu còn vbCr ở cuối)
        s = p.Range.Text
        s = Replace(Replace(s, vbCr, ""), vbLf, "")
        s = Replace(s, Chr(7), "")
        s = Replace(s, Chr(11), "")
        s = Replace(s, Chr(160), " ")
        s = Trim(s)
        
        ' 1. XÓA TẬN GỐC ĐƯỜNG VIỀN ẨN VÀ NỀN TRẮNG (Kể cả dòng trống)
        On Error Resume Next
        p.Borders.Enable = False
        p.Borders(wdBorderTop).LineStyle = wdLineStyleNone
        p.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        p.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        p.Borders(wdBorderRight).LineStyle = wdLineStyleNone
        p.Shading.BackgroundPatternColor = wdColorAutomatic
        p.Shading.Texture = wdTextureNone
        On Error GoTo 0
        
        If Len(s) = 0 Then
            ' Đảm bảo dòng trống (Enter) có chiều cao chuẩn 1 dòng để làm khoảng cách
            p.Range.Font.Size = 15
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.LineSpacingRule = wdLineSpaceSingle
            GoTo NextParagraphMain
        End If
        
        ' 2. FONT CHỮ CHUẨN TIMES NEW ROMAN
        p.Range.Font.Name = "Times New Roman"
        p.Range.Font.ColorIndex = wdBlack
        
        ' --- QUẢN LÝ KHỐI TỰA ĐỀ (TÊN LOẠI VĂN BẢN, TRÍCH YẾU, DÒNG PHÂN CÁCH) ---
        Dim sAscPara As String
        sAscPara = ToAscii(s)
        
        If titleState = 0 Then
            ' Chưa vào Title Block: Kiểm tra xem dòng hiện tại có phải Tên loại văn bản Đảng không
            If IsPartyDocTitle(sAscPara) Then
                titleState = 1
                With p.Format
                    .Alignment = wdAlignParagraphCenter
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    .SpaceBefore = 14
                    .SpaceAfter = 4
                    .LineSpacingRule = wdLineSpaceSingle
                    .DisableLineHeightGrid = True
                End With
                p.Range.Case = wdUpperCase
                p.Range.Font.Name = "Times New Roman"
                p.Range.Font.Size = 16
                p.Range.Font.Bold = True
                p.Range.Font.Italic = False
                p.Range.Font.Underline = wdUnderlineNone
                GoTo NextParagraphMain
            End If
            
        ElseIf titleState = 1 Then
            ' Vừa qua Dòng 1 (Tên loại), đang xét dòng tiếp theo:
            If IsSeparatorLine(s) Then
                titleState = 4
                With p.Format
                    .Alignment = wdAlignParagraphCenter
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 14
                    .LineSpacingRule = wdLineSpaceSingle
                    .DisableLineHeightGrid = True
                End With
                p.Range.Font.Name = "Times New Roman"
                p.Range.Font.Size = 15
                p.Range.Font.Bold = True ' IN ĐẬM DÒNG PHÂN CÁCH
                p.Range.Font.Italic = False
                p.Range.Font.Underline = wdUnderlineNone
                GoTo NextParagraphMain
                
            ElseIf IsBodyStart(s, sAscPara) Then
                titleState = 4
                ' Không GoTo, để trôi thẳng xuống format Thân bài bên dưới
                
            Else
                ' Dòng 1 Trích yếu nội dung (ví dụ: "Về lãnh đạo thực hiện nhiệm vụ công tác tháng 10 năm 2026")
                titleState = 2
                With p.Format
                    .Alignment = wdAlignParagraphCenter
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 6
                    .LineSpacingRule = wdLineSpaceSingle
                    .DisableLineHeightGrid = True
                End With
                p.Range.Font.Name = "Times New Roman"
                p.Range.Font.Size = 15
                p.Range.Font.Bold = True ' IN ĐẬM THEO ĐÚNG CHUẨN HD 36
                p.Range.Font.Italic = False
                p.Range.Font.Underline = wdUnderlineNone
                GoTo NextParagraphMain
            End If
            
        ElseIf titleState = 2 Then
            ' Đã có Dòng 1 Trích yếu. Đang xét dòng tiếp theo:
            If IsSeparatorLine(s) Then
                ' DÒNG PHÂN CÁCH (------ hoặc ---):
                titleState = 4 ' Kết thúc khối tựa đề!
                With p.Format
                    .Alignment = wdAlignParagraphCenter
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 14 ' Giãn dưới 14 pt phân cách với thân bài
                    .LineSpacingRule = wdLineSpaceSingle
                    .DisableLineHeightGrid = True
                End With
                p.Range.Font.Name = "Times New Roman"
                p.Range.Font.Size = 15
                p.Range.Font.Bold = True ' IN ĐẬM THEO YÊU CẦU NGƯỜI DÙNG!
                p.Range.Font.Italic = False
                p.Range.Font.Underline = wdUnderlineNone
                GoTo NextParagraphMain
                
            ElseIf IsBodyStart(s, sAscPara) Or Len(s) > 60 Or Right(s, 1) = "." Or Right(s, 1) = ";" Or Right(s, 1) = ":" Or InStr(s, ". ") > 0 Then
                ' VÔ TÌNH KHÔNG CÓ "------"! Nhận biết ngay đây là bắt đầu Thân bài!
                titleState = 4 ' Thoát khối tựa đề NGAY LẬP TỨC!
                ' Không GoTo, để trôi thẳng xuống format Thân bài (căn đều 2 bên, thụt lề 1.27cm, before/after 6pt)
                
            Else
                ' Dòng 2 nối tiếp của Trích yếu (trích yếu dài phải ngắt 2 dòng)
                titleState = 3 ' Đã có 2 dòng trích yếu, dòng tiếp theo BẮT BUỘC thoát sang thân bài hoặc separator!
                With p.Format
                    .Alignment = wdAlignParagraphCenter
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 6
                    .LineSpacingRule = wdLineSpaceSingle
                    .DisableLineHeightGrid = True
                End With
                p.Range.Font.Name = "Times New Roman"
                p.Range.Font.Size = 15
                p.Range.Font.Bold = True
                p.Range.Font.Italic = False
                p.Range.Font.Underline = wdUnderlineNone
                GoTo NextParagraphMain
            End If
            
        ElseIf titleState = 3 Then
            ' Đã có 2 dòng Trích yếu. Dòng tiếp theo:
            If IsSeparatorLine(s) Then
                titleState = 4 ' Dòng phân cách
                With p.Format
                    .Alignment = wdAlignParagraphCenter
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 14
                    .LineSpacingRule = wdLineSpaceSingle
                    .DisableLineHeightGrid = True
                End With
                p.Range.Font.Name = "Times New Roman"
                p.Range.Font.Size = 15
                p.Range.Font.Bold = True ' IN ĐẬM THEO YÊU CẦU NGƯỜI DÙNG!
                p.Range.Font.Italic = False
                p.Range.Font.Underline = wdUnderlineNone
                GoTo NextParagraphMain
            Else
                ' Dù có hay không có dòng phân cách, dòng này CHẮC CHẮN là Thân bài!
                titleState = 4
                ' Không GoTo, để trôi thẳng xuống format Thân bài bên dưới!
            End If
        End If
        
        ' --- XỬ LÝ KÍNH GỬI / NƠI NHẬN ---
        If Len(s) < 100 Then
            If UCase(Left(s, 9)) Like "N?I NH?N:" Or UCase(Left(s, 8)) Like "K?NH G?I" Then
                With p.Format
                    .Alignment = wdAlignParagraphLeft
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpaceSingle
                    .DisableLineHeightGrid = True
                End With
                p.Range.Font.Size = 14
                p.Range.Font.Underline = wdUnderlineSingle
                p.Range.Font.Bold = False
                GoTo NextParagraphMain
            End If
        End If
        
        ' 4. TOÀN BỘ NỘI DUNG VĂN BẢN (100% CÙNG 1 BỘ FORMAT ĐỒNG NHẤT DUY NHẤT)
        ' - Cỡ chữ chuẩn thể thức: Times New Roman 15 pt
        p.Range.Font.Name = "Times New Roman"
        p.Range.Font.Size = 15
        p.Range.Font.ColorIndex = wdBlack
        ' - Căn lề: Justified (Đều 2 bên)
        ' - Thụt đầu dòng: First line 1.27 cm
        ' - Lề trái: 0 cm, Lề phải: 0 cm
        ' - Giãn đoạn: Before 6 pt, After 6 pt
        ' - Giãn dòng: Single
        With p.Format
            .Alignment = wdAlignParagraphJustify
            .LeftIndent = 0
            .RightIndent = 0
            .FirstLineIndent = CentimetersToPoints(1.27)
            .SpaceBefore = 6
            .SpaceAfter = 6
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .DisableLineHeightGrid = True
        End With
        
        ' Tự động nhận diện in đậm: CHỈ CÁC ĐẦU MỤC MỚI IN ĐẬM, CÒN LẠI NỘI DUNG LÀ CHỮ THƯỜNG (KHÔNG ĐẬM)
        Dim isHeading As Boolean
        isHeading = False
        
        If s Like "I.*" Or s Like "II.*" Or s Like "III.*" Or s Like "IV.*" Or s Like "V.*" Or s Like "VI.*" Or s Like "VII.*" Or s Like "VIII.*" Then
            isHeading = True
            p.Range.Case = wdUpperCase
        ElseIf (s Like "#. *" Or s Like "##. *" Or s Like "#.#. *" Or s Like "#.##. *" Or s Like "##.#. *" Or s Like "#.#.#*" Or Left(s, 1) = "*") And Len(s) < 160 Then
            isHeading = True
        End If
        
        If isHeading Then
            p.Range.Font.Bold = True
        Else
            ' Toàn bộ nội dung bình thường, gạch đầu dòng, dấu cộng, bullet -> CHỮ THƯỜNG (KHÔNG ĐẬM)
            p.Range.Font.Bold = False
        End If
        
NextParagraphMain:
    Next p
    
    '-----------------------------------------------------------------------
    ' BƯỚC 5: TỰ ĐỘNG CĂN VỪA CỬA SỔ TRANG IN (AUTOFIT WINDOW) CHO TOÀN BỘ BẢNG
    '-----------------------------------------------------------------------
    Dim tblAll As Table
    For Each tblAll In doc.Tables
        On Error Resume Next
        tblAll.AutoFitBehavior 2 ' wdAutoFitWindow = 2
        tblAll.AllowAutoFit = True
        On Error GoTo 0
    Next tblAll
    
    Application.ScreenUpdating = True
    MsgBox "Da chinh the thuc theo Huong dan 05-HD/VPTW." & vbCrLf & vbCrLf & _
           "Created by VPDU xa Thanh Phu", _
           vbInformation, "Chuan Hoa The Thuc"
End Sub

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

'-----------------------------------------------------------------------
' HAM KIEM TRA TEN LOAI VAN BAN DANG (HD 36-HD/VPTW & QD 66-QD/TW)
'-----------------------------------------------------------------------
Private Function IsPartyDocTitle(ByVal sAsc As String) As Boolean
    Dim t As String
    t = Trim(sAsc)
    If Len(t) = 0 Or Len(t) > 70 Then
        IsPartyDocTitle = False
        Exit Function
    End If
    
    Dim match As Boolean
    match = False
    
    If t = "NGHI QUYET" Or t Like "NGHI QUYET*" Then
        match = True
    ElseIf t = "QUYET DINH" Or t Like "QUYET DINH*" Then
        match = True
    ElseIf t = "CHI THI" Or t Like "CHI THI*" Then
        match = True
    ElseIf t = "KET LUAN" Or t Like "KET LUAN*" Then
        match = True
    ElseIf t = "QUY CHE" Or t Like "QUY CHE*" Then
        match = True
    ElseIf t = "QUY DINH" Or t Like "QUY DINH*" Then
        match = True
    ElseIf t = "THONG TRI" Or t Like "THONG TRI*" Then
        match = True
    ElseIf t = "HUONG DAN" Or t Like "HUONG DAN*" Then
        match = True
    ElseIf t = "THONG BAO" Or t Like "THONG BAO*" Then
        match = True
    ElseIf t = "KE HOACH" Or t Like "KE HOACH*" Then
        match = True
    ElseIf t = "CHUONG TRINH" Or t Like "CHUONG TRINH*" Then
        match = True
    ElseIf t = "BAO CAO" Or t Like "BAO CAO*" Then
        match = True
    ElseIf t = "BIEN BAN" Or t Like "BIEN BAN*" Then
        match = True
    ElseIf t = "TO TRINH" Or t Like "TO TRINH*" Then
        match = True
    ElseIf t = "CONG DIEN" Or t Like "CONG DIEN*" Then
        match = True
    ElseIf t = "DE AN" Or t Like "DE AN*" Then
        match = True
    ElseIf t = "PHUONG AN" Or t Like "PHUONG AN*" Then
        match = True
    ElseIf t = "DU THAO" Or t Like "DU THAO*" Then
        match = True
    ElseIf t = "GIAY MOI" Or t Like "GIAY MOI*" Then
        match = True
    ElseIf t = "GIAY GIOI THIEU" Or t Like "GIAY GIOI THIEU*" Then
        match = True
    ElseIf t = "BAN CAM KET" Or t Like "BAN CAM KET*" Then
        match = True
    ElseIf t = "BAN DANG KY" Or t Like "BAN DANG KY*" Then
        match = True
    ElseIf t = "DIEU LE" Or t Like "DIEU LE*" Then
        match = True
    End If
    
    IsPartyDocTitle = match
End Function

'-----------------------------------------------------------------------
' HAM KIEM TRA DONG PHAN CACH (---, ------, – – –, ___, * * *)
'-----------------------------------------------------------------------
Private Function IsSeparatorLine(ByVal rawText As String) As Boolean
    Dim cleanStr As String
    cleanStr = Trim(Replace(Replace(Replace(rawText, vbCr, ""), vbLf, ""), " ", ""))
    cleanStr = Replace(cleanStr, Chr(160), "") ' Non-breaking space
    cleanStr = Replace(cleanStr, vbTab, "")
    
    If Len(cleanStr) < 2 Or Len(cleanStr) > 40 Then
        IsSeparatorLine = False
        Exit Function
    End If
    
    Dim isAllDash As Boolean
    isAllDash = True
    Dim i As Long
    For i = 1 To Len(cleanStr)
        Dim wCode As Long
        wCode = AscW(Mid(cleanStr, i, 1))
        If wCode < 0 Then wCode = wCode + 65536
        ' 45 = Hyphen (-), 95 = Underscore (_), 42 = Asterisk (*)
        ' 8211 = En dash (–), 8212 = Em dash (—), 8213 = Horizontal bar
        If wCode <> 45 And wCode <> 95 And wCode <> 42 And _
           wCode <> 8211 And wCode <> 8212 And wCode <> 8213 Then
            isAllDash = False
            Exit For
        End If
    Next i
    
    IsSeparatorLine = isAllDash
End Function

'-----------------------------------------------------------------------
' HAM KIEM TRA BAT DAU THAN VAN BAN (IS BODY START)
' DUNG DE NHAN BIET DIEM KET THUC KHOI TUA DE NGAY CA KHI THIEU "------"
'-----------------------------------------------------------------------
Private Function IsBodyStart(ByVal rawText As String, ByVal sAsc As String) As Boolean
    Dim tAsc As String
    tAsc = Trim(sAsc)
    Dim tRaw As String
    tRaw = Trim(rawText)
    
    If Len(tAsc) = 0 Then
        IsBodyStart = False
        Exit Function
    End If
    
    Dim match As Boolean
    match = False
    
    ' 1. Cac tu khoa phap ly / tham quyen mo dau than bai
    If tAsc Like "CAN CU*" Or tAsc Like "THUC HIEN*" Or tAsc Like "DE THUC HIEN*" Or _
       tAsc Like "XET DE NGHI*" Or tAsc Like "THEO DE NGHI*" Or _
       tAsc Like "BAN THUONG VU*" Or tAsc Like "BAN CHAP HANH*" Or _
       tAsc Like "DANG UY*" Or tAsc Like "CHI BO*" Or _
       tAsc Like "HOI NGHI*" Or tAsc Like "DAI HOI*" Or _
       tAsc Like "QUYET DINH:*" Or tAsc Like "QUYET DINH :*" Or _
       tAsc Like "YEU CAU:*" Or tAsc Like "YEU CAU :*" Then
        match = True
    End If
    
    ' 2. Cac dau muc phan cap (La Ma, Chuong, Muc, Dieu, Khoan, So thu tu)
    If Not match Then
        If tAsc Like "I.*" Or tAsc Like "II.*" Or tAsc Like "III.*" Or _
           tAsc Like "IV.*" Or tAsc Like "V.*" Or tAsc Like "VI.*" Or _
           tAsc Like "VII.*" Or tAsc Like "VIII.*" Or tAsc Like "IX.*" Or tAsc Like "X.*" Then
            match = True
        ElseIf tAsc Like "PHAN I*" Or tAsc Like "PHAN 1*" Or tAsc Like "PHAN THU*" Or _
               tAsc Like "MUC I*" Or tAsc Like "MUC 1*" Or _
               tAsc Like "CHUONG I*" Or tAsc Like "CHUONG 1*" Or _
               tAsc Like "DIEU 1*" Or tAsc Like "DIEU 2*" Or tAsc Like "DIEU *" Or _
               tAsc Like "KHOAN 1*" Or tAsc Like "KHOAN 2*" Then
            match = True
        ElseIf tAsc Like "1.*" Or tAsc Like "2.*" Or tAsc Like "3.*" Or _
               tAsc Like "4.*" Or tAsc Like "5.*" Or tAsc Like "6.*" Or _
               tAsc Like "7.*" Or tAsc Like "8.*" Or tAsc Like "9.*" Or _
               tAsc Like "1/*" Or tAsc Like "2/*" Or tAsc Like "3/*" Then
            match = True
        ElseIf tAsc Like "A.*" Or tAsc Like "B.*" Or tAsc Like "C.*" Or tAsc Like "D.*" Then
            match = True
        ElseIf Left(tRaw, 1) = "-" Or Left(tRaw, 1) = "+" Or Left(tRaw, 1) = "*" Then
            match = True
        End If
    End If
    
    ' 3. Tu khoa thoi gian / tinh hinh / nhiem vu mo dau
    If Not match Then
        If tAsc Like "TRONG THANG*" Or tAsc Like "TRONG QUY*" Or tAsc Like "TRONG NAM*" Or _
           tAsc Like "NAM 20*" Or tAsc Like "THANG 1*" Or tAsc Like "THANG 2*" Or _
           tAsc Like "THANG 3*" Or tAsc Like "THANG 4*" Or tAsc Like "THANG 5*" Or _
           tAsc Like "THANG 6*" Or tAsc Like "THANG 7*" Or tAsc Like "THANG 8*" Or _
           tAsc Like "THANG 9*" Or tAsc Like "THANG 10*" Or tAsc Like "THANG 11*" Or tAsc Like "THANG 12*" Or _
           tAsc Like "TINH HINH*" Or tAsc Like "DANH GIA*" Or _
           tAsc Like "MUC DICH*" Or tAsc Like "NOI DUNG*" Then
            match = True
        End If
    End If
    
    ' 4. Dau hieu cu phap (ket thuc bang ';' hoac ':' hoac do dai > 60)
    If Not match Then
        Dim lastCh As String
        lastCh = Right(tRaw, 1)
        If lastCh = ";" Or lastCh = ":" Then
            match = True
        ElseIf Len(tRaw) > 60 Then
            match = True
        End If
    End If
    
    IsBodyStart = match
End Function

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
