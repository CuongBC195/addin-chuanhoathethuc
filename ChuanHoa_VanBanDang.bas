'===========================================================================
' MACRO CHUẨN HÓA & RÀ SOÁT LOGIC VĂN BẢN ĐẢNG THEO HD 05-HD/VPTW
' (Bản nâng cấp triệt tiêu lỗi ToAscii, bảo toàn in đậm cục bộ & rà soát 10 nhóm Logic)
'===========================================================================

Public Sub RibbonChuanHoa(control As Object)
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
    doc.ConvertNumbersToText
    
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
    
    Dim i As Integer
    For i = 1 To 5
        With doc.Content.Find
            .Text = "^p "
            .Replacement.Text = "^p"
            .Execute Replace:=wdReplaceAll
        End With
        With doc.Content.Find
            .Text = "^p^t"
            .Replacement.Text = "^p"
            .Execute Replace:=wdReplaceAll
        End With
    Next i
    
    ' Xóa các dòng trắng thừa liên tiếp ngoài bảng (Duyệt ngược từ dưới lên để không bị nhảy cóc)
    Dim pIdx As Long
    For pIdx = doc.Paragraphs.Count To 2 Step -1
        Dim pCur As Paragraph, pPrev As Paragraph
        Set pCur = doc.Paragraphs(pIdx)
        Set pPrev = doc.Paragraphs(pIdx - 1)
        
        If Not pCur.Range.Information(wdWithInTable) And Not pPrev.Range.Information(wdWithInTable) Then
            Dim sCur As String, sPrev As String
            sCur = CleanParaText(pCur.Range.Text)
            sPrev = CleanParaText(pPrev.Range.Text)
            If Len(sCur) = 0 And Len(sPrev) = 0 Then
                On Error Resume Next
                pCur.Range.Delete
                On Error GoTo 0
            End If
        End If
    Next pIdx
    
    ' Xóa nhiều dấu cách liên tiếp
    For i = 1 To 5
        With doc.Content.Find
            .Text = "  "
            .Replacement.Text = " "
            .Execute Replace:=wdReplaceAll
        End With
    Next i
    
    '-----------------------------------------------------------------------
    ' BƯỚC 1: CẤU HÌNH TRANG IN A4 CHUẨN 05-HD/VPTW
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
    ' BƯỚC 2: ĐÁNH SỐ TRANG ĐỈNH GIỮA (TRỪ TRANG 1)
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
    Dim tFit As Table
    For Each tFit In doc.Tables
        On Error Resume Next
        tFit.AutoFitBehavior 2 ' wdAutoFitWindow = 2
        tFit.AllowAutoFit = True
        On Error GoTo 0
    Next tFit
    
    Dim tbl As Table
    Dim tblCounter As Integer
    tblCounter = 0
    Dim c As Cell
    Dim pCell As Paragraph
    Dim pDCS As Paragraph
    Dim cRight As Cell
    
    For Each tbl In doc.Tables
        tblCounter = tblCounter + 1
        Dim tAsc As String
        tAsc = ToAscii(tbl.Range.Text)
        
        Dim isHeaderTbl As Boolean
        isHeaderTbl = False
        
        If tblCounter = 1 And (InStr(1, tAsc, "DANG CONG SAN") > 0 Or InStr(1, tAsc, "VIET NAM") > 0 Or InStr(1, tAsc, "DANG BO") > 0 Or InStr(1, tAsc, "DANG UY") > 0 Or InStr(1, tAsc, "CHI BO") > 0) Then
            isHeaderTbl = True
        ElseIf InStr(1, tAsc, "DANG CONG SAN VIET NAM") > 0 Then
            isHeaderTbl = True
        End If
        
        If isHeaderTbl Then
            Set pDCS = Nothing
            Set cRight = Nothing
            
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
                cellTextRaw = CleanParaText(c.Range.Text)
                If Len(cellTextRaw) = 0 Then GoTo NextCellHeader
                
                Dim cAsc As String
                cAsc = ToAscii(c.Range.Text)
                
                If InStr(1, cAsc, "KINH GUI") > 0 Or Left(cellTextRaw, 1) = "-" Then
                    GoTo NextCellHeader
                End If
                
                Dim isRightCell As Boolean
                isRightCell = False
                
                If InStr(1, cAsc, "DANG CONG SAN") > 0 Or InStr(1, cAsc, "VIET NAM") > 0 Then
                    isRightCell = True
                    Set cRight = c
                ElseIf InStr(1, cAsc, "NGAY") > 0 And InStr(1, cAsc, "THANG") > 0 And InStr(1, cAsc, "NAM") > 0 Then
                    isRightCell = True
                    If cRight Is Nothing Then Set cRight = c
                ElseIf cellIdx = totalCells And totalCells >= 2 Then
                    isRightCell = True
                    If cRight Is Nothing Then Set cRight = c
                Else
                    On Error Resume Next
                    If c.ColumnIndex > 1 And tbl.Columns.Count >= 2 And c.ColumnIndex = tbl.Columns.Count Then
                        isRightCell = True
                    End If
                    On Error GoTo 0
                End If
                
                On Error Resume Next
                c.Borders.Enable = False
                c.Shading.BackgroundPatternColor = wdColorAutomatic
                c.Shading.Texture = wdTextureNone
                On Error GoTo 0
                
                For Each pCell In c.Range.Paragraphs
                    Do While pCell.Range.Characters.Count > 1 And (pCell.Range.Characters(1).Text = " " Or pCell.Range.Characters(1).Text = vbTab)
                        pCell.Range.Characters(1).Delete
                    Loop
                    
                    Dim pAsc As String
                    pAsc = Trim(ToAscii(pCell.Range.Text))
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
                        .DisableLineHeightGrid = True
                    End With
                    
                    pCell.Range.Font.Name = "Times New Roman"
                    pCell.Range.Font.ColorIndex = wdBlack
                    pCell.Range.Font.Underline = wdUnderlineNone
                    
                    If isRightCell Then
                        pCell.Format.Alignment = wdAlignParagraphRight
                        
                        If InStr(1, pAsc, "DANG CONG SAN") > 0 Or InStr(1, pAsc, "VIET NAM") > 0 Then
                            pCell.Range.Case = wdUpperCase
                            pCell.Range.Font.Size = 15
                            pCell.Range.Font.Bold = True
                            pCell.Range.Font.Italic = False
                            pCell.Format.SpaceAfter = 4
                            Set pDCS = pCell
                        ElseIf InStr(1, pAsc, "NGAY") > 0 And InStr(1, pAsc, "THANG") > 0 Then
                            Dim origNgay As String, fixNgay As String
                            origNgay = CleanParaText(pCell.Range.Text)
                            fixNgay = ChuanHoa_NgayThang(origNgay)
                            If fixNgay <> origNgay And Len(fixNgay) > 0 Then
                                Call SafeSetParagraphText(doc, pCell, fixNgay)
                            End If
                            pCell.Format.Alignment = wdAlignParagraphRight
                            pCell.Range.Font.Name = "Times New Roman"
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
                        If pAsc Like "V/V*" Or pAsc Like "V.V*" Or pAsc Like "VE VIEC*" Then
                            leftColInTrichYeu = True
                        End If
                        
                        If leftColInTrichYeu Then
                            pCell.Format.Alignment = wdAlignParagraphCenter
                            pCell.Range.Font.Size = 12
                            pCell.Range.Font.Bold = False
                            pCell.Range.Font.Italic = True
                            pCell.Format.SpaceBefore = 0
                            pCell.Format.SpaceAfter = 0
                            pCell.Format.LineSpacingRule = wdLineSpaceSingle
                        Else
                            pCell.Format.Alignment = wdAlignParagraphCenter
                            pCell.Range.Font.Size = 14
                            
                            If pAsc = "*" Or pAsc = "***" Or IsSeparatorLine(pAsc) Then
                                pCell.Range.Font.Bold = False
                                pCell.Range.Font.Italic = False
                            ElseIf pAsc Like "SO *" Or pAsc Like "SO:*" Or pAsc Like "SO-*" Then
                                Dim origSo As String, fixSo As String
                                origSo = CleanParaText(pCell.Range.Text)
                                fixSo = ChuanHoa_SoKyHieu(origSo)
                                If fixSo <> origSo And Len(fixSo) > 0 Then
                                    Call SafeSetParagraphText(doc, pCell, fixSo)
                                End If
                                pCell.Format.Alignment = wdAlignParagraphCenter
                                pCell.Range.Font.Name = "Times New Roman"
                                pCell.Range.Font.Size = 14
                                pCell.Range.Font.Bold = False
                                pCell.Range.Font.Italic = False
                            Else
                                If leftColCqIdx = 1 Then
                                    pCell.Range.Case = wdUpperCase
                                    pCell.Range.Font.Bold = False
                                    pCell.Range.Font.Italic = False
                                ElseIf leftColCqIdx = 2 Then
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
SkipEmptyLineHeader:
                Next pCell
NextCellHeader:
                cellIdx = cellIdx + 1
            Next c
            
            ' Nếu Ô 1 chỉ có đúng 1 dòng tên cơ quan (không có cấp trên) -> Bắt buộc IN ĐẬM dòng đó
            If leftColCqIdx = 2 Then
                On Error Resume Next
                Dim firstLeftP As Paragraph
                Set firstLeftP = tbl.Range.Cells(1).Range.Paragraphs(1)
                firstLeftP.Range.Font.Bold = True
                On Error GoTo 0
            End If
            
            ' --- TỰ ĐỘNG CHÈN SHAPE LINE DƯỚI "ĐẢNG CỘNG SẢN VIỆT NAM" ---
            If Not (pDCS Is Nothing) Then
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
                        If StrComp(sItem.Name, "LineTieuNguDang", vbTextCompare) = 0 Then shouldDel = True
                        On Error Resume Next
                        If sItem.Anchor.InRange(tbl.Range) Then shouldDel = True
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
                
                Dim iShp As InlineShape
                For Each iShp In tbl.Range.InlineShapes
                    On Error Resume Next
                    iShp.Delete
                    On Error GoTo 0
                Next iShp
                
                pDCS.Range.Font.Underline = wdUnderlineNone
                On Error Resume Next
                pDCS.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                On Error GoTo 0
                
                Dim cRightCell As Cell
                Set cRightCell = Nothing
                On Error Resume Next
                Set cRightCell = pDCS.Range.Cells(1)
                On Error GoTo 0
                If cRightCell Is Nothing Then Set cRightCell = cRight
                
                Dim rPad As Single: rPad = 5.4
                Dim cellW As Single: cellW = 248
                If Not (cRightCell Is Nothing) Then
                    cellW = cRightCell.Width
                    On Error Resume Next
                    rPad = cRightCell.RightPadding
                    On Error GoTo 0
                End If
                
                Dim lineWidth As Single: lineWidth = 190
                Dim lPos As Single: lPos = cellW - lineWidth - rPad - 5.5
                If lPos < 0 Then lPos = 0
                
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
                    .Line.DashStyle = 1
                    .WrapFormat.Type = 3
                End With
                On Error GoTo 0
            End If
            
            Exit For
        End If
    Next tbl
    
    '-----------------------------------------------------------------------
    ' BƯỚC 3.2: DÒ BẮT VÀ CHUẨN HÓA BẢNG KÍNH GỬI (CÔNG VĂN)
    '-----------------------------------------------------------------------
    For Each tbl In doc.Tables
        Dim tAscKG As String
        tAscKG = ToAscii(tbl.Range.Text)
        
        If InStr(1, tAscKG, "KINH GUI") > 0 And InStr(1, tAscKG, "NOI NHAN") = 0 Then
            On Error Resume Next
            tbl.Borders.Enable = False
            On Error GoTo 0
            
            Dim cKG As Cell
            For Each cKG In tbl.Range.Cells
                Dim cellTxtKG As String
                cellTxtKG = CleanParaText(cKG.Range.Text)
                
                If Len(cellTxtKG) > 0 Then
                    Dim cAscKG As String
                    cAscKG = ToAscii(cKG.Range.Text)
                    
                    On Error Resume Next
                    cKG.Borders.Enable = False
                    cKG.Shading.BackgroundPatternColor = wdColorAutomatic
                    cKG.Shading.Texture = wdTextureNone
                    On Error GoTo 0
                    
                    Dim isKGOther As Boolean
                    isKGOther = False
                    If InStr(1, cAscKG, "DANG CONG SAN") > 0 Or InStr(1, cAscKG, "DANG UY") > 0 Or InStr(1, cAscKG, "DANG BO") > 0 Then
                        isKGOther = True
                    End If
                    
                    If Not isKGOther Then
                        If InStr(1, cAscKG, "KINH GUI") > 0 Then
                            Dim pKGLeft As Paragraph
                            For Each pKGLeft In cKG.Range.Paragraphs
                                Do While pKGLeft.Range.Characters.Count > 1 And (pKGLeft.Range.Characters(1).Text = " " Or pKGLeft.Range.Characters(1).Text = vbTab)
                                    pKGLeft.Range.Characters(1).Delete
                                Loop
                                
                                Dim pAscLKG As String
                                pAscLKG = Trim(ToAscii(pKGLeft.Range.Text))
                                If Len(pAscLKG) > 0 Then
                                    With pKGLeft.Format
                                        .Alignment = wdAlignParagraphRight
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
                                    pKGLeft.Range.Font.Italic = True
                                    pKGLeft.Range.Font.ColorIndex = wdBlack
                                    pKGLeft.Range.Font.Underline = wdUnderlineNone
                                End If
                            Next pKGLeft
                        Else
                            Dim lastLineKG As Integer: lastLineKG = 0
                            Dim counterKG As Integer: counterKG = 1
                            Dim tPkg As Paragraph
                            For Each tPkg In cKG.Range.Paragraphs
                                If Len(CleanParaText(tPkg.Range.Text)) > 0 Then lastLineKG = counterKG
                                counterKG = counterKG + 1
                            Next tPkg
                            
                            Dim pKGRight As Paragraph
                            Dim lineKGIdx As Integer: lineKGIdx = 1
                            For Each pKGRight In cKG.Range.Paragraphs
                                Do While pKGRight.Range.Characters.Count > 1 And (pKGRight.Range.Characters(1).Text = " " Or pKGRight.Range.Characters(1).Text = vbTab)
                                    pKGRight.Range.Characters(1).Delete
                                Loop
                                
                                If Len(CleanParaText(pKGRight.Range.Text)) > 0 Then
                                    With pKGRight.Format
                                        .Alignment = wdAlignParagraphLeft
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
                                    pKGRight.Range.Font.Italic = False
                                    pKGRight.Range.Font.ColorIndex = wdBlack
                                    pKGRight.Range.Font.Underline = wdUnderlineNone
                                    
                                    Call NormalizeLineEndingPunctuation(doc, pKGRight, (lineKGIdx = lastLineKG))
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
    ' BƯỚC 3.5: DÒ BẮT BẢNG NƠI NHẬN & CHỮ KÝ (PHÂN BIỆT CHÍNH XÁC 2 CỘT)
    '-----------------------------------------------------------------------
    For Each tbl In doc.Tables
        Dim tAscFooter As String
        tAscFooter = ToAscii(tbl.Range.Text)
        
        If InStr(1, tAscFooter, "NOI NHAN") > 0 Or InStr(1, tAscFooter, "T/M") > 0 Or InStr(1, tAscFooter, "TM.") > 0 Or InStr(1, tAscFooter, "K/T") > 0 Or InStr(1, tAscFooter, "KT.") > 0 Or InStr(1, tAscFooter, "BI THU") > 0 Or InStr(1, tAscFooter, "CHANH VAN PHONG") > 0 Or InStr(1, tAscFooter, "TRUONG BAN") > 0 Then
            ' Không xử lý nhầm bảng Header nếu bảng Header có từ "DANG CONG SAN"
            If InStr(1, tAscFooter, "DANG CONG SAN") > 0 Then GoTo NextTableFooter
            
            On Error Resume Next
            tbl.Borders.Enable = False
            On Error GoTo 0
            
            Dim fCellIdx As Integer: fCellIdx = 1
            Dim fTotalCells As Integer: fTotalCells = tbl.Range.Cells.Count
            
            For Each c In tbl.Range.Cells
                Dim cAscFooter As String
                cAscFooter = Trim(ToAscii(c.Range.Text))
                If Len(cAscFooter) = 0 Then GoTo NextFooterCell
                
                ' Nhận diện chính xác: Nếu ô có chữ "NOI NHAN:" -> Chắc chắn 100% là Cột Nơi nhận!
                ' Ngược lại nếu không có "NOI NHAN:" và nằm ở cột phải / chứa chức danh ký -> Là Cột Chữ ký!
                Dim isSignatureCell As Boolean
                isSignatureCell = False
                
                If InStr(1, cAscFooter, "NOI NHAN") > 0 Then
                    isSignatureCell = False
                ElseIf fCellIdx = fTotalCells And fTotalCells >= 2 Then
                    isSignatureCell = True
                ElseIf InStr(1, cAscFooter, "T/M") > 0 Or InStr(1, cAscFooter, "TM.") > 0 Or InStr(1, cAscFooter, "K/T") > 0 Or InStr(1, cAscFooter, "KT.") > 0 Or InStr(1, cAscFooter, "T/L") > 0 Or InStr(1, cAscFooter, "TL.") > 0 Or InStr(1, cAscFooter, "BI THU") > 0 Or InStr(1, cAscFooter, "CHANH VAN PHONG") > 0 Or InStr(1, cAscFooter, "TRUONG BAN") > 0 Or InStr(1, cAscFooter, "CHU NHIEM") > 0 Then
                    isSignatureCell = True
                End If
                
                If Not isSignatureCell Then
                    ' --- CỘT NƠI NHẬN ---
                    Dim lastListItemIdx As Integer: lastListItemIdx = 0
                    Dim tempP As Paragraph
                    Dim tCounter As Integer: tCounter = 1
                    For Each tempP In c.Range.Paragraphs
                        Dim tVal As String
                        tVal = CleanParaText(tempP.Range.Text)
                        If Len(tVal) > 0 Then
                            If InStr(1, ToAscii(tVal), "NOI NHAN") = 0 Then
                                lastListItemIdx = tCounter
                            End If
                        End If
                        tCounter = tCounter + 1
                    Next tempP
                    
                    Dim pNoiNhan As Paragraph
                    Dim pIndexNN As Integer: pIndexNN = 1
                    For Each pNoiNhan In c.Range.Paragraphs
                        Dim pTextNN As String
                        pTextNN = CleanParaText(pNoiNhan.Range.Text)
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
                            pNoiNhan.Range.Font.Size = 14
                            pNoiNhan.Range.Font.Underline = wdUnderlineSingle
                            pNoiNhan.Range.Font.Bold = False
                            pNoiNhan.Range.Font.Italic = False
                        Else
                            pNoiNhan.Range.Font.Size = 12
                            pNoiNhan.Range.Font.Underline = wdUnderlineNone
                            pNoiNhan.Range.Font.Bold = False
                            pNoiNhan.Range.Font.Italic = False
                            pNoiNhan.Format.LeftIndent = CentimetersToPoints(0.5)
                            pNoiNhan.Format.FirstLineIndent = CentimetersToPoints(-0.5)
                            
                            Dim firstChar As String
                            firstChar = Left(pTextNN, 1)
                            If firstChar <> "-" Then
                                If firstChar = ChrW(8211) Or firstChar = ChrW(8212) Or firstChar = "*" Or firstChar = "+" Or firstChar = ChrW(8226) Then
                                    pNoiNhan.Range.Characters(1).Text = "-"
                                Else
                                    pNoiNhan.Range.InsertBefore "- "
                                End If
                            End If
                            
                            Call NormalizeLineEndingPunctuation(doc, pNoiNhan, (pIndexNN = lastListItemIdx))
                        End If
SkipNNLine:
                        pIndexNN = pIndexNN + 1
                    Next pNoiNhan
                Else
                    ' --- CỘT CHỮ KÝ ---
                    On Error Resume Next
                    Dim textIndices(1 To 30) As Integer
                    Dim textCount As Integer: textCount = 0
                    
                    Dim pScan As Paragraph
                    Dim scanIdx As Integer: scanIdx = 1
                    For Each pScan In c.Range.Paragraphs
                        If Len(CleanParaText(pScan.Range.Text)) > 0 Then
                            If textCount < 30 Then
                                textCount = textCount + 1
                                textIndices(textCount) = scanIdx
                            End If
                        End If
                        scanIdx = scanIdx + 1
                    Next pScan
                    
                    If textCount >= 2 Then
                        Dim sigIdx As Integer: sigIdx = textIndices(textCount)
                        Dim prevTitleIdx As Integer: prevTitleIdx = textIndices(textCount - 1)
                        Dim curEmpty As Integer: curEmpty = sigIdx - prevTitleIdx - 1
                        
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
                    
                    ' Định dạng chuẩn cho các dòng có chữ trong ô Chữ ký
                    Dim pSig As Paragraph
                    Dim nonEmptyOrder As Integer: nonEmptyOrder = 0
                    
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
                        curText = CleanParaText(pSig.Range.Text)
                        
                        If Len(curText) > 0 Then
                            nonEmptyOrder = nonEmptyOrder + 1
                            Dim uText As String
                            uText = ToAscii(curText)
                            
                            If uText Like "T/M *" Or uText Like "TM.*" Or uText Like "TM *" Or _
                               uText Like "K/T *" Or uText Like "KT.*" Or uText Like "KT *" Or _
                               uText Like "T/L *" Or uText Like "TL.*" Or uText Like "TL *" Then
                                Dim fixAuth As String
                                fixAuth = ChuanHoa_ThamQuyen(curText)
                                If fixAuth <> curText And Len(fixAuth) > 0 Then
                                    Call SafeSetParagraphText(doc, pSig, fixAuth)
                                End If
                                pSig.Format.Alignment = wdAlignParagraphCenter
                                pSig.Range.Font.Name = "Times New Roman"
                                pSig.Range.Font.Size = 14
                                pSig.Range.Case = wdUpperCase
                                pSig.Range.Font.Bold = True
                            ElseIf nonEmptyOrder = textCount Then
                                ' Dòng cuối cùng có chữ (Họ và tên người ký): In đậm
                                pSig.Range.Font.Bold = True
                            ElseIf nonEmptyOrder = 2 And textCount >= 3 Then
                                ' Chức vụ cụ thể dòng 2
                                pSig.Range.Case = wdUpperCase
                                pSig.Range.Font.Bold = False
                            Else
                                pSig.Range.Case = wdUpperCase
                                pSig.Range.Font.Bold = True
                            End If
                        Else
                            pSig.Range.Font.Bold = False
                            pSig.Range.Font.Size = 14
                        End If
                    Next pSig
                End If
NextFooterCell:
                fCellIdx = fCellIdx + 1
            Next c
        End If
NextTableFooter:
    Next tbl
    
    '-----------------------------------------------------------------------
    ' BƯỚC 4: ĐỊNH DẠNG THÂN VĂN BẢN THÔNG MINH (BẢO TOÀN IN ĐẬM CỤC BỘ)
    '-----------------------------------------------------------------------
    Dim p As Paragraph
    Dim s As String
    Dim titleState As Integer: titleState = 0
    Dim nonTableTextParaCount As Integer: nonTableTextParaCount = 0
    
    For Each p In doc.Paragraphs
        If p.Range.Information(wdWithInTable) Then
            GoTo NextParagraphMain
        End If
        
        s = CleanParaText(p.Range.Text)
        
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
            p.Range.Font.Size = 15
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.LineSpacingRule = wdLineSpaceSingle
            GoTo NextParagraphMain
        End If
        
        nonTableTextParaCount = nonTableTextParaCount + 1
        p.Range.Font.Name = "Times New Roman"
        p.Range.Font.ColorIndex = wdBlack
        
        Dim sAscPara As String
        sAscPara = Trim(ToAscii(s))
        
        ' --- 1. QUẢN LÝ KHỐI TỰA ĐỀ (CHỈ XÉT TỐI ĐA 2 ĐOẠN ĐẦU NGOÀI BẢNG) ---
        If titleState = 0 Then
            If nonTableTextParaCount <= 2 And IsPartyDocTitle(sAscPara) Then
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
                p.Range.Font.Size = 16
                p.Range.Font.Bold = True
                p.Range.Font.Italic = False
                p.Range.Font.Underline = wdUnderlineNone
                GoTo NextParagraphMain
            Else
                ' Nếu ngay đoạn đầu ngoài bảng không phải Tên loại văn bản (VD: Công văn) -> Khóa ngay vào Thân bài!
                titleState = 4
            End If
        ElseIf titleState = 1 Then
            If IsSeparatorLine(s) Then
                titleState = 4
                Call FormatCenteredSeparator(p)
                GoTo NextParagraphMain
            ElseIf IsBodyStart(s, sAscPara) Then
                titleState = 4
            Else
                titleState = 2
                Call FormatCenteredSummary(p)
                GoTo NextParagraphMain
            End If
        ElseIf titleState = 2 Then
            If IsSeparatorLine(s) Then
                titleState = 4
                Call FormatCenteredSeparator(p)
                GoTo NextParagraphMain
            ElseIf IsBodyStart(s, sAscPara) Or Len(s) > 85 Or Right(s, 1) = "." Or Right(s, 1) = ";" Or Right(s, 1) = ":" Then
                titleState = 4
            Else
                titleState = 3
                Call FormatCenteredSummary(p)
                GoTo NextParagraphMain
            End If
        ElseIf titleState = 3 Then
            If IsSeparatorLine(s) Then
                titleState = 4
                Call FormatCenteredSeparator(p)
                GoTo NextParagraphMain
            Else
                titleState = 4
            End If
        End If
        
        ' --- 2. XỬ LÝ DÒNG ĐẶC BIỆT TRONG THÂN BÀI (QUYẾT ĐỊNH:, PHẦN I, CHƯƠNG I, DÒNG KẾT THÚC -----) ---
        If sAscPara = "QUYET DINH:" Or sAscPara = "QUYET DINH :" Or _
           sAscPara Like "PHAN I*" Or sAscPara Like "PHAN II*" Or sAscPara Like "PHAN III*" Or sAscPara Like "PHAN THU *" Or _
           sAscPara Like "CHUONG I*" Or sAscPara Like "CHUONG II*" Or sAscPara Like "CHUONG III*" Then
            With p.Format
                .Alignment = wdAlignParagraphCenter
                .LeftIndent = 0
                .RightIndent = 0
                .FirstLineIndent = 0
                .SpaceBefore = 6
                .SpaceAfter = 6
                .LineSpacingRule = wdLineSpaceSingle
                .DisableLineHeightGrid = True
            End With
            p.Range.Case = wdUpperCase
            p.Range.Font.Size = 15
            p.Range.Font.Bold = True
            p.Range.Font.Italic = False
            GoTo NextParagraphMain
        ElseIf IsSeparatorLine(s) Then
            Call FormatCenteredSeparator(p)
            GoTo NextParagraphMain
        End If
        
        ' --- 3. XỬ LÝ KÍNH GỬI / NƠI NHẬN NGOÀI BẢNG ---
        If Len(s) < 100 Then
            If sAscPara Like "NOI NHAN:*" Or sAscPara Like "KINH GUI:*" Then
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
        
        ' --- 4. ĐỊNH DẠNG ĐOẠN VĂN THÂN BÀI CHUẨN ---
        p.Range.Font.Name = "Times New Roman"
        p.Range.Font.Size = 15
        p.Range.Font.ColorIndex = wdBlack
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
        
        ' Nhận diện đề mục lớn (I., II., 1., 1.1.) để In Đậm;
        ' KHÔNG ép Bold = False lên các đoạn có in đậm cục bộ (như Điều 1., - Về kinh tế:)
        If sAscPara Like "I.*" Or sAscPara Like "II.*" Or sAscPara Like "III.*" Or sAscPara Like "IV.*" Or sAscPara Like "V.*" Or sAscPara Like "VI.*" Or sAscPara Like "VII.*" Or sAscPara Like "VIII.*" Or _
           sAscPara Like "I-*" Or sAscPara Like "II-*" Or sAscPara Like "III-*" Or sAscPara Like "IV-*" Or sAscPara Like "V-*" Then
            p.Range.Case = wdUpperCase
            p.Range.Font.Bold = True
        ElseIf (s Like "#. *" Or s Like "##. *" Or s Like "#.#. *" Or s Like "#.##. *" Or s Like "##.#. *" Or s Like "#.#.#*" Or Left(s, 1) = "*") And Len(s) < 160 And Right(s, 1) <> "." And Right(s, 1) <> ";" Then
            p.Range.Font.Bold = True
        ElseIf sAscPara Like "DIEU #.*" Or sAscPara Like "DIEU ##.*" Or sAscPara Like "DIEU #:*" Then
            ' Chỉ in đậm chữ "Điều X." ở đầu câu, giữ nguyên phần sau
            Dim colonOrDotPos As Long
            colonOrDotPos = InStr(1, s, ".")
            If colonOrDotPos = 0 Or colonOrDotPos > 10 Then colonOrDotPos = InStr(1, s, ":")
            If colonOrDotPos > 0 And colonOrDotPos <= 10 Then
                Dim rDieu As Range
                Set rDieu = doc.Range(p.Range.Start, p.Range.Start + colonOrDotPos)
                rDieu.Font.Bold = True
            End If
        ElseIf p.Range.Font.Bold = True Then
            ' Chỉ tắt Bold nếu TOÀN BỘ đoạn văn xuôi đang bị bôi đậm cả đoạn
            p.Range.Font.Bold = False
        End If
        
NextParagraphMain:
    Next p
    
    '-----------------------------------------------------------------------
    ' BƯỚC 5: AUTOFIT WINDOW CHO TOÀN BỘ BẢNG
    '-----------------------------------------------------------------------
    Dim tblAll As Table
    For Each tblAll In doc.Tables
        On Error Resume Next
        tblAll.AutoFitBehavior 2
        tblAll.AllowAutoFit = True
        On Error GoTo 0
    Next tblAll
    
    Application.ScreenUpdating = True
    
    '-----------------------------------------------------------------------
    ' BƯỚC 6: RÀ SOÁT LOGIC THỂ THỨC DIỆN RỘNG (10 NHÓM LOGIC CHÉO)
    '-----------------------------------------------------------------------
    Dim strWarnings As String
    strWarnings = KiemTra_LogicTheThuc(doc)
    
    If Len(strWarnings) > 0 Then
        MsgBox "DA CHUAN HOA THE THUC THEO HD 05-HD/VPTW!" & vbCrLf & vbCrLf & _
               "Phat hien cac diem bat hop ly ve LOGIC can kiem tra truoc khi trinh ky:" & vbCrLf & _
               strWarnings & vbCrLf & vbCrLf & _
               "Created by VPDU xa Thanh Phu", vbExclamation, "Bao Cao Ra Soat Logic & The Thuc"
    Else
        MsgBox "Da chuan hoa the thuc hoan hao theo Huong dan 05-HD/VPTW!" & vbCrLf & vbCrLf & _
               "Van ban dat chuan 100%, khong phat hien loi logic." & vbCrLf & vbCrLf & _
               "Created by VPDU xa Thanh Phu", vbInformation, "Chuan Hoa The Thuc Thanh Cong"
    End If
End Sub

'===========================================================================
' CÁC HÀM PHỤ TRỢ AN TOÀN CHO BẢNG & ĐOẠN VĂN
'===========================================================================
Private Function CleanParaText(ByVal rawText As String) As String
    Dim s As String
    s = Replace(Replace(rawText, vbCr, ""), vbLf, "")
    s = Replace(s, Chr(7), "")
    s = Replace(s, Chr(11), "")
    s = Replace(s, Chr(160), " ")
    CleanParaText = Trim(s)
End Function

' Gán text an toàn vào Paragraph mà KHÔNG làm mất ký tự kết thúc ô bảng Chr(7)
Private Sub SafeSetParagraphText(ByVal doc As Document, ByVal p As Paragraph, ByVal newText As String)
    On Error Resume Next
    Dim r As Range
    Set r = p.Range
    Do While r.End > r.Start
        Dim code As Long
        code = AscW(Right(r.Text, 1))
        If code < 0 Then code = code + 65536
        If code = 13 Or code = 10 Or code = 7 Then
            r.End = r.End - 1
        Else
            Exit Do
        End If
    Loop
    r.Text = newText
    On Error GoTo 0
End Sub

Private Sub NormalizeLineEndingPunctuation(ByVal doc As Document, ByVal p As Paragraph, ByVal isLastLine As Boolean)
    On Error Resume Next
    Dim rngLine As Range
    Set rngLine = p.Range
    
    Do While rngLine.End > rngLine.Start
        Dim cCode As Long
        cCode = AscW(Right(rngLine.Text, 1))
        If cCode < 0 Then cCode = cCode + 65536
        If cCode = 13 Or cCode = 10 Or cCode = 7 Then
            rngLine.End = rngLine.End - 1
        Else
            Exit Do
        End If
    Loop
    
    Dim rngEndPos As Long
    rngEndPos = rngLine.End
    
    Do While rngLine.End > rngLine.Start
        Dim endChCode As Long
        endChCode = AscW(Right(rngLine.Text, 1))
        If endChCode < 0 Then endChCode = endChCode + 65536
        If endChCode = 32 Or endChCode = 160 Or endChCode = 9 Or endChCode = 46 Or endChCode = 59 Or endChCode = 44 Then
            rngLine.End = rngLine.End - 1
        Else
            Exit Do
        End If
    Loop
    
    Dim rngTail As Range
    Set rngTail = doc.Range(rngLine.End, rngEndPos)
    If isLastLine Then
        rngTail.Text = "."
    Else
        rngTail.Text = ";"
    End If
    On Error GoTo 0
End Sub

Private Sub FormatCenteredSummary(ByVal p As Paragraph)
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
End Sub

Private Sub FormatCenteredSeparator(ByVal p As Paragraph)
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
    p.Range.Font.Bold = True
    p.Range.Font.Italic = False
    p.Range.Font.Underline = wdUnderlineNone
End Sub

'===========================================================================
' CÁC HÀM TỰ ĐỘNG SỬA SỐ KÝ HIỆU, NGÀY THÁNG, THẨM QUYỀN KÝ
'===========================================================================
Private Function ChuanHoa_SoKyHieu(ByVal sRaw As String) As String
    Dim s As String
    s = Trim(sRaw)
    Dim sAsc As String
    sAsc = ToAscii(s)
    
    If Left(sAsc, 3) = "SO:" Then
        s = "S" & ChrW(7889) & " " & Trim(Mid(s, 4))
    End If
    
    Dim p1 As Long
    p1 = InStr(1, s, " ")
    If p1 > 0 Then
        Dim prefix As String, rest As String
        prefix = Left(s, p1)
        rest = Replace(Trim(Mid(s, p1 + 1)), " ", "")
        rest = Replace(rest, ChrW(8211), "-")
        rest = Replace(rest, ChrW(8212), "-")
        
        Dim dashPos As Long, slashPos As Long
        dashPos = InStr(1, rest, "-")
        slashPos = InStr(1, rest, "/")
        
        Dim numPart As String, codePart As String, orgPart As String
        
        If slashPos > 0 And dashPos > slashPos Then
            numPart = Trim(Left(rest, slashPos - 1))
            codePart = Trim(Mid(rest, slashPos + 1, dashPos - slashPos - 1))
            orgPart = Trim(Mid(rest, dashPos + 1))
            
            If IsNumeric(numPart) And Val(numPart) < 10 And Val(numPart) > 0 And Len(numPart) = 1 Then
                numPart = "0" & numPart
            End If
            s = prefix & numPart & "-" & UCase(codePart) & "/" & UCase(orgPart)
        ElseIf dashPos > 0 And slashPos > dashPos Then
            numPart = Trim(Left(rest, dashPos - 1))
            codePart = Trim(Mid(rest, dashPos + 1, slashPos - dashPos - 1))
            orgPart = Trim(Mid(rest, slashPos + 1))
            
            If IsNumeric(numPart) And Val(numPart) < 10 And Val(numPart) > 0 And Len(numPart) = 1 Then
                numPart = "0" & numPart
            End If
            s = prefix & numPart & "-" & UCase(codePart) & "/" & UCase(orgPart)
        ElseIf dashPos > 0 Then
            numPart = Trim(Left(rest, dashPos - 1))
            If IsNumeric(numPart) And Val(numPart) < 10 And Val(numPart) > 0 And Len(numPart) = 1 Then
                s = prefix & "0" & rest
            End If
        End If
    End If
    
    ChuanHoa_SoKyHieu = s
End Function

Private Function ChuanHoa_NgayThang(ByVal sRaw As String) As String
    Dim s As String
    s = Trim(sRaw)
    Dim sAsc As String
    sAsc = ToAscii(s)
    
    Dim pNgay As Long, pThang As Long, pNam As Long
    pNgay = InStr(1, sAsc, "NGAY")
    pThang = InStr(1, sAsc, "THANG")
    pNam = InStr(1, sAsc, "NAM")
    
    If pNgay > 0 And pThang > pNgay And pNam > pThang Then
        Dim locPart As String, dayPart As String, monthPart As String, yearPart As String
        locPart = Left(s, pNgay - 1)
        dayPart = Trim(Mid(s, pNgay + 4, pThang - (pNgay + 4)))
        monthPart = Trim(Mid(s, pThang + 5, pNam - (pThang + 5)))
        yearPart = Trim(Mid(s, pNam + 3))
        
        If IsNumeric(dayPart) Then
            Dim dVal As Long: dVal = Val(dayPart)
            If dVal < 10 And dVal > 0 And Len(dayPart) = 1 Then dayPart = "0" & dayPart
        End If
        
        If IsNumeric(monthPart) Then
            Dim mVal As Long: mVal = Val(monthPart)
            If (mVal = 1 Or mVal = 2) And Len(monthPart) = 1 Then
                monthPart = "0" & monthPart
            ElseIf mVal >= 3 And mVal <= 9 And Len(monthPart) > 1 And Left(monthPart, 1) = "0" Then
                monthPart = CStr(mVal)
            End If
        End If
        
        Dim strNgay As String, strThang As String, strNam As String
        strNgay = "ng" & ChrW(224) & "y "
        strThang = " th" & ChrW(225) & "ng "
        strNam = " n" & ChrW(259) & "m "
        
        s = locPart & strNgay & dayPart & strThang & monthPart & strNam & yearPart
    End If
    
    ChuanHoa_NgayThang = s
End Function

Private Function ChuanHoa_ThamQuyen(ByVal sRaw As String) As String
    Dim s As String
    s = Trim(sRaw)
    Dim sUp As String
    sUp = ToAscii(s)
    
    If Left(sUp, 3) = "TM." Or Left(sUp, 3) = "TM " Then
        s = "T/M " & Trim(Mid(s, 4))
    ElseIf Left(sUp, 3) = "KT." Or Left(sUp, 3) = "KT " Then
        s = "K/T " & Trim(Mid(s, 4))
    ElseIf Left(sUp, 3) = "TL." Or Left(sUp, 3) = "TL " Then
        s = "T/L " & Trim(Mid(s, 4))
    End If
    
    ChuanHoa_ThamQuyen = s
End Function

'===========================================================================
' BỘ RÀ SOÁT LOGIC & THỂ THỨC DIỆN RỘNG (10 NHÓM KIỂM TRA CHÉO)
'===========================================================================
Private Function KiemTra_LogicTheThuc(ByVal doc As Document) As String
    Dim warnings As String: warnings = ""
    Dim warnCount As Integer: warnCount = 0
    
    Dim docTextFull As String
    docTextFull = doc.Content.Text
    Dim docAscFull As String
    docAscFull = ToAscii(docTextFull)
    
    ' ----------------------------------------------------------------------
    ' 1. KIỂM TRA DẤU BA CHẤM (...) CHƯA ĐIỀN THÔNG TIN
    ' ----------------------------------------------------------------------
    If InStr(1, docTextFull, "...") > 0 Or InStr(1, docTextFull, ChrW(8230)) > 0 Then
        warnCount = warnCount + 1
        warnings = warnings & vbCrLf & warnCount & ". [Dữ liệu trống] Phát hiện còn dấu '...' chưa điền đầy đủ thông tin."
    End If
    
    ' ----------------------------------------------------------------------
    ' 2. BÓC TÁCH CẤU TRÚC TỪ HEADER, BODY VÀ FOOTER
    ' ----------------------------------------------------------------------
    Dim foundSoRaw As String: foundSoRaw = ""
    Dim foundNumVal As String: foundNumVal = ""
    Dim foundCode As String: foundCode = ""
    Dim foundDateRaw As String: foundDateRaw = ""
    Dim foundTitle As String: foundTitle = ""
    Dim foundSuperiorOrgan As String: foundSuperiorOrgan = ""
    Dim foundIssuingOrgan As String: foundIssuingOrgan = ""
    Dim hasVvSummary As Boolean: hasVvSummary = False
    Dim hasKinhGui As Boolean: hasKinhGui = False
    Dim hasQuyetDinhVerb As Boolean: hasQuyetDinhVerb = False
    Dim hasNoiNhan As Boolean: hasNoiNhan = False
    Dim hasLuuVT As Boolean: hasLuuVT = False
    Dim hasNhuDieu As Boolean: hasNhuDieu = False
    Dim foundSignAuth As String: foundSignAuth = ""
    Dim foundSignRole As String: foundSignRole = ""
    
    ' A. Bóc tách chính xác từ Bảng Header (Bảng 1) nếu có
    If doc.Tables.Count >= 1 Then
        Dim tblH As Table
        Set tblH = doc.Tables(1)
        Dim tHAsc As String
        tHAsc = ToAscii(tblH.Range.Text)
        
        If InStr(1, tHAsc, "DANG CONG SAN") > 0 Or InStr(1, tHAsc, "DANG UY") > 0 Or InStr(1, tHAsc, "DANG BO") > 0 Then
            Dim cH As Cell
            Dim orgLinesCount As Integer: orgLinesCount = 0
            For Each cH In tblH.Range.Cells
                Dim pH As Paragraph
                For Each pH In cH.Range.Paragraphs
                    Dim sH As String: sH = CleanParaText(pH.Range.Text)
                    Dim sHAsc As String: sHAsc = Trim(ToAscii(sH))
                    If Len(sHAsc) > 0 Then
                        If sHAsc Like "SO *" Or sHAsc Like "SO:*" Or sHAsc Like "SO-*" Then
                            foundSoRaw = sHAsc
                        ElseIf InStr(1, sHAsc, "NGAY") > 0 And InStr(1, sHAsc, "THANG") > 0 Then
                            foundDateRaw = sHAsc
                        ElseIf sHAsc Like "V/V*" Or sHAsc Like "VE VIEC*" Then
                            hasVvSummary = True
                        ElseIf InStr(1, sHAsc, "DANG CONG SAN") = 0 And Not IsSeparatorLine(sH) And sHAsc <> "*" And sHAsc <> "***" Then
                            orgLinesCount = orgLinesCount + 1
                            If orgLinesCount = 1 Then
                                foundIssuingOrgan = sHAsc
                            ElseIf orgLinesCount = 2 Then
                                foundSuperiorOrgan = foundIssuingOrgan
                                foundIssuingOrgan = sHAsc
                            End If
                        End If
                    End If
                Next pH
            Next cH
        End If
    End If
    
    ' B. Duyệt toàn bộ các đoạn văn để kiểm tra cấu trúc Body, Đề mục, Điều khoản & Chữ ký
    Dim pCount As Long: pCount = doc.Paragraphs.Count
    Dim nonTblIdx As Long: nonTblIdx = 0
    
    Dim romanOrder(1 To 10) As String
    romanOrder(1) = "I": romanOrder(2) = "II": romanOrder(3) = "III": romanOrder(4) = "IV": romanOrder(5) = "V"
    romanOrder(6) = "VI": romanOrder(7) = "VII": romanOrder(8) = "VIII": romanOrder(9) = "IX": romanOrder(10) = "X"
    
    Dim nextExpectedRoman As Integer: nextExpectedRoman = 1
    Dim nextExpectedArabic As Integer: nextExpectedArabic = 1
    Dim nextExpectedDieu As Integer: nextExpectedDieu = 1
    Dim totalDieuCount As Integer: totalDieuCount = 0
    
    Dim idx As Long
    For idx = 1 To pCount
        Dim pItem As Paragraph
        Set pItem = doc.Paragraphs(idx)
        Dim pText As String
        pText = CleanParaText(pItem.Range.Text)
        If Len(pText) = 0 Then GoTo NextParaCheck
        
        Dim pAsc As String
        pAsc = Trim(ToAscii(pText))
        Dim inTbl As Boolean
        inTbl = pItem.Range.Information(wdWithInTable)
        
        If Not inTbl Then nonTblIdx = nonTblIdx + 1
        
        ' Nhận diện Tên loại văn bản (chỉ ở 3 đoạn đầu ngoài bảng)
        If Not inTbl And nonTblIdx <= 3 And Len(foundTitle) = 0 Then
            If IsPartyDocTitle(pAsc) Then
                foundTitle = pAsc
                GoTo NextParaCheck
            End If
        End If
        
        ' Nhận diện Kính gửi & V/v
        If InStr(1, pAsc, "KINH GUI") > 0 Then hasKinhGui = True
        If pAsc Like "V/V*" Or pAsc Like "VE VIEC*" Then hasVvSummary = True
        If pAsc = "QUYET DINH:" Or pAsc = "QUYET DINH :" Then hasQuyetDinhVerb = True
        
        ' Nhận diện Nơi nhận & Chữ ký ở nửa cuối văn bản
        If idx >= pCount - 30 Or inTbl Then
            If InStr(1, pAsc, "NOI NHAN") > 0 Then hasNoiNhan = True
            If InStr(1, pAsc, "LUU VT") > 0 Or InStr(1, pAsc, "LUU: VT") > 0 Or InStr(1, pAsc, "LUU VAN THU") > 0 Or InStr(1, pAsc, "LUU: VAN THU") > 0 Then
                hasLuuVT = True
            End If
            If InStr(1, pAsc, "NHU DIEU") > 0 Then hasNhuDieu = True
            
            If pAsc Like "T/M *" Or pAsc Like "TM.*" Or pAsc Like "K/T *" Or pAsc Like "KT.*" Or pAsc Like "T/L *" Or pAsc Like "TL.*" Then
                foundSignAuth = pAsc
            ElseIf pAsc = "BI THU" Or pAsc Like "PHO BI THU*" Or pAsc = "CHANH VAN PHONG" Or pAsc = "PHO CHANH VAN PHONG" Or pAsc Like "TRUONG BAN*" Or pAsc Like "CHU NHIEM*" Then
                foundSignRole = pAsc
            End If
        End If
        
        ' Kiểm tra thứ tự Đề mục và Điều khoản (chỉ xét ngoài bảng)
        If Not inTbl Then
            ' Kiểm tra Điều 1., Điều 2., Điều 3.
            If pAsc Like "DIEU #.*" Or pAsc Like "DIEU ##.*" Or pAsc Like "DIEU #:*" Then
                Dim dieuNumStr As String
                dieuNumStr = Trim(Mid(pAsc, 5, 3))
                dieuNumStr = Replace(Replace(dieuNumStr, ".", ""), ":", "")
                If IsNumeric(dieuNumStr) Then
                    Dim actualDieu As Integer: actualDieu = Val(dieuNumStr)
                    totalDieuCount = totalDieuCount + 1
                    If actualDieu <> nextExpectedDieu Then
                        warnCount = warnCount + 1
                        warnings = warnings & vbCrLf & warnCount & ". [Thứ tự Điều] Nhảy cóc hoặc sai thứ tự: Kỳ vọng 'Điều " & nextExpectedDieu & "' nhưng gặp 'Điều " & actualDieu & "'."
                    End If
                    nextExpectedDieu = actualDieu + 1
                End If
            End If
            
            ' Kiểm tra mục La Mã: I., II., III. hoặc I-, II-
            Dim rIdx As Integer
            Dim matchedRoman As Boolean: matchedRoman = False
            For rIdx = 1 To 10
                If pAsc Like romanOrder(rIdx) & ". *" Or pAsc Like romanOrder(rIdx) & "- *" Or pAsc Like romanOrder(rIdx) & " - *" Then
                    matchedRoman = True
                    If rIdx <> nextExpectedRoman Then
                        warnCount = warnCount + 1
                        warnings = warnings & vbCrLf & warnCount & ". [Đề mục La Mã] Nhảy cóc thứ tự: Kỳ vọng mục '" & romanOrder(nextExpectedRoman) & "' nhưng gặp mục '" & romanOrder(rIdx) & "'."
                    End If
                    nextExpectedRoman = rIdx + 1
                    nextExpectedArabic = 1 ' Reset bộ đếm mục 1., 2. khi sang mục La Mã mới
                    Exit For
                End If
            Next rIdx
            
            ' Kiểm tra mục số Ả Rập: 1., 2., 3. (Không xét 1.1.)
            If Not matchedRoman Then
                If (pAsc Like "#. *" Or pAsc Like "##. *") And Not (pAsc Like "#.#*") Then
                    Dim dotPos As Long: dotPos = InStr(1, pAsc, ".")
                    If dotPos > 1 Then
                        Dim arabStr As String: arabStr = Left(pAsc, dotPos - 1)
                        If IsNumeric(arabStr) Then
                            Dim actualArab As Integer: actualArab = Val(arabStr)
                            If actualArab <> nextExpectedArabic Then
                                warnCount = warnCount + 1
                                warnings = warnings & vbCrLf & warnCount & ". [Đề mục Số] Nhảy cóc thứ tự khoản: Kỳ vọng mục '" & nextExpectedArabic & ".' nhưng gặp '" & actualArab & ".' (" & Left(pText, 35) & "...)."
                            End If
                            nextExpectedArabic = actualArab + 1
                        End If
                    End If
                End If
            End If
        End If
NextParaCheck:
    Next idx
    
    ' ----------------------------------------------------------------------
    ' 3. KIỂM TRA SỐ, KÝ HIỆU VĂN BẢN (THIẾU SỐ / SAI KÝ HIỆU)
    ' ----------------------------------------------------------------------
    If Len(foundSoRaw) > 0 Then
        Dim dashIdx As Long, slashIdx As Long
        dashIdx = InStr(1, foundSoRaw, "-")
        slashIdx = InStr(1, foundSoRaw, "/")
        If dashIdx > 0 And slashIdx > dashIdx Then
            foundNumVal = Trim(Mid(foundSoRaw, 3, dashIdx - 3))
            foundCode = Trim(Mid(foundSoRaw, dashIdx + 1, slashIdx - dashIdx - 1))
        End If
        
        If Len(foundNumVal) = 0 Or Not IsNumeric(foundNumVal) Then
            warnCount = warnCount + 1
            warnings = warnings & vbCrLf & warnCount & ". [Số văn bản] Chưa điền số thứ tự văn bản ('" & foundSoRaw & "')."
        End If
    Else
        warnCount = warnCount + 1
        warnings = warnings & vbCrLf & warnCount & ". [Số văn bản] Không tìm thấy dòng Số, ký hiệu văn bản ở phần Tiêu đề."
    End If
    
    ' ----------------------------------------------------------------------
    ' 4. KIỂM TRA NGÀY THÁNG NĂM HỢP LỆ (LỊCH THỰC TẾ)
    ' ----------------------------------------------------------------------
    If Len(foundDateRaw) > 0 Then
        Dim pNg As Long, pTh As Long, pNm As Long
        pNg = InStr(1, foundDateRaw, "NGAY")
        pTh = InStr(1, foundDateRaw, "THANG")
        pNm = InStr(1, foundDateRaw, "NAM")
        If pNg > 0 And pTh > pNg And pNm > pTh Then
            Dim sDay As String, sMonth As String, sYear As String
            sDay = Trim(Mid(foundDateRaw, pNg + 4, pTh - (pNg + 4)))
            sMonth = Trim(Mid(foundDateRaw, pTh + 5, pNm - (pTh + 5)))
            sYear = Trim(Left(Trim(Mid(foundDateRaw, pNm + 3)), 4))
            
            If Not IsNumeric(sDay) Or Not IsNumeric(sMonth) Then
                warnCount = warnCount + 1
                warnings = warnings & vbCrLf & warnCount & ". [Ngày tháng] Chưa điền đầy đủ ngày hoặc tháng ban hành ('" & foundDateRaw & "')."
            Else
                Dim dInt As Integer, mInt As Integer, yInt As Integer
                dInt = Val(sDay): mInt = Val(sMonth): yInt = Val(sYear)
                If mInt < 1 Or mInt > 12 Or dInt < 1 Or dInt > 31 Then
                    warnCount = warnCount + 1
                    warnings = warnings & vbCrLf & warnCount & ". [Ngày tháng] Ngày tháng không hợp lệ: ngày " & dInt & " tháng " & mInt & "."
                ElseIf (mInt = 4 Or mInt = 6 Or mInt = 9 Or mInt = 11) And dInt > 30 Then
                    warnCount = warnCount + 1
                    warnings = warnings & vbCrLf & warnCount & ". [Ngày tháng] Tháng " & mInt & " chỉ có tối đa 30 ngày (đang ghi ngày " & dInt & ")."
                ElseIf mInt = 2 And dInt > 29 Then
                    warnCount = warnCount + 1
                    warnings = warnings & vbCrLf & warnCount & ". [Ngày tháng] Tháng 02 không thể có ngày " & dInt & "."
                End If
            End If
        End If
    End If
    
    ' ----------------------------------------------------------------------
    ' 5. ĐỐI CHIẾU CHÉO [MÃ KÝ HIỆU] <-> [TÊN LOẠI VĂN BẢN] <-> [CẤU TRÚC]
    ' ----------------------------------------------------------------------
    If Len(foundCode) > 0 Then
        Dim expectTitle As String: expectTitle = ""
        Select Case foundCode
            Case "BC": expectTitle = "BAO CAO"
            Case "KH": expectTitle = "KE HOACH"
            Case "QD": expectTitle = "QUYET DINH"
            Case "NQ": expectTitle = "NGHI QUYET"
            Case "CT": expectTitle = "CHI THI"
            Case "KL": expectTitle = "KET LUAN"
            Case "TB": expectTitle = "THONG BAO"
            Case "TTR": expectTitle = "TO TRINH"
            Case "CTR": expectTitle = "CHUONG TRINH"
            Case "HD": expectTitle = "HUONG DAN"
            Case "QC": expectTitle = "QUY CHE"
            Case "QYD": expectTitle = "QUY DINH"
            Case "CV": expectTitle = "CONG VAN"
        End Select
        
        If expectTitle = "CONG VAN" Then
            If Len(foundTitle) > 0 Then
                warnCount = warnCount + 1
                warnings = warnings & vbCrLf & warnCount & ". [Loại văn bản] Ký hiệu là Công văn ('CV') nhưng lại có tiêu đề '" & foundTitle & "' ở giữa bài."
            End If
            If Not hasVvSummary Or Not hasKinhGui Then
                warnCount = warnCount + 1
                warnings = warnings & vbCrLf & warnCount & ". [Cấu trúc Công văn] Công văn đang thiếu dòng trích yếu 'V/v...' hoặc thiếu phần 'Kính gửi:'."
            End If
        ElseIf Len(expectTitle) > 0 Then
            If Len(foundTitle) = 0 Then
                warnCount = warnCount + 1
                warnings = warnings & vbCrLf & warnCount & ". [Thiếu tiêu đề] Ký hiệu văn bản là '" & foundCode & "' nhưng thiếu dòng tên loại văn bản ('" & expectTitle & "') ở giữa trang."
            ElseIf foundTitle <> expectTitle Then
                warnCount = warnCount + 1
                warnings = warnings & vbCrLf & warnCount & ". [Mâu thuẫn loại VB] Ký hiệu ghi là '" & foundCode & "' nhưng tiêu đề văn bản lại ghi là '" & foundTitle & "'."
            End If
        End If
    End If
    
    ' ----------------------------------------------------------------------
    ' 6. KIỂM TRA LOGIC CẤU TRÚC VĂN BẢN "QUYẾT ĐỊNH"
    ' ----------------------------------------------------------------------
    If foundTitle = "QUYET DINH" Or foundCode = "QD" Then
        If Not hasQuyetDinhVerb Or totalDieuCount = 0 Then
            warnCount = warnCount + 1
            warnings = warnings & vbCrLf & warnCount & ". [Cấu trúc Quyết định] Văn bản Quyết định nhưng thiếu dòng 'QUYẾT ĐỊNH:' hoặc thiếu các 'Điều 1.', 'Điều 2.'."
        End If
        If totalDieuCount > 0 And hasNoiNhan And Not hasNhuDieu Then
            warnCount = warnCount + 1
            warnings = warnings & vbCrLf & warnCount & ". [Logic Nơi nhận] Văn bản Quyết định có Điều khoản thi hành nhưng phần Nơi nhận thiếu dòng '- Như Điều ...;'."
        End If
    End If
    
    ' ----------------------------------------------------------------------
    ' 7. ĐỐI CHIẾU CHÉO [CƠ QUAN BAN HÀNH (Ô 1)] <-> [THẨM QUYỀN KÝ (CHÂN TRANG)]
    ' ----------------------------------------------------------------------
    If Len(foundIssuingOrgan) > 0 Then
        Dim isExecutive As Boolean: isExecutive = False
        If foundIssuingOrgan Like "VAN PHONG*" Or foundIssuingOrgan Like "BAN TO CHUC*" Or _
           foundIssuingOrgan Like "BAN TUYEN GIAO*" Or foundIssuingOrgan Like "BAN XAY DUNG*" Or _
           foundIssuingOrgan Like "BAN DAN VAN*" Or foundIssuingOrgan Like "TRUNG TAM*" Then
            isExecutive = True
        End If
        
        Dim isCollective As Boolean: isCollective = False
        If Not isExecutive And (InStr(1, foundIssuingOrgan, "DANG UY") > 0 Or InStr(1, foundIssuingOrgan, "BAN CHAP HANH") > 0 Or InStr(1, foundIssuingOrgan, "BAN THUONG VU") > 0 Or InStr(1, foundIssuingOrgan, "UY BAN KIEM TRA") > 0 Or InStr(1, foundIssuingOrgan, "CHI BO") > 0) Then
            isCollective = True
        End If
        
        If isCollective And Not (foundSignAuth Like "T/M *" Or foundSignAuth Like "TM.*") Then
            warnCount = warnCount + 1
            warnings = warnings & vbCrLf & warnCount & ". [Thẩm quyền ký] Cơ quan ban hành ở Ô 1 là tập thể ('" & foundIssuingOrgan & "') nhưng chữ ký thiếu thẩm quyền 'T/M ...'."
        ElseIf isExecutive And (foundSignAuth Like "T/M *" Or foundSignAuth Like "TM.*") Then
            warnCount = warnCount + 1
            warnings = warnings & vbCrLf & warnCount & ". [Thẩm quyền ký] Cơ quan ban hành ở Ô 1 là '" & foundIssuingOrgan & "' (chế độ thủ trưởng) nên không được ký 'T/M' (" & foundSignAuth & ")."
        End If
    End If
    
    ' ----------------------------------------------------------------------
    ' 8. KIỂM TRA DÒNG LƯU VĂN THƯ TRONG NƠI NHẬN
    ' ----------------------------------------------------------------------
    If hasNoiNhan And Not hasLuuVT Then
        warnCount = warnCount + 1
        warnings = warnings & vbCrLf & warnCount & ". [Nơi nhận] Phần 'Nơi nhận:' đang thiếu dòng '- Lưu: VT...' ở cuối."
    End If
    
    KiemTra_LogicTheThuc = warnings
End Function

'===========================================================================
' HÀM CHUYỂN ĐỔI ASCII AN TOÀN (ĐÃ BỔ SUNG GIỮ DẤU '.', ':', ';')
'===========================================================================
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
            Case 8211, 8212 ' En-dash, Em-dash -> '-'
                res = res & "-"
            Case Else
                If w >= 97 And w <= 122 Then
                    res = res & Chr(w - 32)
                ElseIf (w >= 65 And w <= 90) Or (w >= 48 And w <= 57) Or _
                       w = 32 Or w = 42 Or w = 45 Or w = 46 Or w = 47 Or w = 58 Or w = 59 Then
                    ' 42=*, 45=-, 46=., 47=/, 58=:, 59=;
                    res = res & Chr(w)
                Else
                    res = res & " "
                End If
        End Select
    Next k
    ToAscii = res
End Function


'-----------------------------------------------------------------------
' HÀM NHẬN DIỆN TÊN LOẠI VĂN BẢN ĐẢNG (KHỚP CHÍNH XÁC, KHÔNG BẮT NHẦM THÂN BÀI)
'-----------------------------------------------------------------------
Private Function IsPartyDocTitle(ByVal sAsc As String) As Boolean
    Dim t As String
    t = Trim(sAsc)
    If Len(t) = 0 Or Len(t) > 35 Then
        IsPartyDocTitle = False
        Exit Function
    End If
    
    Select Case t
        Case "NGHI QUYET", "QUYET DINH", "CHI THI", "KET LUAN", "QUY CHE", _
             "QUY DINH", "THONG TRI", "HUONG DAN", "THONG BAO", "KE HOACH", _
             "CHUONG TRINH", "BAO CAO", "BIEN BAN", "TO TRINH", "CONG DIEN", _
             "DE AN", "PHUONG AN", "GIAY MOI", "GIAY GIOI THIEU", _
             "BAN CAM KET", "BAN DANG KY", "DIEU LE", "CHUONG TRINH HANH DONG"
            IsPartyDocTitle = True
        Case Else
            IsPartyDocTitle = False
    End Select
End Function

Private Function IsSeparatorLine(ByVal rawText As String) As Boolean
    Dim cleanStr As String
    cleanStr = Trim(Replace(Replace(Replace(rawText, vbCr, ""), vbLf, ""), " ", ""))
    cleanStr = Replace(cleanStr, Chr(160), "")
    cleanStr = Replace(cleanStr, vbTab, "")
    
    If Len(cleanStr) < 2 Or Len(cleanStr) > 40 Then
        IsSeparatorLine = False
        Exit Function
    End If
    
    Dim isAllDash As Boolean: isAllDash = True
    Dim i As Long
    For i = 1 To Len(cleanStr)
        Dim wCode As Long
        wCode = AscW(Mid(cleanStr, i, 1))
        If wCode < 0 Then wCode = wCode + 65536
        If wCode <> 45 And wCode <> 95 And wCode <> 42 And _
           wCode <> 8211 And wCode <> 8212 And wCode <> 8213 Then
            isAllDash = False
            Exit For
        End If
    Next i
    
    IsSeparatorLine = isAllDash
End Function

Private Function IsBodyStart(ByVal rawText As String, ByVal sAsc As String) As Boolean
    Dim tAsc As String: tAsc = Trim(sAsc)
    Dim tRaw As String: tRaw = Trim(rawText)
    
    If Len(tAsc) = 0 Then
        IsBodyStart = False
        Exit Function
    End If
    
    Dim match As Boolean: match = False
    
    If tAsc Like "CAN CU*" Or tAsc Like "THUC HIEN*" Or tAsc Like "DE THUC HIEN*" Or _
       tAsc Like "XET DE NGHI*" Or tAsc Like "THEO DE NGHI*" Or _
       tAsc Like "BAN THUONG VU*" Or tAsc Like "BAN CHAP HANH*" Or _
       tAsc Like "DANG UY*" Or tAsc Like "CHI BO*" Or _
       tAsc Like "HOI NGHI*" Or tAsc Like "DAI HOI*" Or _
       tAsc Like "QUYET DINH:*" Or tAsc Like "QUYET DINH :*" Or _
       tAsc Like "YEU CAU:*" Or tAsc Like "YEU CAU :*" Then
        match = True
    End If
    
    If Not match Then
        If tAsc Like "I.*" Or tAsc Like "II.*" Or tAsc Like "III.*" Or _
           tAsc Like "I-*" Or tAsc Like "II-*" Or tAsc Like "III-*" Or _
           tAsc Like "PHAN I*" Or tAsc Like "PHAN 1*" Or tAsc Like "PHAN THU*" Or _
           tAsc Like "MUC I*" Or tAsc Like "MUC 1*" Or _
           tAsc Like "CHUONG I*" Or tAsc Like "CHUONG 1*" Or _
           tAsc Like "DIEU 1*" Or tAsc Like "DIEU 2*" Or tAsc Like "DIEU *" Or _
           tAsc Like "1.*" Or tAsc Like "2.*" Or tAsc Like "A.*" Or tAsc Like "B.*" Then
            match = True
        ElseIf Left(tRaw, 1) = "-" Or Left(tRaw, 1) = "+" Or Left(tRaw, 1) = "*" Then
            match = True
        End If
    End If
    
    If Not match Then
        Dim lastCh As String: lastCh = Right(tRaw, 1)
        If lastCh = ";" Or lastCh = ":" Or Len(tRaw) > 85 Then
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
