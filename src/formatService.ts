/**
 * formatService.ts
 * Module chuẩn hóa văn bản Word theo Hướng dẫn số 05-HD/VPTW ngày 27/5/2026
 * của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng.
 */

import {
  MM_TO_PT,
  mmToPt,
  ptToMm,
  PAGE_SETUP_CONSTANTS,
  TYPOGRAPHY_CONSTANTS,
  HEADER_CONSTANTS,
  FOOTER_CONSTANTS,
  PAGE_NUMBERING_CONSTANTS,
} from "./constants";

import {
  HeaderData,
  SignData,
  FormatBodyOptions,
  PageSetupOptions,
  LintReport,
  LintIssue,
  LintIssueType,
  LintSeverity,
} from "./types";

/**
 * Chuẩn hóa số thứ tự văn bản (bắt buộc có số 0 đằng trước nếu số < 10)
 * Ví dụ: 5 -> "05", "5" -> "05", "12" -> "12"
 */
export function formatDocNumber(num: number | string): string {
  if (typeof num === "number") {
    return num < 10 && num > 0 ? `0${num}` : `${num}`;
  }
  const trimmed = num.trim();
  const parsed = parseInt(trimmed, 10);
  if (!isNaN(parsed) && parsed < 10 && parsed > 0 && !trimmed.startsWith("0")) {
    return `0${parsed}`;
  }
  return trimmed;
}

/**
 * Chuẩn hóa ngày tháng năm theo Hướng dẫn 05-HD/VPTW:
 * - Ngày dưới 10: phải ghi thêm số 0 đằng trước (ví dụ: ngày 05).
 * - Tháng 1, tháng 2: phải ghi thêm số 0 đằng trước (ví dụ: tháng 01, tháng 02).
 * - Tháng 3 đến 12: ghi số bình thường (ví dụ: tháng 3, tháng 12).
 *
 * @param location Địa danh ban hành (ví dụ: "Hà Nội", "Nghệ An")
 * @param dateInput Đối tượng Date, object { day, month, year } hoặc chuỗi có sẵn
 * @returns Chuỗi ngày tháng năm chuẩn hóa (ví dụ: "Hà Nội, ngày 27 tháng 5 năm 2026")
 */
export function formatPartyDate(
  location: string,
  dateInput: Date | { day: number; month: number; year: number } | string
): string {
  const cleanLocation = location.trim().replace(/,\s*$/, "");

  if (typeof dateInput === "string") {
    // Nếu truyền chuỗi text, tiến hành quét và chuẩn hóa số 0 nếu người dùng gõ thiếu
    let formattedText = dateInput.trim();
    // Bổ sung số 0 cho ngày < 10
    formattedText = formattedText.replace(
      /ngày\s+([1-9])(?!\d)\s+tháng/i,
      "ngày 0$1 tháng"
    );
    // Bổ sung số 0 cho tháng 1, 2
    formattedText = formattedText.replace(
      /tháng\s+([1-2])(?!\d)\s+năm/i,
      "tháng 0$1 năm"
    );
    if (!formattedText.toLowerCase().includes(cleanLocation.toLowerCase())) {
      return `${cleanLocation}, ${formattedText}`;
    }
    return formattedText;
  }

  let day: number;
  let month: number;
  let year: number;

  if (dateInput instanceof Date) {
    day = dateInput.getDate();
    month = dateInput.getMonth() + 1;
    year = dateInput.getFullYear();
  } else {
    day = dateInput.day;
    month = dateInput.month;
    year = dateInput.year;
  }

  const dayStr = day < 10 ? `0${day}` : `${day}`;
  const monthStr = month < 3 ? `0${month}` : `${month}`;

  return `${cleanLocation}, ngày ${dayStr} tháng ${monthStr} năm ${year}`;
}

/**
 * Tạo Flat OPC XML chèn trường số trang động (Field PAGE)
 * căn giữa, font Times New Roman, cỡ 13-14 pt
 */
function createPageNumberFlatOpcXml(fontSizePt: number = 13.5): string {
  // OOXML sz đo bằng nửa point (half-points: 13.5 pt = 27)
  const halfPoints = Math.round(fontSizePt * 2);
  return `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:pPr>
              <w:jc w:val="center"/>
              <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
              <w:rPr>
                <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
                <w:color w:val="000000"/>
                <w:sz w:val="${halfPoints}"/>
                <w:szCs w:val="${halfPoints}"/>
              </w:rPr>
            </w:pPr>
            <w:fldSimple w:instr="PAGE"/>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

/**
 * 1. Hàm thiết lập khổ giấy và căn lề (Page Setup) theo Hướng dẫn 05-HD/VPTW
 * - Khổ giấy: A4 (Portrait: 210 x 297 mm)
 * - Lề trên: 20 mm (~56.69 pt)
 * - Lề dưới: 20 mm (~56.69 pt)
 * - Lề trái: 30 mm (~85.04 pt)
 * - Lề phải: 15 mm (~42.52 pt)
 * - Bật in 2 mặt (mirrorMargins = true)
 * - Đánh số trang đỉnh giữa (Header Center), cách mép trên 10 mm (~28.35 pt)
 * - Không đánh số ở trang đầu (differentFirstPageHeaderFooter = true)
 */
export async function applyPageSetup(
  context: Word.RequestContext,
  options?: PageSetupOptions
): Promise<void> {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  const topPt = options?.topMarginMm
    ? mmToPt(options.topMarginMm)
    : PAGE_SETUP_CONSTANTS.MARGIN_TOP_PT;
  const bottomPt = options?.bottomMarginMm
    ? mmToPt(options.bottomMarginMm)
    : PAGE_SETUP_CONSTANTS.MARGIN_BOTTOM_PT;
  const leftPt = options?.leftMarginMm
    ? mmToPt(options.leftMarginMm)
    : PAGE_SETUP_CONSTANTS.MARGIN_LEFT_PT;
  const rightPt = options?.rightMarginMm
    ? mmToPt(options.rightMarginMm)
    : PAGE_SETUP_CONSTANTS.MARGIN_RIGHT_PT;
  const headerDistPt = options?.headerDistanceMm
    ? mmToPt(options.headerDistanceMm)
    : PAGE_SETUP_CONSTANTS.HEADER_DISTANCE_PT;
  const mirror = options?.mirrorMargins ?? PAGE_SETUP_CONSTANTS.MIRROR_MARGINS;
  const diffFirst =
    options?.differentFirstPageHeaderFooter ??
    PAGE_SETUP_CONSTANTS.DIFFERENT_FIRST_PAGE;

  const pageNumberXml = createPageNumberFlatOpcXml(
    PAGE_NUMBERING_CONSTANTS.FONT_SIZE
  );

  for (let i = 0; i < sections.items.length; i++) {
    const section = sections.items[i];
    const pageSetup = section.pageSetup;

    // Cấu hình kích thước A4 tiêu chuẩn và hướng dọc (Portrait)
    pageSetup.pageWidth = PAGE_SETUP_CONSTANTS.A4_WIDTH_PT;
    pageSetup.pageHeight = PAGE_SETUP_CONSTANTS.A4_HEIGHT_PT;
    pageSetup.orientation = Word.PageOrientation.portrait;

    // Cấu hình lề 4 hướng
    pageSetup.topMargin = topPt;
    pageSetup.bottomMargin = bottomPt;
    pageSetup.leftMargin = leftPt;
    pageSetup.rightMargin = rightPt;

    // Khoảng cách từ mép trên đến Header để đánh số trang (10 mm)
    pageSetup.headerDistance = headerDistPt;
    pageSetup.footerDistance = PAGE_SETUP_CONSTANTS.FOOTER_DISTANCE_PT;

    // Bật khác biệt trang đầu (không hiển thị số trang ở trang bìa/đầu)
    pageSetup.differentFirstPageHeaderFooter = diffFirst;

    // In 2 mặt (Mirror Margins)
    try {
      (pageSetup as any).mirrorMargins = mirror;
    } catch {
      // Bỏ qua nếu phiên bản API không hỗ trợ setter trực tiếp trên section
    }

    // Cấu hình đánh số trang ở giữa đỉnh trang trong Header
    try {
      // 1. Primary Header (dùng cho các trang chính)
      const primaryHeader = section.getHeader(Word.HeaderFooterType.primary);
      primaryHeader.insertOoxml(pageNumberXml, Word.InsertLocation.replace);

      // 2. First Page Header: xóa sạch để trang đầu không có số trang
      if (diffFirst) {
        const firstHeader = section.getHeader(Word.HeaderFooterType.firstPage);
        firstHeader.clear();
      }

      // 3. Even Pages Header: nếu in 2 mặt đối xứng, trang chẵn cũng có số trang giữa đỉnh
      if (mirror) {
        try {
          const evenHeader = section.getHeader(Word.HeaderFooterType.evenPages);
          evenHeader.insertOoxml(pageNumberXml, Word.InsertLocation.replace);
        } catch {
          // Trường hợp host không kích hoạt oddAndEvenPagesHeaderFooter
        }
      }
    } catch {
      // Fallback: nếu chèn OOXML thất bại do quyền hạn môi trường, sử dụng API Paragraph thông thường
      const primaryHeader = section.getHeader(Word.HeaderFooterType.primary);
      primaryHeader.clear();
      const p = primaryHeader.insertParagraph("", Word.InsertLocation.start);
      p.alignment = Word.Alignment.centered;
      p.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
      p.font.size = PAGE_NUMBERING_CONSTANTS.FONT_SIZE;
      p.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
    }
  }

  // Thử kích hoạt thuộc tính mirrorMargins tại cấp Document nếu có
  try {
    const docPageSetup = (context.document as any).pageSetup;
    if (docPageSetup) {
      docPageSetup.mirrorMargins = mirror;
      docPageSetup.oddAndEvenPagesHeaderFooter = mirror;
    }
  } catch {
    // Không bắt buộc nếu phiên bản WordApi Desktop thấp hơn 1.3
  }

  await context.sync();
}

/**
 * Kiểm tra xem một đoạn văn bản có phải là tiêu đề/tiêu ngữ căn giữa cần giữ nguyên
 * (ví dụ: QUYẾT ĐỊNH, CHỈ THỊ, KẾ HOẠCH, BÁO CÁO, CHƯƠNG I...)
 */
function isDocumentHeading(text: string): boolean {
  const clean = text.trim();
  if (!clean) return false;

  const headingKeywords = [
    /^QUYẾT ĐỊNH$/i,
    /^CHỈ THỊ$/i,
    /^KẾ HOẠCH$/i,
    /^BÁO CÁO$/i,
    /^QUY ĐỊNH$/i,
    /^HƯỚNG DẪN$/i,
    /^NGHỊ QUYẾT$/i,
    /^KẾT LUẬN$/i,
    /^THÔNG BÁO$/i,
    /^CHƯƠNG\s+[IVXLCDM\d]+/i,
    /^MỤC\s+[\d]+/i,
    /^ĐIỀU\s+[\d]+/i,
  ];

  return headingKeywords.some((regex) => regex.test(clean));
}

/**
 * 2. Hàm chuẩn hóa nội dung văn bản (Format Body Content) theo Hướng dẫn 05-HD/VPTW
 * - Font chữ: Times New Roman, màu #000000
 * - Cỡ chữ nội dung: 14 pt (hoặc 13 pt)
 * - Căn lề: Căn đều 2 bên (Justified)
 * - Thụt lề dòng đầu (First Line Indent): Đúng 10 mm (~28.35 pt)
 * - Giãn dòng: Chế độ Exactly 18 pt (hoặc từ 18-22 pt)
 * - Giãn đoạn: Space Before = 0 pt, Space After = 6 pt (tối thiểu)
 * - Bảo vệ bảng biểu: Không làm biến dạng các bảng đã có trong tài liệu.
 */
export async function formatBodyContent(
  context: Word.RequestContext,
  fontSize?: number,
  options?: FormatBodyOptions
): Promise<void> {
  const targetFontSize =
    fontSize ??
    options?.fontSize ??
    TYPOGRAPHY_CONSTANTS.BODY_FONT_SIZE_DEFAULT;

  const targetLineSpacing =
    options?.lineSpacing ?? TYPOGRAPHY_CONSTANTS.LINE_SPACING_DEFAULT_PT;
  const targetSpaceBefore =
    options?.spaceBefore ?? TYPOGRAPHY_CONSTANTS.SPACE_BEFORE_PT;
  const targetSpaceAfter =
    options?.spaceAfter ?? TYPOGRAPHY_CONSTANTS.SPACE_AFTER_DEFAULT_PT;
  const targetIndentPt = options?.firstLineIndentMm
    ? mmToPt(options.firstLineIndentMm)
    : TYPOGRAPHY_CONSTANTS.FIRST_LINE_INDENT_PT;
  const preserveHeadings = options?.preserveCenteredHeadings ?? true;

  const paragraphs = context.document.body.paragraphs;
  paragraphs.load(["items/text", "items/tableNestingLevel", "items/alignment"]);
  await context.sync();

  for (let i = 0; i < paragraphs.items.length; i++) {
    const p = paragraphs.items[i];

    // QUAN TRỌNG: Bỏ qua đoạn văn nằm trong bảng biểu (như Bảng Header, Bảng Chữ ký, Thống kê)
    // để tránh làm sai lệch thụt lề và căn lề trong ô bảng.
    if (p.tableNestingLevel > 0) {
      continue;
    }

    const trimmedText = p.text.trim();
    if (!trimmedText) {
      // Dòng trắng trống: reset giãn cách để tài liệu gọn gàng
      p.spaceBefore = 0;
      p.spaceAfter = 0;
      p.firstLineIndent = 0;
      continue;
    }

    // Luôn áp dụng phông chữ và màu mực tiêu chuẩn của Đảng
    p.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
    p.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;

    // Kiểm tra xem đoạn văn có phải là tiêu đề chính căn giữa hay không
    const isHeading =
      preserveHeadings &&
      (p.alignment === Word.Alignment.centered ||
        isDocumentHeading(trimmedText));

    if (isHeading) {
      p.alignment = Word.Alignment.centered;
      p.firstLineIndent = 0;
      p.font.size = targetFontSize;
      p.spaceBefore = 6;
      p.spaceAfter = 6;
      p.lineSpacing = targetLineSpacing;
    } else {
      // Đoạn văn nội dung thông thường
      p.font.size = targetFontSize;
      p.alignment = Word.Alignment.justified;
      p.firstLineIndent = targetIndentPt;
      p.spaceBefore = targetSpaceBefore;
      p.spaceAfter = targetSpaceAfter;
      p.lineSpacing = targetLineSpacing;
    }
  }

  await context.sync();
}

/**
 * 3. Hàm chèn bảng Header 2 cột không viền (Ô 1 & Ô 2) theo Hướng dẫn 05-HD/VPTW
 *
 * Ô 1 (Cột trái):
 * - Dòng 1: Cấp trên trực tiếp (13-14 pt, in hoa, đứng) - nếu có.
 * - Dòng 2: Tên cơ quan ban hành (13-14 pt, in hoa, ĐẬM).
 * - Dòng kẻ ngang dưới cơ quan ban hành: nét liền, dài 1/3 đến 1/2 chiều dài chữ.
 * - Dòng 3: Số, ký hiệu văn bản (13 pt, đứng). Tự động có số 0 nếu < 10 (ví dụ: "Số 05-HD/VPTW").
 *
 * Ô 2 (Cột phải):
 * - Dòng 1: "ĐẢNG CỘNG SẢN VIỆT NAM" (13-14 pt, in hoa, ĐẬM).
 * - Dòng kẻ ngang dưới: nét liền có độ dài bằng độ dài dòng chữ.
 * - Dòng 2: Địa danh, ngày tháng năm (13-14 pt, chữ thường, NGHIÊNG).
 *   Quy chuẩn: Ngày < 10 có số 0; Tháng 1, tháng 2 có số 0.
 */
export async function insertHeaderTable(
  context: Word.RequestContext,
  data: HeaderData
): Promise<Word.Table> {
  const body = context.document.body;
  const headerFontSize = data.fontSize ?? HEADER_CONSTANTS.ISSUING_ORGAN_FONT_SIZE;

  // Chuẩn hóa số văn bản (bắt buộc số 0 nếu < 10)
  const formattedNumber = formatDocNumber(data.docNumber);
  const cleanDocCode = data.docCode.trim().replace(/^[-/]+/, "");
  const docNumberString = `Số ${formattedNumber}-${cleanDocCode}`;

  // Chuẩn hóa địa danh và ngày tháng năm
  const dateString = formatPartyDate(data.location, data.date);

  // Chèn bảng 1 hàng x 2 cột vào đầu văn bản
  const table = body.insertTable(1, 2, Word.InsertLocation.start, [["", ""]]);

  // Xóa đường viền toàn bảng (Border = none)
  try {
    table.getBorder(Word.BorderLocation.all).type = Word.BorderType.none;
  } catch {
    // Dự phòng
  }

  // Thiết lập lề ô bằng 0 để chữ bám thẳng lề trang
  try {
    table.setCellPadding(Word.CellPaddingLocation.top, 0);
    table.setCellPadding(Word.CellPaddingLocation.bottom, 0);
    table.setCellPadding(Word.CellPaddingLocation.left, 0);
    table.setCellPadding(Word.CellPaddingLocation.right, 0);
  } catch {
    // Dự phòng
  }

  const cellLeft = table.getCell(0, 0);
  const cellRight = table.getCell(0, 1);

  // Cấu hình độ rộng cột và căn đỉnh cho Ô 1 (~40%) và Ô 2 (~60%)
  try {
    cellLeft.columnWidth = mmToPt(68);
    cellRight.columnWidth = mmToPt(97);
    cellLeft.verticalAlignment = Word.VerticalAlignment.top;
    cellRight.verticalAlignment = Word.VerticalAlignment.top;
  } catch {
    // Dự phòng
  }

  // ----------------------------------------------------
  // Xử lý Ô 1 (Cột trái)
  // ----------------------------------------------------
  const leftFirstP = cellLeft.body.paragraphs.getFirst();
  let currentLeftP = leftFirstP;

  if (data.superiorOrgan && data.superiorOrgan.trim()) {
    // Dòng 1: Cơ quan cấp trên (13-14 pt, HOA đứng)
    currentLeftP.insertText(
      data.superiorOrgan.trim().toUpperCase(),
      Word.InsertLocation.replace
    );
    currentLeftP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
    currentLeftP.font.size = headerFontSize;
    currentLeftP.font.bold = false;
    currentLeftP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
    currentLeftP.alignment = Word.Alignment.centered;
    currentLeftP.firstLineIndent = 0;
    currentLeftP.spaceBefore = 0;
    currentLeftP.spaceAfter = 1;
    currentLeftP.lineSpacing = 15;

    // Dòng 2: Tên cơ quan ban hành (13-14 pt, HOA ĐẬM)
    currentLeftP = currentLeftP.insertParagraph(
      data.issuingOrgan.trim().toUpperCase(),
      Word.InsertLocation.after
    );
  } else {
    // Chỉ có cơ quan ban hành
    currentLeftP.insertText(
      data.issuingOrgan.trim().toUpperCase(),
      Word.InsertLocation.replace
    );
  }

  currentLeftP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
  currentLeftP.font.size = headerFontSize;
  currentLeftP.font.bold = true;
  currentLeftP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
  currentLeftP.alignment = Word.Alignment.centered;
  currentLeftP.firstLineIndent = 0;
  currentLeftP.spaceBefore = 0;
  currentLeftP.spaceAfter = 1;
  currentLeftP.lineSpacing = 15;

  // Dòng kẻ ngang dưới cơ quan ban hành (độ dài 1/3 đến 1/2)
  const linePLeft = currentLeftP.insertParagraph(
    HEADER_CONSTANTS.LINE_ISSUING_ORGAN,
    Word.InsertLocation.after
  );
  linePLeft.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
  linePLeft.font.size = 8;
  linePLeft.font.bold = false;
  linePLeft.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
  linePLeft.alignment = Word.Alignment.centered;
  linePLeft.firstLineIndent = 0;
  linePLeft.spaceBefore = 0;
  linePLeft.spaceAfter = 3;
  linePLeft.lineSpacing = 8;

  // Dòng Số, ký hiệu văn bản (13 pt, đứng)
  const numberP = linePLeft.insertParagraph(
    docNumberString,
    Word.InsertLocation.after
  );
  numberP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
  numberP.font.size = HEADER_CONSTANTS.DOC_NUMBER_FONT_SIZE;
  numberP.font.bold = false;
  numberP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
  numberP.alignment = Word.Alignment.centered;
  numberP.firstLineIndent = 0;
  numberP.spaceBefore = 0;
  numberP.spaceAfter = 0;
  numberP.lineSpacing = 15;

  // ----------------------------------------------------
  // Xử lý Ô 2 (Cột phải)
  // ----------------------------------------------------
  // Dòng 1: "ĐẢNG CỘNG SẢN VIỆT NAM" (13-14 pt, HOA ĐẬM)
  const rightFirstP = cellRight.body.paragraphs.getFirst();
  rightFirstP.insertText(
    HEADER_CONSTANTS.PARTY_HEADER_TEXT,
    Word.InsertLocation.replace
  );
  rightFirstP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
  rightFirstP.font.size = headerFontSize;
  rightFirstP.font.bold = true;
  rightFirstP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
  rightFirstP.alignment = Word.Alignment.centered;
  rightFirstP.firstLineIndent = 0;
  rightFirstP.spaceBefore = 0;
  rightFirstP.spaceAfter = 1;
  rightFirstP.lineSpacing = 15;

  // Dòng kẻ ngang dưới tiêu đề Đảng (độ dài liền bằng dòng chữ)
  const linePRight = rightFirstP.insertParagraph(
    HEADER_CONSTANTS.LINE_PARTY_TITLE,
    Word.InsertLocation.after
  );
  linePRight.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
  linePRight.font.size = 8;
  linePRight.font.bold = false;
  linePRight.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
  linePRight.alignment = Word.Alignment.centered;
  linePRight.firstLineIndent = 0;
  linePRight.spaceBefore = 0;
  linePRight.spaceAfter = 3;
  linePRight.lineSpacing = 8;

  // Dòng Địa danh, ngày tháng năm (13-14 pt, chữ thường, NGHIÊNG)
  const dateP = linePRight.insertParagraph(
    dateString,
    Word.InsertLocation.after
  );
  dateP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
  dateP.font.size = headerFontSize;
  dateP.font.bold = false;
  dateP.font.italic = true;
  dateP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
  dateP.alignment = Word.Alignment.centered;
  dateP.firstLineIndent = 0;
  dateP.spaceBefore = 0;
  dateP.spaceAfter = 0;
  dateP.lineSpacing = 15;

  await context.sync();
  return table;
}

/**
 * 4. Hàm chèn bảng Chân trang (Ô Nơi nhận & Ô Chữ ký) theo Hướng dẫn 05-HD/VPTW
 *
 * Cột trái:
 * - "Nơi nhận:" (11-12 pt, in thường, ĐẬM - NGHIÊNG).
 * - Các gạch đầu dòng nơi nhận: (10-11 pt, in thường, đứng).
 *
 * Cột phải:
 * - Thẩm quyền/chức vụ người ký (13-14 pt, in hoa, ĐẬM).
 * - Chức danh cụ thể dòng 2 nếu có (13-14 pt, in hoa, ĐẬM).
 * - Khoảng trống chèn chữ ký: 35 mm - 40 mm (khoảng cách spaceBefore ~107.7 pt).
 * - Họ và tên người ký (13-14 pt, in thường, ĐẬM).
 */
export async function insertSignatureTable(
  context: Word.RequestContext,
  data: SignData
): Promise<Word.Table> {
  const body = context.document.body;
  const signerFontSize = data.fontSize ?? FOOTER_CONSTANTS.SIGNER_TITLE_FONT_SIZE;
  const sigSpacePt = data.signatureSpaceMm
    ? mmToPt(data.signatureSpaceMm)
    : FOOTER_CONSTANTS.SIGNATURE_SPACE_PT;

  // Chèn bảng 1 hàng x 2 cột vào cuối tài liệu
  const table = body.insertTable(1, 2, Word.InsertLocation.end, [["", ""]]);

  // Xóa đường viền toàn bảng (Border = none)
  try {
    table.getBorder(Word.BorderLocation.all).type = Word.BorderType.none;
  } catch {
    // Dự phòng
  }

  // Thiết lập lề ô = 0
  try {
    table.setCellPadding(Word.CellPaddingLocation.top, 0);
    table.setCellPadding(Word.CellPaddingLocation.bottom, 0);
    table.setCellPadding(Word.CellPaddingLocation.left, 0);
    table.setCellPadding(Word.CellPaddingLocation.right, 0);
  } catch {
    // Dự phòng
  }

  const cellLeft = table.getCell(0, 0);
  const cellRight = table.getCell(0, 1);

  // Cấu hình độ rộng cột và căn đỉnh cho Nơi nhận (~45%) và Chữ ký (~55%)
  try {
    cellLeft.columnWidth = mmToPt(75);
    cellRight.columnWidth = mmToPt(90);
    cellLeft.verticalAlignment = Word.VerticalAlignment.top;
    cellRight.verticalAlignment = Word.VerticalAlignment.top;
  } catch {
    // Dự phòng
  }

  // ----------------------------------------------------
  // Cột trái: Nơi nhận
  // ----------------------------------------------------
  const leftFirstP = cellLeft.body.paragraphs.getFirst();
  leftFirstP.insertText(
    FOOTER_CONSTANTS.RECIPIENTS_LABEL,
    Word.InsertLocation.replace
  );
  leftFirstP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
  leftFirstP.font.size = FOOTER_CONSTANTS.RECIPIENTS_LABEL_FONT_SIZE;
  leftFirstP.font.bold = true;
  leftFirstP.font.italic = true;
  leftFirstP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
  leftFirstP.alignment = Word.Alignment.left;
  leftFirstP.firstLineIndent = 0;
  leftFirstP.spaceBefore = 0;
  leftFirstP.spaceAfter = 2;
  leftFirstP.lineSpacing = 14;

  let currentRecP = leftFirstP;
  for (const item of data.recipients) {
    const raw = item.trim();
    if (!raw) continue;
    // Tự động bổ sung dấu gạch đầu dòng nếu người dùng chưa nhập
    const itemText = raw.startsWith("-") ? raw : `- ${raw}`;

    currentRecP = currentRecP.insertParagraph(
      itemText,
      Word.InsertLocation.after
    );
    currentRecP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
    currentRecP.font.size = FOOTER_CONSTANTS.RECIPIENTS_ITEM_FONT_SIZE;
    currentRecP.font.bold = false;
    currentRecP.font.italic = false;
    currentRecP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
    currentRecP.alignment = Word.Alignment.left;
    currentRecP.firstLineIndent = 0;
    currentRecP.spaceBefore = 0;
    currentRecP.spaceAfter = 1.5;
    currentRecP.lineSpacing = 13;
  }

  // ----------------------------------------------------
  // Cột phải: Chữ ký & Chức vụ
  // ----------------------------------------------------
  // Dòng 1: Thẩm quyền/chức vụ ký (13-14 pt, HOA ĐẬM)
  const rightFirstP = cellRight.body.paragraphs.getFirst();
  rightFirstP.insertText(
    data.signTitle.trim().toUpperCase(),
    Word.InsertLocation.replace
  );
  rightFirstP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
  rightFirstP.font.size = signerFontSize;
  rightFirstP.font.bold = true;
  rightFirstP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
  rightFirstP.alignment = Word.Alignment.centered;
  rightFirstP.firstLineIndent = 0;
  rightFirstP.spaceBefore = 0;
  rightFirstP.spaceAfter = 2;
  rightFirstP.lineSpacing = 15;

  let namePrecedingP = rightFirstP;

  // Dòng 2: Chức vụ cụ thể (nếu có, ví dụ: "BÍ THƯ" sau dòng "T/M BAN BÍ THƯ")
  if (data.signerRole && data.signerRole.trim()) {
    namePrecedingP = rightFirstP.insertParagraph(
      data.signerRole.trim().toUpperCase(),
      Word.InsertLocation.after
    );
    namePrecedingP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
    namePrecedingP.font.size = signerFontSize;
    namePrecedingP.font.bold = true;
    namePrecedingP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
    namePrecedingP.alignment = Word.Alignment.centered;
    namePrecedingP.firstLineIndent = 0;
    namePrecedingP.spaceBefore = 0;
    namePrecedingP.spaceAfter = 2;
    namePrecedingP.lineSpacing = 15;
  }

  // Dòng Họ và tên người ký: tạo khoảng trống 35-40mm chèn chữ ký
  // bằng thuộc tính spaceBefore trực tiếp trên paragraph họ tên
  const nameP = namePrecedingP.insertParagraph(
    data.signerName.trim(),
    Word.InsertLocation.after
  );
  nameP.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
  nameP.font.size = signerFontSize;
  nameP.font.bold = true;
  nameP.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
  nameP.alignment = Word.Alignment.centered;
  nameP.firstLineIndent = 0;
  nameP.spaceBefore = sigSpacePt;
  nameP.spaceAfter = 0;
  nameP.lineSpacing = 15;

  await context.sync();
  return table;
}

/**
 * 5. Hàm kiểm tra thể thức văn bản (Lint Document) theo Hướng dẫn 05-HD/VPTW
 * Quét toàn bộ văn bản để phát hiện các lỗi thể thức phổ biến:
 * - Quên số 0 ở số văn bản nhỏ hơn 10 (ví dụ: Số 5-HD/VPTW thay vì Số 05-HD/VPTW).
 * - Quên số 0 ở ngày nhỏ hơn 10 hoặc tháng 1, tháng 2.
 * - Gõ nhiều dấu cách (space) hoặc ký tự Tab ở đầu đoạn thay vì dùng thụt lề 10mm.
 * - Dùng phông chữ khác "Times New Roman".
 * - Căn lề trang (Margins) không đúng tỷ lệ A4 chuẩn Đảng.
 * - Đoạn văn bản nội dung không được căn đều 2 bên (Justified).
 * - Giãn dòng không nằm trong khoảng quy chuẩn (18-22 pt) hoặc Space After < 6 pt.
 */
export async function lintDocument(
  context: Word.RequestContext
): Promise<LintReport> {
  const issues: LintIssue[] = [];

  // Tải thông tin Sections để kiểm tra khổ giấy & lề
  const sections = context.document.sections;
  sections.load(["items/pageSetup"]);

  // Tải danh sách các đoạn văn bản (paragraphs)
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load([
    "items/text",
    "items/tableNestingLevel",
    "items/alignment",
    "items/firstLineIndent",
    "items/lineSpacing",
    "items/spaceBefore",
    "items/spaceAfter",
    "items/font/name",
    "items/font/size",
    "items/font/color",
  ]);

  await context.sync();

  // ----------------------------------------------------
  // Kiểm tra 1: Khổ giấy & Lề trang (Page Setup)
  // ----------------------------------------------------
  for (let sIdx = 0; sIdx < sections.items.length; sIdx++) {
    const ps = sections.items[sIdx].pageSetup;
    const topMm = ptToMm(ps.topMargin);
    const bottomMm = ptToMm(ps.bottomMargin);
    const leftMm = ptToMm(ps.leftMargin);
    const rightMm = ptToMm(ps.rightMargin);

    const TOLERANCE_MM = 1.5; // Sai số cho phép 1.5mm

    if (
      Math.abs(topMm - PAGE_SETUP_CONSTANTS.MARGIN_TOP_MM) > TOLERANCE_MM ||
      Math.abs(bottomMm - PAGE_SETUP_CONSTANTS.MARGIN_BOTTOM_MM) > TOLERANCE_MM ||
      Math.abs(leftMm - PAGE_SETUP_CONSTANTS.MARGIN_LEFT_MM) > TOLERANCE_MM ||
      Math.abs(rightMm - PAGE_SETUP_CONSTANTS.MARGIN_RIGHT_MM) > TOLERANCE_MM
    ) {
      issues.push({
        id: `PAGE_MARGIN_${sIdx + 1}`,
        type: LintIssueType.INVALID_MARGIN,
        severity: LintSeverity.ERROR,
        message: `Lề trang Section ${sIdx + 1} chưa chuẩn: Trên ${topMm}mm, Dưới ${bottomMm}mm, Trái ${leftMm}mm, Phải ${rightMm}mm.`,
        suggestion: `Điều chỉnh lề về chuẩn 05-HD/VPTW: Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm.`,
        locationInfo: `Section ${sIdx + 1}`,
      });
    }

    if (!ps.differentFirstPageHeaderFooter) {
      issues.push({
        id: `PAGE_DIFF_FIRST_${sIdx + 1}`,
        type: LintIssueType.INVALID_MARGIN,
        severity: LintSeverity.WARNING,
        message: `Chưa bật chế độ không đánh số trang ở trang đầu tiên (differentFirstPageHeaderFooter).`,
        suggestion: `Bật tùy chọn 'Different First Page' để không in số trang tại trang bìa/đầu.`,
        locationInfo: `Section ${sIdx + 1}`,
      });
    }
  }

  // ----------------------------------------------------
  // Kiểm tra 2: Quét từng đoạn văn bản
  // ----------------------------------------------------
  // Regex tìm lỗi thiếu số 0 ở số văn bản nhỏ hơn 10 (ví dụ: "Số 5-HD", "Số: 1/QĐ")
  const missingLeadingZeroDocRegex = /Số\s*:?\s*([1-9])(?!\d)\s*[-/]/i;

  // Regex tìm lỗi thiếu số 0 ở ngày < 10 (ví dụ: "ngày 5 tháng")
  const missingLeadingZeroDayRegex = /ngày\s+([1-9])(?!\d)\s+tháng/i;

  // Regex tìm lỗi thiếu số 0 ở tháng 1, tháng 2 (ví dụ: "tháng 1 năm", "tháng 2 năm")
  const missingLeadingZeroMonthRegex = /tháng\s+([1-2])(?!\d)\s+năm/i;

  // Regex kiểm tra dùng Space hoặc Tab để thụt đầu dòng
  const spaceTabIndentRegex = /^([ ]{2,}|\t+)/;

  for (let pIdx = 0; pIdx < paragraphs.items.length; pIdx++) {
    const p = paragraphs.items[pIdx];
    const text = p.text;
    const trimmed = text.trim();
    if (!trimmed) continue;

    const loc = `Đoạn ${pIdx + 1}`;
    const snippet = trimmed.length > 50 ? `${trimmed.slice(0, 50)}...` : trimmed;

    // 2.1 Kiểm tra lỗi gõ phím Space/Tab thay vì dùng First Line Indent
    if (spaceTabIndentRegex.test(text)) {
      issues.push({
        id: `SPACE_TAB_INDENT_${pIdx + 1}`,
        type: LintIssueType.SPACE_OR_TAB_INDENTATION,
        severity: LintSeverity.ERROR,
        message: `Đoạn văn dùng phím Space hoặc phím Tab ở đầu dòng để thụt lề thay vì thiết lập First Line Indent.`,
        suggestion: `Xóa các khoảng trắng/tab ở đầu đoạn và thiết lập First Line Indent = 10 mm (~28.35 pt).`,
        locationInfo: loc,
        snippet,
      });
    }

    // 2.2 Kiểm tra phông chữ (Phải là Times New Roman)
    if (p.font && p.font.name && p.font.name !== TYPOGRAPHY_CONSTANTS.FONT_FAMILY) {
      issues.push({
        id: `FONT_NAME_${pIdx + 1}`,
        type: LintIssueType.NON_STANDARD_FONT,
        severity: LintSeverity.ERROR,
        message: `Phông chữ hiện tại là '${p.font.name}', không đúng quy chuẩn phông '${TYPOGRAPHY_CONSTANTS.FONT_FAMILY}'.`,
        suggestion: `Chuyển toàn bộ phông chữ về '${TYPOGRAPHY_CONSTANTS.FONT_FAMILY}'.`,
        locationInfo: loc,
        snippet,
      });
    }

    // 2.3 Kiểm tra lỗi thiếu số 0 ở số văn bản
    const docMatch = missingLeadingZeroDocRegex.exec(trimmed);
    if (docMatch) {
      issues.push({
        id: `DOC_NUM_ZERO_${pIdx + 1}`,
        type: LintIssueType.MISSING_LEADING_ZERO_NUMBER,
        severity: LintSeverity.ERROR,
        message: `Số văn bản đang ghi là 'Số ${docMatch[1]}...', thiếu số 0 đằng trước đối với số nhỏ hơn 10.`,
        suggestion: `Sửa thành 'Số 0${docMatch[1]}...' theo quy định tại Hướng dẫn 05-HD/VPTW.`,
        locationInfo: loc,
        snippet,
      });
    }

    // 2.4 Kiểm tra lỗi thiếu số 0 ở ngày tháng
    const dayMatch = missingLeadingZeroDayRegex.exec(trimmed);
    if (dayMatch) {
      issues.push({
        id: `DATE_DAY_ZERO_${pIdx + 1}`,
        type: LintIssueType.MISSING_LEADING_ZERO_DATE,
        severity: LintSeverity.ERROR,
        message: `Ngày ban hành đang ghi là 'ngày ${dayMatch[1]}', thiếu số 0 đối với ngày dưới 10.`,
        suggestion: `Sửa thành 'ngày 0${dayMatch[1]}' theo quy định tại Hướng dẫn 05-HD/VPTW.`,
        locationInfo: loc,
        snippet,
      });
    }

    const monthMatch = missingLeadingZeroMonthRegex.exec(trimmed);
    if (monthMatch) {
      issues.push({
        id: `DATE_MONTH_ZERO_${pIdx + 1}`,
        type: LintIssueType.MISSING_LEADING_ZERO_DATE,
        severity: LintSeverity.ERROR,
        message: `Tháng ban hành đang ghi là 'tháng ${monthMatch[1]}', thiếu số 0 đối với tháng 1 và tháng 2.`,
        suggestion: `Sửa thành 'tháng 0${monthMatch[1]}' theo quy định tại Hướng dẫn 05-HD/VPTW.`,
        locationInfo: loc,
        snippet,
      });
    }

    // 2.5 Kiểm tra quy chuẩn đoạn văn bản nội dung (bỏ qua đoạn trong bảng)
    if (p.tableNestingLevel === 0) {
      const isCenteredHeading =
        p.alignment === Word.Alignment.centered || isDocumentHeading(trimmed);

      if (!isCenteredHeading) {
        // Căn đều 2 bên (Justified)
        if (p.alignment !== Word.Alignment.justified) {
          issues.push({
            id: `ALIGNMENT_${pIdx + 1}`,
            type: LintIssueType.INVALID_ALIGNMENT,
            severity: LintSeverity.WARNING,
            message: `Đoạn nội dung chưa được căn đều hai bên (Justified). Hiện tại là '${p.alignment}'.`,
            suggestion: `Đặt Alignment của đoạn thành Căn đều hai bên (Justified).`,
            locationInfo: loc,
            snippet,
          });
        }

        // Thụt lề dòng đầu: 10mm (~28.35 pt, cho phép sai số ±3 pt)
        const indentDiff = Math.abs(
          p.firstLineIndent - TYPOGRAPHY_CONSTANTS.FIRST_LINE_INDENT_PT
        );
        if (indentDiff > 3) {
          issues.push({
            id: `INDENT_${pIdx + 1}`,
            type: LintIssueType.INVALID_FIRST_LINE_INDENT,
            severity: LintSeverity.WARNING,
            message: `Thụt lề đầu dòng hiện tại là ${ptToMm(p.firstLineIndent)}mm (~${p.firstLineIndent}pt), khác chuẩn 10mm (~28.35pt).`,
            suggestion: `Đặt First Line Indent thành 10 mm (28.35 pt).`,
            locationInfo: loc,
            snippet,
          });
        }

        // Cỡ chữ nội dung: 13-14 pt
        if (p.font && p.font.size && (p.font.size < 13 || p.font.size > 14)) {
          issues.push({
            id: `FONT_SIZE_${pIdx + 1}`,
            type: LintIssueType.NON_STANDARD_FONT,
            severity: LintSeverity.WARNING,
            message: `Cỡ chữ nội dung là ${p.font.size}pt, nằm ngoài khoảng cho phép (13 - 14 pt).`,
            suggestion: `Chuyển cỡ chữ về 14 pt (hoặc 13 pt).`,
            locationInfo: loc,
            snippet,
          });
        }

        // Giãn dòng: 18 - 22 pt
        if (
          p.lineSpacing &&
          (p.lineSpacing < TYPOGRAPHY_CONSTANTS.LINE_SPACING_MIN_PT ||
            p.lineSpacing > TYPOGRAPHY_CONSTANTS.LINE_SPACING_MAX_PT)
        ) {
          issues.push({
            id: `LINE_SPACING_${pIdx + 1}`,
            type: LintIssueType.INVALID_LINE_SPACING,
            severity: LintSeverity.INFO,
            message: `Giãn dòng hiện tại là ${p.lineSpacing}pt (quy chuẩn: chế độ Exactly từ 18 pt đến 22 pt).`,
            suggestion: `Điều chỉnh Line Spacing về 18 pt.`,
            locationInfo: loc,
            snippet,
          });
        }

        // Giãn đoạn phía dưới: tối thiểu 6 pt
        if (p.spaceAfter < TYPOGRAPHY_CONSTANTS.SPACE_AFTER_MIN_PT) {
          issues.push({
            id: `SPACE_AFTER_${pIdx + 1}`,
            type: LintIssueType.INVALID_PARAGRAPH_SPACING,
            severity: LintSeverity.INFO,
            message: `Khoảng cách dưới đoạn (Space After) là ${p.spaceAfter}pt, nhỏ hơn mức tối thiểu 6 pt.`,
            suggestion: `Đặt Space After tối thiểu là 6 pt.`,
            locationInfo: loc,
            snippet,
          });
        }
      }
    }
  }

  const errorCount = issues.filter(
    (i) => i.severity === LintSeverity.ERROR
  ).length;
  const warningCount = issues.filter(
    (i) => i.severity === LintSeverity.WARNING
  ).length;
  const infoCount = issues.filter((i) => i.severity === LintSeverity.INFO).length;

  return {
    isValid: errorCount === 0,
    totalIssues: issues.length,
    errorCount,
    warningCount,
    infoCount,
    issues,
    timestamp: new Date().toISOString(),
  };
}
