/**
 * src/formatService.ts
 * Tầng 2 (Zone-Aware Formatter) & Tầng 3 (Structural & Semantic Linter)
 * Chuẩn hóa và kiểm tra logic văn bản Đảng theo Hướng dẫn số 05-HD/VPTW
 */

import {
  mmToPt,
  ptToMm,
  PAGE_SETUP_CONSTANTS,
  TYPOGRAPHY_CONSTANTS,
  HEADER_CONSTANTS,
  FOOTER_CONSTANTS,
  PAGE_NUMBERING_CONSTANTS,
  PARTY_DOC_CODE_MAP,
  COLLECTIVE_ORGANS,
  EXECUTIVE_ORGANS,
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
  DocumentZone,
  ParagraphRole,
  ParsedDocument,
} from "./types";

import { parseWordDocument } from "./documentParser";

/**
 * Chuẩn hóa số thứ tự văn bản (bắt buộc có số 0 đằng trước nếu số < 10)
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
  if (!trimmed || trimmed === "..." || trimmed === "…") {
    return "   ";
  }
  return trimmed;
}

/**
 * Chuẩn hóa ngày tháng năm theo Hướng dẫn 05-HD/VPTW:
 * - Ngày < 10: thêm số 0 (ngày 05)
 * - Tháng 1, 2: thêm số 0 (tháng 01, tháng 02)
 * - Tháng 3 đến 9: bỏ số 0 thừa nếu có (tháng 03 -> tháng 3)
 */
export function formatPartyDate(
  location: string,
  dateInput: Date | { day: number; month: number; year: number } | string
): string {
  const cleanLocation = location.trim().replace(/,\s*$/, "");

  if (typeof dateInput === "string") {
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
    // Loại bỏ số 0 thừa ở tháng 3 đến tháng 9
    formattedText = formattedText.replace(
      /tháng\s+0([3-9])\s+năm/i,
      "tháng $1 năm"
    );
    // Nếu để trống ngày / tháng khi xin số: tự động để khoảng trống 3 space
    formattedText = formattedText.replace(
      /ngày(?:\s*|\s*\.{2,}\s*|\s*…\s*)tháng/i,
      "ngày   tháng"
    );
    formattedText = formattedText.replace(
      /tháng(?:\s*|\s*\.{2,}\s*|\s*…\s*)năm/i,
      "tháng   năm"
    );
    if (
      cleanLocation &&
      !formattedText.toLowerCase().includes(cleanLocation.toLowerCase())
    ) {
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

function createPageNumberFlatOpcXml(fontSizePt: number = 13.5): string {
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
 * 1. Thiết lập khổ giấy A4 và căn lề chuẩn 05-HD/VPTW
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

    pageSetup.pageWidth = PAGE_SETUP_CONSTANTS.A4_WIDTH_PT;
    pageSetup.pageHeight = PAGE_SETUP_CONSTANTS.A4_HEIGHT_PT;
    pageSetup.orientation = Word.PageOrientation.portrait;

    pageSetup.topMargin = topPt;
    pageSetup.bottomMargin = bottomPt;
    pageSetup.leftMargin = leftPt;
    pageSetup.rightMargin = rightPt;
    pageSetup.headerDistance = headerDistPt;
    pageSetup.footerDistance = PAGE_SETUP_CONSTANTS.FOOTER_DISTANCE_PT;
    pageSetup.differentFirstPageHeaderFooter = diffFirst;

    try {
      (pageSetup as any).mirrorMargins = mirror;
    } catch {
      // Bỏ qua nếu host không hỗ trợ
    }

    try {
      const primaryHeader = section.getHeader(Word.HeaderFooterType.primary);
      primaryHeader.insertOoxml(pageNumberXml, Word.InsertLocation.replace);

      if (diffFirst) {
        const firstHeader = section.getHeader(Word.HeaderFooterType.firstPage);
        firstHeader.clear();
      }

      if (mirror) {
        try {
          const evenHeader = section.getHeader(Word.HeaderFooterType.evenPages);
          evenHeader.insertOoxml(pageNumberXml, Word.InsertLocation.replace);
        } catch {
          // Bỏ qua
        }
      }
    } catch {
      const primaryHeader = section.getHeader(Word.HeaderFooterType.primary);
      primaryHeader.clear();
      const p = primaryHeader.insertParagraph("", Word.InsertLocation.start);
      p.alignment = Word.Alignment.centered;
      p.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
      p.font.size = PAGE_NUMBERING_CONSTANTS.FONT_SIZE;
      p.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;
    }
  }

  await context.sync();
}

/**
 * 2. Chuẩn hóa thông minh theo Phân vùng (Zone-Aware Formatter)
 * - Hiểu và định dạng riêng từng Zone: HEADER, TITLE, SALUTATION, LEGAL_BASIS, BODY, FOOTER.
 * - Tự động chuẩn hóa cả nội dung nằm trong bảng Header & bảng Chữ ký (không còn bị "mù trong bảng").
 * - Bảo vệ các bảng dữ liệu/phụ lục ở giữa bài (`DATA_TABLE_CELL`).
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

  const { parsed, wordParagraphs } = await parseWordDocument(context);

  for (const node of parsed.nodes) {
    const p = wordParagraphs.items[node.index];

    if (node.role === ParagraphRole.EMPTY) {
      p.spaceBefore = 0;
      p.spaceAfter = 0;
      p.firstLineIndent = 0;
      continue;
    }

    // Luôn đồng bộ phông chữ Times New Roman và màu đen chuẩn
    p.font.name = TYPOGRAPHY_CONSTANTS.FONT_FAMILY;
    p.font.color = TYPOGRAPHY_CONSTANTS.FONT_COLOR;

    // ---------------------------------------------------------
    // XỬ LÝ VÙNG HEADER (Ô 1 & Ô 2)
    // ---------------------------------------------------------
    if (node.zone === DocumentZone.HEADER) {
      p.firstLineIndent = 0;
      p.spaceBefore = 0;
      p.lineSpacing = 15;
      if (node.tableNestingLevel > 0) {
        p.alignment = Word.Alignment.centered;
      }

      switch (node.role) {
        case ParagraphRole.HEADER_SUPERIOR_ORGAN:
          p.font.size = HEADER_CONSTANTS.SUPERIOR_ORGAN_FONT_SIZE;
          p.font.bold = false;
          p.font.italic = false;
          p.spaceAfter = 1;
          break;
        case ParagraphRole.HEADER_ISSUING_ORGAN:
          p.font.size = HEADER_CONSTANTS.ISSUING_ORGAN_FONT_SIZE;
          p.font.bold = true;
          p.font.italic = false;
          p.spaceAfter = 1;
          break;
        case ParagraphRole.HEADER_PARTY_MOTTO:
          p.font.size = HEADER_CONSTANTS.PARTY_TITLE_FONT_SIZE;
          p.font.bold = true;
          p.font.italic = false;
          p.spaceAfter = 1;
          break;
        case ParagraphRole.HEADER_SEPARATOR_LINE:
          p.font.size = 8;
          p.spaceAfter = 3;
          break;
        case ParagraphRole.HEADER_DOC_NUMBER: {
          p.font.size = HEADER_CONSTANTS.DOC_NUMBER_FONT_SIZE;
          p.font.bold = false;
          p.font.italic = false;
          p.spaceAfter = 0;
          // Tự động chuẩn hóa định dạng Số 05-BC/ĐU nếu người dùng gõ Số: 5/BC-ĐU
          const fixedCode = node.text
            .trim()
            .replace(
              /^Số\s*:\s*(\d+)\s*\/\s*([A-ZĐa-zđ]+)\s*-\s*([A-ZĐa-zđ0-9]+)/i,
              (_, num, type, org) => `Số ${formatDocNumber(num)}-${type.toUpperCase()}/${org.toUpperCase()}`
            )
            .replace(
              /^Số\s*:?\s*([1-9])(?!\d)\s*([-/])/i,
              "Số 0$1$2"
            );
          if (fixedCode !== node.text.trim() && !node.text.includes("\t")) {
            p.insertText(fixedCode, Word.InsertLocation.replace);
          }
          break;
        }
        case ParagraphRole.HEADER_DATE_LOCATION: {
          p.font.size = HEADER_CONSTANTS.DATE_LOCATION_FONT_SIZE;
          p.font.bold = false;
          p.font.italic = true;
          p.spaceAfter = 0;
          const fixedDate = formatPartyDate("", node.text.trim());
          if (fixedDate !== node.text.trim() && !node.text.includes("\t")) {
            p.insertText(fixedDate, Word.InsertLocation.replace);
          }
          break;
        }
      }
      continue;
    }

    // ---------------------------------------------------------
    // XỬ LÝ VÙNG TITLE (TÊN LOẠI VĂN BẢN & TRÍCH YẾU)
    // ---------------------------------------------------------
    if (node.zone === DocumentZone.TITLE) {
      p.firstLineIndent = 0;
      p.alignment = Word.Alignment.centered;
      p.lineSpacing = targetLineSpacing;

      if (node.role === ParagraphRole.DOC_TYPE) {
        p.font.size = TYPOGRAPHY_CONSTANTS.DOC_TYPE_FONT_SIZE;
        p.font.bold = true;
        p.font.italic = false;
        p.spaceBefore = 12;
        p.spaceAfter = 2;
      } else if (node.role === ParagraphRole.DOC_SUMMARY) {
        p.font.size = targetFontSize;
        p.font.bold = true;
        p.font.italic = false;
        p.spaceBefore = 0;
        p.spaceAfter = 6;
      } else if (node.role === ParagraphRole.CV_SUMMARY) {
        p.font.size = 12;
        p.font.bold = false;
        p.font.italic = false;
        p.spaceBefore = 2;
        p.spaceAfter = 6;
      } else if (node.role === ParagraphRole.TITLE_SEPARATOR_LINE) {
        p.font.size = 10;
        p.font.bold = true;
        p.spaceBefore = 0;
        p.spaceAfter = 6;
      }
      continue;
    }

    // ---------------------------------------------------------
    // XỬ LÝ VÙNG FOOTER (NƠI NHẬN & CHỮ KÝ)
    // ---------------------------------------------------------
    if (node.zone === DocumentZone.FOOTER) {
      p.firstLineIndent = 0;
      p.spaceBefore = 0;

      switch (node.role) {
        case ParagraphRole.RECIPIENT_HEADER:
          p.font.size = FOOTER_CONSTANTS.RECIPIENTS_LABEL_FONT_SIZE;
          p.font.bold = true;
          p.font.italic = true;
          p.alignment = Word.Alignment.left;
          p.spaceAfter = 2;
          p.lineSpacing = 14;
          break;
        case ParagraphRole.RECIPIENT_ITEM:
          p.font.size = FOOTER_CONSTANTS.RECIPIENTS_ITEM_FONT_SIZE;
          p.font.bold = false;
          p.font.italic = false;
          p.alignment = Word.Alignment.left;
          p.spaceAfter = 1.5;
          p.lineSpacing = 13;
          break;
        case ParagraphRole.SIGN_AUTHORITY: {
          p.font.size = FOOTER_CONSTANTS.SIGNER_TITLE_FONT_SIZE;
          p.font.bold = true;
          p.font.italic = false;
          if (node.tableNestingLevel > 0) p.alignment = Word.Alignment.centered;
          p.spaceAfter = 2;
          p.lineSpacing = 15;
          // Tự động sửa TM. -> T/M, KT. -> K/T, TL. -> T/L
          const fixedSignAuth = node.text
            .trim()
            .toUpperCase()
            .replace(/^TM\.?\s+/, "T/M ")
            .replace(/^KT\.?\s+/, "K/T ")
            .replace(/^TL\.?\s+/, "T/L ");
          if (fixedSignAuth !== node.text.trim()) {
            p.insertText(fixedSignAuth, Word.InsertLocation.replace);
          }
          break;
        }
        case ParagraphRole.SIGN_ROLE:
          p.font.size = FOOTER_CONSTANTS.SIGNER_TITLE_FONT_SIZE;
          p.font.bold = true;
          p.font.italic = false;
          if (node.tableNestingLevel > 0) p.alignment = Word.Alignment.centered;
          p.spaceAfter = 2;
          p.lineSpacing = 15;
          break;
        case ParagraphRole.SIGN_NAME:
          p.font.size = FOOTER_CONSTANTS.SIGNER_NAME_FONT_SIZE;
          p.font.bold = true;
          p.font.italic = false;
          if (node.tableNestingLevel > 0) p.alignment = Word.Alignment.centered;
          if (p.spaceBefore < mmToPt(25)) {
            p.spaceBefore = FOOTER_CONSTANTS.SIGNATURE_SPACE_PT;
          }
          p.spaceAfter = 0;
          p.lineSpacing = 15;
          break;
      }
      continue;
    }

    // ---------------------------------------------------------
    // XỬ LÝ VÙNG SALUTATION, LEGAL_BASIS & BODY
    // ---------------------------------------------------------
    if (node.role === ParagraphRole.DATA_TABLE_CELL) {
      // Bảo vệ bảng biểu dữ liệu trong bài: không ép thụt đầu dòng
      continue;
    }

    // Tự động xóa khoảng trắng/tab thừa ở đầu dòng đối với đoạn nội dung
    if (/^([ ]{2,}|\t+)/.test(node.text)) {
      p.insertText(node.text.replace(/^([ ]+|\t+)/, ""), Word.InsertLocation.replace);
    }

    if (
      node.role === ParagraphRole.DECISION_VERB ||
      node.role === ParagraphRole.HEADING_PART ||
      node.role === ParagraphRole.END_MARKER
    ) {
      p.alignment = Word.Alignment.centered;
      p.firstLineIndent = 0;
      p.font.size = targetFontSize;
      p.font.bold = true;
      p.spaceBefore = 6;
      p.spaceAfter = 6;
      p.lineSpacing = targetLineSpacing;
      continue;
    }

    // Các đoạn nội dung, đề mục I., 1., Điều 1., Kính gửi, Căn cứ
    p.font.size = targetFontSize;
    p.alignment = Word.Alignment.justified;
    p.firstLineIndent = targetIndentPt;
    p.spaceBefore = targetSpaceBefore;
    p.spaceAfter = targetSpaceAfter;
    p.lineSpacing = targetLineSpacing;

    if (node.role === ParagraphRole.HEADING_ROMAN) {
      p.font.bold = true;
    }
  }

  await context.sync();
}

/**
 * 3. Chèn hoặc Cập nhật Bảng Header 2 cột không viền (Ô 1 & Ô 2)
 * Tự động kiểm tra nếu đã có bảng Header ở đầu văn bản thì cập nhật tại chỗ, tránh chèn đè.
 */
export async function insertHeaderTable(
  context: Word.RequestContext,
  data: HeaderData
): Promise<Word.Table> {
  const body = context.document.body;
  const headerFontSize = data.fontSize ?? HEADER_CONSTANTS.ISSUING_ORGAN_FONT_SIZE;

  const formattedNumber = formatDocNumber(data.docNumber);
  const cleanDocCode = data.docCode.trim().replace(/^[-/]+/, "");
  const docNumberString = `Số ${formattedNumber}-${cleanDocCode}`;
  const dateString = formatPartyDate(data.location, data.date);

  const { parsed } = await parseWordDocument(context);
  const tables = body.tables;
  tables.load("items");
  await context.sync();

  let table: Word.Table;
  if (parsed.header.hasHeaderTable && tables.items.length > 0) {
    table = tables.items[0];
    table.getCell(0, 0).body.clear();
    table.getCell(0, 1).body.clear();
  } else {
    table = body.insertTable(1, 2, Word.InsertLocation.start, [["", ""]]);
  }

  try {
    table.getBorder(Word.BorderLocation.all).type = Word.BorderType.none;
    table.setCellPadding(Word.CellPaddingLocation.top, 0);
    table.setCellPadding(Word.CellPaddingLocation.bottom, 0);
    table.setCellPadding(Word.CellPaddingLocation.left, 0);
    table.setCellPadding(Word.CellPaddingLocation.right, 0);
  } catch {
    // Dự phòng
  }

  const cellLeft = table.getCell(0, 0);
  const cellRight = table.getCell(0, 1);

  try {
    cellLeft.columnWidth = mmToPt(68);
    cellRight.columnWidth = mmToPt(97);
    cellLeft.verticalAlignment = Word.VerticalAlignment.top;
    cellRight.verticalAlignment = Word.VerticalAlignment.top;
  } catch {
    // Dự phòng
  }

  const leftFirstP = cellLeft.body.paragraphs.getFirst();
  let currentLeftP = leftFirstP;

  if (data.superiorOrgan && data.superiorOrgan.trim()) {
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

    currentLeftP = currentLeftP.insertParagraph(
      data.issuingOrgan.trim().toUpperCase(),
      Word.InsertLocation.after
    );
  } else {
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
 * 4. Chèn hoặc Cập nhật Bảng Chân trang (Ô Nơi nhận & Ô Chữ ký)
 * Tự động kiểm tra nếu đã có bảng Chữ ký ở cuối văn bản thì cập nhật tại chỗ, tránh chèn đè.
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

  const { parsed } = await parseWordDocument(context);
  const tables = body.tables;
  tables.load("items");
  await context.sync();

  let table: Word.Table;
  if (parsed.footer.hasSignatureTable && tables.items.length > 0) {
    table = tables.items[tables.items.length - 1];
    table.getCell(0, 0).body.clear();
    table.getCell(0, 1).body.clear();
  } else {
    table = body.insertTable(1, 2, Word.InsertLocation.end, [["", ""]]);
  }

  try {
    table.getBorder(Word.BorderLocation.all).type = Word.BorderType.none;
    table.setCellPadding(Word.CellPaddingLocation.top, 0);
    table.setCellPadding(Word.CellPaddingLocation.bottom, 0);
    table.setCellPadding(Word.CellPaddingLocation.left, 0);
    table.setCellPadding(Word.CellPaddingLocation.right, 0);
  } catch {
    // Dự phòng
  }

  const cellLeft = table.getCell(0, 0);
  const cellRight = table.getCell(0, 1);

  try {
    cellLeft.columnWidth = mmToPt(75);
    cellRight.columnWidth = mmToPt(90);
    cellLeft.verticalAlignment = Word.VerticalAlignment.top;
    cellRight.verticalAlignment = Word.VerticalAlignment.top;
  } catch {
    // Dự phòng
  }

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
 * 5. Bộ kiểm tra Thể thức & Logic diện rộng trên AST (`lintParsedDocument`)
 * Tách biệt hoàn toàn khỏi Word.RequestContext để vừa chạy trong Add-in vừa chạy Unit Test
 */
export function lintParsedDocument(
  parsed: ParsedDocument,
  existingIssues: LintIssue[] = []
): LintIssue[] {
  const issues: LintIssue[] = [...existingIssues];

  // =========================================================================
  // NHÓM 1: KIỂM TRA SỐ, KÝ HIỆU & NGÀY THÁNG BAN HÀNH (VÙNG HEADER)
  // =========================================================================
  if (parsed.header.docNumberRaw) {
    const rawCode = parsed.header.docNumberRaw;
    const pIdx = parsed.header.docNumberParagraphIndex;
    const loc = pIdx !== undefined ? `Đoạn ${pIdx + 1} (Số, ký hiệu)` : "Ô Số, ký hiệu";

    // 1.1 Bắt lỗi dùng dấu hai chấm "Số:" hoặc gạch chéo sai kiểu Hành chính Nhà nước (Số: 05/BC-ĐU)
    if (/^Số\s*:/i.test(rawCode) || /^Số\s*:?\s*\d+\s*\//i.test(rawCode)) {
      issues.push({
        id: `DOC_CODE_STATE_STYLE_${pIdx ?? 0}`,
        type: LintIssueType.INVALID_DOC_CODE_FORMAT,
        severity: LintSeverity.ERROR,
        message: `Ký hiệu '${rawCode}' đang viết theo thể thức hành chính Nhà nước (có dấu ':' hoặc dùng dấu '/' sau số thứ tự).`,
        suggestion: `Theo Hướng dẫn 05-HD/VPTW, văn bản Đảng không có dấu ':' sau chữ Số, giữa số và tên loại văn bản dùng dấu '-', giữa tên loại văn bản và cơ quan dùng dấu '/' (VD: Số 05-BC/ĐU).`,
        locationInfo: loc,
        snippet: rawCode,
        paragraphIndex: pIdx,
      });
    }

    // 1.2 Bắt lỗi thiếu số 0 ở số văn bản < 10
    const missingZeroDoc = /Số\s*:?\s*([1-9])(?!\d)\s*[-/]/i.exec(rawCode);
    if (missingZeroDoc) {
      issues.push({
        id: `DOC_NUM_ZERO_${pIdx ?? 0}`,
        type: LintIssueType.MISSING_LEADING_ZERO_NUMBER,
        severity: LintSeverity.ERROR,
        message: `Số văn bản đang ghi là 'Số ${missingZeroDoc[1]}...', thiếu số 0 đằng trước đối với số nhỏ hơn 10.`,
        suggestion: `Sửa thành 'Số 0${missingZeroDoc[1]}...' theo quy định tại Hướng dẫn 05-HD/VPTW.`,
        locationInfo: loc,
        snippet: rawCode,
        paragraphIndex: pIdx,
      });
    }
  }

  if (parsed.header.dateLocationRaw) {
    const rawDate = parsed.header.dateLocationRaw;
    const pIdx = parsed.header.dateParagraphIndex;
    const loc = pIdx !== undefined ? `Đoạn ${pIdx + 1} (Ngày tháng)` : "Ô Ngày tháng";

    // 1.3 Bắt lỗi thiếu số 0 ở ngày < 10
    const dayMissingZero = /ngày\s+([1-9])(?!\d)\s+tháng/i.exec(rawDate);
    if (dayMissingZero) {
      issues.push({
        id: `DATE_DAY_ZERO_${pIdx ?? 0}`,
        type: LintIssueType.MISSING_LEADING_ZERO_DATE,
        severity: LintSeverity.ERROR,
        message: `Ngày ban hành đang ghi là 'ngày ${dayMissingZero[1]}', thiếu số 0 đối với ngày dưới 10.`,
        suggestion: `Sửa thành 'ngày 0${dayMissingZero[1]}' theo quy định tại Hướng dẫn 05-HD/VPTW.`,
        locationInfo: loc,
        snippet: rawDate,
        paragraphIndex: pIdx,
      });
    }

    // 1.4 Bắt lỗi thiếu số 0 ở tháng 1, tháng 2
    const monthMissingZero = /tháng\s+([1-2])(?!\d)\s+năm/i.exec(rawDate);
    if (monthMissingZero) {
      issues.push({
        id: `DATE_MONTH_ZERO_${pIdx ?? 0}`,
        type: LintIssueType.MISSING_LEADING_ZERO_DATE,
        severity: LintSeverity.ERROR,
        message: `Tháng ban hành đang ghi là 'tháng ${monthMissingZero[1]}', thiếu số 0 đối với tháng 1 và tháng 2.`,
        suggestion: `Sửa thành 'tháng 0${monthMissingZero[1]}' theo quy định tại Hướng dẫn 05-HD/VPTW.`,
        locationInfo: loc,
        snippet: rawDate,
        paragraphIndex: pIdx,
      });
    }

    // 1.5 Bắt lỗi THỪA số 0 ở các tháng từ 3 đến 9 (tháng 03..09)
    const monthExcessZero = /tháng\s+0([3-9])\s+năm/i.exec(rawDate);
    if (monthExcessZero) {
      issues.push({
        id: `DATE_MONTH_EXCESS_ZERO_${pIdx ?? 0}`,
        type: LintIssueType.EXCESS_LEADING_ZERO_DATE,
        severity: LintSeverity.ERROR,
        message: `Tháng ban hành đang ghi là 'tháng 0${monthExcessZero[1]}', thừa số 0 đằng trước đối với các tháng từ 3 đến 9.`,
        suggestion: `Sửa thành 'tháng ${monthExcessZero[1]}' (chỉ tháng 01 và tháng 02 mới có số 0 đứng trước).`,
        locationInfo: loc,
        snippet: rawDate,
        paragraphIndex: pIdx,
      });
    }

    // 1.6 Kiểm tra tính hợp lệ thực tế của ngày lịch (VD: 30/02, 31/04)
    if (parsed.header.dayRaw && parsed.header.monthRaw && parsed.header.yearRaw) {
      const d = parseInt(parsed.header.dayRaw, 10);
      const m = parseInt(parsed.header.monthRaw, 10);
      const y = parseInt(parsed.header.yearRaw, 10);
      if (!isNaN(d) && !isNaN(m) && !isNaN(y)) {
        const maxDaysInMonth = new Date(y, m, 0).getDate();
        if (m < 1 || m > 12 || d < 1 || d > maxDaysInMonth) {
          issues.push({
            id: `DATE_INVALID_CALENDAR_${pIdx ?? 0}`,
            type: LintIssueType.INVALID_DATE_VALUE,
            severity: LintSeverity.ERROR,
            message: `Ngày lịch không tồn tại: ngày ${d} tháng ${m} năm ${y}.`,
            suggestion: `Kiểm tra lại ngày tháng thực tế trên lịch.`,
            locationInfo: loc,
            snippet: rawDate,
            paragraphIndex: pIdx,
          });
        }
      }
    }
  }

  // =========================================================================
  // NHÓM 2: ĐỐI CHIẾU CHÉO [SỐ, KÝ HIỆU] <-> [TÊN LOẠI VĂN BẢN] <-> [CẤU TRÚC]
  // =========================================================================
  if (parsed.header.docTypeAbbr) {
    const abbr = parsed.header.docTypeAbbr.toUpperCase();
    const expectedDocType = PARTY_DOC_CODE_MAP[abbr];

    if (abbr === "CV" && parsed.docTypeTitle === "CÔNG VĂN") {
      issues.push({
        id: `DOC_TYPE_CV_REDUNDANT`,
        type: LintIssueType.DOC_TYPE_CODE_MISMATCH,
        severity: LintSeverity.ERROR,
        message: `Văn bản Công văn (ký hiệu '${abbr}') không được trình bày dòng tiêu đề 'CÔNG VĂN' đứng riêng.`,
        suggestion: `Xóa dòng chữ 'CÔNG VĂN', chỉ trình bày dòng trích yếu 'V/v...' dưới số ký hiệu và dòng 'Kính gửi:...' ở bên phải.`,
        locationInfo: parsed.docTypeParagraphIndex !== undefined ? `Đoạn ${parsed.docTypeParagraphIndex + 1}` : "Tiêu đề văn bản",
        paragraphIndex: parsed.docTypeParagraphIndex,
      });
    } else if (
      expectedDocType &&
      abbr !== "CV" &&
      parsed.docTypeTitle &&
      parsed.docTypeTitle !== expectedDocType
    ) {
      issues.push({
        id: `DOC_TYPE_CODE_MISMATCH`,
        type: LintIssueType.DOC_TYPE_CODE_MISMATCH,
        severity: LintSeverity.ERROR,
        message: `Mâu thuẫn logic: Ký hiệu văn bản là '${abbr}' (${expectedDocType}) nhưng tiêu đề văn bản lại ghi là '${parsed.docTypeTitle}'.`,
        suggestion: `Đồng bộ lại ký hiệu văn bản hoặc sửa tiêu đề thành '${expectedDocType}'.`,
        locationInfo: parsed.docTypeParagraphIndex !== undefined ? `Đoạn ${parsed.docTypeParagraphIndex + 1}` : "Tiêu đề văn bản",
        paragraphIndex: parsed.docTypeParagraphIndex,
      });
    }
  }

  if (parsed.docTypeTitle === "QUYẾT ĐỊNH") {
    const hasArticles = parsed.nodes.some(
      (n) => n.role === ParagraphRole.ARTICLE_HEADING
    );
    if (!parsed.hasDecisionVerb || !hasArticles) {
      issues.push({
        id: `DECISION_STRUCTURE_MISSING`,
        type: LintIssueType.DOC_TYPE_CODE_MISMATCH,
        severity: LintSeverity.WARNING,
        message: `Văn bản QUYẾT ĐỊNH bị thiếu dòng 'QUYẾT ĐỊNH:' hoặc thiếu các 'Điều 1.', 'Điều 2.' trong phần nội dung.`,
        suggestion: `Bổ sung đầy đủ phần căn cứ, dòng 'QUYẾT ĐỊNH:' và hệ thống các Điều theo mẫu Quyết định của Đảng.`,
        locationInfo: "Cấu trúc Quyết định",
        paragraphIndex: parsed.docTypeParagraphIndex,
      });
    }
  }

  // =========================================================================
  // NHÓM 3: ĐỐI CHIẾU CHÉO [CƠ QUAN BAN HÀNH (Ô 1)] <-> [THẨM QUYỀN KÝ]
  // =========================================================================
  if (parsed.footer.signAuthority) {
    const rawAuth = parsed.footer.signAuthority.trim();
    const pIdx = parsed.footer.signAuthorityIndex;
    const loc = pIdx !== undefined ? `Đoạn ${pIdx + 1} (Thẩm quyền ký)` : "Ô Chữ ký";

    // 3.1 Bắt lỗi viết TM., KT., TL. kiểu hành chính Nhà nước thay vì T/M, K/T, T/L
    if (/^(TM|KT|TL)\.?\s+/i.test(rawAuth)) {
      issues.push({
        id: `SIGN_PREFIX_INVALID_${pIdx ?? 0}`,
        type: LintIssueType.INVALID_SIGN_PREFIX,
        severity: LintSeverity.ERROR,
        message: `Ký hiệu quyền hạn ký '${rawAuth.split(/\s+/)[0]}' viết sai thể thức văn bản của Đảng.`,
        suggestion: `Trong văn bản Đảng, bắt buộc dùng dấu gạch chéo in hoa: 'T/M' (Thay mặt), 'K/T' (Ký thay), 'T/L' (Thừa lệnh).`,
        locationInfo: loc,
        snippet: rawAuth,
        paragraphIndex: pIdx,
      });
    }
  }

  if (parsed.header.issuingOrgan && (parsed.footer.signAuthority || parsed.footer.signRole)) {
    const issuingUpper = parsed.header.issuingOrgan.toUpperCase();
    const authUpper = (parsed.footer.signAuthority || "").toUpperCase();
    const pIdx = parsed.footer.signAuthorityIndex ?? parsed.footer.signRoleIndex;
    const loc = pIdx !== undefined ? `Đoạn ${pIdx + 1} (Thẩm quyền ký)` : "Ô Chữ ký";

    const isExecutiveOrgan = EXECUTIVE_ORGANS.some((kw) => issuingUpper.startsWith(kw));
    const isCollectiveOrgan =
      !isExecutiveOrgan &&
      COLLECTIVE_ORGANS.some((kw) => issuingUpper.includes(kw));

    const hasTM = /^(T\/M|TM\.?)\s+/.test(authUpper);

    if (isCollectiveOrgan && !hasTM) {
      issues.push({
        id: `AUTHORITY_MISSING_TM`,
        type: LintIssueType.AUTHORITY_SIGNATURE_MISMATCH,
        severity: LintSeverity.ERROR,
        message: `Mâu thuẫn thẩm quyền: Cơ quan ban hành ở Ô 1 là tập thể ('${parsed.header.issuingOrgan}') nhưng phần chữ ký lại thiếu dòng thay mặt 'T/M ...'.`,
        suggestion: `Bổ sung dòng thẩm quyền ký thay mặt tập thể (VD: 'T/M BAN THƯỜNG VỤ' hoặc 'T/M ĐẢNG ỦY') phía trên chức danh người ký.`,
        locationInfo: loc,
        paragraphIndex: pIdx,
      });
    } else if (isExecutiveOrgan && hasTM) {
      issues.push({
        id: `AUTHORITY_WRONG_TM_FOR_EXECUTIVE`,
        type: LintIssueType.AUTHORITY_SIGNATURE_MISMATCH,
        severity: LintSeverity.ERROR,
        message: `Mâu thuẫn thẩm quyền: Cơ quan ban hành ở Ô 1 là '${parsed.header.issuingOrgan}' (cơ quan làm việc theo chế độ thủ trưởng) nên không được ký 'T/M' (${parsed.footer.signAuthority}).`,
        suggestion: `Nếu ban hành nhân danh cấp ủy thì sửa Ô 1 thành tên Đảng ủy/Cấp ủy; nếu ban hành nhân danh '${parsed.header.issuingOrgan}' thì người đứng đầu ký trực tiếp hoặc cấp phó ký 'K/T'.`,
        locationInfo: loc,
        paragraphIndex: pIdx,
      });
    }
  }

  // =========================================================================
  // NHÓM 4: KIỂM TRA LOGIC CÂY ĐỀ MỤC (HEADING SEQUENCE & HIERARCHY)
  // =========================================================================
  let expectedRoman = 1;
  let firstRomanSeparator: string | undefined;
  let expectedNumber = 1;
  let expectedArticle = 1;

  for (const node of parsed.nodes) {
    if (!node.headingMeta) continue;
    const loc = `Đoạn ${node.index + 1}`;

    if (node.headingMeta.level === "ROMAN") {
      const actual = node.headingMeta.numericValue;
      if (actual !== expectedRoman) {
        issues.push({
          id: `HEADING_ROMAN_SEQ_${node.index + 1}`,
          type: LintIssueType.HEADING_SEQUENCE_ERROR,
          severity: LintSeverity.ERROR,
          message: `Đề mục La Mã nhảy cóc hoặc đánh sai thứ tự: Đang ở mục '${node.headingMeta.rawIndex}' (giá trị ${actual}) trong khi thứ tự đúng phải là mục số ${expectedRoman}.`,
          suggestion: `Kiểm tra và đánh lại số thứ tự các mục lớn I, II, III... liền mạch.`,
          locationInfo: loc,
          snippet: node.text.trim(),
          paragraphIndex: node.index,
        });
      }
      expectedRoman = actual + 1;
      expectedNumber = 1; // Reset bộ đếm mục con 1., 2. khi sang mục La Mã mới

      if (!firstRomanSeparator) {
        firstRomanSeparator = node.headingMeta.separator;
      } else if (
        node.headingMeta.separator &&
        node.headingMeta.separator !== firstRomanSeparator
      ) {
        issues.push({
          id: `HEADING_ROMAN_SEP_${node.index + 1}`,
          type: LintIssueType.HEADING_SEQUENCE_ERROR,
          severity: LintSeverity.WARNING,
          message: `Dấu phân cách sau số La Mã không đồng nhất: mục đầu dùng '${firstRomanSeparator}' nhưng mục '${node.headingMeta.rawIndex}' lại dùng '${node.headingMeta.separator}'.`,
          suggestion: `Thống nhất dùng một kiểu dấu sau số La Mã trên toàn văn bản (hoặc 'I-' hoặc 'I.').`,
          locationInfo: loc,
          snippet: node.text.trim(),
          paragraphIndex: node.index,
        });
      }
    } else if (node.headingMeta.level === "NUMBER") {
      const actual = node.headingMeta.numericValue;
      if (actual !== expectedNumber) {
        issues.push({
          id: `HEADING_NUM_SEQ_${node.index + 1}`,
          type: LintIssueType.HEADING_SEQUENCE_ERROR,
          severity: LintSeverity.ERROR,
          message: `Đề mục số Ả Rập nhảy cóc hoặc sai thứ tự: Đang ở mục '${actual}.' nhưng thứ tự liên tiếp phải là '${expectedNumber}.'.`,
          suggestion: `Đánh lại số thứ tự các khoản 1., 2., 3. trong cùng một mục lớn.`,
          locationInfo: loc,
          snippet: node.text.trim(),
          paragraphIndex: node.index,
        });
      }
      expectedNumber = actual + 1;
    } else if (node.headingMeta.level === "ARTICLE") {
      const actual = node.headingMeta.numericValue;
      if (actual !== expectedArticle) {
        issues.push({
          id: `HEADING_ARTICLE_SEQ_${node.index + 1}`,
          type: LintIssueType.HEADING_SEQUENCE_ERROR,
          severity: LintSeverity.ERROR,
          message: `Thứ tự Điều trong văn bản bị nhảy cóc: Đang ở 'Điều ${actual}' nhưng thứ tự liên tiếp phải là 'Điều ${expectedArticle}'.`,
          suggestion: `Đánh lại thứ tự Điều 1, Điều 2, Điều 3... liên tục.`,
          locationInfo: loc,
          snippet: node.text.trim(),
          paragraphIndex: node.index,
        });
      }
      expectedArticle = actual + 1;
    }
  }

  // =========================================================================
  // NHÓM 5: KIỂM TRA LOGIC PHẦN "NƠI NHẬN" & DẤU CÂU KẾT THÚC
  // =========================================================================
  const recItems = parsed.footer.recipientItems;
  if (recItems.length > 0) {
    for (let idx = 0; idx < recItems.length; idx++) {
      const item = recItems[idx];
      const isLast = idx === recItems.length - 1;
      const loc = `Đoạn ${item.index + 1} (Nơi nhận)`;

      if (!isLast && !item.text.endsWith(";")) {
        issues.push({
          id: `RECIPIENT_SEMICOLON_${item.index + 1}`,
          type: LintIssueType.RECIPIENTS_FORMAT_ERROR,
          severity: LintSeverity.WARNING,
          message: `Dòng nơi nhận '${item.text}' chưa kết thúc bằng dấu chấm phẩy (;).`,
          suggestion: `Tất cả các dòng liệt kê nơi nhận (trừ dòng cuối cùng) phải kết thúc bằng dấu chấm phẩy (;).`,
          locationInfo: loc,
          snippet: item.text,
          paragraphIndex: item.index,
        });
      }

      if (isLast) {
        if (!item.text.endsWith(".")) {
          issues.push({
            id: `RECIPIENT_PERIOD_${item.index + 1}`,
            type: LintIssueType.RECIPIENTS_FORMAT_ERROR,
            severity: LintSeverity.WARNING,
            message: `Dòng nơi nhận cuối cùng '${item.text}' phải kết thúc bằng dấu chấm (.).`,
            suggestion: `Đặt dấu chấm (.) ở cuối dòng nơi nhận cuối cùng.`,
            locationInfo: loc,
            snippet: item.text,
            paragraphIndex: item.index,
          });
        }
        if (!/Lưu\s*(VT|Văn thư|:)/i.test(item.text)) {
          issues.push({
            id: `RECIPIENT_MISSING_ARCHIVE_${item.index + 1}`,
            type: LintIssueType.RECIPIENTS_FORMAT_ERROR,
            severity: LintSeverity.WARNING,
            message: `Phần 'Nơi nhận:' chưa có dòng lưu hồ sơ ('- Lưu Văn thư...' / '- Lưu: VT...').`,
            suggestion: `Bổ sung dòng '- Lưu Văn thư...' ở vị trí cuối cùng của mục Nơi nhận.`,
            locationInfo: loc,
            snippet: item.text,
            paragraphIndex: item.index,
          });
        }
      }
    }
  }

  // =========================================================================
  // NHÓM 6 & 7: QUÉT LỖI DỮ LIỆU BỎ TRỐNG (PLACEHOLDERS) & THỂ THỨC TỪNG ZONE
  // =========================================================================
  const spaceTabIndentRegex = /^([ ]{2,}|\t+)/;
  const placeholderRegex = /(\.{3,}|…|\[\s*\.+\s*\])/;
  const spaceBeforePunctRegex = /\s+([,.;:!?])/;

  for (const node of parsed.nodes) {
    if (node.role === ParagraphRole.EMPTY) continue;

    const trimmed = node.text.trim();
    const loc = `Đoạn ${node.index + 1}`;
    const snippet = trimmed.length > 55 ? `${trimmed.slice(0, 55)}...` : trimmed;

    // 6.1 Bắt dấu ba chấm (...) hoặc chỗ trống quên điền thông tin trước khi ban hành
    if (
      node.role !== ParagraphRole.HEADER_SEPARATOR_LINE &&
      node.role !== ParagraphRole.TITLE_SEPARATOR_LINE &&
      node.role !== ParagraphRole.END_MARKER &&
      placeholderRegex.test(trimmed)
    ) {
      issues.push({
        id: `PLACEHOLDER_${node.index + 1}`,
        type: LintIssueType.UNFILLED_PLACEHOLDER,
        severity: LintSeverity.WARNING,
        message: `Phát hiện dấu ba chấm ('...') hoặc trường thông tin có thể chưa được điền hoàn chỉnh.`,
        suggestion: `Rà soát và điền đầy đủ số liệu/ngày tháng/tên riêng trước khi trình ký.`,
        locationInfo: loc,
        snippet,
        paragraphIndex: node.index,
      });
    }

    // 6.2 Bắt lỗi gõ khoảng trắng thừa trước dấu câu (, . ; :)
    if (spaceBeforePunctRegex.test(trimmed)) {
      issues.push({
        id: `PUNCT_SPACE_${node.index + 1}`,
        type: LintIssueType.TYPOGRAPHY_PUNCTUATION_ERROR,
        severity: LintSeverity.INFO,
        message: `Có khoảng trắng thừa đứng trước dấu câu (',', '.', ';', ':').`,
        suggestion: `Xóa khoảng trắng thừa trước dấu câu để văn bản chuẩn kỹ thuật trình bày.`,
        locationInfo: loc,
        snippet,
        paragraphIndex: node.index,
      });
    }

    // 7.1 Kiểm tra phông chữ Times New Roman trên TOÀN BỘ các Zone (kể cả trong bảng)
    if (node.fontName && node.fontName !== TYPOGRAPHY_CONSTANTS.FONT_FAMILY) {
      issues.push({
        id: `FONT_NAME_${node.index + 1}`,
        type: LintIssueType.NON_STANDARD_FONT,
        severity: LintSeverity.ERROR,
        message: `Phông chữ hiện tại là '${node.fontName}', không đúng quy chuẩn '${TYPOGRAPHY_CONSTANTS.FONT_FAMILY}'.`,
        suggestion: `Chuyển phông chữ về '${TYPOGRAPHY_CONSTANTS.FONT_FAMILY}'.`,
        locationInfo: loc,
        snippet,
        paragraphIndex: node.index,
      });
    }

    // 7.2 Kiểm tra thể thức chuyên biệt trong vùng HEADER & FOOTER
    if (
      node.role === ParagraphRole.HEADER_ISSUING_ORGAN ||
      node.role === ParagraphRole.HEADER_PARTY_MOTTO ||
      node.role === ParagraphRole.SIGN_AUTHORITY ||
      node.role === ParagraphRole.SIGN_ROLE ||
      node.role === ParagraphRole.SIGN_NAME
    ) {
      if (node.fontBold === false) {
        issues.push({
          id: `HEADER_FOOTER_BOLD_${node.index + 1}`,
          type: LintIssueType.HEADER_FOOTER_STYLE_ERROR,
          severity: LintSeverity.WARNING,
          message: `Thành phần '${snippet}' bắt buộc phải được in đậm (Bold) theo Hướng dẫn 05-HD/VPTW.`,
          suggestion: `Bật định dạng In đậm (Bold) cho dòng này.`,
          locationInfo: loc,
          snippet,
          paragraphIndex: node.index,
        });
      }
    }

    if (node.role === ParagraphRole.HEADER_DATE_LOCATION && node.fontItalic === false) {
      issues.push({
        id: `HEADER_DATE_ITALIC_${node.index + 1}`,
        type: LintIssueType.HEADER_FOOTER_STYLE_ERROR,
        severity: LintSeverity.WARNING,
        message: `Dòng địa danh, ngày tháng năm '${snippet}' chưa được in nghiêng (Italic).`,
        suggestion: `Định dạng chữ in thường, kiểu chữ Nghiêng (Italic) cho dòng ngày tháng.`,
        locationInfo: loc,
        snippet,
        paragraphIndex: node.index,
      });
    }

    // 7.3 Kiểm tra đoạn văn trong vùng BODY (chỉ áp dụng cho văn xuôi và đề mục, KHÔNG báo lỗi nhầm tiêu đề)
    const isBodyProseOrHeading =
      node.zone === DocumentZone.BODY &&
      node.role !== ParagraphRole.DATA_TABLE_CELL &&
      node.role !== ParagraphRole.HEADING_PART &&
      node.role !== ParagraphRole.END_MARKER;

    if (isBodyProseOrHeading || node.zone === DocumentZone.LEGAL_BASIS) {
      if (node.role === ParagraphRole.DECISION_VERB) continue;

      if (spaceTabIndentRegex.test(node.text)) {
        issues.push({
          id: `SPACE_TAB_INDENT_${node.index + 1}`,
          type: LintIssueType.SPACE_OR_TAB_INDENTATION,
          severity: LintSeverity.ERROR,
          message: `Đoạn văn dùng phím Space hoặc Tab ở đầu dòng để thụt lề thay vì dùng First Line Indent.`,
          suggestion: `Xóa khoảng trắng/tab đầu dòng và đặt First Line Indent = 10mm.`,
          locationInfo: loc,
          snippet,
          paragraphIndex: node.index,
        });
      }

      if (
        node.alignment &&
        node.alignment.toLowerCase() !== "justified" &&
        node.alignment.toLowerCase() !== "justify"
      ) {
        issues.push({
          id: `ALIGNMENT_${node.index + 1}`,
          type: LintIssueType.INVALID_ALIGNMENT,
          severity: LintSeverity.WARNING,
          message: `Đoạn nội dung chưa được căn đều hai bên (Justified). Hiện tại là '${node.alignment}'.`,
          suggestion: `Đặt Alignment của đoạn thành Căn đều hai bên (Justified).`,
          locationInfo: loc,
          snippet,
          paragraphIndex: node.index,
        });
      }

      if (node.firstLineIndent !== undefined) {
        // Cho phép biên độ chuẩn từ 10mm (28.35pt) đến 12.7mm (36pt) với sai số ±3pt
        const minAllowedPt = TYPOGRAPHY_CONSTANTS.FIRST_LINE_INDENT_PT - 3;
        const maxAllowedPt = TYPOGRAPHY_CONSTANTS.FIRST_LINE_INDENT_MAX_PT + 3;
        if (node.firstLineIndent < minAllowedPt || node.firstLineIndent > maxAllowedPt) {
          issues.push({
            id: `INDENT_${node.index + 1}`,
            type: LintIssueType.INVALID_FIRST_LINE_INDENT,
            severity: LintSeverity.WARNING,
            message: `Thụt lề đầu dòng hiện tại là ${ptToMm(node.firstLineIndent)}mm (~${node.firstLineIndent}pt), nằm ngoài chuẩn 10mm - 12.7mm.`,
            suggestion: `Đặt First Line Indent chuẩn 10 mm (28.35 pt).`,
            locationInfo: loc,
            snippet,
            paragraphIndex: node.index,
          });
        }
      }

      if (node.fontSize && (node.fontSize < 13 || node.fontSize > 14.5)) {
        issues.push({
          id: `FONT_SIZE_${node.index + 1}`,
          type: LintIssueType.NON_STANDARD_FONT,
          severity: LintSeverity.WARNING,
          message: `Cỡ chữ nội dung là ${node.fontSize}pt, nằm ngoài khoảng quy chuẩn (13 - 14 pt).`,
          suggestion: `Chuyển cỡ chữ nội dung về 14 pt (hoặc 13 pt).`,
          locationInfo: loc,
          snippet,
          paragraphIndex: node.index,
        });
      }
    }
  }

  return issues;
}

/**
 * 6. Hàm tích hợp quét lỗi toàn văn bản Word (`lintDocument`)
 */
export async function lintDocument(
  context: Word.RequestContext
): Promise<LintReport> {
  const pageIssues: LintIssue[] = [];

  const sections = context.document.sections;
  sections.load(["items/pageSetup"]);
  const { parsed } = await parseWordDocument(context);

  // Kiểm tra Khổ giấy & Lề trang (Page Setup)
  for (let sIdx = 0; sIdx < sections.items.length; sIdx++) {
    const ps = sections.items[sIdx].pageSetup;
    const topMm = ptToMm(ps.topMargin);
    const bottomMm = ptToMm(ps.bottomMargin);
    const leftMm = ptToMm(ps.leftMargin);
    const rightMm = ptToMm(ps.rightMargin);

    const TOLERANCE_MM = 1.5;

    if (
      Math.abs(topMm - PAGE_SETUP_CONSTANTS.MARGIN_TOP_MM) > TOLERANCE_MM ||
      Math.abs(bottomMm - PAGE_SETUP_CONSTANTS.MARGIN_BOTTOM_MM) > TOLERANCE_MM ||
      Math.abs(leftMm - PAGE_SETUP_CONSTANTS.MARGIN_LEFT_MM) > TOLERANCE_MM ||
      Math.abs(rightMm - PAGE_SETUP_CONSTANTS.MARGIN_RIGHT_MM) > TOLERANCE_MM
    ) {
      pageIssues.push({
        id: `PAGE_MARGIN_${sIdx + 1}`,
        type: LintIssueType.INVALID_MARGIN,
        severity: LintSeverity.ERROR,
        message: `Lề trang Section ${sIdx + 1} chưa chuẩn: Trên ${topMm}mm, Dưới ${bottomMm}mm, Trái ${leftMm}mm, Phải ${rightMm}mm.`,
        suggestion: `Điều chỉnh lề về chuẩn 05-HD/VPTW: Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm.`,
        locationInfo: `Section ${sIdx + 1}`,
      });
    }

    if (!ps.differentFirstPageHeaderFooter) {
      pageIssues.push({
        id: `PAGE_DIFF_FIRST_${sIdx + 1}`,
        type: LintIssueType.INVALID_MARGIN,
        severity: LintSeverity.WARNING,
        message: `Chưa bật chế độ không đánh số trang ở trang đầu tiên (differentFirstPageHeaderFooter).`,
        suggestion: `Bật tùy chọn 'Different First Page' để không in số trang tại trang 1.`,
        locationInfo: `Section ${sIdx + 1}`,
      });
    }
  }

  const allIssues = lintParsedDocument(parsed, pageIssues);

  const errorCount = allIssues.filter((i) => i.severity === LintSeverity.ERROR).length;
  const warningCount = allIssues.filter((i) => i.severity === LintSeverity.WARNING).length;
  const infoCount = allIssues.filter((i) => i.severity === LintSeverity.INFO).length;

  return {
    isValid: errorCount === 0,
    totalIssues: allIssues.length,
    errorCount,
    warningCount,
    infoCount,
    issues: allIssues,
    timestamp: new Date().toISOString(),
  };
}
