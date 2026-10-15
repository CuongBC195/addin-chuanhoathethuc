/**
 * src/documentParser.ts
 * Tầng 1: Bộ phân tích cấu trúc & Phân vùng văn bản Đảng (AST & Zone Classifier)
 * Khắc phục triệt để nhược điểm "vô trạng thái" và "mù trong bảng"
 */

import {
  DocumentZone,
  ParagraphRole,
  RawParagraphNode,
  ParsedParagraphNode,
  ParsedDocument,
  ExtractedHeaderInfo,
  ExtractedFooterInfo,
} from "./types";

const DOC_TYPE_KEYWORDS = [
  "NGHỊ QUYẾT",
  "QUYẾT ĐỊNH",
  "CHỈ THỊ",
  "KẾT LUẬN",
  "QUY CHẾ",
  "QUY ĐỊNH",
  "HƯỚNG DẪN",
  "THÔNG BÁO",
  "BÁO CÁO",
  "KẾ HOẠCH",
  "CHƯƠNG TRÌNH",
  "TỜ TRÌNH",
  "CÔNG VĂN",
  "GIẤY MỜI",
  "BIÊN BẢN",
  "ĐỀ ÁN",
  "PHƯƠNG ÁN",
];

const SIGN_AUTHORITY_REGEX =
  /^(T\/M|TM\.?|K\/T|KT\.?|T\/L|TL\.?|Q\.|QUYỀN)\s+(.+)$/i;

const SIGN_ROLE_KEYWORDS = [
  "BÍ THƯ",
  "PHÓ BÍ THƯ",
  "PHÓ BÍ THƯ THƯỜNG TRỰC",
  "CHÁNH VĂN PHÒNG",
  "PHÓ CHÁNH VĂN PHÒNG",
  "TRƯỞNG BAN",
  "PHÓ TRƯỞNG BAN",
  "CHỦ NHIỆM",
  "PHÓ CHỦ NHIỆM",
  "GIÁM ĐỐC",
  "PHÓ GIÁM ĐỐC",
];

/** Chuyển số La Mã sang số nguyên (ví dụ: I -> 1, IV -> 4, IX -> 9) */
export function romanToInt(roman: string): number {
  const s = roman.toUpperCase().trim();
  const map: Record<string, number> = {
    I: 1,
    V: 5,
    X: 10,
    L: 50,
    C: 100,
    D: 500,
    M: 1000,
  };
  let total = 0;
  for (let i = 0; i < s.length; i++) {
    const curr = map[s[i]];
    const next = map[s[i + 1]];
    if (!curr) return NaN;
    if (next && curr < next) {
      total -= curr;
    } else {
      total += curr;
    }
  }
  return total;
}

function isSeparatorLine(text: string): boolean {
  const clean = text.trim();
  return /^([—–\-_*]\s*){3,}$/.test(clean);
}

function isExactDocType(text: string): boolean {
  const clean = text.trim().toUpperCase();
  return DOC_TYPE_KEYWORDS.includes(clean);
}

function isSignRoleLine(text: string): boolean {
  const clean = text.trim().toUpperCase();
  return SIGN_ROLE_KEYWORDS.some(
    (kw) => clean === kw || clean.startsWith(`${kw} `)
  );
}

/**
 * Phân tích danh sách các đoạn văn thô thành cấu trúc AST hoàn chỉnh (`ParsedDocument`)
 */
export function parseDocumentStructure(
  rawNodes: RawParagraphNode[]
): ParsedDocument {
  const nodes: ParsedParagraphNode[] = rawNodes.map((n) => ({
    ...n,
    zone: DocumentZone.BODY,
    role: n.text.trim() ? ParagraphRole.BODY_PARAGRAPH : ParagraphRole.EMPTY,
  }));

  const header: ExtractedHeaderInfo = {
    hasHeaderTable: false,
  };

  const footer: ExtractedFooterInfo = {
    hasSignatureTable: false,
    recipientItems: [],
  };

  let docTypeTitle: string | undefined;
  let docTypeParagraphIndex: number | undefined;
  let docSummary: string | undefined;
  let hasDecisionVerb = false;

  // ------------------------------------------------------------------
  // BƯỚC 1: NHẬN DIỆN VÙNG HEADER (Ô 1 & Ô 2) Ở ĐẦU VĂN BẢN
  // ------------------------------------------------------------------
  const maxHeaderScan = Math.min(nodes.length, 15);
  let lastHeaderIndex = -1;

  for (let i = 0; i < maxHeaderScan; i++) {
    const node = nodes[i];
    const trimmed = node.text.trim();
    if (!trimmed) continue;

    // Nếu gặp Tên loại văn bản hoặc Kính gửi thì kết thúc vùng Header
    if (isExactDocType(trimmed) || /^Kính gửi\s*:/i.test(trimmed)) {
      break;
    }

    const upper = trimmed.toUpperCase();
    const isPartyMotto = upper.includes("ĐẢNG CỘNG SẢN VIỆT NAM");
    const isDocNumber = /^Số\s*:?\s*(\d+|\.{2,}|…)/i.test(trimmed);
    const isDateLine =
      /(ngày\s*(\d+|\.{2,}|…)\s*tháng|tháng\s*(\d+|\.{2,}|…)\s*năm)/i.test(
        trimmed
      );
    const isOrganCandidate =
      i <= 6 &&
      upper === trimmed &&
      /(ĐẢNG|BAN|TỈNH ỦY|THÀNH ỦY|HUYỆN ỦY|THỊ ỦY|QUẬN ỦY|VĂN PHÒNG|ỦY BAN|CHI BỘ|TRUNG ƯƠNG)/.test(
        upper
      );

    if (
      isPartyMotto ||
      isDocNumber ||
      isDateLine ||
      isOrganCandidate ||
      (lastHeaderIndex >= 0 && i <= lastHeaderIndex + 2 && isSeparatorLine(trimmed))
    ) {
      node.zone = DocumentZone.HEADER;
      lastHeaderIndex = i;
      if (node.tableNestingLevel > 0) {
        header.hasHeaderTable = true;
      }

      if (isPartyMotto && isOrganCandidate && trimmed.includes("\t")) {
        // Trường hợp gõ cả cơ quan và tiêu ngữ trên cùng 1 dòng bằng phím Tab
        node.role = ParagraphRole.HEADER_RAW_COMBINED;
        const parts = trimmed.split(/\t+/).map((s) => s.trim()).filter(Boolean);
        if (parts.length >= 2) {
          if (!header.issuingOrgan) header.issuingOrgan = parts[0];
          header.partyMottoParagraphIndex = i;
        }
      } else if (isPartyMotto) {
        node.role = ParagraphRole.HEADER_PARTY_MOTTO;
        header.partyMottoParagraphIndex = i;
      } else if (isDocNumber) {
        node.role = ParagraphRole.HEADER_DOC_NUMBER;
        header.docNumberRaw = trimmed;
        header.docNumberParagraphIndex = i;

        // Tách chi tiết Số, loại văn bản viết tắt, cơ quan ban hành viết tắt
        const codeMatch =
          /^Số\s*:?\s*(\d+|\.{2,}|…)\s*[-/]\s*([A-ZĐa-zđ]+)\s*[-/]\s*([A-ZĐa-zđ0-9]+)$/i.exec(
            trimmed.split(/\t/)[0].trim()
          );
        if (codeMatch) {
          header.docNumber = codeMatch[1];
          header.docTypeAbbr = codeMatch[2].toUpperCase();
          header.docOrgAbbr = codeMatch[3].toUpperCase();
        }
      } else if (isDateLine) {
        node.role = ParagraphRole.HEADER_DATE_LOCATION;
        header.dateLocationRaw = trimmed;
        header.dateParagraphIndex = i;

        const dateMatch =
          /^(?:([^,]+),\s*)?ngày\s*([^\s]+)\s*tháng\s*([^\s]+)\s*năm\s*(\d{4}|\.{2,}|…)/i.exec(
            trimmed
          );
        if (dateMatch) {
          header.location = dateMatch[1]?.trim();
          header.dayRaw = dateMatch[2]?.trim();
          header.monthRaw = dateMatch[3]?.trim();
          header.yearRaw = dateMatch[4]?.trim();
        }
      } else if (isSeparatorLine(trimmed)) {
        node.role = ParagraphRole.HEADER_SEPARATOR_LINE;
      } else if (isOrganCandidate) {
        if (!header.issuingOrgan) {
          header.issuingOrgan = trimmed;
          header.issuingOrganParagraphIndex = i;
          node.role = ParagraphRole.HEADER_ISSUING_ORGAN;
        } else {
          // Đã có dòng cơ quan thứ nhất -> dòng 1 là Cấp trên, dòng 2 mới là Cơ quan ban hành
          header.superiorOrgan = header.issuingOrgan;
          if (header.issuingOrganParagraphIndex !== undefined) {
            nodes[header.issuingOrganParagraphIndex].role =
              ParagraphRole.HEADER_SUPERIOR_ORGAN;
          }
          header.issuingOrgan = trimmed.replace(/\*+$/, "").trim();
          header.issuingOrganParagraphIndex = i;
          node.role = ParagraphRole.HEADER_ISSUING_ORGAN;
        }
      }
    } else if (/^(V\/v|Về việc)\s+/i.test(trimmed) && i <= 8) {
      node.zone = DocumentZone.TITLE;
      node.role = ParagraphRole.CV_SUMMARY;
      docSummary = trimmed;
    }
  }

  // ------------------------------------------------------------------
  // BƯỚC 2: NHẬN DIỆN VÙNG FOOTER (NƠI NHẬN & CHỮ KÝ) Ở CUỐI VĂN BẢN
  // ------------------------------------------------------------------
  let footerStartIndex = nodes.length;
  const minFooterScan = Math.max(lastHeaderIndex + 1, nodes.length - 25);

  for (let i = minFooterScan; i < nodes.length; i++) {
    const trimmed = nodes[i].text.trim();
    if (!trimmed) continue;

    if (
      /^Nơi nhận\s*:/i.test(trimmed) ||
      SIGN_AUTHORITY_REGEX.test(trimmed) ||
      isSignRoleLine(trimmed)
    ) {
      footerStartIndex = i;
      break;
    }
  }

  if (footerStartIndex < nodes.length) {
    let inRecipientsList = false;
    for (let i = footerStartIndex; i < nodes.length; i++) {
      const node = nodes[i];
      const trimmed = node.text.trim();
      if (!trimmed) continue;

      node.zone = DocumentZone.FOOTER;
      if (node.tableNestingLevel > 0) {
        footer.hasSignatureTable = true;
      }

      if (/^Nơi nhận\s*:/i.test(trimmed)) {
        node.role = ParagraphRole.RECIPIENT_HEADER;
        footer.recipientHeaderIndex = i;
        inRecipientsList = true;
      } else if (SIGN_AUTHORITY_REGEX.test(trimmed)) {
        node.role = ParagraphRole.SIGN_AUTHORITY;
        footer.signAuthority = trimmed;
        footer.signAuthorityIndex = i;
        inRecipientsList = false;
      } else if (isSignRoleLine(trimmed)) {
        if (!footer.signAuthority && !footer.signRole) {
          // Trường hợp thủ trưởng ký trực tiếp (ví dụ: CHÁNH VĂN PHÒNG hoặc BÍ THƯ)
          node.role = ParagraphRole.SIGN_ROLE;
          footer.signRole = trimmed;
          footer.signRoleIndex = i;
        } else {
          node.role = ParagraphRole.SIGN_ROLE;
          footer.signRole = trimmed;
          footer.signRoleIndex = i;
        }
        inRecipientsList = false;
      } else if (
        inRecipientsList &&
        (trimmed.startsWith("-") ||
          trimmed.startsWith("–") ||
          /^Như\s+Điều/i.test(trimmed) ||
          /^Lưu\s*:/i.test(trimmed) ||
          /^Lưu\s+VT/i.test(trimmed))
      ) {
        node.role = ParagraphRole.RECIPIENT_ITEM;
        footer.recipientItems.push({ index: i, text: trimmed });
      } else if (
        (footer.signAuthority || footer.signRole) &&
        !footer.signerName &&
        !trimmed.startsWith("-")
      ) {
        node.role = ParagraphRole.SIGN_NAME;
        footer.signerName = trimmed;
        footer.signerNameIndex = i;
      }
    }
  }

  // ------------------------------------------------------------------
  // BƯỚC 3: PHÂN LOẠI VÙNG TITLE, SALUTATION, LEGAL_BASIS & BODY
  // ------------------------------------------------------------------
  let bodyStartIdx = lastHeaderIndex + 1;
  let expectingSummary = false;

  for (let i = bodyStartIdx; i < footerStartIndex; i++) {
    const node = nodes[i];
    const trimmed = node.text.trim();
    if (!trimmed) continue;

    // Nếu đoạn văn đã được gán CV_SUMMARY ở bước 1 thì bỏ qua
    if (node.role === ParagraphRole.CV_SUMMARY) continue;

    // Nếu nằm trong bảng dữ liệu ở giữa văn bản -> bảo vệ bảng dữ liệu
    if (node.tableNestingLevel > 0) {
      node.zone = DocumentZone.BODY;
      node.role = ParagraphRole.DATA_TABLE_CELL;
      continue;
    }

    // 1. Nhận diện Tên loại văn bản (NGHỊ QUYẾT, QUYẾT ĐỊNH, KẾ HOẠCH, BÁO CÁO...)
    if (docTypeParagraphIndex === undefined && i <= lastHeaderIndex + 6 && isExactDocType(trimmed)) {
      node.zone = DocumentZone.TITLE;
      node.role = ParagraphRole.DOC_TYPE;
      docTypeTitle = trimmed.toUpperCase();
      docTypeParagraphIndex = i;
      expectingSummary = true;
      continue;
    }

    // 2. Nhận diện Trích yếu văn bản ngay dưới Tên loại văn bản
    if (expectingSummary) {
      if (isSeparatorLine(trimmed)) {
        node.zone = DocumentZone.TITLE;
        node.role = ParagraphRole.TITLE_SEPARATOR_LINE;
        expectingSummary = false;
        continue;
      }
      if (
        !/^Kính gửi\s*:/i.test(trimmed) &&
        !/^(-?\s*Căn cứ|Xét đề nghị)/i.test(trimmed) &&
        !/^([IVXLCDM]+)\s*[.\-]\s+/i.test(trimmed) &&
        !/^Điều\s+\d+/i.test(trimmed)
      ) {
        node.zone = DocumentZone.TITLE;
        node.role = ParagraphRole.DOC_SUMMARY;
        docSummary = docSummary ? `${docSummary} ${trimmed}` : trimmed;
        continue;
      } else {
        expectingSummary = false;
      }
    }

    // 3. Dòng kẻ ngang dưới Trích yếu
    if (i <= (docTypeParagraphIndex ?? lastHeaderIndex) + 4 && isSeparatorLine(trimmed)) {
      node.zone = DocumentZone.TITLE;
      node.role = ParagraphRole.TITLE_SEPARATOR_LINE;
      continue;
    }

    // 4. Dòng Kính gửi
    if (/^(Kính gửi|Đồng kính gửi)\s*:/i.test(trimmed)) {
      node.zone = DocumentZone.SALUTATION;
      node.role = ParagraphRole.SALUTATION_LINE;
      continue;
    }

    // 5. Dòng QUYẾT ĐỊNH: (trong văn bản Quyết định)
    if (/^QUYẾT ĐỊNH\s*:$/i.test(trimmed)) {
      node.zone = DocumentZone.LEGAL_BASIS;
      node.role = ParagraphRole.DECISION_VERB;
      hasDecisionVerb = true;
      continue;
    }

    // 6. Căn cứ pháp lý (Căn cứ..., - Căn cứ..., Xét đề nghị...)
    if (
      !hasDecisionVerb &&
      /^([\-–+•]\s*)?(Căn cứ|Thực hiện|Theo đề nghị|Xét đề nghị)\s+/i.test(trimmed)
    ) {
      node.zone = DocumentZone.LEGAL_BASIS;
      node.role = ParagraphRole.LEGAL_BASIS_LINE;
      continue;
    }

    // 7. Dấu kết thúc bài (----- hoặc * * *)
    if (isSeparatorLine(trimmed) && i >= footerStartIndex - 3) {
      node.zone = DocumentZone.BODY;
      node.role = ParagraphRole.END_MARKER;
      continue;
    }

    // 8. Nhận diện Cây đề mục (Headings) trong BODY
    node.zone = DocumentZone.BODY;

    // 8.1: PHẦN / CHƯƠNG
    const partMatch = /^(PHẦN|CHƯƠNG)\s+([IVXLCDM\d]+|THỨ\s+[^\s:.\-]+)(.*)$/i.exec(trimmed);
    if (partMatch) {
      node.role = ParagraphRole.HEADING_PART;
      continue;
    }

    // 8.2: Điều 1., Điều 2.
    const articleMatch = /^Điều\s+(\d+)\s*([.:]?)\s*(.*)$/i.exec(trimmed);
    if (articleMatch) {
      node.role = ParagraphRole.ARTICLE_HEADING;
      node.headingMeta = {
        level: "ARTICLE",
        rawIndex: articleMatch[1],
        numericValue: parseInt(articleMatch[1], 10),
        separator: articleMatch[2] || ".",
      };
      continue;
    }

    // 8.3: Mục La Mã (I., II., I-, II-)
    const romanMatch = /^([IVXLCDM]+)\s*([.\-])\s+(.+)$/.exec(trimmed);
    if (romanMatch) {
      const val = romanToInt(romanMatch[1]);
      if (!isNaN(val)) {
        node.role = ParagraphRole.HEADING_ROMAN;
        node.headingMeta = {
          level: "ROMAN",
          rawIndex: romanMatch[1],
          numericValue: val,
          separator: romanMatch[2],
        };
        continue;
      }
    }

    // 8.4: Mục con đa cấp (1.1., 1.2.)
    const subNumMatch = /^(\d+\.\d+(?:\.\d+)?)\.?\s+(.+)$/.exec(trimmed);
    if (subNumMatch) {
      node.role = ParagraphRole.HEADING_SUBNUMBER;
      node.headingMeta = {
        level: "SUBNUMBER",
        rawIndex: subNumMatch[1],
        numericValue: parseFloat(subNumMatch[1]),
      };
      continue;
    }

    // 8.5: Mục số Ả Rập (1., 2., 3.)
    const numMatch = /^(\d+)\.\s+(.+)$/.exec(trimmed);
    if (numMatch) {
      node.role = ParagraphRole.HEADING_NUMBER;
      node.headingMeta = {
        level: "NUMBER",
        rawIndex: numMatch[1],
        numericValue: parseInt(numMatch[1], 10),
        separator: ".",
      };
      continue;
    }

    // 8.6: Mục chữ cái a), b), c)
    const letterMatch = /^([a-zđ])\)\s+(.+)$/i.exec(trimmed);
    if (letterMatch) {
      node.role = ParagraphRole.HEADING_LETTER;
      node.headingMeta = {
        level: "LETTER",
        rawIndex: letterMatch[1].toLowerCase(),
        numericValue: letterMatch[1].toLowerCase().charCodeAt(0) - 96,
        separator: ")",
      };
      continue;
    }

    // 8.7: Gạch đầu dòng liệt kê (-, +, •)
    if (/^[\-–+•]\s+/.test(trimmed)) {
      node.role = ParagraphRole.LIST_ITEM;
      continue;
    }

    node.role = ParagraphRole.BODY_PARAGRAPH;
  }

  return {
    nodes,
    header,
    docTypeTitle,
    docTypeParagraphIndex,
    docSummary,
    hasDecisionVerb,
    footer,
  };
}

/**
 * Đọc trực tiếp các đoạn văn từ Word.RequestContext và trả về `ParsedDocument`
 */
export async function parseWordDocument(
  context: Word.RequestContext
): Promise<{ parsed: ParsedDocument; wordParagraphs: Word.ParagraphCollection }> {
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
    "items/font/bold",
    "items/font/italic",
    "items/font/color",
  ]);
  await context.sync();

  const rawNodes: RawParagraphNode[] = paragraphs.items.map((p, idx) => ({
    index: idx,
    text: p.text,
    tableNestingLevel: p.tableNestingLevel,
    alignment: String(p.alignment),
    firstLineIndent: p.firstLineIndent,
    lineSpacing: p.lineSpacing,
    spaceBefore: p.spaceBefore,
    spaceAfter: p.spaceAfter,
    fontName: p.font?.name,
    fontSize: p.font?.size,
    fontBold: p.font?.bold,
    fontItalic: p.font?.italic,
    fontColor: p.font?.color,
  }));

  return {
    parsed: parseDocumentStructure(rawNodes),
    wordParagraphs: paragraphs,
  };
}
