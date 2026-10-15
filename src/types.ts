/**
 * src/types.ts
 * Kiểu dữ liệu, Phân vùng AST (Zones/Roles) và Interfaces cho Word Web Add-in
 * Chuẩn hóa thể thức & rà soát logic văn bản Đảng theo Hướng dẫn số 05-HD/VPTW
 */

export enum LintSeverity {
  ERROR = "error",
  WARNING = "warning",
  INFO = "info",
}

/** Danh mục mã lỗi thể thức & logic văn bản Đảng */
export enum LintIssueType {
  // Nhóm lỗi thể thức cơ bản
  MISSING_LEADING_ZERO_NUMBER = "MISSING_LEADING_ZERO_NUMBER",
  MISSING_LEADING_ZERO_DATE = "MISSING_LEADING_ZERO_DATE",
  EXCESS_LEADING_ZERO_DATE = "EXCESS_LEADING_ZERO_DATE",
  INVALID_DATE_VALUE = "INVALID_DATE_VALUE",
  SPACE_OR_TAB_INDENTATION = "SPACE_OR_TAB_INDENTATION",
  NON_STANDARD_FONT = "NON_STANDARD_FONT",
  INVALID_MARGIN = "INVALID_MARGIN",
  INVALID_ALIGNMENT = "INVALID_ALIGNMENT",
  INVALID_FIRST_LINE_INDENT = "INVALID_FIRST_LINE_INDENT",
  INVALID_LINE_SPACING = "INVALID_LINE_SPACING",
  INVALID_PARAGRAPH_SPACING = "INVALID_PARAGRAPH_SPACING",
  HEADER_FOOTER_STYLE_ERROR = "HEADER_FOOTER_STYLE_ERROR",

  // Nhóm lỗi logic diện rộng & đối chiếu chéo
  INVALID_DOC_CODE_FORMAT = "INVALID_DOC_CODE_FORMAT",
  DOC_TYPE_CODE_MISMATCH = "DOC_TYPE_CODE_MISMATCH",
  AUTHORITY_SIGNATURE_MISMATCH = "AUTHORITY_SIGNATURE_MISMATCH",
  INVALID_SIGN_PREFIX = "INVALID_SIGN_PREFIX",
  HEADING_SEQUENCE_ERROR = "HEADING_SEQUENCE_ERROR",
  RECIPIENTS_FORMAT_ERROR = "RECIPIENTS_FORMAT_ERROR",
  UNFILLED_PLACEHOLDER = "UNFILLED_PLACEHOLDER",
  TYPOGRAPHY_PUNCTUATION_ERROR = "TYPOGRAPHY_PUNCTUATION_ERROR",
}

/** Các phân vùng cấu trúc chính của một văn bản Đảng */
export enum DocumentZone {
  HEADER = "HEADER",
  TITLE = "TITLE",
  SALUTATION = "SALUTATION",
  LEGAL_BASIS = "LEGAL_BASIS",
  BODY = "BODY",
  FOOTER = "FOOTER",
}

/** Vai trò ngữ nghĩa chi tiết của từng đoạn văn (Paragraph Role) */
export enum ParagraphRole {
  EMPTY = "EMPTY",

  // Vùng HEADER (Ô 1 & Ô 2)
  HEADER_SUPERIOR_ORGAN = "HEADER_SUPERIOR_ORGAN",
  HEADER_ISSUING_ORGAN = "HEADER_ISSUING_ORGAN",
  HEADER_PARTY_MOTTO = "HEADER_PARTY_MOTTO",
  HEADER_SEPARATOR_LINE = "HEADER_SEPARATOR_LINE",
  HEADER_DOC_NUMBER = "HEADER_DOC_NUMBER",
  HEADER_DATE_LOCATION = "HEADER_DATE_LOCATION",
  HEADER_RAW_COMBINED = "HEADER_RAW_COMBINED",

  // Vùng TITLE & MỞ ĐẦU
  DOC_TYPE = "DOC_TYPE",
  DOC_SUMMARY = "DOC_SUMMARY",
  CV_SUMMARY = "CV_SUMMARY",
  TITLE_SEPARATOR_LINE = "TITLE_SEPARATOR_LINE",
  SALUTATION_LINE = "SALUTATION_LINE",
  LEGAL_BASIS_LINE = "LEGAL_BASIS_LINE",
  DECISION_VERB = "DECISION_VERB",

  // Vùng BODY (Nội dung & Cây đề mục)
  HEADING_PART = "HEADING_PART",
  HEADING_ROMAN = "HEADING_ROMAN",
  HEADING_NUMBER = "HEADING_NUMBER",
  HEADING_SUBNUMBER = "HEADING_SUBNUMBER",
  HEADING_LETTER = "HEADING_LETTER",
  ARTICLE_HEADING = "ARTICLE_HEADING",
  LIST_ITEM = "LIST_ITEM",
  BODY_PARAGRAPH = "BODY_PARAGRAPH",
  DATA_TABLE_CELL = "DATA_TABLE_CELL",
  END_MARKER = "END_MARKER",

  // Vùng FOOTER (Nơi nhận & Chữ ký)
  RECIPIENT_HEADER = "RECIPIENT_HEADER",
  RECIPIENT_ITEM = "RECIPIENT_ITEM",
  SIGN_AUTHORITY = "SIGN_AUTHORITY",
  SIGN_ROLE = "SIGN_ROLE",
  SIGN_NAME = "SIGN_NAME",
}

/** Thông tin thô của một đoạn văn đọc từ Word hoặc truyền từ Unit Test */
export interface RawParagraphNode {
  index: number;
  text: string;
  tableNestingLevel: number;
  alignment?: string;
  firstLineIndent?: number;
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  fontName?: string;
  fontSize?: number;
  fontBold?: boolean;
  fontItalic?: boolean;
  fontColor?: string;
}

/** Nút đoạn văn sau khi đã được bộ Parser phân vùng và gán nhãn */
export interface ParsedParagraphNode extends RawParagraphNode {
  zone: DocumentZone;
  role: ParagraphRole;
  headingMeta?: {
    level: "PART" | "ROMAN" | "NUMBER" | "SUBNUMBER" | "LETTER" | "ARTICLE";
    rawIndex: string;
    numericValue: number;
    separator?: string;
  };
}

/** Dữ liệu bóc tách từ vùng Header (Ô 1 & Ô 2) */
export interface ExtractedHeaderInfo {
  hasHeaderTable: boolean;
  superiorOrgan?: string;
  issuingOrgan?: string;
  issuingOrganParagraphIndex?: number;
  partyMottoParagraphIndex?: number;
  docNumberRaw?: string;
  docNumberParagraphIndex?: number;
  docNumber?: string;
  docTypeAbbr?: string;
  docOrgAbbr?: string;
  dateLocationRaw?: string;
  dateParagraphIndex?: number;
  location?: string;
  dayRaw?: string;
  monthRaw?: string;
  yearRaw?: string;
}

/** Dữ liệu bóc tách từ vùng Footer (Nơi nhận & Chữ ký) */
export interface ExtractedFooterInfo {
  hasSignatureTable: boolean;
  recipientHeaderIndex?: number;
  recipientItems: { index: number; text: string }[];
  signAuthority?: string;
  signAuthorityIndex?: number;
  signRole?: string;
  signRoleIndex?: number;
  signerName?: string;
  signerNameIndex?: number;
}

/** Cấu trúc AST toàn bộ tài liệu sau khi phân tích */
export interface ParsedDocument {
  nodes: ParsedParagraphNode[];
  header: ExtractedHeaderInfo;
  docTypeTitle?: string;
  docTypeParagraphIndex?: number;
  docSummary?: string;
  hasDecisionVerb: boolean;
  footer: ExtractedFooterInfo;
}

export interface HeaderData {
  superiorOrgan?: string;
  issuingOrgan: string;
  docNumber: number | string;
  docCode: string;
  location: string;
  date: Date | { day: number; month: number; year: number } | string;
  fontSize?: 13 | 14;
}

export interface SignData {
  recipients: string[];
  signTitle: string;
  signerRole?: string;
  signerName: string;
  signatureSpaceMm?: number;
  fontSize?: 13 | 14;
}

export interface FormatBodyOptions {
  fontSize?: number;
  lineSpacing?: number;
  spaceAfter?: number;
  spaceBefore?: number;
  firstLineIndentMm?: number;
  preserveCenteredHeadings?: boolean;
}

export interface PageSetupOptions {
  topMarginMm?: number;
  bottomMarginMm?: number;
  leftMarginMm?: number;
  rightMarginMm?: number;
  mirrorMargins?: boolean;
  differentFirstPageHeaderFooter?: boolean;
  headerDistanceMm?: number;
}

export interface LintIssue {
  id: string;
  type: LintIssueType;
  severity: LintSeverity;
  message: string;
  suggestion: string;
  locationInfo?: string;
  snippet?: string;
  /** Chỉ số đoạn văn trong Word để click nhảy thẳng tới vị trí lỗi */
  paragraphIndex?: number;
}

export interface LintReport {
  isValid: boolean;
  totalIssues: number;
  errorCount: number;
  warningCount: number;
  infoCount: number;
  issues: LintIssue[];
  timestamp: string;
}
