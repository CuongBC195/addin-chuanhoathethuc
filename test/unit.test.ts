/**
 * test/unit.test.ts
 * Kiểm thử đơn vị cho AST Zone Parser, Formatters và Bộ Rà Soát Logic 05-HD/VPTW
 */

import { mmToPt, ptToMm, PAGE_SETUP_CONSTANTS } from "../src/constants";
import {
  formatDocNumber,
  formatPartyDate,
  lintParsedDocument,
} from "../src/formatService";
import { parseDocumentStructure } from "../src/documentParser";
import { LintIssueType, RawParagraphNode, DocumentZone, ParagraphRole } from "../src/types";

function assert(condition: boolean, message: string) {
  if (!condition) {
    console.error(`❌ FAIL: ${message}`);
    throw new Error(`Test failed: ${message}`);
  }
  console.log(`✅ PASS: ${message}`);
}

console.log("=== BẮT ĐẦU KIỂM THỬ THỂ THỨC & LOGIC 05-HD/VPTW ===");

// 1. Kiểm thử chuyển đổi đơn vị
assert(Math.abs(mmToPt(10) - 28.3465) < 0.001, "10 mm ≈ 28.35 pt");
assert(Math.abs(ptToMm(PAGE_SETUP_CONSTANTS.MARGIN_LEFT_PT) - 30) < 0.001, "ptToMm(MARGIN_LEFT_PT) == 30 mm");

// 2. Kiểm thử chuẩn hóa số văn bản & ngày tháng (bao gồm cả xóa số 0 thừa ở tháng 3..9)
assert(formatDocNumber(5) === "05", "Số 5 -> '05'");
assert(formatDocNumber("") === "   ", "Xin số để trống -> '   '");
const date1 = formatPartyDate("Hà Nội", { day: 5, month: 2, year: 2026 });
assert(date1 === "Hà Nội, ngày 05 tháng 02 năm 2026", `Ngày 5 tháng 2 -> '${date1}'`);

const rawDateExcessZero = formatPartyDate("Thạnh Phú", "ngày 5 tháng 08 năm 2026");
assert(
  rawDateExcessZero === "Thạnh Phú, ngày 05 tháng 8 năm 2026",
  `Tự động thêm số 0 cho ngày 5 và xóa số 0 thừa ở tháng 08 -> '${rawDateExcessZero}'`
);

const rawDateXinNgay = formatPartyDate("Thạnh Phú", "ngày ... tháng ... năm 2026");
assert(
  rawDateXinNgay === "Thạnh Phú, ngày   tháng   năm 2026",
  `Xin ngày tháng tự động để 3 khoảng trắng -> '${rawDateXinNgay}'`
);

// 3. Kiểm thử AST Parser & Bộ Rà Soát Logic diện rộng trên một văn bản mẫu có gài nhiều lỗi logic
const sampleRawParagraphs: RawParagraphNode[] = [
  { index: 0, text: "HUYỆN ỦY THẠNH PHÚ", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 13, fontBold: false },
  { index: 1, text: "ĐẢNG ỦY XÃ THẠNH PHÚ", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 13, fontBold: true },
  { index: 2, text: "Số: 5/BC-ĐU", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 13 },
  { index: 3, text: "ĐẢNG CỘNG SẢN VIỆT NAM", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 13, fontBold: true },
  { index: 4, text: "Thạnh Phú, ngày 5 tháng 08 năm 2026", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 13, fontItalic: false },
  { index: 5, text: "KẾ HOẠCH", tableNestingLevel: 0, fontName: "Times New Roman", fontSize: 15, fontBold: true },
  { index: 6, text: "Triển khai công tác kiểm tra, giám sát quý III năm 2026", tableNestingLevel: 0, fontName: "Times New Roman", fontSize: 14, fontBold: true },
  { index: 7, text: "-----", tableNestingLevel: 0, fontName: "Times New Roman", fontSize: 12 },
  { index: 8, text: "I. MỤC ĐÍCH, YÊU CẦU", tableNestingLevel: 0, fontName: "Times New Roman", fontSize: 14, alignment: "Justified", firstLineIndent: 28.35 },
  { index: 9, text: "1. Mục đích thực hiện...", tableNestingLevel: 0, fontName: "Times New Roman", fontSize: 14, alignment: "Justified", firstLineIndent: 28.35 },
  { index: 10, text: "3. Yêu cầu triển khai thực hiện.", tableNestingLevel: 0, fontName: "Times New Roman", fontSize: 14, alignment: "Justified", firstLineIndent: 28.35 },
  { index: 11, text: "III- NỘI DUNG THỰC HIỆN", tableNestingLevel: 0, fontName: "Times New Roman", fontSize: 14, alignment: "Justified", firstLineIndent: 28.35 },
  { index: 12, text: "Nơi nhận:", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 11, fontBold: true, fontItalic: true },
  { index: 13, text: "- Thường trực Huyện ủy (để b/c)", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 11 },
  { index: 14, text: "- Lưu Văn thư", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 11 },
  { index: 15, text: "BÍ THƯ", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 13, fontBold: true },
  { index: 16, text: "Nguyễn Văn A", tableNestingLevel: 1, fontName: "Times New Roman", fontSize: 13, fontBold: true },
];

const parsed = parseDocumentStructure(sampleRawParagraphs);

assert(parsed.header.superiorOrgan === "HUYỆN ỦY THẠNH PHÚ", "Parser nhận diện đúng Cơ quan cấp trên");
assert(parsed.header.issuingOrgan === "ĐẢNG ỦY XÃ THẠNH PHÚ", "Parser nhận diện đúng Cơ quan ban hành");
assert(parsed.docTypeTitle === "KẾ HOẠCH", "Parser nhận diện đúng Tên loại văn bản");
assert(parsed.nodes[6].role === ParagraphRole.DOC_SUMMARY, "Parser nhận diện đúng dòng Trích yếu dưới Tên loại văn bản");
assert(parsed.nodes[12].zone === DocumentZone.FOOTER, "Parser nhận diện đúng vùng FOOTER trong bảng");

const issues = lintParsedDocument(parsed);
const issueTypes = new Set(issues.map((i) => i.type));

assert(issueTypes.has(LintIssueType.INVALID_DOC_CODE_FORMAT), "Bắt được lỗi viết 'Số: 5/BC-ĐU' kiểu hành chính Nhà nước");
assert(issueTypes.has(LintIssueType.MISSING_LEADING_ZERO_NUMBER), "Bắt được lỗi thiếu số 0 ở Số 5");
assert(issueTypes.has(LintIssueType.MISSING_LEADING_ZERO_DATE), "Bắt được lỗi thiếu số 0 ở ngày 5");
assert(issueTypes.has(LintIssueType.EXCESS_LEADING_ZERO_DATE), "Bắt được lỗi thừa số 0 ở tháng 08");
assert(issueTypes.has(LintIssueType.DOC_TYPE_CODE_MISMATCH), "Bắt được lỗi mâu thuẫn ký hiệu 'BC' nhưng tiêu đề là 'KẾ HOẠCH'");
assert(issueTypes.has(LintIssueType.HEADING_SEQUENCE_ERROR), "Bắt được lỗi nhảy cóc đề mục (I -> III và 1. -> 3.)");
assert(issueTypes.has(LintIssueType.AUTHORITY_SIGNATURE_MISMATCH), "Bắt được lỗi cơ quan tập thể 'ĐẢNG ỦY XÃ' nhưng chữ ký thiếu 'T/M'");
assert(issueTypes.has(LintIssueType.RECIPIENTS_FORMAT_ERROR), "Bắt được lỗi thiếu dấu ';' và '.' ở phần Nơi nhận");
assert(issueTypes.has(LintIssueType.UNFILLED_PLACEHOLDER), "Bắt được lỗi còn dấu '...' chưa điền xong");
assert(issueTypes.has(LintIssueType.HEADER_FOOTER_STYLE_ERROR), "Bắt được lỗi ngày tháng trong bảng Header quên in nghiêng");

console.log("=== TẤT CẢ KIỂM THỬ THỂ THỨC & LOGIC ĐÃ THÀNH CÔNG 100% ===");
