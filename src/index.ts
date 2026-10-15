/**
 * doc-party-addin
 * Module chuẩn hóa thể thức và kỹ thuật trình bày văn bản của Đảng
 * Căn cứ: Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng
 */

export * from "./constants";
export * from "./types";
export * from "./formatService";

import {
  applyPageSetup,
  formatBodyContent,
  insertHeaderTable,
  insertSignatureTable,
  lintDocument,
} from "./formatService";

import { HeaderData, SignData, FormatBodyOptions, PageSetupOptions, LintReport } from "./types";

export interface StandardizeConfig {
  /** Cỡ chữ nội dung văn bản (14 pt hoặc 13 pt) */
  fontSize?: number;
  /** Cấu hình tùy biến thiết lập trang */
  pageSetupOptions?: PageSetupOptions;
  /** Cấu hình tùy biến đoạn nội dung */
  bodyOptions?: FormatBodyOptions;
  /** Dữ liệu chèn bảng Header (nếu muốn chèn mới) */
  headerData?: HeaderData;
  /** Dữ liệu chèn bảng Chữ ký (nếu muốn chèn mới) */
  signData?: SignData;
}

/**
 * Hàm tích hợp (All-in-one): Chuẩn hóa toàn diện văn bản đang mở
 * Thực hiện theo thứ tự:
 * 1. Áp dụng Page Setup (A4, căn lề 4 hướng, mirrorMargins, không số trang đầu).
 * 2. Chuẩn hóa toàn bộ nội dung văn bản (Times New Roman, Justified, 10mm indent, 18pt line spacing, 6pt space after).
 * 3. Chèn bảng Header (nếu có dữ liệu headerData).
 * 4. Chèn bảng Chữ ký & Nơi nhận (nếu có dữ liệu signData).
 * 5. Chạy Lint quét kiểm tra toàn văn bản và trả về báo cáo kết quả.
 */
export async function standardizeDocument(
  context: Word.RequestContext,
  config?: StandardizeConfig
): Promise<LintReport> {
  // 1. Áp dụng Page Setup
  await applyPageSetup(context, config?.pageSetupOptions);

  // 2. Chuẩn hóa Body
  await formatBodyContent(context, config?.fontSize, config?.bodyOptions);

  // 3. Chèn Header Table nếu có dữ liệu
  if (config?.headerData) {
    await insertHeaderTable(context, config.headerData);
  }

  // 4. Chèn Signature Table nếu có dữ liệu
  if (config?.signData) {
    await insertSignatureTable(context, config.signData);
  }

  // 5. Quét kiểm tra kết quả
  const report = await lintDocument(context);
  return report;
}
