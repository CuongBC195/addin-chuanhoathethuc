/**
 * src/index.ts
 * Module chuẩn hóa thể thức và kiểm tra logic văn bản của Đảng
 * Căn cứ: Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng
 */

export * from "./constants";
export * from "./types";
export * from "./documentParser";
export * from "./formatService";

import {
  applyPageSetup,
  formatBodyContent,
  insertHeaderTable,
  insertSignatureTable,
  lintDocument,
} from "./formatService";

import {
  HeaderData,
  SignData,
  FormatBodyOptions,
  PageSetupOptions,
  LintReport,
} from "./types";

export interface StandardizeConfig {
  fontSize?: number;
  pageSetupOptions?: PageSetupOptions;
  bodyOptions?: FormatBodyOptions;
  headerData?: HeaderData;
  signData?: SignData;
}

export async function standardizeDocument(
  context: Word.RequestContext,
  config?: StandardizeConfig
): Promise<LintReport> {
  await applyPageSetup(context, config?.pageSetupOptions);
  await formatBodyContent(context, config?.fontSize, config?.bodyOptions);

  if (config?.headerData) {
    await insertHeaderTable(context, config.headerData);
  }

  if (config?.signData) {
    await insertSignatureTable(context, config.signData);
  }

  const report = await lintDocument(context);
  return report;
}
