/**
 * src/constants.ts
 * Quy chuẩn Kỹ thuật, Thể thức và Từ điển Logic Văn bản của Đảng Cộng sản Việt Nam
 * Căn cứ Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng
 */

/** Tỉ lệ chuyển đổi đơn vị chuẩn: 1 mm ≈ 2.83465 pt (points) */
export const MM_TO_PT = 2.83465;

export function mmToPt(mm: number): number {
  return Number((mm * MM_TO_PT).toFixed(4));
}

export function ptToMm(pt: number): number {
  return Number((pt / MM_TO_PT).toFixed(4));
}

/**
 * 1. Khổ giấy & Căn lề trang (Page Setup) theo Hướng dẫn 05-HD/VPTW
 */
export const PAGE_SETUP_CONSTANTS = {
  A4_WIDTH_MM: 210,
  A4_WIDTH_PT: mmToPt(210),

  A4_HEIGHT_MM: 297,
  A4_HEIGHT_PT: mmToPt(297),

  MARGIN_TOP_MM: 20,
  MARGIN_TOP_PT: mmToPt(20),

  MARGIN_BOTTOM_MM: 20,
  MARGIN_BOTTOM_PT: mmToPt(20),

  MARGIN_LEFT_MM: 30,
  MARGIN_LEFT_PT: mmToPt(30),

  MARGIN_RIGHT_MM: 15,
  MARGIN_RIGHT_PT: mmToPt(15),

  HEADER_DISTANCE_MM: 10,
  HEADER_DISTANCE_PT: mmToPt(10),

  FOOTER_DISTANCE_MM: 10,
  FOOTER_DISTANCE_PT: mmToPt(10),

  PRINTABLE_WIDTH_MM: 165,
  PRINTABLE_WIDTH_PT: mmToPt(165),

  MIRROR_MARGINS: true,
  DIFFERENT_FIRST_PAGE: true,
} as const;

/**
 * 2. Quy chuẩn Font & Định dạng Đoạn văn (Typography & Paragraphs)
 */
export const TYPOGRAPHY_CONSTANTS = {
  FONT_FAMILY: "Times New Roman",
  FONT_COLOR: "#000000",

  BODY_FONT_SIZE_DEFAULT: 14,
  BODY_FONT_SIZE_ALTERNATIVE: 13,
  DOC_TYPE_FONT_SIZE: 15,

  /** Chuẩn 05-HD/VPTW: 10 mm (~28.35 pt), chấp nhận biên độ thực tế 10mm - 12.7mm khi quét lỗi */
  FIRST_LINE_INDENT_MM: 10,
  FIRST_LINE_INDENT_PT: mmToPt(10),
  FIRST_LINE_INDENT_MAX_MM: 12.7,
  FIRST_LINE_INDENT_MAX_PT: mmToPt(12.7),

  LINE_SPACING_DEFAULT_PT: 18,
  LINE_SPACING_MIN_PT: 18,
  LINE_SPACING_MAX_PT: 22,

  SPACE_BEFORE_PT: 0,
  SPACE_AFTER_MIN_PT: 6,
  SPACE_AFTER_DEFAULT_PT: 6,
} as const;

/**
 * 3. Quy chuẩn Header & Ô thông tin đầu văn bản (Header Table)
 */
export const HEADER_CONSTANTS = {
  PARTY_HEADER_TEXT: "ĐẢNG CỘNG SẢN VIỆT NAM",
  SUPERIOR_ORGAN_FONT_SIZE: 13,
  ISSUING_ORGAN_FONT_SIZE: 13,
  DOC_NUMBER_FONT_SIZE: 13,
  PARTY_TITLE_FONT_SIZE: 13,
  DATE_LOCATION_FONT_SIZE: 13,
  LINE_ISSUING_ORGAN: "— — — —",
  LINE_PARTY_TITLE: "—————————————",
  COL_LEFT_PERCENT: 40,
  COL_RIGHT_PERCENT: 60,
} as const;

/**
 * 4. Quy chuẩn Footer (Ô Nơi nhận & Ô Chữ ký)
 */
export const FOOTER_CONSTANTS = {
  RECIPIENTS_LABEL: "Nơi nhận:",
  RECIPIENTS_LABEL_FONT_SIZE: 11,
  RECIPIENTS_ITEM_FONT_SIZE: 11,
  SIGNER_TITLE_FONT_SIZE: 13,
  SIGNER_NAME_FONT_SIZE: 13,
  SIGNATURE_SPACE_MM: 38,
  SIGNATURE_SPACE_PT: mmToPt(38),
  COL_LEFT_PERCENT: 45,
  COL_RIGHT_PERCENT: 55,
} as const;

/**
 * 5. Quy chuẩn Đánh số trang (Page Numbering)
 */
export const PAGE_NUMBERING_CONSTANTS = {
  ALIGNMENT: "Centered" as const,
  FONT_SIZE: 13.5,
  SUPPRESS_FIRST_PAGE: true,
} as const;

/**
 * 6. Bảng ánh xạ chữ viết tắt tên loại văn bản Đảng (Hướng dẫn 05-HD/VPTW)
 * Dùng để đối chiếu chéo giữa [Số, ký hiệu] và [Tên loại văn bản]
 */
export const PARTY_DOC_CODE_MAP: Record<string, string> = {
  NQ: "NGHỊ QUYẾT",
  "QĐ": "QUYẾT ĐỊNH",
  QD: "QUYẾT ĐỊNH",
  CT: "CHỈ THỊ",
  KL: "KẾT LUẬN",
  QC: "QUY CHẾ",
  "QYĐ": "QUY ĐỊNH",
  QYD: "QUY ĐỊNH",
  HD: "HƯỚNG DẪN",
  TB: "THÔNG BÁO",
  BC: "BÁO CÁO",
  KH: "KẾ HOẠCH",
  CTR: "CHƯƠNG TRÌNH",
  TTR: "TỜ TRÌNH",
  CV: "CÔNG VĂN",
  GM: "GIẤY MỜI",
  BB: "BIÊN BẢN",
  ĐA: "ĐỀ ÁN",
  PA: "PHƯƠNG ÁN",
};

/**
 * Danh sách từ khóa nhận diện cơ quan ban hành làm việc theo chế độ TẬP THỂ
 * (Khi ký bắt buộc phải có thẩm quyền "T/M ...")
 */
export const COLLECTIVE_ORGANS = [
  "BAN CHẤP HÀNH",
  "BAN THƯỜNG VỤ",
  "BAN BÍ THƯ",
  "BỘ CHÍNH TRỊ",
  "ĐẢNG ỦY",
  "HUYỆN ỦY",
  "THÀNH ỦY",
  "TỈNH ỦY",
  "THỊ ỦY",
  "QUẬN ỦY",
  "CHI BỘ",
  "ỦY BAN KIỂM TRA",
];

/**
 * Danh sách từ khóa nhận diện cơ quan tham mưu/giúp việc làm việc theo chế độ THỦ TRƯỞNG
 * (Người đứng đầu ký trực tiếp hoặc cấp phó ký thay "K/T", KHÔNG được ký "T/M")
 */
export const EXECUTIVE_ORGANS = [
  "VĂN PHÒNG",
  "BAN TỔ CHỨC",
  "BAN TUYÊN GIÁO",
  "BAN DÂN VẬN",
  "BAN NỘI CHÍNH",
  "BAN KINH TẾ",
  "BAN XÂY DỰNG ĐẢNG",
  "TRUNG TÂM CHÍNH TRỊ",
];
