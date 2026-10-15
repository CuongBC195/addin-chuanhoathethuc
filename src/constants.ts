/**
 * Quy chuẩn Kỹ thuật và Thể thức Văn bản của Đảng Cộng sản Việt Nam
 * Căn cứ Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng
 * (thay thế Hướng dẫn 36-HD/VPTW ngày 03/4/2018)
 */

/** Tỉ lệ chuyển đổi đơn vị chuẩn: 1 mm ≈ 2.83465 pt (points) */
export const MM_TO_PT = 2.83465;

/**
 * Chuyển đổi milimét sang point
 * @param mm Chiều dài tính theo mm
 * @returns Chiều dài tính theo pt
 */
export function mmToPt(mm: number): number {
  return Number((mm * MM_TO_PT).toFixed(4));
}

/**
 * Chuyển đổi point sang milimét
 * @param pt Chiều dài tính theo pt
 * @returns Chiều dài tính theo mm
 */
export function ptToMm(pt: number): number {
  return Number((pt / MM_TO_PT).toFixed(4));
}

/**
 * 1. Khổ giấy & Căn lề trang (Page Setup) theo Hướng dẫn 05-HD/VPTW
 */
export const PAGE_SETUP_CONSTANTS = {
  /** Chiều rộng khổ giấy A4: 210 mm (~595.28 pt) */
  A4_WIDTH_MM: 210,
  A4_WIDTH_PT: mmToPt(210),

  /** Chiều cao khổ giấy A4: 297 mm (~841.89 pt) */
  A4_HEIGHT_MM: 297,
  A4_HEIGHT_PT: mmToPt(297),

  /** Lề trên: 20 mm (~56.69 pt) */
  MARGIN_TOP_MM: 20,
  MARGIN_TOP_PT: mmToPt(20),

  /** Lề dưới: 20 mm (~56.69 pt) */
  MARGIN_BOTTOM_MM: 20,
  MARGIN_BOTTOM_PT: mmToPt(20),

  /** Lề trái: 30 mm (~85.04 pt) - Trang lẻ */
  MARGIN_LEFT_MM: 30,
  MARGIN_LEFT_PT: mmToPt(30),

  /** Lề phải: 15 mm (~42.52 pt) - Trang lẻ */
  MARGIN_RIGHT_MM: 15,
  MARGIN_RIGHT_PT: mmToPt(15),

  /** Khoảng cách Header so với mép trên đỉnh trang: 10 mm (~28.35 pt) */
  HEADER_DISTANCE_MM: 10,
  HEADER_DISTANCE_PT: mmToPt(10),

  /** Khoảng cách Footer so với mép dưới đáy trang: 10 mm (~28.35 pt) */
  FOOTER_DISTANCE_MM: 10,
  FOOTER_DISTANCE_PT: mmToPt(10),

  /** Chiều rộng vùng in khả dụng A4 (210 - 30 - 15 = 165 mm ~ 467.72 pt) */
  PRINTABLE_WIDTH_MM: 165,
  PRINTABLE_WIDTH_PT: mmToPt(165),

  /** Bật căn lề in 2 mặt (Mirror Margins) */
  MIRROR_MARGINS: true,

  /** Bật tính năng khác biệt Header/Footer trang đầu (không đánh số trang bìa/trang 1) */
  DIFFERENT_FIRST_PAGE: true,
} as const;

/**
 * 2. Quy chuẩn Font & Định dạng Đoạn văn (Typography & Paragraphs)
 */
export const TYPOGRAPHY_CONSTANTS = {
  /** Phông chữ tiêu chuẩn Unicode duy nhất */
  FONT_FAMILY: "Times New Roman",

  /** Màu chữ đen chuẩn */
  FONT_COLOR: "#000000",

  /** Cỡ chữ nội dung mặc định: 14 pt (tùy chọn 13 pt) */
  BODY_FONT_SIZE_DEFAULT: 14,
  BODY_FONT_SIZE_ALTERNATIVE: 13,

  /** Thụt lề dòng đầu tiên của đoạn văn bản: 1.27 cm (theo thực tế sử dụng) */
  FIRST_LINE_INDENT_MM: 12.7,
  FIRST_LINE_INDENT_PT: mmToPt(12.7),

  /** Giãn dòng (Line Spacing): chế độ Exactly từ 18 pt đến 22 pt */
  LINE_SPACING_DEFAULT_PT: 18,
  LINE_SPACING_MIN_PT: 18,
  LINE_SPACING_MAX_PT: 22,

  /** Giãn đoạn phía trên (Space Before): 6 pt */
  SPACE_BEFORE_PT: 6,

  /** Giãn đoạn phía dưới (Space After): tối thiểu 6 pt */
  SPACE_AFTER_MIN_PT: 6,
  SPACE_AFTER_DEFAULT_PT: 6,
} as const;

/**
 * 3. Quy chuẩn Header & Ô thông tin đầu văn bản (Header Table)
 */
export const HEADER_CONSTANTS = {
  /** Tên tiêu đề Đảng */
  PARTY_HEADER_TEXT: "ĐẢNG CỘNG SẢN VIỆT NAM",

  /** Cỡ chữ cấp trên: 13-14 pt, chữ in hoa, đứng */
  SUPERIOR_ORGAN_FONT_SIZE: 13,

  /** Cỡ chữ cơ quan ban hành: 13-14 pt, chữ in hoa, ĐẬM */
  ISSUING_ORGAN_FONT_SIZE: 13,

  /** Cỡ chữ số, ký hiệu: 13 pt, chữ thường, đứng */
  DOC_NUMBER_FONT_SIZE: 13,

  /** Cỡ chữ "ĐẢNG CỘNG SẢN VIỆT NAM": 13-14 pt, chữ in hoa, ĐẬM */
  PARTY_TITLE_FONT_SIZE: 13,

  /** Cỡ chữ địa danh, ngày tháng năm: 13-14 pt, chữ thường, NGHIÊNG */
  DATE_LOCATION_FONT_SIZE: 13,

  /** Ký tự đường kẻ ngang dưới tên cơ quan ban hành (độ dài 1/3 - 1/2) */
  LINE_ISSUING_ORGAN: "— — — —",

  /** Ký tự đường kẻ ngang dưới tiêu đề Đảng (độ dài liền bằng dòng chữ) */
  LINE_PARTY_TITLE: "—————————————",

  /** Chiều rộng tương đối Ô 1 (Cột trái) và Ô 2 (Cột phải) */
  COL_LEFT_PERCENT: 40,
  COL_RIGHT_PERCENT: 60,
} as const;

/**
 * 4. Quy chuẩn Footer (Ô Nơi nhận & Ô Chữ ký)
 */
export const FOOTER_CONSTANTS = {
  /** Tiêu đề nơi nhận */
  RECIPIENTS_LABEL: "Nơi nhận:",

  /** Cỡ chữ tiêu đề "Nơi nhận:": 11-12 pt, chữ thường, ĐẬM - NGHIÊNG */
  RECIPIENTS_LABEL_FONT_SIZE: 11,

  /** Cỡ chữ thành phần nơi nhận: 10-11 pt, chữ thường, đứng */
  RECIPIENTS_ITEM_FONT_SIZE: 11,

  /** Cỡ chữ chức vụ, thẩm quyền ký: 13-14 pt, chữ in hoa, ĐẬM */
  SIGNER_TITLE_FONT_SIZE: 13,

  /** Cỡ chữ họ và tên người ký: 13-14 pt, chữ thường/hoa chữ cái đầu, ĐẬM */
  SIGNER_NAME_FONT_SIZE: 13,

  /** Khoảng trống để ký tên và đóng dấu: 35 mm - 40 mm (chọn 38 mm ~ 107.72 pt) */
  SIGNATURE_SPACE_MM: 38,
  SIGNATURE_SPACE_PT: mmToPt(38),

  /** Tỉ lệ phân chia cột: Nơi nhận (~45%), Chữ ký (~55%) */
  COL_LEFT_PERCENT: 45,
  COL_RIGHT_PERCENT: 55,
} as const;

/**
 * 5. Quy chuẩn Đánh số trang (Page Numbering)
 */
export const PAGE_NUMBERING_CONSTANTS = {
  /** Vị trí: Ở giữa trên đỉnh trang */
  ALIGNMENT: "Centered" as const,

  /** Cỡ chữ số trang: 13-14 pt, chữ thường, đứng */
  FONT_SIZE: 13.5,

  /** Không hiển thị tại trang đầu tiên */
  SUPPRESS_FIRST_PAGE: true,
} as const;
