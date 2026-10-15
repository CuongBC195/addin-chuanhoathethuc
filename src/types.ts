/**
 * Kiểu dữ liệu và Interfaces cho Word Web Add-in
 * Chuẩn hóa thể thức văn bản Đảng theo Hướng dẫn số 05-HD/VPTW
 */

/** Phân loại mức độ nghiêm trọng của lỗi thể thức */
export enum LintSeverity {
  ERROR = "error",
  WARNING = "warning",
  INFO = "info",
}

/** Danh mục mã lỗi thể thức văn bản thường gặp */
export enum LintIssueType {
  /** Quên số 0 đằng trước số văn bản nhỏ hơn 10 (ví dụ: Số 5 thay vì Số 05) */
  MISSING_LEADING_ZERO_NUMBER = "MISSING_LEADING_ZERO_NUMBER",

  /** Quên số 0 ở ngày nhỏ hơn 10 hoặc tháng 1, tháng 2 */
  MISSING_LEADING_ZERO_DATE = "MISSING_LEADING_ZERO_DATE",

  /** Dùng phím Space hoặc Tab ở đầu đoạn thay vì dùng thụt lề chuẩn (10mm) */
  SPACE_OR_TAB_INDENTATION = "SPACE_OR_TAB_INDENTATION",

  /** Phông chữ khác Times New Roman */
  NON_STANDARD_FONT = "NON_STANDARD_FONT",

  /** Sai kích thước lề giấy (Top/Bottom 20mm, Left 30mm, Right 15mm) */
  INVALID_MARGIN = "INVALID_MARGIN",

  /** Đoạn văn bản không được căn đều 2 bên (Justified) */
  INVALID_ALIGNMENT = "INVALID_ALIGNMENT",

  /** Thụt lề dòng đầu khác 10 mm (~28.35 pt) */
  INVALID_FIRST_LINE_INDENT = "INVALID_FIRST_LINE_INDENT",

  /** Giãn dòng nằm ngoài khoảng quy định 18 pt - 22 pt */
  INVALID_LINE_SPACING = "INVALID_LINE_SPACING",

  /** Giãn đoạn sai quy chuẩn (Space After < 6 pt hoặc Space Before > 0 pt) */
  INVALID_PARAGRAPH_SPACING = "INVALID_PARAGRAPH_SPACING",
}

/** Cấu trúc dữ liệu đầu vào cho Ô 1 và Ô 2 (Header Table) */
export interface HeaderData {
  /** Cơ quan cấp trên (nếu có), e.g. "BAN CHẤP HÀNH TRUNG ƯƠNG" hoặc "TỈNH ỦY NGHỆ AN" */
  superiorOrgan?: string;

  /** Cơ quan, đơn vị ban hành văn bản, e.g. "VĂN PHÒNG" hoặc "BAN TỔ CHỨC" */
  issuingOrgan: string;

  /** Số văn bản (hỗ trợ số nguyên hoặc chuỗi số). Tự động bổ sung số 0 nếu < 10 */
  docNumber: number | string;

  /** Ký hiệu văn bản, e.g. "HD/VPTW", "QĐ/TU", "CT/TW" */
  docCode: string;

  /** Địa danh ban hành văn bản, e.g. "Hà Nội", "Nghệ An", "TP. Hồ Chí Minh" */
  location: string;

  /** Ngày tháng năm ban hành văn bản (chấp nhận Date, object hoặc chuỗi text) */
  date: Date | { day: number; month: number; year: number } | string;

  /** Cỡ chữ cho bảng Header (13 pt hoặc 14 pt, mặc định 13 pt) */
  fontSize?: 13 | 14;
}

/** Cấu trúc dữ liệu đầu vào cho Ô Nơi nhận & Ô Chữ ký (Signature Table) */
export interface SignData {
  /** Danh sách thành phần nhận văn bản (các gạch đầu dòng) */
  recipients: string[];

  /**
   * Thẩm quyền/chức vụ người ký dòng 1:
   * Ví dụ: "T/M BAN BÍ THƯ", "T/M BAN THƯỜNG VỤ", "CHÁNH VĂN PHÒNG"
   */
  signTitle: string;

  /**
   * Chức danh cụ thể dòng 2 (nếu dòng 1 là thay mặt tập thể):
   * Ví dụ: "BÍ THƯ", "PHÓ BÍ THƯ", "PHÓ CHÁNH VĂN PHÒNG"
   */
  signerRole?: string;

  /** Họ và tên đầy đủ của người ký */
  signerName: string;

  /** Khoảng trống để ký tên và đóng dấu (tính bằng mm, từ 35-40 mm, mặc định 38 mm) */
  signatureSpaceMm?: number;

  /** Cỡ chữ cho bảng chân trang (13 pt hoặc 14 pt, mặc định 13 pt cho chữ ký) */
  fontSize?: 13 | 14;
}

/** Cấu hình tùy chọn khi định dạng nội dung văn bản */
export interface FormatBodyOptions {
  /** Cỡ chữ nội dung (14 pt hoặc 13 pt, mặc định 14 pt) */
  fontSize?: number;

  /** Giãn dòng cố định (Exactly, mặc định 18 pt) */
  lineSpacing?: number;

  /** Khoảng cách dưới đoạn (Space After, mặc định 6 pt) */
  spaceAfter?: number;

  /** Khoảng cách trên đoạn (Space Before, mặc định 0 pt) */
  spaceBefore?: number;

  /** Thụt lề dòng đầu (mm, mặc định 10 mm) */
  firstLineIndentMm?: number;

  /** Bỏ qua định dạng thụt lề đối với các tiêu đề căn giữa (như QUYẾT ĐỊNH, KẾ HOẠCH) */
  preserveCenteredHeadings?: boolean;
}

/** Cấu hình tùy chọn khi thiết lập trang */
export interface PageSetupOptions {
  /** Chiều rộng lề trên (mm, mặc định 20) */
  topMarginMm?: number;

  /** Chiều rộng lề dưới (mm, mặc định 20) */
  bottomMarginMm?: number;

  /** Chiều rộng lề trái (mm, mặc định 30) */
  leftMarginMm?: number;

  /** Chiều rộng lề phải (mm, mặc định 15) */
  rightMarginMm?: number;

  /** Bật in 2 mặt đối xứng (mặc định true) */
  mirrorMargins?: boolean;

  /** Không hiển thị số trang và header ở trang đầu tiên (mặc định true) */
  differentFirstPageHeaderFooter?: boolean;

  /** Khoảng cách header so với mép trên (mm, mặc định 10) */
  headerDistanceMm?: number;
}

/** Chi tiết của một lỗi thể thức phát hiện được */
export interface LintIssue {
  /** Mã định danh duy nhất cho vấn đề */
  id: string;

  /** Phân loại lỗi */
  type: LintIssueType;

  /** Mức độ nghiêm trọng */
  severity: LintSeverity;

  /** Thông báo lỗi chi tiết */
  message: string;

  /** Hướng dẫn sửa đúng theo 05-HD/VPTW */
  suggestion: string;

  /** Vị trí phát hiện (chỉ số đoạn paragraph hoặc thông tin section) */
  locationInfo?: string;

  /** Trích đoạn văn bản có lỗi */
  snippet?: string;
}

/** Báo cáo kết quả quét thể thức văn bản */
export interface LintReport {
  /** Tài liệu có hợp lệ hoàn toàn không (không có lỗi ERROR) */
  isValid: boolean;

  /** Tổng số vấn đề phát hiện */
  totalIssues: number;

  /** Số lỗi nghiêm trọng */
  errorCount: number;

  /** Số cảnh báo */
  warningCount: number;

  /** Số thông tin nhắc nhở */
  infoCount: number;

  /** Danh sách chi tiết các lỗi */
  issues: LintIssue[];

  /** Thời điểm thực hiện quét */
  timestamp: string;
}
