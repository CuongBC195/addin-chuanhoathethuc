/**
 * unit.test.ts
 * Kiểm thử đơn vị các hàm logic định dạng theo Hướng dẫn 05-HD/VPTW
 */

import { mmToPt, ptToMm, MM_TO_PT, PAGE_SETUP_CONSTANTS } from "../src/constants";
import { formatDocNumber, formatPartyDate } from "../src/formatService";

function assert(condition: boolean, message: string) {
  if (!condition) {
    console.error(`❌ FAIL: ${message}`);
    throw new Error(`Test failed: ${message}`);
  }
  console.log(`✅ PASS: ${message}`);
}

console.log("=== BẮT ĐẦU KIỂM THỬ ĐƠN VỊ THỂ THỨC 05-HD/VPTW ===");

// 1. Kiểm thử chuyển đổi đơn vị mm sang pt
assert(Math.abs(mmToPt(10) - 28.3465) < 0.001, "10 mm ≈ 28.35 pt");
assert(Math.abs(mmToPt(20) - 56.693) < 0.001, "20 mm ≈ 56.69 pt");
assert(Math.abs(mmToPt(30) - 85.0395) < 0.001, "30 mm ≈ 85.04 pt");
assert(Math.abs(mmToPt(15) - 42.5198) < 0.001, "15 mm ≈ 42.52 pt");
assert(Math.abs(ptToMm(PAGE_SETUP_CONSTANTS.MARGIN_LEFT_PT) - 30) < 0.001, "ptToMm(MARGIN_LEFT_PT) == 30 mm");

// 2. Kiểm thử quy chuẩn số văn bản < 10 phải có số 0 đằng trước
assert(formatDocNumber(5) === "05", "Số 5 -> '05'");
assert(formatDocNumber(1) === "01", "Số 1 -> '01'");
assert(formatDocNumber(9) === "09", "Số 9 -> '09'");
assert(formatDocNumber(10) === "10", "Số 10 -> '10'");
assert(formatDocNumber(125) === "125", "Số 125 -> '125'");
assert(formatDocNumber("5") === "05", "Chuỗi '5' -> '05'");
assert(formatDocNumber("05") === "05", "Chuỗi '05' giữ nguyên '05'");
assert(formatDocNumber("12") === "12", "Chuỗi '12' giữ nguyên '12'");

// 3. Kiểm thử quy chuẩn ngày tháng năm
// Ngày < 10 có số 0, tháng < 3 có số 0
const date1 = formatPartyDate("Hà Nội", { day: 5, month: 2, year: 2026 });
assert(date1 === "Hà Nội, ngày 05 tháng 02 năm 2026", `Ngày 5 tháng 2 -> '${date1}'`);

const date2 = formatPartyDate("Nghệ An", { day: 27, month: 5, year: 2026 });
assert(date2 === "Nghệ An, ngày 27 tháng 5 năm 2026", `Ngày 27 tháng 5 -> '${date2}'`);

const date3 = formatPartyDate("TP. Hồ Chí Minh", { day: 1, month: 1, year: 2026 });
assert(date3 === "TP. Hồ Chí Minh, ngày 01 tháng 01 năm 2026", `Ngày 1 tháng 1 -> '${date3}'`);

const date4 = formatPartyDate("Đà Nẵng", { day: 15, month: 12, year: 2026 });
assert(date4 === "Đà Nẵng, ngày 15 tháng 12 năm 2026", `Ngày 15 tháng 12 -> '${date4}'`);

// Chuỗi nhập thô thiếu số 0 được auto-fix
const rawDateFix1 = formatPartyDate("Hà Nội", "ngày 5 tháng 2 năm 2026");
assert(rawDateFix1 === "Hà Nội, ngày 05 tháng 02 năm 2026", `Fix chuỗi thô -> '${rawDateFix1}'`);

console.log("=== TẤT CẢ KIỂM THỬ ĐƠN VỊ ĐÃ THÀNH CÔNG ===");
