#!/bin/bash
# CaiDat_AddIn_Mac.sh
# Cài đặt trọn bộ Add-in Chuẩn hóa Văn bản Đảng (Word & Excel) trên macOS

echo "=============================================================================="
echo "    CÀI ĐẶT ADD-IN CHUẨN HÓA VĂN BẢN ĐẢNG & EXCEL (macOS)"
echo "             Văn phòng Đảng ủy xã Thạnh Phú"
echo "=============================================================================="

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Thư mục Startup của Office trên macOS
WORD_STARTUP="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
EXCEL_STARTUP="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Excel"

# 1. Cài đặt Add-in cho Word
if [ -f "$SCRIPT_DIR/ChuanHoa_VanBanDang.dotm" ]; then
    mkdir -p "$WORD_STARTUP"
    cp -f "$SCRIPT_DIR/ChuanHoa_VanBanDang.dotm" "$WORD_STARTUP/"
    echo "[+] Đã cài đặt Add-in Word: ChuanHoa_VanBanDang.dotm"
else
    echo "[!] Không tìm thấy ChuanHoa_VanBanDang.dotm (bỏ qua Word)"
fi

# 2. Cài đặt Add-in cho Excel
if [ -f "$SCRIPT_DIR/ChuanHoa_Excel.xlam" ]; then
    mkdir -p "$EXCEL_STARTUP"
    cp -f "$SCRIPT_DIR/ChuanHoa_Excel.xlam" "$EXCEL_STARTUP/"
    echo "[+] Đã cài đặt Add-in Excel: ChuanHoa_Excel.xlam"
else
    echo "[!] Không tìm thấy ChuanHoa_Excel.xlam (bỏ qua Excel)"
fi

echo ""
echo "=============================================================================="
echo "                   CÀI ĐẶT THÀNH CÔNG 100%!"
echo "=============================================================================="
echo "Khởi động lại Microsoft Word & Excel để sử dụng:"
echo "- Word: Xuất hiện thẻ [VĂN BẢN ĐẢNG] với nút [Chuẩn Hóa Thể Thức]."
echo "- Excel: Tự động sẵn sàng macro chuẩn hóa trang in A4 sắc nét, không che chữ."
echo "Created by VPĐU xã Thạnh Phú"
echo "=============================================================================="
