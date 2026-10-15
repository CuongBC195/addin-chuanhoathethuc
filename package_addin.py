#!/usr/bin/env python3
"""
package_addin.py
Tự động đóng gói ChuanHoa_VanBanDang.dotm và ChuanHoa_Excel.xlam
Nhúng Custom Ribbon UI chuẩn cho cả Word và Excel, nén bộ cài Bo_Cai_Dat_AddIn_VanBanDang.zip
"""
import os
import sys
import zipfile
import shutil
import xml.etree.ElementTree as ET

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

NORMAL_DOTM_MAC = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/Normal.dotm"
)
WORD_STARTUP_MAC = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
)
EXCEL_STARTUP_MAC = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Excel"
)

OUTPUT_DOTM = os.path.join(BASE_DIR, "ChuanHoa_VanBanDang.dotm")
EXCEL_XLAM = os.path.join(BASE_DIR, "ChuanHoa_Excel.xlam")

CUSTOM_UI_WORD_XML = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab idMso="TabHome">
        <group id="grpChuanHoaHome" label="VĂN BẢN ĐẢNG" insertAfterMso="GroupClipboard">
          <button id="btnChuanHoaHome"
                  label="Chuẩn Hóa Thể Thức"
                  size="large"
                  imageMso="AcceptTask"
                  screentip="Chuẩn hóa văn bản Đảng (HD 05 &amp; HD 36)"
                  supertip="Tự động chuẩn hóa toàn bộ font chữ Times New Roman 14, tiêu đề, tiêu ngữ, đường kẻ, công văn (V/v size 12, kính gửi), giãn đoạn 6pt, thụt dòng 1.27cm, nơi nhận và chữ ký."
                  onAction="RibbonChuanHoa" />
        </group>
      </tab>
      <tab id="tabVanBanDang" label="VĂN BẢN ĐẢNG">
        <group id="grpChuanHoa" label="Thể Thức Văn Bản">
          <button id="btnChuanHoa"
                  label="Chuẩn Hóa Thể Thức"
                  size="large"
                  imageMso="AcceptTask"
                  screentip="Chuẩn hóa văn bản Đảng (HD 05 &amp; HD 36)"
                  supertip="Tự động chuẩn hóa toàn bộ font chữ Times New Roman 14, tiêu đề, tiêu ngữ, đường kẻ, công văn (V/v size 12, kính gửi), giãn đoạn 6pt, thụt dòng 1.27cm, nơi nhận và chữ ký."
                  onAction="RibbonChuanHoa" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""

CUSTOM_UI_EXCEL_XML = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab idMso="TabHome">
        <group id="grpChuanHoaExcelHome" label="IN ẤN A4">
          <button id="btnChuanHoaExcelHome"
                  label="Chuẩn Hóa Trang In"
                  size="large"
                  imageMso="PrintAreaMenu"
                  screentip="Chuẩn hóa trang in A4 và căn chỉnh bảng biểu"
                  supertip="Tự động tính chiều cao hàng chống che chữ (kể cả Merged cells), xoay A4 Ngang, vừa khít 1 trang ngang."
                  onAction="RibbonChuanHoaExcel" />
        </group>
      </tab>
      <tab id="tabChuanHoaExcel" label="CHUẨN HÓA IN">
        <group id="grpChuanHoaExcel" label="Bảng Biểu &amp; Trang In">
          <button id="btnChuanHoaExcelTab"
                  label="Chuẩn Hóa Trang In"
                  size="large"
                  imageMso="PrintAreaMenu"
                  screentip="Chuẩn hóa trang in A4 và căn chỉnh bảng biểu"
                  supertip="Tự động tính chiều cao hàng chống che chữ (kể cả Merged cells), xoay A4 Ngang, vừa khít 1 trang ngang."
                  onAction="RibbonChuanHoaExcel" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""

def inject_ribbon(file_path, custom_ui_content):
    """Nhúng customUI14.xml vào file OpenXML (.dotm hoặc .xlam)"""
    tmp_file = file_path + ".tmp"
    with zipfile.ZipFile(file_path, "r") as zin:
        rels_xml = zin.read("_rels/.rels").decode("utf-8")
        ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/relationships")
        root = ET.fromstring(rels_xml)
        
        ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        has_rel = False
        for rel in root.findall(f"{ns}Relationship"):
            if "customui" in rel.attrib.get("Target", "").lower():
                has_rel = True
                break
                
        if not has_rel:
            rel_elem = ET.Element(f"{ns}Relationship", {
                "Id": "rIdCustomUI14",
                "Type": "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
                "Target": "customUI/customUI14.xml"
            })
            root.append(rel_elem)
            updated_rels = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        else:
            updated_rels = rels_xml.encode("utf-8")

        with zipfile.ZipFile(tmp_file, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "_rels/.rels":
                    zout.writestr(item, updated_rels)
                elif item.filename == "customUI/customUI14.xml":
                    continue
                else:
                    zout.writestr(item, zin.read(item.filename))
            zout.writestr("customUI/customUI14.xml", custom_ui_content.encode("utf-8"))

    os.replace(tmp_file, file_path)


def package_word():
    # Ưu tiên lấy từ Normal.dotm nếu có, hoặc dùng OUTPUT_DOTM có sẵn
    src_file = None
    if os.path.exists(NORMAL_DOTM_MAC):
        with zipfile.ZipFile(NORMAL_DOTM_MAC, "r") as z:
            if "word/vbaProject.bin" in z.namelist():
                src_file = NORMAL_DOTM_MAC

    if src_file:
        print(f"[*] Đọc template gốc Word từ: {src_file}")
        shutil.copyfile(src_file, OUTPUT_DOTM)
    elif not os.path.exists(OUTPUT_DOTM):
        print(f"LỖI: Không tìm thấy nguồn Word Add-in!")
        return False

    inject_ribbon(OUTPUT_DOTM, CUSTOM_UI_WORD_XML)
    print(f"[+] Đã nhúng Ribbon vào Word Add-in: {OUTPUT_DOTM}")
    return True


def package_excel():
    if not os.path.exists(EXCEL_XLAM):
        print(f"LỖI: Không tìm thấy {EXCEL_XLAM}")
        return False

    inject_ribbon(EXCEL_XLAM, CUSTOM_UI_EXCEL_XML)
    print(f"[+] Đã nhúng Ribbon vào Excel Add-in: {EXCEL_XLAM}")
    return True


def install_local_mac():
    """Tự động copy Add-in vào thư mục Startup Office trên máy Mac hiện tại"""
    if os.path.exists(OUTPUT_DOTM):
        os.makedirs(WORD_STARTUP_MAC, exist_ok=True)
        shutil.copyfile(OUTPUT_DOTM, os.path.join(WORD_STARTUP_MAC, "ChuanHoa_VanBanDang.dotm"))
        print(f"[+] Đã cài vào Word Startup Mac: {WORD_STARTUP_MAC}")

    if os.path.exists(EXCEL_XLAM):
        os.makedirs(EXCEL_STARTUP_MAC, exist_ok=True)
        shutil.copyfile(EXCEL_XLAM, os.path.join(EXCEL_STARTUP_MAC, "ChuanHoa_Excel.xlam"))
        print(f"[+] Đã cài vào Excel Startup Mac: {EXCEL_STARTUP_MAC}")


def make_zip():
    zip_path = os.path.join(BASE_DIR, "Bo_Cai_Dat_AddIn_VanBanDang.zip")
    files_to_zip = [
        "ChuanHoa_VanBanDang.dotm",
        "ChuanHoa_Excel.xlam",
        "CaiDat_AddIn_Windows.bat",
        "GoCaiDat_Windows.bat",
        "CaiDat_AddIn_Mac.sh",
        "HuongDan_DongNghiep.txt"
    ]
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in files_to_zip:
            fpath = os.path.join(BASE_DIR, fname)
            if os.path.exists(fpath):
                zf.write(fpath, arcname=fname)
    print(f"[+] Đã đóng gói bộ cài ZIP gửi đồng nghiệp: {zip_path}")


def main():
    package_word()
    package_excel()
    install_local_mac()
    make_zip()
    print("=====================================================")
    print(" HOÀN TẤT: CẢ WORD VÀ EXCEL ĐỀU ĐÃ CÓ NÚT TRÊN RIBBON!")
    print("=====================================================")

if __name__ == "__main__":
    main()
