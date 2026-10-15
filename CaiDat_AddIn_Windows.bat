@echo off
chcp 65001 >nul
cls
echo ======================================================
echo    DANG CAI DAT ADD-IN CHUAN HOA VAN BAN DANG ^& EXCEL
echo              Van phong Dang uy xa Thanh Phu
echo ======================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "WORD_SOURCE=%SCRIPT_DIR%ChuanHoa_VanBanDang.dotm"
set "EXCEL_SOURCE=%SCRIPT_DIR%ChuanHoa_Excel.xlam"

set "WORD_TARGET=%APPDATA%\Microsoft\Word\STARTUP"
set "EXCEL_TARGET=%APPDATA%\Microsoft\Excel\XLSTART"

set "OK_COUNT=0"

:: 1. Cai dat Word Add-in
if exist "%WORD_SOURCE%" (
    if not exist "%WORD_TARGET%" mkdir "%WORD_TARGET%" 2>nul
    copy /Y "%WORD_SOURCE%" "%WORD_TARGET%\ChuanHoa_VanBanDang.dotm" >nul
    if exist "%WORD_TARGET%\ChuanHoa_VanBanDang.dotm" (
        echo [+] Da cai dat xong Add-in Word: ChuanHoa_VanBanDang.dotm
        set /a OK_COUNT+=1
    )
) else (
    echo [!] Khong tim thay ChuanHoa_VanBanDang.dotm
)

:: 2. Cai dat Excel Add-in
if exist "%EXCEL_SOURCE%" (
    if not exist "%EXCEL_TARGET%" mkdir "%EXCEL_TARGET%" 2>nul
    copy /Y "%EXCEL_SOURCE%" "%EXCEL_TARGET%\ChuanHoa_Excel.xlam" >nul
    if exist "%EXCEL_TARGET%\ChuanHoa_Excel.xlam" (
        echo [+] Da cai dat xong Add-in Excel: ChuanHoa_Excel.xlam
        set /a OK_COUNT+=1
    )
) else (
    echo [!] Khong tim thay ChuanHoa_Excel.xlam
)

echo.
if %OK_COUNT% GEQ 1 (
    echo ======================================================
    echo             CAI DAT THANH CONG 100%%!
    echo ======================================================
    echo.
    echo Tu bay gio:
    echo - Mo Word: Tu dong co the [VAN BAN DANG] tren Ribbon.
    echo - Mo Excel: Tu dong co san tien ich Chuan Hoa Trang In A4.
    echo.
    echo (Created by VPDU xa Thanh Phu)
    echo ======================================================
) else (
    echo [LOI] Khong the sao chep file vao thu muc Startup!
    echo Vui long giai nen toan bo file zip truoc khi chay file .bat nay.
)
echo.
pause
