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

:: 0. Dong Word va Excel neu dang mo de tranh bi khoa file
echo [*] Dang kiem tra va dong tam thoi Word / Excel...
taskkill /F /IM WINWORD.EXE >nul 2>&1
taskkill /F /IM EXCEL.EXE >nul 2>&1

:: 1. Go khoa bao mat Internet (Unblock Mark-of-the-Web) tren file nguon
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -Path '%SCRIPT_DIR%' -Include *.dotm,*.xlam -Recurse | Unblock-File" >nul 2>&1

set "WORD_OK=0"
set "EXCEL_OK=0"

:: 2. Cai dat Word Add-in
if exist "%WORD_SOURCE%" (
    if not exist "%WORD_TARGET%" mkdir "%WORD_TARGET%" 2>nul
    copy /Y "%WORD_SOURCE%" "%WORD_TARGET%\ChuanHoa_VanBanDang.dotm" >nul
    if exist "%WORD_TARGET%\ChuanHoa_VanBanDang.dotm" (
        powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -Path '%WORD_TARGET%\ChuanHoa_VanBanDang.dotm'" >nul 2>&1
        echo [+] Da cai dat xong Add-in Word: ChuanHoa_VanBanDang.dotm
        set "WORD_OK=1"
    )
) else (
    echo [!] CANH BAO: Khong tim thay file ChuanHoa_VanBanDang.dotm trong thu muc giai nen!
)

:: 3. Cai dat Excel Add-in
if exist "%EXCEL_SOURCE%" (
    if not exist "%EXCEL_TARGET%" mkdir "%EXCEL_TARGET%" 2>nul
    copy /Y "%EXCEL_SOURCE%" "%EXCEL_TARGET%\ChuanHoa_Excel.xlam" >nul
    if exist "%EXCEL_TARGET%\ChuanHoa_Excel.xlam" (
        powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -Path '%EXCEL_TARGET%\ChuanHoa_Excel.xlam'" >nul 2>&1
        echo [+] Da cai dat xong Add-in Excel: ChuanHoa_Excel.xlam
        set "EXCEL_OK=1"
    )
) else (
    echo [!] CANH BAO: Khong tim thay file ChuanHoa_Excel.xlam trong thu muc giai nen!
)

:: 4. Dang ky thu muc STARTUP va XLSTART vao Trusted Locations cua Office (15.0 & 16.0)
echo [*] Dang cau hinh quyen tin cay (Trusted Locations) cho Office...
for %%V in (15.0 16.0) do (
    reg add "HKCU\Software\Microsoft\Office\%%V\Word\Security\Trusted Locations\VPDU_WordStartup" /v "Path" /t REG_SZ /d "%WORD_TARGET%\" /f >nul 2>&1
    reg add "HKCU\Software\Microsoft\Office\%%V\Word\Security\Trusted Locations\VPDU_WordStartup" /v "AllowSubfolders" /t REG_DWORD /d 1 /f >nul 2>&1
    reg add "HKCU\Software\Microsoft\Office\%%V\Word\Security" /v "DisableAllAddins" /t REG_DWORD /d 0 /f >nul 2>&1

    reg add "HKCU\Software\Microsoft\Office\%%V\Excel\Security\Trusted Locations\VPDU_ExcelStartup" /v "Path" /t REG_SZ /d "%EXCEL_TARGET%\" /f >nul 2>&1
    reg add "HKCU\Software\Microsoft\Office\%%V\Excel\Security\Trusted Locations\VPDU_ExcelStartup" /v "AllowSubfolders" /t REG_DWORD /d 1 /f >nul 2>&1
    reg add "HKCU\Software\Microsoft\Office\%%V\Excel\Security" /v "DisableAllAddins" /t REG_DWORD /d 0 /f >nul 2>&1
)

echo.
if "%WORD_OK%"=="1" if "%EXCEL_OK%"=="1" (
    echo ======================================================
    echo        CAI DAT THANH CONG TRON BO WORD ^& EXCEL!
    echo ======================================================
    echo.
    echo Hay mo lai Microsoft Word va Excel:
    echo - Tren Word: Xuat hien the [VAN BAN DANG] tren thanh Ribbon.
    echo - Tren Excel: Xuat hien the [CHUAN HOA IN] tren thanh Ribbon.
    echo ======================================================
    pause
    exit /b 0
)

echo ======================================================
echo [!] CHU Y: Cai dat chua du bo (Word=%WORD_OK%, Excel=%EXCEL_OK%).
echo Vui long Giai Nen ^(Extract Here^) toan bo file .zip ra thu muc
echo truoc khi nhap dup vao file CaiDat_AddIn_Windows.bat!
echo ======================================================
pause
