@echo off
chcp 65001 >nul
cls
echo ======================================================
echo    GO BO ADD-IN CHUAN HOA VAN BAN DANG ^& EXCEL
echo ======================================================
echo.

set "WORD_TARGET=%APPDATA%\Microsoft\Word\STARTUP\ChuanHoa_VanBanDang.dotm"
set "EXCEL_TARGET=%APPDATA%\Microsoft\Excel\XLSTART\ChuanHoa_Excel.xlam"

if exist "%WORD_TARGET%" del /F /Q "%WORD_TARGET%" >nul 2>&1
if exist "%EXCEL_TARGET%" del /F /Q "%EXCEL_TARGET%" >nul 2>&1

echo Da go bo thanh cong khoi Word va Excel!
echo.
pause
