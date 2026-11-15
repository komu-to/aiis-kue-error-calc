@echo off
chcp 65001 > nul
echo ========================================================
echo Сборка автономного исполняемого файла .EXE для Windows 10
echo ========================================================
echo.
cd /d "%~dp0"
pyinstaller --noconfirm --onedir --windowed --name "AIIS_KUE_Error_Calc" --add-data "core.py;." --add-data "exporters.py;." --add-data "gui.py;." app.py
echo.
if exist "dist\AIIS_KUE_Error_Calc\AIIS_KUE_Error_Calc.exe" (
    echo [УСПЕХ] Файл собран в: dist\AIIS_KUE_Error_Calc\AIIS_KUE_Error_Calc.exe
) else (
    echo [ОШИБКА] Не удалось собрать исполняемый файл.
)
pause
