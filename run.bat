@echo off
chcp 65001 > nul
title Расчет погрешностей ИК АИИС КУЭ
echo Запуск приложения...
start "" pythonw "%~dp0app.py"
if %errorlevel% neq 0 (
    echo Pythonw не сработал, запуск через python...
    python "%~dp0app.py"
)
