@echo off
REM Thiết lập tiêu đề cho cửa sổ Command Prompt
title Git Auto Uploader - Interactive

echo =========================================================
echo BAT File: Tu dong day code len GitHub
echo =========================================================

REM Hỏi người dùng nhập thông báo commit
set /P commit_message="Nhap thong bao COMMIT (vi du: Cap nhat giao dien): "
echo.

REM Kiem tra neu nguoi dung khong nhap gi
if "%commit_message%"=="" (
    echo LOI: Thong bao commit khong duoc de trong.
    pause
    goto :eof
)

REM Thuc hien cac lenh Git
echo 1. Kiem tra va them tat ca file moi/thay doi...
git add .
echo.

echo 2. Tao COMMIT voi thong bao: "%commit_message%"
git commit -m "%commit_message%"
echo.

echo 3. Day code len GitHub (nhanh main)...
git push -u origin main
echo.

echo =========================================================
echo ✅ Quy trinh day code da HOAN TAT!
echo =========================================================
pause