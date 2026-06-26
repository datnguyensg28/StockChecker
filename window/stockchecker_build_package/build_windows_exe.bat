@echo off
setlocal
cd /d %~dp0

echo =====================================
echo  STOCKCHECKER - BUILD WINDOWS EXE
echo =====================================

python --version
if errorlevel 1 (
  echo Python chua duoc cai hoac chua them vao PATH.
  pause
  exit /b 1
)

echo.
echo [1/4] Tao moi truong ao .venv...
python -m venv .venv
call .venv\Scripts\activate.bat

echo.
echo [2/4] Nang cap pip va cai thu vien...
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [3/4] Build file EXE...
pyinstaller Stockchecker.spec --clean --noconfirm

echo.
echo [4/4] Hoan tat.
echo File EXE nam tai: dist\Stockchecker\Stockchecker.exe
echo.
echo Luu y: neu app can file data\MB52.xlsx, hay dat thu muc data nam cung cap voi file exe.
pause
