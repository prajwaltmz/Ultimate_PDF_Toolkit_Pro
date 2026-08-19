@echo off
echo Installing requirements...
pip install -r requirements.txt

echo.
echo Compiling Ultimate PDF Toolkit Pro into a standalone .exe...
pyinstaller --onefile --windowed "PDF, Word file merger.py"

echo.
echo Build complete! Your .exe is located in the "dist" folder.
pause
