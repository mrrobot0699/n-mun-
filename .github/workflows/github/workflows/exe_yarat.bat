@echo off
rem Python olan kompüterdə bir dəfə işlədin -> dist\DebitorKreditor.exe yaranır (digər kompüterlərə Python lazım deyil)
pip install openpyxl pyinstaller
pyinstaller --onefile --windowed --name DebitorKreditor debitor_kreditor.py
pause
