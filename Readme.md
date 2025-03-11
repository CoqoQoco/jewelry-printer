<!-- install package -->
pip install pyinstaller


<!-- create .exe -->
python -m PyInstaller --onefile --windowed --icon=printer.ico --name="Zebra Print Service" print_service.py 