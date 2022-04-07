pyinstaller --clean --onefile --windowed ^
    --distpath="dist/" ^
    --icon="icon.ico" ^
    --name="SupplyLineUtility" ^
    --add-data="icon.ico";. ^
    app.py