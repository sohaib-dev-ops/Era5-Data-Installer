@echo off
setlocal
cd /d "%~dp0"
set PYTHON_EXE=C:\Users\HP\.python311\python.exe

echo ===============================================================
echo     ERA5 Reanalysis 2m Temperature Download & Excel Pipeline
echo ===============================================================
echo.

if not exist "%USERPROFILE%\.cdsapirc" (
    echo [NOTICE] CDS API key not detected.
    echo Launching CDS credentials setup...
    "%PYTHON_EXE%" setup_cds_key.py
    echo.
)

echo Starting ERA5 Download & Conversion (Year-by-Year)...
echo Target area: Lat [27.5, 34.0], Lon [69.25, 75.5]
echo Output directory: excel_output\
echo.

"%PYTHON_EXE%" era5_yearly_downloader.py --start_year 1940 --end_year 2025 --export_csv

pause
