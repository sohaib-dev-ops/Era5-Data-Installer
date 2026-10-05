# ERA5 2m-Temperature Download & Excel Conversion

This tool downloads hourly 2m-temperature data from the Copernicus Climate Data Store (ERA5 reanalysis) and converts it into structured Excel (`.xlsx`) and CSV files.

---

## 1. Quick Setup (Copernicus CDS API Key)

To download ERA5 data, you need free access to the Copernicus Climate Data Store:

1. **Register / Log in**: [Copernicus Climate Data Store](https://cds.climate.copernicus.eu)
2. **Accept Dataset Terms**: Navigate to the [ERA5 Single Levels Dataset](https://cds.climate.copernicus.eu/datasets/reanalysis-era5-single-levels), scroll down to the bottom, and click **Accept Terms**.
3. **Copy your Personal Access Token**: Go to your [CDS User Profile](https://cds.climate.copernicus.eu/profile) and copy the Personal Access Token.
4. **Configure your key**: Run:
   ```cmd
   python setup_cds_key.py
   ```
   Paste your token when prompted. This automatically writes your `~/.cdsapirc` configuration file.

---

## 2. Running the Pipeline

You can run the entire pipeline with:

```cmd
run_download.bat
```

Or via Python directly:

```cmd
# Process all years from 1940 to 2025:
python era5_yearly_downloader.py --start_year 1940 --end_year 2025

# Or test with a smaller range (e.g., 2023 to 2024):
python era5_yearly_downloader.py --start_year 2023 --end_year 2024
```

### Resume Support
If the process is stopped or interrupted, you can re-run the command anytime. It will automatically detect already-converted files in `excel_output/` and skip them.

---

## 3. Structure of Generated Excel Files (`.xlsx`)

For each year, a file named `ERA5_2m_Temperature_{YEAR}.xlsx` is produced inside `excel_output/`:

1. **`Daily_Summary` Sheet**:
   - `Date`
   - `Regional_Mean_C` (Daily average temperature across the area in Celsius)
   - `Regional_Min_C` (Daily minimum temperature)
   - `Regional_Max_C` (Daily maximum temperature)

2. **`Grid_Coordinates` Sheet**:
   - Master reference table mapping each grid column (e.g., `G001_34.00N_69.25E`) to its exact `Latitude` and `Longitude`.

3. **`Hourly_Grid_Matrix` Sheet**:
   - **Rows**: 8,760 hourly timestamps for the year (`Timestamp`).
   - **Columns**: 702 columns representing all 0.25° grid points covering `[34°N, 69.25°E, 27.5°N, 75.5°E]`.
   - Temperature values are converted from Kelvin to **Celsius** (°C).
   - This matrix format fits well within Excel limits (Excel max: 1,048,576 rows × 16,384 columns).

---

## 4. Multi-Year Compilation Pipeline (`compile_era5_data.py`)

To compile the individual yearly Excel files into decade sets (10-year chunks) and a master unified dataset:

```cmd
# Compile all 10-year decade workbooks AND the full 1940-2026 dataset:
python compile_era5_data.py

# Or compile only the 10-year decade workbooks:
python compile_era5_data.py --decades_only

# Or compile only the master full dataset:
python compile_era5_data.py --full_only
```

### Generated Datasets:
1. **Decade Workbooks (10-Year Sets)**:
   - `ERA5_Daily_Compiled_1940_1949.xlsx`
   - `ERA5_Daily_Compiled_1950_1959.xlsx`
   - `ERA5_Daily_Compiled_1960_1969.xlsx`
   - `ERA5_Daily_Compiled_1970_1979.xlsx`
   - `ERA5_Daily_Compiled_1980_1989.xlsx`
   - `ERA5_Daily_Compiled_1990_1999.xlsx`
   - `ERA5_Daily_Compiled_2000_2009.xlsx`
   - `ERA5_Daily_Compiled_2010_2019.xlsx`
   - `ERA5_Daily_Compiled_2020_2026.xlsx`
   * **Sheets per Decade**: `Grid_Coordinates`, `Regional_Daily_Summary`, `{START}_{END}_Continuous`, and one tab per individual year.

2. **Master Full Dataset Workbook**:
   - `ERA5_Daily_Compiled_1940_2026_Full.xlsx` (107.1 MB)
   * **Sheets**:
     - `Grid_Coordinates`: Master lookup for all 702 grid points.
     - `Regional_Daily_Summary`: 31,679 rows of daily regional statistics (1940–2026).
     - `1940_2026_Continuous`: Complete continuous daily matrix (31,679 rows × 705 columns).

---

## 5. Meteorological Stations & City Transformation Pipeline (`transform_era5_stations.py`)

Transforms raw ERA5 702-point grid matrices into named meteorological stations and full-city entities matching Pakistan Meteorological Department (PMD) coordinates:

```cmd
# Run transformation for all years (1940-2026):
python transform_era5_stations.py

# Or generate decade workbooks in addition to the master workbook:
python transform_era5_stations.py --decades
```

### Generated Master Dataset:
- **`ERA5_Stations_Transformed_1940_2026.xlsx`** (153.9 MB)
  * **Sheet 1: `Hourly_Stations`**: 760,282 continuous hourly rows from 1940 to 2026 across 33 meteorological series (24 full cities + 9 sub-stations).
  * **Sheet 2: `Daily_Summary`**: 31,679 daily rows with Mean, Min, and Max temperatures for every city and station.
  * **Sheet 3: `Station_Directory`**: Complete PMD station catalog with Index Nos, coordinates (DMS and Decimal), elevation, phone numbers, and ERA5 4-point bilinear weights.

