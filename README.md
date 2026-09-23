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
