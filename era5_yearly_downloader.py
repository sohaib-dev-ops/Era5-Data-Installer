"""
ERA5 Hourly 2m-Temperature Yearly Downloader and Excel Converter
Dataset: reanalysis-era5-single-levels
Area: [34.0, 69.25, 27.5, 75.5] (North, West, South, East)

Output per year:
- Excel (.xlsx):
  * Sheet 1 'Hourly_Matrix': Timestamps (rows) x 702 Grid Points (columns) [Fits within Excel's limits]
  * Sheet 2 'Grid_Coordinates': Mapping of Grid_ID -> (Latitude, Longitude)
  * Sheet 3 'Daily_Summary': Daily Mean, Min, and Max temperature
- Optional CSV (long format): [Timestamp, Latitude, Longitude, Temp_Celsius]
"""

import os
import sys

# Ensure local directory is in Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import glob
import zipfile
import argparse
import numpy as np
import pandas as pd
import xarray as xr

# Bounding box for area of interest
AREA = [34.0, 69.25, 27.5, 75.5]

def get_cds_client(api_key=None):
    """Verify CDS credentials and return a cdsapi Client."""
    import cdsapi
    if api_key:
        return cdsapi.Client(url="https://cds.climate.copernicus.eu/api", key=api_key)

    cdsapirc_path = os.path.expanduser("~/.cdsapirc")
    has_env = os.environ.get("CDSAPI_KEY")
    if not os.path.exists(cdsapirc_path) and not has_env:
        print("\n" + "=" * 65)
        print("[ERROR] CDS API credentials file (.cdsapirc) not found!")
        print(f"Target location: {cdsapirc_path}")
        print("=" * 65)
        print("To download from Copernicus CDS, you need to configure your key:")
        print("1. Go to: https://cds.climate.copernicus.eu")
        print("2. Log in and accept the license for 'ERA5 hourly data on single levels':")
        print("   https://cds.climate.copernicus.eu/datasets/reanalysis-era5-single-levels")
        print("3. Copy your Personal Access Token from:")
        print("   https://cds.climate.copernicus.eu/profile")
        print("4. Run: python setup_cds_key.py")
        print("   or pass: --api_key <YOUR_TOKEN>")
        print("=" * 65 + "\n")
        sys.exit(1)
    
    return cdsapi.Client()

import tarfile

def extract_nc_from_zip(zip_path, extract_dir):
    """Extract .nc file from CDS zip or tar download, or return None if already raw."""
    if zipfile.is_zipfile(zip_path):
        try:
            with zipfile.ZipFile(zip_path, 'r') as zf:
                zf.extractall(extract_dir)
                nc_files = [os.path.join(extract_dir, f) for f in zf.namelist() if f.endswith(".nc")]
                if nc_files:
                    return nc_files[0]
        except Exception as e:
            print(f"[!] Warning extracting zip: {e}")
    elif tarfile.is_tarfile(zip_path):
        try:
            with tarfile.open(zip_path, 'r:*') as tf:
                tf.extractall(extract_dir)
                nc_files = [os.path.join(extract_dir, f) for f in tf.getnames() if f.endswith(".nc")]
                if nc_files:
                    return nc_files[0]
        except Exception as e:
            print(f"[!] Warning extracting tar: {e}")
    return None

def download_year(client, year, downloads_dir):
    """Download single year NetCDF from Copernicus CDS API."""
    os.makedirs(downloads_dir, exist_ok=True)
    target_nc = os.path.join(downloads_dir, f"era5_t2m_{year}.nc")
    target_zip = os.path.join(downloads_dir, f"era5_t2m_{year}.zip")

    if os.path.exists(target_nc):
        print(f"[*] Raw NetCDF already exists for {year}: {target_nc}")
        return target_nc

    print(f"\n[+] Requesting ERA5 2m_temperature for year {year} from CDS...")
    request = {
        "product_type": ["reanalysis"],
        "variable": ["2m_temperature"],
        "year": [str(year)],
        "month": [f"{m:02d}" for m in range(1, 13)],
        "day": [f"{d:02d}" for d in range(1, 32)],
        "time": [f"{h:02d}:00" for h in range(24)],
        "data_format": "netcdf",
        "download_format": "zip",
        "area": AREA
    }

    client.retrieve("reanalysis-era5-single-levels", request).download(target_zip)

    # Extract NetCDF
    extracted = extract_nc_from_zip(target_zip, downloads_dir)
    if extracted:
        if extracted != target_nc:
            if os.path.exists(target_nc):
                os.remove(target_nc)
            os.rename(extracted, target_nc)
        if os.path.exists(target_zip):
            try:
                os.remove(target_zip)
            except Exception:
                pass
        return target_nc
    elif os.path.exists(target_zip):
        # If the file downloaded was already an uncompressed .nc named .zip
        os.rename(target_zip, target_nc)
        return target_nc
    else:
        raise FileNotFoundError(f"Failed to find downloaded NetCDF for year {year}")

def convert_nc_to_excel_and_csv(nc_path, year, excel_dir, csv_dir=None):
    """
    Converts 1 year of ERA5 NetCDF data into an optimized Excel workbook and optional CSV.
    """
    os.makedirs(excel_dir, exist_ok=True)
    excel_path = os.path.join(excel_dir, f"ERA5_2m_Temperature_{year}.xlsx")

    print(f"[*] Processing {nc_path} -> {excel_path}...")
    with xr.open_dataset(nc_path) as ds:
        var_name = "t2m" if "t2m" in ds else ("2t" if "2t" in ds else list(ds.data_vars)[0])
        da = ds[var_name]

        # Convert Kelvin to Celsius
        da_celsius = da - 273.15

        # Get coordinates
        lats = da_celsius['latitude'].values
        lons = da_celsius['longitude'].values
        times = da_celsius['valid_time'].values if 'valid_time' in da_celsius.coords else da_celsius['time'].values

        # Build Grid Metadata (Grid_ID -> Lat, Lon)
        grid_rows = []
        grid_col_names = []
        gid = 1
        for lat in lats:
            for lon in lons:
                col_name = f"G{gid:03d}_{lat:.2f}N_{lon:.2f}E"
                grid_col_names.append(col_name)
                grid_rows.append({
                    "Grid_ID": f"G{gid:03d}",
                    "Column_Name": col_name,
                    "Latitude": round(float(lat), 3),
                    "Longitude": round(float(lon), 3)
                })
                gid += 1
        
        coords_df = pd.DataFrame(grid_rows)

        # Reshape temperature data: (time, lat, lon) -> (time, num_grid_points)
        num_times = len(times)
        num_grids = len(grid_col_names)
        t_matrix = da_celsius.values.reshape((num_times, num_grids))
        
        # Build DataFrame: Timestamps as rows, Grid points as columns
        hourly_matrix_df = pd.DataFrame(t_matrix, columns=grid_col_names)
        hourly_matrix_df.insert(0, "Timestamp", pd.to_datetime(times))

        # Calculate Daily Statistics (Mean, Min, Max over entire region)
        hourly_matrix_df["Date"] = hourly_matrix_df["Timestamp"].dt.date
        numeric_cols = grid_col_names
        daily_mean = hourly_matrix_df.groupby("Date")[numeric_cols].mean().mean(axis=1)
        daily_min = hourly_matrix_df.groupby("Date")[numeric_cols].min().min(axis=1)
        daily_max = hourly_matrix_df.groupby("Date")[numeric_cols].max().max(axis=1)

        daily_summary_df = pd.DataFrame({
            "Date": daily_mean.index,
            "Regional_Mean_C": np.round(daily_mean.values, 2),
            "Regional_Min_C": np.round(daily_min.values, 2),
            "Regional_Max_C": np.round(daily_max.values, 2)
        })
        hourly_matrix_df.drop(columns=["Date"], inplace=True)

        # Write to Excel with openpyxl
        print(f"[*] Writing Excel workbook ({hourly_matrix_df.shape[0]} rows x {hourly_matrix_df.shape[1]} columns)...")
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            # Sheet 1: Daily Summary
            daily_summary_df.to_excel(writer, sheet_name="Daily_Summary", index=False)
            # Sheet 2: Grid Coordinates lookup table
            coords_df.to_excel(writer, sheet_name="Grid_Coordinates", index=False)
            # Sheet 3: Hourly Matrix (8,760 rows x 703 columns - perfectly within Excel limits)
            hourly_matrix_df.to_excel(writer, sheet_name="Hourly_Grid_Matrix", index=False)

        print(f"[SUCCESS] Saved Excel file: {excel_path}")

        # Optional CSV export in long format if user requested CSV
        if csv_dir:
            os.makedirs(csv_dir, exist_ok=True)
            csv_path = os.path.join(csv_dir, f"ERA5_2m_Temperature_{year}_long.csv")
            print(f"[*] Exporting companion CSV: {csv_path}...")
            # Unpivot / melt to long format: [Timestamp, Grid, Temp_C]
            long_df = pd.melt(hourly_matrix_df, id_vars=["Timestamp"], value_vars=grid_col_names,
                              var_name="Column_Name", value_name="Temperature_Celsius")
            long_df = long_df.merge(coords_df[["Column_Name", "Latitude", "Longitude"]], on="Column_Name", how="left")
            long_df.drop(columns=["Column_Name"], inplace=True)
            long_df.to_csv(csv_path, index=False)
            print(f"[SUCCESS] Saved CSV file: {csv_path}")

def main():
    parser = argparse.ArgumentParser(description="Download ERA5 and save as Year-by-Year Excel files.")
    parser.add_argument("--start_year", type=int, default=1966, help="First year to process (default: 1966)")
    parser.add_argument("--end_year", type=int, default=1975, help="Last year to process (default: 1975)")
    parser.add_argument("--downloads_dir", type=str, default="./downloads_nc", help="Folder for raw NetCDF files")
    parser.add_argument("--excel_dir", type=str, default="./excel_output", help="Folder for generated Excel files")
    parser.add_argument("--keep_nc", action="store_true", help="Keep raw NetCDF files after Excel conversion (default: False to save disk space)")
    parser.add_argument("--export_csv", action="store_true", help="Also generate long-format CSV files")
    parser.add_argument("--csv_dir", type=str, default="./csv_output", help="Folder for CSV files")
    parser.add_argument("--api_key", type=str, default=None, help="Copernicus CDS API Personal Access Token")
    args = parser.parse_args()

    client = get_cds_client(args.api_key)

    total_years = args.end_year - args.start_year + 1
    print("\n" + "=" * 65)
    print(f" ERA5 DOWNLOAD & EXCEL CONVERSION PIPELINE ")
    print(f" Range: {args.start_year} to {args.end_year} ({total_years} years)")
    print(f" Bounding Box: {AREA} [North, West, South, East]")
    print(f" Output Folder: {os.path.abspath(args.excel_dir)}")
    print(f" Auto-cleanup NetCDFs: {'No (keep_nc=True)' if args.keep_nc else 'Yes (saving ~1GB per 80 years)'}")
    print("=" * 65 + "\n")

    failed_years = []

    for yr in range(args.start_year, args.end_year + 1):
        excel_target = os.path.join(args.excel_dir, f"ERA5_2m_Temperature_{yr}.xlsx")
        if os.path.exists(excel_target):
            print(f"[*] Year {yr} is already completed ({excel_target}). Skipping.")
            continue

        try:
            nc_file = download_year(client, yr, args.downloads_dir)
            convert_nc_to_excel_and_csv(
                nc_path=nc_file,
                year=yr,
                excel_dir=args.excel_dir,
                csv_dir=args.csv_dir if args.export_csv else None
            )
            # Remove raw NetCDF to save disk space unless keep_nc is requested
            if not args.keep_nc and os.path.exists(nc_file):
                try:
                    os.remove(nc_file)
                    print(f"[*] Cleaned up temporary raw NetCDF: {nc_file}")
                except Exception as del_err:
                    print(f"[!] Note: Could not remove {nc_file}: {del_err}")

        except Exception as e:
            print(f"\n[ERROR] Encountered error processing year {yr}: {e}")
            failed_years.append(yr)
            with open("failed_years.log", "a", encoding="utf-8") as fl:
                fl.write(f"Year {yr}: {e}\n")
            print(f"Logged failure for {yr}. Continuing with remaining years...\n")

    print("\n" + "=" * 65)
    print("Pipeline finished.")
    if failed_years:
        print(f"The following years encountered issues and were skipped: {failed_years}")
        print("Check failed_years.log for details.")
    else:
        print("All requested years completed successfully!")
    print("=" * 65)

if __name__ == "__main__":
    main()
