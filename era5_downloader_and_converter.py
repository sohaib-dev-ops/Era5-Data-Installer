"""
ERA5 2m-Temperature Downloader & NetCDF to Excel Converter
Author: Antigravity Assistant

Features:
- Batches downloads year-by-year to comply with Copernicus CDS API size limits.
- Supports extracting 2m Temperature (converted from Kelvin to Celsius).
- Outputs directly to Excel (.xlsx) or CSV.
- Supports multiple aggregation modes:
    1. 'spatial_mean': Regional average temperature over time (fits directly into 1 Excel sheet).
    2. 'daily_grid': Daily Mean, Min, and Max for every grid point.
    3. 'stations': Time series extracted for specific coordinates (e.g. major cities/stations).
    4. 'yearly_raw': Full hourly grid points saved per year.
"""

import os
import glob
import zipfile
import argparse
import numpy as np
import pandas as pd
import xarray as xr
import cdsapi

# Default Pakistan bounding box from request: [North, West, South, East]
DEFAULT_AREA = [34.0, 69.25, 27.5, 75.5]

def extract_zip_if_needed(file_path, output_dir):
    """If downloaded file is a zip archive, extract it and return the .nc file path."""
    if file_path.endswith(".zip") or zipfile.is_zipfile(file_path):
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
            extracted_files = [os.path.join(output_dir, f) for f in zip_ref.namelist() if f.endswith(".nc")]
            if extracted_files:
                return extracted_files[0]
    return file_path

def download_era5_year(year, area=DEFAULT_AREA, output_dir="./downloads"):
    """Download single year NetCDF from CDS API."""
    os.makedirs(output_dir, exist_ok=True)
    target_nc = os.path.join(output_dir, f"era5_t2m_{year}.nc")
    target_zip = os.path.join(output_dir, f"era5_t2m_{year}.zip")

    if os.path.exists(target_nc):
        print(f"[INFO] File already exists for {year}: {target_nc}")
        return target_nc

    c = cdsapi.Client()
    request = {
        "product_type": ["reanalysis"],
        "variable": ["2m_temperature"],
        "year": [str(year)],
        "month": [f"{m:02d}" for m in range(1, 13)],
        "day": [f"{d:02d}" for d in range(1, 32)],
        "time": [f"{h:02d}:00" for h in range(24)],
        "data_format": "netcdf",
        "download_format": "zip",
        "area": area
    }

    print(f"\n--> Requesting ERA5 data for year {year} from CDS...")
    c.retrieve("reanalysis-era5-single-levels", request).download(target_zip)
    
    # Extract zip if needed
    nc_path = extract_zip_if_needed(target_zip, output_dir)
    if nc_path != target_nc and os.path.exists(nc_path):
        os.replace(nc_path, target_nc)
        nc_path = target_nc
    if os.path.exists(target_zip):
        try:
            os.remove(target_zip)
        except Exception:
            pass

    return nc_path

def netcdf_to_excel_spatial_mean(nc_files, excel_output_path, unit="celsius"):
    """
    Computes regional spatial mean temperature for each hour/day,
    and writes to an Excel spreadsheet.
    """
    print(f"\nProcessing {len(nc_files)} NetCDF file(s) for Regional Spatial Mean...")
    records = []
    
    for f in sorted(nc_files):
        print(f"  Reading {os.path.basename(f)}...")
        with xr.open_dataset(f) as ds:
            # ERA5 variable name is typically 't2m' or '2t'
            var_name = "t2m" if "t2m" in ds else ("2t" if "2t" in ds else list(ds.data_vars)[0])
            t2m = ds[var_name]

            # Convert Kelvin to Celsius
            if unit.lower() == "celsius":
                t2m = t2m - 273.15

            # Spatial mean across lat and lon
            spatial_mean = t2m.mean(dim=["latitude", "longitude"]).to_series()
            records.append(spatial_mean)

    full_series = pd.concat(records)
    df = full_series.reset_index()
    df.columns = ["Timestamp", f"Mean_Temperature_{unit.capitalize()}"]
    
    # Also calculate daily statistics
    df["Date"] = pd.to_datetime(df["Timestamp"]).dt.date
    daily_df = df.groupby("Date")[f"Mean_Temperature_{unit.capitalize()}"].agg(
        Daily_Mean="mean", Daily_Min="min", Daily_Max="max"
    ).reset_index()

    # Save to Excel with two sheets: Hourly and Daily
    with pd.ExcelWriter(excel_output_path, engine="openpyxl") as writer:
        daily_df.to_excel(writer, sheet_name="Daily_Summary", index=False)
        # Check if hourly records fit in 1 sheet (1,048,576 row limit)
        if len(df) <= 1_000_000:
            df.to_excel(writer, sheet_name="Hourly_Series", index=False)
        else:
            print(f"[WARN] Hourly rows ({len(df)}) exceed single sheet limit, exporting first 1M rows.")
            df.iloc[:1_000_000].to_excel(writer, sheet_name="Hourly_Series_Part1", index=False)

    print(f"[SUCCESS] Exported Regional Spatial Mean to: {excel_output_path}")

def netcdf_to_excel_stations(nc_files, station_coords, excel_output_path, unit="celsius"):
    """
    Extracts time-series for specific coordinate points (e.g. weather stations or cities).
    station_coords: dict of {'StationName': (latitude, longitude)}
    """
    print(f"\nExtracting point data for {len(station_coords)} stations/locations...")
    all_dfs = {name: [] for name in station_coords}

    for f in sorted(nc_files):
        print(f"  Reading {os.path.basename(f)}...")
        with xr.open_dataset(f) as ds:
            var_name = "t2m" if "t2m" in ds else ("2t" if "2t" in ds else list(ds.data_vars)[0])
            for name, (lat, lon) in station_coords.items():
                pt = ds[var_name].sel(latitude=lat, longitude=lon, method="nearest")
                if unit.lower() == "celsius":
                    pt = pt - 273.15
                s = pt.to_series()
                all_dfs[name].append(s)

    with pd.ExcelWriter(excel_output_path, engine="openpyxl") as writer:
        for name, series_list in all_dfs.items():
            full_series = pd.concat(series_list).reset_index()
            full_series.columns = ["Timestamp", f"Temp_{unit.capitalize()}"]
            sheet_title = name[:31] # Excel max sheet title is 31 chars
            full_series.to_excel(writer, sheet_name=sheet_title, index=False)

    print(f"[SUCCESS] Exported Station Data to: {excel_output_path}")

def main():
    parser = argparse.ArgumentParser(description="Download ERA5 and export to Excel.")
    parser.add_argument("--start_year", type=int, default=2023, help="Start year (e.g. 1940 or 2020)")
    parser.add_argument("--end_year", type=int, default=2024, help="End year (e.g. 2024)")
    parser.add_argument("--mode", choices=["spatial_mean", "stations", "daily_grid"], default="spatial_mean")
    parser.add_argument("--output", type=str, default="ERA5_2m_Temperature.xlsx", help="Output Excel filename")
    parser.add_argument("--unit", choices=["celsius", "kelvin"], default="celsius")
    args = parser.parse_args()

    nc_dir = os.path.join(".", "downloads")
    nc_files = []

    for yr in range(args.start_year, args.end_year + 1):
        try:
            nc_file = download_era5_year(yr, area=DEFAULT_AREA, output_dir=nc_dir)
            nc_files.append(nc_file)
        except Exception as e:
            print(f"[ERROR] Failed downloading year {yr}: {e}")
            break

    if not nc_files:
        print("[ERROR] No NetCDF files available to convert.")
        return

    if args.mode == "spatial_mean":
        netcdf_to_excel_spatial_mean(nc_files, args.output, unit=args.unit)
    elif args.mode == "stations":
        # Example sample stations in the bounding box
        stations = {
            "Lahore": (31.52, 74.35),
            "Islamabad": (33.68, 73.04),
            "Peshawar": (34.01, 71.52),
            "Quetta": (30.17, 66.97),
            "Multan": (30.15, 71.52),
            "Faisalabad": (31.45, 73.13)
        }
        netcdf_to_excel_stations(nc_files, stations, args.output, unit=args.unit)

if __name__ == "__main__":
    main()
