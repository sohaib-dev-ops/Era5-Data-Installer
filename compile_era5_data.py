"""
ERA5 Compilation Pipeline (1940 - 2026)
=======================================
High-performance compilation of yearly ERA5 Excel files into:
1. Standardized 10-year decade workbooks (1940-1949, 1950-1959, ..., 2020-2026)
2. Master unified Full Data workbook (1940-2026)

Key Features:
- Multi-core parallel extraction of daily means from yearly hourly files
- Automated disk caching for instantaneous re-runs
- High-speed, constant-memory Excel streaming via xlsxwriter
- Strict styling matching Copernicus ERA5 specification:
  * Dark Blue header fill (#1F497D) with white Calibri bold text
  * Precision freeze panes (D2 for matrix, E2 for regional stats, A2 for coordinates)
  * Exact numeric formatting ('0.00' for temperatures and coordinates)
  * Column widths optimized for readability
"""

import os
import sys
import glob
import time
import pickle
import argparse
import numpy as np
import pandas as pd
import openpyxl
import xlsxwriter
from concurrent.futures import ProcessPoolExecutor, as_completed

DECADE_RANGES = [
    (1940, 1949),
    (1950, 1959),
    (1960, 1969),
    (1970, 1979),
    (1980, 1989),
    (1990, 1999),
    (2000, 2009),
    (2010, 2019),
    (2020, 2026),
]

def extract_year_worker(fpath):
    """
    Worker function executed in parallel processes.
    Reads a single yearly ERA5 Excel file and aggregates hourly values to daily means.
    """
    t0 = time.time()
    fname = os.path.basename(fpath)
    yr = int(fname.split("_")[-1].replace(".xlsx", ""))

    wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)

    # 1. Read Regional Daily_Summary
    ws_ds = wb["Daily_Summary"]
    ds_iter = ws_ds.iter_rows(values_only=True)
    next(ds_iter)  # Skip header

    reg_summary_records = []
    orig_reg_means = []
    for r in ds_iter:
        d_val = r[0]
        if isinstance(d_val, str):
            d_obj = pd.to_datetime(d_val)
        else:
            d_obj = d_val

        r_mean = round(float(r[1]), 2)
        r_min = round(float(r[2]), 2)
        r_max = round(float(r[3]), 2)
        reg_summary_records.append([
            int(d_obj.year),
            int(d_obj.month),
            int(d_obj.day),
            d_obj.strftime("%Y-%m-%d"),
            r_mean,
            r_min,
            r_max
        ])
        orig_reg_means.append(r_mean)

    # 2. Read Hourly_Grid_Matrix
    ws_hm = wb["Hourly_Grid_Matrix"]
    hm_iter = ws_hm.iter_rows(values_only=True)
    headers = next(hm_iter)
    grid_cols = list(headers[1:])

    raw_dates = []
    raw_vals = []
    for r in hm_iter:
        raw_dates.append(r[0])
        raw_vals.append(r[1:])
    wb.close()

    arr = np.array(raw_vals, dtype=np.float32)

    # Universal day aggregation handling full 24h days and partial days (e.g. 2026)
    dt_series = pd.to_datetime(raw_dates)
    day_dates = dt_series.date
    unique_days, indices, counts = np.unique(day_dates, return_index=True, return_counts=True)
    sort_order = np.argsort(indices)
    unique_days = unique_days[sort_order]
    indices = indices[sort_order]
    counts = counts[sort_order]

    num_days = len(unique_days)
    daily_means = np.round(np.add.reduceat(arr, indices, axis=0) / counts[:, None], 2)  # shape (num_days, 702)

    years = np.array([d.year for d in unique_days], dtype=int)
    months = np.array([d.month for d in unique_days], dtype=int)
    days = np.array([d.day for d in unique_days], dtype=int)

    # Verification: check discrepancy between daily grid mean and Regional_Mean_C
    calc_reg_means = np.round(daily_means.mean(axis=1), 2)
    orig_arr = np.array(orig_reg_means[:num_days], dtype=np.float32)
    max_diff = float(np.max(np.abs(calc_reg_means - orig_arr)))

    elapsed = time.time() - t0
    return {
        "year": yr,
        "filename": fname,
        "num_days": num_days,
        "years": years,
        "months": months,
        "days": days,
        "daily_means": daily_means,
        "grid_cols": grid_cols,
        "reg_summary": reg_summary_records,
        "max_diff": max_diff,
        "elapsed": elapsed
    }


def load_or_extract_all_data(
    input_dir="excel_output",
    start_year=1940,
    end_year=2026,
    cache_path="compiled_cache/daily_cache_1940_2026.pkl",
    max_workers=8,
    force_extract=False
):
    """
    Extracts daily data for all years in parallel, with automatic caching.
    """
    # 1. Master Grid Coordinates Lookup from first year
    first_file = os.path.join(input_dir, f"ERA5_2m_Temperature_{start_year}.xlsx")
    if not os.path.exists(first_file):
        raise FileNotFoundError(f"Base file not found: {first_file}")

    print(f"[*] Extracting Grid Coordinates master lookup from {first_file}...")
    wb_sample = openpyxl.load_workbook(first_file, read_only=True, data_only=True)
    ws_coords = wb_sample["Grid_Coordinates"]
    coords_rows = list(ws_coords.iter_rows(values_only=True))
    coords_headers = list(coords_rows[0])
    coords_data = coords_rows[1:]
    wb_sample.close()
    print(f"    Loaded {len(coords_data)} grid coordinates points with headers: {coords_headers}")

    # 2. Check cache
    if not force_extract and os.path.exists(cache_path):
        print(f"[*] Found existing extraction cache: {cache_path}")
        try:
            with open(cache_path, "rb") as f:
                cached_data = pickle.load(f)
            cached_years = set(cached_data.keys())
            required_years = set(range(start_year, end_year + 1))
            if required_years.issubset(cached_years):
                print(f"[+] Loaded all {len(required_years)} years from cache instantaneously!")
                grid_cols = cached_data[start_year]["grid_cols"]
                return coords_data, coords_headers, grid_cols, cached_data
            else:
                print("[-] Cache missing some required years. Re-extracting...")
        except Exception as e:
            print(f"[!] Warning reading cache ({e}). Re-extracting...")

    # 3. Locate files to process
    yearly_files = {}
    missing = []
    for yr in range(start_year, end_year + 1):
        fpath = os.path.join(input_dir, f"ERA5_2m_Temperature_{yr}.xlsx")
        if not os.path.exists(fpath):
            missing.append(fpath)
        else:
            yearly_files[yr] = fpath

    if missing:
        raise FileNotFoundError(f"Missing {len(missing)} yearly files: {missing[:5]}...")

    total_files = len(yearly_files)
    print(f"\n[*] Extracting daily data from {total_files} yearly files using {max_workers} parallel workers...")
    t0 = time.time()

    year_data_dict = {}
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(extract_year_worker, fpath): yr for yr, fpath in yearly_files.items()}
        completed = 0
        for fut in as_completed(futures):
            res = fut.result()
            yr = res["year"]
            year_data_dict[yr] = res
            completed += 1
            print(f"  [{completed:2d}/{total_files}] Year {yr}: {res['num_days']} days (diff={res['max_diff']:.4f}°C) in {res['elapsed']:.1f}s", flush=True)

    total_extract_time = time.time() - t0
    print(f"[+] All {total_files} years extracted in {total_extract_time:.1f}s (Average: {total_extract_time/total_files:.2f}s per year)", flush=True)

    # 4. Save cache
    os.makedirs(os.path.dirname(os.path.abspath(cache_path)), exist_ok=True)
    try:
        with open(cache_path, "wb") as f:
            pickle.dump(year_data_dict, f, protocol=pickle.HIGHEST_PROTOCOL)
        cache_mb = os.path.getsize(cache_path) / (1024 * 1024)
        print(f"[+] Saved extraction cache to {cache_path} ({cache_mb:.1f} MB)")
    except Exception as e:
        print(f"[!] Warning saving cache: {e}")

    grid_cols = year_data_dict[start_year]["grid_cols"]
    return coords_data, coords_headers, grid_cols, year_data_dict


def write_excel_workbook(
    output_path,
    start_year,
    end_year,
    year_data_dict,
    coords_data,
    coords_headers,
    grid_cols,
    include_year_tabs=True,
    title_label=""
):
    """
    Writes a fully-styled multi-sheet Excel file using constant-memory streaming xlsxwriter.
    """
    t0 = time.time()
    desc = title_label or f"{start_year} to {end_year}"
    print(f"\n[*] Building workbook: {os.path.basename(output_path)} ({desc})...")

    wb = xlsxwriter.Workbook(output_path, {"constant_memory": True})

    # Header style (Navy Blue #1F497D, white bold text, centered)
    header_fmt = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#FFFFFF",
        "bg_color": "#1F497D",
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "border_color": "#D9D9D9",
    })

    # Data styles
    center_int_fmt = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "align": "center",
        "valign": "vcenter"
    })
    center_str_fmt = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "align": "center",
        "valign": "vcenter"
    })
    coord_name_fmt = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "align": "left",
        "valign": "vcenter"
    })
    num_fmt = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "num_format": "0.00",
        "align": "right",
        "valign": "vcenter"
    })

    # -------------------------------------------------------------
    # Sheet 1: Grid_Coordinates
    # -------------------------------------------------------------
    ws_gc = wb.add_worksheet("Grid_Coordinates")
    ws_gc.freeze_panes(1, 0)  # Freeze row 1 (A2)
    ws_gc.set_column(0, 0, 12)  # Col A (Grid_ID)
    ws_gc.set_column(1, 1, 26)  # Col B (Column_Name)
    ws_gc.set_column(2, 2, 12)  # Col C (Latitude)
    ws_gc.set_column(3, 3, 12)  # Col D (Longitude)

    ws_gc.write_row(0, 0, coords_headers, header_fmt)
    for r_idx, row in enumerate(coords_data, start=1):
        ws_gc.write(r_idx, 0, row[0], center_str_fmt)
        ws_gc.write(r_idx, 1, row[1], coord_name_fmt)
        ws_gc.write(r_idx, 2, float(row[2]), num_fmt)
        ws_gc.write(r_idx, 3, float(row[3]), num_fmt)

    # -------------------------------------------------------------
    # Sheet 2: Regional_Daily_Summary
    # -------------------------------------------------------------
    ws_reg = wb.add_worksheet("Regional_Daily_Summary")
    ws_reg.freeze_panes(1, 4)  # Freeze row 1 and cols A-D (E2)
    ws_reg.set_column(0, 2, 8)   # Cols A-C: Year, Month, Day
    ws_reg.set_column(3, 3, 14)  # Col D: Date
    ws_reg.set_column(4, 6, 16)  # Cols E-G: Regional_Mean_C, Min_C, Max_C

    reg_headers = ["Year", "Month", "Day", "Date", "Regional_Mean_C", "Regional_Min_C", "Regional_Max_C"]
    ws_reg.write_row(0, 0, reg_headers, header_fmt)

    reg_row_idx = 1
    for yr in range(start_year, end_year + 1):
        for rec in year_data_dict[yr]["reg_summary"]:
            ws_reg.write(reg_row_idx, 0, rec[0], center_int_fmt)
            ws_reg.write(reg_row_idx, 1, rec[1], center_int_fmt)
            ws_reg.write(reg_row_idx, 2, rec[2], center_int_fmt)
            ws_reg.write(reg_row_idx, 3, rec[3], center_str_fmt)
            ws_reg.write(reg_row_idx, 4, rec[4], num_fmt)
            ws_reg.write(reg_row_idx, 5, rec[5], num_fmt)
            ws_reg.write(reg_row_idx, 6, rec[6], num_fmt)
            reg_row_idx += 1

    # -------------------------------------------------------------
    # Sheet 3: Continuous Grid Matrix Tab
    # -------------------------------------------------------------
    continuous_title = f"{start_year}_{end_year}_Continuous"
    ws_cont = wb.add_worksheet(continuous_title)
    ws_cont.freeze_panes(1, 3)  # Freeze row 1 and cols A-C (D2)
    ws_cont.set_column(0, 2, 8)   # Cols A-C: Year, Month, Day
    ws_cont.set_column(3, 3 + len(grid_cols) - 1, 12)  # Grid columns

    matrix_headers = ["Year", "Month", "Day"] + grid_cols
    ws_cont.write_row(0, 0, matrix_headers, header_fmt)

    cont_row_idx = 1
    for yr in range(start_year, end_year + 1):
        y_data = year_data_dict[yr]
        n_days = y_data["num_days"]
        y_arr = y_data["years"]
        m_arr = y_data["months"]
        d_arr = y_data["days"]
        means_list = y_data["daily_means"].tolist()

        for d_i in range(n_days):
            row_vals = [int(y_arr[d_i]), int(m_arr[d_i]), int(d_arr[d_i])] + means_list[d_i]
            ws_cont.write_row(cont_row_idx, 0, row_vals)
            cont_row_idx += 1

    # -------------------------------------------------------------
    # Optional Sheets: Individual Year Tabs
    # -------------------------------------------------------------
    if include_year_tabs:
        for yr in range(start_year, end_year + 1):
            ws_yr = wb.add_worksheet(str(yr))
            ws_yr.freeze_panes(1, 3)  # Freeze row 1 and cols A-C (D2)
            ws_yr.set_column(0, 2, 8)
            ws_yr.set_column(3, 3 + len(grid_cols) - 1, 12)
            ws_yr.write_row(0, 0, matrix_headers, header_fmt)

            y_data = year_data_dict[yr]
            n_days = y_data["num_days"]
            y_arr = y_data["years"]
            m_arr = y_data["months"]
            d_arr = y_data["days"]
            means_list = y_data["daily_means"].tolist()

            for d_i in range(n_days):
                row_vals = [int(y_arr[d_i]), int(m_arr[d_i]), int(d_arr[d_i])] + means_list[d_i]
                ws_yr.write_row(d_i + 1, 0, row_vals)

    wb.close()
    elapsed = time.time() - t0
    fsize_mb = os.path.getsize(output_path) / (1024 * 1024)
    print(f"[+] Finished {os.path.basename(output_path)}: {fsize_mb:.1f} MB in {elapsed:.1f}s")
    return output_path, fsize_mb


def compile_all_decades_and_full(
    input_dir="excel_output",
    output_dir=".",
    cache_path="compiled_cache/daily_cache_1940_2026.pkl",
    max_workers=8,
    force_extract=False,
    include_year_tabs_in_full=False,
    build_decades=True,
    build_full=True
):
    """
    Main compilation orchestrator.
    Extracts all data once, then produces all requested workbooks.
    """
    overall_start = time.time()
    print("=" * 75)
    print(" ERA5 COMPILATION PIPELINE (1940 - 2026)")
    print(f" Input Directory : {os.path.abspath(input_dir)}")
    print(f" Output Directory: {os.path.abspath(output_dir)}")
    print(f" Workers         : {max_workers}")
    print(f" Build Decades   : {build_decades}")
    print(f" Build Full Data : {build_full}")
    print("=" * 75)

    os.makedirs(output_dir, exist_ok=True)

    # Step 1: Extract all daily data from 1940 to 2026
    coords_data, coords_headers, grid_cols, year_data_dict = load_or_extract_all_data(
        input_dir=input_dir,
        start_year=1940,
        end_year=2026,
        cache_path=cache_path,
        max_workers=max_workers,
        force_extract=force_extract
    )

    created_files = []

    # Step 2: Compile 10-year decade workbooks
    if build_decades:
        print("\n" + "=" * 75)
        print(" GENERATING 10-YEAR DECADE WORKBOOKS (1940 - 2026)")
        print("=" * 75)
        for s_yr, e_yr in DECADE_RANGES:
            fname = f"ERA5_Daily_Compiled_{s_yr}_{e_yr}.xlsx"
            fpath = os.path.join(output_dir, fname)
            out_p, size_mb = write_excel_workbook(
                output_path=fpath,
                start_year=s_yr,
                end_year=e_yr,
                year_data_dict=year_data_dict,
                coords_data=coords_data,
                coords_headers=coords_headers,
                grid_cols=grid_cols,
                include_year_tabs=True,
                title_label=f"{s_yr} to {e_yr} (Decade Workbook)"
            )
            created_files.append((fname, size_mb, f"{s_yr}-{e_yr}"))

    # Step 3: Compile Full Data workbook
    if build_full:
        print("\n" + "=" * 75)
        print(" GENERATING FULL MASTER DATA WORKBOOK (1940 - 2026)")
        print("=" * 75)
        full_fname = "ERA5_Daily_Compiled_1940_2026_Full.xlsx"
        full_fpath = os.path.join(output_dir, full_fname)
        out_p, size_mb = write_excel_workbook(
            output_path=full_fpath,
            start_year=1940,
            end_year=2026,
            year_data_dict=year_data_dict,
            coords_data=coords_data,
            coords_headers=coords_headers,
            grid_cols=grid_cols,
            include_year_tabs=include_year_tabs_in_full,
            title_label="1940 to 2026 Full Master Dataset (31,679 Daily Records)"
        )
        created_files.append((full_fname, size_mb, "1940-2026 Full Dataset"))

    total_time = time.time() - overall_start

    print("\n" + "=" * 75)
    print(" ALL REQUESTED COMPILATIONS COMPLETED SUCCESSFULLY!")
    print(f" Total Elapsed Time: {total_time:.1f}s ({total_time / 60:.1f} minutes)")
    print("=" * 75)
    print(" SUMMARY OF CREATED FILES:")
    for fname, size_mb, desc in created_files:
        print(f"  * {fname:<42} | {size_mb:6.1f} MB | {desc}")
    print("=" * 75)
    return created_files


# Backward compatibility for compile_10_years
def compile_10_years(
    input_dir="excel_output",
    output_file="ERA5_Daily_Compiled_1940_1949.xlsx",
    start_year=1940,
    end_year=1949,
    include_continuous_sheet=True
):
    """Legacy interface maintained for backward compatibility."""
    coords_data, coords_headers, grid_cols, year_data_dict = load_or_extract_all_data(
        input_dir=input_dir,
        start_year=start_year,
        end_year=end_year
    )
    write_excel_workbook(
        output_path=output_file,
        start_year=start_year,
        end_year=end_year,
        year_data_dict=year_data_dict,
        coords_data=coords_data,
        coords_headers=coords_headers,
        grid_cols=grid_cols,
        include_year_tabs=True,
        title_label=f"{start_year} to {end_year}"
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Compile yearly ERA5 files into decade sets and full continuous Excel dataset.")
    parser.add_argument("--all", action="store_true", default=True, help="Compile all 10-year sets and full master dataset (default)")
    parser.add_argument("--decades_only", action="store_true", help="Compile only the 10-year sets (1940-1949 to 2020-2026)")
    parser.add_argument("--full_only", action="store_true", help="Compile only the master full dataset (1940-2026)")
    parser.add_argument("--start_year", type=int, default=None, help="Custom start year (if compiling custom range)")
    parser.add_argument("--end_year", type=int, default=None, help="Custom end year (if compiling custom range)")
    parser.add_argument("--output_file", type=str, default=None, help="Custom output file (if compiling custom range)")
    parser.add_argument("--output_dir", type=str, default=".", help="Directory to save compiled files (default: .)")
    parser.add_argument("--input_dir", type=str, default="excel_output", help="Directory containing yearly ERA5 files")
    parser.add_argument("--workers", type=int, default=8, help="Number of parallel extraction workers (default: 8)")
    parser.add_argument("--force_extract", action="store_true", help="Force re-extraction ignoring disk cache")
    parser.add_argument("--include_year_tabs_in_full", action="store_true", help="Include 87 individual year tabs in the full master file")

    args = parser.parse_args()

    # If user specified a custom range
    if args.start_year is not None and args.end_year is not None:
        out_f = args.output_file or os.path.join(args.output_dir, f"ERA5_Daily_Compiled_{args.start_year}_{args.end_year}.xlsx")
        compile_10_years(
            input_dir=args.input_dir,
            output_file=out_f,
            start_year=args.start_year,
            end_year=args.end_year
        )
    else:
        build_decades = not args.full_only
        build_full = not args.decades_only

        compile_all_decades_and_full(
            input_dir=args.input_dir,
            output_dir=args.output_dir,
            max_workers=args.workers,
            force_extract=args.force_extract,
            include_year_tabs_in_full=args.include_year_tabs_in_full,
            build_decades=build_decades,
            build_full=build_full
        )
