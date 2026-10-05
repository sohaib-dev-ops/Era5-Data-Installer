"""
ERA5 Meteorological Stations Transformation Pipeline (1940 - 2026)
==================================================================
Transforms raw ERA5 702-point grid matrix data (1940-2026) into proper
meteorological stations and full-city entities matching Pakistan
Meteorological Department (PMD) standards.

Features:
- High-precision 4-point Bilinear Interpolation from 0.25° ERA5 grid.
- Consolidated Full City columns for multi-station cities:
  * Lahore (City Average): Lahore A.P. + Lahore PBO. + MCC Lahore
  * Bahawalpur (City Average): Bahawal Pur (City) + Bahawal Pur (Airport)
  * Sialkot (City Average): Sialkot (City) + Sialkot Airport
  * D.G. Khan (City Average): D.G. Khan (City) + D.G. Khan (Aeromet)
- Detailed individual station columns for all PMD Punjab stations + Islamabad.
- Multi-sheet Master Workbook:
  1. 'Hourly_Stations': 760,000+ continuous hourly temperatures (1940-2026).
  2. 'Daily_Summary': Daily Mean, Min, and Max for every station/city (1940-2026).
  3. 'Station_Directory': Comprehensive PMD-aligned station catalog and metadata.
- Multi-core parallel extraction and automated disk caching.
- Memory-efficient streaming Excel workbook generation via xlsxwriter.
"""

import os
import sys
import time
import pickle
import argparse
import numpy as np
import pandas as pd
import openpyxl
import xlsxwriter
from concurrent.futures import ProcessPoolExecutor, as_completed

# ---------------------------------------------------------------------------
# Station Metadata (aligned with PMD Official Records)
# ---------------------------------------------------------------------------
STATIONS_DEF = [
    # (sno, index_no, station_name, city, lat_dms, lon_dms, lat_dec, lon_dec, elev, phone, is_multi)
    ("01", "41678", "Bahawal Nagar", "Bahawal Nagar", "29° 20'", "73° 51'", 29 + 20/60, 73 + 51/60, "161.05 meter", "063-9239046, 0331-7947019", False),
    ("02", "41700", "Bahawal Pur", "Bahawalpur", "29° 20'", "71° 47'", 29 + 20/60, 71 + 47/60, "110.00 meter", "0344-3668731", True),
    ("03", "41701", "Bahawal Pur (A/P)", "Bahawalpur", "29° 23'", "71° 41'", 29 + 23/60, 71 + 41/60, "119.00 meter", "0301-6503510, 0302-4527757, 0301-7700447", True),
    ("04", "41625(AP)", "Bhakkar", "Bhakkar", "31° 37'", "71° 04'", 31 + 37/60, 71 + 4/60, "162 meter", "-----", False),
    ("05", "41626(AP)", "Noorpur Thal", "Noorpur Thal", "31° 52'", "71° 54'", 31 + 52/60, 71 + 54/60, "186 meter", "0346-8611808", False),
    ("06", "41627(AP)", "Jauharabad", "Jauharabad", "32° 30'", "72° 26'", 32 + 30/60, 72 + 26/60, "187 meter", "0300-6055391, 0300-7736294", False),
    ("07", "41630", "Faisalabad", "Faisalabad", "31° 26'", "73° 08'", 31 + 26/60, 73 + 8/60, "185.6 meter", "041-9201623, 0333-4751778", False),
    ("08", "41598", "Jhelum", "Jhelum", "32° 56'", "73° 44'", 32 + 56/60, 73 + 44/60, "287.19 meter", "0544-627922, 0346-5994136", False),
    ("09", "41718", "Khanpur", "Khanpur", "28° 39'", "70° 41'", 28 + 39/60, 70 + 41/60, "88.41 meter", "0302-7542967, 0334-3668731", False),
    ("10", "41641", "Lahore A.P.", "Lahore", "31° 35'", "74° 24'", 31 + 35/60, 74 + 24/60, "216.15 meter", "042-9240402-3", True),
    ("11", "41640", "Lahore PBO.", "Lahore", "31° 33'", "74° 20'", 31 + 33/60, 74 + 20/60, "214.00 meter", "042-37581496", True),
    ("12", "41675", "Multan", "Multan", "30° 12'", "71° 26'", 30 + 12/60, 71 + 26/60, "121.95 meter", "061-9201150, 061-9200763", False),
    ("13", "41597", "Mandi Bahauddin", "Mandi Bahauddin", "32° 58'", "73° 48'", 32 + 58/60, 73 + 48/60, "252.97 meter", "0334-4903292, 0300-7562425", False),
    ("14", "41600", "Sialkot", "Sialkot", "32° 31'", "74° 32'", 32 + 31/60, 74 + 32/60, "255.1 meter", "052-4267195", True),
    ("15", "41600", "Sialkot Airport", "Sialkot", "32° 32'", "74° 02'", 32 + 32/60, 74 + 2/60, "240 meter", "052-6633033", True),
    ("16", "41593", "Sargodha", "Sargodha", "32° 03'", "72° 40'", 32 + 3/60, 72 + 40/60, "187 meter", "048-9230206, 0306-6705566", False),
    ("17", "41638", "Toba Tek Sing", "Toba Tek Singh", "30° 59'", "72° 47'", 30 + 59/60, 72 + 47/60, "155 meter", "0347-6899541, 0300-2345426", False),
    ("18", "41652", "D.G. Khan", "D.G. Khan", "30° 03'", "70° 38'", 30 + 3/60, 70 + 38/60, "148.1 meter", "0332-0799908, 0333-5303956", True),
    ("19", "NIL", "D.G. Khan (Aeromet)", "D.G. Khan", "29° 58'", "70° 20'", 29 + 58/60, 70 + 20/60, "141 meter", "---", True),
    ("20", "41636", "Jhang", "Jhang", "31° 16'", "72° 19'", 31 + 16/60, 72 + 19/60, "---", "0336-7924489, 0345-7609030", False),
    ("21", "41576", "Mangla", "Mangla", "33° 04'", "73° 38'", 33 + 4/60, 73 + 38/60, "283.3 meter", "0301-7123667, 0346-2730326", False),
    ("22", "41680", "Sahiwal", "Sahiwal", "30° 39'", "73° 10'", 30 + 39/60, 73 + 10/60, "172 meter", "0334-6622924, 0345-8281133", False),
    ("23", "41628(AP)", "Chakwal", "Chakwal", "32° 55'", "72° 51'", 32 + 55/60, 72 + 51/60, "519 meter", "0333-5903226", False),
    ("24", "41642(AP)", "Gujranwala", "Gujranwala", "32° 22'", "74° 21'", 32 + 22/60, 74 + 21/60, "227 meter", "0331-4628601", False),
    ("25", "41646", "Okara", "Okara", "30° 48'", "73° 26'", 30 + 48/60, 73 + 26/60, "180 meter", "0334-6622924, 0345-8281133", False),
    ("26", "41716", "Rahim Yar Khan", "Rahim Yar Khan", "28° 26'", "70° 19'", 28 + 26/60, 70 + 19/60, "82.93 meter", "0336-6800891", False),
    ("27", "NIL", "Gujrat", "Gujrat", "32° 34'", "74° 4'", 32 + 34/60, 74 + 4/60, "---", "0346-4277110", False),
    ("28", "NIL", "MCC Lahore", "Lahore", "31° 35'", "74° 24'", 31 + 35/60, 74 + 24/60, "216.15 meter", "042-37580645, 042-3755546", True),
    ("29", "41571", "Islamabad", "Islamabad", "33° 41'", "73° 04'", 33 + 41/60, 73 + 4/60, "508 meter", "051-9250365", False),
]

# Order of City Columns (24 total)
FULL_CITY_COLUMNS = [
    "Islamabad (City)",
    "Lahore (City Average)",
    "Faisalabad (City)",
    "Multan (City)",
    "Gujranwala (City)",
    "Sialkot (City Average)",
    "Bahawalpur (City Average)",
    "Sargodha (City)",
    "D.G. Khan (City Average)",
    "Sahiwal (City)",
    "Okara (City)",
    "Rahim Yar Khan (City)",
    "Jhang (City)",
    "Gujrat (City)",
    "Jhelum (City)",
    "Chakwal (City)",
    "Bahawal Nagar (City)",
    "Mandi Bahauddin (City)",
    "Toba Tek Singh (City)",
    "Khanpur (City)",
    "Bhakkar (City)",
    "Jauharabad (City)",
    "Noorpur Thal (City)",
    "Mangla (City)"
]

# Order of Sub-station Columns (9 total)
SUBSTATION_COLUMNS = [
    "Lahore - Airport (A.P.)",
    "Lahore - PBO.",
    "Lahore - MCC",
    "Bahawal Pur - City",
    "Bahawal Pur - Airport (A/P)",
    "Sialkot - City",
    "Sialkot - Airport",
    "D.G. Khan - City",
    "D.G. Khan - Aeromet"
]

ALL_SERIES_COLUMNS = FULL_CITY_COLUMNS + SUBSTATION_COLUMNS

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


def build_bilinear_weights_matrix(first_yearly_file):
    """
    Constructs the (702, 33) matrix W where column j represents the bilinear weights
    of 702 ERA5 grid points for series j.
    """
    wb = openpyxl.load_workbook(first_yearly_file, read_only=True)
    coords_rows = list(wb["Grid_Coordinates"].iter_rows(values_only=True))[1:]
    wb.close()

    coord_to_idx = {}
    for i, r in enumerate(coords_rows):
        lat = round(float(r[2]), 2)
        lon = round(float(r[3]), 2)
        coord_to_idx[(lat, lon)] = i

    W = np.zeros((702, len(ALL_SERIES_COLUMNS)), dtype=np.float32)

    def get_station_pts(target_lat, target_lon):
        lat_idx_top = int(np.floor((34.0 - target_lat) / 0.25))
        lat_top = round(34.0 - lat_idx_top * 0.25, 2)
        lat_bot = round(lat_top - 0.25, 2)
        lon_idx_left = int(np.floor((target_lon - 69.25) / 0.25))
        lon_left = round(69.25 + lon_idx_left * 0.25, 2)
        lon_right = round(lon_left + 0.25, 2)

        y = (target_lat - lat_bot) / 0.25
        x = (target_lon - lon_left) / 0.25

        return [
            ((lat_top, lon_left), (1.0 - x) * y),
            ((lat_top, lon_right), x * y),
            ((lat_bot, lon_left), (1.0 - x) * (1.0 - y)),
            ((lat_bot, lon_right), x * (1.0 - y))
        ]

    station_lookup = {}
    for item in STATIONS_DEF:
        sname = item[2]
        pts = get_station_pts(item[6], item[7])
        station_lookup[sname] = pts

    def add_weights(col_name, pts, scale=1.0):
        c_idx = ALL_SERIES_COLUMNS.index(col_name)
        for (lat, lon), w in pts:
            g_idx = coord_to_idx[(lat, lon)]
            W[g_idx, c_idx] += (w * scale)

    # 1. Single-station cities
    single_mappings = {
        "Islamabad (City)": "Islamabad",
        "Faisalabad (City)": "Faisalabad",
        "Multan (City)": "Multan",
        "Gujranwala (City)": "Gujranwala",
        "Sargodha (City)": "Sargodha",
        "Sahiwal (City)": "Sahiwal",
        "Okara (City)": "Okara",
        "Rahim Yar Khan (City)": "Rahim Yar Khan",
        "Jhang (City)": "Jhang",
        "Gujrat (City)": "Gujrat",
        "Jhelum (City)": "Jhelum",
        "Chakwal (City)": "Chakwal",
        "Bahawal Nagar (City)": "Bahawal Nagar",
        "Mandi Bahauddin (City)": "Mandi Bahauddin",
        "Toba Tek Singh (City)": "Toba Tek Sing",
        "Khanpur (City)": "Khanpur",
        "Bhakkar (City)": "Bhakkar",
        "Jauharabad (City)": "Jauharabad",
        "Noorpur Thal (City)": "Noorpur Thal",
        "Mangla (City)": "Mangla",
    }
    for cname, sname in single_mappings.items():
        add_weights(cname, station_lookup[sname])

    # 2. Multi-station City Averages
    # Lahore (3 stations)
    for s in ["Lahore A.P.", "Lahore PBO.", "MCC Lahore"]:
        add_weights("Lahore (City Average)", station_lookup[s], scale=1.0 / 3.0)

    # Sialkot (2 stations)
    for s in ["Sialkot", "Sialkot Airport"]:
        add_weights("Sialkot (City Average)", station_lookup[s], scale=0.5)

    # Bahawalpur (2 stations)
    for s in ["Bahawal Pur", "Bahawal Pur (A/P)"]:
        add_weights("Bahawalpur (City Average)", station_lookup[s], scale=0.5)

    # DG Khan (2 stations)
    for s in ["D.G. Khan", "D.G. Khan (Aeromet)"]:
        add_weights("D.G. Khan (City Average)", station_lookup[s], scale=0.5)

    # 3. Individual Sub-stations
    add_weights("Lahore - Airport (A.P.)", station_lookup["Lahore A.P."])
    add_weights("Lahore - PBO.", station_lookup["Lahore PBO."])
    add_weights("Lahore - MCC", station_lookup["MCC Lahore"])
    add_weights("Bahawal Pur - City", station_lookup["Bahawal Pur"])
    add_weights("Bahawal Pur - Airport (A/P)", station_lookup["Bahawal Pur (A/P)"])
    add_weights("Sialkot - City", station_lookup["Sialkot"])
    add_weights("Sialkot - Airport", station_lookup["Sialkot Airport"])
    add_weights("D.G. Khan - City", station_lookup["D.G. Khan"])
    add_weights("D.G. Khan - Aeromet", station_lookup["D.G. Khan (Aeromet)"])

    # Sanity verification
    col_sums = W.sum(axis=0)
    for idx, cname in enumerate(ALL_SERIES_COLUMNS):
        assert abs(col_sums[idx] - 1.0) < 1e-4, f"Column {cname} sum is {col_sums[idx]} != 1.0"

    return W, station_lookup, coord_to_idx


def extract_station_year_worker(args):
    """
    Worker function executed in parallel processes.
    Reads yearly Hourly_Grid_Matrix, applies bilinear weight matrix W,
    computes hourly series and daily summary (Mean, Min, Max) for all stations.
    """
    fpath, W = args
    t0 = time.time()
    fname = os.path.basename(fpath)
    yr = int(fname.split("_")[-1].replace(".xlsx", ""))

    wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
    ws = wb["Hourly_Grid_Matrix"]
    row_iter = ws.iter_rows(values_only=True)
    next(row_iter)  # Skip header

    raw_dates = []
    raw_vals = []
    for r in row_iter:
        raw_dates.append(r[0])
        raw_vals.append(r[1:])
    wb.close()

    arr = np.array(raw_vals, dtype=np.float32)
    # Bilinear matrix multiplication: shape (N_hours, 33)
    station_hourly = np.round(arr @ W, 2)

    dt_series = pd.to_datetime(raw_dates)
    years = dt_series.year.values
    months = dt_series.month.values
    days = dt_series.day.values
    hours = dt_series.hour.values
    dt_strings = dt_series.strftime("%Y-%m-%d %H:00").values

    # Daily aggregation
    day_dates = dt_series.date
    unique_days, indices, counts = np.unique(day_dates, return_index=True, return_counts=True)
    sort_order = np.argsort(indices)
    unique_days = unique_days[sort_order]
    indices = indices[sort_order]
    counts = counts[sort_order]

    daily_summary_records = []
    for idx, cnt in zip(indices, counts):
        chunk = station_hourly[idx : idx + cnt]
        d_mean = np.round(chunk.mean(axis=0), 2)
        d_min = np.round(chunk.min(axis=0), 2)
        d_max = np.round(chunk.max(axis=0), 2)
        
        sample_dt = dt_series[idx]
        # Record format: [Year, Month, Day, Date_str, stats for each of 33 series (mean, min, max)]
        rec = [sample_dt.year, sample_dt.month, sample_dt.day, sample_dt.strftime("%Y-%m-%d")]
        for s_i in range(len(ALL_SERIES_COLUMNS)):
            rec.extend([float(d_mean[s_i]), float(d_min[s_i]), float(d_max[s_i])])
        daily_summary_records.append(rec)

    elapsed = time.time() - t0
    return {
        "year": yr,
        "num_hours": len(dt_series),
        "num_days": len(unique_days),
        "years": years,
        "months": months,
        "days": days,
        "hours": hours,
        "dt_strings": dt_strings,
        "hourly_data": station_hourly,  # shape (N_hours, 33) float32
        "daily_summary": daily_summary_records,
        "elapsed": elapsed
    }


def load_or_extract_all_stations(
    input_dir="excel_output",
    start_year=1940,
    end_year=2026,
    cache_path="compiled_cache/stations_transformed_cache_1940_2026.pkl",
    max_workers=8,
    force_extract=False
):
    """
    Extracts station hourly and daily summary for all years in parallel, with disk caching.
    """
    first_file = os.path.join(input_dir, f"ERA5_2m_Temperature_{start_year}.xlsx")
    if not os.path.exists(first_file):
        raise FileNotFoundError(f"Base file not found: {first_file}")

    print(f"[*] Building 4-point Bilinear Interpolation weight matrix from {first_file}...")
    W, station_lookup, coord_to_idx = build_bilinear_weights_matrix(first_file)
    print(f"[+] Weight matrix active: 702 ERA5 grid points -> {len(ALL_SERIES_COLUMNS)} meteorological series.")

    # Check cache
    if not force_extract and os.path.exists(cache_path):
        print(f"[*] Found station extraction cache: {cache_path}")
        try:
            with open(cache_path, "rb") as f:
                cached_data = pickle.load(f)
            cached_years = set(cached_data.keys())
            required_years = set(range(start_year, end_year + 1))
            if required_years.issubset(cached_years):
                print(f"[+] Loaded all {len(required_years)} years from cache instantaneously!")
                return cached_data, station_lookup, coord_to_idx
            else:
                print("[-] Cache incomplete. Re-extracting missing years...")
        except Exception as e:
            print(f"[!] Warning reading cache ({e}). Re-extracting...")

    # Locate files
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
    print(f"\n[*] Extracting station hourly & daily data from {total_files} files using {max_workers} workers...")
    t0 = time.time()

    worker_args = [(fpath, W) for yr, fpath in yearly_files.items()]

    year_data_dict = {}
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(extract_station_year_worker, arg): arg[0] for arg in worker_args}
        completed = 0
        for fut in as_completed(futures):
            res = fut.result()
            yr = res["year"]
            year_data_dict[yr] = res
            completed += 1
            print(f"  [{completed:2d}/{total_files}] Year {yr}: {res['num_hours']:,} hours, {res['num_days']} days in {res['elapsed']:.1f}s", flush=True)

    total_time = time.time() - t0
    print(f"[+] All {total_files} years extracted in {total_time:.1f}s (Average: {total_time/total_files:.2f}s per year)", flush=True)

    # Save cache
    os.makedirs(os.path.dirname(os.path.abspath(cache_path)), exist_ok=True)
    try:
        with open(cache_path, "wb") as f:
            pickle.dump(year_data_dict, f, protocol=pickle.HIGHEST_PROTOCOL)
        cache_mb = os.path.getsize(cache_path) / (1024 * 1024)
        print(f"[+] Saved station extraction cache to {cache_path} ({cache_mb:.1f} MB)")
    except Exception as e:
        print(f"[!] Warning saving cache: {e}")

    return year_data_dict, station_lookup, coord_to_idx


def write_stations_workbook(
    output_path,
    start_year,
    end_year,
    year_data_dict,
    station_lookup,
    coord_to_idx,
    title_label=""
):
    """
    Writes a multi-sheet master workbook:
    1. 'Hourly_Stations' (1940-2026 continuous hourly temperatures)
    2. 'Daily_Summary' (1940-2026 daily Mean, Min, Max)
    3. 'Station_Directory' (PMD catalog and coordinates)
    """
    t0 = time.time()
    desc = title_label or f"{start_year} to {end_year}"
    print(f"\n[*] Generating Excel workbook: {os.path.basename(output_path)} ({desc})...")

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

    # Header style for Sub-stations (Dark Slate #2C3E50)
    sub_header_fmt = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#FFFFFF",
        "bg_color": "#2C3E50",
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "border_color": "#D9D9D9",
    })

    # Data formatting
    center_int_fmt = wb.add_format({"font_name": "Calibri", "font_size": 11, "align": "center", "valign": "vcenter"})
    center_str_fmt = wb.add_format({"font_name": "Calibri", "font_size": 11, "align": "center", "valign": "vcenter"})
    left_str_fmt = wb.add_format({"font_name": "Calibri", "font_size": 11, "align": "left", "valign": "vcenter"})
    num_fmt = wb.add_format({"font_name": "Calibri", "font_size": 11, "num_format": "0.00", "align": "right", "valign": "vcenter"})

    # =========================================================================
    # SHEET 1: Hourly_Stations (1940 - 2026 All Hours)
    # =========================================================================
    print("  -> Writing Sheet 1: Hourly_Stations (760,000+ hourly rows)...")
    ws_hr = wb.add_worksheet("Hourly_Stations")
    ws_hr.freeze_panes(1, 5)  # Freeze row 1 and columns A-E (F2)

    # Column widths & formats
    ws_hr.set_column(0, 3, 8, center_int_fmt)    # Year, Month, Day, Hour
    ws_hr.set_column(4, 4, 18, center_str_fmt)   # Date_Time
    ws_hr.set_column(5, 5 + len(ALL_SERIES_COLUMNS) - 1, 15, num_fmt)  # Station series

    time_headers = ["Year", "Month", "Day", "Hour", "Date_Time"]
    for c_idx, h in enumerate(time_headers):
        ws_hr.write(0, c_idx, h, header_fmt)

    for c_idx, cname in enumerate(ALL_SERIES_COLUMNS, start=len(time_headers)):
        fmt = sub_header_fmt if cname in SUBSTATION_COLUMNS else header_fmt
        ws_hr.write(0, c_idx, cname, fmt)

    hr_row_idx = 1
    for yr in range(start_year, end_year + 1):
        y_data = year_data_dict[yr]
        n_hours = y_data["num_hours"]
        y_arr = y_data["years"]
        m_arr = y_data["months"]
        d_arr = y_data["days"]
        h_arr = y_data["hours"]
        dt_strs = y_data["dt_strings"]
        hourly_matrix = y_data["hourly_data"].tolist()

        for h_i in range(n_hours):
            row_vals = [
                int(y_arr[h_i]),
                int(m_arr[h_i]),
                int(d_arr[h_i]),
                int(h_arr[h_i]),
                str(dt_strs[h_i])
            ] + [round(float(v), 2) for v in hourly_matrix[h_i]]
            ws_hr.write_row(hr_row_idx, 0, row_vals)
            hr_row_idx += 1

    print(f"     Written {hr_row_idx - 1:,} hourly rows.")

    # =========================================================================
    # SHEET 2: Daily_Summary (1940 - 2026 Daily Mean, Min, Max)
    # =========================================================================
    print("  -> Writing Sheet 2: Daily_Summary (31,600+ daily summary rows)...")
    ws_ds = wb.add_worksheet("Daily_Summary")
    ws_ds.freeze_panes(1, 4)  # Freeze row 1 and columns A-D (E2)

    ws_ds.set_column(0, 2, 8, center_int_fmt)   # Year, Month, Day
    ws_ds.set_column(3, 3, 14, center_str_fmt)  # Date

    # Build Daily Summary Headers
    daily_headers = ["Year", "Month", "Day", "Date"]
    for cname in ALL_SERIES_COLUMNS:
        # Create clear clean column labels
        clean_name = cname.replace(" (City Average)", "_City").replace(" (City)", "_City")
        clean_name = clean_name.replace(" - ", "_").replace(" ", "_").replace("(", "").replace(")", "").replace(".", "")
        daily_headers.extend([f"{clean_name}_Mean_C", f"{clean_name}_Min_C", f"{clean_name}_Max_C"])

    ws_ds.set_column(4, len(daily_headers) - 1, 14, num_fmt)
    ws_ds.write_row(0, 0, daily_headers, header_fmt)

    ds_row_idx = 1
    for yr in range(start_year, end_year + 1):
        for rec in year_data_dict[yr]["daily_summary"]:
            clean_rec = [rec[0], rec[1], rec[2], rec[3]] + [round(float(v), 2) for v in rec[4:]]
            ws_ds.write_row(ds_row_idx, 0, clean_rec)
            ds_row_idx += 1

    print(f"     Written {ds_row_idx - 1:,} daily summary rows.")

    # =========================================================================
    # SHEET 3: Station_Directory (Official PMD Metadata Table)
    # =========================================================================
    print("  -> Writing Sheet 3: Station_Directory...")
    ws_dir = wb.add_worksheet("Station_Directory")
    ws_dir.freeze_panes(1, 0)

    dir_headers = [
        "S. No.",
        "Index No.",
        "Station Name",
        "City Name",
        "Classification",
        "Latitude (DMS)",
        "Longitude (DMS)",
        "Latitude (Decimal)",
        "Longitude (Decimal)",
        "Elevation (a.m.s.l.)",
        "Telephone No.",
        "Spatial Interpolation Method",
        "ERA5 Surrounding Grid Points & Weights"
    ]

    ws_dir.set_column(0, 0, 8)   # S. No.
    ws_dir.set_column(1, 1, 12)  # Index No.
    ws_dir.set_column(2, 2, 24)  # Station Name
    ws_dir.set_column(3, 3, 18)  # City Name
    ws_dir.set_column(4, 4, 30)  # Classification
    ws_dir.set_column(5, 6, 15)  # DMS
    ws_dir.set_column(7, 8, 18)  # Decimal
    ws_dir.set_column(9, 9, 20)  # Elevation
    ws_dir.set_column(10, 10, 30) # Telephone
    ws_dir.set_column(11, 11, 28) # Method
    ws_dir.set_column(12, 12, 50) # Surrounding grid weights

    ws_dir.write_row(0, 0, dir_headers, header_fmt)

    dir_row_idx = 1

    # First, list the 4 Consolidated Full City Entities
    consolidated_cities = [
        ("C1", "CONS-LHR", "Lahore (Full City Average)", "Lahore", "Consolidated City Average (3 Stations)", "31° 34' N", "74° 23' E", 31.5722, 74.3778, "215.4 m (avg)", "042-9240402", "Average of Lahore A.P., Lahore PBO., MCC Lahore", "Bilinear Interpolation across 3 observatory coordinates"),
        ("C2", "CONS-BWP", "Bahawalpur (Full City Average)", "Bahawalpur", "Consolidated City Average (2 Stations)", "29° 21' N", "71° 44' E", 29.3583, 71.7333, "114.5 m (avg)", "0344-3668731", "Average of Bahawal Pur City & Bahawal Pur Airport", "Bilinear Interpolation across 2 observatory coordinates"),
        ("C3", "CONS-SKT", "Sialkot (Full City Average)", "Sialkot", "Consolidated City Average (2 Stations)", "32° 31' N", "74° 17' E", 32.5250, 74.2833, "247.5 m (avg)", "052-4267195", "Average of Sialkot City & Sialkot Airport", "Bilinear Interpolation across 2 observatory coordinates"),
        ("C4", "CONS-DGK", "D.G. Khan (Full City Average)", "D.G. Khan", "Consolidated City Average (2 Stations)", "30° 00' N", "70° 29' E", 30.0083, 70.4833, "144.5 m (avg)", "0332-0799908", "Average of D.G. Khan City & D.G. Khan Aeromet", "Bilinear Interpolation across 2 observatory coordinates"),
    ]

    for item in consolidated_cities:
        ws_dir.write(dir_row_idx, 0, item[0], center_str_fmt)
        ws_dir.write(dir_row_idx, 1, item[1], center_str_fmt)
        ws_dir.write(dir_row_idx, 2, item[2], left_str_fmt)
        ws_dir.write(dir_row_idx, 3, item[3], left_str_fmt)
        ws_dir.write(dir_row_idx, 4, item[4], left_str_fmt)
        ws_dir.write(dir_row_idx, 5, item[5], center_str_fmt)
        ws_dir.write(dir_row_idx, 6, item[6], center_str_fmt)
        ws_dir.write(dir_row_idx, 7, item[7], num_fmt)
        ws_dir.write(dir_row_idx, 8, item[8], num_fmt)
        ws_dir.write(dir_row_idx, 9, item[9], center_str_fmt)
        ws_dir.write(dir_row_idx, 10, item[10], left_str_fmt)
        ws_dir.write(dir_row_idx, 11, item[11], left_str_fmt)
        ws_dir.write(dir_row_idx, 12, item[12], left_str_fmt)
        dir_row_idx += 1

    # Individual Stations
    for item in STATIONS_DEF:
        sno, idx_no, sname, city, lat_dms, lon_dms, lat_dec, lon_dec, elev, phone, is_multi = item
        classification = "Full City Station" if not is_multi else f"Observatory / Sub-Station of {city}"
        
        # Grid weights string
        pts = station_lookup[sname]
        pts_str = ", ".join([f"{pt[0][0]:.2f}N_{pt[0][1]:.2f}E (w={pt[1]:.3f})" for pt in pts])

        ws_dir.write(dir_row_idx, 0, sno, center_str_fmt)
        ws_dir.write(dir_row_idx, 1, idx_no, center_str_fmt)
        ws_dir.write(dir_row_idx, 2, sname, left_str_fmt)
        ws_dir.write(dir_row_idx, 3, city, left_str_fmt)
        ws_dir.write(dir_row_idx, 4, classification, left_str_fmt)
        ws_dir.write(dir_row_idx, 5, lat_dms, center_str_fmt)
        ws_dir.write(dir_row_idx, 6, lon_dms, center_str_fmt)
        ws_dir.write(dir_row_idx, 7, round(lat_dec, 4), num_fmt)
        ws_dir.write(dir_row_idx, 8, round(lon_dec, 4), num_fmt)
        ws_dir.write(dir_row_idx, 9, elev, center_str_fmt)
        ws_dir.write(dir_row_idx, 10, phone, left_str_fmt)
        ws_dir.write(dir_row_idx, 11, "Bilinear 4-Point Weighted Interpolation", left_str_fmt)
        ws_dir.write(dir_row_idx, 12, pts_str, left_str_fmt)
        dir_row_idx += 1

    print("     Written Station_Directory table.")

    wb.close()
    elapsed = time.time() - t0
    fsize_mb = os.path.getsize(output_path) / (1024 * 1024)
    print(f"[SUCCESS] Workbook saved: {output_path} ({fsize_mb:.1f} MB in {elapsed:.1f}s)")
    return output_path, fsize_mb


def main():
    parser = argparse.ArgumentParser(description="Transform ERA5 702-point Matrix into Meteorological Stations & Cities (1940-2026).")
    parser.add_argument("--start_year", type=int, default=1940, help="First year to process (default: 1940)")
    parser.add_argument("--end_year", type=int, default=2026, help="Last year to process (default: 2026)")
    parser.add_argument("--input_dir", type=str, default="excel_output", help="Folder with yearly ERA5 Excel files")
    parser.add_argument("--output_dir", type=str, default=".", help="Output directory for transformed Excel file")
    parser.add_argument("--output_file", type=str, default="ERA5_Stations_Transformed_1940_2026.xlsx", help="Filename for master transformed workbook")
    parser.add_argument("--cache_path", type=str, default="compiled_cache/stations_transformed_cache_1940_2026.pkl", help="Path to pickle cache")
    parser.add_argument("--max_workers", type=int, default=8, help="Number of parallel extraction processes")
    parser.add_argument("--force_extract", action="store_true", help="Force re-extraction from yearly files even if cache exists")
    parser.add_argument("--decades", action="store_true", help="Also generate 10-year decade transformed workbooks")
    args = parser.parse_args()

    overall_t0 = time.time()
    print("=" * 80)
    print(" ERA5 METEOROLOGICAL STATIONS & FULL-CITY TRANSFORMATION PIPELINE")
    print(f" Years             : {args.start_year} - {args.end_year}")
    print(f" Input Directory   : {os.path.abspath(args.input_dir)}")
    print(f" Output Master File: {os.path.abspath(os.path.join(args.output_dir, args.output_file))}")
    print(f" Parallel Workers  : {args.max_workers}")
    print("=" * 80)

    # Step 1: Extract all stations data
    year_data_dict, station_lookup, coord_to_idx = load_or_extract_all_stations(
        input_dir=args.input_dir,
        start_year=args.start_year,
        end_year=args.end_year,
        cache_path=args.cache_path,
        max_workers=args.max_workers,
        force_extract=args.force_extract
    )

    # Step 2: Write Master Full Workbook (1940 - 2026)
    master_path = os.path.join(args.output_dir, args.output_file)
    write_stations_workbook(
        output_path=master_path,
        start_year=args.start_year,
        end_year=args.end_year,
        year_data_dict=year_data_dict,
        station_lookup=station_lookup,
        coord_to_idx=coord_to_idx,
        title_label="Master Unified Stations Dataset (1940-2026)"
    )

    # Step 3: Optional Decade Workbooks
    if args.decades:
        print("\n" + "=" * 80)
        print(" GENERATING 10-YEAR DECADE STATION WORKBOOKS (1940 - 2026)")
        print("=" * 80)
        for s_yr, e_yr in DECADE_RANGES:
            dec_fname = f"ERA5_Stations_Transformed_{s_yr}_{e_yr}.xlsx"
            dec_path = os.path.join(args.output_dir, dec_fname)
            write_stations_workbook(
                output_path=dec_path,
                start_year=s_yr,
                end_year=e_yr,
                year_data_dict=year_data_dict,
                station_lookup=station_lookup,
                coord_to_idx=coord_to_idx,
                title_label=f"Decade Stations Dataset ({s_yr}-{e_yr})"
            )

    total_time = time.time() - overall_t0
    print("\n" + "=" * 80)
    print(f" [COMPLETE] All operations finished in {total_time:.1f}s ({total_time/60:.2f} minutes)!")
    print(f" Master file ready: {os.path.abspath(master_path)}")
    print("=" * 80)


if __name__ == "__main__":
    main()
