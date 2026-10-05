import os
import re
import glob
import calendar
import zipfile
import xml.etree.ElementTree as ET

def verify_all(excel_dir="excel_output", start_year=1960, end_year=None):
    if end_year is None:
        files = glob.glob(os.path.join(excel_dir, "ERA5_2m_Temperature_*.xlsx"))
        found_years = []
        for f in files:
            m = re.search(r"_(\d{4})\.xlsx", os.path.basename(f))
            if m:
                found_years.append(int(m.group(1)))
        end_year = max(found_years) if found_years else 2016

    report = []
    print(f"Starting verification of Excel files from {start_year} to {end_year} in '{excel_dir}'...\n")

    for year in range(start_year, end_year + 1):
        fpath = os.path.join(excel_dir, f"ERA5_2m_Temperature_{year}.xlsx")
        if not os.path.exists(fpath):
            report.append((year, "MISSING", f"File does not exist: {fpath}"))
            continue

        size_mb = os.path.getsize(fpath) / (1024 * 1024)
        is_leap = calendar.isleap(year)
        expected_hours = 8784 if is_leap else 8760
        expected_days = 366 if is_leap else 365
        expected_h_dim = f"A1:AAA{expected_hours + 1}"
        expected_d_dim = f"A1:D{expected_days + 1}"
        expected_c_dim = "A1:D703"

        try:
            with zipfile.ZipFile(fpath, "r") as zf:
                bad_crc = zf.testzip()
                if bad_crc:
                    report.append((year, "CORRUPT_CRC", f"CRC check failed on {bad_crc}"))
                    continue

                # Check sheets
                with zf.open("xl/workbook.xml") as wb_file:
                    root = ET.fromstring(wb_file.read())
                    sheets = [
                        c.attrib["name"]
                        for c in root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets")
                    ]

                expected_sheets = ["Daily_Summary", "Grid_Coordinates", "Hourly_Grid_Matrix"]
                if sheets != expected_sheets:
                    report.append((year, "SHEET_MISMATCH", f"Found sheets: {sheets}"))
                    continue

                # Check dimensions
                dims = {}
                for i, sname in enumerate(sheets, 1):
                    with zf.open(f"xl/worksheets/sheet{i}.xml") as s_file:
                        chunk = s_file.read(1500)
                        m = re.search(rb'<dimension ref="([^"]+)"', chunk)
                        dims[sname] = m.group(1).decode() if m else "NOT_FOUND"

                if dims["Daily_Summary"] != expected_d_dim:
                    report.append(
                        (year, "DIM_MISMATCH", f"Daily_Summary dim {dims['Daily_Summary']} != {expected_d_dim}")
                    )
                    continue
                if dims["Grid_Coordinates"] != expected_c_dim:
                    report.append(
                        (year, "DIM_MISMATCH", f"Grid_Coordinates dim {dims['Grid_Coordinates']} != {expected_c_dim}")
                    )
                    continue
                if dims["Hourly_Grid_Matrix"] != expected_h_dim:
                    report.append(
                        (year, "DIM_MISMATCH", f"Hourly_Grid_Matrix dim {dims['Hourly_Grid_Matrix']} != {expected_h_dim}")
                    )
                    continue

                report.append((year, "OK", f"{size_mb:.1f} MB, leap={is_leap}, {expected_hours} hrs, CRC verified"))
        except Exception as e:
            report.append((year, "CORRUPT_FILE", str(e)))

    print("=" * 70)
    print("VERIFICATION SUMMARY:")
    print("=" * 70)
    failed = [r for r in report if r[1] != "OK"]
    passed = [r for r in report if r[1] == "OK"]

    for yr, status, detail in report:
        sym = "[OK]  " if status == "OK" else "[FAIL]"
        print(f"{sym} {yr}: {status:<15} | {detail}")

    print("=" * 70)
    print(f"Total Passed: {len(passed)} / {len(report)}")
    print(f"Total Failed / Issues: {len(failed)} / {len(report)}")
    print("=" * 70)
    return passed, failed

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Verify ERA5 Excel files integrity")
    parser.add_argument("--start_year", type=int, default=1960, help="First year to verify (default: 1960)")
    parser.add_argument("--end_year", type=int, default=None, help="Last year to verify (default: all available)")
    parser.add_argument("--excel_dir", type=str, default="excel_output", help="Directory containing Excel files")
    args = parser.parse_args()

    verify_all(excel_dir=args.excel_dir, start_year=args.start_year, end_year=args.end_year)
