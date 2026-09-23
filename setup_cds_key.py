"""
Setup CDS API Key for Copernicus Climate Data Store
"""
import os
import sys

def setup_key():
    print("=" * 60)
    print(" Copernicus Climate Data Store (CDS) API Setup ")
    print("=" * 60)
    print("To download ERA5 data, you need an account on CDS:")
    print("1. Register or login at: https://cds.climate.copernicus.eu")
    print("2. Accept dataset Terms of Use for 'ERA5 hourly data on single levels':")
    print("   https://cds.climate.copernicus.eu/datasets/reanalysis-era5-single-levels")
    print("3. Go to your Profile page to copy your Personal Access Token:")
    print("   https://cds.climate.copernicus.eu/profile")
    print("-" * 60)
    
    cdsapirc_path = os.path.expanduser("~/.cdsapirc")
    if os.path.exists(cdsapirc_path):
        print(f"Existing configuration found at: {cdsapirc_path}")
        with open(cdsapirc_path, "r") as f:
            print("Current contents:")
            print(f.read())
        resp = input("Do you want to overwrite it? (y/n): ").strip().lower()
        if resp != "y":
            print("Aborted.")
            return

    if len(sys.argv) > 1:
        token = sys.argv[1].strip()
    else:
        token = input("Enter your CDS Personal Access Token: ").strip()
    if not token:
        print("Error: Token cannot be empty.")
        return

    content = f"url: https://cds.climate.copernicus.eu/api\nkey: {token}\n"
    with open(cdsapirc_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"\n[SUCCESS] Credentials saved to: {cdsapirc_path}")
    print("You can now run 'python era5_downloader_and_converter.py'.")

if __name__ == "__main__":
    setup_key()
