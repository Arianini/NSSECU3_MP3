import os
import subprocess
import json
import pandas as pd
from datetime import datetime, timedelta, timezone
from dateutil import parser, tz
import re
import pytz

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

TOOLS_DIR = os.path.join(SCRIPT_DIR, "tools")
PHOTOREC_DIR = os.path.join(TOOLS_DIR, "testdisk-7.3-WIP")
EXIFTOOL_DIR = os.path.join(SCRIPT_DIR, "tools", "exiftool-13.19_64")
AMCACHE_DIR = os.path.join(TOOLS_DIR, "AmcacheParser")

def find_executable(tool_folder, tool_name):
    for file in os.listdir(tool_folder):
        if tool_name.lower() in file.lower() and file.endswith(".exe"):
            return os.path.join(tool_folder, file)
    return None

PHOTOREC_PATH = find_executable(PHOTOREC_DIR, "photorec")
EXIFTOOL_PATH = find_executable(EXIFTOOL_DIR, "exiftool")
AMCACHE_PARSER_PATH = find_executable(AMCACHE_DIR, "AmcacheParser")

RECOVERED_ROOT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, "RecoveredFiles"))
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
RECOVERED_DIR = os.path.join(RECOVERED_ROOT_DIR, f"Recovery_{TIMESTAMP}")

BASE_DIR = os.path.join(SCRIPT_DIR, f"ForensicSession_{TIMESTAMP}")
AMCACHE_OUTPUT_DIR = os.path.join(BASE_DIR, "AmcacheAnalysis")

for directory in [BASE_DIR, AMCACHE_OUTPUT_DIR, RECOVERED_DIR]:
    os.makedirs(directory, exist_ok=True)

EXIF_OUTPUT_FILE = os.path.join(BASE_DIR, "metadata.json")
FINAL_CSV = os.path.join(BASE_DIR, "consolidated_timeline.csv")

DISK_PATH = "E:\\"  # Set your target disk path here
FILE_TYPES = ["jpg", "mp4"]

def enable_file_types():
    print(f"\n🔍 [PhotoRec] Enabling file types: {FILE_TYPES}...")
    disable_cmd = [PHOTOREC_PATH, "/cmd", DISK_PATH, "fileopt", "disable"]
    subprocess.run(disable_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    for ftype in FILE_TYPES:
        enable_cmd = [PHOTOREC_PATH, "/cmd", DISK_PATH, "fileopt", "enable", ftype]
        subprocess.run(enable_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    print("\n✅ [PhotoRec] File type filtering applied successfully!")

def run_photorec():
    if not PHOTOREC_PATH:
        print("\n[!] PhotoRec executable not found! Skipping recovery...")
        return False
    print("\n🔍 [PhotoRec] Starting controlled file recovery...")
    enable_file_types() 
    photorec_cmd = [PHOTOREC_PATH, "/d", RECOVERED_DIR, "/cmd", DISK_PATH, "search"]
    result = subprocess.run(photorec_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

    print("\n✅ [PhotoRec] Recovery process completed successfully!")
    return True

def run_exiftool(exiftool_path, root_dir, output_file):
    if not exiftool_path or not os.path.exists(root_dir):
        print("\n[!] ExifTool skipped: Root directory not found.")
        return []
    scan_dirs = []
    for dirpath, _, filenames in os.walk(root_dir):
        if any(f.lower().endswith((".jpg", ".pdf", ".png", ".doc", ".docx")) for f in filenames):
            scan_dirs.append(dirpath)
    if not scan_dirs:
        print("\n[!] ExifTool skipped: No relevant files found in recovered files or subdirectories.")
        return []
    print("\n🔍 [ExifTool] Extracting metadata from recovered files...")
    exif_cmd = [exiftool_path, "-r", "-json", "-ext", "jpg", "-ext", "pdf", "-ext", "png", "-ext", "doc", "-ext", "docx"] + scan_dirs 
    result = subprocess.run(exif_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace")
    if result.stdout is None or not result.stdout.strip():
        print("\n❌ [ExifTool] Encountered an error:", result.stderr)
        return []
    try:
        metadata = json.loads(result.stdout)
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=4)
        print(f"\n✅ [ExifTool] Metadata saved at: {output_file}")
        return metadata
    except json.JSONDecodeError:
        print("\n❌ [ExifTool] Error parsing metadata JSON.")
        return []

def run_amcache_parser(amcache_parser_path, output_dir):
    if not amcache_parser_path:
        print("\n[!] AmcacheParser not found! Skipping execution history extraction...")
        return {}
    print("\n🔍 [AmcacheParser] Extracting execution history...")
    existing_csvs = set(f for f in os.listdir(output_dir) if f.endswith(".csv"))
    amcache_hve = r"C:\Windows\AppCompat\Programs\Amcache.hve"
    if not os.path.exists(amcache_hve):
        print("\n❌ [AmcacheParser] Amcache.hve not found! Skipping analysis...")
        return {}
    amcache_cmd = [amcache_parser_path, "-f", amcache_hve, "--csv", output_dir]
    result = subprocess.run(amcache_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        print("\n❌ [AmcacheParser] Failed:\n", result.stderr)
        return {}
    new_csvs = set(f for f in os.listdir(output_dir) if f.endswith(".csv")) - existing_csvs
    file_count = len(new_csvs)
    if file_count == 0:
        print("\n❌ [AmcacheParser] No new data extracted.")
        return {}
    print(f"\n✅ [AmcacheParser] {file_count} CSV files generated. Output directory: {output_dir}")
    return {csv_file: os.path.join(output_dir, csv_file) for csv_file in new_csvs}

def parse_timestamp(ts_str):
    if not ts_str or str(ts_str).strip() == "" or pd.isna(ts_str):
        return None
    ts_str = str(ts_str).strip()
    try:
        if re.match(r"^\d{4}:\d{2}:\d{2}\s", ts_str):
            fmt = "%Y:%m:%d %H:%M:%S%z" if re.search(r"(\+|-)\d{2}:\d{2}$", ts_str) else "%Y:%m:%d %H:%M:%S"
            dt = datetime.strptime(ts_str, fmt)
        else:
            dt = parser.parse(ts_str)
    except Exception:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=tz.tzlocal())
    return dt.astimezone(timezone.utc)

def merge_artifacts(exif_data, amcache_csvs, photorec_data):
    """Merge artifacts from ExifTool, Amcache, and PhotoRec without merging rows by timestamp."""
    # For simplicity, we will concatenate all data without merging matching timestamps.
    exif_rows = []
    for meta in exif_data:
        timestamp_fields = ["DateTimeOriginal", "CreateDate", "ModifyDate", "FileModifyDate"]
        ts_val = None
        for field in timestamp_fields:
            if field in meta and str(meta[field]).strip():
                ts_val = meta[field]
                break
        ts_utc = parse_timestamp(ts_val)
        exif_rows.append({
            "Timestamp": ts_utc,
            "Artifact_Type": "File Metadata",
            "Source": "ExifTool",
            "FileName": meta.get("FileName", ""),
            "Original_File": meta.get("SourceFile", ""),
            "FileSize": meta.get("FileSize", ""),
            "FileType": meta.get("FileType", ""),
            "FileModifyDate": meta.get("FileModifyDate", ""),
            "FileAccessDate": meta.get("FileAccessDate", ""),
            "FileCreateDate": meta.get("FileCreateDate", "")
        })
    amcache_rows = []
    if isinstance(amcache_csvs, dict):
        amcache_frames = []
        for csv_path in amcache_csvs.values():
            try:
                am_df = pd.read_csv(csv_path, encoding="utf-8")
                amcache_frames.append(am_df)
            except Exception as e:
                print(f"\n[!] Could not read Amcache CSV '{csv_path}': {e}")
        if amcache_frames:
            merged_amcache = pd.concat(amcache_frames, ignore_index=True, sort=False)
            for _, row in merged_amcache.iterrows():
                am_ts = None
                for field in ["LastModifiedTime", "LastWriteTimestamp", "LastExecutionTime", "LastRunTime"]:
                    if field in row and pd.notna(row[field]) and str(row[field]).strip():
                        am_ts = row[field]
                        break
                ts_utc = parse_timestamp(am_ts)
                program_path = row["Path"] if "Path" in row and pd.notna(row["Path"]) else ""
                program_name = os.path.basename(str(program_path)) if program_path else (row["Name"] if "Name" in row else "")
                size_val = row["Size"] if "Size" in row else (row["FileSize"] if "FileSize" in row else "")
                amcache_rows.append({
                    "Timestamp": ts_utc,
                    "Artifact_Type": "Execution History",
                    "Source": "Amcache",
                    "FileName": program_name,
                    "Original_File": program_path,
                    "FileSize": size_val if pd.notna(size_val) else "",
                    "FileType": "Executable",
                    "FileModifyDate": "",
                    "FileAccessDate": "",
                    "FileCreateDate": ""
                })
    photorec_rows = []
    for data in photorec_data:
        # photorec_data is a list of dictionaries collected from recovered files
        photorec_rows.append(data)
    return exif_rows + amcache_rows + photorec_rows

def get_photorec_data():
    """Extract data from PhotoRec recovered files."""
    rows = []
    for root, _, files in os.walk(RECOVERED_DIR):
        for file in files:
            file_path = os.path.join(root, file)
            ts_utc = None
            try:
                ts = os.path.getctime(file_path)
                ts_utc = datetime.utcfromtimestamp(ts).replace(tzinfo=timezone.utc)
            except Exception:
                ts_utc = None
            rows.append({
                "Timestamp": ts_utc,
                "Artifact_Type": "Recovered File",
                "Source": "PhotoRec",
                "FileName": file,
                "Original_File": file_path,
                "FileSize": os.path.getsize(file_path),
                "FileType": os.path.splitext(file)[1],
                "FileModifyDate": "",
                "FileAccessDate": "",
                "FileCreateDate": ""
            })
    return rows

def merge_and_generate_timeline(exif_data, amcache_csvs):
    photorec_data = get_photorec_data()
    all_rows = merge_artifacts(exif_data, amcache_csvs, photorec_data)
    df = pd.DataFrame(all_rows)
    if df.empty:
        print("\n[!] No artifacts to include in the timeline.")
        return
    df.dropna(subset=["Timestamp"], inplace=True)
    df.sort_values(by="Timestamp", inplace=True)
    df["Timestamp"] = df["Timestamp"].apply(lambda dt: dt.strftime("%Y-%m-%dT%H:%M:%SZ") if pd.notna(dt) else "")
    df.fillna("NA", inplace=True)
    columns_order = ["Timestamp", "Artifact_Type", "Source", "FileName", "Original_File", 
                     "FileSize", "FileType", "FileModifyDate", "FileAccessDate", "FileCreateDate"]
    df = df.reindex(columns=columns_order, fill_value="NA")
    df.to_csv(FINAL_CSV, index=False, encoding="utf-8")
    print(f"\n[+] Timeline successfully written to '{FINAL_CSV}'")

def main():
    print("\n🚀 Starting forensic analysis workflow...\n")
    run_photorec()
    exif_metadata = run_exiftool(EXIFTOOL_PATH, RECOVERED_ROOT_DIR, EXIF_OUTPUT_FILE)
    amcache_csvs = run_amcache_parser(AMCACHE_PARSER_PATH, AMCACHE_OUTPUT_DIR)
    merge_and_generate_timeline(exif_metadata, amcache_csvs)
    print("\n🎯 Forensic analysis complete!")

if __name__ == "__main__":
    main()
