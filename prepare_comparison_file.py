# import pandas as pd
# import os
# import time
# import warnings
#
# # Suppress warnings
# warnings.filterwarnings("ignore")
#
#
# def find_header_row(file_path, sheet_name, engine):
#     """
#     Scans the first 50 rows to find the row index where the actual table starts.
#     It looks for a row containing 'Code', 'CPT Code', or 'Service Code'.
#     """
#     try:
#         # Read only the first 50 rows, no header initially
#         df_preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=50, dtype=str, engine=engine)
#     except Exception:
#         return 0
#
#     target_headers = ["code", "cpt code", "procedure code", "hcpcs", "service code"]
#
#     for idx, row in df_preview.iterrows():
#         # Convert row values to lowercase strings
#         row_values = [str(x).strip().lower() for x in row.values]
#
#         # Check if any target header exists in this row
#         if any(t in row_values for t in target_headers):
#             return idx
#
#     return 0  # Fallback to first row if not found
#
#
# def merge_quarters_clean(file_old, file_new, output_filename):
#     print("=" * 60)
#     print("       🧹 SMART CLEANING & MERGING TOOL")
#     print("=" * 60)
#
#     # 1. Detect Fast Engine
#     read_engine = 'openpyxl'
#     try:
#         import python_calamine
#         read_engine = 'calamine'
#         print("✅ Fast 'calamine' engine detected!")
#     except ImportError:
#         print("⚠️ Using 'openpyxl' (Slower). Install 'python-calamine' for speed.")
#
#     # 2. Get Sheet Names
#     print(f"\n📂 Scanning files...")
#     try:
#         xls_old = pd.ExcelFile(file_old, engine=read_engine)
#         xls_new = pd.ExcelFile(file_new, engine=read_engine)
#     except Exception as e:
#         print(f"❌ Error reading files: {e}")
#         return
#
#     # 3. Identify Common Sheets
#     ignore_sheets = [
#         "Updates", "UPDATES", "UPDATES Temp", "Sheet1",
#         "Lookup Tool", "Reference", "Change Log", "Instructions",
#         "Evolent Delegated Codes", "ICD 10 Codes", "MEDICAID Evolent Archive",
#         "Introduction", "Table of Contents"
#     ]
#
#     sheets_old = set(xls_old.sheet_names)
#     sheets_new = set(xls_new.sheet_names)
#     common_sheets = sorted([s for s in sheets_old if s in sheets_new and s not in ignore_sheets])
#
#     print(f"✅ Found {len(common_sheets)} states to merge.")
#
#     # 4. Processing & Cleaning
#     config_list = []
#
#     print(f"\n💾 Cleaning and writing to {output_filename}...")
#
#     with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
#         for i, sheet in enumerate(common_sheets, 1):
#             print(f"   [{i}/{len(common_sheets)}] {sheet:<5}", end=" ", flush=True)
#
#             try:
#                 # A. Find the real header row for both files
#                 header_idx_old = find_header_row(file_old, sheet, read_engine)
#                 header_idx_new = find_header_row(file_new, sheet, read_engine)
#
#                 # B. Read Data starting from that row
#                 # header=header_idx tells pandas to use that row as columns and skip everything above
#                 df_old = pd.read_excel(file_old, sheet_name=sheet, header=header_idx_old, dtype=str, engine=read_engine)
#                 df_new = pd.read_excel(file_new, sheet_name=sheet, header=header_idx_new, dtype=str, engine=read_engine)
#
#                 # C. Basic cleanup (drop empty columns/rows)
#                 df_old = df_old.dropna(how='all', axis=1).dropna(how='all', axis=0)
#                 df_new = df_new.dropna(how='all', axis=1).dropna(how='all', axis=0)
#
#                 rows_count = len(df_old) + len(df_new)
#                 print(f"| Cleaned Rows: {rows_count:<6} | Writing...", end=" ", flush=True)
#
#                 sheet_name_old = f"{sheet} 25Q4"
#                 sheet_name_new = f"{sheet} 26Q1"
#
#                 # D. Write Clean Data
#                 df_old.to_excel(writer, sheet_name=sheet_name_old, index=False)
#                 df_new.to_excel(writer, sheet_name=sheet_name_new, index=False)
#
#                 print("Done. ✅")
#
#                 # E. Build Config for Step 2
#                 config_entry = {
#                     "q3_sheet": sheet_name_old,
#                     "q4_sheet": sheet_name_new,
#                     "output_sheet": f"{sheet} 25Q4 vs 26Q1",
#                     "notes_col_candidates": ["Code Notes", "Notes", "Additional Notes", "Comments"],
#                     "medicaid_col_candidates": ["Medicaid", "MHI Medicaid", "Medicaid PA", "MHI Medicaid PA Status", "MediCal"]
#                 }
#                 config_list.append(config_entry)
#
#             except Exception as e:
#                 print(f"\n❌ Error processing {sheet}: {e}")
#
#     print(f"\n🎉 Clean Merge Complete: {output_filename}")
#
#     # Output Config for main script
#     print("\n" + "=" * 50)
#     print("COPY THIS CONFIG INTO 'run_step2_analysis.py':")
#     print("=" * 50)
#     print("my_comparisons = [")
#     for item in config_list:
#         print(f"    {item},")
#     print("]")
#     print("=" * 50)
#
#
# if __name__ == "__main__":
#     file_old = "Authorization Business Matrix 2025 Q4 - All States and LOBs - Reference.xlsx"
#     file_new = "Authorization Business Matrix 2026 Q1 - All States and LOBs - Reference.xlsx"
#     output_file = "Merged_25Q4_26Q1_Analysis_1.xlsx"
#
#     if os.path.exists(file_old) and os.path.exists(file_new):
#         merge_quarters_clean(file_old, file_new, output_file)
#     else:
#         print("❌ Input files not found. Check filenames.")

# try:
#     import python_calamine
#     READ_ENGINE = 'calamine'
#     print("✅ Fast 'calamine' engine detected!")
# except ImportError:
#     READ_ENGINE = 'openpyxl'
#     print("⚠️ Using 'openpyxl' (Slower). Install 'python-calamine' for speed.")
import pandas as pd
import os

try:
    import python_calamine
    READ_ENGINE = 'calamine'
    print("✅ Fast 'calamine' engine detected!")
except ImportError:
    READ_ENGINE = 'openpyxl'
    print("⚠️ Using 'openpyxl' (Slower). Install 'python-calamine' for speed.")


    print("⚠️ Using 'openpyxl' (Slower). Install 'python-calamine' for speed.")


def find_header_row(file_path, sheet_name, engine):
    """
    Scans the first 50 rows to find the row index where the actual table starts.
    It looks for a row containing 'Code', 'CPT Code', or 'Service Code'.
    """
    try:
        # Read only the first 50 rows, no header initially
        # Use openpyxl here if calamine is problematic for partial reads, but calamine is fast enough to read all?
        # Let's stick to the main engine passed in.
        # But wait, pd.read_excel with nrows might behave differently with calamine in older pandas versions?
        # Assuming typical pandas behavior:
        df_preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=50, dtype=str, engine=engine)
    except Exception:
        return 0

    target_headers = ["code", "cpt code", "procedure code", "hcpcs", "service code"]

    for idx, row in df_preview.iterrows():
        # Convert row values to lowercase strings
        row_values = [str(x).strip().lower() for x in row.values]

        # Check if any target header exists in this row
        if any(t in row_values for t in target_headers):
            return idx

    return 0  # Fallback to first row if not found


def merge_quarters_and_generate_config(file_old, file_new, output_filename):
    """
    Merges matching sheets from two Excel files into one and generates
    the config list for the analysis script.
    """

    print(f"Loading 2025 Q4 (Old): {file_old}...")
    try:
        xls_old = pd.ExcelFile(file_old, engine=READ_ENGINE)
    except FileNotFoundError:
        print(f"Error: File not found - {file_old}")
        return

    print(f"Loading 2026 Q1 (New): {file_new}...")
    try:
        xls_new = pd.ExcelFile(file_new, engine=READ_ENGINE)
    except FileNotFoundError:
        print(f"Error: File not found - {file_new}")
        return

    # 1. Define Explicit Sheet Mappings
    # (Old Sheet Name, New Sheet Name, Output Label)
    target_mappings = [
        ("MEDICAID", "MEDICAID", "MEDICAID"),
        ("WA", "WA", "WA"),
        ("NY", "NY", "NY"),
        ("KY", "KY", "KY"),
        ("SWH MA", "MA", "MA"),  # Handled mapping
        ("MS", "MS", "MS"),
        ("ID", "ID", "ID")
    ]
    
    # Filter to only those that exist
    sheets_old_set = set(xls_old.sheet_names)
    sheets_new_set = set(xls_new.sheet_names)
    
    to_process = []
    for s_old, s_new, label in target_mappings:
        if s_old in sheets_old_set and s_new in sheets_new_set:
            to_process.append((s_old, s_new, label))
        else:
            print(f"⚠️ Warning: skipping {label} - {s_old} in Old: {s_old in sheets_old_set}, {s_new} in New: {s_new in sheets_new_set}")

    print(f"\nFound {len(to_process)} valid sheet pairs to process: {[x[2] for x in to_process]}")
    
    if not to_process:
        print("No matching sheets found.")
        return

    # 2. Merge and Write to New File
    config_list = []

    print(f"\nWriting merged data to {output_filename}...")
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        for s_old, s_new, label in to_process:
            print(f"  Processing: {label} ({s_old} -> {s_new})", flush=True)

            # Define new sheet names
            # Using specific quarters in names to be clear
            sheet_name_old = f"{label} 25Q4"
            sheet_name_new = f"{label} 26Q1"

            # Read from source files
            
            # A. Find the real header row for both files
            header_idx_old = find_header_row(file_old, s_old, READ_ENGINE)
            header_idx_new = find_header_row(file_new, s_new, READ_ENGINE)
            
            print(f"    Headers found at rows: Old={header_idx_old}, New={header_idx_new}")

            df_old = pd.read_excel(xls_old, sheet_name=s_old, header=header_idx_old)
            df_new = pd.read_excel(xls_new, sheet_name=s_new, header=header_idx_new)
            
            # Clean up empty columns and rows which cause massive performance hits if present (e.g. 16k empty cols)
            # Normalize empty strings and whitespace to NaN
            df_old = df_old.replace(r'^\s*$', pd.NA, regex=True)
            df_new = df_new.replace(r'^\s*$', pd.NA, regex=True)
            
            df_old = df_old.dropna(how='all', axis=1).dropna(how='all', axis=0)
            df_new = df_new.dropna(how='all', axis=1).dropna(how='all', axis=0)
            
            print(f"    Cleaned Rows: {len(df_old) + len(df_new)}", end=" | ", flush=True)

            # Write to the single merged output file
            df_old.to_excel(writer, sheet_name=sheet_name_old, index=False)
            df_new.to_excel(writer, sheet_name=sheet_name_new, index=False)
            print("Done.")

            # 3. Build Config Entry
            # We map the new specific names to the keys 'q3_sheet' and 'q4_sheet'
            # expected by your analysis script, even though they represent Q4 and Q1.
            config_entry = {
                "q3_sheet": sheet_name_old,  # Acts as the "Old" comparison point
                "q4_sheet": sheet_name_new,  # Acts as the "New" comparison point
                "output_sheet": f"{label} 25Q4 vs 26Q1",
                # Including candidates seen in your uploaded files
                "notes_col_candidates": ["Code Notes", "Notes", "Additional Notes"],
                "medicaid_col_candidates": ["Medicaid", "MHI Medicaid", "Medicaid PA", "MHI Medicaid PA Status"]
            }
            config_list.append(config_entry)

    print(f"\n✅ Success! Merged file saved as: {output_filename}")

    # 4. Output the Python List
    print("\n" + "=" * 50)
    print("COPY THIS LIST INTO YOUR new_code_analysis.py:")
    print("=" * 50)
    print("my_comparisons = [")
    for item in config_list:
        print(f"    {item},")
    print("]")
    print("=" * 50)


if __name__ == "__main__":
    # --- UPDATE THESE PATHS TO YOUR ACTUAL FILE LOCATIONS ---
    path_2025_q4 = "Authorization Business Matrix 2025 Q4 - All States and LOBs - Reference.xlsx"
    path_2026_q1 = "Authorization Business Matrix 2026 Q1 - All States and LOBs - Reference.xlsx"

    output_file = "Merged_25Q4_26Q1_Analysis_3.xlsx"

    merge_quarters_and_generate_config(path_2025_q4, path_2026_q1, output_file)