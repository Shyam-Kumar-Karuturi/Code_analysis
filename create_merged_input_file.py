import pandas as pd
import os


def prepare_comparison_file(file_old, file_new, output_filename):
    """
    Reads two Authorization Matrix files, finds matching state sheets,
    and merges them into a single workbook suitable for the analysis script.
    """
    print(f"📂 Loading Old File (2025 Q4): {file_old}")
    try:
        xls_old = pd.ExcelFile(file_old)
    except FileNotFoundError:
        print(f"❌ Error: Could not find file {file_old}")
        return

    print(f"📂 Loading New File (2026 Q1): {file_new}")
    try:
        xls_new = pd.ExcelFile(file_new)
    except FileNotFoundError:
        print(f"❌ Error: Could not find file {file_new}")
        return

    # Filter out non-state sheets (Summary/Index tabs)
    ignored_sheets = [
        "Sheet1", "UPDATES", "Updates", "UPDATES Temp", "Lookup Tool",
        "ICD 10 Codes", "Evolent Delegated Codes", "MEDICARE", "MEDICAID",
        "MARKETPLACE", "MEDICAID Evolent Archive", "Reference",
        "Change Log", "Instructions"
    ]

    # Find sheets that exist in BOTH files
    sheets_old = set(xls_old.sheet_names)
    sheets_new = set(xls_new.sheet_names)
    common_sheets = sorted([s for s in sheets_old if s in sheets_new and s not in ignored_sheets])

    if not common_sheets:
        print("⚠️ No matching state sheets found between the two files.")
        return

    print(f"✅ Found {len(common_sheets)} common states to process: {common_sheets}")

    config_output = []

    # Create the merged Excel file
    print(f"\n📝 Writing merged data to {output_filename}...")
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        for sheet in common_sheets:
            print(f"   Processing State: {sheet}")

            # Read Data
            df_old = pd.read_excel(xls_old, sheet_name=sheet)
            df_new = pd.read_excel(xls_new, sheet_name=sheet)

            # Define specific sheet names for the output file
            # Naming convention: "State 25Q4" and "State 26Q1"
            name_old = f"{sheet} 25Q4"
            name_new = f"{sheet} 26Q1"

            # Write to the new workbook
            df_old.to_excel(writer, sheet_name=name_old, index=False)
            df_new.to_excel(writer, sheet_name=name_new, index=False)

            # Generate the config dictionary for your analysis script
            config_entry = {
                "q3_sheet": name_old,  # The 'Old' Quarter
                "q4_sheet": name_new,  # The 'New' Quarter
                "output_sheet": f"{sheet} Comparison",
                "notes_col_candidates": ["Code Notes", "Notes", "Additional Notes"],
                "medicaid_col_candidates": ["Medicaid", "MHI Medicaid", "Medicaid PA"]
            }
            config_output.append(config_entry)

    print(f"\n🎉 File created successfully: {output_filename}")

    # Print the config list for the user to copy
    print("\n" + "=" * 60)
    print("📋 COPY THIS LIST INTO YOUR 'new_code_analysis.py' SCRIPT:")
    print("=" * 60)
    print("my_comparisons = [")
    for item in config_output:
        print(f"    {item},")
    print("]")
    print("=" * 60)


if __name__ == "__main__":
    # Update these filenames if yours are different
    file_2025 = "Authorization Business Matrix 2025 Q4 - All States and LOBs - Reference.xlsx"
    file_2026 = "Authorization Business Matrix 2026 Q1 - All States and LOBs - Reference.xlsx"

    output_name = "Merged_25Q4_26Q1_Input.xlsx"

    prepare_comparison_file(file_2025, file_2026, output_name)