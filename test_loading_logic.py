import os
import pandas as pd
import time
import shutil
from excel_io import get_sheet_names, read_header_file

def create_sample_files():
    print("Creating sample files...")
    # Create a dummy Excel file with multiple sheets
    df = pd.DataFrame({"A": [1, 2, 3], "B": ["x", "y", "z"]})
    df.to_excel("test_load.xlsx", sheet_name="Sheet1", index=False)
    
    with pd.ExcelWriter("test_load.xlsx", engine='openpyxl', mode='a') as writer:
        pd.DataFrame({"C": [4, 5], "D": [6, 7]}).to_excel(writer, sheet_name="Sheet2", index=False)
        
    # Create a CSV file
    df.to_csv("test_load.csv", index=False)
    print("Files created: test_load.xlsx, test_load.csv")

def test_sheet_loading():
    print("\n--- Testing Sheet Loading ---")
    start = time.time()
    sheets = get_sheet_names(os.path.abspath("test_load.xlsx"))
    duration = time.time() - start
    print(f"[XLSX] Sheets found: {sheets} (Time: {duration:.4f}s)")
    
    if "Sheet1" in sheets and "Sheet2" in sheets:
        print("PASS: Sheet names correct.")
    else:
        print("FAIL: Sheet names mismatch.")

    start = time.time()
    sheets_csv = get_sheet_names(os.path.abspath("test_load.csv"))
    duration = time.time() - start
    print(f"[CSV] Sheets found: {sheets_csv} (Time: {duration:.4f}s)")
    
    if sheets_csv == ["CSV"]:
        print("PASS: CSV handled correctly.")
    else:
        print("FAIL: CSV sheet name error.")

def test_column_loading():
    print("\n--- Testing Column Loading ---")
    start = time.time()
    cols = read_header_file(os.path.abspath("test_load.xlsx"), sheet_name="Sheet1", header_row=1)
    duration = time.time() - start
    print(f"[XLSX] Columns (Sheet1): {cols} (Time: {duration:.4f}s)")
    
    if cols == ["A", "B"]:
        print("PASS: Columns correct.")
    else:
        print("FAIL: Columns mismatch.")

    start = time.time()
    cols2 = read_header_file(os.path.abspath("test_load.xlsx"), sheet_name="Sheet2", header_row=1)
    duration = time.time() - start
    print(f"[XLSX] Columns (Sheet2): {cols2} (Time: {duration:.4f}s)")
    
    if cols2 == ["C", "D"]:
        print("PASS: Columns correct.")
    else:
        print("FAIL: Columns mismatch.")

def clean_up():
    print("\nCleaning up...")
    if os.path.exists("test_load.xlsx"): os.remove("test_load.xlsx")
    if os.path.exists("test_load.csv"): os.remove("test_load.csv")

if __name__ == "__main__":
    try:
        create_sample_files()
        test_sheet_loading()
        test_column_loading()
    except Exception as e:
        import traceback
        traceback.print_exc()
    finally:
        clean_up()
