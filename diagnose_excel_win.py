import sys
import os

def check_excel_recognition():
    print("=== Excel Recognition Diagnostic (v1.0.1) ===")
    
    # 1. Check xlwings availability
    try:
        import xlwings as xw
        print(f"[OK] xlwings version: {xw.__version__}")
    except ImportError:
        print("[FAIL] xlwings is not installed.")
        return

    # 2. Check for open Excel books
    print("\n--- Listing Open Workbooks ---")
    try:
        from open_excel import list_open_books
        books = list_open_books()
        if not books:
            print("[INFO] No open workbooks found. Please open an Excel file and try again.")
        for b in books:
            print(f"- {b}")
            
            # test sheet recognition
            from open_excel import list_sheets
            sheets = list_sheets(b)
            print(f"  Sheets: {', '.join(sheets) if sheets else '(Error/None)'}")
            
            if sheets:
                from open_excel import read_header_open
                header = read_header_open(b, sheets[0], 1)
                print(f"  Header (Row 1): {header if header else '(Error/Empty)'}")
                
    except Exception as e:
        print(f"[ERROR] Recognition test failed: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_excel_recognition()
    input("\nEnter to exit...")
