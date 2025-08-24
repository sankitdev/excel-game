import sys
from pathlib import Path
from openpyxl import load_workbook
import os

def is_missing(v):
    return v is None or (isinstance(v, str) and v.strip() in ("", "-"))

def find_header_row_and_cols(ws, search_rows=20, max_cols=50):
    """
    Scan the top of the sheet to find the header row containing:
    - Name or Employee Name
    - Check In Time
    - Check Out Time
    Returns (header_row_index, name_col, checkin_col, checkout_col)
    """
    name_headers = {"name", "employee name"}
    cin_headers = {"check in time", "checkin time", "check in", "checkin"}
    cout_headers = {"check out time", "checkout time", "check out", "checkout"}

    for r in range(1, min(search_rows, ws.max_row) + 1):
        labels = {}
        for c in range(1, min(max_cols, ws.max_column) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                val = v.strip().lower()
                labels[val] = c
        # find required columns in this row
        name_col = next((labels[k] for k in name_headers if k in labels), None)
        cin_col = next((labels[k] for k in cin_headers if k in labels), None)
        cout_col = next((labels[k] for k in cout_headers if k in labels), None)

        if name_col and cin_col and cout_col:
            return r, name_col, cin_col, cout_col

    raise RuntimeError("Could not find header row with Name / Check In Time / Check Out Time")

def process_ws(ws):
    # 1) detect header + columns
    header_row, name_col, cin_col, cout_col = find_header_row_and_cols(ws)

    # 2) iterate from first data row to bottom
    first_data_row = header_row + 1
    last_row = ws.max_row

    r = first_data_row
    while r <= last_row:
        name = ws.cell(r, name_col).value

        # group is a contiguous block of same name (consecutive rows)
        start = r
        while r <= last_row and ws.cell(r, name_col).value == name:
            r += 1
        end = r - 1  # inclusive

        # walk inside this block, pulling missing values from prior rows in the same block
        for i in range(start, end + 1):
            cin_cell = ws.cell(i, cin_col)
            cout_cell = ws.cell(i, cout_col)
            cin_val = cin_cell.value
            cout_val = cout_cell.value

            # skip if both missing
            if is_missing(cin_val) and is_missing(cout_val):
                continue

            need_cin = is_missing(cin_val) and not is_missing(cout_val)
            need_cout = is_missing(cout_val) and not is_missing(cin_val)

            if not need_cin and not need_cout:
                continue

            # look upward within the same block for the closest previous row with the needed value
            j = i - 1
            while j >= start and ws.cell(j, name_col).value == name and (need_cin or need_cout):
                prev_cin = ws.cell(j, cin_col).value
                prev_cout = ws.cell(j, cout_col).value

                if need_cin and not is_missing(prev_cin):
                    cin_cell.value = prev_cin
                    need_cin = False

                if need_cout and not is_missing(prev_cout):
                    cout_cell.value = prev_cout
                    need_cout = False

                j -= 1
            # If still missing after search, we leave it as-is (no guess)

def process_excel(file):
    """
    file: either a file path (str/Path) or file-like object (BytesIO / Streamlit UploadedFile)
    Returns: openpyxl Workbook object
    """
    # detect if .xlsm
    name = getattr(file, "name", "")  # UploadedFile has .name
    suffix = Path(name).suffix.lower() if name else ""
    keep_vba = suffix == ".xlsm"

    wb = load_workbook(file, keep_vba=keep_vba)
    ws = wb.worksheets[0]  # first sheet only
    process_ws(ws)
    return wb

def main():
    if len(sys.argv) < 2:
        print("Usage: python process_attendance.py <input.xlsx>")
        sys.exit(1)

    inp = Path(sys.argv[1])
    if not inp.exists():
        print(f"Input file not found: {inp}")
        sys.exit(1)

    # Auto-generate output filename
    base, ext = os.path.splitext(str(inp))
    out = f"{base}_processed{ext}"

    # Load workbook safely (don't force keep_vba=True unless it's .xlsm)
    if inp.suffix.lower() == ".xlsm":
        wb = load_workbook(inp, keep_vba=True)
    else:
        wb = load_workbook(inp)

    # Process ONLY first sheet
    ws = wb.worksheets[0]
    process_ws(ws)

    wb.save(out)
    print(f"Done. Wrote: {out}")

if __name__ == "__main__":
    main()
