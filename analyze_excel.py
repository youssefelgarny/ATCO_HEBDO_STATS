import openpyxl
from openpyxl.utils import get_column_letter
import json

FILE = r"C:\Users\user\Downloads\ATCO HEBDO (1) du 02-03-26 au 08-03-26 Fichier de JArrou 1.xlsx"

wb = openpyxl.load_workbook(FILE, data_only=True)
wb_formula = openpyxl.load_workbook(FILE, data_only=False)

report = {}

for sheetname in wb.sheetnames:
    ws = wb[sheetname]
    ws_f = wb_formula[sheetname]
    state = ws.sheet_state  # 'visible', 'hidden', 'veryHidden'
    
    # Get dimensions
    max_row = ws.max_row
    max_col = ws.max_column
    
    # Get column headers (first non-empty row)
    headers = []
    first_data_row = None
    for row in ws.iter_rows(min_row=1, max_row=min(10, max_row)):
        row_vals = [cell.value for cell in row if cell.value is not None]
        if row_vals:
            headers = [cell.value for cell in row]
            first_data_row = row[0].row
            break
    
    # Sample data (first 5 rows after header)
    sample_data = []
    if first_data_row:
        data_start = first_data_row + 1
        for row in ws.iter_rows(min_row=data_start, max_row=min(data_start + 4, max_row)):
            sample_data.append([cell.value for cell in row])
    
    # Get formulas from formula workbook (first 5 rows)
    formulas = {}
    for row in ws_f.iter_rows(min_row=1, max_row=min(15, max_row)):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                formulas[f"{cell.coordinate}"] = cell.value
    
    # Charts
    charts = []
    for chart in ws._charts:
        chart_info = {
            "type": type(chart).__name__,
            "title": str(chart.title) if chart.title else None,
        }
        try:
            chart_info["series"] = [str(s.title) if s.title else str(s.val) for s in chart.series]
        except:
            pass
        charts.append(chart_info)
    
    # Merged cells
    merged = [str(m) for m in ws.merged_cells.ranges]
    
    # Defined names / tables
    tables = list(ws.tables.keys()) if hasattr(ws, 'tables') else []
    
    report[sheetname] = {
        "state": state,
        "dimensions": f"{max_col} cols x {max_row} rows",
        "headers": headers,
        "sample_data": sample_data,
        "formulas_sample": dict(list(formulas.items())[:20]),
        "charts": charts,
        "merged_cells_count": len(merged),
        "tables": tables,
    }

print(json.dumps(report, ensure_ascii=False, default=str, indent=2))
