import openpyxl
import zipfile

FILE = r"C:\Users\user\Downloads\ATCO HEBDO (1) du 02-03-26 au 08-03-26 Fichier de JArrou 1.xlsx"

# 1. List all sheets with visibility
print("=== SHEETS ===")
wb = openpyxl.load_workbook(FILE, data_only=True)
for name in wb.sheetnames:
    ws = wb[name]
    state = ws.sheet_state
    ncharts = len(ws._charts)
    print(f"  [{state}] {name} | rows={ws.max_row} cols={ws.max_column} charts={ncharts}")

# 2. List SQL queries from connections
print("\n=== SQL QUERIES (from xl/connections.xml) ===")
with zipfile.ZipFile(FILE, 'r') as z:
    if 'xl/connections.xml' in z.namelist():
        import xml.etree.ElementTree as ET
        content = z.read('xl/connections.xml')
        root = ET.fromstring(content)
        ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for conn in root.iter():
            if conn.tag.endswith('}connection') or conn.tag == 'connection':
                cid = conn.get('id','')
                cname = conn.get('name','')
                print(f"\n  Connection id={cid} name={cname}")
            if conn.tag.endswith('}dbPr'):
                print(f"    SQL: {conn.get('command','')}")

# 3. Charts detail per sheet
print("\n=== CHARTS DETAIL ===")
for name in wb.sheetnames:
    ws = wb[name]
    if ws._charts:
        for i, chart in enumerate(ws._charts):
            ctype = type(chart).__name__
            title = None
            try:
                if chart.title:
                    if hasattr(chart.title, 'text'):
                        title = ''.join(r.t for p in chart.title.text.rich.p for r in (p.r or []) if r.t)
                    else:
                        title = str(chart.title)
            except: pass
            print(f"  Sheet '{name}' Chart#{i+1}: {ctype} | title={title}")

# 4. Sample data from first visible sheets
print("\n=== SAMPLE DATA (first 3 rows, visible sheets) ===")
for name in wb.sheetnames:
    ws = wb[name]
    if ws.sheet_state == 'visible':
        print(f"\n  [{name}]")
        for row in ws.iter_rows(min_row=1, max_row=4, values_only=True):
            vals = [str(v)[:30] if v is not None else '' for v in row[:15]]
            if any(vals):
                print(f"    {vals}")
