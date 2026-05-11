import zipfile
import xml.etree.ElementTree as ET

FILE = r"C:\Users\user\Downloads\ATCO HEBDO (1) du 02-03-26 au 08-03-26 Fichier de JArrou 1.xlsx"

with zipfile.ZipFile(FILE, 'r') as z:
    # List all files
    print("=== ZIP FILES ===")
    for n in sorted(z.namelist()):
        print(f"  {n}")
    
    # connections.xml
    print("\n=== CONNECTIONS ===")
    if 'xl/connections.xml' in z.namelist():
        content = z.read('xl/connections.xml').decode('utf-8', errors='replace')
        root = ET.fromstring(content)
        # Find all elements with command attribute (SQL)
        for elem in root.iter():
            tag = elem.tag.split('}')[-1]
            if tag in ('connection', 'dbPr', 'textPr'):
                attrs = dict(elem.attrib)
                if attrs:
                    print(f"  <{tag}> {attrs}")
    
    print("\n=== POWER QUERY (Formulas/Section1.m) ===")
    # Try to find M code
    for name in z.namelist():
        if 'Formulas' in name or 'Section' in name:
            try:
                raw = z.read(name)
                # try utf-8
                try:
                    text = raw.decode('utf-8')
                except:
                    text = raw.decode('utf-16', errors='replace')
                print(f"\n--- {name} ---")
                print(text[:5000])
            except Exception as e:
                print(f"  Error reading {name}: {e}")
