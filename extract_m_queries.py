import zipfile
import base64
import re
import io

FILE = r"C:\Users\user\Downloads\ATCO HEBDO (1) du 02-03-26 au 08-03-26 Fichier de JArrou 1.xlsx"

with zipfile.ZipFile(FILE, 'r') as z:
    content = z.read('customXml/item1.xml')
    text = content.decode('utf-8', errors='replace')
    
    # Extract AAAA...
    b64_match = re.search(r'>([A-Za-z0-9+/]{100,}={0,2})<', text)
    if not b64_match:
        b64_match = re.search(r'([A-Za-z0-9+/]{100,}={0,2})', text)
        
    if b64_match:
        b64 = b64_match.group(1)
        decoded = base64.b64decode(b64)
        
        # In Mashup Data, there is a header. Usually zip starts with PK (50 4B 03 04)
        idx = decoded.find(b'PK\x03\x04')
        if idx != -1:
            zip_data = decoded[idx:]
            try:
                with zipfile.ZipFile(io.BytesIO(zip_data)) as inner:
                    for name in inner.namelist():
                        if name.endswith('.m'):
                            print(f"\n======== {name} ========")
                            # Read and decode
                            m_code = inner.read(name).decode('utf-8', errors='replace')
                            print(m_code)
            except Exception as e:
                print("Error extracting inner zip:", e)
        else:
            print("Zip signature not found in decoded blob")
