import zipfile
import base64
import zlib
import struct

FILE = r"C:\Users\user\Downloads\ATCO HEBDO (1) du 02-03-26 au 08-03-26 Fichier de JArrou 1.xlsx"

with zipfile.ZipFile(FILE, 'r') as z:
    # The M code is stored in customXml/item1.xml as base64 encoded blob
    content = z.read('customXml/item1.xml')
    # Find base64 encoded section
    text = content.decode('utf-8', errors='replace')
    
    # The mashup blob is base64 inside the XML
    # Extract from DataMashup element
    import re
    # Find the base64 content
    match = re.search(r'<DataMashup[^>]*>([\s\S]+?)</DataMashup>', text)
    if not match:
        # Try to get raw content after XML header
        # The content is actually the entire blob
        print("RAW XML preview:")
        print(text[:500])
        print("...")
        # Try finding AAAA pattern (base64)
        b64_match = re.search(r'([A-Za-z0-9+/]{100,}={0,2})', text)
        if b64_match:
            b64 = b64_match.group(1)
            print(f"\nFound base64 block ({len(b64)} chars):")
            try:
                decoded = base64.b64decode(b64)
                print(f"Decoded length: {len(decoded)} bytes")
                # This is a ZIP within the mashup package
                # Try to read as inner zip
                import io
                # Skip first 4 bytes (length prefix)
                inner_zip_data = decoded[4:]
                try:
                    with zipfile.ZipFile(io.BytesIO(inner_zip_data)) as inner:
                        print("Inner zip files:")
                        for name in inner.namelist():
                            print(f"  {name}")
                        # Read the M formula file
                        for name in inner.namelist():
                            if 'Section' in name or 'Formula' in name:
                                m_code = inner.read(name).decode('utf-8', errors='replace')
                                print(f"\n=== {name} ===")
                                print(m_code[:8000])
                except Exception as e:
                    print(f"Inner zip error: {e}")
                    # Try raw decompression
                    print("First 200 decoded bytes (hex):", decoded[:200].hex())
            except Exception as e:
                print(f"Base64 decode error: {e}")
    else:
        b64 = match.group(1).strip()
        print(f"Found DataMashup base64 ({len(b64)} chars)")
