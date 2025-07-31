import zipfile
import xml.etree.ElementTree as ET
import os

def extract_fonts_from_docx(docx_path):
    font_names = set()

    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            # Read document.xml inside the zipped .docx
            with docx_zip.open('word/document.xml') as document_xml:
                tree = ET.parse(document_xml)
                root = tree.getroot()

                namespaces = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                }

                # Find all rFonts tags and extract font names
                for rFonts in root.findall('.//w:rFonts', namespaces):
                    ascii_font = rFonts.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii')
                    hAnsi_font = rFonts.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi')
                    cs_font = rFonts.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs')

                    for font in [ascii_font, hAnsi_font, cs_font]:
                        if font:
                            font_names.add(font)

        return sorted(font_names)

    except zipfile.BadZipFile:
        print("Error: The file is not a valid .docx file (zip format).")
        return []
    except KeyError:
        print("Error: 'word/document.xml' not found in the .docx file.")
        return []
    except Exception as e:
        print(f"Unexpected error: {e}")
        return []

# Ask for file path
docx_path = input("Enter the full path to the .docx file: ").strip()

# Check if file exists
if not os.path.isfile(docx_path):
    print("File does not exist. Please check the path and try again.")
else:
    fonts = extract_fonts_from_docx(docx_path)

    if fonts:
        print("\nUnique fonts found in the Word document:")
        for font in fonts:
            print(font)
    else:
        print("No fonts found or there was an error processing the file.")
