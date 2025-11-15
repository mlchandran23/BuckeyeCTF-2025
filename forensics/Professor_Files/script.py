import zipfile
import xml.etree.ElementTree as ET
from pprint import pprint

def extract_metadata(docx_path):
    metadata = {}

    with zipfile.ZipFile(docx_path, 'r') as docx:
        # Core properties (standard metadata)
        if 'docProps/core.xml' in docx.namelist():
            core = ET.fromstring(docx.read('docProps/core.xml'))
            ns = {'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                  'dc': 'http://purl.org/dc/elements/1.1/',
                  'dcterms': 'http://purl.org/dc/terms/'}
            for tag in ['title', 'subject', 'creator', 'description', 'created', 'modified']:
                el = core.find(f'dc:{tag}', ns) or core.find(f'dcterms:{tag}', ns)
                if el is not None:
                    metadata[tag] = el.text

        # Custom properties (user-defined metadata)
        if 'docProps/custom.xml' in docx.namelist():
            custom = ET.fromstring(docx.read('docProps/custom.xml'))
            ns = {'cp': 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
                  'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'}
            for prop in custom.findall('cp:property', ns):
                name = prop.attrib.get('name')
                value = next(iter(prop))
                metadata[name] = value.text

    return metadata

if __name__ == "__main__":
    path = "OSU_Ethics_Report.docx"  # Replace with your file path
    meta = extract_metadata(path)
    print("\n📄 Extracted Metadata:")
    pprint(meta)
