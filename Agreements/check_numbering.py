#!/usr/bin/env python3
"""Check numbering definitions in the 2025 reference document (see README.md)."""

import sys

from docx import Document
from docx.oxml.ns import qn
from reference_doc import find_reference_doc

reference_path = find_reference_doc()
if reference_path is None:
    sys.exit("No 2025 reference agreement (.docx) found in this folder — see README.md.")

doc = Document(str(reference_path))

# Check numbering part
numbering_part = doc.part.numbering_part
if numbering_part:
    print("Numbering part exists")
    numbering = numbering_part.numbering_definitions
    print(f"Number of numbering definitions: {len(numbering)}")
    
    # Check abstract numbering
    if hasattr(numbering_part, 'abstract_numbering'):
        print("Has abstract numbering")
    
    # Try to access the XML directly
    numbering_xml = numbering_part.element
    print(f"Numbering XML root: {numbering_xml.tag}")
    
    # Look for num elements
    num_elements = numbering_xml.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num')
    print(f"Found {len(num_elements)} num elements")
    
    for i, num in enumerate(num_elements[:5]):
        numId = num.get(qn('w:numId'))
        abstractNumId = num.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId')
        if abstractNumId is not None:
            abstract_val = abstractNumId.get(qn('w:val'))
            print(f"  Num {i+1}: numId={numId}, abstractNumId={abstract_val}")
else:
    print("No numbering part found")

