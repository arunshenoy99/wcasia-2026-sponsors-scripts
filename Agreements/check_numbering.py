#!/usr/bin/env python3
"""Check numbering definitions in original document"""

from docx import Document
from docx.oxml.ns import qn

doc = Document("1-Elementor WordCamp Asia 2025 Super Admin Sponsorship Agreement.docx")

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

