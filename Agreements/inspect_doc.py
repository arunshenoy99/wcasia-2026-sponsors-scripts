#!/usr/bin/env python3
"""Inspect Word document structure"""

from docx import Document

doc = Document("1-Elementor WordCamp Asia 2025 Super Admin Sponsorship Agreement.docx")

print("=== Document Structure ===\n")

for i, para in enumerate(doc.paragraphs[:50]):
    style_name = para.style.name if para.style else "None"
    is_bold = any(run.bold for run in para.runs) if para.runs else False
    text = para.text[:100] if para.text else ""
    
    # Check for numbering
    num_info = ""
    if para._element.pPr is not None and para._element.pPr.numPr is not None:
        numPr = para._element.pPr.numPr
        ilvl = numPr.ilvl
        numId = numPr.numId
        if ilvl is not None:
            num_info = f" [List Level: {ilvl.val}]"
        if numId is not None:
            num_info += f" [NumId: {numId.val}]"
    
    print(f"{i+1:3d}. Style: {style_name:20s} Bold: {str(is_bold):5s}{num_info}")
    if text:
        print(f"     Text: {text}")
    print()

