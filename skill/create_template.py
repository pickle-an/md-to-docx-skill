#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from docx import Document
import shutil
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_TEMPLATE = ""
OUTPUT_TEMPLATE = os.path.join(SCRIPT_DIR, 'template.docx')

def create_template_from_source():
    if not os.path.exists(SOURCE_TEMPLATE):
        print(f"Source template not found: {SOURCE_TEMPLATE}")
        print("Creating new template with default styles...")
        create_new_template()
        return
    
    shutil.copy(SOURCE_TEMPLATE, OUTPUT_TEMPLATE)
    print(f"Copied: {SOURCE_TEMPLATE} -> {OUTPUT_TEMPLATE}")
    
    doc = Document(OUTPUT_TEMPLATE)
    
    for element in doc.element.body[:]:
        if element.tag.endswith('p') or element.tag.endswith('tbl'):
            doc.element.body.remove(element)
    
    doc.save(OUTPUT_TEMPLATE)
    print(f"Template created: {OUTPUT_TEMPLATE}")
    print("Content cleared, styles preserved.")

def create_new_template():
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn
    
    doc = Document()
    
    for section in doc.sections:
        section.page_height = Cm(29.7)
        section.page_width = Cm(21.0)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
    
    styles_config = {
        'Heading 1': {'size': 22, 'bold': True, 'font_cn': '宋体', 'font_en': 'Times New Roman'},
        'Heading 2': {'size': 16, 'bold': True, 'font_cn': '宋体', 'font_en': 'Times New Roman'},
        'Heading 3': {'size': 15, 'bold': True, 'font_cn': '宋体', 'font_en': 'Times New Roman'},
        'Heading 4': {'size': 14, 'bold': True, 'font_cn': '宋体', 'font_en': 'Times New Roman'},
        'Heading 5': {'size': 14, 'bold': True, 'font_cn': '宋体', 'font_en': 'Times New Roman'},
        'Heading 6': {'size': 12, 'bold': True, 'font_cn': '宋体', 'font_en': 'Times New Roman'},
        'Normal': {'size': 12, 'bold': False, 'font_cn': '宋体', 'font_en': 'Times New Roman'},
    }
    
    for style_name, config in styles_config.items():
        try:
            style = doc.styles[style_name]
            style.font.size = Pt(config['size'])
            style.font.bold = config['bold']
            style.font.name = config['font_en']
            style._element.rPr.rFonts.set(qn('w:eastAsia'), config['font_cn'])
        except KeyError:
            pass
    
    doc.save(OUTPUT_TEMPLATE)
    print(f"New template created: {OUTPUT_TEMPLATE}")

if __name__ == '__main__':
    create_template_from_source()
