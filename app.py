#!/usr/bin/env python3
"""
Simple script to convert LaTeX in Word documents
Just put your .docx file in the same folder and run this script
"""

import os
import sys
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
from datetime import datetime

def convert_latex_symbols(text):
    """Convert LaTeX symbols to Unicode"""
    if not text:
        return text
    
    result = text
    
    # Prime notation: A^{prime} -> A'
    result = re.sub(r'([A-Za-z]+)\^\{prime\}', r"\1'", result)
    result = re.sub(r'([A-Za-z]+)\s*\(prime\)', r"\1'", result)
    result = result.replace('(prime)', "'").replace('prime', "'")
    
    # Superscript: A^{2} -> A²
    superscript_map = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', 
                      '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    
    # Handle ^{digit}
    for digit, sup in superscript_map.items():
        result = result.replace(f'^{{{digit}}}', sup)
        result = result.replace(f'^{digit}', sup)
    
    # Math symbols
    symbols = {
        '\\pi': 'π', '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
        '\\theta': 'θ', '\\lambda': 'λ', '\\mu': 'μ', '\\sigma': 'σ', '\\phi': 'φ',
        '\\in': '∈', '\\leq': '≤', '\\geq': '≥', '\\neq': '≠', '\\approx': '≈',
        '\\infty': '∞', '\\pm': '±', '\\times': '×', '\\div': '÷',
        '\\mathbb{R}': 'ℝ', '\\mathbb{Z}': 'ℤ', '\\mathbb{N}': 'ℕ', '\\mathbb{Q}': 'ℚ',
        '\\sqrt{3}': '√3', '\\cot': 'cot', '\\tan': 'tan', '\\lim': 'lim'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Clean brackets
    result = result.replace('\\left(', '(').replace('\\right)', ')')
    result = result.replace('\\left[', '[').replace('\\right]', ']')
    result = result.replace('\\lbrack', '[').replace('\\rbrack', ']')
    
    # Fractions: \frac{a}{b} -> (a)/(b)
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        try:
            # Find numerator
            num_start = start + 6
            brace_count = 1
            num_end = num_start
            
            while num_end < len(result) and brace_count > 0:
                if result[num_end] == '{':
                    brace_count += 1
                elif result[num_end] == '}':
                    brace_count -= 1
                num_end += 1
            
            numerator = result[num_start:num_end-1]
            
            # Find denominator
            if num_end < len(result) and result[num_end] == '{':
                denom_start = num_end + 1
                brace_count = 1
                denom_end = denom_start
                
                while denom_end < len(result) and brace_count > 0:
                    if result[denom_end] == '{':
                        brace_count += 1
                    elif result[denom_end] == '}':
                        brace_count -= 1
                    denom_end += 1
                
                denominator = result[denom_start:denom_end-1]
                fraction = f"({numerator})/({denominator})"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        except:
            break
    
    # Clean up
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Find $...$ and ${...}$ expressions"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Handle ${...}$ format
            if i < len(text) and text[i] == '{':
                i += 1
                brace_count = 1
                content_start = i
                
                while i < len(text) and brace_count > 0:
                    if text[i] == '{':
                        brace_count += 1
                    elif text[i] == '}':
                        brace_count -= 1
                    i += 1
                
                if i < len(text) and text[i] == '$':
                    end = i + 1
                    content = text[content_start:i-1]
                    if content.strip():
                        expressions.append((start, end, content.strip()))
                    i += 1
                else:
                    i += 1
            else:
                # Handle $...$ format
                content_start = i
                
                while i < len(text) and text[i] != '$':
                    i += 1
                
                if i < len(text):
                    end = i + 1
                    content = text[content_start:i]
                    if content.strip():
                        expressions.append((start, end, content.strip()))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_equation_object(paragraph, latex_text):
    """Create equation object in Word"""
    unicode_text = convert_latex_symbols(latex_text)
    
    try:
        # Try OMML equation object
        safe_text = unicode_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:rPr>
                    <m:scr m:val="roman"/>
                    <m:sty m:val="i"/>
                </m:rPr>
                <m:t>{safe_text}</m:t>
            </m:r>
        </m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except:
        pass
    
    try:
        # Fallback: styled text
        run = paragraph.add_run(unicode_text)
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 32, 96)
        return True
    except:
        # Final fallback
        run = paragraph.add_run(unicode_text)
        run.italic = True
        return True

def process_text_with_math(paragraph, text):
    """Process text with math expressions"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    
    for start, end, latex_content in expressions:
        # Add text before math
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Create equation
        create_equation_object(paragraph, latex_content)
        math_count += 1
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format questions and answers"""
    paragraph.clear()
    
    # Question pattern: "Câu 1."
    question_match = re.match(r'^(Câu\s+\d+[\.:])\s*(.*)', text, re.IGNORECASE)
    if question_match:
        question_part = question_match.group(1)
        content_part = question_match.group(2)
        
        # Bold question number
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Answer pattern: "A.", "B.", etc.
    answer_match = re.match(r'^([A-Da-d])[\.\)]\s*(.*)', text)
    if answer_match:
        answer_letter = answer_match.group(1).upper() + '.'
        answer_content = answer_match.group(2)
        
        # Bold answer letter
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Not Q&A format
    process_text_with_math(paragraph, text)
    return False

def process_document(input_file):
    """Main processing function"""
    try:
        doc = Document(input_file)
        new_doc = Document()
        
        math_count = 0
        qa_count = 0
        
        print(f"Processing: {input_file}")
        print(f"Found: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
        
        # Process paragraphs
        for para in doc.paragraphs:
            text = para.text.strip()
            
            if not text:
                new_doc.add_paragraph()
                continue
            
            new_para = new_doc.add_paragraph()
            
            if format_question_answer(new_para, text):
                qa_count += 1
            else:
                math_count += process_text_with_math(new_para, text)
        
        # Process tables
        for table in doc.tables:
            if not table.rows:
                continue
            
            num_rows = len(table.rows)
            num_cols = len(table.rows[0].cells) if table.rows else 0
            
            new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
            
            for r in range(num_rows):
                for c in range(min(len(table.rows[r].cells), num_cols)):
                    original_cell = table.rows[r].cells[c]
                    new_cell = new_table.rows[r].cells[c]
                    
                    cell_text = original_cell.text.strip()
                    if cell_text:
                        new_cell.text = ""
                        cell_para = new_cell.paragraphs[0] if new_cell.paragraphs else new_cell.add_paragraph()
                        cell_para.clear()
                        
                        if not format_question_answer(cell_para, cell_text):
                            math_count += process_text_with_math(cell_para, cell_text)
                        
                        cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Format table
            try:
                new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                if new_table.rows:
                    for cell in new_table.rows[0].cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.bold = True
            except:
                pass
        
        # Save output
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = input_file.replace('.docx', f'_converted_{timestamp}.docx')
        new_doc.save(output_file)
        
        print(f"\n✅ Success!")
        print(f"📄 Output: {output_file}")
        print(f"🔢 Math expressions converted: {math_count}")
        print(f"❓ Q&A items formatted: {qa_count}")
        
        return output_file
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def main():
    """Main function - automatically find and process .docx files"""
    print("🔍 Looking for .docx files in current directory...")
    
    docx_files = [f for f in os.listdir('.') if f.endswith('.docx') and not f.startswith('~')]
    
    if not docx_files:
        print("❌ No .docx files found in current directory!")
        print("📋 Put your Word file in the same folder as this script and try again.")
        return
    
    print(f"📁 Found {len(docx_files)} file(s):")
    for i, file in enumerate(docx_files, 1):
        print(f"   {i}. {file}")
    
    if len(docx_files) == 1:
        selected_file = docx_files[0]
        print(f"\n🚀 Processing: {selected_file}")
    else:
        try:
            choice = int(input(f"\nSelect file (1-{len(docx_files)}): ")) - 1
            selected_file = docx_files[choice]
        except (ValueError, IndexError):
            print("❌ Invalid choice!")
            return
    
    # Process the file
    output_file = process_document(selected_file)
    
    if output_file:
        print(f"\n🎉 Done! Your converted file is ready: {output_file}")
    else:
        print("\n❌ Processing failed!")

if __name__ == "__main__":
    main()
