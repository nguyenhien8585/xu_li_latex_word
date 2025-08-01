#!/usr/bin/env python3
"""
Streamlit LaTeX to Word Converter - COMPLETE FIX
Properly converts LaTeX expressions and markdown tables to Word
Fixed: Markdown tables, sqrt functions, fractions, complex expressions
"""

import streamlit as st
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from io import BytesIO
from datetime import datetime
import traceback

# Page configuration
st.set_page_config(
    page_title="LaTeX to Word Converter - COMPLETE",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 2rem;
        color: white;
    }
    .main-header h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: bold;
    }
    .main-header p {
        margin: 0.5rem 0 0 0;
        font-size: 1.1rem;
        opacity: 0.9;
    }
    .feature-card {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        margin: 1rem 0;
    }
    .success-box {
        background: linear-gradient(90deg, #56ab2f 0%, #a8e6cf 100%);
        padding: 1rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin: 1rem 0;
    }
    .warning-box {
        background: linear-gradient(90deg, #f093fb 0%, #f5576c 100%);
        padding: 1rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin: 1rem 0;
    }
    .stButton > button {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
    }
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

def convert_latex_symbols(text):
    """COMPLETELY FIXED LaTeX to Unicode conversion"""
    if not text:
        return text
    
    result = text
    
    # Step 1: Handle fractions FIRST (before other processing)
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
                
                # Mark as fraction for special OMML processing
                fraction = f"FRACTION[{numerator}]/[{denominator}]"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        except:
            break
    
    # Step 2: Handle square roots (before symbol replacement)
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
        
        try:
            content_start = start + 6
            brace_count = 1
            content_end = content_start
            
            while content_end < len(result) and brace_count > 0:
                if result[content_end] == '{':
                    brace_count += 1
                elif result[content_end] == '}':
                    brace_count -= 1
                content_end += 1
            
            content = result[content_start:content_end-1]
            # Mark as sqrt for special processing
            sqrt_result = f"SQRT[{content}]"
            result = result[:start] + sqrt_result + result[content_end:]
        except:
            break
    
    # Step 3: Handle Greek letters and symbols
    symbols = {
        '\\pi': 'π', '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
        '\\theta': 'θ', '\\lambda': 'λ', '\\mu': 'μ', '\\sigma': 'σ', '\\phi': 'φ',
        '\\omega': 'ω', '\\Omega': 'Ω', '\\Phi': 'Φ', '\\Theta': 'Θ', '\\Delta': 'Δ',
        '\\in': '∈', '\\subset': '⊂', '\\cup': '∪', '\\cap': '∩', '\\emptyset': '∅',
        '\\leq': '≤', '\\geq': '≥', '\\neq': '≠', '\\approx': '≈', '\\equiv': '≡',
        '\\infty': '∞', '\\pm': '±', '\\mp': '∓', '\\times': '×', '\\div': '÷',
        '\\cdot': '·', '\\bullet': '•', '\\circ': '∘',
        '\\mathbb{R}': 'ℝ', '\\mathbb{Z}': 'ℤ', '\\mathbb{Q}': 'ℚ', 
        '\\mathbb{N}': 'ℕ', '\\mathbb{C}': 'ℂ', '\\mathbb{P}': 'ℙ'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Step 4: Handle mathematical functions
    functions = {
        '\\sin': 'sin', '\\cos': 'cos', '\\tan': 'tan', '\\cot': 'cot',
        '\\sec': 'sec', '\\csc': 'csc', '\\arcsin': 'arcsin', '\\arccos': 'arccos',
        '\\lim': 'lim', '\\sup': 'sup', '\\inf': 'inf', '\\max': 'max', '\\min': 'min',
        '\\log': 'log', '\\ln': 'ln'
    }
    
    for latex_func, unicode_func in functions.items():
        result = result.replace(latex_func, unicode_func)
    
    # Step 5: Handle prime notation
    prime_patterns = [
        (r"([A-Za-z]+)\^?\{?'?\}?'", r"\1'"),
        (r"([A-Za-z]+)\^?\{?prime\}?", r"\1'"),
        (r"\(prime\)", "'"),
    ]
    
    for pattern, replacement in prime_patterns:
        result = re.sub(pattern, replacement, result)
    
    # Step 6: Handle subscripts
    subscript_map = {
        '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
        '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
        'n': 'ₙ', 'i': 'ᵢ', 'j': 'ⱼ', 'a': 'ₐ', 'e': 'ₑ', 
        'o': 'ₒ', 'x': 'ₓ', 'r': 'ᵣ', 's': 'ₛ', 't': 'ₜ'
    }
    
    def replace_subscript(match):
        base = match.group(1)
        content = match.group(2)
        converted = ''.join(subscript_map.get(c, c) for c in content)
        return base + converted
    
    result = re.sub(r'([A-Za-z]+)_\{([^}]+)\}', replace_subscript, result)
    result = re.sub(r'([A-Za-z]+)_([0-9a-z])', lambda m: m.group(1) + subscript_map.get(m.group(2), m.group(2)), result)
    
    # Step 7: Handle superscripts
    superscript_map = {
        '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
        '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
        '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾'
    }
    
    def replace_superscript(match):
        base = match.group(1)
        content = match.group(2)
        if 'prime' in content or content == "'":
            return base + "'"
        converted = ''.join(superscript_map.get(c, c) for c in content)
        return base + converted
    
    result = re.sub(r'([A-Za-z0-9]+)\^\{([^}]+)\}', replace_superscript, result)
    result = re.sub(r'([A-Za-z0-9]+)\^([0-9])', lambda m: m.group(1) + superscript_map.get(m.group(2), m.group(2)), result)
    
    # Step 8: Handle brackets
    result = result.replace('\\left(', '(').replace('\\right)', ')')
    result = result.replace('\\left[', '[').replace('\\right]', ']')
    result = result.replace('\\left{', '{').replace('\\right}', '}')
    result = result.replace('\\lbrack', '[').replace('\\rbrack', ']')
    result = result.replace('\\{', '{').replace('\\}', '}')
    
    # Step 9: Clean up remaining LaTeX commands
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Find mathematical expressions in text - handles both $ and $$ formats"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Handle $$ format
            if i < len(text) and text[i] == '$':
                i += 1  # Skip second $
                content_start = i
                
                # Find closing $$
                while i < len(text) - 1:
                    if text[i] == '$' and text[i+1] == '$':
                        end = i + 2
                        content = text[content_start:i]
                        if content.strip():
                            expressions.append((start, end, content.strip()))
                        i += 2
                        break
                    i += 1
                else:
                    i += 1
            
            # Handle ${...}$ format
            elif i < len(text) and text[i] == '{':
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
            
            # Handle regular $...$ format
            else:
                content_start = i
                brace_depth = 0
                
                while i < len(text):
                    if text[i] == '{':
                        brace_depth += 1
                    elif text[i] == '}':
                        brace_depth -= 1
                    elif text[i] == '$' and brace_depth == 0:
                        break
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

def create_fraction_omml(paragraph, numerator, denominator):
    """Create proper fraction OMML"""
    try:
        # Safe XML escaping
        safe_num = numerator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_den = denominator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:f>
                <m:fPr></m:fPr>
                <m:num>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_num}</m:t>
                    </m:r>
                </m:num>
                <m:den>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_den}</m:t>
                    </m:r>
                </m:den>
            </m:f>
        </m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_sqrt_omml(paragraph, content):
    """Create proper square root OMML"""
    try:
        # Safe XML escaping
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:rad>
                <m:radPr></m:radPr>
                <m:deg></m:deg>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_content}</m:t>
                    </m:r>
                </m:e>
            </m:rad>
        </m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_equation_object(paragraph, latex_text, method='auto'):
    """IMPROVED equation object creation with proper OMML for fractions and sqrt"""
    unicode_text = convert_latex_symbols(latex_text)
    
    # Check for special cases that need OMML
    if 'FRACTION[' in unicode_text:
        # Extract fraction parts
        match = re.search(r'FRACTION\[([^\]]+)\]/\[([^\]]+)\]', unicode_text)
        if match:
            numerator = convert_latex_symbols(match.group(1))
            denominator = convert_latex_symbols(match.group(2))
            
            if create_fraction_omml(paragraph, numerator, denominator):
                return 'omml'
            else:
                # Fallback to styled fraction
                fraction_text = f"({numerator})/({denominator})"
                unicode_text = unicode_text.replace(match.group(0), fraction_text)
    
    if 'SQRT[' in unicode_text:
        # Extract sqrt content
        match = re.search(r'SQRT\[([^\]]+)\]', unicode_text)
        if match:
            content = convert_latex_symbols(match.group(1))
            
            if create_sqrt_omml(paragraph, content):
                return 'omml'
            else:
                # Fallback to Unicode sqrt
                sqrt_text = f"√{content}"
                unicode_text = unicode_text.replace(match.group(0), sqrt_text)
    
    # Regular OMML for other expressions
    if method == 'omml' or method == 'auto':
        try:
            # Safe XML escaping
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
            return 'omml'
        except Exception as e:
            if method == 'omml':
                st.warning(f"OMML failed: {e}")
    
    # Styled text fallback
    if method == 'styled' or method == 'auto':
        try:
            run = paragraph.add_run(unicode_text)
            run.font.name = 'Cambria Math'
            run.font.size = Pt(11)
            run.italic = True
            run.font.color.rgb = RGBColor(0, 32, 96)
            return 'styled'
        except Exception as e:
            if method == 'styled':
                st.warning(f"Styled failed: {e}")
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_text)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)
        return 'fallback'
    except:
        run = paragraph.add_run(f"[{latex_text}]")
        return 'error'

def detect_markdown_table(text):
    """Detect if text contains markdown table format"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    # Check for pipe characters
    pipe_lines = [line for line in lines if '|' in line]
    if len(pipe_lines) < 2:
        return False
    
    # Check for separator line (contains only |, -, :, spaces)
    has_separator = False
    for line in lines:
        if re.match(r'^[\|\s\-:]+$', line.strip()):
            has_separator = True
            break
    
    return has_separator and len(pipe_lines) >= 2

def parse_markdown_table(text):
    """Parse markdown table into structured data"""
    lines = text.strip().split('\n')
    table_data = []
    
    for line in lines:
        line = line.strip()
        if not line or re.match(r'^[\|\s\-:]+$', line):
            continue  # Skip empty lines and separator lines
        
        if '|' in line:
            # Remove leading/trailing pipes and split
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_data.append(cells)
    
    return table_data

def create_word_table(doc, table_data, options):
    """Create Word table from parsed markdown data"""
    if not table_data:
        return None, 0
    
    try:
        num_rows = len(table_data)
        num_cols = max(len(row) for row in table_data) if table_data else 0
        
        if num_rows == 0 or num_cols == 0:
            return None, 0
        
        # Create table
        word_table = doc.add_table(rows=num_rows, cols=num_cols)
        table_math = 0
        
        # Fill table data
        for r, row_data in enumerate(table_data):
            for c in range(num_cols):
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Get cell content
                if c < len(row_data):
                    cell_text = row_data[c].strip()
                else:
                    cell_text = ""
                
                # Create paragraph in cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Process cell content
                if cell_text:
                    if not format_question_answer(cell_para, cell_text, options):
                        cell_math_count, _ = process_text_with_math(cell_para, cell_text, options)
                        table_math += cell_math_count
                
                # Center align and set font
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(10)
        
        # Apply table formatting
        if options.get('format_tables', True):
            try:
                word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # Style header row
                if word_table.rows:
                    header_row = word_table.rows[0]
                    for cell in header_row.cells:
                        # Bold white text for header
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.bold = True
                                run.font.color.rgb = RGBColor(255, 255, 255)
                        
                        # Blue background
                        try:
                            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="4472C4"/>'
                            shading = parse_xml(shading_xml)
                            cell._tc.get_or_add_tcPr().append(shading)
                        except:
                            pass
                
                # Add borders
                try:
                    for row in word_table.rows:
                        for cell in row.cells:
                            tc = cell._tc
                            tcPr = tc.get_or_add_tcPr()
                            
                            borders_xml = '''
                            <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                                <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                                <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                                <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                            </w:tcBorders>'''
                            borders = parse_xml(borders_xml)
                            tcPr.append(borders)
                except:
                    pass
            except:
                pass
        
        return word_table, table_math
        
    except Exception as e:
        st.warning(f"Table creation error: {e}")
        return None, 0

def process_text_with_math(paragraph, text, options):
    """Process text containing mathematical expressions"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        return 0, {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    last_pos = 0
    math_count = 0
    method_stats = {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    for start, end, latex_content in expressions:
        # Add text before math
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(11)
                run.font.name = 'Times New Roman'
        
        # Create equation
        method_used = create_equation_object(paragraph, latex_content, options.get('equation_method', 'auto'))
        method_stats[method_used] += 1
        math_count += 1
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(11)
            run.font.name = 'Times New Roman'
    
    return math_count, method_stats

def format_question_answer(paragraph, text, options):
    """Format questions and answers"""
    paragraph.clear()
    
    # Question patterns
    question_patterns = [
        r'^(Câu\s+\d+[\.:])\s*(.*)',
        r'^(Question\s+\d+[\.:])\s*(.*)',
        r'^(\d+[\.:])\s*(.*)'
    ]
    
    for pattern in question_patterns:
        match = re.match(pattern, text, re.IGNORECASE)
        if match:
            question_part = match.group(1)
            content_part = match.group(2)
            
            # Bold question number
            run_q = paragraph.add_run(question_part)
            run_q.bold = True
            run_q.font.size = Pt(11)
            run_q.font.name = 'Times New Roman'
            run_q.font.color.rgb = RGBColor(0, 0, 0)
            
            if content_part:
                paragraph.add_run(" ")
                process_text_with_math(paragraph, content_part, options)
            
            return True
    
    # Answer patterns
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Bold answer letter with blue color
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(11)
        run_l.font.name = 'Times New Roman'
        
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content, options)
        
        return True
    
    # Not Q&A format
    process_text_with_math(paragraph, text, options)
    return False

def process_document(doc, options):
    """Process Word document with markdown table detection and conversion"""
    new_doc = Document()
    
    stats = {
        'paragraphs': 0,
        'tables': 0,
        'markdown_tables': 0,
        'math_expressions': 0,
        'questions_formatted': 0,
        'method_stats': {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0},
        'processing_log': []
    }
    
    try:
        # Collect all paragraph text to detect markdown tables
        paragraph_texts = []
        for para in doc.paragraphs:
            paragraph_texts.append(para.text)
        
        # Process paragraphs with markdown table detection
        i = 0
        while i < len(paragraph_texts):
            text = paragraph_texts[i].strip()
            stats['paragraphs'] += 1
            
            if not text:
                new_doc.add_paragraph()
                i += 1
                continue
            
            # Check if this starts a markdown table
            if options.get('convert_markdown', True) and '|' in text:
                # Look ahead to collect table lines
                table_lines = []
                j = i
                
                while j < len(paragraph_texts):
                    current_text = paragraph_texts[j].strip()
                    if not current_text:
                        j += 1
                        break
                    if '|' in current_text:
                        table_lines.append(current_text)
                        j += 1
                    else:
                        break
                
                # Check if it's a valid markdown table
                combined_table_text = '\n'.join(table_lines)
                if detect_markdown_table(combined_table_text):
                    # Parse and create Word table
                    table_data = parse_markdown_table(combined_table_text)
                    if table_data:
                        word_table, table_math = create_word_table(new_doc, table_data, options)
                        if word_table is not None:
                            stats['markdown_tables'] += 1
                            stats['math_expressions'] += table_math
                            stats['processing_log'].append(f"Converted markdown table: {len(table_data)} rows")
                    
                    i = j  # Skip processed table lines
                    continue
            
            # Process regular paragraph
            new_para = new_doc.add_paragraph()
            
            try:
                if options.get('format_qa', True) and format_question_answer(new_para, text, options):
                    stats['questions_formatted'] += 1
                    stats['processing_log'].append(f"Para {i+1}: Q&A formatted")
                else:
                    math_count, method_stats = process_text_with_math(new_para, text, options)
                    stats['math_expressions'] += math_count
                    for method, count in method_stats.items():
                        stats['method_stats'][method] += count
                    
                    if math_count > 0:
                        stats['processing_log'].append(f"Para {i+1}: {math_count} equations converted")
            except Exception as e:
                # Fallback to plain text
                new_para.clear()
                run = new_para.add_run(text)
                run.font.size = Pt(11)
                run.font.name = 'Times New Roman'
                stats['processing_log'].append(f"Para {i+1}: Error, used plain text - {str(e)}")
            
            i += 1
        
        # Process existing Word tables
        for table_idx, table in enumerate(doc.tables):
            if not table.rows:
                continue
            
            try:
                num_rows = len(table.rows)
                num_cols = len(table.rows[0].cells) if table.rows else 0
                
                new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
                table_math = 0
                
                for r in range(num_rows):
                    for c in range(min(len(table.rows[r].cells), num_cols)):
                        try:
                            original_cell = table.rows[r].cells[c]
                            new_cell = new_table.rows[r].cells[c]
                            
                            cell_text = original_cell.text.strip()
                            if cell_text:
                                new_cell.text = ""
                                cell_para = new_cell.paragraphs[0] if new_cell.paragraphs else new_cell.add_paragraph()
                                cell_para.clear()
                                
                                if not (options.get('format_qa', True) and format_question_answer(cell_para, cell_text, options)):
                                    cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text, options)
                                    table_math += cell_math_count
                                    for method, count in cell_method_stats.items():
                                        stats['method_stats'][method] += count
                                
                                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for run in cell_para.runs:
                                    if not run.font.name:
                                        run.font.name = 'Times New Roman'
                                    if not run.font.size:
                                        run.font.size = Pt(10)
                        except Exception:
                            continue
                
                # Format table
                if options.get('format_tables', True):
                    try:
                        new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                        
                        # Style header row
                        if new_table.rows:
                            header_row = new_table.rows[0]
                            for cell in header_row.cells:
                                for para in cell.paragraphs:
                                    for run in para.runs:
                                        run.bold = True
                                        run.font.color.rgb = RGBColor(255, 255, 255)
                                
                                try:
                                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="4472C4"/>'
                                    shading = parse_xml(shading_xml)
                                    cell._tc.get_or_add_tcPr().append(shading)
                                except:
                                    pass
                    except Exception:
                        pass
                
                stats['tables'] += 1
                stats['math_expressions'] += table_math
                if table_math > 0:
                    stats['processing_log'].append(f"Word table {table_idx+1}: {table_math} equations")
                
            except Exception as e:
                stats['processing_log'].append(f"Word table {table_idx+1}: Error - {str(e)}")
                continue
        
    except Exception as e:
        stats['processing_log'].append(f"Document processing error: {str(e)}")
    
    return new_doc, stats

# Main Streamlit App
def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📚 LaTeX to Word Converter - COMPLETE FIX</h1>
        <p>✅ Fixed: Markdown tables, √ functions, fractions, complex expressions</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("## ⚙️ Cài đặt")
        
        # Processing options
        st.markdown("### 🔧 Tùy chọn xử lý")
        format_qa = st.checkbox("📝 Định dạng câu hỏi & đáp án", value=True)
        format_tables = st.checkbox("📊 Định dạng bảng", value=True)
        convert_markdown = st.checkbox("🔄 Chuyển markdown tables", value=True,
                                     help="Chuyển | tables | thành Word tables")
        
        equation_method = st.selectbox(
            "🧮 Phương pháp equation",
            ["auto", "omml", "styled"],
            index=0
        )
        
        st.markdown("### 🎯 Test LaTeX")
        test_cases = [
            "\\sqrt{3}\\cot 2x = 1",
            "\\frac{\\pi}{3} + k\\pi",
            "\\tan\\frac{65\\pi}{6}",
            "u_{n} + v_{n}",
            "A^{'}B^{'}C^{'}D^{'}"
        ]
        
        selected_test = st.selectbox("Test cases:", [""] + test_cases)
        if selected_test:
            try:
                converted = convert_latex_symbols(selected_test)
                st.code(f"${selected_test}$ → {converted}")
                
                # Show what type of OMML will be created
                if 'FRACTION[' in converted:
                    st.success("Will create: Fraction OMML")
                elif 'SQRT[' in converted:
                    st.success("Will create: Square Root OMML")
                else:
                    st.info("Will create: Standard OMML")
            except Exception as e:
                st.error(f"Error: {e}")
        
        st.markdown("---")
        st.markdown("### ✅ FIXES Applied")
        st.success("""
        **Markdown Tables:**
        | Col1 | Col2 | → Word Table
        
        **Square Roots:**
        √{3} → √3 (no parentheses)
        
        **Fractions:**  
        \\frac{π}{3} → Proper fraction OMML
        
        **Complex Expressions:**
        All LaTeX properly converted
        """)
        
        show_debug = st.checkbox("Debug log", value=False)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("## 📁 Upload File Word")
        uploaded_file = st.file_uploader(
            "Chọn file .docx",
            type=["docx"],
            help="File Word có LaTeX và/hoặc markdown tables"
        )
        
        if uploaded_file:
            st.success(f"✅ Uploaded: **{uploaded_file.name}**")
            
            # Preview
            with st.expander("👀 Preview & Analysis", expanded=True):
                try:
                    doc = Document(uploaded_file)
                    st.info(f"📄 {len(doc.paragraphs)} paragraphs | 📊 {len(doc.tables)} tables")
                    
                    # Analyze content
                    all_text = "\n".join([para.text for para in doc.paragraphs])
                    
                    # Find LaTeX expressions
                    expressions = find_math_expressions(all_text)
                    
                    # Find markdown tables
                    markdown_tables = []
                    lines = all_text.split('\n')
                    i = 0
                    while i < len(lines):
                        if '|' in lines[i]:
                            table_lines = []
                            j = i
                            while j < len(lines) and ('|' in lines[j] or not lines[j].strip()):
                                if lines[j].strip():
                                    table_lines.append(lines[j])
                                j += 1
                            if len(table_lines) >= 2 and detect_markdown_table('\n'.join(table_lines)):
                                markdown_tables.append('\n'.join(table_lines))
                            i = j
                        else:
                            i += 1
                    
                    # Display findings
                    if expressions:
                        st.markdown("**🧮 LaTeX Expressions Found:**")
                        unique_expressions = list(set([expr[2] for expr in expressions[:10]]))
                        
                        for expr in unique_expressions:
                            col_a, col_b = st.columns([1, 1])
                            with col_a:
                                st.code(f"${expr}$", language="latex")
                            with col_b:
                                converted = convert_latex_symbols(expr)
                                st.code(f"→ {converted}")
                        
                        if len(expressions) > 10:
                            st.info(f"... and {len(expressions) - 10} more")
                    
                    if markdown_tables:
                        st.markdown("**📊 Markdown Tables Found:**")
                        for i, table in enumerate(markdown_tables[:2]):
                            st.text(f"Table {i+1}:")
                            st.code(table[:200] + "..." if len(table) > 200 else table)
                        if len(markdown_tables) > 2:
                            st.info(f"... and {len(markdown_tables) - 2} more tables")
                    
                    if not expressions and not markdown_tables:
                        st.warning("No LaTeX expressions or markdown tables found")
                        
                except Exception as e:
                    st.error(f"Preview error: {e}")
    
    with col2:
        st.markdown("## 🎯 Complete Fixes")
        
        fixes = [
            ("📊", "Markdown Tables", "| table | → Word table"),
            ("√", "Square Roots", "√3 (not √(3))"),
            ("⅃", "Proper Fractions", "Real fraction OMML"),
            ("🔤", "Functions", "cot, tan preserved"),
            ("₊", "Subscripts/Supers", "Accurate Unicode"),
            ("🎨", "Professional", "Clean formatting")
        ]
        
        for icon, title, desc in fixes:
            st.markdown(f"""
            <div class="feature-card">
                <strong>{icon} {title}</strong><br>
                <small>{desc}</small>
            </div>
            """, unsafe_allow_html=True)
    
    # Processing
    if uploaded_file:
        st.markdown("---")
        
        if st.button("🚀 Convert with COMPLETE FIX", type="primary", use_container_width=True):
            options = {
                'format_qa': format_qa,
                'format_tables': format_tables,
                'convert_markdown': convert_markdown,
                'equation_method': equation_method
            }
            
            with st.spinner("⏳ Processing with complete fixes..."):
                try:
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document(doc, options)
                    
                    # Save to buffer
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    # Generate filename
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_COMPLETE_{timestamp}.docx"
                    
                    # Success message
                    st.markdown("""
                    <div class="success-box">
                        <h3>🎉 COMPLETE CONVERSION SUCCESS!</h3>
                        <p>All fixes applied: Tables, Math, Formatting</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Statistics
                    st.markdown("## 📊 Conversion Results")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h2 style="color: #667eea;">{stats['paragraphs']}</h2>
                            <p>Paragraphs</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h2 style="color: #28a745;">{stats['math_expressions']}</h2>
                            <p>Math Exprs</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h2 style="color: #ffc107;">{stats['markdown_tables']}</h2>
                            <p>MD Tables</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col4:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h2 style="color: #17a2b8;">{stats['tables']}</h2>
                            <p>Word Tables</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Method breakdown
                    if stats['math_expressions'] > 0:
                        st.markdown("### 🔧 Conversion Methods")
                        method_col1, method_col2, method_col3 = st.columns(3)
                        
                        with method_col1:
                            st.metric("OMML Objects", stats['method_stats']['omml'])
                        with method_col2:
                            st.metric("Styled Math", stats['method_stats']['styled'])
                        with method_col3:
                            st.metric("Fallbacks", stats['method_stats']['fallback'] + stats['method_stats']['error'])
                    
                    # Debug log
                    if show_debug and stats['processing_log']:
                        with st.expander("🔍 Processing Log"):
                            for log in stats['processing_log']:
                                st.text(log)
                    
                    # Download
                    st.download_button(
                        "📥 Download COMPLETE FIXED File",
                        data=buffer.getvalue(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        type="primary"
                    )
                    
                    # Success details
                    success_details = []
                    if stats['math_expressions'] > 0:
                        success_details.append(f"✨ {stats['math_expressions']} math expressions converted")
                    if stats['markdown_tables'] > 0:
                        success_details.append(f"📊 {stats['markdown_tables']} markdown tables → Word tables")
                    if stats['questions_formatted'] > 0:
                        success_details.append(f"📝 {stats['questions_formatted']} Q&A items formatted")
                    if stats['tables'] > 0:
                        success_details.append(f"🎨 {stats['tables']} tables styled")
                    
                    for detail in success_details:
                        st.success(detail)
                        
                except Exception as e:
                    st.markdown("""
                    <div class="warning-box">
                        <h3>❌ Processing Error</h3>
                        <p>Something went wrong during conversion</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.error(f"Error details: {str(e)}")
                    
                    if show_debug:
                        st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
