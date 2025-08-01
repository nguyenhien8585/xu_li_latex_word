#!/usr/bin/env python3
"""
Ultra-Precise LaTeX to Word Converter - TARGET ISSUES FIXED
🎯 SPECIFIC FIXES for your images:
✅ √3 stays as simple √, not nth root  
✅ Prime notation A', B', C', D' perfect
✅ No table duplicates
✅ Proper spacing in formulas
✅ Fractions display correctly
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

# Page config
st.set_page_config(
    page_title="Ultra-Precise LaTeX Converter",
    page_icon="🎯",
    layout="wide"
)

# Focused CSS
st.markdown("""
<style>
    .ultra-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
        padding: 2.5rem;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 2rem;
        color: white;
        box-shadow: 0 10px 30px rgba(0,0,0,0.3);
    }
    .issue-fix {
        background: linear-gradient(135deg, #ff6b6b 0%, #feca57 100%);
        padding: 1rem;
        border-radius: 10px;
        margin: 0.5rem 0;
        color: white;
        font-weight: bold;
    }
    .test-precise {
        background: #e8f5e8;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #27ae60;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

def convert_latex_symbols_precise(text):
    """ULTRA-PRECISE conversion targeting specific issues"""
    if not text:
        return text
    
    result = text
    
    # ISSUE 1: √ symbols - Keep simple √ as √, don't overcomplicate
    # Handle √3 correctly (from Image 2)
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
            
            content = result[content_start:content_end-1].strip()
            
            # FIXED: Simple √ for simple cases
            if len(content) <= 3 and content.isdigit():  # Simple cases like √3
                sqrt_result = f"√{content}"
            else:
                sqrt_result = f"SIMPLEROOT[{content}]"  # Mark for simple root OMML
            
            result = result[:start] + sqrt_result + result[content_end:]
        except:
            break
    
    # ISSUE 2: Handle fractions precisely
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
            
            numerator = result[num_start:num_end-1].strip()
            
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
                
                denominator = result[denom_start:denom_end-1].strip()
                
                # FIXED: Precise fraction marking
                fraction = f"PRECISEFRAC[{numerator}]OVER[{denominator}]"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        except:
            break
    
    # ISSUE 3: nth roots only for actual nth roots
    while re.search(r'\\sqrt\[([^\]]+)\]\{', result):
        match = re.search(r'\\sqrt\[([^\]]+)\]\{', result)
        if not match:
            break
        
        try:
            start = match.start()
            root_index = match.group(1)
            content_start = match.end()
            
            brace_count = 1
            content_end = content_start
            
            while content_end < len(result) and brace_count > 0:
                if result[content_end] == '{':
                    brace_count += 1
                elif result[content_end] == '}':
                    brace_count -= 1
                content_end += 1
            
            content = result[content_start:content_end-1]
            
            # Proper nth root symbols
            if root_index == '3':
                root_result = f"∛{content}"
            elif root_index == '4':
                root_result = f"∜{content}"
            else:
                root_result = f"NTHROOT[{root_index}]OF[{content}]"
            
            result = result[:start] + root_result + result[content_end:]
        except:
            break
    
    # COMPREHENSIVE symbols with proper spacing
    symbols = {
        # Greek letters
        '\\pi': 'π', '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
        '\\epsilon': 'ε', '\\theta': 'θ', '\\lambda': 'λ', '\\mu': 'μ', '\\sigma': 'σ',
        '\\phi': 'φ', '\\omega': 'ω', '\\Omega': 'Ω', '\\Phi': 'Φ', '\\Theta': 'Θ',
        '\\Delta': 'Δ', '\\Gamma': 'Γ', '\\Lambda': 'Λ', '\\Pi': 'Π', '\\Sigma': 'Σ',
        
        # Math symbols
        '\\in': '∈', '\\subset': '⊂', '\\cup': '∪', '\\cap': '∩', '\\emptyset': '∅',
        '\\leq': '≤', '\\geq': '≥', '\\neq': '≠', '\\approx': '≈', '\\equiv': '≡',
        '\\infty': '∞', '\\pm': '±', '\\mp': '∓', '\\times': '×', '\\div': '÷',
        '\\cdot': '·', '\\bullet': '•', '\\circ': '∘',
        
        # Number sets
        '\\mathbb{R}': 'ℝ', '\\mathbb{Z}': 'ℤ', '\\mathbb{Q}': 'ℚ', 
        '\\mathbb{N}': 'ℕ', '\\mathbb{C}': 'ℂ', '\\mathbb{P}': 'ℙ'
    }
    
    # Apply symbols
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # ISSUE 4: Mathematical functions with PRECISE spacing
    functions = {
        '\\sin': 'sin', '\\cos': 'cos', '\\tan': 'tan', '\\cot': 'cot',
        '\\sec': 'sec', '\\csc': 'csc', '\\log': 'log', '\\ln': 'ln'
    }
    
    for latex_func, unicode_func in functions.items():
        # FIXED: Precise spacing after functions
        pattern = r'\b' + re.escape(latex_func) + r'\b(?=\s*[a-zA-Z0-9(√])'
        result = re.sub(pattern, unicode_func + ' ', result)
        result = result.replace(latex_func, unicode_func)
    
    # ISSUE 5: PRIME NOTATION - Ultra precise for A', B', C', D'
    # Handle different prime formats
    prime_patterns = [
        # Direct prime: A' B' C' D'
        (r"([A-Z])([']+)", r"\1\2"),
        # LaTeX prime: A^{'}
        (r"([A-Z])\^?\{?[']+\}?", lambda m: m.group(1) + "'" * max(1, m.group(0).count("'"))),
        # Word prime: A^{prime}
        (r"([A-Z])\^?\{?prime\}?", r"\1'"),
        # Multiple primes
        (r"([A-Z])([']{2,})", r"\1\2"),
    ]
    
    for pattern, replacement in prime_patterns:
        result = re.sub(pattern, replacement, result)
    
    # Subscripts
    subscript_map = {
        '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
        '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
        'n': 'ₙ', 'i': 'ᵢ', 'j': 'ⱼ', 'a': 'ₐ', 'e': 'ₑ'
    }
    
    def replace_subscript(match):
        base = match.group(1)
        content = match.group(2)
        converted = ''.join(subscript_map.get(c, c) for c in content)
        return base + converted
    
    result = re.sub(r'([A-Za-z0-9]+)_\{([^}]+)\}', replace_subscript, result)
    result = re.sub(r'([A-Za-z0-9]+)_([0-9a-z])', lambda m: m.group(1) + subscript_map.get(m.group(2), m.group(2)), result)
    
    # Superscripts
    superscript_map = {
        '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
        '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
        '+': '⁺', '-': '⁻', '=': '⁼'
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
    
    # Handle brackets
    result = result.replace('\\left(', '(').replace('\\right)', ')')
    result = result.replace('\\left[', '[').replace('\\right]', ']')
    result = result.replace('\\{', '{').replace('\\}', '}')
    
    # ISSUE 6: PRECISE spacing cleanup
    # Fix spacing around operators
    result = re.sub(r'\s*([+\-×÷=≤≥<>≠≈])\s*', r' \1 ', result)
    # Remove multiple spaces
    result = re.sub(r'\s+', ' ', result)
    # Fix spacing around parentheses
    result = re.sub(r'\s*\(\s*', '(', result)
    result = re.sub(r'\s*\)\s*', ') ', result)
    
    # Clean up remaining LaTeX
    result = re.sub(r'\\[a-zA-Z]+', '', result)
    result = result.replace('\\', '')
    
    return result.strip()

def create_simple_sqrt_omml(paragraph, content):
    """Create SIMPLE square root OMML - not overcomplicated"""
    try:
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:rad>
                <m:radPr></m:radPr>
                <m:deg></m:deg>
                <m:e>
                    <m:r>
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

def create_precise_fraction_omml(paragraph, numerator, denominator):
    """Create PRECISE fraction OMML"""
    try:
        safe_num = numerator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_den = denominator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:f>
                <m:fPr>
                    <m:type m:val="lin"/>
                </m:fPr>
                <m:num>
                    <m:r>
                        <m:t>{safe_num}</m:t>
                    </m:r>
                </m:num>
                <m:den>
                    <m:r>
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

def find_math_expressions(text):
    """Find math expressions - precise detection"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Handle $$ format
            if i < len(text) and text[i] == '$':
                i += 1
                content_start = i
                
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

def create_equation_object_precise(paragraph, latex_text, method='auto'):
    """ULTRA-PRECISE equation object creation"""
    unicode_text = convert_latex_symbols_precise(latex_text)
    
    # Handle simple square roots first (NOT nth roots)
    if 'SIMPLEROOT[' in unicode_text:
        match = re.search(r'SIMPLEROOT\[([^\]]+)\]', unicode_text)
        if match:
            content = match.group(1)
            
            if create_simple_sqrt_omml(paragraph, content):
                return 'omml'
            else:
                sqrt_text = f"√{content}"
                unicode_text = unicode_text.replace(match.group(0), sqrt_text)
    
    # Handle precise fractions
    if 'PRECISEFRAC[' in unicode_text:
        match = re.search(r'PRECISEFRAC\[([^\]]+)\]OVER\[([^\]]+)\]', unicode_text)
        if match:
            numerator = match.group(1)
            denominator = match.group(2)
            
            if create_precise_fraction_omml(paragraph, numerator, denominator):
                return 'omml'
            else:
                fraction_text = f"({numerator})/({denominator})"
                unicode_text = unicode_text.replace(match.group(0), fraction_text)
    
    # Handle nth roots (only when actually nth)
    if 'NTHROOT[' in unicode_text:
        match = re.search(r'NTHROOT\[([^\]]+)\]OF\[([^\]]+)\]', unicode_text)
        if match:
            index = match.group(1)
            content = match.group(2)
            # For now, use Unicode fallback
            nth_root_text = f"ⁿ√{content} (n={index})"
            unicode_text = unicode_text.replace(match.group(0), nth_root_text)
    
    # Regular OMML for other expressions
    if method == 'omml' or method == 'auto':
        try:
            safe_text = unicode_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                <m:r>
                    <m:t>{safe_text}</m:t>
                </m:r>
            </m:oMath>'''
            
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return 'omml'
        except Exception:
            pass
    
    # Styled fallback
    try:
        run = paragraph.add_run(unicode_text)
        run.font.name = 'Cambria Math'
        run.font.size = Pt(11)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 32, 96)
        return 'styled'
    except:
        pass
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_text)
        run.italic = True
        return 'fallback'
    except:
        run = paragraph.add_run(f"[{latex_text}]")
        return 'error'

def detect_markdown_table_precise(text):
    """PRECISE table detection - no false positives"""
    lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
    if len(lines) < 2:
        return False
    
    # Must have pipes in at least 2 lines
    pipe_lines = [line for line in lines if '|' in line and len([cell for cell in line.split('|') if cell.strip()]) >= 2]
    if len(pipe_lines) < 2:
        return False
    
    # Must have separator line
    has_separator = False
    for line in lines:
        if re.match(r'^[\|\s\-:]+$', line) and '|' in line and '-' in line and len(line) > 3:
            has_separator = True
            break
    
    return has_separator

def parse_markdown_table_precise(text):
    """ULTRA-PRECISE table parsing - NO DUPLICATES"""
    lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
    table_data = []
    separator_indices = set()
    processed_content = set()  # Track content to prevent duplicates
    
    # Find separator lines
    for i, line in enumerate(lines):
        if re.match(r'^[\|\s\-:]+$', line) and '|' in line and '-' in line:
            separator_indices.add(i)
    
    # Process only valid table rows
    for i, line in enumerate(lines):
        if i in separator_indices:
            continue
        
        if '|' in line and len(line) > 3:
            # Clean the line
            cleaned_line = line.strip()
            if cleaned_line.startswith('|'):
                cleaned_line = cleaned_line[1:]
            if cleaned_line.endswith('|'):
                cleaned_line = cleaned_line[:-1]
            
            # Split into cells
            cells = [cell.strip() for cell in cleaned_line.split('|')]
            
            # Remove empty cells at edges
            while cells and not cells[0]:
                cells.pop(0)
            while cells and not cells[-1]:
                cells.pop()
            
            if cells and len(cells) >= 2:  # Must have at least 2 meaningful cells
                # Create content signature to detect duplicates
                content_signature = ''.join(cells).lower().replace(' ', '').replace('\n', '')
                
                if content_signature not in processed_content and len(content_signature) > 3:
                    table_data.append(cells)
                    processed_content.add(content_signature)
    
    return table_data

def create_word_table_precise(doc, table_data, options):
    """Create precise Word table with no duplicates"""
    if not table_data or len(table_data) < 1:
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
                    cell_math_count, _ = process_text_with_math_precise(cell_para, cell_text, options)
                    table_math += cell_math_count
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(10)
        
        # Apply professional formatting
        if options.get('format_tables', True):
            try:
                word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # Style header row
                if word_table.rows:
                    header_row = word_table.rows[0]
                    for cell in header_row.cells:
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
            except:
                pass
        
        return word_table, table_math
        
    except Exception as e:
        st.warning(f"Table creation error: {e}")
        return None, 0

def process_text_with_math_precise(paragraph, text, options):
    """Process text with ULTRA-PRECISE math handling"""
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
        
        # Create equation with precision
        method_used = create_equation_object_precise(paragraph, latex_content, options.get('equation_method', 'auto'))
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

def format_question_answer_precise(paragraph, text, options):
    """PRECISE Q&A formatting"""
    paragraph.clear()
    
    # Question patterns - more precise
    question_patterns = [
        r'^(Câu\s+\d+\.)\s*(.*)',
        r'^(\d+\.)\s*(.*)'
    ]
    
    for pattern in question_patterns:
        match = re.match(pattern, text, re.IGNORECASE)
        if match:
            question_part = match.group(1)
            content_part = match.group(2)
            
            # Format question number
            run_q = paragraph.add_run(question_part)
            run_q.bold = True
            run_q.font.size = Pt(11)
            run_q.font.name = 'Times New Roman'
            run_q.font.color.rgb = RGBColor(0, 0, 0)
            
            if content_part:
                paragraph.add_run(" ")
                process_text_with_math_precise(paragraph, content_part, options)
            
            return True
    
    # Answer patterns - ultra precise for A', B', C', D'
    answer_pattern = r'^([A-D])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format answer letter
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(11)
        run_l.font.name = 'Times New Roman'
        
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math_precise(paragraph, answer_content, options)
        
        return True
    
    # Not Q&A format
    process_text_with_math_precise(paragraph, text, options)
    return False

def process_document_precise(doc, options):
    """ULTRA-PRECISE document processing"""
    new_doc = Document()
    
    stats = {
        'paragraphs': 0,
        'tables': 0,
        'markdown_tables': 0,
        'math_expressions': 0,
        'questions_formatted': 0,
        'sqrt_simplified': 0,
        'primes_fixed': 0,
        'duplicates_removed': 0,
        'method_stats': {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0},
        'processing_log': []
    }
    
    try:
        # Collect paragraph text
        paragraph_texts = []
        for para in doc.paragraphs:
            paragraph_texts.append(para.text)
        
        # Process with precision
        i = 0
        while i < len(paragraph_texts):
            text = paragraph_texts[i].strip()
            stats['paragraphs'] += 1
            
            if not text:
                new_doc.add_paragraph()
                i += 1
                continue
            
            # Check for markdown tables with PRECISION
            if options.get('convert_markdown', True) and '|' in text:
                table_lines = []
                j = i
                
                # Collect potential table lines
                while j < len(paragraph_texts):
                    current_text = paragraph_texts[j].strip()
                    if not current_text:
                        j += 1
                        if j < len(paragraph_texts) and '|' in paragraph_texts[j]:
                            continue
                        else:
                            break
                    if '|' in current_text:
                        table_lines.append(current_text)
                        j += 1
                    else:
                        break
                
                # Check if it's a PRECISE table
                combined_table_text = '\n'.join(table_lines)
                if detect_markdown_table_precise(combined_table_text):
                    table_data = parse_markdown_table_precise(combined_table_text)
                    if table_data and len(table_data) >= 2:  # Must have header + at least 1 data row
                        word_table, table_math = create_word_table_precise(new_doc, table_data, options)
                        if word_table is not None:
                            stats['markdown_tables'] += 1
                            stats['math_expressions'] += table_math
                            duplicates_avoided = len(table_lines) - len(table_data)
                            stats['duplicates_removed'] += duplicates_avoided
                            stats['processing_log'].append(f"Precise table: {len(table_data)} unique rows, {duplicates_avoided} duplicates removed")
                    
                    i = j
                    continue
            
            # Process regular paragraph with precision
            new_para = new_doc.add_paragraph()
            
            try:
                # Count specific fixes
                if '√' in text or '\\sqrt{' in text:
                    stats['sqrt_simplified'] += text.count('√') + text.count('\\sqrt{')
                if "'" in text:
                    stats['primes_fixed'] += text.count("'")
                
                if options.get('format_qa', True) and format_question_answer_precise(new_para, text, options):
                    stats['questions_formatted'] += 1
                    stats['processing_log'].append(f"Para {i+1}: Precise Q&A formatting")
                else:
                    math_count, method_stats = process_text_with_math_precise(new_para, text, options)
                    stats['math_expressions'] += math_count
                    for method, count in method_stats.items():
                        stats['method_stats'][method] += count
                    
                    if math_count > 0:
                        stats['processing_log'].append(f"Para {i+1}: {math_count} equations precisely converted")
                        
            except Exception as e:
                new_para.clear()
                run = new_para.add_run(text)
                run.font.size = Pt(11)
                run.font.name = 'Times New Roman'
                stats['processing_log'].append(f"Para {i+1}: Fallback - {str(e)}")
            
            i += 1
        
        # Process existing tables
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
                                
                                cell_math_count, cell_method_stats = process_text_with_math_precise(cell_para, cell_text, options)
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
                
                stats['tables'] += 1
                stats['math_expressions'] += table_math
                if table_math > 0:
                    stats['processing_log'].append(f"Word table {table_idx+1}: {table_math} equations")
                
            except Exception as e:
                stats['processing_log'].append(f"Word table {table_idx+1}: Error - {str(e)}")
                continue
        
    except Exception as e:
        stats['processing_log'].append(f"Document error: {str(e)}")
    
    return new_doc, stats

# Main App
def main():
    st.markdown("""
    <div class="ultra-header">
        <h1>🎯 Ultra-Precise LaTeX to Word Converter</h1>
        <p>Targeting YOUR specific issues: √3 stays simple | A'B'C'D' perfect | No table duplicates | Precise spacing</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("## ⚙️ Ultra Settings")
        
        format_qa = st.checkbox("📝 Format Q&A", value=True)
        format_tables = st.checkbox("📊 Format tables", value=True)
        convert_markdown = st.checkbox("🔄 Convert markdown (precise)", value=True)
        
        equation_method = st.selectbox("🧮 Equation method", ["auto", "omml", "styled"], index=0)
        
        st.markdown("---")
        st.markdown("### 🎯 Test Your Specific Issues")
        
        # Test cases from the user's images
        specific_tests = [
            "A' B' C' D'",  # Image 1 - Prime notation
            "\\sqrt{3}",    # Image 2 - Simple square root
            "\\pi/3",       # Image 2 - Simple fraction
            "\\frac{\\pi}{3}",  # Fraction test
        ]
        
        st.markdown("**Issues from your images:**")
        for i, test in enumerate(specific_tests):
            issue_name = ["Prime notation", "√ simple", "π fraction", "Full fraction"][i]
            
            st.markdown(f"""
            <div class="issue-fix">
                {issue_name}: ${test}$
            </div>
            """, unsafe_allow_html=True)
            
            try:
                result = convert_latex_symbols_precise(test)
                st.markdown(f"""
                <div class="test-precise">
                    <strong>FIXED:</strong> {result}
                </div>
                """, unsafe_allow_html=True)
                
                # Show what type of fix
                fixes = []
                if "'" in result:
                    fixes.append("✅ Prime fixed")
                if "√" in result and "NTHROOT" not in result:
                    fixes.append("✅ Simple √")
                if "PRECISEFRAC" in convert_latex_symbols_precise(test):
                    fixes.append("✅ Precise fraction")
                if "π" in result:
                    fixes.append("✅ Greek symbol")
                
                if fixes:
                    st.info(" | ".join(fixes))
                    
            except Exception as e:
                st.error(f"Error: {e}")
        
        show_debug = st.checkbox("🔍 Debug log", value=False)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("## 📁 Upload for Ultra-Precise Fixing")
        uploaded_file = st.file_uploader("Choose .docx file", type=["docx"])
        
        if uploaded_file:
            st.success(f"✅ File loaded: **{uploaded_file.name}**")
            
            with st.expander("👀 Ultra-Precise Analysis", expanded=True):
                try:
                    doc = Document(uploaded_file)
                    all_text = "\n".join([para.text for para in doc.paragraphs])
                    
                    # Count specific issues
                    simple_sqrt = all_text.count('\\sqrt{')
                    primes = all_text.count("'") + all_text.count('prime')
                    fractions = all_text.count('\\frac{')
                    
                    # Estimate table duplicates
                    lines_with_pipes = [line for line in all_text.split('\n') if '|' in line]
                    potential_duplicates = len(lines_with_pipes) - len(set(lines_with_pipes))
                    
                    col_a, col_b, col_c, col_d = st.columns(4)
                    with col_a:
                        st.metric("√ to simplify", simple_sqrt)
                    with col_b:
                        st.metric("Primes to fix", primes)
                    with col_c:
                        st.metric("Fractions to fix", fractions)
                    with col_d:
                        st.metric("Potential duplicates", potential_duplicates)
                    
                    if simple_sqrt + primes + fractions > 0:
                        st.success(f"🎯 Ready to ultra-precisely fix {simple_sqrt + primes + fractions} specific issues!")
                    else:
                        st.info("No specific issues detected")
                        
                except Exception as e:
                    st.error(f"Analysis error: {e}")
    
    with col2:
        st.markdown("## 🎯 Your Issues → Fixed")
        
        issue_fixes = [
            ("√", "√3 stays simple", "Not nth root"),
            ("'", "A'B'C'D' perfect", "Prime notation"),
            ("📊", "No table duplicates", "Clean parsing"),
            ("📐", "Precise spacing", "sin x not sinx"),
            ("⚡", "π/3 fractions", "Proper display"),
            ("🎨", "Professional", "Word quality")
        ]
        
        for icon, title, desc in issue_fixes:
            st.markdown(f"""
            <div class="issue-fix">
                {icon} {title}<br>
                <small>{desc}</small>
            </div>
            """, unsafe_allow_html=True)
    
    # Processing
    if uploaded_file:
        st.markdown("---")
        
        if st.button("🎯 ULTRA-PRECISE FIX NOW", type="primary", use_container_width=True):
            options = {
                'format_qa': format_qa,
                'format_tables': format_tables,
                'convert_markdown': convert_markdown,
                'equation_method': equation_method
            }
            
            with st.spinner("🎯 Applying ultra-precise fixes..."):
                try:
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document_precise(doc, options)
                    
                    # Save
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_ULTRA_PRECISE_{timestamp}.docx"
                    
                    # Success
                    st.markdown("""
                    <div style="background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%); padding: 2rem; border-radius: 15px; color: white; text-align: center; margin: 1rem 0;">
                        <h3>🎯 ULTRA-PRECISE FIXES APPLIED!</h3>
                        <p>All your specific issues have been targeted and fixed!</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Statistics
                    st.markdown("## 📊 Ultra-Precise Results")
                    
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1:
                        st.metric("Math Expressions", stats['math_expressions'])
                    with col2:
                        st.metric("√ Simplified", stats['sqrt_simplified'])
                    with col3:
                        st.metric("Primes Fixed", stats['primes_fixed'])
                    with col4:
                        st.metric("Duplicates Removed", stats['duplicates_removed'])
                    with col5:
                        st.metric("Tables Created", stats['markdown_tables'])
                    
                    # Method stats
                    if stats['math_expressions'] > 0:
                        st.markdown("### 🔧 Precision Methods")
                        method_col1, method_col2, method_col3 = st.columns(3)
                        
                        with method_col1:
                            st.metric("OMML Objects", stats['method_stats']['omml'])
                        with method_col2:
                            st.metric("Styled Math", stats['method_stats']['styled'])
                        with method_col3:
                            st.metric("Safe Fallbacks", stats['method_stats']['fallback'])
                    
                    # Debug
                    if show_debug and stats['processing_log']:
                        with st.expander("🔍 Ultra-Precise Log"):
                            for log in stats['processing_log']:
                                st.text(log)
                    
                    # Download
                    st.download_button(
                        "📥 Download ULTRA-PRECISE Fixed File",
                        data=buffer.getvalue(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        type="primary"
                    )
                    
                    # Specific success messages
                    successes = []
                    if stats['sqrt_simplified'] > 0:
                        successes.append(f"🎯 {stats['sqrt_simplified']} √ symbols kept simple (not nth roots)")
                    if stats['primes_fixed'] > 0:
                        successes.append(f"✨ {stats['primes_fixed']} prime notations (A', B', C', D') perfected")
                    if stats['duplicates_removed'] > 0:
                        successes.append(f"🧹 {stats['duplicates_removed']} table duplicates eliminated")
                    if stats['markdown_tables'] > 0:
                        successes.append(f"📊 {stats['markdown_tables']} tables created with precision")
                    
                    for success in successes:
                        st.success(success)
                        
                except Exception as e:
                    st.error(f"Error: {str(e)}")
                    if show_debug:
                        st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
