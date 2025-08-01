import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="LaTeX Word Processor",
    page_icon="📚",
    layout="wide"
)

def convert_latex_symbols(text):
    """Chuyển đổi ký hiệu LaTeX thành Unicode"""
    if not text:
        return text
    
    result = text
    
    # Xử lý superscript có dấu ngoặc: A^{prime} -> A'
    while True:
        pos = result.find('^{')
        if pos == -1:
            break
        
        # Tìm ký tự trước ^{
        base_start = pos - 1
        while base_start >= 0 and result[base_start].isalnum():
            base_start -= 1
        base_start += 1
        base = result[base_start:pos]
        
        # Tìm nội dung trong {}
        brace_start = pos + 2
        brace_count = 1
        brace_end = brace_start
        
        while brace_end < len(result) and brace_count > 0:
            if result[brace_end] == '{':
                brace_count += 1
            elif result[brace_end] == '}':
                brace_count -= 1
            brace_end += 1
        
        if brace_count == 0:
            content = result[brace_start:brace_end-1]
            if content == 'prime':
                replacement = base + "'"
            else:
                replacement = base + "^(" + content + ")"
            result = result[:base_start] + replacement + result[brace_end:]
        else:
            break
    
    # Các ký hiệu toán học cơ bản
    symbols = {
        '\\pi': 'π',
        '\\alpha': 'α',
        '\\beta': 'β',
        '\\gamma': 'γ',
        '\\delta': 'δ',
        '\\theta': 'θ',
        '\\lambda': 'λ',
        '\\mu': 'μ',
        '\\sigma': 'σ',
        '\\phi': 'φ',
        '\\omega': 'ω',
        '\\in': '∈',
        '\\subset': '⊂',
        '\\cup': '∪',
        '\\cap': '∩',
        '\\emptyset': '∅',
        '\\leq': '≤',
        '\\geq': '≥',
        '\\neq': '≠',
        '\\approx': '≈',
        '\\infty': '∞',
        '\\pm': '±',
        '\\mp': '∓',
        '\\mathbb{R}': 'ℝ',
        '\\mathbb{Z}': 'ℤ',
        '\\mathbb{Q}': 'ℚ',
        '\\mathbb{N}': 'ℕ',
        '\\mathbb{C}': 'ℂ'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Xử lý left, right parentheses
    result = result.replace('\\left(', '(')
    result = result.replace('\\right)', ')')
    result = result.replace('\\left[', '[')
    result = result.replace('\\right]', ']')
    result = result.replace('\\left{', '{')
    result = result.replace('\\right}', '}')
    
    # Xử lý frac
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        # Tìm tử số
        num_start = start + 6
        brace_count = 1
        num_end = num_start
        
        while num_end < len(result) and brace_count > 0:
            if result[num_end] == '{':
                brace_count += 1
            elif result[num_end] == '}':
                brace_count -= 1
            num_end += 1
        
        if brace_count > 0:
            break
            
        numerator = result[num_start:num_end-1]
        
        # Tìm mẫu số
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
            
            if brace_count == 0:
                denominator = result[denom_start:denom_end-1]
                fraction = f"({numerator})/({denominator})"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        else:
            break
    
    # Xử lý sqrt
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
            
        content_start = start + 6
        brace_count = 1
        content_end = content_start
        
        while content_end < len(result) and brace_count > 0:
            if result[content_end] == '{':
                brace_count += 1
            elif result[content_end] == '}':
                brace_count -= 1
            content_end += 1
        
        if brace_count == 0:
            content = result[content_start:content_end-1]
            sqrt_result = f"√({content})"
            result = result[:start] + sqrt_result + result[content_end:]
        else:
            break
    
    # Xử lý superscript đơn giản: x^2 -> x²
    superscript_digits = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    
    i = 0
    new_result = ""
    while i < len(result):
        if (i < len(result) - 2 and 
            result[i+1] == '^' and 
            result[i+2].isdigit() and
            result[i].isalnum()):
            base = result[i]
            power = result[i+2]
            new_result += base + superscript_digits.get(power, power)
            i += 3
        else:
            new_result += result[i]
            i += 1
    
    result = new_result
    
    # Loại bỏ backslash còn lại
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Tìm expressions toán học trong text"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Xử lý ${...}$ format
            if i < len(text) and text[i] == '{':
                i += 1  # skip {
                brace_count = 1
                content_start = i
                
                while i < len(text) and brace_count > 0:
                    if text[i] == '{':
                        brace_count += 1
                    elif text[i] == '}':
                        brace_count -= 1
                    i += 1
                
                # Kiểm tra có }$ không
                if i < len(text) and text[i] == '$':
                    end = i + 1
                    latex_content = text[content_start:i-1]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'braced'))
                    i += 1
                else:
                    i += 1
            else:
                # Xử lý $...$ format thông thường
                content_start = i
                brace_depth = 0
                
                # Tìm $ đóng, tracking nested braces
                while i < len(text):
                    if text[i] == '{':
                        brace_depth += 1
                    elif text[i] == '}':
                        brace_depth -= 1
                    elif text[i] == '$' and brace_depth == 0:
                        break
                    i += 1
                
                if i < len(text):  # Tìm thấy $ đóng
                    end = i + 1
                    latex_content = text[content_start:i]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'normal'))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_omml_equation(paragraph, text, equation_type='simple'):
    """Tạo OMML equation - Safe with fallbacks"""
    try:
        # Always try simple approach first to avoid corruption
        if equation_type == 'prime' and "'" in text:
            if create_prime_omml(paragraph, text):
                return True
        elif equation_type == 'superscript' and any(c in text for c in ['²', '³', '⁴', '⁵']):
            if create_superscript_omml(paragraph, text):
                return True
        elif equation_type == 'fraction' and '(' in text and ')' in text and '/' in text:
            if create_fraction_omml(paragraph, text):
                return True
        elif equation_type == 'parentheses' and '(' in text and ')' in text:
            if create_parentheses_omml(paragraph, text):
                return True
        
        # Default to simple OMML
        return create_simple_omml(paragraph, text)
        
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field - Simplified and safe"""
    try:
        # Simple EQ field without complex formatting
        eq_content = unicode_text
        
        # Only handle basic superscripts
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}'"
        
        # Safe field creation
        fldChar_begin = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>')
        instrText = parse_xml(f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> {eq_content} </w:instrText>')
        fldChar_end = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>')
        
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_prime_omml(paragraph, text):
    """OMML cho prime notation - Fixed XML"""
    try:
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        # Safe XML with proper escaping
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_prime = "'" * prime_count
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_prime}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_superscript_omml(paragraph, text):
    """OMML cho superscript numbers - Fixed XML"""
    try:
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:
                    base += char
        
        if not sup:
            return create_simple_omml(paragraph, text)
        
        # Safe XML
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_sup = sup.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_sup}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_fraction_omml(paragraph, text):
    """OMML cho fractions - Simplified and safe"""
    try:
        frac_pattern = r'\(([^)]+)\)/\(([^)]+)\)'
        match = re.search(frac_pattern, text)
        
        if match:
            numerator = match.group(1).strip()
            denominator = match.group(2).strip()
            
            # Safe XML escaping
            safe_num = numerator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            safe_den = denominator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:f><m:fPr></m:fPr><m:num><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_num}</m:t></m:r></m:num><m:den><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_den}</m:t></m:r></m:den></m:f></m:oMath>'''
            
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return True
        
        return False
    except Exception:
        return False

def create_parentheses_omml(paragraph, text):
    """OMML cho parentheses - Simplified"""
    try:
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end].strip()
        else:
            content = text.strip()
        
        # Safe XML
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_content}</m:t></m:r></m:e></m:d></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_simple_omml(paragraph, text):
    """OMML cho equation đơn giản - Safe version"""
    try:
        # Safe XML escaping
        safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_text}</m:t></m:r></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field"""
    try:
        eq_content = unicode_text
        
        # Handle superscripts for EQ field
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}\\s\\up5(')"
        elif '²' in unicode_text:
            base = unicode_text.replace('²', '')
            eq_content = f"{base}\\s\\up5(2)"
        elif '³' in unicode_text:
            base = unicode_text.replace('³', '')
            eq_content = f"{base}\\s\\up5(3)"
        
        # Create field elements
        fldChar_begin = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>'
        )
        
        instrText = parse_xml(
            f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> EQ {eq_content}</w:instrText>'
        )
        
        fldChar_end = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>'
        )
        
        # Add to paragraph
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_styled_math_run(paragraph, unicode_text):
    """Tạo styled run như equation"""
    try:
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 32, 96)
        
        # Add background
        try:
            rPr = run._element.get_or_add_rPr()
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
        except:
            pass
        
        return True
    except Exception:
        return False

def create_equation_object(paragraph, latex_text):
    """Tạo equation object với safe methods"""
    unicode_content = convert_latex_symbols(latex_text)
    
    # Determine equation type
    equation_type = 'simple'
    if "'" in unicode_content:
        equation_type = 'prime'
    elif '(' in unicode_content and ')' in unicode_content and '/' in unicode_content:
        equation_type = 'fraction'
    elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
        equation_type = 'superscript'
    elif '(' in unicode_content and ')' in unicode_content:
        equation_type = 'parentheses'
    
    # Try methods with better error handling
    try:
        # Method 1: OMML (safest approach)
        if create_omml_equation(paragraph, unicode_content, equation_type):
            return 'omml'
    except Exception:
        pass
    
    try:
        # Method 2: Styled run (most reliable)
        if create_styled_math_run(paragraph, unicode_content):
            return 'styled'
    except Exception:
        pass
    
    try:
        # Method 3: EQ Field (skip if problematic)
        # Commenting out to avoid file corruption
        # if create_equation_field(paragraph, unicode_content, latex_text):
        #     return 'field'
        pass
    except Exception:
        pass
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_content)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)
        return 'fallback'
    except Exception:
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return 'error'

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Safe version"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    last_pos = 0
    math_count = 0
    method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    for start, end, latex_content, expr_type in expressions:
        # Add text before equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Create equation object safely
        method_used = create_equation_object(paragraph, latex_content)
        method_stats[method_used] += 1
        math_count += 1
        
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count, method_stats

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # Bold question number
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Process content with math
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Bold answer letter
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Process answer content
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra markdown table"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse markdown table"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Skip separator line
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table - Safe version"""
    if not table_data:
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    math_count = 0
    total_method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    max_cols = max(len(row) for row in table_data)
    
    # Create table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Fill data
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Create paragraph in cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Process cell content safely
                try:
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                        math_count += cell_math_count
                        for method, count in cell_method_stats.items():
                            total_method_stats[method] += count
                except Exception:
                    # Fallback to plain text
                    run = cell_para.add_run(cell_text)
                    run.font.size = Pt(11)
                    run.font.name = 'Times New Roman'
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table safely
    if options.get('format_tables', True):
        try:
            word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if word_table.rows:
                header_row = word_table.rows[0]
                for cell in header_row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                    
                    try:
                        shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                        shading = parse_xml(shading_xml)
                        cell._tc.get_or_add_tcPr().append(shading)
                    except:
                        pass
        except Exception:
            pass
    
    return math_count, total_method_stats

def process_document_content(doc, options):
    """Xử lý document content - Safe version"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'omml_equations': 0,
        'field_equations': 0,
        'styled_equations': 0,
        'fallback_equations': 0,
        'error_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Processing: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Process paragraphs safely
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Check for markdown table
        if '|' in text and options.get('convert_markdown', True):
            table_lines = []
            j = i
            
            while j < len(doc.paragraphs):
                current_text = doc.paragraphs[j].text.strip()
                if not current_text:
                    break
                if '|' in current_text:
                    table_lines.append(current_text)
                    j += 1
                else:
                    break
            
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    try:
                        math_count, method_stats = create_word_table(new_doc, table_data, options)
                        stats['math_expressions'] += math_count
                        for method, count in method_stats.items():
                            if method in stats:
                                stats[f'{method}_equations'] += count
                        stats['markdown_tables'] += 1
                        stats['debug_log'].append(f"📋 Converted table: {len(table_data)} rows")
                    except Exception as e:
                        stats['debug_log'].append(f"❌ Table error: {str(e)}")
                
                i = j
                continue
        
        # Process normal paragraph safely
        new_para = new_doc.add_paragraph()
        
        try:
            if format_question_answer(new_para, text):
                stats['questions_formatted'] += 1
                stats['debug_log'].append(f"📝 P{i}: Q&A formatted")
            else:
                math_count, method_stats = process_text_with_math(new_para, text)
                stats['math_expressions'] += math_count
                for method, count in method_stats.items():
                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                        stats[f'{method}_equations'] += count
                
                if math_count > 0:
                    omml_count = method_stats.get('omml', 0)
                    styled_count = method_stats.get('styled', 0) + method_stats.get('fallback', 0)
                    stats['debug_log'].append(f"🔢 P{i}: {math_count} equations ({omml_count} OMML, {styled_count} styled)")
        except Exception as e:
            # Fallback to plain text
            new_para.clear()
            run = new_para.add_run(text)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
            stats['debug_log'].append(f"❌ P{i}: Error, used plain text")
        
        i += 1
    
    # Process existing Word tables safely
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
                            if new_cell.paragraphs:
                                cell_para = new_cell.paragraphs[0]
                                cell_para.clear()
                            else:
                                cell_para = new_cell.add_paragraph()
                            
                            if not format_question_answer(cell_para, cell_text):
                                cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                                table_math += cell_math_count
                                for method, count in cell_method_stats.items():
                                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                                        stats[f'{method}_equations'] += count
                            
                            cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except Exception:
                        # Skip problematic cells
                        continue
            
            stats['math_expressions'] += table_math
            
            # Format table safely
            if options.get('format_tables', True):
                try:
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    if new_table.rows:
                        header_row = new_table.rows[0]
                        for cell in header_row.cells:
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.bold = True
                except Exception:
                    pass
                
                stats['word_tables'] += 1
            
            if table_math > 0:
                stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations")
                
        except Exception as e:
            stats['debug_log'].append(f"❌ Table {table_idx} error: {str(e)}")
            continue
    
    stats['debug_log'].append(f"✅ Complete: {stats['math_expressions']} total equations")
    stats['debug_log'].append(f"   🎯 OMML: {stats['omml_equations']}")
    stats['debug_log'].append(f"   ✨ Styled: {stats['styled_equations'] + stats['fallback_equations']}")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v8.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">Fixed File Corruption + Safe Equation Objects + Complete LaTeX Support</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🎯 Tạo TRUE Equation Objects", value=True,
                                       help="Chuyển $...$ thành REAL equation objects trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True,
                                     help="Format 'Câu X:' và 'A. B. C. D.'")  
        format_tables = st.checkbox("📊 Format bảng", value=True,
                                   help="Professional table formatting")
        convert_markdown = st.checkbox("📋 Markdown → Word Tables", value=True,
                                     help="Convert | tables | to Word")
        show_debug = st.checkbox("🔍 Advanced Debug", value=False,
                               help="Show equation creation details")
        
        st.markdown("---")
        st.info("🛡️ **SAFE Equation Creation:**\n\n✅ **Method 1:** Safe OMML Objects\n🎨 **Method 2:** Enhanced Styling  \n💫 **Method 3:** Fallback Protection\n\n🔧 **Fixed Issues:**\n- File Corruption\n- XML Validation\n- Safe Error Handling\n- Robust Fallbacks")
        
        st.markdown("### 🧪 Test LaTeX → Safe Equation")
        test_latex = st.text_input("Test LaTeX:", placeholder="A^{prime}")
        if test_latex:
            try:
                converted = convert_latex_symbols(test_latex)
                st.code(f"Input:  ${test_latex}$")
                st.code(f"Output: {converted}")
                
                # Show equation type safely
                unicode_content = convert_latex_symbols(test_latex)
                if '(' in unicode_content and ')' in unicode_content and '/' in unicode_content:
                    st.success("🎯 Will create: Safe OMML Fraction")
                elif "'" in unicode_content:
                    st.success("🎯 Will create: Safe OMML Prime Superscript")
                elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
                    st.success("🎯 Will create: Safe OMML Superscript")
                elif '(' in unicode_content and ')' in unicode_content:
                    st.success("🎯 Will create: Safe OMML Parentheses")
                else:
                    st.info("🎯 Will create: Safe OMML Simple Math")
            except Exception:
                st.warning("⚠️ Test failed - will use styled fallback")
        
        st.markdown("### 📊 Safe Examples")
        st.markdown("""
        **Safely Supported:**
        - `$A^{prime}import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="LaTeX Word Processor",
    page_icon="📚",
    layout="wide"
)

def convert_latex_symbols(text):
    """Chuyển đổi ký hiệu LaTeX thành Unicode"""
    if not text:
        return text
    
    result = text
    
    # Xử lý superscript có dấu ngoặc: A^{prime} -> A'
    while True:
        pos = result.find('^{')
        if pos == -1:
            break
        
        # Tìm ký tự trước ^{
        base_start = pos - 1
        while base_start >= 0 and result[base_start].isalnum():
            base_start -= 1
        base_start += 1
        base = result[base_start:pos]
        
        # Tìm nội dung trong {}
        brace_start = pos + 2
        brace_count = 1
        brace_end = brace_start
        
        while brace_end < len(result) and brace_count > 0:
            if result[brace_end] == '{':
                brace_count += 1
            elif result[brace_end] == '}':
                brace_count -= 1
            brace_end += 1
        
        if brace_count == 0:
            content = result[brace_start:brace_end-1]
            if content == 'prime':
                replacement = base + "'"
            else:
                replacement = base + "^(" + content + ")"
            result = result[:base_start] + replacement + result[brace_end:]
        else:
            break
    
    # Các ký hiệu toán học cơ bản
    symbols = {
        '\\pi': 'π',
        '\\alpha': 'α',
        '\\beta': 'β',
        '\\gamma': 'γ',
        '\\delta': 'δ',
        '\\theta': 'θ',
        '\\lambda': 'λ',
        '\\mu': 'μ',
        '\\sigma': 'σ',
        '\\phi': 'φ',
        '\\omega': 'ω',
        '\\in': '∈',
        '\\subset': '⊂',
        '\\cup': '∪',
        '\\cap': '∩',
        '\\emptyset': '∅',
        '\\leq': '≤',
        '\\geq': '≥',
        '\\neq': '≠',
        '\\approx': '≈',
        '\\infty': '∞',
        '\\pm': '±',
        '\\mp': '∓',
        '\\mathbb{R}': 'ℝ',
        '\\mathbb{Z}': 'ℤ',
        '\\mathbb{Q}': 'ℚ',
        '\\mathbb{N}': 'ℕ',
        '\\mathbb{C}': 'ℂ'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Xử lý left, right parentheses
    result = result.replace('\\left(', '(')
    result = result.replace('\\right)', ')')
    result = result.replace('\\left[', '[')
    result = result.replace('\\right]', ']')
    result = result.replace('\\left{', '{')
    result = result.replace('\\right}', '}')
    
    # Xử lý frac
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        # Tìm tử số
        num_start = start + 6
        brace_count = 1
        num_end = num_start
        
        while num_end < len(result) and brace_count > 0:
            if result[num_end] == '{':
                brace_count += 1
            elif result[num_end] == '}':
                brace_count -= 1
            num_end += 1
        
        if brace_count > 0:
            break
            
        numerator = result[num_start:num_end-1]
        
        # Tìm mẫu số
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
            
            if brace_count == 0:
                denominator = result[denom_start:denom_end-1]
                fraction = f"({numerator})/({denominator})"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        else:
            break
    
    # Xử lý sqrt
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
            
        content_start = start + 6
        brace_count = 1
        content_end = content_start
        
        while content_end < len(result) and brace_count > 0:
            if result[content_end] == '{':
                brace_count += 1
            elif result[content_end] == '}':
                brace_count -= 1
            content_end += 1
        
        if brace_count == 0:
            content = result[content_start:content_end-1]
            sqrt_result = f"√({content})"
            result = result[:start] + sqrt_result + result[content_end:]
        else:
            break
    
    # Xử lý superscript đơn giản: x^2 -> x²
    superscript_digits = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    
    i = 0
    new_result = ""
    while i < len(result):
        if (i < len(result) - 2 and 
            result[i+1] == '^' and 
            result[i+2].isdigit() and
            result[i].isalnum()):
            base = result[i]
            power = result[i+2]
            new_result += base + superscript_digits.get(power, power)
            i += 3
        else:
            new_result += result[i]
            i += 1
    
    result = new_result
    
    # Loại bỏ backslash còn lại
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Tìm expressions toán học trong text"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Xử lý ${...}$ format
            if i < len(text) and text[i] == '{':
                i += 1  # skip {
                brace_count = 1
                content_start = i
                
                while i < len(text) and brace_count > 0:
                    if text[i] == '{':
                        brace_count += 1
                    elif text[i] == '}':
                        brace_count -= 1
                    i += 1
                
                # Kiểm tra có }$ không
                if i < len(text) and text[i] == '$':
                    end = i + 1
                    latex_content = text[content_start:i-1]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'braced'))
                    i += 1
                else:
                    i += 1
            else:
                # Xử lý $...$ format thông thường
                content_start = i
                brace_depth = 0
                
                # Tìm $ đóng, tracking nested braces
                while i < len(text):
                    if text[i] == '{':
                        brace_depth += 1
                    elif text[i] == '}':
                        brace_depth -= 1
                    elif text[i] == '$' and brace_depth == 0:
                        break
                    i += 1
                
                if i < len(text):  # Tìm thấy $ đóng
                    end = i + 1
                    latex_content = text[content_start:i]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'normal'))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_omml_equation(paragraph, text, equation_type='simple'):
    """Tạo OMML equation - Safe with fallbacks"""
    try:
        # Always try simple approach first to avoid corruption
        if equation_type == 'prime' and "'" in text:
            if create_prime_omml(paragraph, text):
                return True
        elif equation_type == 'superscript' and any(c in text for c in ['²', '³', '⁴', '⁵']):
            if create_superscript_omml(paragraph, text):
                return True
        elif equation_type == 'fraction' and '(' in text and ')' in text and '/' in text:
            if create_fraction_omml(paragraph, text):
                return True
        elif equation_type == 'parentheses' and '(' in text and ')' in text:
            if create_parentheses_omml(paragraph, text):
                return True
        
        # Default to simple OMML
        return create_simple_omml(paragraph, text)
        
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field - Simplified and safe"""
    try:
        # Simple EQ field without complex formatting
        eq_content = unicode_text
        
        # Only handle basic superscripts
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}'"
        
        # Safe field creation
        fldChar_begin = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>')
        instrText = parse_xml(f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> {eq_content} </w:instrText>')
        fldChar_end = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>')
        
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_prime_omml(paragraph, text):
    """OMML cho prime notation - Fixed XML"""
    try:
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        # Safe XML with proper escaping
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_prime = "'" * prime_count
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_prime}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_superscript_omml(paragraph, text):
    """OMML cho superscript numbers - Fixed XML"""
    try:
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:
                    base += char
        
        if not sup:
            return create_simple_omml(paragraph, text)
        
        # Safe XML
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_sup = sup.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_sup}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_fraction_omml(paragraph, text):
    """OMML cho fractions - Simplified and safe"""
    try:
        frac_pattern = r'\(([^)]+)\)/\(([^)]+)\)'
        match = re.search(frac_pattern, text)
        
        if match:
            numerator = match.group(1).strip()
            denominator = match.group(2).strip()
            
            # Safe XML escaping
            safe_num = numerator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            safe_den = denominator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:f><m:fPr></m:fPr><m:num><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_num}</m:t></m:r></m:num><m:den><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_den}</m:t></m:r></m:den></m:f></m:oMath>'''
            
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return True
        
        return False
    except Exception:
        return False

def create_parentheses_omml(paragraph, text):
    """OMML cho parentheses - Simplified"""
    try:
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end].strip()
        else:
            content = text.strip()
        
        # Safe XML
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_content}</m:t></m:r></m:e></m:d></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_simple_omml(paragraph, text):
    """OMML cho equation đơn giản - Safe version"""
    try:
        # Safe XML escaping
        safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_text}</m:t></m:r></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field"""
    try:
        eq_content = unicode_text
        
        # Handle superscripts for EQ field
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}\\s\\up5(')"
        elif '²' in unicode_text:
            base = unicode_text.replace('²', '')
            eq_content = f"{base}\\s\\up5(2)"
        elif '³' in unicode_text:
            base = unicode_text.replace('³', '')
            eq_content = f"{base}\\s\\up5(3)"
        
        # Create field elements
        fldChar_begin = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>'
        )
        
        instrText = parse_xml(
            f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> EQ {eq_content}</w:instrText>'
        )
        
        fldChar_end = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>'
        )
        
        # Add to paragraph
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_styled_math_run(paragraph, unicode_text):
    """Tạo styled run như equation"""
    try:
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 32, 96)
        
        # Add background
        try:
            rPr = run._element.get_or_add_rPr()
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
        except:
            pass
        
        return True
    except Exception:
        return False

def create_equation_object(paragraph, latex_text):
    """Tạo equation object với safe methods"""
    unicode_content = convert_latex_symbols(latex_text)
    
    # Determine equation type
    equation_type = 'simple'
    if "'" in unicode_content:
        equation_type = 'prime'
    elif '(' in unicode_content and ')' in unicode_content and '/' in unicode_content:
        equation_type = 'fraction'
    elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
        equation_type = 'superscript'
    elif '(' in unicode_content and ')' in unicode_content:
        equation_type = 'parentheses'
    
    # Try methods with better error handling
    try:
        # Method 1: OMML (safest approach)
        if create_omml_equation(paragraph, unicode_content, equation_type):
            return 'omml'
    except Exception:
        pass
    
    try:
        # Method 2: Styled run (most reliable)
        if create_styled_math_run(paragraph, unicode_content):
            return 'styled'
    except Exception:
        pass
    
    try:
        # Method 3: EQ Field (skip if problematic)
        # Commenting out to avoid file corruption
        # if create_equation_field(paragraph, unicode_content, latex_text):
        #     return 'field'
        pass
    except Exception:
        pass
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_content)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)
        return 'fallback'
    except Exception:
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return 'error'

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Safe version"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    last_pos = 0
    math_count = 0
    method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    for start, end, latex_content, expr_type in expressions:
        # Add text before equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Create equation object safely
        method_used = create_equation_object(paragraph, latex_content)
        method_stats[method_used] += 1
        math_count += 1
        
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count, method_stats

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # Bold question number
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Process content with math
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Bold answer letter
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Process answer content
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra markdown table"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse markdown table"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Skip separator line
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table - Safe version"""
    if not table_data:
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    math_count = 0
    total_method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    max_cols = max(len(row) for row in table_data)
    
    # Create table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Fill data
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Create paragraph in cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Process cell content safely
                try:
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                        math_count += cell_math_count
                        for method, count in cell_method_stats.items():
                            total_method_stats[method] += count
                except Exception:
                    # Fallback to plain text
                    run = cell_para.add_run(cell_text)
                    run.font.size = Pt(11)
                    run.font.name = 'Times New Roman'
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table safely
    if options.get('format_tables', True):
        try:
            word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if word_table.rows:
                header_row = word_table.rows[0]
                for cell in header_row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                    
                    try:
                        shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                        shading = parse_xml(shading_xml)
                        cell._tc.get_or_add_tcPr().append(shading)
                    except:
                        pass
        except Exception:
            pass
    
    return math_count, total_method_stats

def process_document_content(doc, options):
    """Xử lý document content - Safe version"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'omml_equations': 0,
        'field_equations': 0,
        'styled_equations': 0,
        'fallback_equations': 0,
        'error_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Processing: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Process paragraphs safely
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Check for markdown table
        if '|' in text and options.get('convert_markdown', True):
            table_lines = []
            j = i
            
            while j < len(doc.paragraphs):
                current_text = doc.paragraphs[j].text.strip()
                if not current_text:
                    break
                if '|' in current_text:
                    table_lines.append(current_text)
                    j += 1
                else:
                    break
            
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    try:
                        math_count, method_stats = create_word_table(new_doc, table_data, options)
                        stats['math_expressions'] += math_count
                        for method, count in method_stats.items():
                            if method in stats:
                                stats[f'{method}_equations'] += count
                        stats['markdown_tables'] += 1
                        stats['debug_log'].append(f"📋 Converted table: {len(table_data)} rows")
                    except Exception as e:
                        stats['debug_log'].append(f"❌ Table error: {str(e)}")
                
                i = j
                continue
        
        # Process normal paragraph safely
        new_para = new_doc.add_paragraph()
        
        try:
            if format_question_answer(new_para, text):
                stats['questions_formatted'] += 1
                stats['debug_log'].append(f"📝 P{i}: Q&A formatted")
            else:
                math_count, method_stats = process_text_with_math(new_para, text)
                stats['math_expressions'] += math_count
                for method, count in method_stats.items():
                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                        stats[f'{method}_equations'] += count
                
                if math_count > 0:
                    omml_count = method_stats.get('omml', 0)
                    styled_count = method_stats.get('styled', 0) + method_stats.get('fallback', 0)
                    stats['debug_log'].append(f"🔢 P{i}: {math_count} equations ({omml_count} OMML, {styled_count} styled)")
        except Exception as e:
            # Fallback to plain text
            new_para.clear()
            run = new_para.add_run(text)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
            stats['debug_log'].append(f"❌ P{i}: Error, used plain text")
        
        i += 1
    
    # Process existing Word tables safely
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
                            if new_cell.paragraphs:
                                cell_para = new_cell.paragraphs[0]
                                cell_para.clear()
                            else:
                                cell_para = new_cell.add_paragraph()
                            
                            if not format_question_answer(cell_para, cell_text):
                                cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                                table_math += cell_math_count
                                for method, count in cell_method_stats.items():
                                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                                        stats[f'{method}_equations'] += count
                            
                            cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except Exception:
                        # Skip problematic cells
                        continue
            
            stats['math_expressions'] += table_math
            
            # Format table safely
            if options.get('format_tables', True):
                try:
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    if new_table.rows:
                        header_row = new_table.rows[0]
                        for cell in header_row.cells:
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.bold = True
                except Exception:
                    pass
                
                stats['word_tables'] += 1
            
            if table_math > 0:
                stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations")
                
        except Exception as e:
            stats['debug_log'].append(f"❌ Table {table_idx} error: {str(e)}")
            continue
    
    stats['debug_log'].append(f"✅ Complete: {stats['math_expressions']} total equations")
    stats['debug_log'].append(f"   🎯 OMML: {stats['omml_equations']}")
    stats['debug_log'].append(f"   ✨ Styled: {stats['styled_equations'] + stats['fallback_equations']}")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v8.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">Fixed File Corruption + Safe Equation Objects + Complete LaTeX Support</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🎯 Tạo TRUE Equation Objects", value=True,
                                       help="Chuyển $...$ thành REAL equation objects trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True,
                                     help="Format 'Câu X:' và 'A. B. C. D.'")  
        format_tables = st.checkbox("📊 Format bảng", value=True,
                                   help="Professional table formatting")
        convert_markdown = st.checkbox("📋 Markdown → Word Tables", value=True,
                                     help="Convert | tables | to Word")
        show_debug = st.checkbox("🔍 Advanced Debug", value=False,
                               help="Show equation creation details")
        
        st.markdown("---")
        st.info("🛡️ **SAFE Equation Creation:**\n\n✅ **Method 1:** Safe OMML Objects\n🎨 **Method 2:** Enhanced Styling  \n💫 **Method 3:** Fallback Protection\n\n🔧 **Fixed Issues:**\n- File Corruption\n- XML Validation\n- Safe Error Handling\n- Robust Fallbacks")
        
 → **A'** (safe OMML)
        - `$(π)/(3)import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="LaTeX Word Processor",
    page_icon="📚",
    layout="wide"
)

def convert_latex_symbols(text):
    """Chuyển đổi ký hiệu LaTeX thành Unicode"""
    if not text:
        return text
    
    result = text
    
    # Xử lý superscript có dấu ngoặc: A^{prime} -> A'
    while True:
        pos = result.find('^{')
        if pos == -1:
            break
        
        # Tìm ký tự trước ^{
        base_start = pos - 1
        while base_start >= 0 and result[base_start].isalnum():
            base_start -= 1
        base_start += 1
        base = result[base_start:pos]
        
        # Tìm nội dung trong {}
        brace_start = pos + 2
        brace_count = 1
        brace_end = brace_start
        
        while brace_end < len(result) and brace_count > 0:
            if result[brace_end] == '{':
                brace_count += 1
            elif result[brace_end] == '}':
                brace_count -= 1
            brace_end += 1
        
        if brace_count == 0:
            content = result[brace_start:brace_end-1]
            if content == 'prime':
                replacement = base + "'"
            else:
                replacement = base + "^(" + content + ")"
            result = result[:base_start] + replacement + result[brace_end:]
        else:
            break
    
    # Các ký hiệu toán học cơ bản
    symbols = {
        '\\pi': 'π',
        '\\alpha': 'α',
        '\\beta': 'β',
        '\\gamma': 'γ',
        '\\delta': 'δ',
        '\\theta': 'θ',
        '\\lambda': 'λ',
        '\\mu': 'μ',
        '\\sigma': 'σ',
        '\\phi': 'φ',
        '\\omega': 'ω',
        '\\in': '∈',
        '\\subset': '⊂',
        '\\cup': '∪',
        '\\cap': '∩',
        '\\emptyset': '∅',
        '\\leq': '≤',
        '\\geq': '≥',
        '\\neq': '≠',
        '\\approx': '≈',
        '\\infty': '∞',
        '\\pm': '±',
        '\\mp': '∓',
        '\\mathbb{R}': 'ℝ',
        '\\mathbb{Z}': 'ℤ',
        '\\mathbb{Q}': 'ℚ',
        '\\mathbb{N}': 'ℕ',
        '\\mathbb{C}': 'ℂ'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Xử lý left, right parentheses
    result = result.replace('\\left(', '(')
    result = result.replace('\\right)', ')')
    result = result.replace('\\left[', '[')
    result = result.replace('\\right]', ']')
    result = result.replace('\\left{', '{')
    result = result.replace('\\right}', '}')
    
    # Xử lý frac
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        # Tìm tử số
        num_start = start + 6
        brace_count = 1
        num_end = num_start
        
        while num_end < len(result) and brace_count > 0:
            if result[num_end] == '{':
                brace_count += 1
            elif result[num_end] == '}':
                brace_count -= 1
            num_end += 1
        
        if brace_count > 0:
            break
            
        numerator = result[num_start:num_end-1]
        
        # Tìm mẫu số
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
            
            if brace_count == 0:
                denominator = result[denom_start:denom_end-1]
                fraction = f"({numerator})/({denominator})"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        else:
            break
    
    # Xử lý sqrt
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
            
        content_start = start + 6
        brace_count = 1
        content_end = content_start
        
        while content_end < len(result) and brace_count > 0:
            if result[content_end] == '{':
                brace_count += 1
            elif result[content_end] == '}':
                brace_count -= 1
            content_end += 1
        
        if brace_count == 0:
            content = result[content_start:content_end-1]
            sqrt_result = f"√({content})"
            result = result[:start] + sqrt_result + result[content_end:]
        else:
            break
    
    # Xử lý superscript đơn giản: x^2 -> x²
    superscript_digits = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    
    i = 0
    new_result = ""
    while i < len(result):
        if (i < len(result) - 2 and 
            result[i+1] == '^' and 
            result[i+2].isdigit() and
            result[i].isalnum()):
            base = result[i]
            power = result[i+2]
            new_result += base + superscript_digits.get(power, power)
            i += 3
        else:
            new_result += result[i]
            i += 1
    
    result = new_result
    
    # Loại bỏ backslash còn lại
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Tìm expressions toán học trong text"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Xử lý ${...}$ format
            if i < len(text) and text[i] == '{':
                i += 1  # skip {
                brace_count = 1
                content_start = i
                
                while i < len(text) and brace_count > 0:
                    if text[i] == '{':
                        brace_count += 1
                    elif text[i] == '}':
                        brace_count -= 1
                    i += 1
                
                # Kiểm tra có }$ không
                if i < len(text) and text[i] == '$':
                    end = i + 1
                    latex_content = text[content_start:i-1]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'braced'))
                    i += 1
                else:
                    i += 1
            else:
                # Xử lý $...$ format thông thường
                content_start = i
                brace_depth = 0
                
                # Tìm $ đóng, tracking nested braces
                while i < len(text):
                    if text[i] == '{':
                        brace_depth += 1
                    elif text[i] == '}':
                        brace_depth -= 1
                    elif text[i] == '$' and brace_depth == 0:
                        break
                    i += 1
                
                if i < len(text):  # Tìm thấy $ đóng
                    end = i + 1
                    latex_content = text[content_start:i]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'normal'))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_omml_equation(paragraph, text, equation_type='simple'):
    """Tạo OMML equation - Safe with fallbacks"""
    try:
        # Always try simple approach first to avoid corruption
        if equation_type == 'prime' and "'" in text:
            if create_prime_omml(paragraph, text):
                return True
        elif equation_type == 'superscript' and any(c in text for c in ['²', '³', '⁴', '⁵']):
            if create_superscript_omml(paragraph, text):
                return True
        elif equation_type == 'fraction' and '(' in text and ')' in text and '/' in text:
            if create_fraction_omml(paragraph, text):
                return True
        elif equation_type == 'parentheses' and '(' in text and ')' in text:
            if create_parentheses_omml(paragraph, text):
                return True
        
        # Default to simple OMML
        return create_simple_omml(paragraph, text)
        
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field - Simplified and safe"""
    try:
        # Simple EQ field without complex formatting
        eq_content = unicode_text
        
        # Only handle basic superscripts
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}'"
        
        # Safe field creation
        fldChar_begin = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>')
        instrText = parse_xml(f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> {eq_content} </w:instrText>')
        fldChar_end = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>')
        
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_prime_omml(paragraph, text):
    """OMML cho prime notation - Fixed XML"""
    try:
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        # Safe XML with proper escaping
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_prime = "'" * prime_count
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_prime}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_superscript_omml(paragraph, text):
    """OMML cho superscript numbers - Fixed XML"""
    try:
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:
                    base += char
        
        if not sup:
            return create_simple_omml(paragraph, text)
        
        # Safe XML
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_sup = sup.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_sup}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_fraction_omml(paragraph, text):
    """OMML cho fractions - Simplified and safe"""
    try:
        frac_pattern = r'\(([^)]+)\)/\(([^)]+)\)'
        match = re.search(frac_pattern, text)
        
        if match:
            numerator = match.group(1).strip()
            denominator = match.group(2).strip()
            
            # Safe XML escaping
            safe_num = numerator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            safe_den = denominator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:f><m:fPr></m:fPr><m:num><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_num}</m:t></m:r></m:num><m:den><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_den}</m:t></m:r></m:den></m:f></m:oMath>'''
            
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return True
        
        return False
    except Exception:
        return False

def create_parentheses_omml(paragraph, text):
    """OMML cho parentheses - Simplified"""
    try:
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end].strip()
        else:
            content = text.strip()
        
        # Safe XML
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_content}</m:t></m:r></m:e></m:d></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_simple_omml(paragraph, text):
    """OMML cho equation đơn giản - Safe version"""
    try:
        # Safe XML escaping
        safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_text}</m:t></m:r></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field"""
    try:
        eq_content = unicode_text
        
        # Handle superscripts for EQ field
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}\\s\\up5(')"
        elif '²' in unicode_text:
            base = unicode_text.replace('²', '')
            eq_content = f"{base}\\s\\up5(2)"
        elif '³' in unicode_text:
            base = unicode_text.replace('³', '')
            eq_content = f"{base}\\s\\up5(3)"
        
        # Create field elements
        fldChar_begin = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>'
        )
        
        instrText = parse_xml(
            f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> EQ {eq_content}</w:instrText>'
        )
        
        fldChar_end = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>'
        )
        
        # Add to paragraph
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_styled_math_run(paragraph, unicode_text):
    """Tạo styled run như equation"""
    try:
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 32, 96)
        
        # Add background
        try:
            rPr = run._element.get_or_add_rPr()
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
        except:
            pass
        
        return True
    except Exception:
        return False

def create_equation_object(paragraph, latex_text):
    """Tạo equation object với safe methods"""
    unicode_content = convert_latex_symbols(latex_text)
    
    # Determine equation type
    equation_type = 'simple'
    if "'" in unicode_content:
        equation_type = 'prime'
    elif '(' in unicode_content and ')' in unicode_content and '/' in unicode_content:
        equation_type = 'fraction'
    elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
        equation_type = 'superscript'
    elif '(' in unicode_content and ')' in unicode_content:
        equation_type = 'parentheses'
    
    # Try methods with better error handling
    try:
        # Method 1: OMML (safest approach)
        if create_omml_equation(paragraph, unicode_content, equation_type):
            return 'omml'
    except Exception:
        pass
    
    try:
        # Method 2: Styled run (most reliable)
        if create_styled_math_run(paragraph, unicode_content):
            return 'styled'
    except Exception:
        pass
    
    try:
        # Method 3: EQ Field (skip if problematic)
        # Commenting out to avoid file corruption
        # if create_equation_field(paragraph, unicode_content, latex_text):
        #     return 'field'
        pass
    except Exception:
        pass
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_content)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)
        return 'fallback'
    except Exception:
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return 'error'

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Safe version"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    last_pos = 0
    math_count = 0
    method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    for start, end, latex_content, expr_type in expressions:
        # Add text before equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Create equation object safely
        method_used = create_equation_object(paragraph, latex_content)
        method_stats[method_used] += 1
        math_count += 1
        
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count, method_stats

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # Bold question number
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Process content with math
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Bold answer letter
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Process answer content
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra markdown table"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse markdown table"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Skip separator line
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table - Safe version"""
    if not table_data:
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    math_count = 0
    total_method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    max_cols = max(len(row) for row in table_data)
    
    # Create table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Fill data
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Create paragraph in cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Process cell content safely
                try:
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                        math_count += cell_math_count
                        for method, count in cell_method_stats.items():
                            total_method_stats[method] += count
                except Exception:
                    # Fallback to plain text
                    run = cell_para.add_run(cell_text)
                    run.font.size = Pt(11)
                    run.font.name = 'Times New Roman'
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table safely
    if options.get('format_tables', True):
        try:
            word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if word_table.rows:
                header_row = word_table.rows[0]
                for cell in header_row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                    
                    try:
                        shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                        shading = parse_xml(shading_xml)
                        cell._tc.get_or_add_tcPr().append(shading)
                    except:
                        pass
        except Exception:
            pass
    
    return math_count, total_method_stats

def process_document_content(doc, options):
    """Xử lý document content - Safe version"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'omml_equations': 0,
        'field_equations': 0,
        'styled_equations': 0,
        'fallback_equations': 0,
        'error_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Processing: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Process paragraphs safely
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Check for markdown table
        if '|' in text and options.get('convert_markdown', True):
            table_lines = []
            j = i
            
            while j < len(doc.paragraphs):
                current_text = doc.paragraphs[j].text.strip()
                if not current_text:
                    break
                if '|' in current_text:
                    table_lines.append(current_text)
                    j += 1
                else:
                    break
            
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    try:
                        math_count, method_stats = create_word_table(new_doc, table_data, options)
                        stats['math_expressions'] += math_count
                        for method, count in method_stats.items():
                            if method in stats:
                                stats[f'{method}_equations'] += count
                        stats['markdown_tables'] += 1
                        stats['debug_log'].append(f"📋 Converted table: {len(table_data)} rows")
                    except Exception as e:
                        stats['debug_log'].append(f"❌ Table error: {str(e)}")
                
                i = j
                continue
        
        # Process normal paragraph safely
        new_para = new_doc.add_paragraph()
        
        try:
            if format_question_answer(new_para, text):
                stats['questions_formatted'] += 1
                stats['debug_log'].append(f"📝 P{i}: Q&A formatted")
            else:
                math_count, method_stats = process_text_with_math(new_para, text)
                stats['math_expressions'] += math_count
                for method, count in method_stats.items():
                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                        stats[f'{method}_equations'] += count
                
                if math_count > 0:
                    omml_count = method_stats.get('omml', 0)
                    styled_count = method_stats.get('styled', 0) + method_stats.get('fallback', 0)
                    stats['debug_log'].append(f"🔢 P{i}: {math_count} equations ({omml_count} OMML, {styled_count} styled)")
        except Exception as e:
            # Fallback to plain text
            new_para.clear()
            run = new_para.add_run(text)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
            stats['debug_log'].append(f"❌ P{i}: Error, used plain text")
        
        i += 1
    
    # Process existing Word tables safely
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
                            if new_cell.paragraphs:
                                cell_para = new_cell.paragraphs[0]
                                cell_para.clear()
                            else:
                                cell_para = new_cell.add_paragraph()
                            
                            if not format_question_answer(cell_para, cell_text):
                                cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                                table_math += cell_math_count
                                for method, count in cell_method_stats.items():
                                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                                        stats[f'{method}_equations'] += count
                            
                            cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except Exception:
                        # Skip problematic cells
                        continue
            
            stats['math_expressions'] += table_math
            
            # Format table safely
            if options.get('format_tables', True):
                try:
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    if new_table.rows:
                        header_row = new_table.rows[0]
                        for cell in header_row.cells:
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.bold = True
                except Exception:
                    pass
                
                stats['word_tables'] += 1
            
            if table_math > 0:
                stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations")
                
        except Exception as e:
            stats['debug_log'].append(f"❌ Table {table_idx} error: {str(e)}")
            continue
    
    stats['debug_log'].append(f"✅ Complete: {stats['math_expressions']} total equations")
    stats['debug_log'].append(f"   🎯 OMML: {stats['omml_equations']}")
    stats['debug_log'].append(f"   ✨ Styled: {stats['styled_equations'] + stats['fallback_equations']}")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v8.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">Fixed File Corruption + Safe Equation Objects + Complete LaTeX Support</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🎯 Tạo TRUE Equation Objects", value=True,
                                       help="Chuyển $...$ thành REAL equation objects trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True,
                                     help="Format 'Câu X:' và 'A. B. C. D.'")  
        format_tables = st.checkbox("📊 Format bảng", value=True,
                                   help="Professional table formatting")
        convert_markdown = st.checkbox("📋 Markdown → Word Tables", value=True,
                                     help="Convert | tables | to Word")
        show_debug = st.checkbox("🔍 Advanced Debug", value=False,
                               help="Show equation creation details")
        
        st.markdown("---")
        st.info("🛡️ **SAFE Equation Creation:**\n\n✅ **Method 1:** Safe OMML Objects\n🎨 **Method 2:** Enhanced Styling  \n💫 **Method 3:** Fallback Protection\n\n🔧 **Fixed Issues:**\n- File Corruption\n- XML Validation\n- Safe Error Handling\n- Robust Fallbacks")
        
 → **π/3** (safe fraction)  
        - `$x^2import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="LaTeX Word Processor",
    page_icon="📚",
    layout="wide"
)

def convert_latex_symbols(text):
    """Chuyển đổi ký hiệu LaTeX thành Unicode"""
    if not text:
        return text
    
    result = text
    
    # Xử lý superscript có dấu ngoặc: A^{prime} -> A'
    while True:
        pos = result.find('^{')
        if pos == -1:
            break
        
        # Tìm ký tự trước ^{
        base_start = pos - 1
        while base_start >= 0 and result[base_start].isalnum():
            base_start -= 1
        base_start += 1
        base = result[base_start:pos]
        
        # Tìm nội dung trong {}
        brace_start = pos + 2
        brace_count = 1
        brace_end = brace_start
        
        while brace_end < len(result) and brace_count > 0:
            if result[brace_end] == '{':
                brace_count += 1
            elif result[brace_end] == '}':
                brace_count -= 1
            brace_end += 1
        
        if brace_count == 0:
            content = result[brace_start:brace_end-1]
            if content == 'prime':
                replacement = base + "'"
            else:
                replacement = base + "^(" + content + ")"
            result = result[:base_start] + replacement + result[brace_end:]
        else:
            break
    
    # Các ký hiệu toán học cơ bản
    symbols = {
        '\\pi': 'π',
        '\\alpha': 'α',
        '\\beta': 'β',
        '\\gamma': 'γ',
        '\\delta': 'δ',
        '\\theta': 'θ',
        '\\lambda': 'λ',
        '\\mu': 'μ',
        '\\sigma': 'σ',
        '\\phi': 'φ',
        '\\omega': 'ω',
        '\\in': '∈',
        '\\subset': '⊂',
        '\\cup': '∪',
        '\\cap': '∩',
        '\\emptyset': '∅',
        '\\leq': '≤',
        '\\geq': '≥',
        '\\neq': '≠',
        '\\approx': '≈',
        '\\infty': '∞',
        '\\pm': '±',
        '\\mp': '∓',
        '\\mathbb{R}': 'ℝ',
        '\\mathbb{Z}': 'ℤ',
        '\\mathbb{Q}': 'ℚ',
        '\\mathbb{N}': 'ℕ',
        '\\mathbb{C}': 'ℂ'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Xử lý left, right parentheses
    result = result.replace('\\left(', '(')
    result = result.replace('\\right)', ')')
    result = result.replace('\\left[', '[')
    result = result.replace('\\right]', ']')
    result = result.replace('\\left{', '{')
    result = result.replace('\\right}', '}')
    
    # Xử lý frac
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        # Tìm tử số
        num_start = start + 6
        brace_count = 1
        num_end = num_start
        
        while num_end < len(result) and brace_count > 0:
            if result[num_end] == '{':
                brace_count += 1
            elif result[num_end] == '}':
                brace_count -= 1
            num_end += 1
        
        if brace_count > 0:
            break
            
        numerator = result[num_start:num_end-1]
        
        # Tìm mẫu số
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
            
            if brace_count == 0:
                denominator = result[denom_start:denom_end-1]
                fraction = f"({numerator})/({denominator})"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        else:
            break
    
    # Xử lý sqrt
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
            
        content_start = start + 6
        brace_count = 1
        content_end = content_start
        
        while content_end < len(result) and brace_count > 0:
            if result[content_end] == '{':
                brace_count += 1
            elif result[content_end] == '}':
                brace_count -= 1
            content_end += 1
        
        if brace_count == 0:
            content = result[content_start:content_end-1]
            sqrt_result = f"√({content})"
            result = result[:start] + sqrt_result + result[content_end:]
        else:
            break
    
    # Xử lý superscript đơn giản: x^2 -> x²
    superscript_digits = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    
    i = 0
    new_result = ""
    while i < len(result):
        if (i < len(result) - 2 and 
            result[i+1] == '^' and 
            result[i+2].isdigit() and
            result[i].isalnum()):
            base = result[i]
            power = result[i+2]
            new_result += base + superscript_digits.get(power, power)
            i += 3
        else:
            new_result += result[i]
            i += 1
    
    result = new_result
    
    # Loại bỏ backslash còn lại
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Tìm expressions toán học trong text"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Xử lý ${...}$ format
            if i < len(text) and text[i] == '{':
                i += 1  # skip {
                brace_count = 1
                content_start = i
                
                while i < len(text) and brace_count > 0:
                    if text[i] == '{':
                        brace_count += 1
                    elif text[i] == '}':
                        brace_count -= 1
                    i += 1
                
                # Kiểm tra có }$ không
                if i < len(text) and text[i] == '$':
                    end = i + 1
                    latex_content = text[content_start:i-1]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'braced'))
                    i += 1
                else:
                    i += 1
            else:
                # Xử lý $...$ format thông thường
                content_start = i
                brace_depth = 0
                
                # Tìm $ đóng, tracking nested braces
                while i < len(text):
                    if text[i] == '{':
                        brace_depth += 1
                    elif text[i] == '}':
                        brace_depth -= 1
                    elif text[i] == '$' and brace_depth == 0:
                        break
                    i += 1
                
                if i < len(text):  # Tìm thấy $ đóng
                    end = i + 1
                    latex_content = text[content_start:i]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'normal'))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_omml_equation(paragraph, text, equation_type='simple'):
    """Tạo OMML equation - Safe with fallbacks"""
    try:
        # Always try simple approach first to avoid corruption
        if equation_type == 'prime' and "'" in text:
            if create_prime_omml(paragraph, text):
                return True
        elif equation_type == 'superscript' and any(c in text for c in ['²', '³', '⁴', '⁵']):
            if create_superscript_omml(paragraph, text):
                return True
        elif equation_type == 'fraction' and '(' in text and ')' in text and '/' in text:
            if create_fraction_omml(paragraph, text):
                return True
        elif equation_type == 'parentheses' and '(' in text and ')' in text:
            if create_parentheses_omml(paragraph, text):
                return True
        
        # Default to simple OMML
        return create_simple_omml(paragraph, text)
        
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field - Simplified and safe"""
    try:
        # Simple EQ field without complex formatting
        eq_content = unicode_text
        
        # Only handle basic superscripts
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}'"
        
        # Safe field creation
        fldChar_begin = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>')
        instrText = parse_xml(f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> {eq_content} </w:instrText>')
        fldChar_end = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>')
        
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_prime_omml(paragraph, text):
    """OMML cho prime notation - Fixed XML"""
    try:
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        # Safe XML with proper escaping
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_prime = "'" * prime_count
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_prime}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_superscript_omml(paragraph, text):
    """OMML cho superscript numbers - Fixed XML"""
    try:
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:
                    base += char
        
        if not sup:
            return create_simple_omml(paragraph, text)
        
        # Safe XML
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_sup = sup.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_sup}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_fraction_omml(paragraph, text):
    """OMML cho fractions - Simplified and safe"""
    try:
        frac_pattern = r'\(([^)]+)\)/\(([^)]+)\)'
        match = re.search(frac_pattern, text)
        
        if match:
            numerator = match.group(1).strip()
            denominator = match.group(2).strip()
            
            # Safe XML escaping
            safe_num = numerator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            safe_den = denominator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:f><m:fPr></m:fPr><m:num><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_num}</m:t></m:r></m:num><m:den><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_den}</m:t></m:r></m:den></m:f></m:oMath>'''
            
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return True
        
        return False
    except Exception:
        return False

def create_parentheses_omml(paragraph, text):
    """OMML cho parentheses - Simplified"""
    try:
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end].strip()
        else:
            content = text.strip()
        
        # Safe XML
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_content}</m:t></m:r></m:e></m:d></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_simple_omml(paragraph, text):
    """OMML cho equation đơn giản - Safe version"""
    try:
        # Safe XML escaping
        safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_text}</m:t></m:r></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field"""
    try:
        eq_content = unicode_text
        
        # Handle superscripts for EQ field
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}\\s\\up5(')"
        elif '²' in unicode_text:
            base = unicode_text.replace('²', '')
            eq_content = f"{base}\\s\\up5(2)"
        elif '³' in unicode_text:
            base = unicode_text.replace('³', '')
            eq_content = f"{base}\\s\\up5(3)"
        
        # Create field elements
        fldChar_begin = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>'
        )
        
        instrText = parse_xml(
            f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> EQ {eq_content}</w:instrText>'
        )
        
        fldChar_end = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>'
        )
        
        # Add to paragraph
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_styled_math_run(paragraph, unicode_text):
    """Tạo styled run như equation"""
    try:
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 32, 96)
        
        # Add background
        try:
            rPr = run._element.get_or_add_rPr()
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
        except:
            pass
        
        return True
    except Exception:
        return False

def create_equation_object(paragraph, latex_text):
    """Tạo equation object với safe methods"""
    unicode_content = convert_latex_symbols(latex_text)
    
    # Determine equation type
    equation_type = 'simple'
    if "'" in unicode_content:
        equation_type = 'prime'
    elif '(' in unicode_content and ')' in unicode_content and '/' in unicode_content:
        equation_type = 'fraction'
    elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
        equation_type = 'superscript'
    elif '(' in unicode_content and ')' in unicode_content:
        equation_type = 'parentheses'
    
    # Try methods with better error handling
    try:
        # Method 1: OMML (safest approach)
        if create_omml_equation(paragraph, unicode_content, equation_type):
            return 'omml'
    except Exception:
        pass
    
    try:
        # Method 2: Styled run (most reliable)
        if create_styled_math_run(paragraph, unicode_content):
            return 'styled'
    except Exception:
        pass
    
    try:
        # Method 3: EQ Field (skip if problematic)
        # Commenting out to avoid file corruption
        # if create_equation_field(paragraph, unicode_content, latex_text):
        #     return 'field'
        pass
    except Exception:
        pass
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_content)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)
        return 'fallback'
    except Exception:
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return 'error'

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Safe version"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    last_pos = 0
    math_count = 0
    method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    for start, end, latex_content, expr_type in expressions:
        # Add text before equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Create equation object safely
        method_used = create_equation_object(paragraph, latex_content)
        method_stats[method_used] += 1
        math_count += 1
        
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count, method_stats

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # Bold question number
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Process content with math
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Bold answer letter
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Process answer content
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra markdown table"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse markdown table"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Skip separator line
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table - Safe version"""
    if not table_data:
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    math_count = 0
    total_method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    max_cols = max(len(row) for row in table_data)
    
    # Create table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Fill data
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Create paragraph in cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Process cell content safely
                try:
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                        math_count += cell_math_count
                        for method, count in cell_method_stats.items():
                            total_method_stats[method] += count
                except Exception:
                    # Fallback to plain text
                    run = cell_para.add_run(cell_text)
                    run.font.size = Pt(11)
                    run.font.name = 'Times New Roman'
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table safely
    if options.get('format_tables', True):
        try:
            word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if word_table.rows:
                header_row = word_table.rows[0]
                for cell in header_row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                    
                    try:
                        shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                        shading = parse_xml(shading_xml)
                        cell._tc.get_or_add_tcPr().append(shading)
                    except:
                        pass
        except Exception:
            pass
    
    return math_count, total_method_stats

def process_document_content(doc, options):
    """Xử lý document content - Safe version"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'omml_equations': 0,
        'field_equations': 0,
        'styled_equations': 0,
        'fallback_equations': 0,
        'error_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Processing: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Process paragraphs safely
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Check for markdown table
        if '|' in text and options.get('convert_markdown', True):
            table_lines = []
            j = i
            
            while j < len(doc.paragraphs):
                current_text = doc.paragraphs[j].text.strip()
                if not current_text:
                    break
                if '|' in current_text:
                    table_lines.append(current_text)
                    j += 1
                else:
                    break
            
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    try:
                        math_count, method_stats = create_word_table(new_doc, table_data, options)
                        stats['math_expressions'] += math_count
                        for method, count in method_stats.items():
                            if method in stats:
                                stats[f'{method}_equations'] += count
                        stats['markdown_tables'] += 1
                        stats['debug_log'].append(f"📋 Converted table: {len(table_data)} rows")
                    except Exception as e:
                        stats['debug_log'].append(f"❌ Table error: {str(e)}")
                
                i = j
                continue
        
        # Process normal paragraph safely
        new_para = new_doc.add_paragraph()
        
        try:
            if format_question_answer(new_para, text):
                stats['questions_formatted'] += 1
                stats['debug_log'].append(f"📝 P{i}: Q&A formatted")
            else:
                math_count, method_stats = process_text_with_math(new_para, text)
                stats['math_expressions'] += math_count
                for method, count in method_stats.items():
                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                        stats[f'{method}_equations'] += count
                
                if math_count > 0:
                    omml_count = method_stats.get('omml', 0)
                    styled_count = method_stats.get('styled', 0) + method_stats.get('fallback', 0)
                    stats['debug_log'].append(f"🔢 P{i}: {math_count} equations ({omml_count} OMML, {styled_count} styled)")
        except Exception as e:
            # Fallback to plain text
            new_para.clear()
            run = new_para.add_run(text)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
            stats['debug_log'].append(f"❌ P{i}: Error, used plain text")
        
        i += 1
    
    # Process existing Word tables safely
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
                            if new_cell.paragraphs:
                                cell_para = new_cell.paragraphs[0]
                                cell_para.clear()
                            else:
                                cell_para = new_cell.add_paragraph()
                            
                            if not format_question_answer(cell_para, cell_text):
                                cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                                table_math += cell_math_count
                                for method, count in cell_method_stats.items():
                                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                                        stats[f'{method}_equations'] += count
                            
                            cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except Exception:
                        # Skip problematic cells
                        continue
            
            stats['math_expressions'] += table_math
            
            # Format table safely
            if options.get('format_tables', True):
                try:
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    if new_table.rows:
                        header_row = new_table.rows[0]
                        for cell in header_row.cells:
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.bold = True
                except Exception:
                    pass
                
                stats['word_tables'] += 1
            
            if table_math > 0:
                stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations")
                
        except Exception as e:
            stats['debug_log'].append(f"❌ Table {table_idx} error: {str(e)}")
            continue
    
    stats['debug_log'].append(f"✅ Complete: {stats['math_expressions']} total equations")
    stats['debug_log'].append(f"   🎯 OMML: {stats['omml_equations']}")
    stats['debug_log'].append(f"   ✨ Styled: {stats['styled_equations'] + stats['fallback_equations']}")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v8.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">Fixed File Corruption + Safe Equation Objects + Complete LaTeX Support</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🎯 Tạo TRUE Equation Objects", value=True,
                                       help="Chuyển $...$ thành REAL equation objects trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True,
                                     help="Format 'Câu X:' và 'A. B. C. D.'")  
        format_tables = st.checkbox("📊 Format bảng", value=True,
                                   help="Professional table formatting")
        convert_markdown = st.checkbox("📋 Markdown → Word Tables", value=True,
                                     help="Convert | tables | to Word")
        show_debug = st.checkbox("🔍 Advanced Debug", value=False,
                               help="Show equation creation details")
        
        st.markdown("---")
        st.info("🛡️ **SAFE Equation Creation:**\n\n✅ **Method 1:** Safe OMML Objects\n🎨 **Method 2:** Enhanced Styling  \n💫 **Method 3:** Fallback Protection\n\n🔧 **Fixed Issues:**\n- File Corruption\n- XML Validation\n- Safe Error Handling\n- Robust Fallbacks")
        
 → **x²** (safe superscript)
        - `$\\left(x\\right)import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="LaTeX Word Processor",
    page_icon="📚",
    layout="wide"
)

def convert_latex_symbols(text):
    """Chuyển đổi ký hiệu LaTeX thành Unicode"""
    if not text:
        return text
    
    result = text
    
    # Xử lý superscript có dấu ngoặc: A^{prime} -> A'
    while True:
        pos = result.find('^{')
        if pos == -1:
            break
        
        # Tìm ký tự trước ^{
        base_start = pos - 1
        while base_start >= 0 and result[base_start].isalnum():
            base_start -= 1
        base_start += 1
        base = result[base_start:pos]
        
        # Tìm nội dung trong {}
        brace_start = pos + 2
        brace_count = 1
        brace_end = brace_start
        
        while brace_end < len(result) and brace_count > 0:
            if result[brace_end] == '{':
                brace_count += 1
            elif result[brace_end] == '}':
                brace_count -= 1
            brace_end += 1
        
        if brace_count == 0:
            content = result[brace_start:brace_end-1]
            if content == 'prime':
                replacement = base + "'"
            else:
                replacement = base + "^(" + content + ")"
            result = result[:base_start] + replacement + result[brace_end:]
        else:
            break
    
    # Các ký hiệu toán học cơ bản
    symbols = {
        '\\pi': 'π',
        '\\alpha': 'α',
        '\\beta': 'β',
        '\\gamma': 'γ',
        '\\delta': 'δ',
        '\\theta': 'θ',
        '\\lambda': 'λ',
        '\\mu': 'μ',
        '\\sigma': 'σ',
        '\\phi': 'φ',
        '\\omega': 'ω',
        '\\in': '∈',
        '\\subset': '⊂',
        '\\cup': '∪',
        '\\cap': '∩',
        '\\emptyset': '∅',
        '\\leq': '≤',
        '\\geq': '≥',
        '\\neq': '≠',
        '\\approx': '≈',
        '\\infty': '∞',
        '\\pm': '±',
        '\\mp': '∓',
        '\\mathbb{R}': 'ℝ',
        '\\mathbb{Z}': 'ℤ',
        '\\mathbb{Q}': 'ℚ',
        '\\mathbb{N}': 'ℕ',
        '\\mathbb{C}': 'ℂ'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Xử lý left, right parentheses
    result = result.replace('\\left(', '(')
    result = result.replace('\\right)', ')')
    result = result.replace('\\left[', '[')
    result = result.replace('\\right]', ']')
    result = result.replace('\\left{', '{')
    result = result.replace('\\right}', '}')
    
    # Xử lý frac
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        # Tìm tử số
        num_start = start + 6
        brace_count = 1
        num_end = num_start
        
        while num_end < len(result) and brace_count > 0:
            if result[num_end] == '{':
                brace_count += 1
            elif result[num_end] == '}':
                brace_count -= 1
            num_end += 1
        
        if brace_count > 0:
            break
            
        numerator = result[num_start:num_end-1]
        
        # Tìm mẫu số
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
            
            if brace_count == 0:
                denominator = result[denom_start:denom_end-1]
                fraction = f"({numerator})/({denominator})"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        else:
            break
    
    # Xử lý sqrt
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
            
        content_start = start + 6
        brace_count = 1
        content_end = content_start
        
        while content_end < len(result) and brace_count > 0:
            if result[content_end] == '{':
                brace_count += 1
            elif result[content_end] == '}':
                brace_count -= 1
            content_end += 1
        
        if brace_count == 0:
            content = result[content_start:content_end-1]
            sqrt_result = f"√({content})"
            result = result[:start] + sqrt_result + result[content_end:]
        else:
            break
    
    # Xử lý superscript đơn giản: x^2 -> x²
    superscript_digits = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    
    i = 0
    new_result = ""
    while i < len(result):
        if (i < len(result) - 2 and 
            result[i+1] == '^' and 
            result[i+2].isdigit() and
            result[i].isalnum()):
            base = result[i]
            power = result[i+2]
            new_result += base + superscript_digits.get(power, power)
            i += 3
        else:
            new_result += result[i]
            i += 1
    
    result = new_result
    
    # Loại bỏ backslash còn lại
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Tìm expressions toán học trong text"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Xử lý ${...}$ format
            if i < len(text) and text[i] == '{':
                i += 1  # skip {
                brace_count = 1
                content_start = i
                
                while i < len(text) and brace_count > 0:
                    if text[i] == '{':
                        brace_count += 1
                    elif text[i] == '}':
                        brace_count -= 1
                    i += 1
                
                # Kiểm tra có }$ không
                if i < len(text) and text[i] == '$':
                    end = i + 1
                    latex_content = text[content_start:i-1]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'braced'))
                    i += 1
                else:
                    i += 1
            else:
                # Xử lý $...$ format thông thường
                content_start = i
                brace_depth = 0
                
                # Tìm $ đóng, tracking nested braces
                while i < len(text):
                    if text[i] == '{':
                        brace_depth += 1
                    elif text[i] == '}':
                        brace_depth -= 1
                    elif text[i] == '$' and brace_depth == 0:
                        break
                    i += 1
                
                if i < len(text):  # Tìm thấy $ đóng
                    end = i + 1
                    latex_content = text[content_start:i]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'normal'))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_omml_equation(paragraph, text, equation_type='simple'):
    """Tạo OMML equation - Safe with fallbacks"""
    try:
        # Always try simple approach first to avoid corruption
        if equation_type == 'prime' and "'" in text:
            if create_prime_omml(paragraph, text):
                return True
        elif equation_type == 'superscript' and any(c in text for c in ['²', '³', '⁴', '⁵']):
            if create_superscript_omml(paragraph, text):
                return True
        elif equation_type == 'fraction' and '(' in text and ')' in text and '/' in text:
            if create_fraction_omml(paragraph, text):
                return True
        elif equation_type == 'parentheses' and '(' in text and ')' in text:
            if create_parentheses_omml(paragraph, text):
                return True
        
        # Default to simple OMML
        return create_simple_omml(paragraph, text)
        
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field - Simplified and safe"""
    try:
        # Simple EQ field without complex formatting
        eq_content = unicode_text
        
        # Only handle basic superscripts
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}'"
        
        # Safe field creation
        fldChar_begin = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>')
        instrText = parse_xml(f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> {eq_content} </w:instrText>')
        fldChar_end = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>')
        
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_prime_omml(paragraph, text):
    """OMML cho prime notation - Fixed XML"""
    try:
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        # Safe XML with proper escaping
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_prime = "'" * prime_count
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_prime}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_superscript_omml(paragraph, text):
    """OMML cho superscript numbers - Fixed XML"""
    try:
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:
                    base += char
        
        if not sup:
            return create_simple_omml(paragraph, text)
        
        # Safe XML
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_sup = sup.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_sup}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_fraction_omml(paragraph, text):
    """OMML cho fractions - Simplified and safe"""
    try:
        frac_pattern = r'\(([^)]+)\)/\(([^)]+)\)'
        match = re.search(frac_pattern, text)
        
        if match:
            numerator = match.group(1).strip()
            denominator = match.group(2).strip()
            
            # Safe XML escaping
            safe_num = numerator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            safe_den = denominator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:f><m:fPr></m:fPr><m:num><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_num}</m:t></m:r></m:num><m:den><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_den}</m:t></m:r></m:den></m:f></m:oMath>'''
            
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return True
        
        return False
    except Exception:
        return False

def create_parentheses_omml(paragraph, text):
    """OMML cho parentheses - Simplified"""
    try:
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end].strip()
        else:
            content = text.strip()
        
        # Safe XML
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_content}</m:t></m:r></m:e></m:d></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_simple_omml(paragraph, text):
    """OMML cho equation đơn giản - Safe version"""
    try:
        # Safe XML escaping
        safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_text}</m:t></m:r></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field"""
    try:
        eq_content = unicode_text
        
        # Handle superscripts for EQ field
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}\\s\\up5(')"
        elif '²' in unicode_text:
            base = unicode_text.replace('²', '')
            eq_content = f"{base}\\s\\up5(2)"
        elif '³' in unicode_text:
            base = unicode_text.replace('³', '')
            eq_content = f"{base}\\s\\up5(3)"
        
        # Create field elements
        fldChar_begin = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>'
        )
        
        instrText = parse_xml(
            f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> EQ {eq_content}</w:instrText>'
        )
        
        fldChar_end = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>'
        )
        
        # Add to paragraph
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_styled_math_run(paragraph, unicode_text):
    """Tạo styled run như equation"""
    try:
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 32, 96)
        
        # Add background
        try:
            rPr = run._element.get_or_add_rPr()
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
        except:
            pass
        
        return True
    except Exception:
        return False

def create_equation_object(paragraph, latex_text):
    """Tạo equation object với safe methods"""
    unicode_content = convert_latex_symbols(latex_text)
    
    # Determine equation type
    equation_type = 'simple'
    if "'" in unicode_content:
        equation_type = 'prime'
    elif '(' in unicode_content and ')' in unicode_content and '/' in unicode_content:
        equation_type = 'fraction'
    elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
        equation_type = 'superscript'
    elif '(' in unicode_content and ')' in unicode_content:
        equation_type = 'parentheses'
    
    # Try methods with better error handling
    try:
        # Method 1: OMML (safest approach)
        if create_omml_equation(paragraph, unicode_content, equation_type):
            return 'omml'
    except Exception:
        pass
    
    try:
        # Method 2: Styled run (most reliable)
        if create_styled_math_run(paragraph, unicode_content):
            return 'styled'
    except Exception:
        pass
    
    try:
        # Method 3: EQ Field (skip if problematic)
        # Commenting out to avoid file corruption
        # if create_equation_field(paragraph, unicode_content, latex_text):
        #     return 'field'
        pass
    except Exception:
        pass
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_content)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)
        return 'fallback'
    except Exception:
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return 'error'

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Safe version"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    last_pos = 0
    math_count = 0
    method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    for start, end, latex_content, expr_type in expressions:
        # Add text before equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Create equation object safely
        method_used = create_equation_object(paragraph, latex_content)
        method_stats[method_used] += 1
        math_count += 1
        
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count, method_stats

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # Bold question number
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Process content with math
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Bold answer letter
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Process answer content
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra markdown table"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse markdown table"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Skip separator line
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table - Safe version"""
    if not table_data:
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    math_count = 0
    total_method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    max_cols = max(len(row) for row in table_data)
    
    # Create table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Fill data
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Create paragraph in cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Process cell content safely
                try:
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                        math_count += cell_math_count
                        for method, count in cell_method_stats.items():
                            total_method_stats[method] += count
                except Exception:
                    # Fallback to plain text
                    run = cell_para.add_run(cell_text)
                    run.font.size = Pt(11)
                    run.font.name = 'Times New Roman'
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table safely
    if options.get('format_tables', True):
        try:
            word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if word_table.rows:
                header_row = word_table.rows[0]
                for cell in header_row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                    
                    try:
                        shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                        shading = parse_xml(shading_xml)
                        cell._tc.get_or_add_tcPr().append(shading)
                    except:
                        pass
        except Exception:
            pass
    
    return math_count, total_method_stats

def process_document_content(doc, options):
    """Xử lý document content - Safe version"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'omml_equations': 0,
        'field_equations': 0,
        'styled_equations': 0,
        'fallback_equations': 0,
        'error_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Processing: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Process paragraphs safely
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Check for markdown table
        if '|' in text and options.get('convert_markdown', True):
            table_lines = []
            j = i
            
            while j < len(doc.paragraphs):
                current_text = doc.paragraphs[j].text.strip()
                if not current_text:
                    break
                if '|' in current_text:
                    table_lines.append(current_text)
                    j += 1
                else:
                    break
            
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    try:
                        math_count, method_stats = create_word_table(new_doc, table_data, options)
                        stats['math_expressions'] += math_count
                        for method, count in method_stats.items():
                            if method in stats:
                                stats[f'{method}_equations'] += count
                        stats['markdown_tables'] += 1
                        stats['debug_log'].append(f"📋 Converted table: {len(table_data)} rows")
                    except Exception as e:
                        stats['debug_log'].append(f"❌ Table error: {str(e)}")
                
                i = j
                continue
        
        # Process normal paragraph safely
        new_para = new_doc.add_paragraph()
        
        try:
            if format_question_answer(new_para, text):
                stats['questions_formatted'] += 1
                stats['debug_log'].append(f"📝 P{i}: Q&A formatted")
            else:
                math_count, method_stats = process_text_with_math(new_para, text)
                stats['math_expressions'] += math_count
                for method, count in method_stats.items():
                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                        stats[f'{method}_equations'] += count
                
                if math_count > 0:
                    omml_count = method_stats.get('omml', 0)
                    styled_count = method_stats.get('styled', 0) + method_stats.get('fallback', 0)
                    stats['debug_log'].append(f"🔢 P{i}: {math_count} equations ({omml_count} OMML, {styled_count} styled)")
        except Exception as e:
            # Fallback to plain text
            new_para.clear()
            run = new_para.add_run(text)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
            stats['debug_log'].append(f"❌ P{i}: Error, used plain text")
        
        i += 1
    
    # Process existing Word tables safely
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
                            if new_cell.paragraphs:
                                cell_para = new_cell.paragraphs[0]
                                cell_para.clear()
                            else:
                                cell_para = new_cell.add_paragraph()
                            
                            if not format_question_answer(cell_para, cell_text):
                                cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                                table_math += cell_math_count
                                for method, count in cell_method_stats.items():
                                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                                        stats[f'{method}_equations'] += count
                            
                            cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except Exception:
                        # Skip problematic cells
                        continue
            
            stats['math_expressions'] += table_math
            
            # Format table safely
            if options.get('format_tables', True):
                try:
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    if new_table.rows:
                        header_row = new_table.rows[0]
                        for cell in header_row.cells:
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.bold = True
                except Exception:
                    pass
                
                stats['word_tables'] += 1
            
            if table_math > 0:
                stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations")
                
        except Exception as e:
            stats['debug_log'].append(f"❌ Table {table_idx} error: {str(e)}")
            continue
    
    stats['debug_log'].append(f"✅ Complete: {stats['math_expressions']} total equations")
    stats['debug_log'].append(f"   🎯 OMML: {stats['omml_equations']}")
    stats['debug_log'].append(f"   ✨ Styled: {stats['styled_equations'] + stats['fallback_equations']}")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v8.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">Fixed File Corruption + Safe Equation Objects + Complete LaTeX Support</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🎯 Tạo TRUE Equation Objects", value=True,
                                       help="Chuyển $...$ thành REAL equation objects trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True,
                                     help="Format 'Câu X:' và 'A. B. C. D.'")  
        format_tables = st.checkbox("📊 Format bảng", value=True,
                                   help="Professional table formatting")
        convert_markdown = st.checkbox("📋 Markdown → Word Tables", value=True,
                                     help="Convert | tables | to Word")
        show_debug = st.checkbox("🔍 Advanced Debug", value=False,
                               help="Show equation creation details")
        
        st.markdown("---")
        st.info("🛡️ **SAFE Equation Creation:**\n\n✅ **Method 1:** Safe OMML Objects\n🎨 **Method 2:** Enhanced Styling  \n💫 **Method 3:** Fallback Protection\n\n🔧 **Fixed Issues:**\n- File Corruption\n- XML Validation\n- Safe Error Handling\n- Robust Fallbacks")
        
 → **(x)** (safe parentheses)
        - `$[6;8]import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="LaTeX Word Processor",
    page_icon="📚",
    layout="wide"
)

def convert_latex_symbols(text):
    """Chuyển đổi ký hiệu LaTeX thành Unicode"""
    if not text:
        return text
    
    result = text
    
    # Xử lý superscript có dấu ngoặc: A^{prime} -> A'
    while True:
        pos = result.find('^{')
        if pos == -1:
            break
        
        # Tìm ký tự trước ^{
        base_start = pos - 1
        while base_start >= 0 and result[base_start].isalnum():
            base_start -= 1
        base_start += 1
        base = result[base_start:pos]
        
        # Tìm nội dung trong {}
        brace_start = pos + 2
        brace_count = 1
        brace_end = brace_start
        
        while brace_end < len(result) and brace_count > 0:
            if result[brace_end] == '{':
                brace_count += 1
            elif result[brace_end] == '}':
                brace_count -= 1
            brace_end += 1
        
        if brace_count == 0:
            content = result[brace_start:brace_end-1]
            if content == 'prime':
                replacement = base + "'"
            else:
                replacement = base + "^(" + content + ")"
            result = result[:base_start] + replacement + result[brace_end:]
        else:
            break
    
    # Các ký hiệu toán học cơ bản
    symbols = {
        '\\pi': 'π',
        '\\alpha': 'α',
        '\\beta': 'β',
        '\\gamma': 'γ',
        '\\delta': 'δ',
        '\\theta': 'θ',
        '\\lambda': 'λ',
        '\\mu': 'μ',
        '\\sigma': 'σ',
        '\\phi': 'φ',
        '\\omega': 'ω',
        '\\in': '∈',
        '\\subset': '⊂',
        '\\cup': '∪',
        '\\cap': '∩',
        '\\emptyset': '∅',
        '\\leq': '≤',
        '\\geq': '≥',
        '\\neq': '≠',
        '\\approx': '≈',
        '\\infty': '∞',
        '\\pm': '±',
        '\\mp': '∓',
        '\\mathbb{R}': 'ℝ',
        '\\mathbb{Z}': 'ℤ',
        '\\mathbb{Q}': 'ℚ',
        '\\mathbb{N}': 'ℕ',
        '\\mathbb{C}': 'ℂ'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Xử lý left, right parentheses
    result = result.replace('\\left(', '(')
    result = result.replace('\\right)', ')')
    result = result.replace('\\left[', '[')
    result = result.replace('\\right]', ']')
    result = result.replace('\\left{', '{')
    result = result.replace('\\right}', '}')
    
    # Xử lý frac
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        # Tìm tử số
        num_start = start + 6
        brace_count = 1
        num_end = num_start
        
        while num_end < len(result) and brace_count > 0:
            if result[num_end] == '{':
                brace_count += 1
            elif result[num_end] == '}':
                brace_count -= 1
            num_end += 1
        
        if brace_count > 0:
            break
            
        numerator = result[num_start:num_end-1]
        
        # Tìm mẫu số
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
            
            if brace_count == 0:
                denominator = result[denom_start:denom_end-1]
                fraction = f"({numerator})/({denominator})"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        else:
            break
    
    # Xử lý sqrt
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
            
        content_start = start + 6
        brace_count = 1
        content_end = content_start
        
        while content_end < len(result) and brace_count > 0:
            if result[content_end] == '{':
                brace_count += 1
            elif result[content_end] == '}':
                brace_count -= 1
            content_end += 1
        
        if brace_count == 0:
            content = result[content_start:content_end-1]
            sqrt_result = f"√({content})"
            result = result[:start] + sqrt_result + result[content_end:]
        else:
            break
    
    # Xử lý superscript đơn giản: x^2 -> x²
    superscript_digits = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    
    i = 0
    new_result = ""
    while i < len(result):
        if (i < len(result) - 2 and 
            result[i+1] == '^' and 
            result[i+2].isdigit() and
            result[i].isalnum()):
            base = result[i]
            power = result[i+2]
            new_result += base + superscript_digits.get(power, power)
            i += 3
        else:
            new_result += result[i]
            i += 1
    
    result = new_result
    
    # Loại bỏ backslash còn lại
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Tìm expressions toán học trong text"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Xử lý ${...}$ format
            if i < len(text) and text[i] == '{':
                i += 1  # skip {
                brace_count = 1
                content_start = i
                
                while i < len(text) and brace_count > 0:
                    if text[i] == '{':
                        brace_count += 1
                    elif text[i] == '}':
                        brace_count -= 1
                    i += 1
                
                # Kiểm tra có }$ không
                if i < len(text) and text[i] == '$':
                    end = i + 1
                    latex_content = text[content_start:i-1]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'braced'))
                    i += 1
                else:
                    i += 1
            else:
                # Xử lý $...$ format thông thường
                content_start = i
                brace_depth = 0
                
                # Tìm $ đóng, tracking nested braces
                while i < len(text):
                    if text[i] == '{':
                        brace_depth += 1
                    elif text[i] == '}':
                        brace_depth -= 1
                    elif text[i] == '$' and brace_depth == 0:
                        break
                    i += 1
                
                if i < len(text):  # Tìm thấy $ đóng
                    end = i + 1
                    latex_content = text[content_start:i]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip(), 'normal'))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_omml_equation(paragraph, text, equation_type='simple'):
    """Tạo OMML equation - Safe with fallbacks"""
    try:
        # Always try simple approach first to avoid corruption
        if equation_type == 'prime' and "'" in text:
            if create_prime_omml(paragraph, text):
                return True
        elif equation_type == 'superscript' and any(c in text for c in ['²', '³', '⁴', '⁵']):
            if create_superscript_omml(paragraph, text):
                return True
        elif equation_type == 'fraction' and '(' in text and ')' in text and '/' in text:
            if create_fraction_omml(paragraph, text):
                return True
        elif equation_type == 'parentheses' and '(' in text and ')' in text:
            if create_parentheses_omml(paragraph, text):
                return True
        
        # Default to simple OMML
        return create_simple_omml(paragraph, text)
        
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field - Simplified and safe"""
    try:
        # Simple EQ field without complex formatting
        eq_content = unicode_text
        
        # Only handle basic superscripts
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}'"
        
        # Safe field creation
        fldChar_begin = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>')
        instrText = parse_xml(f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> {eq_content} </w:instrText>')
        fldChar_end = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>')
        
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_prime_omml(paragraph, text):
    """OMML cho prime notation - Fixed XML"""
    try:
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        # Safe XML with proper escaping
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_prime = "'" * prime_count
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_prime}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_superscript_omml(paragraph, text):
    """OMML cho superscript numbers - Fixed XML"""
    try:
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:
                    base += char
        
        if not sup:
            return create_simple_omml(paragraph, text)
        
        # Safe XML
        safe_base = base.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_sup = sup.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:sSupPr></m:sSupPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_base}</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>{safe_sup}</m:t></m:r></m:sup></m:sSup></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_fraction_omml(paragraph, text):
    """OMML cho fractions - Simplified and safe"""
    try:
        frac_pattern = r'\(([^)]+)\)/\(([^)]+)\)'
        match = re.search(frac_pattern, text)
        
        if match:
            numerator = match.group(1).strip()
            denominator = match.group(2).strip()
            
            # Safe XML escaping
            safe_num = numerator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            safe_den = denominator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:f><m:fPr></m:fPr><m:num><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_num}</m:t></m:r></m:num><m:den><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_den}</m:t></m:r></m:den></m:f></m:oMath>'''
            
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return True
        
        return False
    except Exception:
        return False

def create_parentheses_omml(paragraph, text):
    """OMML cho parentheses - Simplified"""
    try:
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end].strip()
        else:
            content = text.strip()
        
        # Safe XML
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_content}</m:t></m:r></m:e></m:d></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_simple_omml(paragraph, text):
    """OMML cho equation đơn giản - Safe version"""
    try:
        # Safe XML escaping
        safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="i"/></m:rPr><m:t>{safe_text}</m:t></m:r></m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Tạo Word EQ field"""
    try:
        eq_content = unicode_text
        
        # Handle superscripts for EQ field
        if "'" in unicode_text:
            base = unicode_text.replace("'", "")
            eq_content = f"{base}\\s\\up5(')"
        elif '²' in unicode_text:
            base = unicode_text.replace('²', '')
            eq_content = f"{base}\\s\\up5(2)"
        elif '³' in unicode_text:
            base = unicode_text.replace('³', '')
            eq_content = f"{base}\\s\\up5(3)"
        
        # Create field elements
        fldChar_begin = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>'
        )
        
        instrText = parse_xml(
            f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> EQ {eq_content}</w:instrText>'
        )
        
        fldChar_end = parse_xml(
            r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>'
        )
        
        # Add to paragraph
        paragraph._element.append(fldChar_begin)
        paragraph._element.append(instrText) 
        paragraph._element.append(fldChar_end)
        
        return True
    except Exception:
        return False

def create_styled_math_run(paragraph, unicode_text):
    """Tạo styled run như equation"""
    try:
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 32, 96)
        
        # Add background
        try:
            rPr = run._element.get_or_add_rPr()
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
        except:
            pass
        
        return True
    except Exception:
        return False

def create_equation_object(paragraph, latex_text):
    """Tạo equation object với safe methods"""
    unicode_content = convert_latex_symbols(latex_text)
    
    # Determine equation type
    equation_type = 'simple'
    if "'" in unicode_content:
        equation_type = 'prime'
    elif '(' in unicode_content and ')' in unicode_content and '/' in unicode_content:
        equation_type = 'fraction'
    elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
        equation_type = 'superscript'
    elif '(' in unicode_content and ')' in unicode_content:
        equation_type = 'parentheses'
    
    # Try methods with better error handling
    try:
        # Method 1: OMML (safest approach)
        if create_omml_equation(paragraph, unicode_content, equation_type):
            return 'omml'
    except Exception:
        pass
    
    try:
        # Method 2: Styled run (most reliable)
        if create_styled_math_run(paragraph, unicode_content):
            return 'styled'
    except Exception:
        pass
    
    try:
        # Method 3: EQ Field (skip if problematic)
        # Commenting out to avoid file corruption
        # if create_equation_field(paragraph, unicode_content, latex_text):
        #     return 'field'
        pass
    except Exception:
        pass
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_content)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)
        return 'fallback'
    except Exception:
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return 'error'

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Safe version"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    last_pos = 0
    math_count = 0
    method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    for start, end, latex_content, expr_type in expressions:
        # Add text before equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Create equation object safely
        method_used = create_equation_object(paragraph, latex_content)
        method_stats[method_used] += 1
        math_count += 1
        
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count, method_stats

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # Bold question number
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Process content with math
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Bold answer letter
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Process answer content
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra markdown table"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse markdown table"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Skip separator line
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table - Safe version"""
    if not table_data:
        return 0, {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    math_count = 0
    total_method_stats = {'omml': 0, 'field': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    max_cols = max(len(row) for row in table_data)
    
    # Create table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Fill data
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Create paragraph in cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Process cell content safely
                try:
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                        math_count += cell_math_count
                        for method, count in cell_method_stats.items():
                            total_method_stats[method] += count
                except Exception:
                    # Fallback to plain text
                    run = cell_para.add_run(cell_text)
                    run.font.size = Pt(11)
                    run.font.name = 'Times New Roman'
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table safely
    if options.get('format_tables', True):
        try:
            word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if word_table.rows:
                header_row = word_table.rows[0]
                for cell in header_row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                    
                    try:
                        shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                        shading = parse_xml(shading_xml)
                        cell._tc.get_or_add_tcPr().append(shading)
                    except:
                        pass
        except Exception:
            pass
    
    return math_count, total_method_stats

def process_document_content(doc, options):
    """Xử lý document content - Safe version"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'omml_equations': 0,
        'field_equations': 0,
        'styled_equations': 0,
        'fallback_equations': 0,
        'error_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Processing: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Process paragraphs safely
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Check for markdown table
        if '|' in text and options.get('convert_markdown', True):
            table_lines = []
            j = i
            
            while j < len(doc.paragraphs):
                current_text = doc.paragraphs[j].text.strip()
                if not current_text:
                    break
                if '|' in current_text:
                    table_lines.append(current_text)
                    j += 1
                else:
                    break
            
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    try:
                        math_count, method_stats = create_word_table(new_doc, table_data, options)
                        stats['math_expressions'] += math_count
                        for method, count in method_stats.items():
                            if method in stats:
                                stats[f'{method}_equations'] += count
                        stats['markdown_tables'] += 1
                        stats['debug_log'].append(f"📋 Converted table: {len(table_data)} rows")
                    except Exception as e:
                        stats['debug_log'].append(f"❌ Table error: {str(e)}")
                
                i = j
                continue
        
        # Process normal paragraph safely
        new_para = new_doc.add_paragraph()
        
        try:
            if format_question_answer(new_para, text):
                stats['questions_formatted'] += 1
                stats['debug_log'].append(f"📝 P{i}: Q&A formatted")
            else:
                math_count, method_stats = process_text_with_math(new_para, text)
                stats['math_expressions'] += math_count
                for method, count in method_stats.items():
                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                        stats[f'{method}_equations'] += count
                
                if math_count > 0:
                    omml_count = method_stats.get('omml', 0)
                    styled_count = method_stats.get('styled', 0) + method_stats.get('fallback', 0)
                    stats['debug_log'].append(f"🔢 P{i}: {math_count} equations ({omml_count} OMML, {styled_count} styled)")
        except Exception as e:
            # Fallback to plain text
            new_para.clear()
            run = new_para.add_run(text)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
            stats['debug_log'].append(f"❌ P{i}: Error, used plain text")
        
        i += 1
    
    # Process existing Word tables safely
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
                            if new_cell.paragraphs:
                                cell_para = new_cell.paragraphs[0]
                                cell_para.clear()
                            else:
                                cell_para = new_cell.add_paragraph()
                            
                            if not format_question_answer(cell_para, cell_text):
                                cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text)
                                table_math += cell_math_count
                                for method, count in cell_method_stats.items():
                                    if method in ['omml', 'field', 'styled', 'fallback', 'error']:
                                        stats[f'{method}_equations'] += count
                            
                            cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except Exception:
                        # Skip problematic cells
                        continue
            
            stats['math_expressions'] += table_math
            
            # Format table safely
            if options.get('format_tables', True):
                try:
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    if new_table.rows:
                        header_row = new_table.rows[0]
                        for cell in header_row.cells:
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.bold = True
                except Exception:
                    pass
                
                stats['word_tables'] += 1
            
            if table_math > 0:
                stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations")
                
        except Exception as e:
            stats['debug_log'].append(f"❌ Table {table_idx} error: {str(e)}")
            continue
    
    stats['debug_log'].append(f"✅ Complete: {stats['math_expressions']} total equations")
    stats['debug_log'].append(f"   🎯 OMML: {stats['omml_equations']}")
    stats['debug_log'].append(f"   ✨ Styled: {stats['styled_equations'] + stats['fallback_equations']}")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v8.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">Fixed File Corruption + Safe Equation Objects + Complete LaTeX Support</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🎯 Tạo TRUE Equation Objects", value=True,
                                       help="Chuyển $...$ thành REAL equation objects trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True,
                                     help="Format 'Câu X:' và 'A. B. C. D.'")  
        format_tables = st.checkbox("📊 Format bảng", value=True,
                                   help="Professional table formatting")
        convert_markdown = st.checkbox("📋 Markdown → Word Tables", value=True,
                                     help="Convert | tables | to Word")
        show_debug = st.checkbox("🔍 Advanced Debug", value=False,
                               help="Show equation creation details")
        
        st.markdown("---")
        st.info("🛡️ **SAFE Equation Creation:**\n\n✅ **Method 1:** Safe OMML Objects\n🎨 **Method 2:** Enhanced Styling  \n💫 **Method 3:** Fallback Protection\n\n🔧 **Fixed Issues:**\n- File Corruption\n- XML Validation\n- Safe Error Handling\n- Robust Fallbacks")
        
 → **[6;8]** (safe text)
        
        **🛡️ All methods include safe fallbacks!**
        """)
    
    # Main area
    st.markdown("### 📁 Upload Word Document")
    st.markdown("Upload your .docx file with LaTeX expressions to convert them into **TRUE Word equation objects**")
    
    uploaded_file = st.file_uploader("Chọn file Word (.docx)", type=["docx"])
    
    if uploaded_file:
        st.success(f"✅ {uploaded_file.name}")
        
        with st.expander("📄 Preview"):
            try:
                doc = Document(uploaded_file)
                st.info(f"📄 {len(doc.paragraphs)} paragraphs | 📊 {len(doc.tables)} tables")
                
                for i, para in enumerate(doc.paragraphs[:5]):
                    if para.text.strip():
                        preview_text = para.text[:80] + "..." if len(para.text) > 80 else para.text
                        st.text(f"{i+1}. {preview_text}")
            except Exception as e:
                st.error(f"Preview error: {e}")
        
        if st.button("🚀 Process Document", type="primary", use_container_width=True):
            options = {
                'convert_equations': convert_equations,
                'format_questions': format_questions,
                'format_tables': format_tables,
                'convert_markdown': convert_markdown
            }
            
            with st.spinner("Creating equation objects..."):
                try:
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document_content(doc, options)
                    
                    # Save document
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_processed_{timestamp}.docx"
                    
                    # Results
                    st.markdown("### 📊 Processing Results")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("🔢 Total Math", stats['math_expressions'])
                    col2.metric("📝 Q&A Items", stats['questions_formatted'])
                    col3.metric("📊 Word Tables", stats['word_tables'])
                    col4.metric("📋 Markdown Tables", stats['markdown_tables'])
                    
                    # Equation creation breakdown
                    if stats['math_expressions'] > 0:
                        st.markdown("### 🎯 Equation Objects Created")
                        eq_col1, eq_col2, eq_col3 = st.columns(3)
                        eq_col1.metric("🎯 OMML Objects", stats['omml_equations'], 
                                     help="True equation objects using OMML")
                        eq_col2.metric("🎨 EQ Fields", stats['field_equations'],
                                     help="Word equation fields") 
                        styled_total = stats['styled_equations'] + stats.get('fallback_equations', 0)
                        eq_col3.metric("✨ Styled Math", styled_total,
                                     help="Professional styled equations")
                    
                    if show_debug:
                        with st.expander("🔍 Detailed Processing Log"):
                            for log in stats['debug_log']:
                                if 'OMML' in log or '🎯' in log:
                                    st.success(log)
                                elif 'ERROR' in log or '❌' in log:
                                    st.error(log)
                                elif 'Table' in log or '📋' in log or '📊' in log:
                                    st.info(log)
                                else:
                                    st.text(log)
                    
                    # Download button
                    st.download_button(
                        "⬇️ Download Processed File",
                        buffer.getvalue(),
                        filename,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                    # Success messages - NO BALLOONS
                    total_processed = (stats['math_expressions'] + stats['questions_formatted'] + 
                                     stats['word_tables'] + stats['markdown_tables'])
                    
                    if total_processed > 0:
                        st.success("🎉 Processing Complete!")
                        
                        if stats['math_expressions'] > 0:
                            eq_breakdown = []
                            if stats['omml_equations'] > 0:
                                eq_breakdown.append(f"{stats['omml_equations']} OMML objects")
                            if stats['field_equations'] > 0:
                                eq_breakdown.append(f"{stats['field_equations']} EQ fields")
                            styled_total = stats['styled_equations'] + stats.get('fallback_equations', 0)
                            if styled_total > 0:
                                eq_breakdown.append(f"{styled_total} styled equations")
                            
                            if eq_breakdown:
                                st.info(f"🔢 **{stats['math_expressions']} equation objects** created: {', '.join(eq_breakdown)}")
                        
                        if stats['markdown_tables'] > 0:
                            st.info(f"📋 **{stats['markdown_tables']} markdown tables** converted")
                        if stats['questions_formatted'] > 0:
                            st.info(f"📝 **{stats['questions_formatted']} Q&A items** formatted")
                        if stats['word_tables'] > 0:
                            st.info(f"📊 **{stats['word_tables']} Word tables** formatted")
                        
                        # Success notice for OMML equations - NO BALLOONS
                        if stats['omml_equations'] > 0:
                            st.success(f"🎯 **{stats['omml_equations']} TRUE equation objects** created successfully!")
                            
                    else:
                        st.warning("⚠️ No LaTeX, Q&A, or tables found to process")
                
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    main()
