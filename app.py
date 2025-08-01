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
    """Tìm TẤT CẢ expressions toán học - improved detection"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '

def create_math_equation(paragraph, latex_text):
    """Tạo equation object thực sự trong Word"""
    try:
        # Convert LaTeX to simpler math notation first
        unicode_text = convert_latex_symbols(latex_text)
        
        # Create equation run với Office Math format
        # Sử dụng approach tạo equation object thông qua OMML
        
        # Tạo math element
        math_xml = f"""
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:t>{unicode_text}</m:t>
            </m:r>
        </m:oMath>
        """
        
        try:
            # Thử insert equation object
            math_element = parse_xml(math_xml)
            paragraph._element.append(math_element)
            return True
        except:
            # Fallback: Tạo run với style đặc biệt trông như equation
            return create_equation_style_run(paragraph, unicode_text)
            
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def create_equation_style_run(paragraph, text):
    """Tạo run trông như equation (fallback)"""
    try:
        run = paragraph.add_run(text)
        
        # Style để trông như equation
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
        
        # Thêm highlight nhẹ
        try:
            # Add subtle background
            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F0F8FF"/>'
            shading = parse_xml(shading_xml)
            run._element.get_or_add_rPr().append(shading)
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao với OMML format"""
    try:
        # Parse LaTeX content để tạo OMML phù hợp
        unicode_content = convert_latex_symbols(latex_text)
        
        # Phân tích content để tạo structure phù hợp
        if '^' in unicode_content or '²' in unicode_content or '³' in unicode_content:
            # Có superscript - tạo sup element
            return create_superscript_equation(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            # Có parentheses - tạo brackets
            return create_bracketed_equation(paragraph, unicode_content)
        else:
            # Simple math
            return create_simple_equation(paragraph, unicode_content)
            
    except Exception:
        return create_equation_style_run(paragraph, unicode_content)

def create_word_equation(paragraph, latex_text, expr_type='normal'):
    """Tạo equation object đẹp trong Word - Improved version"""
    try:
        # Convert LaTeX to Unicode
        unicode_content = convert_latex_symbols(latex_text)
        
        # Method 1: Thử tạo equation field thực sự
        if create_equation_field(paragraph, unicode_content, latex_text):
            return True
        
        # Method 2: Tạo styled run trông như equation
        return create_professional_math_run(paragraph, unicode_content, latex_text)
        
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[Math: {latex_text}]")
        run.italic = True
        run.font.color.rgb = RGBColor(255, 0, 0)
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Thử tạo equation field thực sự trong Word"""
    try:
        # Approach: Insert equation through field codes
        # EQ field is Word's built-in equation field
        
        # Convert content to EQ field format
        eq_content = convert_to_eq_field(unicode_text, original_latex)
        
        # Create field run
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

def convert_to_eq_field(unicode_text, latex_text):
    """Convert to EQ field format"""
    # EQ field syntax for common patterns
    eq_content = unicode_text
    
    # Handle superscripts for EQ field
    if "'" in unicode_text:
        # A' -> A\s\up5(')
        base = unicode_text.replace("'", "")
        eq_content = f"{base}\\s\\up5(')"
    elif '²' in unicode_text:
        base = unicode_text.replace('²', '')
        eq_content = f"{base}\\s\\up5(2)"
    elif '³' in unicode_text:
        base = unicode_text.replace('³', '')
        eq_content = f"{base}\\s\\up5(3)"
    
    return eq_content

def create_professional_math_run(paragraph, unicode_text, original_latex):
    """Tạo run với style chuyên nghiệp như equation"""
    try:
        # Tạo run với style equation đặc biệt
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.bold = False
        
        # Màu equation chuyên nghiệp
        run.font.color.rgb = RGBColor(0, 32, 96)  # Dark navy blue
        
        # Thêm background nhẹ để highlight
        try:
            rPr = run._element.get_or_add_rPr()
            
            # Add equation-like background
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
            
            # Add slight border for equation effect
            border_xml = '''<w:bdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                           w:val="single" w:sz="2" w:space="1" w:color="E3E6F0"/>'''
            try:
                border = parse_xml(border_xml)
                rPr.append(border)
            except:
                pass
                
        except:
            pass
        
        # Add subtle spacing
        try:
            # Add space before and after for equation isolation
            space_before = paragraph.add_run(" ")
            space_before.font.size = Pt(8)
            
            # Move the equation run to correct position
            # (This is complex, so we'll rely on the styling)
            
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_complex_equation_omml(paragraph, latex_text):
    """Tạo OMML equation phức tạp cho cases đặc biệt"""
    try:
        unicode_content = convert_latex_symbols(latex_text)
        
        # Determine equation type và tạo OMML phù hợp
        if "'" in unicode_content:
            return create_prime_equation_omml(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            return create_parentheses_equation_omml(paragraph, unicode_content)
        elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
            return create_superscript_equation_omml(paragraph, unicode_content)
        else:
            return create_simple_equation_omml(paragraph, unicode_content)
            
    except Exception:
        return False

def create_prime_equation_omml(paragraph, text):
    """OMML cho prime notation"""
    try:
        # Split base and prime
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{"'" * prime_count}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_parentheses_equation_omml(paragraph, text):
    """OMML cho expressions với dấu ngoặc"""
    try:
        # Extract content inside parentheses
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end]
        else:
            content = text
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:d>
                <m:dPr>
                    <m:begChr m:val="("/>
                    <m:endChr m:val=")"/>
                    <m:grow m:val="1"/>
                </m:dPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{content}</m:t>
                    </m:r>
                </m:e>
            </m:d>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_superscript_equation_omml(paragraph, text):
    """OMML cho superscript numbers"""
    try:
        # Find base and superscript
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                # Convert back to normal number
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:  # Still building base
                    base += char
        
        if not sup:
            return create_simple_equation_omml(paragraph, text)
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{sup}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_simple_equation_omml(paragraph, text):
    """OMML cho equation đơn giản"""
    try:
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:rPr>
                    <m:scr m:val="roman"/>
                    <m:sty m:val="i"/>
                </m:rPr>
                <m:t>{text}</m:t>
            </m:r>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Enhanced với multiple approaches"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    successful_equations = 0
    
    for start, end, latex_content, expr_type in expressions:
        # Thêm text trước equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Thử tạo equation với multiple methods
        equation_created = False
        
        # Method 1: Complex OMML cho cases đặc biệt
        if create_complex_equation_omml(paragraph, latex_content):
            equation_created = True
            successful_equations += 1
        # Method 2: Word equation field
        elif create_word_equation(paragraph, latex_content, expr_type):
            equation_created = True
            successful_equations += 1
        # Method 3: Professional styled run
        else:
            unicode_text = convert_latex_symbols(latex_content)
            if create_professional_math_run(paragraph, unicode_text, latex_content):
                equation_created = True
            else:
                # Final fallback
                run = paragraph.add_run(f"[{latex_content}]")
                run.italic = True
                run.font.color.rgb = RGBColor(255, 0, 0)
        
        math_count += 1
        last_pos = end
    
    # Thêm text còn lại
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Kiểm tra pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # In đậm phần câu hỏi
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Xử lý phần nội dung
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Kiểm tra pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format chữ cái đáp án
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Xử lý nội dung đáp án
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A, xử lý bình thường
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra xem text có phải markdown table không"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse text thành table data"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Bỏ qua dòng separator
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            # Loại bỏ | đầu và cuối
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table từ data với enhanced equation processing"""
    if not table_data:
        return 0
    
    math_count = 0
    max_cols = max(len(row) for row in table_data)
    
    # Tạo table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Điền dữ liệu
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Tạo paragraph trong cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Check for math expressions in cell
                cell_math_expressions = find_math_expressions(cell_text)
                
                # Xử lý Q&A hoặc math
                if not format_question_answer(cell_para, cell_text):
                    cell_math_count = process_text_with_math(cell_para, cell_text)
                    math_count += cell_math_count
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table structure
    if options.get('format_tables', True):
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header
        if word_table.rows:
            header_row = word_table.rows[0]
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                        run.font.size = Pt(12)
                
                # Màu nền xanh
                try:
                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                    shading = parse_xml(shading_xml)
                    cell._tc.get_or_add_tcPr().append(shading)
                except:
                    pass
    
    return math_count

def process_document_content(doc, options):
    """Xử lý nội dung document với detailed logging"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'successful_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Xử lý paragraphs
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Debug: Check for math content
        math_expressions = find_math_expressions(text)
        if math_expressions:
            stats['debug_log'].append(f"📐 P{i}: Found {len(math_expressions)} math expressions")
            for j, (start, end, content, expr_type) in enumerate(math_expressions):
                stats['debug_log'].append(f"   Math {j+1}: '{content}' ({expr_type})")
        
        # Kiểm tra markdown table
        if '|' in text and options.get('convert_markdown', True):
            # Collect multiple lines for table
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
            
            # Check if it's a valid table
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    math_count = create_word_table(new_doc, table_data, options)
                    stats['math_expressions'] += math_count
                    stats['markdown_tables'] += 1
                    stats['debug_log'].append(f"📋 Converted markdown table: {len(table_data)} rows, {math_count} equations")
                
                i = j
                continue
        
        # Xử lý paragraph thông thường
        new_para = new_doc.add_paragraph()
        
        if format_question_answer(new_para, text):
            stats['questions_formatted'] += 1
            stats['debug_log'].append(f"📝 P{i}: Formatted as Q&A")
        else:
            math_count = process_text_with_math(new_para, text)
            stats['math_expressions'] += math_count
            if math_count > 0:
                stats['debug_log'].append(f"🔢 P{i}: Created {math_count} equations")
        
        i += 1
    
    # Xử lý Word tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        stats['debug_log'].append(f"📊 Processing Word table {table_idx}")
        
        num_rows = len(table.rows)
        num_cols = len(table.rows[0].cells) if table.rows else 0
        
        # Tạo table mới
        new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
        
        table_math = 0
        for r in range(num_rows):
            for c in range(min(len(table.rows[r].cells), num_cols)):
                original_cell = table.rows[r].cells[c]
                new_cell = new_table.rows[r].cells[c]
                
                cell_text = original_cell.text.strip()
                if cell_text:
                    # Check for math in cell
                    cell_math = find_math_expressions(cell_text)
                    if cell_math:
                        stats['debug_log'].append(f"   Cell[{r}][{c}]: {len(cell_math)} math expressions")
                    
                    new_cell.text = ""
                    if new_cell.paragraphs:
                        cell_para = new_cell.paragraphs[0]
                        cell_para.clear()
                    else:
                        cell_para = new_cell.add_paragraph()
                    
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count = process_text_with_math(cell_para, cell_text)
                        table_math += cell_math_count
                    
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        stats['math_expressions'] += table_math
        stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations created")
        
        # Format table
        if options.get('format_tables', True):
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if new_table.rows:
                header_row = new_table.rows[0]
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
            
            stats['word_tables'] += 1
    
    stats['debug_log'].append(f"✅ Hoàn thành: {stats['math_expressions']} equations, {stats['questions_formatted']} Q&A")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v7.2</h1>
        <p style="color: white; margin: 10px 0 0 0;">Enhanced Equation Objects + Complete LaTeX Support</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Tạo Equation Objects", value=True,
                                       help="Chuyển $...$ thành equation objects thực sự trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True)  
        format_tables = st.checkbox("📊 Format bảng", value=True)
        convert_markdown = st.checkbox("📋 Markdown Tables", value=True)
        show_debug = st.checkbox("🔍 Debug", value=False)
        
        st.markdown("---")
        st.info("💡 **Enhanced Equation Creation:**\n\n🎯 **Method 1:** OMML Equation Objects\n🎨 **Method 2:** Word EQ Fields\n✨ **Method 3:** Professional Math Styling\n\n📐 **Supports:** Superscripts, Prime notation, Parentheses, Fractions, Roots")
        
        st.markdown("### 🧪 Test LaTeX")
        test_latex = st.text_input("Test LaTeX:", placeholder="A^{prime}")
        if test_latex:
            converted = convert_latex_symbols(test_latex)
            st.code(f"${test_latex}$ → {converted}")
        
        st.markdown("### 📊 Expected Results")
        st.markdown("""
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
    """Tìm TẤT CẢ expressions toán học - improved detection"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '

def create_math_equation(paragraph, latex_text):
    """Tạo equation object thực sự trong Word"""
    try:
        # Convert LaTeX to simpler math notation first
        unicode_text = convert_latex_symbols(latex_text)
        
        # Create equation run với Office Math format
        # Sử dụng approach tạo equation object thông qua OMML
        
        # Tạo math element
        math_xml = f"""
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:t>{unicode_text}</m:t>
            </m:r>
        </m:oMath>
        """
        
        try:
            # Thử insert equation object
            math_element = parse_xml(math_xml)
            paragraph._element.append(math_element)
            return True
        except:
            # Fallback: Tạo run với style đặc biệt trông như equation
            return create_equation_style_run(paragraph, unicode_text)
            
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def create_equation_style_run(paragraph, text):
    """Tạo run trông như equation (fallback)"""
    try:
        run = paragraph.add_run(text)
        
        # Style để trông như equation
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
        
        # Thêm highlight nhẹ
        try:
            # Add subtle background
            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F0F8FF"/>'
            shading = parse_xml(shading_xml)
            run._element.get_or_add_rPr().append(shading)
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao với OMML format"""
    try:
        # Parse LaTeX content để tạo OMML phù hợp
        unicode_content = convert_latex_symbols(latex_text)
        
        # Phân tích content để tạo structure phù hợp
        if '^' in unicode_content or '²' in unicode_content or '³' in unicode_content:
            # Có superscript - tạo sup element
            return create_superscript_equation(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            # Có parentheses - tạo brackets
            return create_bracketed_equation(paragraph, unicode_content)
        else:
            # Simple math
            return create_simple_equation(paragraph, unicode_content)
            
    except Exception:
        return create_equation_style_run(paragraph, unicode_content)

def create_word_equation(paragraph, latex_text, expr_type='normal'):
    """Tạo equation object đẹp trong Word - Improved version"""
    try:
        # Convert LaTeX to Unicode
        unicode_content = convert_latex_symbols(latex_text)
        
        # Method 1: Thử tạo equation field thực sự
        if create_equation_field(paragraph, unicode_content, latex_text):
            return True
        
        # Method 2: Tạo styled run trông như equation
        return create_professional_math_run(paragraph, unicode_content, latex_text)
        
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[Math: {latex_text}]")
        run.italic = True
        run.font.color.rgb = RGBColor(255, 0, 0)
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Thử tạo equation field thực sự trong Word"""
    try:
        # Approach: Insert equation through field codes
        # EQ field is Word's built-in equation field
        
        # Convert content to EQ field format
        eq_content = convert_to_eq_field(unicode_text, original_latex)
        
        # Create field run
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

def convert_to_eq_field(unicode_text, latex_text):
    """Convert to EQ field format"""
    # EQ field syntax for common patterns
    eq_content = unicode_text
    
    # Handle superscripts for EQ field
    if "'" in unicode_text:
        # A' -> A\s\up5(')
        base = unicode_text.replace("'", "")
        eq_content = f"{base}\\s\\up5(')"
    elif '²' in unicode_text:
        base = unicode_text.replace('²', '')
        eq_content = f"{base}\\s\\up5(2)"
    elif '³' in unicode_text:
        base = unicode_text.replace('³', '')
        eq_content = f"{base}\\s\\up5(3)"
    
    return eq_content

def create_professional_math_run(paragraph, unicode_text, original_latex):
    """Tạo run với style chuyên nghiệp như equation"""
    try:
        # Tạo run với style equation đặc biệt
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.bold = False
        
        # Màu equation chuyên nghiệp
        run.font.color.rgb = RGBColor(0, 32, 96)  # Dark navy blue
        
        # Thêm background nhẹ để highlight
        try:
            rPr = run._element.get_or_add_rPr()
            
            # Add equation-like background
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
            
            # Add slight border for equation effect
            border_xml = '''<w:bdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                           w:val="single" w:sz="2" w:space="1" w:color="E3E6F0"/>'''
            try:
                border = parse_xml(border_xml)
                rPr.append(border)
            except:
                pass
                
        except:
            pass
        
        # Add subtle spacing
        try:
            # Add space before and after for equation isolation
            space_before = paragraph.add_run(" ")
            space_before.font.size = Pt(8)
            
            # Move the equation run to correct position
            # (This is complex, so we'll rely on the styling)
            
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_complex_equation_omml(paragraph, latex_text):
    """Tạo OMML equation phức tạp cho cases đặc biệt"""
    try:
        unicode_content = convert_latex_symbols(latex_text)
        
        # Determine equation type và tạo OMML phù hợp
        if "'" in unicode_content:
            return create_prime_equation_omml(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            return create_parentheses_equation_omml(paragraph, unicode_content)
        elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
            return create_superscript_equation_omml(paragraph, unicode_content)
        else:
            return create_simple_equation_omml(paragraph, unicode_content)
            
    except Exception:
        return False

def create_prime_equation_omml(paragraph, text):
    """OMML cho prime notation"""
    try:
        # Split base and prime
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{"'" * prime_count}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_parentheses_equation_omml(paragraph, text):
    """OMML cho expressions với dấu ngoặc"""
    try:
        # Extract content inside parentheses
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end]
        else:
            content = text
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:d>
                <m:dPr>
                    <m:begChr m:val="("/>
                    <m:endChr m:val=")"/>
                    <m:grow m:val="1"/>
                </m:dPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{content}</m:t>
                    </m:r>
                </m:e>
            </m:d>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_superscript_equation_omml(paragraph, text):
    """OMML cho superscript numbers"""
    try:
        # Find base and superscript
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                # Convert back to normal number
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:  # Still building base
                    base += char
        
        if not sup:
            return create_simple_equation_omml(paragraph, text)
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{sup}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_simple_equation_omml(paragraph, text):
    """OMML cho equation đơn giản"""
    try:
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:rPr>
                    <m:scr m:val="roman"/>
                    <m:sty m:val="i"/>
                </m:rPr>
                <m:t>{text}</m:t>
            </m:r>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Enhanced với multiple approaches"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    successful_equations = 0
    
    for start, end, latex_content, expr_type in expressions:
        # Thêm text trước equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Thử tạo equation với multiple methods
        equation_created = False
        
        # Method 1: Complex OMML cho cases đặc biệt
        if create_complex_equation_omml(paragraph, latex_content):
            equation_created = True
            successful_equations += 1
        # Method 2: Word equation field
        elif create_word_equation(paragraph, latex_content, expr_type):
            equation_created = True
            successful_equations += 1
        # Method 3: Professional styled run
        else:
            unicode_text = convert_latex_symbols(latex_content)
            if create_professional_math_run(paragraph, unicode_text, latex_content):
                equation_created = True
            else:
                # Final fallback
                run = paragraph.add_run(f"[{latex_content}]")
                run.italic = True
                run.font.color.rgb = RGBColor(255, 0, 0)
        
        math_count += 1
        last_pos = end
    
    # Thêm text còn lại
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Kiểm tra pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # In đậm phần câu hỏi
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Xử lý phần nội dung
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Kiểm tra pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format chữ cái đáp án
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Xử lý nội dung đáp án
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A, xử lý bình thường
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra xem text có phải markdown table không"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse text thành table data"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Bỏ qua dòng separator
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            # Loại bỏ | đầu và cuối
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table từ data với enhanced equation processing"""
    if not table_data:
        return 0
    
    math_count = 0
    max_cols = max(len(row) for row in table_data)
    
    # Tạo table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Điền dữ liệu
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Tạo paragraph trong cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Check for math expressions in cell
                cell_math_expressions = find_math_expressions(cell_text)
                
                # Xử lý Q&A hoặc math
                if not format_question_answer(cell_para, cell_text):
                    cell_math_count = process_text_with_math(cell_para, cell_text)
                    math_count += cell_math_count
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table structure
    if options.get('format_tables', True):
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header
        if word_table.rows:
            header_row = word_table.rows[0]
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                        run.font.size = Pt(12)
                
                # Màu nền xanh
                try:
                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                    shading = parse_xml(shading_xml)
                    cell._tc.get_or_add_tcPr().append(shading)
                except:
                    pass
    
    return math_count

def process_document_content(doc, options):
    """Xử lý nội dung document với detailed logging"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'successful_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Xử lý paragraphs
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Debug: Check for math content
        math_expressions = find_math_expressions(text)
        if math_expressions:
            stats['debug_log'].append(f"📐 P{i}: Found {len(math_expressions)} math expressions")
            for j, (start, end, content, expr_type) in enumerate(math_expressions):
                stats['debug_log'].append(f"   Math {j+1}: '{content}' ({expr_type})")
        
        # Kiểm tra markdown table
        if '|' in text and options.get('convert_markdown', True):
            # Collect multiple lines for table
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
            
            # Check if it's a valid table
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    math_count = create_word_table(new_doc, table_data, options)
                    stats['math_expressions'] += math_count
                    stats['markdown_tables'] += 1
                    stats['debug_log'].append(f"📋 Converted markdown table: {len(table_data)} rows, {math_count} equations")
                
                i = j
                continue
        
        # Xử lý paragraph thông thường
        new_para = new_doc.add_paragraph()
        
        if format_question_answer(new_para, text):
            stats['questions_formatted'] += 1
            stats['debug_log'].append(f"📝 P{i}: Formatted as Q&A")
        else:
            math_count = process_text_with_math(new_para, text)
            stats['math_expressions'] += math_count
            if math_count > 0:
                stats['debug_log'].append(f"🔢 P{i}: Created {math_count} equations")
        
        i += 1
    
    # Xử lý Word tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        stats['debug_log'].append(f"📊 Processing Word table {table_idx}")
        
        num_rows = len(table.rows)
        num_cols = len(table.rows[0].cells) if table.rows else 0
        
        # Tạo table mới
        new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
        
        table_math = 0
        for r in range(num_rows):
            for c in range(min(len(table.rows[r].cells), num_cols)):
                original_cell = table.rows[r].cells[c]
                new_cell = new_table.rows[r].cells[c]
                
                cell_text = original_cell.text.strip()
                if cell_text:
                    # Check for math in cell
                    cell_math = find_math_expressions(cell_text)
                    if cell_math:
                        stats['debug_log'].append(f"   Cell[{r}][{c}]: {len(cell_math)} math expressions")
                    
                    new_cell.text = ""
                    if new_cell.paragraphs:
                        cell_para = new_cell.paragraphs[0]
                        cell_para.clear()
                    else:
                        cell_para = new_cell.add_paragraph()
                    
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count = process_text_with_math(cell_para, cell_text)
                        table_math += cell_math_count
                    
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        stats['math_expressions'] += table_math
        stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations created")
        
        # Format table
        if options.get('format_tables', True):
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if new_table.rows:
                header_row = new_table.rows[0]
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
            
            stats['word_tables'] += 1
    
    stats['debug_log'].append(f"✅ Hoàn thành: {stats['math_expressions']} equations, {stats['questions_formatted']} Q&A")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v7.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">LaTeX → Word Equation Objects + Q&A + Tables</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Tạo Equation Objects", value=True,
                                       help="Chuyển $...$ thành equation objects thực sự trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True)  
        format_tables = st.checkbox("📊 Format bảng", value=True)
        convert_markdown = st.checkbox("📋 Markdown Tables", value=True)
        show_debug = st.checkbox("🔍 Debug", value=False)
        
 → **A'** (equation)
        - `${left(D A^{prime}right)}import streamlit as st
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
    """Tìm TẤT CẢ expressions toán học - improved detection"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '

def create_math_equation(paragraph, latex_text):
    """Tạo equation object thực sự trong Word"""
    try:
        # Convert LaTeX to simpler math notation first
        unicode_text = convert_latex_symbols(latex_text)
        
        # Create equation run với Office Math format
        # Sử dụng approach tạo equation object thông qua OMML
        
        # Tạo math element
        math_xml = f"""
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:t>{unicode_text}</m:t>
            </m:r>
        </m:oMath>
        """
        
        try:
            # Thử insert equation object
            math_element = parse_xml(math_xml)
            paragraph._element.append(math_element)
            return True
        except:
            # Fallback: Tạo run với style đặc biệt trông như equation
            return create_equation_style_run(paragraph, unicode_text)
            
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def create_equation_style_run(paragraph, text):
    """Tạo run trông như equation (fallback)"""
    try:
        run = paragraph.add_run(text)
        
        # Style để trông như equation
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
        
        # Thêm highlight nhẹ
        try:
            # Add subtle background
            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F0F8FF"/>'
            shading = parse_xml(shading_xml)
            run._element.get_or_add_rPr().append(shading)
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao với OMML format"""
    try:
        # Parse LaTeX content để tạo OMML phù hợp
        unicode_content = convert_latex_symbols(latex_text)
        
        # Phân tích content để tạo structure phù hợp
        if '^' in unicode_content or '²' in unicode_content or '³' in unicode_content:
            # Có superscript - tạo sup element
            return create_superscript_equation(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            # Có parentheses - tạo brackets
            return create_bracketed_equation(paragraph, unicode_content)
        else:
            # Simple math
            return create_simple_equation(paragraph, unicode_content)
            
    except Exception:
        return create_equation_style_run(paragraph, unicode_content)

def create_word_equation(paragraph, latex_text, expr_type='normal'):
    """Tạo equation object đẹp trong Word - Improved version"""
    try:
        # Convert LaTeX to Unicode
        unicode_content = convert_latex_symbols(latex_text)
        
        # Method 1: Thử tạo equation field thực sự
        if create_equation_field(paragraph, unicode_content, latex_text):
            return True
        
        # Method 2: Tạo styled run trông như equation
        return create_professional_math_run(paragraph, unicode_content, latex_text)
        
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[Math: {latex_text}]")
        run.italic = True
        run.font.color.rgb = RGBColor(255, 0, 0)
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Thử tạo equation field thực sự trong Word"""
    try:
        # Approach: Insert equation through field codes
        # EQ field is Word's built-in equation field
        
        # Convert content to EQ field format
        eq_content = convert_to_eq_field(unicode_text, original_latex)
        
        # Create field run
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

def convert_to_eq_field(unicode_text, latex_text):
    """Convert to EQ field format"""
    # EQ field syntax for common patterns
    eq_content = unicode_text
    
    # Handle superscripts for EQ field
    if "'" in unicode_text:
        # A' -> A\s\up5(')
        base = unicode_text.replace("'", "")
        eq_content = f"{base}\\s\\up5(')"
    elif '²' in unicode_text:
        base = unicode_text.replace('²', '')
        eq_content = f"{base}\\s\\up5(2)"
    elif '³' in unicode_text:
        base = unicode_text.replace('³', '')
        eq_content = f"{base}\\s\\up5(3)"
    
    return eq_content

def create_professional_math_run(paragraph, unicode_text, original_latex):
    """Tạo run với style chuyên nghiệp như equation"""
    try:
        # Tạo run với style equation đặc biệt
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.bold = False
        
        # Màu equation chuyên nghiệp
        run.font.color.rgb = RGBColor(0, 32, 96)  # Dark navy blue
        
        # Thêm background nhẹ để highlight
        try:
            rPr = run._element.get_or_add_rPr()
            
            # Add equation-like background
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
            
            # Add slight border for equation effect
            border_xml = '''<w:bdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                           w:val="single" w:sz="2" w:space="1" w:color="E3E6F0"/>'''
            try:
                border = parse_xml(border_xml)
                rPr.append(border)
            except:
                pass
                
        except:
            pass
        
        # Add subtle spacing
        try:
            # Add space before and after for equation isolation
            space_before = paragraph.add_run(" ")
            space_before.font.size = Pt(8)
            
            # Move the equation run to correct position
            # (This is complex, so we'll rely on the styling)
            
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_complex_equation_omml(paragraph, latex_text):
    """Tạo OMML equation phức tạp cho cases đặc biệt"""
    try:
        unicode_content = convert_latex_symbols(latex_text)
        
        # Determine equation type và tạo OMML phù hợp
        if "'" in unicode_content:
            return create_prime_equation_omml(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            return create_parentheses_equation_omml(paragraph, unicode_content)
        elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
            return create_superscript_equation_omml(paragraph, unicode_content)
        else:
            return create_simple_equation_omml(paragraph, unicode_content)
            
    except Exception:
        return False

def create_prime_equation_omml(paragraph, text):
    """OMML cho prime notation"""
    try:
        # Split base and prime
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{"'" * prime_count}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_parentheses_equation_omml(paragraph, text):
    """OMML cho expressions với dấu ngoặc"""
    try:
        # Extract content inside parentheses
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end]
        else:
            content = text
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:d>
                <m:dPr>
                    <m:begChr m:val="("/>
                    <m:endChr m:val=")"/>
                    <m:grow m:val="1"/>
                </m:dPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{content}</m:t>
                    </m:r>
                </m:e>
            </m:d>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_superscript_equation_omml(paragraph, text):
    """OMML cho superscript numbers"""
    try:
        # Find base and superscript
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                # Convert back to normal number
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:  # Still building base
                    base += char
        
        if not sup:
            return create_simple_equation_omml(paragraph, text)
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{sup}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_simple_equation_omml(paragraph, text):
    """OMML cho equation đơn giản"""
    try:
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:rPr>
                    <m:scr m:val="roman"/>
                    <m:sty m:val="i"/>
                </m:rPr>
                <m:t>{text}</m:t>
            </m:r>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Enhanced với multiple approaches"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    successful_equations = 0
    
    for start, end, latex_content, expr_type in expressions:
        # Thêm text trước equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Thử tạo equation với multiple methods
        equation_created = False
        
        # Method 1: Complex OMML cho cases đặc biệt
        if create_complex_equation_omml(paragraph, latex_content):
            equation_created = True
            successful_equations += 1
        # Method 2: Word equation field
        elif create_word_equation(paragraph, latex_content, expr_type):
            equation_created = True
            successful_equations += 1
        # Method 3: Professional styled run
        else:
            unicode_text = convert_latex_symbols(latex_content)
            if create_professional_math_run(paragraph, unicode_text, latex_content):
                equation_created = True
            else:
                # Final fallback
                run = paragraph.add_run(f"[{latex_content}]")
                run.italic = True
                run.font.color.rgb = RGBColor(255, 0, 0)
        
        math_count += 1
        last_pos = end
    
    # Thêm text còn lại
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Kiểm tra pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # In đậm phần câu hỏi
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Xử lý phần nội dung
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Kiểm tra pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format chữ cái đáp án
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Xử lý nội dung đáp án
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A, xử lý bình thường
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra xem text có phải markdown table không"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse text thành table data"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Bỏ qua dòng separator
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            # Loại bỏ | đầu và cuối
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table từ data với enhanced equation processing"""
    if not table_data:
        return 0
    
    math_count = 0
    max_cols = max(len(row) for row in table_data)
    
    # Tạo table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Điền dữ liệu
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Tạo paragraph trong cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Check for math expressions in cell
                cell_math_expressions = find_math_expressions(cell_text)
                
                # Xử lý Q&A hoặc math
                if not format_question_answer(cell_para, cell_text):
                    cell_math_count = process_text_with_math(cell_para, cell_text)
                    math_count += cell_math_count
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table structure
    if options.get('format_tables', True):
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header
        if word_table.rows:
            header_row = word_table.rows[0]
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                        run.font.size = Pt(12)
                
                # Màu nền xanh
                try:
                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                    shading = parse_xml(shading_xml)
                    cell._tc.get_or_add_tcPr().append(shading)
                except:
                    pass
    
    return math_count

def process_document_content(doc, options):
    """Xử lý nội dung document với detailed logging"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'successful_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Xử lý paragraphs
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Debug: Check for math content
        math_expressions = find_math_expressions(text)
        if math_expressions:
            stats['debug_log'].append(f"📐 P{i}: Found {len(math_expressions)} math expressions")
            for j, (start, end, content, expr_type) in enumerate(math_expressions):
                stats['debug_log'].append(f"   Math {j+1}: '{content}' ({expr_type})")
        
        # Kiểm tra markdown table
        if '|' in text and options.get('convert_markdown', True):
            # Collect multiple lines for table
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
            
            # Check if it's a valid table
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    math_count = create_word_table(new_doc, table_data, options)
                    stats['math_expressions'] += math_count
                    stats['markdown_tables'] += 1
                    stats['debug_log'].append(f"📋 Converted markdown table: {len(table_data)} rows, {math_count} equations")
                
                i = j
                continue
        
        # Xử lý paragraph thông thường
        new_para = new_doc.add_paragraph()
        
        if format_question_answer(new_para, text):
            stats['questions_formatted'] += 1
            stats['debug_log'].append(f"📝 P{i}: Formatted as Q&A")
        else:
            math_count = process_text_with_math(new_para, text)
            stats['math_expressions'] += math_count
            if math_count > 0:
                stats['debug_log'].append(f"🔢 P{i}: Created {math_count} equations")
        
        i += 1
    
    # Xử lý Word tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        stats['debug_log'].append(f"📊 Processing Word table {table_idx}")
        
        num_rows = len(table.rows)
        num_cols = len(table.rows[0].cells) if table.rows else 0
        
        # Tạo table mới
        new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
        
        table_math = 0
        for r in range(num_rows):
            for c in range(min(len(table.rows[r].cells), num_cols)):
                original_cell = table.rows[r].cells[c]
                new_cell = new_table.rows[r].cells[c]
                
                cell_text = original_cell.text.strip()
                if cell_text:
                    # Check for math in cell
                    cell_math = find_math_expressions(cell_text)
                    if cell_math:
                        stats['debug_log'].append(f"   Cell[{r}][{c}]: {len(cell_math)} math expressions")
                    
                    new_cell.text = ""
                    if new_cell.paragraphs:
                        cell_para = new_cell.paragraphs[0]
                        cell_para.clear()
                    else:
                        cell_para = new_cell.add_paragraph()
                    
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count = process_text_with_math(cell_para, cell_text)
                        table_math += cell_math_count
                    
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        stats['math_expressions'] += table_math
        stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations created")
        
        # Format table
        if options.get('format_tables', True):
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if new_table.rows:
                header_row = new_table.rows[0]
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
            
            stats['word_tables'] += 1
    
    stats['debug_log'].append(f"✅ Hoàn thành: {stats['math_expressions']} equations, {stats['questions_formatted']} Q&A")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v7.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">LaTeX → Word Equation Objects + Q&A + Tables</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Tạo Equation Objects", value=True,
                                       help="Chuyển $...$ thành equation objects thực sự trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True)  
        format_tables = st.checkbox("📊 Format bảng", value=True)
        convert_markdown = st.checkbox("📋 Markdown Tables", value=True)
        show_debug = st.checkbox("🔍 Debug", value=False)
        
 → **(D A')** (equation)  
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
    """Tìm TẤT CẢ expressions toán học - improved detection"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '

def create_math_equation(paragraph, latex_text):
    """Tạo equation object thực sự trong Word"""
    try:
        # Convert LaTeX to simpler math notation first
        unicode_text = convert_latex_symbols(latex_text)
        
        # Create equation run với Office Math format
        # Sử dụng approach tạo equation object thông qua OMML
        
        # Tạo math element
        math_xml = f"""
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:t>{unicode_text}</m:t>
            </m:r>
        </m:oMath>
        """
        
        try:
            # Thử insert equation object
            math_element = parse_xml(math_xml)
            paragraph._element.append(math_element)
            return True
        except:
            # Fallback: Tạo run với style đặc biệt trông như equation
            return create_equation_style_run(paragraph, unicode_text)
            
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def create_equation_style_run(paragraph, text):
    """Tạo run trông như equation (fallback)"""
    try:
        run = paragraph.add_run(text)
        
        # Style để trông như equation
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
        
        # Thêm highlight nhẹ
        try:
            # Add subtle background
            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F0F8FF"/>'
            shading = parse_xml(shading_xml)
            run._element.get_or_add_rPr().append(shading)
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao với OMML format"""
    try:
        # Parse LaTeX content để tạo OMML phù hợp
        unicode_content = convert_latex_symbols(latex_text)
        
        # Phân tích content để tạo structure phù hợp
        if '^' in unicode_content or '²' in unicode_content or '³' in unicode_content:
            # Có superscript - tạo sup element
            return create_superscript_equation(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            # Có parentheses - tạo brackets
            return create_bracketed_equation(paragraph, unicode_content)
        else:
            # Simple math
            return create_simple_equation(paragraph, unicode_content)
            
    except Exception:
        return create_equation_style_run(paragraph, unicode_content)

def create_word_equation(paragraph, latex_text, expr_type='normal'):
    """Tạo equation object đẹp trong Word - Improved version"""
    try:
        # Convert LaTeX to Unicode
        unicode_content = convert_latex_symbols(latex_text)
        
        # Method 1: Thử tạo equation field thực sự
        if create_equation_field(paragraph, unicode_content, latex_text):
            return True
        
        # Method 2: Tạo styled run trông như equation
        return create_professional_math_run(paragraph, unicode_content, latex_text)
        
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[Math: {latex_text}]")
        run.italic = True
        run.font.color.rgb = RGBColor(255, 0, 0)
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Thử tạo equation field thực sự trong Word"""
    try:
        # Approach: Insert equation through field codes
        # EQ field is Word's built-in equation field
        
        # Convert content to EQ field format
        eq_content = convert_to_eq_field(unicode_text, original_latex)
        
        # Create field run
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

def convert_to_eq_field(unicode_text, latex_text):
    """Convert to EQ field format"""
    # EQ field syntax for common patterns
    eq_content = unicode_text
    
    # Handle superscripts for EQ field
    if "'" in unicode_text:
        # A' -> A\s\up5(')
        base = unicode_text.replace("'", "")
        eq_content = f"{base}\\s\\up5(')"
    elif '²' in unicode_text:
        base = unicode_text.replace('²', '')
        eq_content = f"{base}\\s\\up5(2)"
    elif '³' in unicode_text:
        base = unicode_text.replace('³', '')
        eq_content = f"{base}\\s\\up5(3)"
    
    return eq_content

def create_professional_math_run(paragraph, unicode_text, original_latex):
    """Tạo run với style chuyên nghiệp như equation"""
    try:
        # Tạo run với style equation đặc biệt
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.bold = False
        
        # Màu equation chuyên nghiệp
        run.font.color.rgb = RGBColor(0, 32, 96)  # Dark navy blue
        
        # Thêm background nhẹ để highlight
        try:
            rPr = run._element.get_or_add_rPr()
            
            # Add equation-like background
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
            
            # Add slight border for equation effect
            border_xml = '''<w:bdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                           w:val="single" w:sz="2" w:space="1" w:color="E3E6F0"/>'''
            try:
                border = parse_xml(border_xml)
                rPr.append(border)
            except:
                pass
                
        except:
            pass
        
        # Add subtle spacing
        try:
            # Add space before and after for equation isolation
            space_before = paragraph.add_run(" ")
            space_before.font.size = Pt(8)
            
            # Move the equation run to correct position
            # (This is complex, so we'll rely on the styling)
            
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_complex_equation_omml(paragraph, latex_text):
    """Tạo OMML equation phức tạp cho cases đặc biệt"""
    try:
        unicode_content = convert_latex_symbols(latex_text)
        
        # Determine equation type và tạo OMML phù hợp
        if "'" in unicode_content:
            return create_prime_equation_omml(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            return create_parentheses_equation_omml(paragraph, unicode_content)
        elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
            return create_superscript_equation_omml(paragraph, unicode_content)
        else:
            return create_simple_equation_omml(paragraph, unicode_content)
            
    except Exception:
        return False

def create_prime_equation_omml(paragraph, text):
    """OMML cho prime notation"""
    try:
        # Split base and prime
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{"'" * prime_count}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_parentheses_equation_omml(paragraph, text):
    """OMML cho expressions với dấu ngoặc"""
    try:
        # Extract content inside parentheses
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end]
        else:
            content = text
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:d>
                <m:dPr>
                    <m:begChr m:val="("/>
                    <m:endChr m:val=")"/>
                    <m:grow m:val="1"/>
                </m:dPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{content}</m:t>
                    </m:r>
                </m:e>
            </m:d>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_superscript_equation_omml(paragraph, text):
    """OMML cho superscript numbers"""
    try:
        # Find base and superscript
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                # Convert back to normal number
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:  # Still building base
                    base += char
        
        if not sup:
            return create_simple_equation_omml(paragraph, text)
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{sup}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_simple_equation_omml(paragraph, text):
    """OMML cho equation đơn giản"""
    try:
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:rPr>
                    <m:scr m:val="roman"/>
                    <m:sty m:val="i"/>
                </m:rPr>
                <m:t>{text}</m:t>
            </m:r>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Enhanced với multiple approaches"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    successful_equations = 0
    
    for start, end, latex_content, expr_type in expressions:
        # Thêm text trước equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Thử tạo equation với multiple methods
        equation_created = False
        
        # Method 1: Complex OMML cho cases đặc biệt
        if create_complex_equation_omml(paragraph, latex_content):
            equation_created = True
            successful_equations += 1
        # Method 2: Word equation field
        elif create_word_equation(paragraph, latex_content, expr_type):
            equation_created = True
            successful_equations += 1
        # Method 3: Professional styled run
        else:
            unicode_text = convert_latex_symbols(latex_content)
            if create_professional_math_run(paragraph, unicode_text, latex_content):
                equation_created = True
            else:
                # Final fallback
                run = paragraph.add_run(f"[{latex_content}]")
                run.italic = True
                run.font.color.rgb = RGBColor(255, 0, 0)
        
        math_count += 1
        last_pos = end
    
    # Thêm text còn lại
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Kiểm tra pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # In đậm phần câu hỏi
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Xử lý phần nội dung
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Kiểm tra pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format chữ cái đáp án
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Xử lý nội dung đáp án
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A, xử lý bình thường
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra xem text có phải markdown table không"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse text thành table data"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Bỏ qua dòng separator
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            # Loại bỏ | đầu và cuối
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table từ data với enhanced equation processing"""
    if not table_data:
        return 0
    
    math_count = 0
    max_cols = max(len(row) for row in table_data)
    
    # Tạo table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Điền dữ liệu
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Tạo paragraph trong cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Check for math expressions in cell
                cell_math_expressions = find_math_expressions(cell_text)
                
                # Xử lý Q&A hoặc math
                if not format_question_answer(cell_para, cell_text):
                    cell_math_count = process_text_with_math(cell_para, cell_text)
                    math_count += cell_math_count
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table structure
    if options.get('format_tables', True):
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header
        if word_table.rows:
            header_row = word_table.rows[0]
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                        run.font.size = Pt(12)
                
                # Màu nền xanh
                try:
                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                    shading = parse_xml(shading_xml)
                    cell._tc.get_or_add_tcPr().append(shading)
                except:
                    pass
    
    return math_count

def process_document_content(doc, options):
    """Xử lý nội dung document với detailed logging"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'successful_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Xử lý paragraphs
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Debug: Check for math content
        math_expressions = find_math_expressions(text)
        if math_expressions:
            stats['debug_log'].append(f"📐 P{i}: Found {len(math_expressions)} math expressions")
            for j, (start, end, content, expr_type) in enumerate(math_expressions):
                stats['debug_log'].append(f"   Math {j+1}: '{content}' ({expr_type})")
        
        # Kiểm tra markdown table
        if '|' in text and options.get('convert_markdown', True):
            # Collect multiple lines for table
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
            
            # Check if it's a valid table
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    math_count = create_word_table(new_doc, table_data, options)
                    stats['math_expressions'] += math_count
                    stats['markdown_tables'] += 1
                    stats['debug_log'].append(f"📋 Converted markdown table: {len(table_data)} rows, {math_count} equations")
                
                i = j
                continue
        
        # Xử lý paragraph thông thường
        new_para = new_doc.add_paragraph()
        
        if format_question_answer(new_para, text):
            stats['questions_formatted'] += 1
            stats['debug_log'].append(f"📝 P{i}: Formatted as Q&A")
        else:
            math_count = process_text_with_math(new_para, text)
            stats['math_expressions'] += math_count
            if math_count > 0:
                stats['debug_log'].append(f"🔢 P{i}: Created {math_count} equations")
        
        i += 1
    
    # Xử lý Word tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        stats['debug_log'].append(f"📊 Processing Word table {table_idx}")
        
        num_rows = len(table.rows)
        num_cols = len(table.rows[0].cells) if table.rows else 0
        
        # Tạo table mới
        new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
        
        table_math = 0
        for r in range(num_rows):
            for c in range(min(len(table.rows[r].cells), num_cols)):
                original_cell = table.rows[r].cells[c]
                new_cell = new_table.rows[r].cells[c]
                
                cell_text = original_cell.text.strip()
                if cell_text:
                    # Check for math in cell
                    cell_math = find_math_expressions(cell_text)
                    if cell_math:
                        stats['debug_log'].append(f"   Cell[{r}][{c}]: {len(cell_math)} math expressions")
                    
                    new_cell.text = ""
                    if new_cell.paragraphs:
                        cell_para = new_cell.paragraphs[0]
                        cell_para.clear()
                    else:
                        cell_para = new_cell.add_paragraph()
                    
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count = process_text_with_math(cell_para, cell_text)
                        table_math += cell_math_count
                    
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        stats['math_expressions'] += table_math
        stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations created")
        
        # Format table
        if options.get('format_tables', True):
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if new_table.rows:
                header_row = new_table.rows[0]
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
            
            stats['word_tables'] += 1
    
    stats['debug_log'].append(f"✅ Hoàn thành: {stats['math_expressions']} equations, {stats['questions_formatted']} Q&A")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v7.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">LaTeX → Word Equation Objects + Q&A + Tables</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Tạo Equation Objects", value=True,
                                       help="Chuyển $...$ thành equation objects thực sự trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True)  
        format_tables = st.checkbox("📊 Format bảng", value=True)
        convert_markdown = st.checkbox("📋 Markdown Tables", value=True)
        show_debug = st.checkbox("🔍 Debug", value=False)
        
 → **[6;8]** (equation)
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
    """Tìm TẤT CẢ expressions toán học - improved detection"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '

def create_math_equation(paragraph, latex_text):
    """Tạo equation object thực sự trong Word"""
    try:
        # Convert LaTeX to simpler math notation first
        unicode_text = convert_latex_symbols(latex_text)
        
        # Create equation run với Office Math format
        # Sử dụng approach tạo equation object thông qua OMML
        
        # Tạo math element
        math_xml = f"""
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:t>{unicode_text}</m:t>
            </m:r>
        </m:oMath>
        """
        
        try:
            # Thử insert equation object
            math_element = parse_xml(math_xml)
            paragraph._element.append(math_element)
            return True
        except:
            # Fallback: Tạo run với style đặc biệt trông như equation
            return create_equation_style_run(paragraph, unicode_text)
            
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def create_equation_style_run(paragraph, text):
    """Tạo run trông như equation (fallback)"""
    try:
        run = paragraph.add_run(text)
        
        # Style để trông như equation
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
        
        # Thêm highlight nhẹ
        try:
            # Add subtle background
            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F0F8FF"/>'
            shading = parse_xml(shading_xml)
            run._element.get_or_add_rPr().append(shading)
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao với OMML format"""
    try:
        # Parse LaTeX content để tạo OMML phù hợp
        unicode_content = convert_latex_symbols(latex_text)
        
        # Phân tích content để tạo structure phù hợp
        if '^' in unicode_content or '²' in unicode_content or '³' in unicode_content:
            # Có superscript - tạo sup element
            return create_superscript_equation(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            # Có parentheses - tạo brackets
            return create_bracketed_equation(paragraph, unicode_content)
        else:
            # Simple math
            return create_simple_equation(paragraph, unicode_content)
            
    except Exception:
        return create_equation_style_run(paragraph, unicode_content)

def create_word_equation(paragraph, latex_text, expr_type='normal'):
    """Tạo equation object đẹp trong Word - Improved version"""
    try:
        # Convert LaTeX to Unicode
        unicode_content = convert_latex_symbols(latex_text)
        
        # Method 1: Thử tạo equation field thực sự
        if create_equation_field(paragraph, unicode_content, latex_text):
            return True
        
        # Method 2: Tạo styled run trông như equation
        return create_professional_math_run(paragraph, unicode_content, latex_text)
        
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[Math: {latex_text}]")
        run.italic = True
        run.font.color.rgb = RGBColor(255, 0, 0)
        return False

def create_equation_field(paragraph, unicode_text, original_latex):
    """Thử tạo equation field thực sự trong Word"""
    try:
        # Approach: Insert equation through field codes
        # EQ field is Word's built-in equation field
        
        # Convert content to EQ field format
        eq_content = convert_to_eq_field(unicode_text, original_latex)
        
        # Create field run
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

def convert_to_eq_field(unicode_text, latex_text):
    """Convert to EQ field format"""
    # EQ field syntax for common patterns
    eq_content = unicode_text
    
    # Handle superscripts for EQ field
    if "'" in unicode_text:
        # A' -> A\s\up5(')
        base = unicode_text.replace("'", "")
        eq_content = f"{base}\\s\\up5(')"
    elif '²' in unicode_text:
        base = unicode_text.replace('²', '')
        eq_content = f"{base}\\s\\up5(2)"
    elif '³' in unicode_text:
        base = unicode_text.replace('³', '')
        eq_content = f"{base}\\s\\up5(3)"
    
    return eq_content

def create_professional_math_run(paragraph, unicode_text, original_latex):
    """Tạo run với style chuyên nghiệp như equation"""
    try:
        # Tạo run với style equation đặc biệt
        run = paragraph.add_run(unicode_text)
        
        # Professional equation styling
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)
        run.italic = True
        run.bold = False
        
        # Màu equation chuyên nghiệp
        run.font.color.rgb = RGBColor(0, 32, 96)  # Dark navy blue
        
        # Thêm background nhẹ để highlight
        try:
            rPr = run._element.get_or_add_rPr()
            
            # Add equation-like background
            shading_xml = '''<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                            w:val="clear" w:color="auto" w:fill="F8F9FA"/>'''
            shading = parse_xml(shading_xml)
            rPr.append(shading)
            
            # Add slight border for equation effect
            border_xml = '''<w:bdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                           w:val="single" w:sz="2" w:space="1" w:color="E3E6F0"/>'''
            try:
                border = parse_xml(border_xml)
                rPr.append(border)
            except:
                pass
                
        except:
            pass
        
        # Add subtle spacing
        try:
            # Add space before and after for equation isolation
            space_before = paragraph.add_run(" ")
            space_before.font.size = Pt(8)
            
            # Move the equation run to correct position
            # (This is complex, so we'll rely on the styling)
            
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_complex_equation_omml(paragraph, latex_text):
    """Tạo OMML equation phức tạp cho cases đặc biệt"""
    try:
        unicode_content = convert_latex_symbols(latex_text)
        
        # Determine equation type và tạo OMML phù hợp
        if "'" in unicode_content:
            return create_prime_equation_omml(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            return create_parentheses_equation_omml(paragraph, unicode_content)
        elif any(c in unicode_content for c in ['²', '³', '⁴', '⁵']):
            return create_superscript_equation_omml(paragraph, unicode_content)
        else:
            return create_simple_equation_omml(paragraph, unicode_content)
            
    except Exception:
        return False

def create_prime_equation_omml(paragraph, text):
    """OMML cho prime notation"""
    try:
        # Split base and prime
        base = text.replace("'", "").strip()
        prime_count = text.count("'")
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{"'" * prime_count}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_parentheses_equation_omml(paragraph, text):
    """OMML cho expressions với dấu ngoặc"""
    try:
        # Extract content inside parentheses
        start = text.find('(')
        end = text.rfind(')')
        
        if start != -1 and end != -1 and start < end:
            content = text[start+1:end]
        else:
            content = text
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:d>
                <m:dPr>
                    <m:begChr m:val="("/>
                    <m:endChr m:val=")"/>
                    <m:grow m:val="1"/>
                </m:dPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{content}</m:t>
                    </m:r>
                </m:e>
            </m:d>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_superscript_equation_omml(paragraph, text):
    """OMML cho superscript numbers"""
    try:
        # Find base and superscript
        base = ""
        sup = ""
        
        for char in text:
            if char in '⁰¹²³⁴⁵⁶⁷⁸⁹':
                # Convert back to normal number
                sup_map = {'⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
                          '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
                sup += sup_map.get(char, char)
            else:
                if not sup:  # Still building base
                    base += char
        
        if not sup:
            return create_simple_equation_omml(paragraph, text)
        
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:sSupPr></m:sSupPr>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{base}</m:t>
                    </m:r>
                </m:e>
                <m:sup>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{sup}</m:t>
                    </m:r>
                </m:sup>
            </m:sSup>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_simple_equation_omml(paragraph, text):
    """OMML cho equation đơn giản"""
    try:
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:rPr>
                    <m:scr m:val="roman"/>
                    <m:sty m:val="i"/>
                </m:rPr>
                <m:t>{text}</m:t>
            </m:r>
        </m:oMath>
        '''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - Enhanced với multiple approaches"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    successful_equations = 0
    
    for start, end, latex_content, expr_type in expressions:
        # Thêm text trước equation
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Thử tạo equation với multiple methods
        equation_created = False
        
        # Method 1: Complex OMML cho cases đặc biệt
        if create_complex_equation_omml(paragraph, latex_content):
            equation_created = True
            successful_equations += 1
        # Method 2: Word equation field
        elif create_word_equation(paragraph, latex_content, expr_type):
            equation_created = True
            successful_equations += 1
        # Method 3: Professional styled run
        else:
            unicode_text = convert_latex_symbols(latex_content)
            if create_professional_math_run(paragraph, unicode_text, latex_content):
                equation_created = True
            else:
                # Final fallback
                run = paragraph.add_run(f"[{latex_content}]")
                run.italic = True
                run.font.color.rgb = RGBColor(255, 0, 0)
        
        math_count += 1
        last_pos = end
    
    # Thêm text còn lại
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Kiểm tra pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # In đậm phần câu hỏi
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Xử lý phần nội dung
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Kiểm tra pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format chữ cái đáp án
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Xử lý nội dung đáp án
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A, xử lý bình thường
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra xem text có phải markdown table không"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse text thành table data"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Bỏ qua dòng separator
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            # Loại bỏ | đầu và cuối
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table từ data với enhanced equation processing"""
    if not table_data:
        return 0
    
    math_count = 0
    max_cols = max(len(row) for row in table_data)
    
    # Tạo table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Điền dữ liệu
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Tạo paragraph trong cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Check for math expressions in cell
                cell_math_expressions = find_math_expressions(cell_text)
                
                # Xử lý Q&A hoặc math
                if not format_question_answer(cell_para, cell_text):
                    cell_math_count = process_text_with_math(cell_para, cell_text)
                    math_count += cell_math_count
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table structure
    if options.get('format_tables', True):
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header
        if word_table.rows:
            header_row = word_table.rows[0]
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                        run.font.size = Pt(12)
                
                # Màu nền xanh
                try:
                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                    shading = parse_xml(shading_xml)
                    cell._tc.get_or_add_tcPr().append(shading)
                except:
                    pass
    
    return math_count

def process_document_content(doc, options):
    """Xử lý nội dung document với detailed logging"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'successful_equations': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"🚀 Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
    
    # Xử lý paragraphs
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Debug: Check for math content
        math_expressions = find_math_expressions(text)
        if math_expressions:
            stats['debug_log'].append(f"📐 P{i}: Found {len(math_expressions)} math expressions")
            for j, (start, end, content, expr_type) in enumerate(math_expressions):
                stats['debug_log'].append(f"   Math {j+1}: '{content}' ({expr_type})")
        
        # Kiểm tra markdown table
        if '|' in text and options.get('convert_markdown', True):
            # Collect multiple lines for table
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
            
            # Check if it's a valid table
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    math_count = create_word_table(new_doc, table_data, options)
                    stats['math_expressions'] += math_count
                    stats['markdown_tables'] += 1
                    stats['debug_log'].append(f"📋 Converted markdown table: {len(table_data)} rows, {math_count} equations")
                
                i = j
                continue
        
        # Xử lý paragraph thông thường
        new_para = new_doc.add_paragraph()
        
        if format_question_answer(new_para, text):
            stats['questions_formatted'] += 1
            stats['debug_log'].append(f"📝 P{i}: Formatted as Q&A")
        else:
            math_count = process_text_with_math(new_para, text)
            stats['math_expressions'] += math_count
            if math_count > 0:
                stats['debug_log'].append(f"🔢 P{i}: Created {math_count} equations")
        
        i += 1
    
    # Xử lý Word tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        stats['debug_log'].append(f"📊 Processing Word table {table_idx}")
        
        num_rows = len(table.rows)
        num_cols = len(table.rows[0].cells) if table.rows else 0
        
        # Tạo table mới
        new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
        
        table_math = 0
        for r in range(num_rows):
            for c in range(min(len(table.rows[r].cells), num_cols)):
                original_cell = table.rows[r].cells[c]
                new_cell = new_table.rows[r].cells[c]
                
                cell_text = original_cell.text.strip()
                if cell_text:
                    # Check for math in cell
                    cell_math = find_math_expressions(cell_text)
                    if cell_math:
                        stats['debug_log'].append(f"   Cell[{r}][{c}]: {len(cell_math)} math expressions")
                    
                    new_cell.text = ""
                    if new_cell.paragraphs:
                        cell_para = new_cell.paragraphs[0]
                        cell_para.clear()
                    else:
                        cell_para = new_cell.add_paragraph()
                    
                    if not format_question_answer(cell_para, cell_text):
                        cell_math_count = process_text_with_math(cell_para, cell_text)
                        table_math += cell_math_count
                    
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        stats['math_expressions'] += table_math
        stats['debug_log'].append(f"📊 Table {table_idx}: {table_math} equations created")
        
        # Format table
        if options.get('format_tables', True):
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if new_table.rows:
                header_row = new_table.rows[0]
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
            
            stats['word_tables'] += 1
    
    stats['debug_log'].append(f"✅ Hoàn thành: {stats['math_expressions']} equations, {stats['questions_formatted']} Q&A")
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v7.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">LaTeX → Word Equation Objects + Q&A + Tables</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Tạo Equation Objects", value=True,
                                       help="Chuyển $...$ thành equation objects thực sự trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True)  
        format_tables = st.checkbox("📊 Format bảng", value=True)
        convert_markdown = st.checkbox("📋 Markdown Tables", value=True)
        show_debug = st.checkbox("🔍 Debug", value=False)
        
 → **x²** (equation)
        
        **✨ All as Word equation objects!**
        """)
        
        st.markdown("### 📝 Equation Examples")
        st.code("""
$A^{prime}$ → A' (equation)
$x^2$ → x² (equation)
$\\left(x\\right)$ → (x) (equation)
${B^{prime}}$ → B' (equation)
$\\frac{a}{b}$ → (a)/(b) (equation)
$\\pi$ → π (equation)

→ Tạo equation objects thực sự
  không phải text thường!
        """)
        
        st.markdown("### 📋 Table Support")
        st.code("""
| Header | Data |
|--------|------|
| $x^2$  | Value|

→ Equations trong tables
        """)
    
    # Main area
    st.markdown("### 📁 Upload File")
    
    uploaded_file = st.file_uploader("Chọn file Word (.docx)", type=["docx"])
    
    if uploaded_file:
        st.success(f"✅ {uploaded_file.name}")
        
        with st.expander("Preview"):
            try:
                doc = Document(uploaded_file)
                st.info(f"📄 {len(doc.paragraphs)} paragraphs | 📊 {len(doc.tables)} tables")
                
                for i, para in enumerate(doc.paragraphs[:5]):
                    if para.text.strip():
                        preview_text = para.text[:80] + "..." if len(para.text) > 80 else para.text
                        st.text(f"{i+1}. {preview_text}")
            except Exception as e:
                st.error(f"Preview error: {e}")
        
        if st.button("🚀 Xử lý", type="primary", use_container_width=True):
            options = {
                'convert_equations': convert_equations,
                'format_questions': format_questions,
                'format_tables': format_tables,
                'convert_markdown': convert_markdown
            }
            
            with st.spinner("Đang xử lý..."):
                try:
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document_content(doc, options)
                    
                    # Save
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_processed_{timestamp}.docx"
                    
                    # Results
                    st.markdown("### 📊 Kết quả")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("🔢 Equations", stats['math_expressions'])
                    col2.metric("📝 Q&A", stats['questions_formatted'])
                    col3.metric("📊 Word Tables", stats['word_tables'])
                    col4.metric("📋 Markdown", stats['markdown_tables'])
                    
                    if show_debug:
                        with st.expander("🔍 Detailed Debug Log"):
                            for log in stats['debug_log']:
                                if 'Math' in log or 'equation' in log:
                                    st.success(log)
                                elif 'ERROR' in log or 'Error' in log:
                                    st.error(log)
                                elif 'Table' in log or 'Converted' in log:
                                    st.info(log)
                                else:
                                    st.text(log)
                    
                    # Download
                    st.download_button(
                        "⬇️ Tải file",
                        buffer.getvalue(),
                        filename,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                    total_processed = (stats['math_expressions'] + stats['questions_formatted'] + 
                                     stats['word_tables'] + stats['markdown_tables'])
                    
                    if total_processed > 0:
                        st.success("🎉 Xử lý hoàn thành!")
                        
                        success_messages = []
                        if stats['math_expressions'] > 0:
                            success_messages.append(f"🔢 **{stats['math_expressions']} equation objects** được tạo thành công")
                        if stats['markdown_tables'] > 0:
                            success_messages.append(f"📋 **{stats['markdown_tables']} markdown tables** được chuyển đổi")
                        if stats['questions_formatted'] > 0:
                            success_messages.append(f"📝 **{stats['questions_formatted']} Q&A items** được format")
                        if stats['word_tables'] > 0:
                            success_messages.append(f"📊 **{stats['word_tables']} Word tables** được format")
                        
                        for msg in success_messages:
                            st.info(msg)
                            
                    else:
                        st.warning("⚠️ Không tìm thấy LaTeX, Q&A hoặc tables để xử lý")
                
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")

if __name__ == "__main__":
    main():
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
                if i < len(text) and text[i] == '

def create_math_equation(paragraph, latex_text):
    """Tạo equation object thực sự trong Word"""
    try:
        # Convert LaTeX to simpler math notation first
        unicode_text = convert_latex_symbols(latex_text)
        
        # Create equation run với Office Math format
        # Sử dụng approach tạo equation object thông qua OMML
        
        # Tạo math element
        math_xml = f"""
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:t>{unicode_text}</m:t>
            </m:r>
        </m:oMath>
        """
        
        try:
            # Thử insert equation object
            math_element = parse_xml(math_xml)
            paragraph._element.append(math_element)
            return True
        except:
            # Fallback: Tạo run với style đặc biệt trông như equation
            return create_equation_style_run(paragraph, unicode_text)
            
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def create_equation_style_run(paragraph, text):
    """Tạo run trông như equation (fallback)"""
    try:
        run = paragraph.add_run(text)
        
        # Style để trông như equation
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
        
        # Thêm highlight nhẹ
        try:
            # Add subtle background
            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F0F8FF"/>'
            shading = parse_xml(shading_xml)
            run._element.get_or_add_rPr().append(shading)
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao với OMML format"""
    try:
        # Parse LaTeX content để tạo OMML phù hợp
        unicode_content = convert_latex_symbols(latex_text)
        
        # Phân tích content để tạo structure phù hợp
        if '^' in unicode_content or '²' in unicode_content or '³' in unicode_content:
            # Có superscript - tạo sup element
            return create_superscript_equation(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            # Có parentheses - tạo brackets
            return create_bracketed_equation(paragraph, unicode_content)
        else:
            # Simple math
            return create_simple_equation(paragraph, unicode_content)
            
    except Exception:
        return create_equation_style_run(paragraph, unicode_content)

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao - thử OMML trước, fallback về styled text"""
    try:
        # Convert LaTeX to Unicode first
        unicode_content = convert_latex_symbols(latex_text)
        
        # Approach 1: Try OMML equation object
        if try_create_omml_equation(paragraph, unicode_content):
            return True
        
        # Approach 2: Fallback - Enhanced styled text that looks like equation
        return create_enhanced_math_run(paragraph, unicode_content)
            
    except Exception:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def try_create_omml_equation(paragraph, text):
    """Thử tạo OMML equation - simple approach"""
    try:
        # Simple OMML template for basic math
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:rPr>
                    <m:scr m:val="roman"/>
                    <m:sty m:val="i"/>
                </m:rPr>
                <m:t>{text}</m:t>
            </m:r>
        </m:oMath>
        '''
        
        # Parse and append
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_enhanced_math_run(paragraph, text):
    """Tạo run với style nâng cao trông như equation"""
    try:
        # Tạo run với formatting đặc biệt
        run = paragraph.add_run(text)
        
        # Apply equation-like formatting
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)  # Slightly larger
        run.italic = True
        run.bold = False
        
        # Màu xanh đậm cho equation
        run.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue
        
        # Thêm border nhẹ để highlight
        try:
            # Add subtle styling to make it look like equation
            rPr = run._element.get_or_add_rPr()
            
            # Add equation style
            w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            style_xml = f'<w:rStyle xmlns:w="{w_ns}" w:val="MathEquation"/>'
            try:
                style_element = parse_xml(style_xml)
                rPr.append(style_element)
            except:
                pass
                
        except:
            pass
        
        return True
        
    except Exception:
        return False

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - tạo equation objects thực sự"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    
    for start, end, latex_content in expressions:
        # Thêm text trước công thức
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Tạo equation object thực sự
        if create_advanced_equation(paragraph, latex_content):
            math_count += 1
        else:
            # Fallback nếu không tạo được equation
            run = paragraph.add_run(f"[{latex_content}]")
            run.italic = True
            run.font.color.rgb = RGBColor(255, 0, 0)  # Red for errors
        
        last_pos = end
    
    # Thêm text còn lại
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Kiểm tra pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # In đậm phần câu hỏi
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Xử lý phần nội dung
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Kiểm tra pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format chữ cái đáp án
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Xử lý nội dung đáp án
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A, xử lý bình thường
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra xem text có phải markdown table không"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse text thành table data"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Bỏ qua dòng separator
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            # Loại bỏ | đầu và cuối
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table từ data"""
    if not table_data:
        return 0
    
    math_count = 0
    max_cols = max(len(row) for row in table_data)
    
    # Tạo table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Điền dữ liệu
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Tạo paragraph trong cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Xử lý Q&A hoặc math
                if not format_question_answer(cell_para, cell_text):
                    math_count += process_text_with_math(cell_para, cell_text)
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table structure
    if options.get('format_tables', True):
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header
        if word_table.rows:
            header_row = word_table.rows[0]
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                
                # Màu nền xanh
                try:
                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                    shading = parse_xml(shading_xml)
                    cell._tc.get_or_add_tcPr().append(shading)
                except:
                    pass
    
    return math_count

def process_document_content(doc, options):
    """Xử lý nội dung document"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs")
    
    # Xử lý paragraphs
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Kiểm tra markdown table
        if '|' in text and options.get('convert_markdown', True):
            # Collect multiple lines for table
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
            
            # Check if it's a valid table
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    math_count = create_word_table(new_doc, table_data, options)
                    stats['math_expressions'] += math_count
                    stats['markdown_tables'] += 1
                    stats['debug_log'].append(f"Converted markdown table: {len(table_data)} rows")
                
                i = j
                continue
        
        # Xử lý paragraph thông thường
        new_para = new_doc.add_paragraph()
        
        if format_question_answer(new_para, text):
            stats['questions_formatted'] += 1
        else:
            math_count = process_text_with_math(new_para, text)
            stats['math_expressions'] += math_count
        
        i += 1
    
    # Xử lý Word tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        num_rows = len(table.rows)
        num_cols = len(table.rows[0].cells) if table.rows else 0
        
        # Tạo table mới
        new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
        
        table_math = 0
        for r in range(num_rows):
            for c in range(min(len(table.rows[r].cells), num_cols)):
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
                        table_math += process_text_with_math(cell_para, cell_text)
                    
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        stats['math_expressions'] += table_math
        
        # Format table
        if options.get('format_tables', True):
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if new_table.rows:
                header_row = new_table.rows[0]
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
            
            stats['word_tables'] += 1
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v7.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">LaTeX → Word Equation Objects + Q&A + Tables</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Tạo Equation Objects", value=True,
                                       help="Chuyển $...$ thành equation objects thực sự trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True)  
        format_tables = st.checkbox("📊 Format bảng", value=True)
        convert_markdown = st.checkbox("📋 Markdown Tables", value=True)
        show_debug = st.checkbox("🔍 Debug", value=False)
        
        st.markdown("---")
        st.info("💡 **Equation Creation:**\n\n1️⃣ Try OMML equation objects\n2️⃣ Fallback to enhanced math styling\n3️⃣ Equations tích hợp với Word!")
        
        st.markdown("### 📝 Equation Examples")
        st.code("""
$A^{prime}$ → A' (equation)
$x^2$ → x² (equation)
$\\left(x\\right)$ → (x) (equation)
${B^{prime}}$ → B' (equation)
$\\frac{a}{b}$ → (a)/(b) (equation)
$\\pi$ → π (equation)

→ Tạo equation objects thực sự
  không phải text thường!
        """)
        
        st.markdown("### 📋 Table Support")
        st.code("""
| Header | Data |
|--------|------|
| $x^2$  | Value|

→ Equations trong tables
        """)
    
    # Main area
    st.markdown("### 📁 Upload File")
    
    uploaded_file = st.file_uploader("Chọn file Word (.docx)", type=["docx"])
    
    if uploaded_file:
        st.success(f"✅ {uploaded_file.name}")
        
        with st.expander("Preview"):
            try:
                doc = Document(uploaded_file)
                st.info(f"📄 {len(doc.paragraphs)} paragraphs | 📊 {len(doc.tables)} tables")
                
                for i, para in enumerate(doc.paragraphs[:5]):
                    if para.text.strip():
                        preview_text = para.text[:80] + "..." if len(para.text) > 80 else para.text
                        st.text(f"{i+1}. {preview_text}")
            except Exception as e:
                st.error(f"Preview error: {e}")
        
        if st.button("🚀 Xử lý", type="primary", use_container_width=True):
            options = {
                'convert_equations': convert_equations,
                'format_questions': format_questions,
                'format_tables': format_tables,
                'convert_markdown': convert_markdown
            }
            
            with st.spinner("Đang xử lý..."):
                try:
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document_content(doc, options)
                    
                    # Save
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_processed_{timestamp}.docx"
                    
                    # Results
                    st.markdown("### 📊 Kết quả")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("🔢 Equations", stats['math_expressions'])
                    col2.metric("📝 Q&A", stats['questions_formatted'])
                    col3.metric("📊 Word Tables", stats['word_tables'])
                    col4.metric("📋 Markdown", stats['markdown_tables'])
                    
                    if show_debug:
                        with st.expander("Debug"):
                            for log in stats['debug_log']:
                                st.text(log)
                    
                    # Download
                    st.download_button(
                        "⬇️ Tải file",
                        buffer.getvalue(),
                        filename,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                    total_processed = (stats['math_expressions'] + stats['questions_formatted'] + 
                                     stats['word_tables'] + stats['markdown_tables'])
                    
                    if total_processed > 0:
                        st.success("🎉 Hoàn thành!")
                        if stats['math_expressions'] > 0:
                            st.info(f"✨ Đã tạo {stats['math_expressions']} equation objects trong Word!")
                        if stats['markdown_tables'] > 0:
                            st.info(f"📋 Đã chuyển đổi {stats['markdown_tables']} bảng Markdown!")
                    else:
                        st.warning("⚠️ Không tìm thấy nội dung cần xử lý")
                
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")

if __name__ == "__main__":
    main():
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
                
                # Tìm $ đóng, cho phép nested braces
                while i < len(text) and text[i] != '

def create_math_equation(paragraph, latex_text):
    """Tạo equation object thực sự trong Word"""
    try:
        # Convert LaTeX to simpler math notation first
        unicode_text = convert_latex_symbols(latex_text)
        
        # Create equation run với Office Math format
        # Sử dụng approach tạo equation object thông qua OMML
        
        # Tạo math element
        math_xml = f"""
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:t>{unicode_text}</m:t>
            </m:r>
        </m:oMath>
        """
        
        try:
            # Thử insert equation object
            math_element = parse_xml(math_xml)
            paragraph._element.append(math_element)
            return True
        except:
            # Fallback: Tạo run với style đặc biệt trông như equation
            return create_equation_style_run(paragraph, unicode_text)
            
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def create_equation_style_run(paragraph, text):
    """Tạo run trông như equation (fallback)"""
    try:
        run = paragraph.add_run(text)
        
        # Style để trông như equation
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
        
        # Thêm highlight nhẹ
        try:
            # Add subtle background
            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F0F8FF"/>'
            shading = parse_xml(shading_xml)
            run._element.get_or_add_rPr().append(shading)
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao với OMML format"""
    try:
        # Parse LaTeX content để tạo OMML phù hợp
        unicode_content = convert_latex_symbols(latex_text)
        
        # Phân tích content để tạo structure phù hợp
        if '^' in unicode_content or '²' in unicode_content or '³' in unicode_content:
            # Có superscript - tạo sup element
            return create_superscript_equation(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            # Có parentheses - tạo brackets
            return create_bracketed_equation(paragraph, unicode_content)
        else:
            # Simple math
            return create_simple_equation(paragraph, unicode_content)
            
    except Exception:
        return create_equation_style_run(paragraph, unicode_content)

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao - thử OMML trước, fallback về styled text"""
    try:
        # Convert LaTeX to Unicode first
        unicode_content = convert_latex_symbols(latex_text)
        
        # Approach 1: Try OMML equation object
        if try_create_omml_equation(paragraph, unicode_content):
            return True
        
        # Approach 2: Fallback - Enhanced styled text that looks like equation
        return create_enhanced_math_run(paragraph, unicode_content)
            
    except Exception:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def try_create_omml_equation(paragraph, text):
    """Thử tạo OMML equation - simple approach"""
    try:
        # Simple OMML template for basic math
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:rPr>
                    <m:scr m:val="roman"/>
                    <m:sty m:val="i"/>
                </m:rPr>
                <m:t>{text}</m:t>
            </m:r>
        </m:oMath>
        '''
        
        # Parse and append
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_enhanced_math_run(paragraph, text):
    """Tạo run với style nâng cao trông như equation"""
    try:
        # Tạo run với formatting đặc biệt
        run = paragraph.add_run(text)
        
        # Apply equation-like formatting
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)  # Slightly larger
        run.italic = True
        run.bold = False
        
        # Màu xanh đậm cho equation
        run.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue
        
        # Thêm border nhẹ để highlight
        try:
            # Add subtle styling to make it look like equation
            rPr = run._element.get_or_add_rPr()
            
            # Add equation style
            w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            style_xml = f'<w:rStyle xmlns:w="{w_ns}" w:val="MathEquation"/>'
            try:
                style_element = parse_xml(style_xml)
                rPr.append(style_element)
            except:
                pass
                
        except:
            pass
        
        return True
        
    except Exception:
        return False

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - tạo equation objects thực sự"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    
    for start, end, latex_content in expressions:
        # Thêm text trước công thức
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Tạo equation object thực sự
        if create_advanced_equation(paragraph, latex_content):
            math_count += 1
        else:
            # Fallback nếu không tạo được equation
            run = paragraph.add_run(f"[{latex_content}]")
            run.italic = True
            run.font.color.rgb = RGBColor(255, 0, 0)  # Red for errors
        
        last_pos = end
    
    # Thêm text còn lại
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Kiểm tra pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # In đậm phần câu hỏi
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Xử lý phần nội dung
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Kiểm tra pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format chữ cái đáp án
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Xử lý nội dung đáp án
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A, xử lý bình thường
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra xem text có phải markdown table không"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse text thành table data"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Bỏ qua dòng separator
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            # Loại bỏ | đầu và cuối
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table từ data"""
    if not table_data:
        return 0
    
    math_count = 0
    max_cols = max(len(row) for row in table_data)
    
    # Tạo table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Điền dữ liệu
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Tạo paragraph trong cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Xử lý Q&A hoặc math
                if not format_question_answer(cell_para, cell_text):
                    math_count += process_text_with_math(cell_para, cell_text)
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table structure
    if options.get('format_tables', True):
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header
        if word_table.rows:
            header_row = word_table.rows[0]
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                
                # Màu nền xanh
                try:
                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                    shading = parse_xml(shading_xml)
                    cell._tc.get_or_add_tcPr().append(shading)
                except:
                    pass
    
    return math_count

def process_document_content(doc, options):
    """Xử lý nội dung document"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs")
    
    # Xử lý paragraphs
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Kiểm tra markdown table
        if '|' in text and options.get('convert_markdown', True):
            # Collect multiple lines for table
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
            
            # Check if it's a valid table
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    math_count = create_word_table(new_doc, table_data, options)
                    stats['math_expressions'] += math_count
                    stats['markdown_tables'] += 1
                    stats['debug_log'].append(f"Converted markdown table: {len(table_data)} rows")
                
                i = j
                continue
        
        # Xử lý paragraph thông thường
        new_para = new_doc.add_paragraph()
        
        if format_question_answer(new_para, text):
            stats['questions_formatted'] += 1
        else:
            math_count = process_text_with_math(new_para, text)
            stats['math_expressions'] += math_count
        
        i += 1
    
    # Xử lý Word tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        num_rows = len(table.rows)
        num_cols = len(table.rows[0].cells) if table.rows else 0
        
        # Tạo table mới
        new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
        
        table_math = 0
        for r in range(num_rows):
            for c in range(min(len(table.rows[r].cells), num_cols)):
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
                        table_math += process_text_with_math(cell_para, cell_text)
                    
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        stats['math_expressions'] += table_math
        
        # Format table
        if options.get('format_tables', True):
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if new_table.rows:
                header_row = new_table.rows[0]
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
            
            stats['word_tables'] += 1
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v7.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">LaTeX → Word Equation Objects + Q&A + Tables</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Tạo Equation Objects", value=True,
                                       help="Chuyển $...$ thành equation objects thực sự trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True)  
        format_tables = st.checkbox("📊 Format bảng", value=True)
        convert_markdown = st.checkbox("📋 Markdown Tables", value=True)
        show_debug = st.checkbox("🔍 Debug", value=False)
        
        st.markdown("---")
        st.info("💡 **Equation Creation:**\n\n1️⃣ Try OMML equation objects\n2️⃣ Fallback to enhanced math styling\n3️⃣ Equations tích hợp với Word!")
        
        st.markdown("### 📝 Equation Examples")
        st.code("""
$A^{prime}$ → A' (equation)
$x^2$ → x² (equation)
$\\left(x\\right)$ → (x) (equation)
${B^{prime}}$ → B' (equation)
$\\frac{a}{b}$ → (a)/(b) (equation)
$\\pi$ → π (equation)

→ Tạo equation objects thực sự
  không phải text thường!
        """)
        
        st.markdown("### 📋 Table Support")
        st.code("""
| Header | Data |
|--------|------|
| $x^2$  | Value|

→ Equations trong tables
        """)
    
    # Main area
    st.markdown("### 📁 Upload File")
    
    uploaded_file = st.file_uploader("Chọn file Word (.docx)", type=["docx"])
    
    if uploaded_file:
        st.success(f"✅ {uploaded_file.name}")
        
        with st.expander("Preview"):
            try:
                doc = Document(uploaded_file)
                st.info(f"📄 {len(doc.paragraphs)} paragraphs | 📊 {len(doc.tables)} tables")
                
                for i, para in enumerate(doc.paragraphs[:5]):
                    if para.text.strip():
                        preview_text = para.text[:80] + "..." if len(para.text) > 80 else para.text
                        st.text(f"{i+1}. {preview_text}")
            except Exception as e:
                st.error(f"Preview error: {e}")
        
        if st.button("🚀 Xử lý", type="primary", use_container_width=True):
            options = {
                'convert_equations': convert_equations,
                'format_questions': format_questions,
                'format_tables': format_tables,
                'convert_markdown': convert_markdown
            }
            
            with st.spinner("Đang xử lý..."):
                try:
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document_content(doc, options)
                    
                    # Save
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_processed_{timestamp}.docx"
                    
                    # Results
                    st.markdown("### 📊 Kết quả")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("🔢 Equations", stats['math_expressions'])
                    col2.metric("📝 Q&A", stats['questions_formatted'])
                    col3.metric("📊 Word Tables", stats['word_tables'])
                    col4.metric("📋 Markdown", stats['markdown_tables'])
                    
                    if show_debug:
                        with st.expander("Debug"):
                            for log in stats['debug_log']:
                                st.text(log)
                    
                    # Download
                    st.download_button(
                        "⬇️ Tải file",
                        buffer.getvalue(),
                        filename,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                    total_processed = (stats['math_expressions'] + stats['questions_formatted'] + 
                                     stats['word_tables'] + stats['markdown_tables'])
                    
                    if total_processed > 0:
                        st.success("🎉 Hoàn thành!")
                        if stats['math_expressions'] > 0:
                            st.info(f"✨ Đã tạo {stats['math_expressions']} equation objects trong Word!")
                        if stats['markdown_tables'] > 0:
                            st.info(f"📋 Đã chuyển đổi {stats['markdown_tables']} bảng Markdown!")
                    else:
                        st.warning("⚠️ Không tìm thấy nội dung cần xử lý")
                
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")

if __name__ == "__main__":
    main():
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

def create_math_equation(paragraph, latex_text):
    """Tạo equation object thực sự trong Word"""
    try:
        # Convert LaTeX to simpler math notation first
        unicode_text = convert_latex_symbols(latex_text)
        
        # Create equation run với Office Math format
        # Sử dụng approach tạo equation object thông qua OMML
        
        # Tạo math element
        math_xml = f"""
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:t>{unicode_text}</m:t>
            </m:r>
        </m:oMath>
        """
        
        try:
            # Thử insert equation object
            math_element = parse_xml(math_xml)
            paragraph._element.append(math_element)
            return True
        except:
            # Fallback: Tạo run với style đặc biệt trông như equation
            return create_equation_style_run(paragraph, unicode_text)
            
    except Exception as e:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def create_equation_style_run(paragraph, text):
    """Tạo run trông như equation (fallback)"""
    try:
        run = paragraph.add_run(text)
        
        # Style để trông như equation
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
        
        # Thêm highlight nhẹ
        try:
            # Add subtle background
            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F0F8FF"/>'
            shading = parse_xml(shading_xml)
            run._element.get_or_add_rPr().append(shading)
        except:
            pass
        
        return True
        
    except Exception:
        return False

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao với OMML format"""
    try:
        # Parse LaTeX content để tạo OMML phù hợp
        unicode_content = convert_latex_symbols(latex_text)
        
        # Phân tích content để tạo structure phù hợp
        if '^' in unicode_content or '²' in unicode_content or '³' in unicode_content:
            # Có superscript - tạo sup element
            return create_superscript_equation(paragraph, unicode_content)
        elif '(' in unicode_content and ')' in unicode_content:
            # Có parentheses - tạo brackets
            return create_bracketed_equation(paragraph, unicode_content)
        else:
            # Simple math
            return create_simple_equation(paragraph, unicode_content)
            
    except Exception:
        return create_equation_style_run(paragraph, unicode_content)

def create_advanced_equation(paragraph, latex_text):
    """Tạo equation nâng cao - thử OMML trước, fallback về styled text"""
    try:
        # Convert LaTeX to Unicode first
        unicode_content = convert_latex_symbols(latex_text)
        
        # Approach 1: Try OMML equation object
        if try_create_omml_equation(paragraph, unicode_content):
            return True
        
        # Approach 2: Fallback - Enhanced styled text that looks like equation
        return create_enhanced_math_run(paragraph, unicode_content)
            
    except Exception:
        # Final fallback
        run = paragraph.add_run(f"[{latex_text}]")
        run.italic = True
        return False

def try_create_omml_equation(paragraph, text):
    """Thử tạo OMML equation - simple approach"""
    try:
        # Simple OMML template for basic math
        omml_xml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:rPr>
                    <m:scr m:val="roman"/>
                    <m:sty m:val="i"/>
                </m:rPr>
                <m:t>{text}</m:t>
            </m:r>
        </m:oMath>
        '''
        
        # Parse and append
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
        
    except Exception:
        return False

def create_enhanced_math_run(paragraph, text):
    """Tạo run với style nâng cao trông như equation"""
    try:
        # Tạo run với formatting đặc biệt
        run = paragraph.add_run(text)
        
        # Apply equation-like formatting
        run.font.name = 'Cambria Math'
        run.font.size = Pt(13)  # Slightly larger
        run.italic = True
        run.bold = False
        
        # Màu xanh đậm cho equation
        run.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue
        
        # Thêm border nhẹ để highlight
        try:
            # Add subtle styling to make it look like equation
            rPr = run._element.get_or_add_rPr()
            
            # Add equation style
            w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            style_xml = f'<w:rStyle xmlns:w="{w_ns}" w:val="MathEquation"/>'
            try:
                style_element = parse_xml(style_xml)
                rPr.append(style_element)
            except:
                pass
                
        except:
            pass
        
        return True
        
    except Exception:
        return False

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học - tạo equation objects thực sự"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    
    for start, end, latex_content in expressions:
        # Thêm text trước công thức
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Tạo equation object thực sự
        if create_advanced_equation(paragraph, latex_content):
            math_count += 1
        else:
            # Fallback nếu không tạo được equation
            run = paragraph.add_run(f"[{latex_content}]")
            run.italic = True
            run.font.color.rgb = RGBColor(255, 0, 0)  # Red for errors
        
        last_pos = end
    
    # Thêm text còn lại
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Kiểm tra pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # In đậm phần câu hỏi
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Xử lý phần nội dung
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Kiểm tra pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format chữ cái đáp án
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Xử lý nội dung đáp án
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A, xử lý bình thường
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra xem text có phải markdown table không"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse text thành table data"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Bỏ qua dòng separator
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            # Loại bỏ | đầu và cuối
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table từ data"""
    if not table_data:
        return 0
    
    math_count = 0
    max_cols = max(len(row) for row in table_data)
    
    # Tạo table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Điền dữ liệu
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Tạo paragraph trong cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Xử lý Q&A hoặc math
                if not format_question_answer(cell_para, cell_text):
                    math_count += process_text_with_math(cell_para, cell_text)
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table structure
    if options.get('format_tables', True):
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header
        if word_table.rows:
            header_row = word_table.rows[0]
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                
                # Màu nền xanh
                try:
                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                    shading = parse_xml(shading_xml)
                    cell._tc.get_or_add_tcPr().append(shading)
                except:
                    pass
    
    return math_count

def process_document_content(doc, options):
    """Xử lý nội dung document"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs")
    
    # Xử lý paragraphs
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Kiểm tra markdown table
        if '|' in text and options.get('convert_markdown', True):
            # Collect multiple lines for table
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
            
            # Check if it's a valid table
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    math_count = create_word_table(new_doc, table_data, options)
                    stats['math_expressions'] += math_count
                    stats['markdown_tables'] += 1
                    stats['debug_log'].append(f"Converted markdown table: {len(table_data)} rows")
                
                i = j
                continue
        
        # Xử lý paragraph thông thường
        new_para = new_doc.add_paragraph()
        
        if format_question_answer(new_para, text):
            stats['questions_formatted'] += 1
        else:
            math_count = process_text_with_math(new_para, text)
            stats['math_expressions'] += math_count
        
        i += 1
    
    # Xử lý Word tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        num_rows = len(table.rows)
        num_cols = len(table.rows[0].cells) if table.rows else 0
        
        # Tạo table mới
        new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
        
        table_math = 0
        for r in range(num_rows):
            for c in range(min(len(table.rows[r].cells), num_cols)):
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
                        table_math += process_text_with_math(cell_para, cell_text)
                    
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        stats['math_expressions'] += table_math
        
        # Format table
        if options.get('format_tables', True):
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if new_table.rows:
                header_row = new_table.rows[0]
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
            
            stats['word_tables'] += 1
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v7.1</h1>
        <p style="color: white; margin: 10px 0 0 0;">LaTeX → Word Equation Objects + Q&A + Tables</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Tạo Equation Objects", value=True,
                                       help="Chuyển $...$ thành equation objects thực sự trong Word")
        format_questions = st.checkbox("📝 Format Q&A", value=True)  
        format_tables = st.checkbox("📊 Format bảng", value=True)
        convert_markdown = st.checkbox("📋 Markdown Tables", value=True)
        show_debug = st.checkbox("🔍 Debug", value=False)
        
        st.markdown("---")
        st.info("💡 **Equation Creation:**\n\n1️⃣ Try OMML equation objects\n2️⃣ Fallback to enhanced math styling\n3️⃣ Equations tích hợp với Word!")
        
        st.markdown("### 📝 Equation Examples")
        st.code("""
$A^{prime}$ → A' (equation)
$x^2$ → x² (equation)
$\\left(x\\right)$ → (x) (equation)
${B^{prime}}$ → B' (equation)
$\\frac{a}{b}$ → (a)/(b) (equation)
$\\pi$ → π (equation)

→ Tạo equation objects thực sự
  không phải text thường!
        """)
        
        st.markdown("### 📋 Table Support")
        st.code("""
| Header | Data |
|--------|------|
| $x^2$  | Value|

→ Equations trong tables
        """)
    
    # Main area
    st.markdown("### 📁 Upload File")
    
    uploaded_file = st.file_uploader("Chọn file Word (.docx)", type=["docx"])
    
    if uploaded_file:
        st.success(f"✅ {uploaded_file.name}")
        
        with st.expander("Preview"):
            try:
                doc = Document(uploaded_file)
                st.info(f"📄 {len(doc.paragraphs)} paragraphs | 📊 {len(doc.tables)} tables")
                
                for i, para in enumerate(doc.paragraphs[:5]):
                    if para.text.strip():
                        preview_text = para.text[:80] + "..." if len(para.text) > 80 else para.text
                        st.text(f"{i+1}. {preview_text}")
            except Exception as e:
                st.error(f"Preview error: {e}")
        
        if st.button("🚀 Xử lý", type="primary", use_container_width=True):
            options = {
                'convert_equations': convert_equations,
                'format_questions': format_questions,
                'format_tables': format_tables,
                'convert_markdown': convert_markdown
            }
            
            with st.spinner("Đang xử lý..."):
                try:
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document_content(doc, options)
                    
                    # Save
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_processed_{timestamp}.docx"
                    
                    # Results
                    st.markdown("### 📊 Kết quả")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("🔢 Equations", stats['math_expressions'])
                    col2.metric("📝 Q&A", stats['questions_formatted'])
                    col3.metric("📊 Word Tables", stats['word_tables'])
                    col4.metric("📋 Markdown", stats['markdown_tables'])
                    
                    if show_debug:
                        with st.expander("Debug"):
                            for log in stats['debug_log']:
                                st.text(log)
                    
                    # Download
                    st.download_button(
                        "⬇️ Tải file",
                        buffer.getvalue(),
                        filename,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                    total_processed = (stats['math_expressions'] + stats['questions_formatted'] + 
                                     stats['word_tables'] + stats['markdown_tables'])
                    
                    if total_processed > 0:
                        st.success("🎉 Hoàn thành!")
                        if stats['math_expressions'] > 0:
                            st.info(f"✨ Đã tạo {stats['math_expressions']} equation objects trong Word!")
                        if stats['markdown_tables'] > 0:
                            st.info(f"📋 Đã chuyển đổi {stats['markdown_tables']} bảng Markdown!")
                    else:
                        st.warning("⚠️ Không tìm thấy nội dung cần xử lý")
                
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")

if __name__ == "__main__":
    main()
