#!/usr/bin/env python3
"""
📝 Module xử lý đề thi Word thành bảng chuẩn hóa
Hỗ trợ LaTeX math và Markdown tables
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml.parser import OxmlElement
import re
from io import BytesIO
from datetime import datetime
import markdown
from markdown.extensions import tables
import html2text

def convert_latex_to_word_equation(text):
    """Chuyển đổi LaTeX math thành Word equation"""
    
    # Pattern để tìm các công thức LaTeX
    patterns = [
        (r'\$\$([^$]+)\$\$', 'display'),  # Display math $$...$$
        (r'\$([^$]+)\$', 'inline'),       # Inline math $...$
        (r'\\\[([^\]]+)\\\]', 'display'), # Display math \[...\]
        (r'\\\(([^\)]+)\\\)', 'inline')   # Inline math \(...\)
    ]
    
    equations = []
    processed_text = text
    
    for pattern, math_type in patterns:
        matches = re.finditer(pattern, text)
        for match in matches:
            latex_code = match.group(1).strip()
            
            # Lưu thông tin equation
            equations.append({
                'latex': latex_code,
                'type': math_type,
                'original': match.group(0),
                'position': match.span()
            })
            
            # Thay thế trong text bằng placeholder
            placeholder = f"[EQUATION_{len(equations)-1}]"
            processed_text = processed_text.replace(match.group(0), placeholder, 1)
    
    return processed_text, equations

def insert_equation_to_paragraph(paragraph, latex_code, math_type='inline'):
    """Chèn equation vào paragraph Word"""
    
    try:
        # Tạo OMML (Office Math Markup Language) từ LaTeX
        omml = convert_latex_to_omml(latex_code)
        
        if omml:
            # Thêm equation vào paragraph
            run = paragraph.add_run()
            math_element = OxmlElement('m:oMath')
            math_element.append(parse_xml(omml))
            run._element.append(math_element)
        else:
            # Fallback: thêm LaTeX code gốc với formatting
            run = paragraph.add_run(f"${latex_code}$")
            run.font.italic = True
            run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
            
    except Exception as e:
        # Fallback: thêm LaTeX code gốc
        run = paragraph.add_run(f"${latex_code}$")
        run.font.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)

def convert_latex_to_omml(latex_code):
    """Chuyển đổi LaTeX thành OMML (Office Math Markup Language)"""
    
    # Mapping cơ bản từ LaTeX sang OMML
    latex_to_omml_map = {
        # Superscript
        r'\^{([^}]+)}': r'<m:sSup><m:e><m:r><m:t>\1</m:t></m:r></m:e></m:sSup>',
        r'\^(\w)': r'<m:sSup><m:e><m:r><m:t>\1</m:t></m:r></m:e></m:sSup>',
        
        # Subscript  
        r'_{([^}]+)}': r'<m:sSub><m:e><m:r><m:t>\1</m:t></m:r></m:e></m:sSub>',
        r'_(\w)': r'<m:sSub><m:e><m:r><m:t>\1</m:t></m:r></m:e></m:sSub>',
        
        # Fractions
        r'\\frac{([^}]+)}{([^}]+)}': r'<m:f><m:num><m:r><m:t>\1</m:t></m:r></m:num><m:den><m:r><m:t>\2</m:t></m:r></m:den></m:f>',
        
        # Square root
        r'\\sqrt{([^}]+)}': r'<m:rad><m:deg></m:deg><m:e><m:r><m:t>\1</m:t></m:r></m:e></m:rad>',
        
        # Greek letters
        r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ',
        r'\\epsilon': 'ε', r'\\zeta': 'ζ', r'\\eta': 'η', r'\\theta': 'θ',
        r'\\iota': 'ι', r'\\kappa': 'κ', r'\\lambda': 'λ', r'\\mu': 'μ',
        r'\\nu': 'ν', r'\\xi': 'ξ', r'\\pi': 'π', r'\\rho': 'ρ',
        r'\\sigma': 'σ', r'\\tau': 'τ', r'\\upsilon': 'υ', r'\\phi': 'φ',
        r'\\chi': 'χ', r'\\psi': 'ψ', r'\\omega': 'ω',
        
        # Operators
        r'\\sum': '∑', r'\\prod': '∏', r'\\int': '∫',
        r'\\infty': '∞', r'\\pm': '±', r'\\mp': '∓',
        r'\\times': '×', r'\\div': '÷', r'\\cdot': '·',
        r'\\leq': '≤', r'\\geq': '≥', r'\\neq': '≠',
        r'\\approx': '≈', r'\\equiv': '≡',
        
        # Parentheses
        r'\\left\(': '(', r'\\right\)': ')',
        r'\\left\[': '[', r'\\right\]': ']',
        r'\\left\{': '{', r'\\right\}': '}',
    }
    
    try:
        processed = latex_code
        
        # Áp dụng các transformations
        for latex_pattern, omml_replacement in latex_to_omml_map.items():
            processed = re.sub(latex_pattern, omml_replacement, processed)
        
        # Wrap trong OMML structure cơ bản
        omml = f'''
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r>
                <m:t>{processed}</m:t>
            </m:r>
        </m:oMath>
        '''
        
        return omml.strip()
        
    except Exception:
        return None

def parse_markdown_table(text):
    """Phân tích bảng Markdown và trả về dữ liệu bảng"""
    
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|[^|\n]*\|(?:\n\|[^|\n]*\|)*)'
    
    tables = []
    matches = re.finditer(table_pattern, text, re.MULTILINE)
    
    for match in matches:
        table_text = match.group(1).strip()
        lines = table_text.split('\n')
        
        if len(lines) < 2:
            continue
        
        # Parse header
        header_line = lines[0]
        headers = [cell.strip() for cell in header_line.split('|')[1:-1]]
        
        # Skip separator line (usually contains dashes)
        if len(lines) > 2 and '---' in lines[1]:
            data_lines = lines[2:]
        else:
            data_lines = lines[1:]
        
        # Parse data rows
        rows = []
        for line in data_lines:
            if line.strip():
                cells = [cell.strip() for cell in line.split('|')[1:-1]]
                # Ensure same number of columns
                while len(cells) < len(headers):
                    cells.append('')
                rows.append(cells[:len(headers)])
        
        if headers and rows:
            tables.append({
                'headers': headers,
                'rows': rows,
                'original': match.group(0),
                'position': match.span()
            })
    
    return tables

def process_text_with_math_and_tables(text):
    """Xử lý text chứa công thức LaTeX và bảng Markdown"""
    
    result = {
        'processed_text': text,
        'equations': [],
        'tables': [],
        'has_math': False,
        'has_tables': False
    }
    
    # Xử lý công thức LaTeX
    processed_text, equations = convert_latex_to_word_equation(text)
    if equations:
        result['equations'] = equations
        result['has_math'] = True
        result['processed_text'] = processed_text
    
    # Xử lý bảng Markdown
    tables = parse_markdown_table(result['processed_text'])
    if tables:
        result['tables'] = tables
        result['has_tables'] = True
        
        # Thay thế bảng bằng placeholder
        for i, table in enumerate(tables):
            placeholder = f"[TABLE_{i}]"
            result['processed_text'] = result['processed_text'].replace(
                table['original'], placeholder, 1
            )
    
    return result

def add_enhanced_paragraph(doc, text, style=None):
    """Thêm paragraph có hỗ trợ LaTeX và Markdown tables"""
    
    processed = process_text_with_math_and_tables(text)
    
    if not processed['has_math'] and not processed['has_tables']:
        # Text thường không có công thức hay bảng
        para = doc.add_paragraph(processed['processed_text'])
        if style:
            para.style = style
        return para
    
    # Xử lý text có công thức và/hoặc bảng
    para = doc.add_paragraph()
    if style:
        para.style = style
    
    current_text = processed['processed_text']
    
    # Thêm từng phần của text
    parts = re.split(r'(\[(?:EQUATION|TABLE)_\d+\])', current_text)
    
    for part in parts:
        if not part:
            continue
            
        # Kiểm tra nếu là equation placeholder
        eq_match = re.match(r'\[EQUATION_(\d+)\]', part)
        if eq_match:
            eq_index = int(eq_match.group(1))
            if eq_index < len(processed['equations']):
                equation = processed['equations'][eq_index]
                insert_equation_to_paragraph(para, equation['latex'], equation['type'])
            continue
        
        # Kiểm tra nếu là table placeholder
        table_match = re.match(r'\[TABLE_(\d+)\]', part)
        if table_match:
            table_index = int(table_match.group(1))
            if table_index < len(processed['tables']):
                table_data = processed['tables'][table_index]
                # Tạo bảng Word từ dữ liệu Markdown
                create_table_from_markdown(doc, table_data)
            continue
        
        # Text thường
        if part.strip():
            para.add_run(part)
    
    return para

def create_table_from_markdown(doc, table_data):
    """Tạo bảng Word từ dữ liệu Markdown table"""
    
    headers = table_data['headers']
    rows = table_data['rows']
    
    if not headers or not rows:
        return None
    
    # Tạo bảng
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Thiết lập độ rộng cột tự động
    table.autofit = True
    
    # Định dạng header
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = header_cells[i]
        cell.text = header
        
        # Định dạng text trong cell
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(10)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Màu nền header
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shading_elm = parse_xml(r'<w:shd {} w:fill="4472C4"/>'.format(
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'))
        cell._tc.get_or_add_tcPr().append(shading_elm)
    
    # Thêm dữ liệu
    for row_data in rows:
        row_cells = table.add_row().cells
        
        for i, cell_data in enumerate(row_data):
            if i < len(row_cells):
                cell = row_cells[i]
                
                # Xử lý cell có thể chứa LaTeX
                processed = process_text_with_math_and_tables(cell_data)
                
                if processed['has_math']:
                    # Thêm text và equations vào cell
                    para = cell.paragraphs[0]
                    para.clear()
                    
                    parts = re.split(r'(\[EQUATION_\d+\])', processed['processed_text'])
                    for part in parts:
                        eq_match = re.match(r'\[EQUATION_(\d+)\]', part)
                        if eq_match:
                            eq_index = int(eq_match.group(1))
                            if eq_index < len(processed['equations']):
                                equation = processed['equations'][eq_index]
                                insert_equation_to_paragraph(para, equation['latex'])
                        else:
                            if part.strip():
                                para.add_run(part)
                else:
                    cell.text = cell_data
                
                # Định dạng text
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    
    # Thêm border cho bảng
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            
            borders = parse_xml(r'''
                <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>
                </w:tcBorders>
            ''')
            tcPr.append(borders)
    
    # Thêm khoảng cách sau bảng
    doc.add_paragraph()
    
    return table

def analyze_exam_structure(doc):
    """Phân tích cấu trúc đề thi với hỗ trợ LaTeX và Markdown"""
    
    analysis = {
        'total_questions': 0,
        'multiple_choice': 0,
        'true_false': 0,
        'essay': 0,
        'sample_questions': [],
        'sections': [],
        'has_math_formulas': False,
        'has_markdown_tables': False,
        'math_count': 0,
        'table_count': 0
    }
    
    current_section = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        
        if not text:
            continue
        
        # Kiểm tra có công thức LaTeX không
        if re.search(r'\$[^$]+\$|\\\[[^\]]+\\\]', text):
            analysis['has_math_formulas'] = True
            analysis['math_count'] += len(re.findall(r'\$[^$]+\$|\\\[[^\]]+\\\]', text))
        
        # Kiểm tra có bảng Markdown không
        if re.search(r'\|[^|\n]*\|', text):
            analysis['has_markdown_tables'] = True
            tables = parse_markdown_table(text)
            analysis['table_count'] += len(tables)
        
        # Nhận dạng phần đề thi
        section_match = re.match(r'(PHẦN|Phần)\s*(I+|1|2|3|II|III)', text, re.IGNORECASE)
        if section_match:
            current_section = {
                'name': text,
                'type': determine_section_type(text),
                'questions': []
            }
            analysis['sections'].append(current_section)
            continue
        
        # Nhận dạng câu hỏi
        question_match = re.match(r'(Câu|CÂU)\s*(\d+)', text, re.IGNORECASE)
        if question_match:
            analysis['total_questions'] += 1
            
            # Phân loại câu hỏi
            question_type = classify_question_type(text, para)
            
            if question_type == 'multiple_choice':
                analysis['multiple_choice'] += 1
            elif question_type == 'true_false':
                analysis['true_false'] += 1
            elif question_type == 'essay':
                analysis['essay'] += 1
            
            # Lưu mẫu câu hỏi
            if len(analysis['sample_questions']) < 5:
                parsed_question = parse_question(text, question_type)
                if parsed_question:
                    analysis['sample_questions'].append(parsed_question)
    
    return analysis

def extract_choice_question(text):
    """Trích xuất câu hỏi trắc nghiệm có hỗ trợ LaTeX"""
    
    # Xử lý công thức LaTeX trước
    processed = process_text_with_math_and_tables(text)
    working_text = processed['processed_text']
    
    # Pattern để bắt câu hỏi trắc nghiệm
    patterns = [
        r'(Câu\s*\d+)\.(.*?)\s*A\.\s*(.*?)\s*B\.\s*(.*?)\s*C\.\s*(.*?)\s*D\.\s*(.*?)$',
        r'(Câu\s*\d+)\.(.*?)(?=\s*A\.)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, working_text, re.DOTALL | re.IGNORECASE)
        if match:
            groups = match.groups()
            
            result = {
                'number': groups[0].replace('Câu', '').strip() if groups[0] else '',
                'question': groups[1].strip() if len(groups) > 1 and groups[1] else '',
                'A': groups[2].strip() if len(groups) > 2 and groups[2] else '',
                'B': groups[3].strip() if len(groups) > 3 and groups[3] else '',
                'C': groups[4].strip() if len(groups) > 4 and groups[4] else '',
                'D': groups[5].strip() if len(groups) > 5 and groups[5] else '',
                'has_math': processed['has_math'],
                'equations': processed['equations']
            }
            
            # Khôi phục công thức LaTeX vào các trường
            if processed['has_math']:
                for field in ['question', 'A', 'B', 'C', 'D']:
                    if result[field]:
                        result[field] = restore_equations_in_text(result[field], processed['equations'])
            
            if result['question']:
                return result
    
    # Fallback parsing
    return extract_choice_question_fallback(text)

def restore_equations_in_text(text, equations):
    """Khôi phục equations từ placeholder về LaTeX gốc"""
    
    for i, eq in enumerate(equations):
        placeholder = f"[EQUATION_{i}]"
        if placeholder in text:
            text = text.replace(placeholder, eq['original'])
    
    return text

def extract_choice_question_fallback(text):
    """Fallback method cho parsing câu trắc nghiệm"""
    
    lines = text.split('\n')
    result = {'number': '', 'question': '', 'A': '', 'B': '', 'C': '', 'D': ''}
    
    current_key = 'question'
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Tìm số câu hỏi
        if re.match(r'Câu\s*\d+', line, re.IGNORECASE):
            result['number'] = re.search(r'\d+', line).group()
            result['question'] = re.sub(r'Câu\s*\d+\.?\s*', '', line, flags=re.IGNORECASE)
            continue
        
        # Tìm đáp án
        for option in ['A', 'B', 'C', 'D']:
            if line.startswith(f'{option}.'):
                result[option] = line[2:].strip()
                current_key = option
                break
        else:
            # Nếu không phải đáp án mới, nối vào key hiện tại
            if current_key and result[current_key]:
                result[current_key] += ' ' + line
            elif current_key == 'question':
                result['question'] += ' ' + line
    
    return result if result['question'] else None

def create_professional_table_enhanced(doc, title, headers, data):
    """Tạo bảng Word chuyên nghiệp có hỗ trợ LaTeX"""
    
    # Thêm tiêu đề bảng
    title_para = doc.add_paragraph()
    title_run = title_para.add_run(title)
    title_run.font.size = Pt(14)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(31, 78, 121)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.space_after = Pt(12)
    
    # Tạo bảng
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Thiết lập độ rộng cột
    if len(headers) == 6:  # Bảng trắc nghiệm
        widths = [Inches(0.5), Inches(3.0), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5)]
    elif len(headers) == 3:  # Bảng đúng/sai hoặc tự luận
        widths = [Inches(0.5), Inches(4.5), Inches(1.5)]
    else:
        widths = [Inches(1.5)] * len(headers)
    
    for i, width in enumerate(widths[:len(headers)]):
        table.columns[i].width = width
    
    # Định dạng header
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = header_cells[i]
        cell.text = header
        
        # Định dạng text trong cell
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(11)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Màu nền header
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shading_elm = parse_xml(r'<w:shd {} w:fill="1f4e79"/>'.format(
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'))
        cell._tc.get_or_add_tcPr().append(shading_elm)
    
    # Thêm dữ liệu
    for row_data in data:
        row_cells = table.add_row().cells
        
        for i, cell_data in enumerate(row_data):
            if i < len(row_cells):
                cell = row_cells[i]
                
                # Xử lý cell có thể chứa LaTeX
                if isinstance(cell_data, str) and ('$' in cell_data or '\\' in cell_data):
                    processed = process_text_with_math_and_tables(cell_data)
                    
                    if processed['has_math'] or processed['has_tables']:
                        # Xóa nội dung cell hiện tại
                        para = cell.paragraphs[0]
                        para.clear()
                        
                        # Thêm nội dung có công thức
                        parts = re.split(r'(\[(?:EQUATION|TABLE)_\d+\])', processed['processed_text'])
                        
                        for part in parts:
                            eq_match = re.match(r'\[EQUATION_(\d+)\]', part)
                            table_match = re.match(r'\[TABLE_(\d+)\]', part)
                            
                            if eq_match:
                                eq_index = int(eq_match.group(1))
                                if eq_index < len(processed['equations']):
                                    equation = processed['equations'][eq_index]
                                    insert_equation_to_paragraph(para, equation['latex'])
                            elif table_match:
                                # Bảng trong cell - chuyển thành text
                                table_index = int(table_match.group(1))
                                if table_index < len(processed['tables']):
                                    table_data = processed['tables'][table_index]
                                    table_text = format_table_as_text(table_data)
                                    para.add_run(table_text)
                            else:
                                if part.strip():
                                    para.add_run(part)
                    else:
                        cell.text = str(cell_data) if cell_data else ''
                else:
                    cell.text = str(cell_data) if cell_data else ''
                
                # Định dạng text
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if not hasattr(run, '_element') or 'm:oMath' not in str(run._element.xml):
                            run.font.size = Pt(10)
                    
                    if i == 0:  # Cột STT
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    
    # Thêm border cho bảng
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            
            borders = parse_xml(r'''
                <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="1f4e79"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="1f4e79"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="1f4e79"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="1f4e79"/>
                </w:tcBorders>
            ''')
            tcPr.append(borders)
    
    # Thêm khoảng cách sau bảng
    doc.add_paragraph().space_after = Pt(18)
    
    return table

def format_table_as_text(table_data):
    """Chuyển bảng Markdown thành text cho cell"""
    
    headers = table_data['headers']
    rows = table_data['rows']
    
    result = []
    result.append(' | '.join(headers))
    result.append(' | '.join(['---'] * len(headers)))
    
    for row in rows:
        result.append(' | '.join(row))
    
    return '\n'.join(result)

def determine_section_type(section_text):
    """Xác định loại phần đề thi"""
    section_lower = section_text.lower()
    
    if any(keyword in section_lower for keyword in ['trắc nghiệm', 'lựa chọn', 'multiple']):
        return 'multiple_choice'
    elif any(keyword in section_lower for keyword in ['đúng', 'sai', 'true', 'false']):
        return 'true_false'
    elif any(keyword in section_lower for keyword in ['tự luận', 'essay', 'viết']):
        return 'essay'
    else:
        return 'unknown'

def classify_question_type(text, paragraph=None):
    """Phân loại loại câu hỏi dựa trên nội dung"""
    
    # Kiểm tra có đáp án A, B, C, D không
    if re.search(r'[A-D]\.', text):
        return 'multiple_choice'
    
    # Kiểm tra câu đúng/sai
    if any(keyword in text.lower() for keyword in ['đúng', 'sai', 'true', 'false']):
        return 'true_false'
    
    # Mặc định là tự luận
    return 'essay'

def parse_question(text, question_type):
    """Phân tích chi tiết câu hỏi với hỗ trợ LaTeX"""
    
    result = {
        'type': question_type,
        'number': '',
        'question': '',
        'raw_text': text,
        'has_math': False,
        'has_tables': False
    }
    
    # Tách số câu hỏi
    number_match = re.match(r'(Câu|CÂU)\s*(\d+)', text, re.IGNORECASE)
    if number_match:
        result['number'] = number_match.group(2)
    
    # Kiểm tra có LaTeX và Markdown
    processed = process_text_with_math_and_tables(text)
    result['has_math'] = processed['has_math']
    result['has_tables'] = processed['has_tables']
    
    if question_type == 'multiple_choice':
        # Phân tích câu trắc nghiệm
        mc_parsed = extract_choice_question(text)
        if mc_parsed:
            result.update(mc_parsed)
    
    elif question_type == 'true_false':
        # Phân tích câu đúng/sai
        tf_parsed = extract_true_false_question(text)
        if tf_parsed:
            result.update(tf_parsed)
    
    else:
        # Câu tự luận
        essay_parsed = extract_essay_question(text)
        if essay_parsed:
            result.update(essay_parsed)
    
    return result

def extract_true_false_question(text):
    """Trích xuất câu hỏi đúng/sai"""
    
    result = {'number': '', 'question': '', 'type': 'true_false'}
    
    # Tìm số câu hỏi
    number_match = re.search(r'Câu\s*(\d+)', text, re.IGNORECASE)
    if number_match:
        result['number'] = number_match.group(1)
    
    # Lấy nội dung câu hỏi
    question_text = re.sub(r'Câu\s*\d+\.?\s*', '', text, flags=re.IGNORECASE)
    result['question'] = question_text.strip()
    
    return result

def extract_essay_question(text):
    """Trích xuất câu hỏi tự luận"""
    
    result = {'number': '', 'question': '', 'type': 'essay'}
    
    # Tìm số câu hỏi
    number_match = re.search(r'Câu\s*(\d+)', text, re.IGNORECASE)
    if number_match:
        result['number'] = number_match.group(1)
    
    # Lấy nội dung câu hỏi
    question_text = re.sub(r'Câu\s*\d+\.?\s*', '', text, flags=re.IGNORECASE)
    result['question'] = question_text.strip()
    
    return result

def process_docx(uploaded_file):
    """Xử lý file Word đề thi thành bảng chuẩn hóa với hỗ trợ LaTeX và Markdown"""
    
    # Đọc document gốc
    doc = Document(uploaded_file)
    
    # Tạo document mới
    new_doc = Document()
    
    # Thiết lập margins
    sections = new_doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Thêm header chính
    header = new_doc.add_heading('ĐỀ THI ĐÃ CHUẨN HÓA', level=1)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in header.runs:
        run.font.color.rgb = RGBColor(31, 78, 121)
        run.font.size = Pt(16)
    
    # Thêm thông tin ngày tạo
    date_para = new_doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in date_para.runs:
        run.font.italic = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(128, 128, 128)
    
    new_doc.add_paragraph()  # Khoảng trống
    
    # Phân tích cấu trúc
    analysis = analyze_exam_structure(doc)
    
    # Thêm thông tin về math và tables nếu có
    if analysis['has_math_formulas'] or analysis['has_markdown_tables']:
        info_para = new_doc.add_paragraph()
        info_text = "📊 Đề thi này có chứa: "
        
        features = []
        if analysis['has_math_formulas']:
            features.append(f"{analysis['math_count']} công thức toán học")
        if analysis['has_markdown_tables']:
            features.append(f"{analysis['table_count']} bảng dữ liệu")
        
        info_text += ", ".join(features)
        info_para.add_run(info_text).font.italic = True
        info_para.space_after = Pt(12)
    
    # Thu thập tất cả câu hỏi
    all_questions = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if re.match(r'Câu\s*\d+', text, re.IGNORECASE):
            question_type = classify_question_type(text)
            parsed = parse_question(text, question_type)
            if parsed:
                all_questions.append(parsed)
    
    # Phân loại câu hỏi
    multiple_choice_questions = [q for q in all_questions if q.get('type') == 'multiple_choice']
    true_false_questions = [q for q in all_questions if q.get('type') == 'true_false']
    essay_questions = [q for q in all_questions if q.get('type') == 'essay']
    
    # Tạo bảng cho câu hỏi trắc nghiệm
    if multiple_choice_questions:
        mc_headers = ['STT', 'Câu hỏi', 'A', 'B', 'C', 'D']
        mc_data = []
        
        for i, q in enumerate(multiple_choice_questions, 1):
            mc_data.append([
                str(i),
                q.get('question', ''),
                q.get('A', ''),
                q.get('B', ''),
                q.get('C', ''),
                q.get('D', '')
            ])
        
        create_professional_table_enhanced(new_doc, "PHẦN I: CÂU HỎI TRẮC NGHIỆM", mc_headers, mc_data)
    
    # Tạo bảng cho câu hỏi đúng/sai
    if true_false_questions:
        tf_headers = ['STT', 'Câu hỏi', 'Đáp án']
        tf_data = []
        
        for i, q in enumerate(true_false_questions, 1):
            tf_data.append([
                str(i),
                q.get('question', ''),
                ''  # Để trống cho giáo viên điền
            ])
        
        create_professional_table_enhanced(new_doc, "PHẦN II: CÂU HỎI ĐÚNG/SAI", tf_headers, tf_data)
    
    # Tạo bảng cho câu hỏi tự luận
    if essay_questions:
        essay_headers = ['STT', 'Câu hỏi', 'Ghi chú']
        essay_data = []
        
        for i, q in enumerate(essay_questions, 1):
            essay_data.append([
                str(i),
                q.get('question', ''),
                ''  # Để trống cho ghi chú
            ])
        
        create_professional_table_enhanced(new_doc, "PHẦN III: CÂU HỎI TỰ LUẬN", essay_headers, essay_data)
    
    # Thống kê
    stats = {
        'total_processed': len(all_questions),
        'multiple_choice_processed': len(multiple_choice_questions),
        'true_false_processed': len(true_false_questions),
        'essay_processed': len(essay_questions),
        'math_formulas': analysis['math_count'],
        'markdown_tables': analysis['table_count'],
        'processing_details': [
            f"Đã phân tích {len(all_questions)} câu hỏi tổng cộng",
            f"Trắc nghiệm: {len(multiple_choice_questions)} câu",
            f"Đúng/Sai: {len(true_false_questions)} câu", 
            f"Tự luận: {len(essay_questions)} câu"
        ]
    }
    
    if analysis['math_count'] > 0:
        stats['processing_details'].append(f"Đã xử lý {analysis['math_count']} công thức toán học")
    
    if analysis['table_count'] > 0:
        stats['processing_details'].append(f"Đã chuyển đổi {analysis['table_count']} bảng Markdown")
    
    return new_doc, stats
