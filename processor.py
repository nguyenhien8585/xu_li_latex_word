#!/usr/bin/env python3
"""
🧮 Module xử lý LaTeX và Markdown trong Word
Chuyển $...$ thành Equation và bảng Markdown thành Table Word
"""

import os
import sys
import re
import tempfile
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import OxmlElement
from io import BytesIO
from datetime import datetime

# Kiểm tra hỗ trợ Win32COM (chỉ trên Windows)
HAS_WIN32COM = False
try:
    if sys.platform == "win32":
        import win32com.client
        HAS_WIN32COM = True
except ImportError:
    pass

def detect_latex_patterns(text):
    """Phát hiện các pattern LaTeX trong text"""
    
    patterns = [
        (r'\$\{([^}]+)\}\$', 'inline'),      # ${...}$
        (r'\$([^$]+)\$', 'inline'),          # $...$
        (r'\\?\\\(([^)]+)\\\)', 'inline'),   # \(...\)
        (r'\\?\\\[([^\]]+)\\\]', 'display')  # \[...\]
    ]
    
    found_formulas = []
    
    for pattern, math_type in patterns:
        matches = re.finditer(pattern, text)
        for match in matches:
            latex_content = match.group(1)
            found_formulas.append({
                'original': match.group(0),
                'latex': latex_content,
                'type': math_type,
                'start': match.start(),
                'end': match.end()
            })
    
    # Sắp xếp theo vị trí để xử lý từ cuối về đầu
    found_formulas.sort(key=lambda x: x['start'], reverse=True)
    
    return found_formulas

def detect_markdown_tables(text):
    """Phát hiện bảng Markdown trong text"""
    
    # Pattern cho bảng Markdown
    table_pattern = r'(\|[^\n]*\|\n\|[\s\-:]*\|(?:\n\|[^\n]*\|)*)'
    
    tables = []
    matches = re.finditer(table_pattern, text, re.MULTILINE)
    
    for match in matches:
        table_text = match.group(1).strip()
        lines = table_text.split('\n')
        
        if len(lines) >= 2:  # Ít nhất header + separator
            # Parse header
            header = [cell.strip() for cell in lines[0].strip('|').split('|')]
            
            # Skip separator line
            data_lines = lines[2:] if len(lines) > 2 else []
            
            # Parse data rows
            rows = []
            for line in data_lines:
                if line.strip():
                    cells = [cell.strip() for cell in line.strip('|').split('|')]
                    # Đảm bảo số cột giống header
                    while len(cells) < len(header):
                        cells.append('')
                    rows.append(cells[:len(header)])
            
            if header and len(header) > 0:
                tables.append({
                    'original': match.group(0),
                    'header': header,
                    'rows': rows,
                    'start': match.start(),
                    'end': match.end()
                })
    
    return tables

def convert_latex_to_mathml(latex_code):
    """Chuyển LaTeX thành MathML (nếu có thư viện hỗ trợ)"""
    
    try:
        # Thử dùng latex2mathml
        from latex2mathml import converter
        mathml = converter.convert(latex_code)
        return mathml
    except ImportError:
        pass
    
    try:
        # Thử dùng sympy
        import sympy as sp
        expr = sp.sympify(latex_code)
        # Chuyển thành presentation MathML
        from sympy.printing.mathml import mathml
        mathml_output = mathml(expr, printer='presentation')
        return mathml_output
    except:
        pass
    
    return None

def create_equation_placeholder(latex_code):
    """Tạo placeholder cho equation để xử lý sau bằng Win32COM"""
    return f"[EQUATION:{latex_code}]"

def process_latex_in_paragraph(doc, text):
    """Xử lý LaTeX trong đoạn văn và tạo paragraph mới"""
    
    formulas = detect_latex_patterns(text)
    
    if not formulas:
        # Không có công thức, tạo paragraph thường
        return doc.add_paragraph(text)
    
    # Có công thức, xử lý từng phần
    para = doc.add_paragraph()
    current_pos = 0
    
    # Sắp xếp lại theo thứ tự xuất hiện
    formulas.sort(key=lambda x: x['start'])
    
    for formula in formulas:
        # Thêm text trước công thức
        if formula['start'] > current_pos:
            before_text = text[current_pos:formula['start']]
            if before_text:
                para.add_run(before_text)
        
        # Thêm công thức
        if HAS_WIN32COM:
            # Sẽ xử lý bằng Win32COM sau
            placeholder = create_equation_placeholder(formula['latex'])
            run = para.add_run(placeholder)
            run.font.color.rgb = RGBColor(0, 100, 200)  # Màu xanh để nhận biết
        else:
            # Fallback: hiển thị dạng italic
            formula_run = para.add_run(f"${formula['latex']}$")
            formula_run.font.italic = True
            formula_run.font.color.rgb = RGBColor(139, 0, 0)  # Màu đỏ để phân biệt
        
        current_pos = formula['end']
    
    # Thêm text cuối
    if current_pos < len(text):
        remaining_text = text[current_pos:]
        if remaining_text:
            para.add_run(remaining_text)
    
    return para

def create_word_table(doc, header, rows):
    """Tạo bảng Word từ dữ liệu header và rows"""
    
    if not header or not rows:
        return None
    
    # Tạo bảng
    table = doc.add_table(rows=1, cols=len(header))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Thiết lập header
    header_cells = table.rows[0].cells
    for i, header_text in enumerate(header):
        cell = header_cells[i]
        cell.text = header_text
        
        # Định dạng header
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(11)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Màu nền header
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        try:
            shading_elm = parse_xml(r'<w:shd {} w:fill="4472C4"/>'.format(
                'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'))
            cell._tc.get_or_add_tcPr().append(shading_elm)
        except:
            pass
    
    # Thêm dữ liệu
    for row_data in rows:
        row_cells = table.add_row().cells
        for i, cell_data in enumerate(row_data):
            if i < len(row_cells):
                cell = row_cells[i]
                
                # Kiểm tra nếu cell có LaTeX
                if '$' in cell_data or '\\(' in cell_data:
                    # Xử lý LaTeX trong cell
                    formulas = detect_latex_patterns(cell_data)
                    if formulas:
                        # Xóa nội dung cũ
                        cell.paragraphs[0].clear()
                        
                        # Thêm nội dung mới với LaTeX
                        current_pos = 0
                        formulas.sort(key=lambda x: x['start'])
                        
                        para = cell.paragraphs[0]
                        for formula in formulas:
                            # Text trước công thức
                            if formula['start'] > current_pos:
                                before_text = cell_data[current_pos:formula['start']]
                                if before_text:
                                    para.add_run(before_text)
                            
                            # Công thức
                            if HAS_WIN32COM:
                                placeholder = create_equation_placeholder(formula['latex'])
                                run = para.add_run(placeholder)
                                run.font.color.rgb = RGBColor(0, 100, 200)
                            else:
                                formula_run = para.add_run(f"${formula['latex']}$")
                                formula_run.font.italic = True
                                formula_run.font.color.rgb = RGBColor(139, 0, 0)
                            
                            current_pos = formula['end']
                        
                        # Text cuối
                        if current_pos < len(cell_data):
                            remaining_text = cell_data[current_pos:]
                            if remaining_text:
                                para.add_run(remaining_text)
                    else:
                        cell.text = cell_data
                else:
                    cell.text = cell_data
                
                # Định dạng cell
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if not ('[EQUATION:' in run.text or run.font.italic):
                            run.font.size = Pt(10)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    
    return table

def process_equations_with_word_com(doc_path):
    """Sử dụng Win32COM để chuyển placeholder thành Equation thực sự"""
    
    if not HAS_WIN32COM:
        print("⚠️ Win32COM không khả dụng. Bỏ qua việc tạo Equation.")
        return False
    
    try:
        # Khởi động Word application
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        
        # Mở document
        doc = word_app.Documents.Open(doc_path)
        
        # Tìm và thay thế tất cả placeholder
        find = word_app.Selection.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        
        # Pattern để tìm [EQUATION:...]
        equation_pattern = r'\[EQUATION:([^\]]+)\]'
        
        # Duyệt qua document để tìm placeholder
        for paragraph in doc.Paragraphs:
            para_text = paragraph.Range.Text
            matches = re.finditer(equation_pattern, para_text)
            
            for match in matches:
                latex_code = match.group(1)
                
                # Tìm vị trí trong Word
                find.Text = match.group(0)
                if find.Execute():
                    # Xóa placeholder
                    word_app.Selection.Delete()
                    
                    # Chèn equation
                    eq_range = word_app.Selection.Range
                    eq = eq_range.OMaths.Add(eq_range)
                    
                    # Chuyển LaTeX thành OMath (đơn giản hóa)
                    try:
                        eq.BuildUp()
                        # Thay thế một số ký hiệu cơ bản
                        simplified_latex = convert_basic_latex_to_omath(latex_code)
                        eq.Range.Text = simplified_latex
                        eq.BuildUp()
                    except:
                        # Nếu lỗi, để lại LaTeX gốc
                        eq.Range.Text = latex_code
        
        # Lưu và đóng
        doc.Save()
        doc.Close()
        word_app.Quit()
        
        print("✅ Đã chuyển đổi placeholder thành Equation thành công!")
        return True
        
    except Exception as e:
        print(f"❌ Lỗi khi xử lý Equation với Win32COM: {e}")
        try:
            if 'doc' in locals():
                doc.Close(False)
            if 'word_app' in locals():
                word_app.Quit()
        except:
            pass
        return False

def convert_basic_latex_to_omath(latex_code):
    """Chuyển đổi LaTeX cơ bản thành format OMath của Word"""
    
    # Mapping cơ bản
    conversions = {
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'\1/\2',  # Fraction
        r'\^(\w)': r'^\1',  # Superscript đơn
        r'\^{([^}]+)}': r'^\1',  # Superscript phức tạp
        r'_(\w)': r'_\1',  # Subscript đơn  
        r'_{([^}]+)}': r'_\1',  # Subscript phức tạp
        r'\\sqrt\{([^}]+)\}': r'√(\1)',  # Square root
        r'\\sum': '∑',  # Summation
        r'\\int': '∫',  # Integral
        r'\\infty': '∞',  # Infinity
        r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ',
        r'\\delta': 'δ', r'\\pi': 'π', r'\\theta': 'θ',
        r'\\lambda': 'λ', r'\\mu': 'μ', r'\\sigma': 'σ',
        r'\\leq': '≤', r'\\geq': '≥', r'\\neq': '≠',
        r'\\pm': '±', r'\\times': '×', r'\\div': '÷'
    }
    
    result = latex_code
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def convert_docx_with_latex_and_tables(uploaded_file):
    """Chuyển đổi file Word: LaTeX → Equation, Markdown → Table"""
    
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
    
    # Header
    header = new_doc.add_heading('DOCUMENT ĐÃ CHUYỂN ĐỔI LATEX VÀ MARKDOWN', level=1)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in header.runs:
        run.font.color.rgb = RGBColor(31, 78, 121)
        run.font.size = Pt(16)
    
    # Thông tin
    info_para = new_doc.add_paragraph(f"Ngày chuyển đổi: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in info_para.runs:
        run.font.italic = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(128, 128, 128)
    
    if HAS_WIN32COM:
        info_para = new_doc.add_paragraph("✅ Hỗ trợ Equation thực sự với Win32COM")
    else:
        info_para = new_doc.add_paragraph("⚠️ Không có Win32COM - LaTeX hiển thị dạng italic")
    
    info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in info_para.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(100, 100, 100)
    
    new_doc.add_paragraph()
    
    # Thống kê
    stats = {
        'latex_count': 0,
        'table_count': 0,
        'processed_paragraphs': 0
    }
    
    # Xử lý từng paragraph
    for para in doc.paragraphs:
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()  # Giữ khoảng trống
            continue
        
        stats['processed_paragraphs'] += 1
        
        # Kiểm tra có bảng Markdown không
        tables = detect_markdown_tables(text)
        if tables:
            stats['table_count'] += len(tables)
            
            # Xử lý text và bảng
            current_pos = 0
            for table in sorted(tables, key=lambda x: x['start']):
                # Text trước bảng
                if table['start'] > current_pos:
                    before_text = text[current_pos:table['start']].strip()
                    if before_text:
                        process_latex_in_paragraph(new_doc, before_text)
                
                # Tạo bảng
                create_word_table(new_doc, table['header'], table['rows'])
                new_doc.add_paragraph()  # Khoảng cách sau bảng
                
                current_pos = table['end']
            
            # Text sau bảng cuối
            if current_pos < len(text):
                remaining_text = text[current_pos:].strip()
                if remaining_text:
                    process_latex_in_paragraph(new_doc, remaining_text)
        else:
            # Không có bảng, chỉ xử lý LaTeX
            formulas = detect_latex_patterns(text)
            if formulas:
                stats['latex_count'] += len(formulas)
            
            process_latex_in_paragraph(new_doc, text)
    
    # Thêm thống kê cuối document
    new_doc.add_paragraph()
    stats_para = new_doc.add_paragraph("📊 THỐNG KÊ CHUYỂN ĐỔI:")
    stats_para.runs[0].font.bold = True
    stats_para.runs[0].font.size = Pt(12)
    
    new_doc.add_paragraph(f"• Đã xử lý: {stats['processed_paragraphs']} đoạn văn")
    new_doc.add_paragraph(f"• Công thức LaTeX: {stats['latex_count']} công thức")
    new_doc.add_paragraph(f"• Bảng Markdown: {stats['table_count']} bảng")
    
    return new_doc, stats

def post_process_with_word_com(doc_stream):
    """Xử lý sau với Win32COM để tạo Equation thực sự"""
    
    if not HAS_WIN32COM:
        return doc_stream, False
    
    try:
        # Lưu tạm document
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_file:
            tmp_file.write(doc_stream.getvalue())
            tmp_path = tmp_file.name
        
        # Xử lý với Win32COM
        success = process_equations_with_word_com(tmp_path)
        
        if success:
            # Đọc lại file đã xử lý
            with open(tmp_path, 'rb') as f:
                processed_stream = BytesIO(f.read())
        else:
            processed_stream = doc_stream
        
        # Xóa file tạm
        try:
            os.unlink(tmp_path)
        except:
            pass
        
        return processed_stream, success
        
    except Exception as e:
        print(f"❌ Lỗi post-processing: {e}")
        return doc_stream, False
