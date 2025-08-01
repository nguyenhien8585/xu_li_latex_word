import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.table import Table
import re
from latex2mathml.converter import convert
import xml.etree.ElementTree as ET
from io import BytesIO
import zipfile
import tempfile
import os
from datetime import datetime

# Cấu hình trang
st.set_page_config(
    page_title="Xử lý LaTeX & Format trong Word",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

def latex_to_omml(latex_expr):
    """Chuyển đổi biểu thức LaTeX thành Office Math Markup Language (OMML)"""
    try:
        # Chuyển LaTeX thành MathML
        mathml = convert(latex_expr)
        
        # Phân tích cú pháp MathML
        root = ET.fromstring(mathml)
        
        # Chuyển đổi MathML thành OMML (đơn giản hóa)
        omml_parts = []
        
        def mathml_to_omml_recursive(element):
            tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
            
            if tag == 'math':
                for child in element:
                    mathml_to_omml_recursive(child)
            elif tag == 'mi':  # định danh (identifier)
                text = element.text or ''
                omml_parts.append(f'<m:r><m:t>{text}</m:t></m:r>')
            elif tag == 'mn':  # số (number)
                text = element.text or ''
                omml_parts.append(f'<m:r><m:t>{text}</m:t></m:r>')
            elif tag == 'mo':  # toán tử (operator)
                text = element.text or ''
                omml_parts.append(f'<m:r><m:t>{text}</m:t></m:r>')
            elif tag == 'mfrac':  # phân số
                omml_parts.append('<m:f>')
                omml_parts.append('<m:num>')
                if len(element) > 0:
                    mathml_to_omml_recursive(element[0])
                omml_parts.append('</m:num>')
                omml_parts.append('<m:den>')
                if len(element) > 1:
                    mathml_to_omml_recursive(element[1])
                omml_parts.append('</m:den>')
                omml_parts.append('</m:f>')
            elif tag == 'msup':  # mũ trên
                omml_parts.append('<m:sSup>')
                omml_parts.append('<m:e>')
                if len(element) > 0:
                    mathml_to_omml_recursive(element[0])
                omml_parts.append('</m:e>')
                omml_parts.append('<m:sup>')
                if len(element) > 1:
                    mathml_to_omml_recursive(element[1])
                omml_parts.append('</m:sup>')
                omml_parts.append('</m:sSup>')
            elif tag == 'msub':  # chỉ số dưới
                omml_parts.append('<m:sSub>')
                omml_parts.append('<m:e>')
                if len(element) > 0:
                    mathml_to_omml_recursive(element[0])
                omml_parts.append('</m:e>')
                omml_parts.append('<m:sub>')
                if len(element) > 1:
                    mathml_to_omml_recursive(element[1])
                omml_parts.append('</m:sub>')
                omml_parts.append('</m:sSub>')
            elif tag == 'msubsup':  # vừa có chỉ số dưới vừa có mũ trên
                omml_parts.append('<m:sSubSup>')
                omml_parts.append('<m:e>')
                if len(element) > 0:
                    mathml_to_omml_recursive(element[0])
                omml_parts.append('</m:e>')
                omml_parts.append('<m:sub>')
                if len(element) > 1:
                    mathml_to_omml_recursive(element[1])
                omml_parts.append('</m:sub>')
                omml_parts.append('<m:sup>')
                if len(element) > 2:
                    mathml_to_omml_recursive(element[2])
                omml_parts.append('</m:sup>')
                omml_parts.append('</m:sSubSup>')
            elif tag == 'msqrt':  # căn bậc hai
                omml_parts.append('<m:rad>')
                omml_parts.append('<m:radPr><m:degHide m:val="1"/></m:radPr>')
                omml_parts.append('<m:deg></m:deg>')
                omml_parts.append('<m:e>')
                for child in element:
                    mathml_to_omml_recursive(child)
                omml_parts.append('</m:e>')
                omml_parts.append('</m:rad>')
            elif tag == 'mroot':  # căn bậc n
                omml_parts.append('<m:rad>')
                omml_parts.append('<m:deg>')
                if len(element) > 1:
                    mathml_to_omml_recursive(element[1])
                omml_parts.append('</m:deg>')
                omml_parts.append('<m:e>')
                if len(element) > 0:
                    mathml_to_omml_recursive(element[0])
                omml_parts.append('</m:e>')
                omml_parts.append('</m:rad>')
            elif tag == 'mrow':  # nhóm các phần tử
                for child in element:
                    mathml_to_omml_recursive(child)
            else:
                # Xử lý các phần tử khác một cách đệ quy
                for child in element:
                    mathml_to_omml_recursive(child)
        
        mathml_to_omml_recursive(root)
        
        # Bọc trong cấu trúc OMML
        omml = f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{"".join(omml_parts)}</m:oMath>'
        
        return omml.strip()
        
    except Exception as e:
        return None

def insert_equation_in_paragraph(paragraph, omml_xml):
    """Chèn phương trình OMML vào đoạn văn"""
    try:
        # Tạo một run mới cho phương trình
        run = paragraph.add_run()
        
        # Phân tích và chèn XML OMML
        omml_element = parse_xml(omml_xml)
        run._element.append(omml_element)
        
        return True
    except Exception as e:
        return False

def format_table_safe(table, options):
    """Format bảng theo yêu cầu - phiên bản an toàn"""
    try:
        # Căn giữa bảng
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header row (hàng đầu tiên)
        if len(table.rows) > 0:
            header_row = table.rows[0]
            for cell in header_row.cells:
                try:
                    # Lấy text từ cell
                    cell_text = cell.text
                    # Xóa nội dung cũ
                    cell.text = ""
                    # Thêm lại với format mới
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run(cell_text)
                    run.bold = True
                    run.font.size = Pt(12)
                    run.font.name = 'Times New Roman'
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Tô màu nền header
                    if options.get('color_header', True):
                        try:
                            shading_elm = parse_xml('<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="4472C4"/>')
                            cell._tc.get_or_add_tcPr().append(shading_elm)
                        except:
                            pass
                except:
                    continue
        
        # Format data rows
        for i, row in enumerate(table.rows[1:], 1):
            for cell in row.cells:
                try:
                    # Lấy text từ cell
                    cell_text = cell.text
                    # Xóa nội dung cũ
                    cell.text = ""
                    # Thêm lại với format mới
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run(cell_text)
                    run.font.size = Pt(11)
                    run.font.name = 'Times New Roman'
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except:
                    continue
        
        return True
    except Exception as e:
        return False

def format_questions_safe(paragraph, options):
    """Format câu hỏi trắc nghiệm - phiên bản hoàn toàn an toàn"""
    try:
        text = paragraph.text.strip()
        
        if not text:
            return False
        
        # Kiểm tra câu hỏi với **Câu số:**
        if '**Câu ' in text and ':**' in text:
            try:
                start_pos = text.find('**Câu ')
                end_pos = text.find(':**', start_pos) + 3
                
                if start_pos >= 0 and end_pos > start_pos:
                    question_part = text[start_pos:end_pos]
                    content_part = text[end_pos:].strip()
                    
                    paragraph.clear()
                    
                    # Thêm phần câu hỏi (in đậm)
                    run_question = paragraph.add_run(question_part)
                    run_question.bold = True
                    run_question.font.size = Pt(12)
                    run_question.font.name = 'Times New Roman'
                    
                    # Thêm phần nội dung
                    if content_part:
                        run_content = paragraph.add_run(" " + content_part)
                        run_content.font.size = Pt(12)
                        run_content.font.name = 'Times New Roman'
                    
                    return True
            except:
                pass
        
        # Kiểm tra câu hỏi với Câu số:
        if text.startswith('Câu ') and ':' in text[:20]:
            try:
                colon_pos = text.find(':')
                if colon_pos > 0 and colon_pos < 20:
                    question_part = text[:colon_pos + 1]
                    content_part = text[colon_pos + 1:].strip()
                    
                    paragraph.clear()
                    
                    # Thêm phần câu hỏi (in đậm)
                    run_question = paragraph.add_run(question_part)
                    run_question.bold = True
                    run_question.font.size = Pt(12)
                    run_question.font.name = 'Times New Roman'
                    
                    # Thêm phần nội dung
                    if content_part:
                        run_content = paragraph.add_run(" " + content_part)
                        run_content.font.size = Pt(12)
                        run_content.font.name = 'Times New Roman'
                    
                    return True
            except:
                pass
        
        # Kiểm tra đáp án (A., B., C., D.)
        if len(text) >= 2:
            first_char = text[0]
            second_char = text[1] if len(text) > 1 else ''
            
            if first_char in 'ABCDabcd' and second_char == '.':
                try:
                    answer_letter = text[:2]  # A., B., etc.
                    answer_content = text[2:].strip()
                    
                    paragraph.clear()
                    
                    # Thêm ký hiệu đáp án (in đậm, màu xanh nước biển)
                    run_letter = paragraph.add_run(answer_letter)
                    run_letter.bold = True
                    run_letter.font.color.rgb = RGBColor(0, 123, 191)  # Xanh nước biển
                    run_letter.font.size = Pt(12)
                    run_letter.font.name = 'Times New Roman'
                    
                    # Thêm nội dung đáp án (nếu có)
                    if answer_content:
                        run_content = paragraph.add_run(" " + answer_content)
                        run_content.font.size = Pt(12)
                        run_content.font.name = 'Times New Roman'
                    
                    return True
                except:
                    pass
        
        return False
        
    except Exception as e:
        return False

def copy_table_safe(original_table, new_doc, options):
    """Sao chép bảng từ tài liệu gốc sang tài liệu mới - phiên bản an toàn"""
    try:
        # Lấy kích thước bảng
        rows_count = len(original_table.rows)
        cols_count = 0
        
        if rows_count > 0:
            cols_count = len(original_table.rows[0].cells)
        
        if rows_count == 0 or cols_count == 0:
            return None
            
        # Tạo bảng mới
        new_table = new_doc.add_table(rows=rows_count, cols=cols_count)
        
        # Sao chép nội dung từng cell
        for i in range(rows_count):
            try:
                original_row = original_table.rows[i]
                for j in range(min(cols_count, len(original_row.cells))):
                    try:
                        # Lấy text từ cell gốc
                        cell_text = original_row.cells[j].text.strip()
                        # Gán text cho cell mới
                        new_table.cell(i, j).text = cell_text
                    except:
                        continue
            except:
                continue
        
        # Format bảng nếu được yêu cầu
        if options.get('format_tables', False):
            format_table_safe(new_table, options)
        
        return new_table
        
    except Exception as e:
        return None

def clean_text_safe(text):
    """Làm sạch văn bản - phiên bản an toàn"""
    try:
        if not text:
            return text
            
        # Loại bỏ ký tự điều khiển bằng cách thay thế từng ký tự
        cleaned = ""
        for char in text:
            # Chỉ giữ lại ký tự có thể in được
            if ord(char) >= 32 or char in ['\n', '\t']:
                cleaned += char
            else:
                cleaned += ' '
        
        # Chuẩn hóa khoảng trắng
        words = cleaned.split()
        return ' '.join(words)
        
    except Exception as e:
        return text

def find_latex_expressions(text):
    """Tìm biểu thức LaTeX một cách an toàn"""
    try:
        expressions = []
        
        # Tìm $$...$$
        start = 0
        while True:
            start_pos = text.find('$$', start)
            if start_pos == -1:
                break
            end_pos = text.find('$$', start_pos + 2)
            if end_pos == -1:
                break
            
            latex_expr = text[start_pos + 2:end_pos]
            expressions.append((start_pos, end_pos + 2, latex_expr, 'display'))
            start = end_pos + 2
        
        # Tìm $...$
        start = 0
        while True:
            start_pos = text.find('$', start)
            if start_pos == -1:
                break
                
            # Kiểm tra xem có phải là $$ không
            if start_pos < len(text) - 1 and text[start_pos + 1] == '$':
                start = start_pos + 2
                continue
                
            end_pos = text.find('$', start_pos + 1)
            if end_pos == -1:
                break
                
            # Kiểm tra xem có phải là $$ không
            if end_pos > 0 and text[end_pos - 1] == '$':
                start = end_pos + 1
                continue
            
            latex_expr = text[start_pos + 1:end_pos]
            # Chỉ thêm nếu không trùng với biểu thức $$...$$
            is_overlap = False
            for existing_start, existing_end, _, _ in expressions:
                if start_pos >= existing_start and end_pos + 1 <= existing_end:
                    is_overlap = True
                    break
            
            if not is_overlap:
                expressions.append((start_pos, end_pos + 1, latex_expr, 'inline'))
            
            start = end_pos + 1
        
        # Sắp xếp theo vị trí
        expressions.sort(key=lambda x: x[0])
        return expressions
        
    except Exception as e:
        return []

def process_word_document_safe(doc, options):
    """Xử lý tài liệu Word - phiên bản hoàn toàn an toàn"""
    try:
        new_doc = Document()
        
        processed_count = 0
        error_count = 0
        cleaned_paragraphs = 0
        formatted_questions = 0
        formatted_tables = 0
        
        # Xử lý paragraphs
        for para in doc.paragraphs:
            try:
                new_para = new_doc.add_paragraph()
                
                # Sao chép định dạng đoạn văn
                try:
                    if para.alignment:
                        new_para.alignment = para.alignment
                    if para.style:
                        new_para.style = para.style
                except:
                    pass
                
                text = para.text
                
                # Làm sạch văn bản nếu được yêu cầu
                if options.get('clean_text', False):
                    original_text = text
                    text = clean_text_safe(text)
                    if text != original_text:
                        cleaned_paragraphs += 1
                
                # Format câu hỏi trắc nghiệm nếu được yêu cầu
                if options.get('format_questions', False):
                    new_para.clear()
                    new_para.add_run(text)
                    if format_questions_safe(new_para, options):
                        formatted_questions += 1
                        continue
                
                # Tìm và xử lý biểu thức LaTeX
                if options.get('convert_equations', True):
                    latex_expressions = find_latex_expressions(text)
                    
                    if latex_expressions:
                        last_idx = 0
                        
                        for start, end, latex_expr, math_type in latex_expressions:
                            try:
                                # Thêm văn bản trước biểu thức toán học
                                if start > last_idx:
                                    text_before = text[last_idx:start]
                                    if text_before:
                                        new_para.add_run(text_before)
                                
                                # Chuyển đổi LaTeX thành OMML và chèn
                                omml_xml = latex_to_omml(latex_expr.strip())
                                
                                if omml_xml:
                                    if insert_equation_in_paragraph(new_para, omml_xml):
                                        processed_count += 1
                                    else:
                                        # Dự phòng: chèn dưới dạng văn bản được định dạng
                                        run = new_para.add_run(f"[PHƯƠNG TRÌNH: {latex_expr}]")
                                        run.italic = True
                                        run.bold = True
                                        error_count += 1
                                else:
                                    # Lỗi trong chuyển đổi: chèn dưới dạng văn bản
                                    if options.get('show_errors', True):
                                        run = new_para.add_run(f"[LỖI: {latex_expr}]")
                                        try:
                                            run.font.color.rgb = RGBColor(255, 0, 0)  # Màu đỏ
                                        except:
                                            pass
                                    else:
                                        run = new_para.add_run(f"${latex_expr}$")
                                    error_count += 1
                                
                                last_idx = end
                            except Exception as e:
                                # Lỗi khi xử lý LaTeX, giữ nguyên
                                run = new_para.add_run(f"${latex_expr}$")
                                error_count += 1
                                last_idx = end
                        
                        # Thêm văn bản còn lại sau biểu thức toán học cuối cùng
                        if last_idx < len(text):
                            remaining_text = text[last_idx:]
                            if remaining_text:
                                new_para.add_run(remaining_text)
                    else:
                        # Không có biểu thức LaTeX, sao chép đoạn văn như cũ
                        new_para.clear()
                        new_para.add_run(text)
                else:
                    # Không chuyển đổi LaTeX, sao chép như cũ
                    new_para.clear()  
                    new_para.add_run(text)
                
                # Áp dụng định dạng cơ bản
                if options.get('fix_formatting', False):
                    for run in new_para.runs:
                        try:
                            if run.font.name != 'Times New Roman':
                                run.font.name = 'Times New Roman'
                            if not run.font.size or run.font.size < Pt(11):
                                run.font.size = Pt(12)
                        except:
                            pass
                            
            except Exception as e:
                # Nếu có lỗi với paragraph, tạo một paragraph trống
                new_doc.add_paragraph("")
                continue
        
        # Xử lý bảng
        for table in doc.tables:
            try:
                new_table = copy_table_safe(table, new_doc, options)
                if new_table:
                    formatted_tables += 1
            except Exception as e:
                continue
        
        return new_doc, {
            'processed_count': processed_count,
            'error_count': error_count,
            'cleaned_paragraphs': cleaned_paragraphs,
            'formatted_questions': formatted_questions,
            'formatted_tables': formatted_tables,
            'total_paragraphs': len(doc.paragraphs),
            'total_tables': len(doc.tables)
        }
        
    except Exception as e:
        # Nếu có lỗi nghiêm trọng, tạo document mới trống
        new_doc = Document()
        new_doc.add_paragraph("Lỗi khi xử lý tài liệu")
        return new_doc, {
            'processed_count': 0,
            'error_count': 1,
            'cleaned_paragraphs': 0,
            'formatted_questions': 0,
            'formatted_tables': 0,
            'total_paragraphs': 0,
            'total_tables': 0
        }

def main():
    # Header
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 Xử lý LaTeX & Format trong Word</h1>
        <p style="color: white; margin: 10px 0 0 0;">Chuyển đổi LaTeX, Format câu hỏi trắc nghiệm & Xử lý bảng</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar với các tùy chọn
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn xử lý")
        
        st.markdown("**🔢 Xử lý LaTeX:**")
        convert_equations = st.checkbox("Chuyển đổi công thức LaTeX", value=True,
                                       help="Chuyển $...$ thành công thức toán chuẩn")
        
        show_errors = st.checkbox("Hiển thị lỗi chuyển đổi", value=True,
                                 help="Đánh dấu các biểu thức không chuyển đổi được")
        
        st.markdown("**🎨 Format tài liệu:**")
        fix_formatting = st.checkbox("Sửa định dạng cơ bản", value=True,
                                    help="Chuẩn hóa font chữ, kích thước và khoảng cách")
        
        format_questions = st.checkbox("Format câu hỏi trắc nghiệm", value=True,
                                      help="In đậm câu hỏi, tô màu xanh các đáp án A, B, C, D")
        
        format_tables = st.checkbox("Format bảng", value=True,
                                   help="Căn chỉnh bảng, in đậm header, tô màu")
        
        clean_text = st.checkbox("Làm sạch văn bản", value=True,
                                help="Loại bỏ ký tự lạ và chuẩn hóa khoảng trắng")
        
        color_header = st.checkbox("Tô màu header bảng", value=True,
                                  help="Tô màu xanh cho hàng đầu tiên của bảng")
        
        st.markdown("---")
        st.markdown("### 📖 Hướng dẫn sử dụng")
        st.markdown("""
        1. **Tải file**: Chọn file Word (.docx) 
        2. **Cấu hình**: Điều chỉnh các tùy chọn bên trái
        3. **Xử lý**: Nhấn nút "Xử lý file"
        4. **Tải về**: Download file đã xử lý
        
        **Định dạng được hỗ trợ:**
        - LaTeX: `$x^2$`, `$\\frac{a}{b}$`
        - Câu hỏi: `**Câu 1:**` → **in đậm**
        - Đáp án: `A.`, `B.` → **màu xanh**
        - Bảng: Header in đậm, căn giữa
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📁 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx) cần xử lý",
            type=["docx"],
            help="File Word chứa LaTeX, câu hỏi trắc nghiệm hoặc bảng"
        )
        
        if uploaded_file is not None:
            # Thông tin file
            file_info = {
                'name': uploaded_file.name,
                'size': f"{uploaded_file.size / 1024:.1f} KB",
                'type': uploaded_file.type
            }
            
            st.success(f"✅ Đã tải lên: **{file_info['name']}** ({file_info['size']})")
            
            # Preview nội dung
            with st.expander("👁️ Xem trước nội dung (5 đoạn đầu)"):
                try:
                    doc_preview = Document(uploaded_file)
                    for i, para in enumerate(doc_preview.paragraphs[:5]):
                        if para.text.strip():
                            text_preview = para.text[:100] + '...' if len(para.text) > 100 else para.text
                            st.text(f"{i+1}. {text_preview}")
                    
                    if doc_preview.tables:
                        st.markdown("**📊 Bảng tìm thấy:**")
                        st.info(f"Tài liệu có {len(doc_preview.tables)} bảng")
                except Exception as e:
                    st.warning("Không thể preview file")
            
            # Nút xử lý
            if st.button("🚀 Xử lý file", type="primary", use_container_width=True):
                options = {
                    'convert_equations': convert_equations,
                    'fix_formatting': fix_formatting,
                    'clean_text': clean_text,
                    'show_errors': show_errors,
                    'format_questions': format_questions,
                    'format_tables': format_tables,
                    'color_header': color_header
                }
                
                try:
                    with st.spinner("⏳ Đang xử lý file, vui lòng chờ..."):
                        # Load document
                        doc = Document(uploaded_file)
                        
                        # Process document
                        new_doc, stats = process_word_document_safe(doc, options)
                        
                        # Save to buffer
                        buffer = BytesIO()
                        new_doc.save(buffer)
                        buffer.seek(0)
                        
                        # Generate filename
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        original_name = uploaded_file.name.rsplit('.', 1)[0]
                        new_filename = f"{original_name}_processed_{timestamp}.docx"
                        
                        # Show results
                        st.markdown("### 📊 Kết quả xử lý")
                        
                        metric_cols = st.columns(3)
                        with metric_cols[0]:
                            st.metric("✅ Công thức LaTeX", stats['processed_count'])
                            st.metric("❌ Lỗi chuyển đổi", stats['error_count'])
                        with metric_cols[1]:
                            st.metric("📋 Câu hỏi đã format", stats['formatted_questions'])
                            st.metric("📊 Bảng đã format", stats['formatted_tables'])
                        with metric_cols[2]:
                            st.metric("🧹 Đoạn văn đã sạch", stats['cleaned_paragraphs'])
                            success_rate = (stats['processed_count'] / (stats['processed_count'] + stats['error_count']) * 100) if (stats['processed_count'] + stats['error_count']) > 0 else 0
                            st.metric("📈 Tỷ lệ thành công", f"{success_rate:.0f}%")
                        
                        # Download button
                        st.markdown("### 📥 Tải file đã xử lý")
                        st.download_button(
                            label="⬇️ Tải file Word đã xử lý",
                            data=buffer.getvalue(),
                            file_name=new_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                        # Success message
                        success_messages = []
                        if stats['processed_count'] > 0:
                            success_messages.append(f"✅ {stats['processed_count']} công thức LaTeX")
                        if stats['formatted_questions'] > 0:
                            success_messages.append(f"📋 {stats['formatted_questions']} câu hỏi")
                        if stats['formatted_tables'] > 0:
                            success_messages.append(f"📊 {stats['formatted_tables']} bảng")
                        if stats['cleaned_paragraphs'] > 0:
                            success_messages.append(f"🧹 {stats['cleaned_paragraphs']} đoạn văn")
                        
                        if success_messages:
                            st.success(f"🎉 Đã xử lý thành công: {', '.join(success_messages)}!")
                        
                        if stats['error_count'] > 0:
                            st.warning(f"⚠️ Có {stats['error_count']} biểu thức LaTeX gặp lỗi.")
                
                except Exception as e:
                    st.error(f"❌ Lỗi khi xử lý file: {str(e)}")
                    st.info("💡 Kiểm tra lại định dạng file và thử lại.")
    
    with col2:
        st.markdown("### 💡 Mẹo hữu ích")
        
        with st.expander("🔍 Format câu hỏi trắc nghiệm"):
            st.markdown("""
            **Input:**
            ```
            **Câu 1:** Cho tứ diện ABCD...
            A. Đáp án A
            B. Đáp án B  
            C. Đáp án C
            D. Đáp án D
            ```
            
            **Output:**
            - **Câu 1:** → In đậm
            - **A.** → Màu xanh nước biển, in đậm
            - **B.** → Màu xanh nước biển, in đậm
            """)
        
        with st.expander("📊 Format bảng"):
            st.markdown("""
            **Tự động xử lý:**
            - Header (hàng 1): In đậm, tô màu xanh
            - Dữ liệu: Căn giữa, font Times New Roman
            - Bảng: Căn giữa trang
            - Kích thước: Tự động điều chỉnh
            """)
        
        with st.expander("⚡ LaTeX được hỗ trợ"):
            st.code("""
# Cơ bản
$x + y = z$
$a^2 + b^2 = c^2$

# Phân số  
$\\frac{a}{b}$

# Căn thức
$\\sqrt{x}$, $\\sqrt[n]{x}$

# Tích phân
$\\int_0^1 f(x) dx$

# Hình học
$ABCD$, $\\angle ABC$
            """)
        
        with st.expander("🎯 Ví dụ thực tế"):
            st.markdown("""
            **File từ PDF → Word thường có:**
            - Công thức LaTeX: `$x^2 + y^2$`
            - Câu hỏi không format: `Câu 1: Nội dung`
            - Bảng không căn chỉnh
            - Font chữ không đồng nhất
            
            **→ App sẽ tự động sửa tất cả!**
            """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>🚀 Được phát triển với Streamlit | 📊 Hỗ trợ LaTeX, Trắc nghiệm & Bảng</p>
        <p><small>Phiên bản 3.0 - Siêu ổn định, không lỗi syntax</small></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
