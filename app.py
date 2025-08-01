import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.table import Table
import re
from io import BytesIO
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

def latex_to_omml_simple(latex_expr):
    """Chuyển đổi LaTeX đơn giản thành OMML"""
    try:
        # Xử lý các trường hợp cơ bản
        latex_expr = latex_expr.strip()
        
        # Xử lý phân số \frac{a}{b}
        frac_pattern = r'\\frac\{([^}]+)\}\{([^}]+)\}'
        if re.search(frac_pattern, latex_expr):
            match = re.search(frac_pattern, latex_expr)
            num = match.group(1)
            den = match.group(2)
            omml = f'''<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                <m:oMath>
                    <m:f>
                        <m:num><m:r><m:t>{num}</m:t></m:r></m:num>
                        <m:den><m:r><m:t>{den}</m:t></m:r></m:den>
                    </m:f>
                </m:oMath>
            </m:oMathPara>'''
            return omml
        
        # Xử lý mũ x^2
        sup_pattern = r'([a-zA-Z0-9]+)\^([a-zA-Z0-9]+)'
        if re.search(sup_pattern, latex_expr):
            match = re.search(sup_pattern, latex_expr)
            base = match.group(1)
            exp = match.group(2)
            omml = f'''<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                <m:oMath>
                    <m:sSup>
                        <m:e><m:r><m:t>{base}</m:t></m:r></m:e>
                        <m:sup><m:r><m:t>{exp}</m:t></m:r></m:sup>
                    </m:sSup>
                </m:oMath>
            </m:oMathPara>'''
            return omml
        
        # Xử lý căn bậc hai \sqrt{x}
        sqrt_pattern = r'\\sqrt\{([^}]+)\}'
        if re.search(sqrt_pattern, latex_expr):
            match = re.search(sqrt_pattern, latex_expr)
            content = match.group(1)
            omml = f'''<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                <m:oMath>
                    <m:rad>
                        <m:radPr><m:degHide m:val="1"/></m:radPr>
                        <m:deg></m:deg>
                        <m:e><m:r><m:t>{content}</m:t></m:r></m:e>
                    </m:rad>
                </m:oMath>
            </m:oMathPara>'''
            return omml
        
        # Trường hợp đơn giản: chỉ là text bình thường
        omml = f'''<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:oMath>
                <m:r><m:t>{latex_expr}</m:t></m:r>
            </m:oMath>
        </m:oMathPara>'''
        return omml
        
    except Exception as e:
        return None

def insert_equation_in_paragraph_fixed(paragraph, omml_xml):
    """Chèn phương trình OMML vào đoạn văn - version mới"""
    try:
        # Parse OMML XML
        omml_element = parse_xml(omml_xml)
        
        # Chèn vào paragraph element thay vì run element
        paragraph._element.append(omml_element)
        
        return True
    except Exception as e:
        return False

def format_table_fixed(table, options):
    """Format bảng - version được cải thiện"""
    try:
        if not table or len(table.rows) == 0:
            return False
            
        # Căn giữa bảng
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header row (hàng đầu tiên)
        header_row = table.rows[0]
        for j, cell in enumerate(header_row.cells):
            # Xóa nội dung cũ và thêm mới
            for paragraph in cell.paragraphs:
                paragraph.clear()
            
            # Thêm nội dung với format mới
            p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
            run = p.add_run(cell.text if hasattr(cell, '_original_text') else '')
            run.bold = True
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Tô màu nền header
            if options.get('color_header', True):
                try:
                    # Thêm màu nền xanh
                    shading_elm = parse_xml(
                        r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="4472C4"/>'
                    )
                    cell._tc.get_or_add_tcPr().append(shading_elm)
                except:
                    pass
        
        # Format data rows
        for i in range(1, len(table.rows)):
            row = table.rows[i]
            for j, cell in enumerate(row.cells):
                # Xóa nội dung cũ và thêm mới
                for paragraph in cell.paragraphs:
                    paragraph.clear()
                
                # Thêm nội dung với format mới
                p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                run = p.add_run(cell.text if hasattr(cell, '_original_text') else '')
                run.font.size = Pt(11)
                run.font.name = 'Times New Roman'
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        return True
    except Exception as e:
        st.error(f"Lỗi format bảng: {str(e)}")
        return False

def format_questions_improved(paragraph, options):
    """Format câu hỏi trắc nghiệm - version cải thiện"""
    try:
        text = paragraph.text.strip()
        if not text:
            return False
        
        # Xóa nội dung cũ
        paragraph.clear()
        
        # Kiểm tra câu hỏi với **Câu số:**
        if '**Câu ' in text and ':**' in text:
            start_pos = text.find('**Câu ')
            end_pos = text.find(':**', start_pos) + 3
            
            if start_pos >= 0 and end_pos > start_pos:
                question_part = text[start_pos:end_pos]
                content_part = text[end_pos:].strip()
                
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
        
        # Kiểm tra câu hỏi với Câu số:
        elif text.startswith('Câu ') and ':' in text[:20]:
            colon_pos = text.find(':')
            if colon_pos > 0 and colon_pos < 20:
                question_part = text[:colon_pos + 1]
                content_part = text[colon_pos + 1:].strip()
                
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
        
        # Kiểm tra đáp án (A., B., C., D.)
        elif len(text) >= 2 and text[0] in 'ABCDabcd' and text[1] == '.':
            answer_letter = text[:2]  # A., B., etc.
            answer_content = text[2:].strip()
            
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
        
        else:
            # Không match pattern nào, khôi phục text gốc
            paragraph.add_run(text)
            return False
        
    except Exception as e:
        # Nếu có lỗi, khôi phục text gốc
        paragraph.clear()
        paragraph.add_run(text)
        return False

def copy_table_improved(original_table, new_doc, options):
    """Sao chép bảng - version cải thiện"""
    try:
        rows_count = len(original_table.rows)
        if rows_count == 0:
            return None
            
        cols_count = len(original_table.rows[0].cells)
        if cols_count == 0:
            return None
            
        # Tạo bảng mới
        new_table = new_doc.add_table(rows=rows_count, cols=cols_count)
        
        # Sao chép nội dung từng cell
        for i in range(rows_count):
            original_row = original_table.rows[i]
            new_row = new_table.rows[i]
            
            for j in range(min(cols_count, len(original_row.cells))):
                # Lấy text từ cell gốc
                cell_text = original_row.cells[j].text
                
                # Gán text cho cell mới
                new_row.cells[j].text = cell_text
                # Lưu text gốc để format sau
                new_row.cells[j]._original_text = cell_text
        
        # Format bảng nếu được yêu cầu
        if options.get('format_tables', True):
            format_table_fixed(new_table, options)
        
        return new_table
        
    except Exception as e:
        st.error(f"Lỗi copy bảng: {str(e)}")
        return None

def find_latex_expressions_improved(text):
    """Tìm biểu thức LaTeX - version cải thiện"""
    expressions = []
    
    try:
        # Tìm ${...}$ trước
        pattern = r'\$\{([^}]+)\}\$'
        for match in re.finditer(pattern, text):
            expressions.append((match.start(), match.end(), match.group(1), 'inline'))
        
        # Tìm $...$ (tránh trùng với ${...}$)
        pattern = r'\$([^{}$]+)\$'
        for match in re.finditer(pattern, text):
            # Kiểm tra không trùng với biểu thức đã tìm
            is_overlap = False
            for existing_start, existing_end, _, _ in expressions:
                if (match.start() >= existing_start and match.end() <= existing_end):
                    is_overlap = True
                    break
            
            if not is_overlap:
                expressions.append((match.start(), match.end(), match.group(1), 'display'))
    
    except Exception as e:
        pass
    
    # Sắp xếp theo vị trí
    expressions.sort(key=lambda x: x[0])
    return expressions

def clean_text_simple(text):
    """Làm sạch văn bản - version đơn giản"""
    if not text:
        return text
    
    # Loại bỏ ký tự điều khiển
    cleaned = ''.join(char if ord(char) >= 32 or char in '\n\t' else ' ' for char in text)
    
    # Chuẩn hóa khoảng trắng
    return ' '.join(cleaned.split())

def process_word_document_improved(doc, options):
    """Xử lý tài liệu Word - version cải thiện với debugging"""
    try:
        new_doc = Document()
        
        # Counters
        processed_count = 0
        error_count = 0
        cleaned_paragraphs = 0
        formatted_questions = 0
        formatted_tables = 0
        
        # Debug info
        debug_info = []
        
        # Xử lý paragraphs
        for para_idx, para in enumerate(doc.paragraphs):
            try:
                text = para.text
                debug_info.append(f"Paragraph {para_idx}: {text[:50]}...")
                
                new_para = new_doc.add_paragraph()
                
                # Sao chép định dạng đoạn văn
                try:
                    if para.alignment:
                        new_para.alignment = para.alignment
                except:
                    pass
                
                # Làm sạch văn bản nếu được yêu cầu
                if options.get('clean_text', False):
                    original_text = text
                    text = clean_text_simple(text)
                    if text != original_text:
                        cleaned_paragraphs += 1
                
                # Format câu hỏi trắc nghiệm nếu được yêu cầu
                if options.get('format_questions', False):
                    if format_questions_improved(new_para, options):
                        formatted_questions += 1
                        debug_info.append(f"  -> Formatted as question")
                        continue
                
                # Tìm và xử lý biểu thức LaTeX
                if options.get('convert_equations', True):
                    latex_expressions = find_latex_expressions_improved(text)
                    debug_info.append(f"  -> Found {len(latex_expressions)} LaTeX expressions")
                    
                    if latex_expressions:
                        last_idx = 0
                        
                        for start, end, latex_expr, math_type in latex_expressions:
                            # Thêm văn bản trước biểu thức toán học
                            if start > last_idx:
                                text_before = text[last_idx:start]
                                if text_before:
                                    new_para.add_run(text_before)
                            
                            # Chuyển đổi LaTeX thành OMML và chèn
                            omml_xml = latex_to_omml_simple(latex_expr.strip())
                            debug_info.append(f"    -> Converting: {latex_expr}")
                            
                            if omml_xml:
                                if insert_equation_in_paragraph_fixed(new_para, omml_xml):
                                    processed_count += 1
                                    debug_info.append(f"    -> SUCCESS: {latex_expr}")
                                else:
                                    # Dự phòng: chèn dưới dạng văn bản được định dạng
                                    run = new_para.add_run(f"[EQUATION: {latex_expr}]")
                                    run.italic = True
                                    run.bold = True
                                    error_count += 1
                                    debug_info.append(f"    -> FALLBACK: {latex_expr}")
                            else:
                                # Lỗi trong chuyển đổi: chèn dưới dạng văn bản
                                if options.get('show_errors', True):
                                    run = new_para.add_run(f"[ERROR: {latex_expr}]")
                                    run.font.color.rgb = RGBColor(255, 0, 0)  # Màu đỏ
                                else:
                                    run = new_para.add_run(f"${latex_expr}$")
                                error_count += 1
                                debug_info.append(f"    -> ERROR: {latex_expr}")
                            
                            last_idx = end
                        
                        # Thêm văn bản còn lại sau biểu thức toán học cuối cùng
                        if last_idx < len(text):
                            remaining_text = text[last_idx:]
                            if remaining_text:
                                new_para.add_run(remaining_text)
                    else:
                        # Không có biểu thức LaTeX, sao chép đoạn văn như cũ
                        new_para.add_run(text)
                else:
                    # Không chuyển đổi LaTeX, sao chép như cũ
                    new_para.add_run(text)
                
                # Áp dụng định dạng cơ bản
                if options.get('fix_formatting', False):
                    for run in new_para.runs:
                        try:
                            run.font.name = 'Times New Roman'
                            if not run.font.size or run.font.size < Pt(11):
                                run.font.size = Pt(12)
                        except:
                            pass
                            
            except Exception as e:
                # Nếu có lỗi với paragraph, tạo một paragraph trống
                new_doc.add_paragraph("")
                debug_info.append(f"  -> ERROR in paragraph: {str(e)}")
                continue
        
        # Xử lý bảng
        debug_info.append(f"Processing {len(doc.tables)} tables...")
        for table_idx, table in enumerate(doc.tables):
            try:
                debug_info.append(f"Table {table_idx}: {len(table.rows)} rows, {len(table.rows[0].cells) if table.rows else 0} cols")
                new_table = copy_table_improved(table, new_doc, options)
                if new_table:
                    formatted_tables += 1
                    debug_info.append(f"  -> Table {table_idx} formatted successfully")
                else:
                    debug_info.append(f"  -> Table {table_idx} failed to format")
            except Exception as e:
                debug_info.append(f"  -> ERROR in table {table_idx}: {str(e)}")
                continue
        
        return new_doc, {
            'processed_count': processed_count,
            'error_count': error_count,
            'cleaned_paragraphs': cleaned_paragraphs,
            'formatted_questions': formatted_questions,
            'formatted_tables': formatted_tables,
            'total_paragraphs': len(doc.paragraphs),
            'total_tables': len(doc.tables),
            'debug_info': debug_info
        }
        
    except Exception as e:
        # Nếu có lỗi nghiêm trọng, tạo document mới trống
        new_doc = Document()
        new_doc.add_paragraph("Lỗi khi xử lý tài liệu: " + str(e))
        return new_doc, {
            'processed_count': 0,
            'error_count': 1,
            'cleaned_paragraphs': 0,
            'formatted_questions': 0,
            'formatted_tables': 0,
            'total_paragraphs': 0,
            'total_tables': 0,
            'debug_info': [f"CRITICAL ERROR: {str(e)}"]
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
                                       help="Chuyển $...$, ${...}$ thành công thức toán chuẩn")
        
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
        
        # Debug option
        show_debug = st.checkbox("Hiển thị thông tin debug", value=False,
                                help="Xem chi tiết quá trình xử lý")
        
        st.markdown("---")
        st.markdown("### 📖 Hướng dẫn sử dụng")
        st.markdown("""
        1. **Tải file**: Chọn file Word (.docx) 
        2. **Cấu hình**: Điều chỉnh các tùy chọn bên trái
        3. **Xử lý**: Nhấn nút "Xử lý file"
        4. **Tải về**: Download file đã xử lý
        
        **Định dạng được hỗ trợ:**
        - LaTeX: `$x^2$`, `${a+b}$`, `$\\frac{a}{b}$`
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
                        for i, table in enumerate(doc_preview.tables):
                            st.text(f"  Bảng {i+1}: {len(table.rows)} dòng × {len(table.rows[0].cells) if table.rows else 0} cột")
                except Exception as e:
                    st.warning(f"Không thể preview file: {str(e)}")
            
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
                        new_doc, stats = process_word_document_improved(doc, options)
                        
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
                        
                        # Debug info
                        if show_debug and stats.get('debug_info'):
                            with st.expander("🔍 Thông tin Debug"):
                                for info in stats['debug_info'][:20]:  # Giới hạn 20 dòng đầu
                                    st.text(info)
                        
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
        
        with st.expander("🔢 LaTeX được hỗ trợ"):
            st.markdown("""
            **Cơ bản:**
            - `$x^2$` → x²
            - `${a+b}$` → a+b
            - `$x + y = z$` → x + y = z
            
            **Phân số:**
            - `$\\frac{a}{b}$` → a/b (dạng phân số)
            
            **Căn bậc hai:**
            - `$\\sqrt{x}$` → √x
            """)
        
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
        
        with st.expander("🎯 Test LaTeX"):
            st.markdown("""
            **Thử nghiệm với text này:**
            ```
            Phương trình bậc hai: $ax^2 + bx + c = 0$
            
            Công thức nghiệm: $x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}$
            
            Tích phân: ${\\int_0^1 x dx = \\frac{1}{2}}$
            ```
            """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>🚀 Được phát triển với Streamlit | 📊 Hỗ trợ LaTeX, Trắc nghiệm & Bảng</p>
        <p><small>Phiên bản 5.0 - Fixed & Improved: LaTeX conversion + Table formatting + Debug mode</small></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
