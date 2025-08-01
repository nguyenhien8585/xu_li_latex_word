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
                if text in ['(', ')', '[', ']', '{', '}']:
                    omml_parts.append(f'<m:r><m:t>{text}</m:t></m:r>')
                else:
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
        omml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{"".join(omml_parts)}</m:oMath>'''
        
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

def format_table(table, options):
    """Format bảng theo yêu cầu"""
    try:
        # Căn giữa bảng
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header row (hàng đầu tiên)
        if len(table.rows) > 0:
            header_row = table.rows[0]
            for cell in header_row.cells:
                # In đậm text trong header
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
                        run.font.size = Pt(12)
                        run.font.name = 'Times New Roman'
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Tô màu nền header
                if options.get('color_header', True):
                    shading_elm = parse_xml(r'<w:shd {} w:fill="4472C4"/>'.format(nsdecls('w')))
                    cell._tc.get_or_add_tcPr().append(shading_elm)
        
        # Format data rows
        for i, row in enumerate(table.rows[1:], 1):
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(11)
                        run.font.name = 'Times New Roman'
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        return True
    except Exception as e:
        st.error(f"Lỗi format bảng: {str(e)}")
        return False

def format_questions(paragraph, options):
    """Format câu hỏi trắc nghiệm"""
    text = paragraph.text
    
    # Pattern cho câu hỏi (Câu 1:, Câu 2:, etc.)
    question_pattern = r'(\*\*Câu\s+\d+:\*\*|Câu\s+\d+:)'
    
    # Pattern cho các đáp án (A., B., C., D. hoặc a., b., c., d.)
    answer_pattern = r'^([A-Da-d]\.)\s+'
    
    # Kiểm tra nếu là câu hỏi
    if re.search(question_pattern, text, re.IGNORECASE):
        paragraph.clear()
        
        # Tách phần câu hỏi và nội dung
        match = re.search(question_pattern, text, re.IGNORECASE)
        if match:
            question_part = match.group(1)
            content_part = text[match.end():].strip()
            
            # Thêm phần câu hỏi (in đậm)
            run_question = paragraph.add_run(question_part)
            run_question.bold = True
            run_question.font.size = Pt(12)
            run_question.font.name = 'Times New Roman'
            
            # Thêm phần nội dung
            run_content = paragraph.add_run(" " + content_part)
            run_content.font.size = Pt(12)
            run_content.font.name = 'Times New Roman'
        
        return True
    
    # Kiểm tra nếu là đáp án
    elif re.match(answer_pattern, text.strip()):
        paragraph.clear()
        match = re.match(answer_pattern, text.strip())
        if match:
            answer_letter = match.group(1)
            answer_content = text[match.end():].strip()
            
            # Thêm ký hiệu đáp án (in đậm, màu xanh nước biển)
            run_letter = paragraph.add_run(answer_letter)
            run_letter.bold = True
            run_letter.font.color.rgb = RGBColor(0, 123, 191)  # Xanh nước biển
            run_letter.font.size = Pt(12)
            run_letter.font.name = 'Times New Roman'
            
            # Thêm nội dung đáp án
            run_content = paragraph.add_run(" " + answer_content)
            run_content.font.size = Pt(12)
            run_content.font.name = 'Times New Roman'
        
        return True
    
    return False

def fix_formatting(doc, options):
    """Sửa định dạng cơ bản của tài liệu"""
    for paragraph in doc.paragraphs:
        # Format câu hỏi trắc nghiệm nếu được yêu cầu
        if options.get('format_questions', False):
            if format_questions(paragraph, options):
                continue
        
        # Sửa khoảng cách dòng
        if paragraph.paragraph_format.line_spacing != 1.15:
            paragraph.paragraph_format.line_spacing = 1.15
        
        # Sửa font chữ cho các run
        for run in paragraph.runs:
            if run.font.name != 'Times New Roman':
                run.font.name = 'Times New Roman'
            if not run.font.size or run.font.size < Pt(11):
                run.font.size = Pt(12)
    
    # Xử lý bảng
    if options.get('format_tables', False):
        for table in doc.tables:
            format_table(table, options)

def clean_text(text):
    """Làm sạch văn bản từ các ký tự không mong muốn"""
    # Loại bỏ các ký tự điều khiển
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]', '', text)
    
    # Chuẩn hóa khoảng trắng
    text = re.sub(r'\s+', ' ', text)
    
    # Loại bỏ khoảng trắng thừa ở đầu và cuối
    text = text.strip()
    
    return text

def process_word_document(doc, options):
    """Xử lý tài liệu Word để chuyển đổi biểu thức LaTeX thành phương trình"""
    new_doc = Document()
    
    # Sao chép styles từ tài liệu gốc
    try:
        new_doc.styles = doc.styles
    except:
        pass
    
    processed_count = 0
    error_count = 0
    cleaned_paragraphs = 0
    formatted_questions = 0
    formatted_tables = 0
    
    # Xử lý các đoạn văn
    for para in doc.paragraphs:
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
            text = clean_text(text)
            if text != original_text:
                cleaned_paragraphs += 1
        
        # Tìm tất cả các biểu thức LaTeX
        patterns = [
            (r'\$\$(.+?)\$\$', 'display'),  # Toán học hiển thị
            (r'\$(.+?)\$', 'inline')        # Toán học nội tuyến
        ]
        
        matches = []
        for pattern, math_type in patterns:
            for match in re.finditer(pattern, text, re.DOTALL):
                matches.append((match.start(), match.end(), match.group(1), math_type))
        
        # Sắp xếp matches theo vị trí
        matches.sort(key=lambda x: x[0])
        
        last_idx = 0
        
        for start, end, latex_expr, math_type in matches:
            # Thêm văn bản trước biểu thức toán học
            if start > last_idx:
                text_before = text[last_idx:start]
                if text_before:
                    new_para.add_run(text_before)
            
            # Chuyển đổi LaTeX thành OMML và chèn
            omml_xml = latex_to_omml(latex_expr.strip())
            
            if omml_xml and options.get('convert_equations', True):
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
                    run.font.color.rgb = RGBColor(255, 0, 0)  # Màu đỏ
                else:
                    run = new_para.add_run(f"${latex_expr}$")
                error_count += 1
            
            last_idx = end
        
        # Thêm văn bản còn lại sau biểu thức toán học cuối cùng
        if last_idx < len(text):
            remaining_text = text[last_idx:]
            if remaining_text:
                new_para.add_run(remaining_text)
        
        # Nếu không tìm thấy biểu thức LaTeX, sao chép đoạn văn như cũ
        if not matches:
            new_para.clear()
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                # Sao chép định dạng run
                try:
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    if run.font.size:
                        new_run.font.size = run.font.size
                    if run.font.name:
                        new_run.font.name = run.font.name
                except:
                    pass
    
    # Sao chép các bảng từ tài liệu gốc
    for table in doc.tables:
        # Tạo bảng mới
        new_table = new_doc.add_table(rows=len(table.rows), cols=len(table.columns))
        
        # Sao chép nội dung bảng
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                new_table.cell(i, j).text = cell.text
        
        # Format bảng nếu được yêu cầu
        if options.get('format_tables', False):
            if format_table(new_table, options):
                formatted_tables += 1
    
    # Sửa định dạng tài liệu nếu được yêu cầu
    if options.get('fix_formatting', False):
        fix_formatting(new_doc, options)
        
        # Đếm số câu hỏi đã được format
        for para in new_doc.paragraphs:
            text = para.text
            if re.search(r'(\*\*Câu\s+\d+:\*\*|Câu\s+\d+:)', text, re.IGNORECASE):
                formatted_questions += 1
    
    return new_doc, {
        'processed_count': processed_count,
        'error_count': error_count,
        'cleaned_paragraphs': cleaned_paragraphs,
        'formatted_questions': formatted_questions,
        'formatted_tables': formatted_tables,
        'total_paragraphs': len(doc.paragraphs),
        'total_tables': len(doc.tables)
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
                            st.text(f"{i+1}. {para.text[:100]}{'...' if len(para.text) > 100 else ''}")
                    
                    if doc_preview.tables:
                        st.markdown("**📊 Bảng tìm thấy:**")
                        st.info(f"Tài liệu có {len(doc_preview.tables)} bảng")
                except:
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
                        new_doc, stats = process_word_document(doc, options)
                        
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
        <p><small>Phiên bản 2.0 - Tối ưu cho file PDF → Word từ AI</small></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
