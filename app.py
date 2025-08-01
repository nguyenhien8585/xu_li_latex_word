import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
    page_title="Xử lý LaTeX trong Word",
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

def fix_formatting(doc):
    """Sửa định dạng cơ bản của tài liệu"""
    for paragraph in doc.paragraphs:
        # Sửa khoảng cách dòng
        if paragraph.paragraph_format.line_spacing != 1.15:
            paragraph.paragraph_format.line_spacing = 1.15
        
        # Sửa font chữ cho các run
        for run in paragraph.runs:
            if run.font.name != 'Times New Roman':
                run.font.name = 'Times New Roman'
            if not run.font.size or run.font.size < Pt(11):
                run.font.size = Pt(12)

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
                    run.font.color.rgb = None  # Màu đỏ cần giá trị RGB
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
    
    # Sửa định dạng tài liệu nếu được yêu cầu
    if options.get('fix_formatting', False):
        fix_formatting(new_doc)
    
    return new_doc, {
        'processed_count': processed_count,
        'error_count': error_count,
        'cleaned_paragraphs': cleaned_paragraphs,
        'total_paragraphs': len(doc.paragraphs)
    }

def main():
    # Header
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 Xử lý LaTeX trong Word</h1>
        <p style="color: white; margin: 10px 0 0 0;">Chuyển đổi biểu thức LaTeX thành công thức toán chuẩn trong Word</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar với các tùy chọn
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn xử lý")
        
        convert_equations = st.checkbox("🔢 Chuyển đổi công thức LaTeX", value=True,
                                       help="Chuyển $...$ thành công thức toán chuẩn")
        
        fix_formatting = st.checkbox("🎨 Sửa định dạng cơ bản", value=True,
                                    help="Chuẩn hóa font chữ, kích thước và khoảng cách")
        
        clean_text = st.checkbox("🧹 Làm sạch văn bản", value=True,
                                help="Loại bỏ ký tự lạ và chuẩn hóa khoảng trắng")
        
        show_errors = st.checkbox("⚠️ Hiển thị lỗi chuyển đổi", value=True,
                                 help="Đánh dấu các biểu thức không chuyển đổi được")
        
        st.markdown("---")
        st.markdown("### 📖 Hướng dẫn sử dụng")
        st.markdown("""
        1. **Tải file**: Chọn file Word (.docx) 
        2. **Cấu hình**: Điều chỉnh các tùy chọn bên trái
        3. **Xử lý**: Nhấn nút "Xử lý file"
        4. **Tải về**: Download file đã xử lý
        
        **Định dạng LaTeX hỗ trợ:**
        - `$x^2$` → x²  
        - `$\\frac{a}{b}$` → phân số
        - `$\\sqrt{x}$` → căn bậc hai
        - `$a_n$` → chỉ số dưới
        - `$\\alpha$` → ký tự Hy Lạp
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📁 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx) cần xử lý",
            type=["docx"],
            help="File Word đã được chuyển từ PDF hoặc tạo bằng AI"
        )
        
        if uploaded_file is not None:
            # Thông tin file
            file_info = {
                'name': uploaded_file.name,
                'size': f"{uploaded_file.size / 1024:.1f} KB",
                'type': uploaded_file.type
            }
            
            st.success(f"✅ Đã tải lên: **{file_info['name']}** ({file_info['size']})")
            
            # Nút xử lý
            if st.button("🚀 Xử lý file", type="primary", use_container_width=True):
                options = {
                    'convert_equations': convert_equations,
                    'fix_formatting': fix_formatting,
                    'clean_text': clean_text,
                    'show_errors': show_errors
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
                        
                        metric_cols = st.columns(4)
                        with metric_cols[0]:
                            st.metric("✅ Công thức đã chuyển", stats['processed_count'])
                        with metric_cols[1]:
                            st.metric("❌ Lỗi chuyển đổi", stats['error_count'])
                        with metric_cols[2]:
                            st.metric("🧹 Đoạn văn đã sạch", stats['cleaned_paragraphs'])
                        with metric_cols[3]:
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
                        if stats['processed_count'] > 0:
                            st.success(f"🎉 Đã xử lý thành công {stats['processed_count']} công thức LaTeX!")
                        if stats['error_count'] > 0:
                            st.warning(f"⚠️ Có {stats['error_count']} biểu thức gặp lỗi hoặc không được hỗ trợ.")
                        if stats['cleaned_paragraphs'] > 0:
                            st.info(f"🧹 Đã làm sạch {stats['cleaned_paragraphs']} đoạn văn.")
                
                except Exception as e:
                    st.error(f"❌ Lỗi khi xử lý file: {str(e)}")
                    st.info("💡 Kiểm tra lại định dạng file và thử lại.")
    
    with col2:
        st.markdown("### 💡 Mẹo hữu ích")
        
        with st.expander("🔍 Kiểm tra file trước khi xử lý"):
            st.markdown("""
            - Đảm bảo file là định dạng `.docx`
            - Kiểm tra các biểu thức LaTeX có đúng cú pháp
            - File không bị khóa hoặc mã hóa
            - Kích thước file không quá lớn (< 10MB)
            """)
        
        with st.expander("🎯 Tối ưu kết quả"):
            st.markdown("""
            - **Bật "Sửa định dạng"** cho file từ PDF
            - **Bật "Làm sạch văn bản"** nếu có ký tự lạ  
            - **Tắt "Hiển thị lỗi"** nếu muốn ẩn lỗi
            - Kiểm tra kết quả trước khi sử dụng
            """)
        
        with st.expander("⚡ Ví dụ LaTeX phổ biến"):
            st.code("""
# Cơ bản
$x + y = z$
$a^2 + b^2 = c^2$

# Phân số
$\\frac{numerator}{denominator}$

# Căn thức
$\\sqrt{x}$
$\\sqrt[n]{x}$

# Tích phân
$\\int_0^1 f(x) dx$

# Ma trận
$\\begin{matrix} a & b \\\\ c & d \\end{matrix}$
            """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>🚀 Được phát triển với Streamlit | 📧 Hỗ trợ: support@example.com</p>
        <p><small>Phiên bản 1.0 - Hỗ trợ chuyển đổi LaTeX thành công thức Word chuẩn</small></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
