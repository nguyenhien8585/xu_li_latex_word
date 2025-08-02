import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement, ns
from docx.oxml.ns import qn
import zipfile
import xml.etree.ElementTree as ET
import base64

def create_math_equation(paragraph, latex_text):
    """Tạo công thức toán học trong Word từ LaTeX"""
    # Loại bỏ dấu $ 
    clean_latex = latex_text.strip('$').strip('{}').strip()
    
    # Tạo equation object trong Word
    math_element = OxmlElement('m:oMathPara')
    math_element.set(qn('xmlns:m'), 'http://schemas.openxmlformats.org/officeDocument/2006/math')
    
    # Tạo math run
    math_run = OxmlElement('m:oMath')
    
    # Tạo math text
    math_text = OxmlElement('m:r')
    math_text_element = OxmlElement('m:t')
    math_text_element.text = clean_latex
    math_text.append(math_text_element)
    math_run.append(math_text)
    math_element.append(math_run)
    
    # Thêm vào paragraph
    paragraph._element.append(math_element)

def latex_to_word_equation(latex_text):
    """Chuyển đổi LaTeX sang định dạng Word equation"""
    # Loại bỏ dấu $
    clean_text = latex_text.strip('$').strip('{}').strip()
    
    # Các chuyển đổi cơ bản từ LaTeX sang Word
    conversions = {
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\alpha': 'α',
        r'\\beta': 'β',
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\phi': 'φ',
        r'\\omega': 'ω',
        r'\\infty': '∞',
        r'\\pm': '±',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\approx': '≈',
        r'\^(\w+)': r'^\1',
        r'_(\w+)': r'_\1'
    }
    
    result = clean_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def parse_markdown_table(table_text):
    """Phân tích bảng Markdown và trả về dữ liệu"""
    lines = table_text.strip().split('\n')
    
    # Lọc các dòng trống
    lines = [line.strip() for line in lines if line.strip()]
    
    if len(lines) < 2:
        return None
    
    # Lấy header
    header = [cell.strip() for cell in lines[0].split('|') if cell.strip()]
    
    # Bỏ qua dòng separator (dòng thứ 2)
    data_lines = lines[2:] if len(lines) > 2 else []
    
    # Lấy dữ liệu
    rows = []
    for line in data_lines:
        row = [cell.strip() for cell in line.split('|') if cell.strip()]
        if row:  # Chỉ thêm nếu row không rỗng
            rows.append(row)
    
    return {'header': header, 'rows': rows}

def add_table_to_doc(doc, table_data):
    """Thêm bảng vào document Word"""
    if not table_data or not table_data['header']:
        return
    
    # Tạo bảng với số hàng và cột phù hợp
    num_cols = len(table_data['header'])
    num_rows = len(table_data['rows']) + 1  # +1 cho header
    
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Table Grid'
    
    # Thêm header
    header_row = table.rows[0]
    for i, header_text in enumerate(table_data['header']):
        if i < len(header_row.cells):
            header_row.cells[i].text = header_text
            # Làm đậm header
            for paragraph in header_row.cells[i].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
    
    # Thêm dữ liệu
    for row_idx, row_data in enumerate(table_data['rows']):
        if row_idx + 1 < len(table.rows):
            table_row = table.rows[row_idx + 1]
            for col_idx, cell_text in enumerate(row_data):
                if col_idx < len(table_row.cells):
                    table_row.cells[col_idx].text = cell_text

def process_word_document(uploaded_file):
    """Xử lý file Word đã upload"""
    try:
        # Đọc file Word
        doc = Document(uploaded_file)
        
        # Tìm và thay thế công thức toán học
        math_pattern = r'\$\{?([^$}]+)\}?\$'
        
        # Tìm và thay thế bảng Markdown
        table_pattern = r'\|[^|\n]*\|[\s\S]*?\n(?=\s*\n|\s*$)'
        
        processed_paragraphs = []
        
        for paragraph in doc.paragraphs:
            original_text = paragraph.text
            
            # Xử lý công thức toán học
            if re.search(math_pattern, original_text):
                # Tìm tất cả công thức trong paragraph
                matches = re.finditer(math_pattern, original_text)
                
                # Thay thế từng công thức
                new_text = original_text
                for match in reversed(list(matches)):  # Đảo ngược để không ảnh hưởng index
                    latex_formula = match.group(0)
                    converted_formula = latex_to_word_equation(latex_formula)
                    new_text = new_text[:match.start()] + converted_formula + new_text[match.end():]
                
                paragraph.text = new_text
                processed_paragraphs.append(f"✓ Đã chuyển đổi công thức: {original_text[:50]}...")
            
            # Xử lý bảng Markdown
            elif '|' in original_text and original_text.count('|') >= 2:
                # Kiểm tra xem có phải bảng Markdown không
                lines = original_text.split('\n')
                table_lines = [line for line in lines if '|' in line]
                
                if len(table_lines) >= 2:
                    table_text = '\n'.join(table_lines)
                    table_data = parse_markdown_table(table_text)
                    
                    if table_data:
                        # Xóa text gốc
                        paragraph.clear()
                        
                        # Thêm bảng mới
                        add_table_to_doc(doc, table_data)
                        processed_paragraphs.append(f"✓ Đã chuyển đổi bảng: {len(table_data['header'])} cột, {len(table_data['rows'])} hàng")
        
        return doc, processed_paragraphs
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        return None, []

def create_download_link(doc, filename):
    """Tạo link tải xuống file Word"""
    # Lưu document vào BytesIO
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    
    # Tạo base64 string
    b64 = base64.b64encode(doc_io.read()).decode()
    
    # Tạo link tải xuống
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}">📥 Tải xuống file đã chuyển đổi</a>'
    
    return href

def main():
    st.set_page_config(
        page_title="Word Converter",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Converter")
    st.markdown("**Chuyển đổi công thức toán học và bảng Markdown trong file Word**")
    
    # Sidebar với hướng dẫn
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        **Bước 1:** Upload file Word (.docx)
        
        **Bước 2:** Ứng dụng sẽ tự động:
        - Chuyển `$formula$` thành equation Word
        - Chuyển bảng Markdown thành bảng Word
        
        **Bước 3:** Tải xuống file đã chuyển đổi
        
        ---
        
        **Ví dụ công thức hỗ trợ:**
        - `$x^2 + y^2 = z^2$`
        - `$\\frac{a}{b}$`
        - `$\\sqrt{x}$`
        - `$\\alpha + \\beta$`
        
        **Ví dụ bảng Markdown:**
        ```
        | Cột 1 | Cột 2 |
        |-------|-------|
        | A     | B     |
        | C     | D     |
        ```
        """)
    
    # Main content
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("📤 Upload File")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)", 
            type=['docx'],
            help="Upload file Word chứa công thức toán học hoặc bảng Markdown"
        )
        
        if uploaded_file is not None:
            st.success(f"✅ Đã upload: {uploaded_file.name}")
            
            # Hiển thị thông tin file
            file_size = len(uploaded_file.getvalue()) / 1024  # KB
            st.info(f"📊 Kích thước file: {file_size:.1f} KB")
    
    with col2:
        st.header("⚙️ Xử lý & Kết quả")
        
        if uploaded_file is not None:
            if st.button("🔄 Bắt đầu chuyển đổi", type="primary", use_container_width=True):
                with st.spinner("Đang xử lý file..."):
                    # Xử lý file
                    processed_doc, processing_log = process_word_document(uploaded_file)
                    
                    if processed_doc:
                        st.success("✅ Chuyển đổi thành công!")
                        
                        # Hiển thị log xử lý
                        if processing_log:
                            st.subheader("📋 Chi tiết xử lý:")
                            for log_entry in processing_log:
                                st.text(log_entry)
                        else:
                            st.info("ℹ️ Không tìm thấy công thức hoặc bảng nào cần chuyển đổi")
                        
                        # Tạo tên file mới
                        original_name = uploaded_file.name
                        new_filename = f"converted_{original_name}"
                        
                        # Tạo link tải xuống
                        download_link = create_download_link(processed_doc, new_filename)
                        st.markdown(download_link, unsafe_allow_html=True)
                        
                        # Thêm button tải xuống alternative
                        doc_io = io.BytesIO()
                        processed_doc.save(doc_io)
                        doc_io.seek(0)
                        
                        st.download_button(
                            label="💾 Tải xuống (Alternative)",
                            data=doc_io.getvalue(),
                            file_name=new_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error("❌ Có lỗi xảy ra khi xử lý file")
        else:
            st.info("👆 Vui lòng upload file Word để bắt đầu")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #888;'>
        <p>🚀 Word Document Converter | Chuyển đổi công thức và bảng tự động</p>
        <p>💡 Hỗ trợ LaTeX equations và Markdown tables</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
