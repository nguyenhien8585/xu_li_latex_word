import streamlit as st
import re
import io
import zipfile
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pandas as pd
import tempfile
import os

def convert_latex_to_word_equation(latex_text):
    """
    Chuyển đổi công thức LaTeX thành OMML (Office Math Markup Language)
    Đây là phiên bản đơn giản, có thể mở rộng thêm
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản
    conversions = {
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'<m:f><m:num>\1</m:num><m:den>\2</m:den></m:f>',
        r'\\sqrt\{([^}]+)\}': r'<m:rad><m:deg></m:deg><m:e>\1</m:e></m:rad>',
        r'\^([^{]|\{[^}]*\})': r'<m:sSup><m:e></m:e><m:sup>\1</m:sup></m:sSup>',
        r'_([^{]|\{[^}]*\})': r'<m:sSub><m:e></m:e><m:sub>\1</m:sub></m:sSub>',
        r'\\alpha': 'α',
        r'\\beta': 'β',
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\times': '×',
        r'\\div': '÷',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def create_math_element(math_text):
    """
    Tạo phần tử toán học trong Word
    """
    # Tạo phần tử math cơ bản
    math_element = OxmlElement('m:oMath')
    math_element.set(qn('xmlns:m'), "http://schemas.openxmlformats.org/officeDocument/2006/math")
    
    # Tạo run text
    run_element = OxmlElement('m:r')
    text_element = OxmlElement('m:t')
    text_element.text = math_text
    run_element.append(text_element)
    math_element.append(run_element)
    
    return math_element

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\$'
    
    # Lấy text từ paragraph
    full_text = paragraph.text
    
    # Tìm tất cả công thức
    matches = list(re.finditer(math_pattern, full_text))
    
    if not matches:
        return False
    
    # Xóa tất cả runs hiện tại
    for run in paragraph.runs:
        run._element.getparent().remove(run._element)
    
    # Xử lý text và thêm lại
    last_end = 0
    
    for match in matches:
        # Thêm text trước công thức
        if match.start() > last_end:
            run = paragraph.add_run(full_text[last_end:match.start()])
        
        # Thêm công thức toán học
        math_text = convert_latex_to_word_equation(match.group(1))
        try:
            # Tạo run cho công thức
            run = paragraph.add_run()
            math_element = create_math_element(math_text)
            run._element.append(math_element)
        except:
            # Nếu không tạo được công thức, thêm text thường
            run = paragraph.add_run(f"[MATH: {match.group(1)}]")
        
        last_end = match.end()
    
    # Thêm text còn lại
    if last_end < len(full_text):
        run = paragraph.add_run(full_text[last_end:])
    
    return True

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Lấy header
    header_line = lines[0]
    headers = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Bỏ qua dòng separator (dòng thứ 2)
    if len(lines) < 3:
        return [headers]
    
    # Lấy dữ liệu
    data = [headers]
    for line in lines[2:]:
        if line.strip():
            row = [cell.strip() for cell in line.split('|') if cell.strip()]
            if row:
                # Đảm bảo số cột bằng header
                while len(row) < len(headers):
                    row.append('')
                row = row[:len(headers)]
                data.append(row)
    
    return data

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word
    """
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|[^|\n]+\|[\n\r]+\|[-\s|:]+\|[\n\r]+(?:\|[^|\n]*\|[\n\r]*)+)'
    
    paragraphs_to_remove = []
    tables_to_add = []
    
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text
        matches = list(re.finditer(table_pattern, text, re.MULTILINE))
        
        for match in matches:
            table_text = match.group(1)
            table_data = parse_markdown_table(table_text)
            
            if table_data and len(table_data) > 1:
                # Lưu thông tin để thêm bảng
                tables_to_add.append({
                    'position': i,
                    'data': table_data,
                    'paragraph': paragraph
                })
                
                # Thay thế text trong paragraph
                new_text = text[:match.start()] + "[TABLE_PLACEHOLDER]" + text[match.end():]
                paragraph.text = new_text
    
    # Thêm bảng vào document
    for table_info in reversed(tables_to_add):  # Reverse để không ảnh hưởng index
        paragraph = table_info['paragraph']
        data = table_info['data']
        
        # Tạo bảng mới
        table = doc.add_table(rows=len(data), cols=len(data[0]))
        table.style = 'Table Grid'
        
        # Điền dữ liệu
        for row_idx, row_data in enumerate(data):
            row = table.rows[row_idx]
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < len(row.cells):
                    row.cells[col_idx].text = cell_data
                    # Định dạng header
                    if row_idx == 0:
                        row.cells[col_idx].paragraphs[0].runs[0].bold = True
        
        # Thay thế placeholder
        paragraph.text = paragraph.text.replace("[TABLE_PLACEHOLDER]", "")
        
        # Di chuyển bảng đến đúng vị trí
        table._element.getparent().remove(table._element)
        paragraph._element.addnext(table._element)

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên
    """
    try:
        # Đọc file Word
        doc = Document(uploaded_file)
        
        # Đếm số thay đổi
        math_count = 0
        table_count = 0
        
        # Xử lý công thức toán học
        for paragraph in doc.paragraphs:
            if process_paragraph_math(paragraph):
                math_count += 1
        
        # Xử lý bảng Markdown
        replace_markdown_tables_with_word_tables(doc)
        
        # Đếm bảng (đơn giản)
        table_pattern = r'(\|[^|\n]+\|[\n\r]+\|[-\s|:]+\|[\n\r]+(?:\|[^|\n]*\|[\n\r]*)+)'
        for paragraph in doc.paragraphs:
            table_count += len(re.findall(table_pattern, paragraph.text, re.MULTILINE))
        
        return doc, math_count, table_count
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download
    """
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
    # Sidebar với hướng dẫn
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        **Bước 1:** Tải lên file Word (.docx)
        
        **Bước 2:** Ứng dụng sẽ tự động:
        - Chuyển `$công thức$` → Equation Word
        - Chuyển `${công thức}$` → Equation Word  
        - Chuyển bảng Markdown → Bảng Word
        
        **Bước 3:** Tải xuống file đã xử lý
        
        ---
        
        **Ví dụ công thức LaTeX:**
        - `$x^2 + y^2 = z^2$`
        - `$\\frac{a}{b} + \\sqrt{c}$`
        - `$\\alpha + \\beta = \\gamma$`
        
        **Ví dụ bảng Markdown:**
        ```
        | Cột 1 | Cột 2 |
        |-------|-------|
        | A     | B     |
        | C     | D     |
        ```
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📤 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)",
            type=['docx'],
            help="Tải lên file Word chứa công thức LaTeX và/hoặc bảng Markdown"
        )
        
        if uploaded_file is not None:
            # Hiển thị thông tin file
            file_details = {
                "Tên file": uploaded_file.name,
                "Kích thước": f"{uploaded_file.size:,} bytes",
                "Loại file": uploaded_file.type
            }
            
            st.subheader("📄 Thông tin file")
            for key, value in file_details.items():
                st.write(f"**{key}:** {value}")
            
            # Xử lý file
            if st.button("🔄 Xử lý file", type="primary"):
                with st.spinner("Đang xử lý file..."):
                    doc, math_count, table_count = process_word_document(uploaded_file)
                    
                    if doc is not None:
                        st.success("✅ Xử lý file thành công!")
                        
                        # Hiển thị kết quả
                        col_result1, col_result2 = st.columns(2)
                        with col_result1:
                            st.metric("Công thức đã chuyển đổi", math_count)
                        with col_result2:
                            st.metric("Bảng đã chuyển đổi", table_count)
                        
                        # Tạo file download
                        processed_file = save_document(doc, uploaded_file.name)
                        
                        # Tạo tên file mới
                        base_name = uploaded_file.name.rsplit('.', 1)[0]
                        new_filename = f"{base_name}_processed.docx"
                        
                        # Nút download
                        st.download_button(
                            label="📥 Tải xuống file đã xử lý",
                            data=processed_file,
                            file_name=new_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="secondary"
                        )
                        
                        st.info("💡 File đã được xử lý và sẵn sàng tải xuống!")
    
    with col2:
        st.header("🛠️ Tính năng")
        
        features = [
            {
                "icon": "🧮",
                "title": "Chuyển đổi công thức",
                "desc": "Tự động chuyển $...$ và ${...}$ thành Equation Word"
            },
            {
                "icon": "📊", 
                "title": "Chuyển đổi bảng",
                "desc": "Chuyển bảng Markdown thành bảng Word đẹp mắt"
            },
            {
                "icon": "⚡",
                "title": "Xử lý nhanh",
                "desc": "Xử lý tự động và nhanh chóng"
            },
            {
                "icon": "💾",
                "title": "Tải xuống dễ dàng", 
                "desc": "Tải file đã xử lý với một click"
            }
        ]
        
        for feature in features:
            with st.container():
                st.markdown(f"""
                <div style="
                    border: 1px solid #e0e0e0;
                    border-radius: 10px;
                    padding: 15px;
                    margin: 10px 0;
                    background-color: #f9f9f9;
                ">
                    <h4>{feature['icon']} {feature['title']}</h4>
                    <p style="margin: 0; color: #666;">{feature['desc']}</p>
                </div>
                """, unsafe_allow_html=True)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; margin-top: 50px;">
        <p>🚀 <strong>Word Document Processor</strong> - Công cụ xử lý file Word với LaTeX và Markdown</p>
        <p>📧 Phát triển để hỗ trợ việc chuyển đổi công thức và bảng trong tài liệu Word</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
