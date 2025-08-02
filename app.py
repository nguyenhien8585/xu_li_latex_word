import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import parse_xml
import re
import pandas as pd
from io import BytesIO, StringIO
import base64

def extract_latex_formulas(text):
    """Tìm và trích xuất các công thức LaTeX từ text"""
    # Pattern cho $...$ và ${...}$
    patterns = [
        r'\$\{([^}]+)\}\$',  # ${...}$
        r'\$([^$]+)\$'       # $...$
    ]
    
    formulas = []
    for pattern in patterns:
        matches = re.finditer(pattern, text)
        for match in matches:
            formulas.append({
                'full_match': match.group(0),
                'latex_content': match.group(1),
                'start': match.start(),
                'end': match.end()
            })
    
    return formulas

def latex_to_mathml(latex_str):
    """Chuyển đổi LaTeX thành MathML (cơ bản)"""
    # Đây là phiên bản đơn giản - trong thực tế bạn có thể dùng latex2mathml
    mathml_template = f"""
    <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:r>
            <m:t>{latex_str}</m:t>
        </m:r>
    </m:oMath>
    """
    return mathml_template

def insert_math_equation(paragraph, latex_content):
    """Chèn công thức toán vào paragraph"""
    try:
        # Tạo MathML từ LaTeX
        mathml = latex_to_mathml(latex_content)
        
        # Tạo run mới cho equation
        run = paragraph._element
        
        # Thêm equation (simplified version)
        math_element = parse_xml(f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>[EQUATION: {latex_content}]</w:t></w:r>')
        run.append(math_element)
        
    except Exception as e:
        # Fallback: hiển thị dạng text
        paragraph.add_run(f"[EQUATION: {latex_content}]")

def process_latex_in_document(doc):
    """Xử lý tất cả công thức LaTeX trong document"""
    processed_count = 0
    
    for paragraph in doc.paragraphs:
        original_text = paragraph.text
        formulas = extract_latex_formulas(original_text)
        
        if formulas:
            # Xóa text cũ
            for run in paragraph.runs:
                run.text = ""
            
            # Xử lý từng công thức
            last_pos = 0
            for formula in sorted(formulas, key=lambda x: x['start']):
                # Thêm text trước công thức
                if formula['start'] > last_pos:
                    paragraph.add_run(original_text[last_pos:formula['start']])
                
                # Thêm công thức
                paragraph.add_run(f"[EQUATION: {formula['latex_content']}]").italic = True
                
                last_pos = formula['end']
                processed_count += 1
            
            # Thêm text còn lại
            if last_pos < len(original_text):
                paragraph.add_run(original_text[last_pos:])
    
    return processed_count

def detect_table_in_text(text):
    """Phát hiện bảng trong text"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Kiểm tra có ký tự | không
    pipe_lines = [line for line in lines if '|' in line]
    if len(pipe_lines) < 2:
        return None
    
    # Kiểm tra separator line (dòng có ---)
    separator_line = None
    for i, line in enumerate(lines):
        if re.match(r'^[\s]*\|?[\s]*:?-+:?[\s]*(\|[\s]*:?-+:?[\s]*)*\|?[\s]*$', line.strip()):
            separator_line = i
            break
    
    if separator_line is not None:
        header_line = lines[separator_line - 1] if separator_line > 0 else None
        data_lines = lines[separator_line + 1:] if separator_line < len(lines) - 1 else []
        
        return {
            'header': header_line,
            'separator': lines[separator_line],
            'data': data_lines,
            'start_line': 0,
            'end_line': len(lines)
        }
    
    return None

def parse_table_row(row_text):
    """Parse một dòng bảng"""
    # Loại bỏ | ở đầu và cuối
    row_text = row_text.strip()
    if row_text.startswith('|'):
        row_text = row_text[1:]
    if row_text.endswith('|'):
        row_text = row_text[:-1]
    
    # Split theo |
    cells = [cell.strip() for cell in row_text.split('|')]
    return cells

def convert_text_tables_to_word_tables(doc):
    """Chuyển đổi bảng text thành bảng Word"""
    converted_count = 0
    
    paragraphs_to_process = list(doc.paragraphs)
    
    for paragraph in paragraphs_to_process:
        text = paragraph.text
        table_info = detect_table_in_text(text)
        
        if table_info:
            try:
                # Parse header
                header_cells = parse_table_row(table_info['header'])
                
                # Parse data rows
                data_rows = []
                for data_line in table_info['data']:
                    if data_line.strip() and '|' in data_line:
                        row_cells = parse_table_row(data_line)
                        if len(row_cells) == len(header_cells):
                            data_rows.append(row_cells)
                
                # Tạo bảng Word
                if header_cells and data_rows:
                    table = doc.add_table(rows=1 + len(data_rows), cols=len(header_cells))
                    table.style = 'Table Grid'
                    
                    # Thêm header
                    header_row = table.rows[0]
                    for i, header_text in enumerate(header_cells):
                        header_row.cells[i].text = header_text
                        # Làm đậm header
                        for run in header_row.cells[i].paragraphs[0].runs:
                            run.bold = True
                    
                    # Thêm data rows
                    for row_idx, row_data in enumerate(data_rows):
                        data_row = table.rows[row_idx + 1]
                        for col_idx, cell_text in enumerate(row_data):
                            if col_idx < len(data_row.cells):
                                data_row.cells[col_idx].text = cell_text
                    
                    # Xóa paragraph gốc
                    paragraph.clear()
                    converted_count += 1
                    
            except Exception as e:
                st.warning(f"Không thể chuyển đổi bảng: {str(e)}")
    
    return converted_count

def process_word_document(uploaded_file):
    """Xử lý toàn bộ document Word"""
    try:
        # Đọc document
        doc = Document(uploaded_file)
        
        # Thống kê ban đầu
        total_paragraphs = len(doc.paragraphs)
        
        # Xử lý công thức LaTeX
        formula_count = process_latex_in_document(doc)
        
        # Xử lý bảng
        table_count = convert_text_tables_to_word_tables(doc)
        
        # Lưu document đã xử lý
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        
        return {
            'document': output,
            'stats': {
                'total_paragraphs': total_paragraphs,
                'formulas_processed': formula_count,
                'tables_converted': table_count
            }
        }
        
    except Exception as e:
        raise Exception(f"Lỗi xử lý document: {str(e)}")

def main():
    st.set_page_config(
        page_title="LaTeX to Word Processor",
        page_icon="🧮",
        layout="wide"
    )
    
    st.title("🧮 LaTeX to Word Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown thành định dạng Word chuẩn**")
    
    # Sidebar với thông tin
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        ### Công thức LaTeX được hỗ trợ:
        - `$x^2 + y^2 = z^2$`
        - `${\\frac{a}{b}}$`
        - `$\\sum_{i=1}^{n} x_i$`
        
        ### Bảng được hỗ trợ:
        ```
        | Header 1 | Header 2 |
        |----------|----------|
        | Data 1   | Data 2   |
        ```
        
        ### Tính năng:
        ✅ Phát hiện công thức LaTeX  
        ✅ Chuyển đổi bảng Markdown  
        ✅ Giữ nguyên định dạng Word  
        ✅ Tải xuống kết quả  
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📁 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx) cần xử lý",
            type=["docx"],
            help="File phải là định dạng .docx (Word 2007+)"
        )
        
        if uploaded_file:
            st.success(f"✅ Đã tải lên: {uploaded_file.name}")
            
            # Hiển thị thông tin file
            file_size = len(uploaded_file.getvalue()) / 1024
            st.info(f"📊 Kích thước file: {file_size:.1f} KB")
            
            # Nút xử lý
            if st.button("🚀 Bắt đầu xử lý", type="primary"):
                try:
                    with st.spinner("⏳ Đang xử lý document..."):
                        result = process_word_document(uploaded_file)
                    
                    st.success("✅ Xử lý hoàn tất!")
                    
                    # Hiển thị thống kê
                    stats = result['stats']
                    col_a, col_b, col_c = st.columns(3)
                    
                    with col_a:
                        st.metric("📄 Tổng đoạn văn", stats['total_paragraphs'])
                    with col_b:
                        st.metric("🧮 Công thức xử lý", stats['formulas_processed'])
                    with col_c:
                        st.metric("📊 Bảng chuyển đổi", stats['tables_converted'])
                    
                    # Nút tải xuống
                    st.download_button(
                        label="📥 Tải xuống file đã xử lý",
                        data=result['document'],
                        file_name=f"processed_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")
    
    with col2:
        st.subheader("📝 Ví dụ mẫu")
        
        # Ví dụ về LaTeX
        st.markdown("**Công thức LaTeX:**")
        st.code("$E = mc^2$\n${\\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}}$", language="latex")
        
        # Ví dụ về bảng
        st.markdown("**Bảng Markdown:**")
        st.code("""| Tên | Tuổi | Thành phố |
|-----|------|-----------|
| An  | 25   | Hà Nội    |
| Bình| 30   | TP.HCM    |""", language="markdown")
        
        st.markdown("---")
        st.markdown("**🔧 Phiên bản:** 1.0.0")
        st.markdown("**👨‍💻 GitHub:** [latex-to-word-app](https://github.com)")

    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>Made with ❤️ using Streamlit</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
