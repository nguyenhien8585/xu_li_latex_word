import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import base64

def create_equation_xml(latex_formula):
    """Tạo XML cho equation Word từ LaTeX"""
    
    # Loại bỏ dấu $
    clean_formula = latex_formula.strip('$').strip('{}').strip()
    
    # Chuyển đổi các patterns LaTeX phổ biến
    xml_parts = []
    
    # Xử lý phân số \frac{a}{b}
    frac_pattern = r'\\frac\{([^}]+)\}\{([^}]+)\}'
    while re.search(frac_pattern, clean_formula):
        match = re.search(frac_pattern, clean_formula)
        numerator = match.group(1)
        denominator = match.group(2)
        
        frac_xml = f'''
        <m:f>
            <m:num>
                <m:r><m:t>{numerator}</m:t></m:r>
            </m:num>
            <m:den>
                <m:r><m:t>{denominator}</m:t></m:r>
            </m:den>
        </m:f>
        '''
        
        clean_formula = clean_formula[:match.start()] + "FRACTION_PLACEHOLDER" + clean_formula[match.end():]
        xml_parts.append(('FRACTION_PLACEHOLDER', frac_xml.strip()))
    
    # Xử lý căn bậc hai \sqrt{x}
    sqrt_pattern = r'\\sqrt\{([^}]+)\}'
    while re.search(sqrt_pattern, clean_formula):
        match = re.search(sqrt_pattern, clean_formula)
        content = match.group(1)
        
        sqrt_xml = f'''
        <m:rad>
            <m:deg></m:deg>
            <m:e>
                <m:r><m:t>{content}</m:t></m:r>
            </m:e>
        </m:rad>
        '''
        
        clean_formula = clean_formula[:match.start()] + "SQRT_PLACEHOLDER" + clean_formula[match.end():]
        xml_parts.append(('SQRT_PLACEHOLDER', sqrt_xml.strip()))
    
    # Xử lý lũy thừa x^{n}
    power_pattern = r'([a-zA-Z_]\w*)\^\{([^}]+)\}'
    while re.search(power_pattern, clean_formula):
        match = re.search(power_pattern, clean_formula)
        base = match.group(1)
        exponent = match.group(2)
        
        power_xml = f'''
        <m:sSup>
            <m:e>
                <m:r><m:t>{base}</m:t></m:r>
            </m:e>
            <m:sup>
                <m:r><m:t>{exponent}</m:t></m:r>
            </m:sup>
        </m:sSup>
        '''
        
        clean_formula = clean_formula[:match.start()] + "POWER_PLACEHOLDER" + clean_formula[match.end():]
        xml_parts.append(('POWER_PLACEHOLDER', power_xml.strip()))
    
    # Xử lý chỉ số dưới x_{n}
    sub_pattern = r'([a-zA-Z_]\w*)_\{([^}]+)\}'
    while re.search(sub_pattern, clean_formula):
        match = re.search(sub_pattern, clean_formula)
        base = match.group(1)
        subscript = match.group(2)
        
        sub_xml = f'''
        <m:sSub>
            <m:e>
                <m:r><m:t>{base}</m:t></m:r>
            </m:e>
            <m:sub>
                <m:r><m:t>{subscript}</m:t></m:r>
            </m:sub>
        </m:sSub>
        '''
        
        clean_formula = clean_formula[:match.start()] + "SUB_PLACEHOLDER" + clean_formula[match.end():]
        xml_parts.append(('SUB_PLACEHOLDER', sub_xml.strip()))
    
    # Thay thế các ký hiệu LaTeX thành Unicode
    symbol_replacements = {
        r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ',
        r'\\epsilon': 'ε', r'\\theta': 'θ', r'\\lambda': 'λ', r'\\mu': 'μ',
        r'\\pi': 'π', r'\\sigma': 'σ', r'\\phi': 'φ', r'\\omega': 'ω',
        r'\\Omega': 'Ω', r'\\Delta': 'Δ', r'\\Gamma': 'Γ', r'\\Lambda': 'Λ',
        r'\\infty': '∞', r'\\pm': '±', r'\\times': '×', r'\\div': '÷',
        r'\\leq': '≤', r'\\geq': '≥', r'\\neq': '≠', r'\\approx': '≈',
        r'\\sum': '∑', r'\\int': '∫', r'\\cot': 'cot', r'\\tan': 'tan',
        r'\\cos': 'cos', r'\\sin': 'sin', r'\\ln': 'ln', r'\\log': 'log',
        r'\\lim': 'lim', r'\\partial': '∂', r'\\nabla': '∇',
        r'\\left\(': '(', r'\\right\)': ')', r'\\left\{': '{', r'\\right\}': '}',
        r'\\left\[': '[', r'\\right\]': ']', r'\\mathbb\{([^}]+)\}': r'\1'
    }
    
    for latex_sym, unicode_sym in symbol_replacements.items():
        clean_formula = re.sub(latex_sym, unicode_sym, clean_formula)
    
    # Tạo XML cuối cùng
    result_xml = f'<m:r><m:t>{clean_formula}</m:t></m:r>'
    
    # Thay thế placeholders bằng XML thực
    for placeholder, xml_content in xml_parts:
        result_xml = result_xml.replace(placeholder, xml_content)
    
    return result_xml

def insert_equation_into_paragraph(paragraph, latex_formula):
    """Chèn equation vào paragraph"""
    try:
        # Tạo math XML
        equation_xml = create_equation_xml(latex_formula)
        
        # Tạo complete math XML với namespace
        complete_xml = f'''
        <m:oMathPara {nsdecls('m')}>
            <m:oMath>
                {equation_xml}
            </m:oMath>
        </m:oMathPara>
        '''
        
        # Parse và insert
        math_element = parse_xml(complete_xml)
        paragraph._element.append(math_element)
        
        return True
    except Exception as e:
        st.error(f"Lỗi tạo equation: {str(e)}")
        # Fallback: thêm text thường
        paragraph.add_run(latex_formula)
        return False

def process_enhanced_document(uploaded_file):
    """Xử lý file Word với equation support nâng cao"""
    try:
        doc = Document(uploaded_file)
        processed_count = 0
        processing_log = []
        
        # Patterns để detect math formulas
        math_patterns = [
            r'\$[^$]+\$',                    # $formula$
            r'\$\{[^}]+\}\$',               # ${formula}$
            r'\\[a-zA-Z]+\{[^}]*\}',        # \frac{}, \sqrt{}, etc.
        ]
        
        for para_idx, paragraph in enumerate(doc.paragraphs):
            original_text = paragraph.text
            
            # Tìm tất cả formulas trong paragraph
            all_matches = []
            for pattern in math_patterns:
                matches = list(re.finditer(pattern, original_text))
                all_matches.extend(matches)
            
            if all_matches:
                # Sắp xếp theo vị trí
                all_matches.sort(key=lambda x: x.start())
                
                # Tạo paragraph mới
                new_paragraph = doc.add_paragraph()
                
                # Xử lý từng phần
                last_end = 0
                equation_count = 0
                
                for match in all_matches:
                    # Thêm text trước formula
                    if match.start() > last_end:
                        text_before = original_text[last_end:match.start()]
                        if text_before.strip():
                            new_paragraph.add_run(text_before)
                    
                    # Thêm equation
                    formula = match.group(0)
                    success = insert_equation_into_paragraph(new_paragraph, formula)
                    if success:
                        equation_count += 1
                    
                    last_end = match.end()
                
                # Thêm text còn lại
                if last_end < len(original_text):
                    remaining_text = original_text[last_end:]
                    if remaining_text.strip():
                        new_paragraph.add_run(remaining_text)
                
                # Xóa paragraph cũ và thay thế
                paragraph._element.getparent().replace(paragraph._element, new_paragraph._element)
                
                if equation_count > 0:
                    processed_count += equation_count
                    processing_log.append(f"✓ Paragraph {para_idx + 1}: Chuyển đổi {equation_count} công thức")
            
            # Xử lý bảng Markdown (giữ nguyên logic cũ)
            elif '|' in original_text and original_text.count('|') >= 2:
                lines = original_text.split('\n')
                table_lines = [line for line in lines if '|' in line]
                
                if len(table_lines) >= 2:
                    table_text = '\n'.join(table_lines)
                    table_data = parse_markdown_table(table_text)
                    
                    if table_data:
                        paragraph.clear()
                        add_table_to_doc(doc, table_data)
                        processing_log.append(f"✓ Đã chuyển đổi bảng: {len(table_data['header'])} cột, {len(table_data['rows'])} hàng")
        
        if processed_count == 0:
            processing_log.append("ℹ️ Không tìm thấy công thức nào để chuyển đổi")
        else:
            processing_log.insert(0, f"🎉 Tổng cộng đã chuyển đổi {processed_count} công thức thành equation Word!")
        
        return doc, processing_log
        
    except Exception as e:
        st.error(f"Lỗi xử lý file: {str(e)}")
        return None, [f"❌ Lỗi: {str(e)}"]

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

def create_download_link(doc, filename):
    """Tạo link tải xuống file Word"""
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    
    b64 = base64.b64encode(doc_io.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" style="text-decoration: none; color: #0066CC; font-weight: bold;">📥 Tải xuống file đã chuyển đổi</a>'
    
    return href

def main():
    st.set_page_config(
        page_title="Enhanced Word Converter",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Enhanced Word Document Converter")
    st.markdown("**Chuyển đổi công thức toán học thành Equation Word thực sự + Bảng Markdown**")
    
    # Info box
    st.info("🆕 **Phiên bản nâng cao**: Tạo equation Word thực sự từ LaTeX, không chỉ là text!")
    
    # Sidebar
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        **✨ Tính năng mới:**
        - Tạo **equation Word thực sự** từ LaTeX
        - Hỗ trợ nhiều cú pháp LaTeX phức tạp
        - Chuyển đổi bảng Markdown
        
        **🧮 Công thức được hỗ trợ:**
        - `$x^{2} + y^{2} = z^{2}$`
        - `$\\frac{a}{b}$` → Phân số
        - `$\\sqrt{x}$` → Căn bậc hai
        - `$x_{1}, x_{2}$` → Chỉ số dưới
        - `$\\alpha, \\beta, \\pi$` → Ký hiệu Hy Lạp
        - `$\\sum, \\int$` → Tổng, tích phân
        
        **📊 Bảng Markdown:**
        ```
        | Cột 1 | Cột 2 |
        |-------|-------|
        | A     | B     |
        ```
        
        **🎯 Kết quả:**
        - Equation Word có thể edit được
        - Bảng Word với format đẹp
        - Giữ nguyên text khác
        """)
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("📤 Upload File")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)", 
            type=['docx'],
            help="Upload file Word chứa công thức LaTeX hoặc bảng Markdown"
        )
        
        if uploaded_file is not None:
            st.success(f"✅ Đã upload: {uploaded_file.name}")
            file_size = len(uploaded_file.getvalue()) / 1024
            st.info(f"📊 Kích thước: {file_size:.1f} KB")
            
            # Preview mode
            with st.expander("🔍 Preview nội dung file"):
                try:
                    preview_doc = Document(uploaded_file)
                    preview_text = ""
                    for i, para in enumerate(preview_doc.paragraphs[:5]):  # Show first 5 paragraphs
                        if para.text.strip():
                            preview_text += f"Đoạn {i+1}: {para.text[:100]}...\n\n"
                    st.text_area("Nội dung mẫu:", preview_text, height=200)
                    
                    if len(preview_doc.paragraphs) > 5:
                        st.info(f"... và {len(preview_doc.paragraphs) - 5} đoạn văn khác")
                except:
                    st.warning("Không thể preview file")
    
    with col2:
        st.header("⚙️ Xử lý & Kết quả")
        
        if uploaded_file is not None:
            if st.button("🚀 Chuyển đổi thành Equation Word", type="primary", use_container_width=True):
                with st.spinner("🔄 Đang xử lý file và tạo equation..."):
                    # Xử lý file với enhanced algorithm
                    processed_doc, processing_log = process_enhanced_document(uploaded_file)
                    
                    if processed_doc:
                        st.success("🎉 Chuyển đổi thành công!")
                        
                        # Hiển thị log chi tiết
                        with st.expander("📋 Chi tiết quá trình xử lý", expanded=True):
                            for log_entry in processing_log:
                                if log_entry.startswith("✓"):
                                    st.success(log_entry)
                                elif log_entry.startswith("ℹ️"):
                                    st.info(log_entry)
                                elif log_entry.startswith("🎉"):
                                    st.balloons()
                                    st.success(log_entry)
                                else:
                                    st.text(log_entry)
                        
                        # Tạo file tải xuống
                        original_name = uploaded_file.name
                        new_filename = f"enhanced_{original_name}"
                        
                        # Tạo link tải xuống
                        download_link = create_download_link(processed_doc, new_filename)
                        st.markdown(download_link, unsafe_allow_html=True)
                        
                        # Download button
                        doc_io = io.BytesIO()
                        processed_doc.save(doc_io)
                        doc_io.seek(0)
                        
                        st.download_button(
                            label="💾 Tải xuống (Button)",
                            data=doc_io.getvalue(),
                            file_name=new_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                        st.success("✅ File đã sẵn sàng! Mở trong Word để xem equation thực sự.")
                        
                    else:
                        st.error("❌ Có lỗi xảy ra khi xử lý file")
        else:
            st.info("👆 Upload file Word để bắt đầu chuyển đổi")
            
            # Example preview
            with st.expander("📝 Ví dụ về kết quả"):
                st.markdown("""
                **Trước khi chuyển đổi:**
                ```
                Công thức Einstein: $E = mc^{2}$
                Phân số: $\\frac{a+b}{c-d}$
                Căn bậc hai: $\\sqrt{x^{2} + y^{2}}$
                ```
                
                **Sau khi chuyển đổi:**
                - Tạo equation Word thực sự có thể edit
                - Hiển thị đẹp như equation editor
                - Có thể copy/paste giữa các document
                """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666; font-size: 0.9em;'>
        <p><strong>🚀 Enhanced Word Converter</strong> - Phiên bản nâng cao</p>
        <p>💡 Tạo equation Word thực sự + Chuyển đổi bảng Markdown tự động</p>
        <p>🔧 Hỗ trợ LaTeX syntax phức tạp và OMML format</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
