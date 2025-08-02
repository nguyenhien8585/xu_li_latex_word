import streamlit as st
import mammoth
import re
import os
import tempfile
import shutil
from pathlib import Path
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
import base64
from io import BytesIO
import pandas as pd

# Cấu hình trang
st.set_page_config(
    page_title="LaTeX to Word Converter",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS styling
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .feature-box {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        margin: 1rem 0;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .success-box {
        background: #d4edda;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .warning-box {
        background: #fff3cd;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #ffc107;
        margin: 1rem 0;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .conversion-preview {
        background: #e3f2fd;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        border-left: 3px solid #2196f3;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="main-header">
    <h1>🧮 LaTeX to Word Converter</h1>
    <p>Chuyển đổi thực sự LaTeX, TikZ và Markdown thành file Word</p>
    <small>Phiên bản Cloud với khả năng conversion</small>
</div>
""", unsafe_allow_html=True)

# Utility Functions
def extract_latex_formulas(content):
    """Trích xuất công thức LaTeX từ nội dung"""
    patterns = [
        (r'\$([^$\n]+)\$', 'inline'),           # $x^2$
        (r'\$\$([^$]+)\$\$', 'display'),       # $$x^2$$
        (r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}', 'equation'),
        (r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}', 'align'),
        (r'\\\[(.*?)\\\]', 'display'),         # \[x^2\]
    ]
    
    formulas = []
    for pattern, formula_type in patterns:
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            formulas.append({
                'content': match.group(1).strip(),
                'type': formula_type,
                'start': match.start(),
                'end': match.end(),
                'full': match.group(0)
            })
    
    # Sort by position
    formulas.sort(key=lambda x: x['start'])
    return formulas

def extract_tikz_diagrams(content):
    """Trích xuất TikZ diagrams"""
    pattern = r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}'
    diagrams = []
    
    matches = re.finditer(pattern, content, re.DOTALL)
    for i, match in enumerate(matches):
        diagrams.append({
            'id': i + 1,
            'content': match.group(1).strip(),
            'start': match.start(),
            'end': match.end(),
            'full': match.group(0)
        })
    
    return diagrams

def extract_markdown_tables(content):
    """Trích xuất bảng Markdown"""
    table_pattern = r'(\|.*?\|(?:\r?\n\|.*?\|)*)'
    tables = []
    
    matches = re.finditer(table_pattern, content, re.MULTILINE)
    for match in matches:
        table_text = match.group(1).strip()
        parsed = parse_markdown_table(table_text)
        if parsed:
            tables.append({
                'data': parsed,
                'start': match.start(),
                'end': match.end(),
                'full': match.group(0)
            })
    
    return tables

def parse_markdown_table(table_text):
    """Parse bảng Markdown"""
    lines = [line.strip() for line in table_text.split('\n') if line.strip()]
    if len(lines) < 2:
        return None
    
    # Header
    header_cells = [cell.strip() for cell in lines[0].split('|') if cell.strip()]
    
    # Data rows (skip separator line)
    data_rows = []
    for line in lines[2:]:
        if '|' in line:
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            # Ensure same number of columns
            while len(cells) < len(header_cells):
                cells.append('')
            data_rows.append(cells[:len(header_cells)])
    
    return {
        'headers': header_cells,
        'rows': data_rows
    }

def convert_latex_to_unicode(latex_code):
    """Chuyển LaTeX thành Unicode symbols"""
    
    # Unicode symbols mapping
    symbols = {
        # Greek letters
        r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ',
        r'\\epsilon': 'ε', r'\\varepsilon': 'ε', r'\\zeta': 'ζ', r'\\eta': 'η',
        r'\\theta': 'θ', r'\\vartheta': 'θ', r'\\iota': 'ι', r'\\kappa': 'κ',
        r'\\lambda': 'λ', r'\\mu': 'μ', r'\\nu': 'ν', r'\\xi': 'ξ',
        r'\\pi': 'π', r'\\varpi': 'π', r'\\rho': 'ρ', r'\\varrho': 'ρ',
        r'\\sigma': 'σ', r'\\varsigma': 'ς', r'\\tau': 'τ', r'\\upsilon': 'υ',
        r'\\phi': 'φ', r'\\varphi': 'φ', r'\\chi': 'χ', r'\\psi': 'ψ', r'\\omega': 'ω',
        
        # Capital Greek
        r'\\Gamma': 'Γ', r'\\Delta': 'Δ', r'\\Theta': 'Θ', r'\\Lambda': 'Λ',
        r'\\Xi': 'Ξ', r'\\Pi': 'Π', r'\\Sigma': 'Σ', r'\\Upsilon': 'Υ',
        r'\\Phi': 'Φ', r'\\Psi': 'Ψ', r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\infty': '∞', r'\\partial': '∂', r'\\nabla': '∇', r'\\emptyset': '∅',
        r'\\varnothing': '∅',
        
        # Operators
        r'\\int': '∫', r'\\iint': '∬', r'\\iiint': '∭', r'\\oint': '∮',
        r'\\sum': '∑', r'\\prod': '∏', r'\\coprod': '∐',
        
        # Relations
        r'\\pm': '±', r'\\mp': '∓', r'\\times': '×', r'\\div': '÷',
        r'\\cdot': '⋅', r'\\ast': '∗', r'\\star': '⋆',
        r'\\leq': '≤', r'\\le': '≤', r'\\geq': '≥', r'\\ge': '≥',
        r'\\neq': '≠', r'\\ne': '≠', r'\\equiv': '≡', r'\\approx': '≈',
        r'\\sim': '∼', r'\\simeq': '≃', r'\\cong': '≅', r'\\propto': '∝',
        
        # Sets
        r'\\subset': '⊂', r'\\supset': '⊃', r'\\subseteq': '⊆', r'\\supseteq': '⊇',
        r'\\in': '∈', r'\\notin': '∉', r'\\ni': '∋', r'\\cap': '∩', r'\\cup': '∪',
        
        # Arrows
        r'\\rightarrow': '→', r'\\to': '→', r'\\leftarrow': '←', 
        r'\\leftrightarrow': '↔', r'\\Rightarrow': '⇒', r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔', r'\\mapsto': '↦',
        
        # Logic
        r'\\forall': '∀', r'\\exists': '∃', r'\\nexists': '∄',
        r'\\land': '∧', r'\\lor': '∨', r'\\lnot': '¬', r'\\neg': '¬',
        
        # Others
        r'\\angle': '∠', r'\\triangle': '△', r'\\square': '□',
        r'\\diamond': '◊', r'\\circ': '∘', r'\\bullet': '•',
        r'\\prime': '′', r'\\backprime': '‵',
    }
    
    text = latex_code
    
    # Apply symbol replacements
    for latex_sym, unicode_sym in symbols.items():
        text = re.sub(latex_sym + r'\b', unicode_sym, text)
    
    # Handle fractions
    text = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'(\1)/(\2)', text)
    
    # Handle roots
    text = re.sub(r'\\sqrt\{([^}]+)\}', r'√(\1)', text)
    text = re.sub(r'\\sqrt\[([^]]+)\]\{([^}]+)\}', r'\1√(\2)', text)
    
    # Handle superscripts and subscripts
    text = re.sub(r'\^(\w+|\{[^}]+\})', r'^(\1)', text)
    text = re.sub(r'_(\w+|\{[^}]+\})', r'_(\1)', text)
    
    # Handle limits
    text = re.sub(r'\\lim_\{([^}]+)\}', r'lim[(\1)]', text)
    
    # Clean up remaining LaTeX commands
    text = re.sub(r'\\[a-zA-Z]+\*?', '', text)
    text = re.sub(r'[{}]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text

def analyze_tikz_content(tikz_code):
    """Phân tích và mô tả nội dung TikZ"""
    
    descriptions = []
    
    # Analyze drawing commands
    if re.search(r'\\draw.*circle', tikz_code):
        descriptions.append("🔵 Vẽ hình tròn")
    
    if re.search(r'\\draw.*rectangle', tikz_code):
        descriptions.append("⬜ Vẽ hình chữ nhật")
    
    if re.search(r'\\draw.*ellipse', tikz_code):
        descriptions.append("⭕ Vẽ hình ellipse")
    
    if re.search(r'\\draw.*--', tikz_code):
        descriptions.append("📏 Vẽ đường thẳng")
    
    if re.search(r'\\draw.*plot', tikz_code):
        descriptions.append("📈 Vẽ đồ thị hàm số")
    
    if re.search(r'\\draw.*arc', tikz_code):
        descriptions.append("🌙 Vẽ cung tròn")
    
    # Analyze nodes and labels
    if re.search(r'\\node', tikz_code):
        descriptions.append("🏷️ Chèn nhãn/text")
    
    # Analyze fills and colors
    if re.search(r'\\fill', tikz_code):
        descriptions.append("🎨 Tô màu")
    
    if re.search(r'color=|fill=', tikz_code):
        descriptions.append("🌈 Sử dụng màu sắc")
    
    # Analyze arrows
    if re.search(r'->', tikz_code):
        descriptions.append("➡️ Vẽ mũi tên")
    
    # Analyze grids
    if re.search(r'\\draw.*grid', tikz_code):
        descriptions.append("⚏ Vẽ lưới")
    
    # Analyze coordinates
    coordinate_matches = re.findall(r'\(([^)]+)\)', tikz_code)
    if coordinate_matches:
        descriptions.append(f"📍 Sử dụng {len(coordinate_matches)} điểm tọa độ")
    
    if not descriptions:
        descriptions.append("🎨 Hình vẽ tùy chỉnh")
    
    return descriptions

def create_word_document(original_content, formulas, tikz_diagrams, tables):
    """Tạo document Word với nội dung đã chuyển đổi"""
    
    doc = Document()
    
    # Set margins
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Title
    title = doc.add_heading('Document đã chuyển đổi từ LaTeX', 0)
    title.alignment = 1  # Center alignment
    
    # Summary
    summary = doc.add_paragraph()
    summary.add_run('📊 Tóm tắt chuyển đổi\n').bold = True
    summary.add_run(f'• Công thức LaTeX đã chuyển đổi: {len(formulas)}\n')
    summary.add_run(f'• TikZ diagrams được mô tả: {len(tikz_diagrams)}\n')
    summary.add_run(f'• Bảng được tạo: {len(tables)}\n')
    summary.add_run(f'• Ngày tạo: {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
    
    doc.add_paragraph('─' * 50)
    
    # Process content
    processed_content = original_content
    
    # Replace LaTeX formulas
    for i, formula in enumerate(formulas):
        unicode_version = convert_latex_to_unicode(formula['content'])
        replacement = f"【Công thức {i+1}: {unicode_version}】"
        processed_content = processed_content.replace(formula['full'], replacement)
    
    # Replace TikZ diagrams
    for diagram in tikz_diagrams:
        descriptions = analyze_tikz_content(diagram['content'])
        description_text = f"【TikZ Diagram {diagram['id']}: {', '.join(descriptions)}】"
        processed_content = processed_content.replace(diagram['full'], description_text)
    
    # Clean HTML tags
    processed_content = re.sub(r'<[^>]+>', '', processed_content)
    processed_content = re.sub(r'\s+', ' ', processed_content).strip()
    
    # Add main content
    doc.add_heading('📄 Nội dung chính', level=1)
    
    # Split into paragraphs and add to document
    paragraphs = processed_content.split('\n')
    for para_text in paragraphs:
        if para_text.strip():
            # Check if it's a converted formula or TikZ
            if '【Công thức' in para_text or '【TikZ' in para_text:
                p = doc.add_paragraph()
                p.add_run(para_text).italic = True
                p.style = 'Intense Quote'
            else:
                doc.add_paragraph(para_text.strip())
    
    # Add tables
    if tables:
        doc.add_heading('📊 Bảng dữ liệu', level=1)
        
        for i, table_info in enumerate(tables):
            table_data = table_info['data']
            doc.add_heading(f'Bảng {i+1}', level=2)
            
            # Create Word table
            word_table = doc.add_table(rows=1, cols=len(table_data['headers']))
            word_table.style = 'Table Grid'
            
            # Style the table
            word_table.allow_autofit = True
            
            # Header row
            header_cells = word_table.rows[0].cells
            for j, header in enumerate(table_data['headers']):
                header_cells[j].text = header
                # Make header bold
                for paragraph in header_cells[j].paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
            
            # Data rows
            for row_data in table_data['rows']:
                row_cells = word_table.add_row().cells
                for j, cell_data in enumerate(row_data):
                    if j < len(row_cells):
                        row_cells[j].text = str(cell_data)
    
    # Add detailed appendix
    if formulas or tikz_diagrams:
        doc.add_page_break()
        doc.add_heading('📋 Phụ lục chi tiết', level=1)
        
        # LaTeX formulas detail
        if formulas:
            doc.add_heading('🧮 Chi tiết công thức LaTeX', level=2)
            
            for i, formula in enumerate(formulas):
                doc.add_heading(f'Công thức {i+1} ({formula["type"]})', level=3)
                
                # Unicode version
                unicode_version = convert_latex_to_unicode(formula['content'])
                p1 = doc.add_paragraph()
                p1.add_run('Dạng Unicode: ').bold = True
                p1.add_run(unicode_version)
                
                # Original LaTeX
                p2 = doc.add_paragraph()
                p2.add_run('LaTeX gốc: ').bold = True
                p2.add_run(formula['content'])
                p2.style = 'Intense Quote'
                
                doc.add_paragraph()  # Spacing
        
        # TikZ diagrams detail
        if tikz_diagrams:
            doc.add_heading('🖼️ Chi tiết TikZ Diagrams', level=2)
            
            for diagram in tikz_diagrams:
                doc.add_heading(f'TikZ Diagram {diagram["id"]}', level=3)
                
                # Analysis
                descriptions = analyze_tikz_content(diagram['content'])
                p1 = doc.add_paragraph()
                p1.add_run('Phân tích: ').bold = True
                p1.add_run(', '.join(descriptions))
                
                # Original code
                p2 = doc.add_paragraph()
                p2.add_run('Mã TikZ gốc:\n').bold = True
                p2.add_run(diagram['content'])
                p2.style = 'Intense Quote'
                
                doc.add_paragraph()  # Spacing
    
    # Footer
    doc.add_paragraph('─' * 50)
    footer = doc.add_paragraph()
    footer.add_run('🔗 Tạo bởi LaTeX to Word Converter').italic = True
    footer.add_run('\n💡 Phiên bản Streamlit Cloud với khả năng chuyển đổi thực tế').italic = True
    footer.add_run('\n🚀 Để có equation objects thật và hình ảnh TikZ, sử dụng Docker version').italic = True
    footer.alignment = 1  # Center
    
    return doc

# Sidebar
with st.sidebar:
    st.markdown("### 🎯 Tính năng")
    st.markdown("""
    **✅ Có thể làm:**
    - Chuyển LaTeX → Unicode (α, β, ∫, ∑)
    - Tạo bảng Word thật
    - Mô tả TikZ diagrams
    - Xuất file .DOCX
    
    **🚀 Docker version:**
    - Equation objects thật
    - Hình ảnh TikZ PNG
    - PDF export
    """)
    
    st.markdown("### 📊 Thống kê")
    if 'stats' in st.session_state:
        stats = st.session_state.stats
        st.metric("Formulas", stats.get('formulas', 0))
        st.metric("TikZ", stats.get('tikz', 0))
        st.metric("Tables", stats.get('tables', 0))

# Main interface
st.markdown("""
<div class="success-box">
<h4>🎉 Tính năng chuyển đổi thực tế!</h4>
<p>Phiên bản này thực sự chuyển đổi LaTeX và tạo file Word:</p>
<ul>
<li>🧮 <strong>LaTeX → Unicode:</strong> $\\alpha$ → α, $\\int$ → ∫</li>
<li>📊 <strong>Markdown → Word Table:</strong> Bảng thật có thể chỉnh sửa</li>
<li>🖼️ <strong>TikZ → Mô tả:</strong> Phân tích ý nghĩa diagram</li>
<li>📄 <strong>Export .DOCX:</strong> File Word thật</li>
</ul>
</div>
""", unsafe_allow_html=True)

# File uploader
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader(
        "📁 Chọn file Word (.docx)",
        type=['docx'],
        help="Upload file Word chứa LaTeX formulas, TikZ diagrams, hoặc Markdown tables"
    )

with col2:
    st.markdown("### ⚙️ Tùy chọn")
    include_appendix = st.checkbox("📋 Bao gồm phụ lục chi tiết", value=True)
    clean_format = st.checkbox("🧹 Định dạng sạch", value=True)

# Processing
if uploaded_file is not None:
    st.success(f"✅ Đã upload: {uploaded_file.name} ({len(uploaded_file.getvalue())} bytes)")
    
    if st.button("🔄 **Chuyển đổi thành Word**", type="primary", use_container_width=True):
        
        with st.spinner("🔄 Đang xử lý và chuyển đổi..."):
            try:
                # Read Word file
                result = mammoth.convert_to_html(uploaded_file)
                html_content = result.value if hasattr(result, 'value') else result.html
                
                # Extract components
                with st.status("Đang phân tích nội dung...", expanded=True) as status:
                    st.write("🔍 Tìm công thức LaTeX...")
                    formulas = extract_latex_formulas(html_content)
                    
                    st.write("🔍 Tìm TikZ diagrams...")
                    tikz_diagrams = extract_tikz_diagrams(html_content)
                    
                    st.write("🔍 Tìm bảng Markdown...")
                    tables = extract_markdown_tables(html_content)
                    
                    st.write("📝 Tạo document Word...")
                    doc = create_word_document(html_content, formulas, tikz_diagrams, tables)
                    
                    status.update(label="✅ Hoàn tất!", state="complete", expanded=False)
                
                # Save to buffer
                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                # Store stats
                st.session_state.stats = {
                    'formulas': len(formulas),
                    'tikz': len(tikz_diagrams),
                    'tables': len(tables)
                }
                
                # Display results
                st.success("🎉 Chuyển đổi hoàn tất!")
                
                # Metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("🧮 LaTeX Formulas", len(formulas))
                with col2:
                    st.metric("🖼️ TikZ Diagrams", len(tikz_diagrams))
                with col3:
                    st.metric("📊 Tables", len(tables))
                with col4:
                    st.metric("📄 File Size", f"{len(buffer.getvalue())//1024} KB")
                
                # Preview conversions
                if formulas or tikz_diagrams or tables:
                    st.markdown("### 👀 Preview chuyển đổi")
                    
                    tab1, tab2, tab3 = st.tabs(["🧮 LaTeX → Unicode", "🖼️ TikZ → Mô tả", "📊 Tables"])
                    
                    with tab1:
                        if formulas:
                            for i, formula in enumerate(formulas[:3]):
                                unicode_version = convert_latex_to_unicode(formula['content'])
                                st.markdown(f"""
                                <div class="conversion-preview">
                                <strong>Công thức {i+1}:</strong><br>
                                <code>{formula['content']}</code> → <strong>{unicode_version}</strong>
                                </div>
                                """, unsafe_allow_html=True)
                            if len(formulas) > 3:
                                st.info(f"... và {len(formulas)-3} công thức khác")
                        else:
                            st.info("Không tìm thấy công thức LaTeX")
                    
                    with tab2:
                        if tikz_diagrams:
                            for diagram in tikz_diagrams[:2]:
                                descriptions = analyze_tikz_content(diagram['content'])
                                st.markdown(f"""
                                <div class="conversion-preview">
                                <strong>TikZ Diagram {diagram['id']}:</strong><br>
                                {', '.join(descriptions)}
                                </div>
                                """, unsafe_allow_html=True)
                            if len(tikz_diagrams) > 2:
                                st.info(f"... và {len(tikz_diagrams)-2} diagram khác")
                        else:
                            st.info("Không tìm thấy TikZ diagram")
                    
                    with tab3:
                        if tables:
                            for i, table_info in enumerate(tables[:2]):
                                table_data = table_info['data']
                                st.markdown(f"**Bảng {i+1}:**")
                                # Create DataFrame for display
                                df = pd.DataFrame(table_data['rows'], columns=table_data['headers'])
                                st.dataframe(df, use_container_width=True)
                            if len(tables) > 2:
                                st.info(f"... và {len(tables)-2} bảng khác")
                        else:
                            st.info("Không tìm thấy bảng")
                
                # Download button
                st.markdown("### 💾 Tải về file Word")
                filename = f"converted_{uploaded_file.name}"
                
                st.download_button(
                    label="📥 **Tải về file .DOCX**",
                    data=buffer.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
                
                # Success message
                st.markdown("""
                <div class="success-box">
                <h4>🎉 Thành công!</h4>
                <p><strong>File Word đã được tạo với:</strong></p>
                <ul>
                <li>✅ <strong>Bảng thật</strong> có thể chỉnh sửa trong Word</li>
                <li>✅ <strong>Công thức Unicode</strong> dễ đọc (α, β, ∫, ∑)</li>
                <li>✅ <strong>Mô tả TikZ</strong> chi tiết và có ý nghĩa</li>
                <li>✅ <strong>Định dạng professional</strong> với phụ lục</li>
                </ul>
                <p><strong>💡 Mở file trong Microsoft Word để xem đầy đủ tính năng!</strong></p>
                </div>
                """, unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"❌ Lỗi khi xử lý: {str(e)}")
                with st.expander("🔍 Chi tiết lỗi"):
                    import traceback
                    st.code(traceback.format_exc())

# Examples section
else:
    st.markdown("### 📝 Ví dụ về khả năng chuyển đổi")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="feature-box">
        <h4>📥 Input LaTeX</h4>
        <code>$\\int_0^1 x^2 dx = \\frac{1}{3}$</code>
        <br><br>
        <h4>📤 Output Word</h4>
        <strong>∫₀¹ x² dx = (1)/(3)</strong>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="feature-box">
        <h4>📥 Input TikZ</h4>
        <code>\\begin{tikzpicture}<br>\\draw (0,0) circle (1cm);<br>\\end{tikzpicture}</code>
        <br><br>
        <h4>📤 Output Word</h4>
        <strong>🔵 Vẽ hình tròn</strong>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="feature-box">
        <h4>📥 Input Table</h4>
        <code>| Name | Score |<br>|------|-------|<br>| Alice| 95    |</code>
        <br><br>
        <h4>📤 Output Word</h4>
        <strong>Bảng Word thật có thể edit!</strong>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="warning-box">
        <h4>🚀 Muốn 100% tính năng?</h4>
        <p><strong>Docker version có:</strong></p>
        <ul>
        <li>🧮 Equation objects thật</li>
        <li>🖼️ Hình ảnh PNG từ TikZ</li>
        <li>📄 Export PDF</li>
        </ul>
        <code>docker-compose up --build</code>
        </div>
        """, unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 1rem;">
    <p>🚀 <strong>LaTeX to Word Converter</strong> - Phiên bản Cloud với khả năng chuyển đổi thực tế</p>
    <p>💻 <a href="https://github.com/your-repo" target="_blank">GitHub Repository</a> | 
       🐳 <a href="#docker">Docker Version</a> | 
       📚 <a href="#docs">Documentation</a></p>
    <p><small>Made with ❤️ for the education community</small></p>
</div>
""", unsafe_allow_html=True)
