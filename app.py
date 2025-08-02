import streamlit as st
import mammoth
import re
import os
import tempfile
import shutil
from pathlib import Path
import base64

# Cấu hình trang
st.set_page_config(
    page_title="LaTeX to Word Converter",
    page_icon="📝",
    layout="wide"
)

# CSS để làm đẹp giao diện
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .feature-box {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        margin: 1rem 0;
    }
    .warning-box {
        background: #fff3cd;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #ffc107;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Header chính
st.markdown("""
<div class="main-header">
    <h1>🧮 LaTeX to Word Converter</h1>
    <p>Chuyển đổi file Word chứa LaTeX, TikZ và Markdown thành Word hoàn chỉnh</p>
</div>
""", unsafe_allow_html=True)

# Warning về Streamlit Cloud limitations
st.markdown("""
<div class="warning-box">
    <h4>⚠️ Streamlit Cloud Version - Tính năng hạn chế</h4>
    <p>Do Streamlit Cloud không hỗ trợ system dependencies (pandoc, pdflatex, imagemagick), phiên bản này chỉ có thể:</p>
    <ul>
        <li>✅ Phân tích và hiển thị LaTeX formulas</li>
        <li>✅ Phát hiện TikZ diagrams</li>  
        <li>✅ Xử lý bảng Markdown</li>
        <li>✅ Xuất HTML preview</li>
        <li>❌ Không thể compile TikZ thành hình ảnh</li>
        <li>❌ Không thể tạo file Word với Equation thật</li>
    </ul>
    <p><strong>Để có đầy đủ tính năng, hãy chạy local hoặc dùng Docker!</strong></p>
</div>
""", unsafe_allow_html=True)

# Sidebar với thông tin
with st.sidebar:
    st.header("📋 Hướng dẫn sử dụng")
    st.markdown("""
    **Ứng dụng này giúp bạn:**
    - ✅ Phân tích file Word chứa LaTeX
    - ✅ Thống kê LaTeX formulas và TikZ
    - ✅ Preview nội dung với highlighting
    - ✅ Xuất HTML để xem trước
    
    **Để có đầy đủ tính năng:**
    1. Clone repo về local
    2. Cài đặt system dependencies  
    3. Chạy `streamlit run app.py`
    """)
    
    st.header("🔧 System Requirements")
    st.markdown("""
    **Để deploy đầy đủ cần:**
    - Pandoc
    - TeX Live (pdflatex)
    - ImageMagick
    """)

# Utility functions
def extract_latex_formulas(content):
    """Tìm và trích xuất công thức LaTeX"""
    patterns = [
        r'\$+([^$]+)\$+',  # $...$ và $$...$$
        r'\\begin\{equation\}(.*?)\\end\{equation\}',
        r'\\begin\{align\}(.*?)\\end\{align\}',
        r'\\\[(.*?)\\\]'
    ]
    
    formulas = []
    for pattern in patterns:
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            formulas.append({
                'content': match.group(1).strip() if match.groups() else match.group(0),
                'type': 'inline' if '$' in pattern else 'display',
                'start': match.start(),
                'end': match.end(),
                'full': match.group(0)
            })
    
    return formulas

def extract_tikz_diagrams(content):
    """Tìm và trích xuất TikZ diagrams"""
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

def extract_tables(content):
    """Tìm và trích xuất bảng"""
    # HTML tables
    html_tables = re.findall(r'<table.*?</table>', content, re.DOTALL | re.IGNORECASE)
    
    # Markdown-style tables 
    markdown_tables = re.findall(r'\|.*?\|(?:\r?\n\|.*?\|)*', content, re.MULTILINE)
    
    return {
        'html': html_tables,
        'markdown': markdown_tables
    }

def highlight_content(content, formulas, tikz_diagrams, tables, show_latex, show_tikz, show_tables):
    """Highlight các thành phần trong nội dung"""
    highlighted = content
    
    if show_latex:
        for formula in formulas:
            original = formula['full']
            highlighted_formula = f'<mark style="background-color: #ffeb3b; padding: 2px;">{original}</mark>'
            highlighted = highlighted.replace(original, highlighted_formula)
    
    if show_tikz:
        for diagram in tikz_diagrams:
            original = diagram['full']
            highlighted_diagram = f'<mark style="background-color: #4caf50; padding: 2px; color: white;">{original}</mark>'
            highlighted = highlighted.replace(original, highlighted_diagram)
    
    if show_tables:
        for table in tables['html']:
            highlighted_table = f'<div style="border: 2px solid #2196f3; padding: 5px; margin: 5px;">{table}</div>'
            highlighted = highlighted.replace(table, highlighted_table)
    
    return highlighted

# Main content
col1, col2 = st.columns([1, 1])

with col1:
    st.header("📁 Tải lên file Word")
    uploaded_file = st.file_uploader(
        "Chọn file Word (.docx) chứa LaTeX, TikZ, bảng",
        type=['docx'],
        help="File nên chứa công thức LaTeX ($...$), TikZ, hoặc bảng Markdown"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ Đã tải lên: {uploaded_file.name}")
        st.info(f"📊 Kích thước: {len(uploaded_file.getvalue())} bytes")

with col2:
    st.header("⚙️ Tùy chọn xử lý")
    
    show_html = st.checkbox("📄 Hiển thị HTML", value=True)
    show_latex = st.checkbox("🧮 Highlight LaTeX", value=True)  
    show_tikz = st.checkbox("🖼️ Highlight TikZ", value=True)
    show_tables = st.checkbox("📊 Highlight Tables", value=True)
    
    output_format = st.selectbox(
        "📄 Định dạng đầu ra",
        ["html", "text"],
        help="Chỉ hỗ trợ HTML và text trong phiên bản Cloud"
    )

# Nút xử lý
if uploaded_file is not None:
    if st.button("🚀 Bắt đầu xử lý", type="primary"):
        
        with st.spinner("⏳ Đang xử lý file..."):
            try:
                # Tạo thư mục tạm thời
                temp_dir = tempfile.mkdtemp()
                input_path = os.path.join(temp_dir, "input.docx")
                
                # Lưu file tải lên
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getvalue())
                
                # Bước 1: Đọc nội dung từ Word
                st.write("📖 Đang đọc nội dung từ file Word...")
                with open(input_path, "rb") as docx_file:
                    result = mammoth.convert_to_html(docx_file)
                    
                    # Handle different mammoth API versions
                    if hasattr(result, 'value'):
                        html_content = result.value
                    elif hasattr(result, 'html'):
                        html_content = result.html
                    else:
                        st.error("❌ Không thể đọc nội dung từ file Word")
                        st.stop()
                    
                    # Get messages/warnings
                    warnings = []
                    if hasattr(result, 'messages'):
                        warnings = result.messages
                    elif hasattr(result, 'warnings'):
                        warnings = result.warnings
                
                # Bước 2: Xử lý nội dung
                st.write("🔍 Đang phân tích LaTeX, TikZ và bảng...")
                
                formulas = extract_latex_formulas(html_content)
                tikz_diagrams = extract_tikz_diagrams(html_content)
                tables = extract_tables(html_content)
                
                # Highlight content
                processed_content = highlight_content(
                    html_content, formulas, tikz_diagrams, tables,
                    show_latex, show_tikz, show_tables
                )
                
                st.success("✅ Xử lý hoàn tất!")
                
                # Hiển thị kết quả
                st.header("📋 Xem trước kết quả")
                
                # Tab để hiển thị nội dung
                tab1, tab2, tab3 = st.tabs(["📊 Thống kê", "🔍 Chi tiết", "📄 Nội dung"])
                
                with tab1:
                    # Thống kê xử lý
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("🧮 Công thức LaTeX", len(formulas))
                    
                    with col2:
                        st.metric("🖼️ TikZ diagrams", len(tikz_diagrams))
                    
                    with col3:
                        st.metric("📊 Bảng HTML", len(tables['html']))
                    
                    with col4:
                        st.metric("📋 Bảng Markdown", len(tables['markdown']))
                
                with tab2:
                    # Chi tiết các thành phần
                    if formulas:
                        st.subheader(f"🧮 {len(formulas)} Công thức LaTeX:")
                        for i, formula in enumerate(formulas, 1):
                            with st.expander(f"Công thức {i} ({formula['type']})"):
                                st.code(formula['content'], language='latex')
                                st.write(f"**Vị trí:** {formula['start']}-{formula['end']}")
                    
                    if tikz_diagrams:
                        st.subheader(f"🖼️ {len(tikz_diagrams)} TikZ Diagrams:")
                        for diagram in tikz_diagrams:
                            with st.expander(f"TikZ Diagram {diagram['id']}"):
                                st.code(diagram['content'], language='latex')
                                st.info("💡 Để tạo hình ảnh từ TikZ, cần chạy ứng dụng local với pdflatex.")
                    
                    if tables['html'] or tables['markdown']:
                        if tables['html']:
                            st.subheader(f"📊 {len(tables['html'])} Bảng HTML:")
                            for i, table in enumerate(tables['html'], 1):
                                with st.expander(f"Bảng HTML {i}"):
                                    st.markdown(table, unsafe_allow_html=True)
                        
                        if tables['markdown']:
                            st.subheader(f"📋 {len(tables['markdown'])} Bảng Markdown:")
                            for i, table in enumerate(tables['markdown'], 1):
                                with st.expander(f"Bảng Markdown {i}"):
                                    st.code(table, language='markdown')
                
                with tab3:
                    # Hiển thị nội dung
                    if output_format == "html" and show_html:
                        st.subheader("📄 HTML với highlighting:")
                        st.components.v1.html(processed_content, height=600, scrolling=True)
                    else:
                        st.subheader("📄 Nội dung text:")
                        preview_length = 2000
                        if len(processed_content) > preview_length:
                            st.code(processed_content[:preview_length] + "\n\n... (truncated)", language='html')
                        else:
                            st.code(processed_content, language='html')
                
                # Nút tải về
                st.header("💾 Tải về kết quả")
                
                # Tạo HTML file với thống kê
                html_template = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>LaTeX Analysis Result - {uploaded_file.name}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }}
        .header {{ background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; text-align: center; }}
        .stats {{ background: #f5f5f5; padding: 15px; border-radius: 5px; margin: 20px 0; }}
        .stat-item {{ display: inline-block; margin: 10px; padding: 10px; background: white; border-radius: 5px; min-width: 100px; text-align: center; }}
        .content {{ margin: 20px 0; }}
        mark {{ padding: 2px 4px; border-radius: 2px; }}
        .latex-highlight {{ background-color: #ffeb3b; }}
        .tikz-highlight {{ background-color: #4caf50; color: white; }}
        .table-highlight {{ border: 2px solid #2196f3; padding: 5px; margin: 5px; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Kết quả phân tích LaTeX</h1>
        <p>File: {uploaded_file.name}</p>
    </div>
    
    <div class="stats">
        <h2>📈 Thống kê:</h2>
        <div class="stat-item">
            <strong>{len(formulas)}</strong><br>
            🧮 LaTeX Formulas
        </div>
        <div class="stat-item">
            <strong>{len(tikz_diagrams)}</strong><br>
            🖼️ TikZ Diagrams
        </div>
        <div class="stat-item">
            <strong>{len(tables['html'])}</strong><br>
            📊 HTML Tables
        </div>
        <div class="stat-item">
            <strong>{len(tables['markdown'])}</strong><br>
            📋 Markdown Tables
        </div>
    </div>
    
    <div class="content">
        <h2>📄 Nội dung với highlighting:</h2>
        {processed_content}
    </div>
    
    <footer style="margin-top: 40px; text-align: center; color: #666; border-top: 1px solid #eee; padding-top: 20px;">
        <p>Generated by LaTeX to Word Converter | Streamlit Cloud Version</p>
        <p>For full features, run locally with Docker</p>
    </footer>
</body>
</html>
                """
                
                st.download_button(
                    label=f"💾 Tải file HTML phân tích",
                    data=html_template,
                    file_name=f"latex_analysis_{uploaded_file.name.split('.')[0]}.html",
                    mime="text/html",
                    type="primary"
                )
                
                # Hiển thị warnings nếu có
                if warnings:
                    with st.expander("⚠️ Warnings từ mammoth"):
                        for warning in warnings:
                            st.warning(str(warning))
                
                # Dọn dẹp file tạm thời
                shutil.rmtree(temp_dir)
                
            except Exception as e:
                st.error(f"❌ Lỗi khi xử lý: {str(e)}")
                st.write("💡 Hãy đảm bảo:")
                st.write("- File Word hợp lệ (.docx)")
                st.write("- File không bị corrupt")
                st.write("- Thử với file Word đơn giản trước")
                
                # Debug info
                import traceback
                st.code(traceback.format_exc())

# Footer và ví dụ
if uploaded_file is None:
    st.header("📝 Ví dụ định dạng đầu vào")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="feature-box">
        <h4>🧮 Công thức LaTeX</h4>
        <code>Phương trình bậc hai: $ax^2 + bx + c = 0$</code><br>
        <code>Tích phân: $\\int_0^1 x^2 dx = \\frac{1}{3}$</code><br>
        <code>\\begin{equation}E = mc^2\\end{equation}</code>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="feature-box">
        <h4>📊 Bảng</h4>
        <code>&lt;table&gt;&lt;tr&gt;&lt;td&gt;Data&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;</code><br>
        <code>| A | B |<br>|---|---|<br>| 1 | 2 |</code>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="feature-box">
        <h4>🖼️ TikZ Diagram</h4>
        <code>\\begin{tikzpicture}<br>
        \\draw (0,0) circle (1cm);<br>
        \\draw (0,0) -- (1,0);<br>
        \\end{tikzpicture}</code>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="feature-box">
        <h4>🚀 Deploy đầy đủ</h4>
        <p>Để có đầy đủ tính năng:</p>
        <code>git clone &lt;repo&gt;<br>
        docker-compose up --build<br>
        # Hoặc chạy local với pandoc, texlive</code>
        </div>
        """, unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666;">
    <p>🚀 Streamlit Cloud Demo Version | 💻 <a href="https://github.com/your-repo" target="_blank">GitHub Repo</a></p>
    <p>💡 Để có đầy đủ tính năng (compile TikZ, tạo Word Equations), hãy chạy local hoặc dùng Docker</p>
</div>
""", unsafe_allow_html=True)
