import streamlit as st
import mammoth
import re
import os
import tempfile
from pathlib import Path
import base64
from io import StringIO
import pandas as pd

# Cấu hình trang
st.set_page_config(
    page_title="LaTeX to Word Converter - Demo",
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
    .demo-box {
        background: #d1ecf1;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #17a2b8;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Header chính
st.markdown("""
<div class="main-header">
    <h1>🧮 LaTeX to Word Converter - Demo Version</h1>
    <p>Phân tích và hiển thị nội dung LaTeX, TikZ từ file Word</p>
</div>
""", unsafe_allow_html=True)

# Warning về limitation
st.markdown("""
<div class="warning-box">
    <h4>⚠️ Lưu ý về phiên bản Demo</h4>
    <p>Streamlit Cloud không hỗ trợ system dependencies như <code>pandoc</code>, <code>pdflatex</code>, <code>imagemagick</code>. 
    Phiên bản này chỉ có thể:</p>
    <ul>
        <li>✅ Đọc và phân tích file Word</li>
        <li>✅ Phát hiện công thức LaTeX và TikZ</li>
        <li>✅ Hiển thị thống kê chi tiết</li>
        <li>✅ Xuất file HTML</li>
        <li>❌ Không thể compile TikZ thành hình ảnh</li>
        <li>❌ Không thể tạo file Word với Equation thật</li>
    </ul>
    <p><strong>Để sử dụng đầy đủ tính năng, hãy chạy local hoặc deploy lên VPS.</strong></p>
</div>
""", unsafe_allow_html=True)

# Sidebar với thông tin
with st.sidebar:
    st.header("📋 Hướng dẫn sử dụng")
    st.markdown("""
    **Phiên bản Demo này giúp bạn:**
    - ✅ Upload và phân tích file Word
    - ✅ Phát hiện LaTeX và TikZ
    - ✅ Xem thống kê chi tiết
    - ✅ Xuất HTML để xem trước
    
    **Để có đầy đủ tính năng:**
    1. Clone repo về local
    2. Cài đặt system dependencies
    3. Chạy `streamlit run app.py`
    """)
    
    st.header("🚀 Deploy đầy đủ")
    st.markdown("""
    **Các option deploy:**
    - 🖥️ **Local**: Đầy đủ tính năng
    - ☁️ **VPS/Cloud**: Cần cài dependencies
    - 🐳 **Docker**: Tự động setup môi trường
    """)

# Main content
col1, col2 = st.columns([1, 1])

with col1:
    st.header("📁 Tải lên file Word")
    uploaded_file = st.file_uploader(
        "Chọn file Word (.docx) chứa LaTeX, TikZ, bảng",
        type=['docx'],
        help="File sẽ được phân tích để tìm LaTeX, TikZ và bảng"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ Đã tải lên: {uploaded_file.name}")
        st.info(f"📊 Kích thước: {len(uploaded_file.getvalue())} bytes")

with col2:
    st.header("⚙️ Tùy chọn phân tích")
    
    show_html = st.checkbox("📄 Hiển thị HTML", value=True)
    show_latex = st.checkbox("🧮 Highlight LaTeX", value=True)  
    show_tikz = st.checkbox("🖼️ Highlight TikZ", value=True)
    show_tables = st.checkbox("📊 Highlight Tables", value=True)

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

# Xử lý file được upload
if uploaded_file is not None:
    if st.button("🔍 Phân tích file", type="primary"):
        
        with st.spinner("⏳ Đang phân tích file..."):
            try:
                # Tạo thư mục tạm thời
                temp_dir = tempfile.mkdtemp()
                input_path = os.path.join(temp_dir, "input.docx")
                
                # Lưu file
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getvalue())
                
                # Debug info
                st.write(f"🔍 Debug: File saved to {input_path}")
                st.write(f"📏 File size: {os.path.getsize(input_path)} bytes")
                
                # Đọc nội dung bằng mammoth
                st.write("📖 Đang đọc nội dung từ file Word...")
                
                # Test mammoth version
                try:
                    import mammoth
                    st.write(f"📦 Mammoth version: {mammoth.__version__}")
                except:
                    st.write("⚠️ Cannot get mammoth version")
                
                with open(input_path, "rb") as docx_file:
                    result = mammoth.convert_to_html(docx_file)
                    
                    # Debug result object
                    st.write(f"🔍 Result type: {type(result)}")
                    st.write(f"🔍 Result attributes: {[attr for attr in dir(result) if not attr.startswith('_')]}")
                    
                    # Try different attributes
                    if hasattr(result, 'value'):
                        html_content = result.value
                        st.write("✅ Using result.value")
                    elif hasattr(result, 'html'):
                        html_content = result.html  
                        st.write("✅ Using result.html")
                    else:
                        st.error("❌ Cannot find HTML content in result object")
                        st.write("Available attributes:", dir(result))
                        return
                    
                    # Try different message attributes  
                    warnings = []
                    if hasattr(result, 'messages'):
                        warnings = result.messages
                    elif hasattr(result, 'warnings'):
                        warnings = result.warnings
                    
                    st.write(f"📄 HTML content length: {len(html_content)}")
                    st.write(f"⚠️ Warnings count: {len(warnings)}")
                
                # Phân tích nội dung
                st.write("🔍 Đang phân tích LaTeX, TikZ và bảng...")
                formulas = extract_latex_formulas(html_content)
                tikz_diagrams = extract_tikz_diagrams(html_content)
                tables = extract_tables(html_content)
                
                st.success("✅ Phân tích hoàn tất!")
                
                # Hiển thị thống kê
                st.header("📊 Thống kê phân tích")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("🧮 Công thức LaTeX", len(formulas))
                
                with col2:
                    st.metric("🖼️ TikZ Diagrams", len(tikz_diagrams))
                
                with col3:
                    st.metric("📊 Bảng HTML", len(tables['html']))
                
                with col4:
                    st.metric("📋 Bảng Markdown", len(tables['markdown']))
                
                # Hiển thị chi tiết
                if formulas or tikz_diagrams or tables['html'] or tables['markdown']:
                    
                    # Tab hiển thị chi tiết
                    tab1, tab2, tab3, tab4 = st.tabs(["🧮 LaTeX Formulas", "🖼️ TikZ Diagrams", "📊 Tables", "📄 Full Content"])
                    
                    with tab1:
                        if formulas:
                            st.subheader(f"Tìm thấy {len(formulas)} công thức LaTeX:")
                            for i, formula in enumerate(formulas, 1):
                                with st.expander(f"Công thức {i} ({formula['type']})"):
                                    st.code(formula['content'], language='latex')
                                    st.write(f"**Vị trí:** {formula['start']}-{formula['end']}")
                        else:
                            st.info("Không tìm thấy công thức LaTeX nào.")
                    
                    with tab2:
                        if tikz_diagrams:
                            st.subheader(f"Tìm thấy {len(tikz_diagrams)} TikZ diagrams:")
                            for diagram in tikz_diagrams:
                                with st.expander(f"TikZ Diagram {diagram['id']}"):
                                    st.code(diagram['content'], language='latex')
                                    st.info("💡 Để tạo hình ảnh từ TikZ, cần chạy ứng dụng local với pdflatex.")
                        else:
                            st.info("Không tìm thấy TikZ diagram nào.")
                    
                    with tab3:
                        if tables['html'] or tables['markdown']:
                            if tables['html']:
                                st.subheader(f"Tìm thấy {len(tables['html'])} bảng HTML:")
                                for i, table in enumerate(tables['html'], 1):
                                    with st.expander(f"Bảng HTML {i}"):
                                        st.markdown(table, unsafe_allow_html=True)
                            
                            if tables['markdown']:
                                st.subheader(f"Tìm thấy {len(tables['markdown'])} bảng Markdown:")
                                for i, table in enumerate(tables['markdown'], 1):
                                    with st.expander(f"Bảng Markdown {i}"):
                                        st.code(table, language='markdown')
                        else:
                            st.info("Không tìm thấy bảng nào.")
                    
                    with tab4:
                        st.subheader("Nội dung đầy đủ với highlighting:")
                        
                        # Highlight content
                        highlighted_content = highlight_content(
                            html_content, formulas, tikz_diagrams, tables, 
                            show_latex, show_tikz, show_tables
                        )
                        
                        if show_html:
                            st.components.v1.html(highlighted_content, height=600, scrolling=True)
                        else:
                            st.code(highlighted_content[:2000] + "..." if len(highlighted_content) > 2000 else highlighted_content, language='html')
                
                # Xuất file HTML
                st.header("💾 Xuất kết quả")
                
                # Tạo HTML file
                html_with_highlights = highlight_content(
                    html_content, formulas, tikz_diagrams, tables, True, True, True
                )
                
                html_template = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>LaTeX Analysis Result</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        .stats {{ background: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
        .highlight-latex {{ background-color: #ffeb3b; padding: 2px; }}
        .highlight-tikz {{ background-color: #4caf50; padding: 2px; color: white; }}
    </style>
</head>
<body>
    <h1>📊 Kết quả phân tích LaTeX</h1>
    <div class="stats">
        <p><strong>File:</strong> {uploaded_file.name}</p>
        <p><strong>LaTeX Formulas:</strong> {len(formulas)}</p>
        <p><strong>TikZ Diagrams:</strong> {len(tikz_diagrams)}</p>
        <p><strong>HTML Tables:</strong> {len(tables['html'])}</p>
        <p><strong>Markdown Tables:</strong> {len(tables['markdown'])}</p>
    </div>
    <hr>
    <h2>📄 Nội dung với highlighting</h2>
    {html_with_highlights}
</body>
</html>
                """
                
                st.download_button(
                    label="💾 Tải file HTML phân tích",
                    data=html_template,
                    file_name=f"analysis_{uploaded_file.name.split('.')[0]}.html",
                    mime="text/html",
                    type="primary"
                )
                
                # Thông tin warnings
                if warnings:
                    with st.expander("⚠️ Warnings từ mammoth"):
                        for warning in warnings:
                            st.warning(warning)
                
                # Dọn dẹp
                os.remove(input_path)
                os.rmdir(temp_dir)
                
            except Exception as e:
                st.error(f"❌ Lỗi khi xử lý: {str(e)}")
                st.write("💡 Hãy đảm bảo file Word hợp lệ và không bị corrupt.")

# Demo instructions
if uploaded_file is None:
    st.header("📝 Hướng dẫn sử dụng Demo")
    
    st.markdown("""
    <div class="demo-box">
    <h4>🔍 Demo này có thể làm gì?</h4>
    <ol>
        <li>📁 <strong>Upload file Word</strong> chứa LaTeX, TikZ hoặc bảng</li>
        <li>🔍 <strong>Phân tích tự động</strong> và đếm số lượng mỗi loại</li>
        <li>👁️ <strong>Xem trước chi tiết</strong> từng công thức và diagram</li>
        <li>🎨 <strong>Highlight</strong> các thành phần trong HTML</li>
        <li>💾 <strong>Xuất file HTML</strong> để lưu trữ hoặc chia sẻ</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)
    
    st.header("📝 Ví dụ định dạng được hỗ trợ")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="feature-box">
        <h4>🧮 Công thức LaTeX</h4>
        <code>$x^2 + y^2 = z^2$</code><br>
        <code>$$\\int_0^1 x^2 dx = \\frac{1}{3}$$</code><br>
        <code>\\begin{equation}E = mc^2\\end{equation}</code>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="feature-box">
        <h4>📊 Bảng HTML/Markdown</h4>
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
        <p>Để có đầy đủ tính năng (compile TikZ, tạo Equation Word):</p>
        <code>git clone &lt;repo&gt;<br>
        pip install -r requirements.txt<br>
        # Cài pandoc, texlive, imagemagick<br>
        streamlit run app.py</code>
        </div>
        """, unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666;">
    <p>🚀 Demo Version | 💻 <a href="#" target="_blank">GitHub Repo</a> | 📧 Support</p>
    <p>💡 Để có đầy đủ tính năng, hãy chạy local với đầy đủ dependencies</p>
</div>
""", unsafe_allow_html=True)
