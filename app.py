import streamlit as st
import mammoth
import pypandoc
import os
import tempfile
import shutil
import re
from pathlib import Path
import base64
from tikz_utils import process_tikz_code
from convert_utils import process_latex_formulas, process_markdown_tables, get_math_statistics

# Cấu hình trang
st.set_page_config(
    page_title="LaTeX to Word Converter - Full Version",
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
    .success-box {
        background: #d4edda;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
    .step-box {
        background: #e8f4fd;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Header chính
st.markdown("""
<div class="main-header">
    <h1>🧮 LaTeX to Word Converter - Full Version</h1>
    <p>Chuyển đổi file Word chứa LaTeX, TikZ và Markdown thành Word hoàn chỉnh với Equation thật</p>
</div>
""", unsafe_allow_html=True)

# Check system dependencies
def check_system_dependencies():
    """Kiểm tra các dependencies cần thiết"""
    deps = {
        'pandoc': False,
        'pdflatex': False, 
        'convert': False
    }
    
    for cmd in deps.keys():
        try:
            import subprocess
            result = subprocess.run([cmd, '--version'], capture_output=True, timeout=5)
            deps[cmd] = result.returncode == 0
        except:
            deps[cmd] = False
    
    return deps

# Status của dependencies
with st.sidebar:
    st.header("🔧 System Status")
    deps = check_system_dependencies()
    
    for tool, available in deps.items():
        if available:
            st.success(f"✅ {tool}")
        else:
            st.error(f"❌ {tool}")
    
    all_available = all(deps.values())
    
    if all_available:
        st.markdown("""
        <div class="success-box">
        <h4>🎉 Đầy đủ tính năng!</h4>
        <p>Tất cả system dependencies đã sẵn sàng</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="feature-box">
        <h4>⚠️ Thiếu dependencies</h4>
        <p>Cần cài đặt:</p>
        <code>sudo apt install pandoc texlive-latex-base imagemagick</code>
        </div>
        """, unsafe_allow_html=True)

# Main sidebar
with st.sidebar:
    st.header("📋 Hướng dẫn sử dụng")
    st.markdown("""
    **Phiên bản đầy đủ này có thể:**
    - ✅ Chuyển công thức LaTeX `$...$` thành Equation thật
    - ✅ Compile TikZ thành hình ảnh PNG chèn vào Word
    - ✅ Chuyển bảng Markdown thành bảng Word chuẩn
    - ✅ Xuất file Word (.docx) hoàn chỉnh
    - ✅ Xuất PDF, HTML
    
    **Quy trình:**
    1. Tải lên file Word (.docx)
    2. Chọn tùy chọn xử lý
    3. Xem preview kết quả
    4. Tải về file đã chuyển đổi
    """)

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
        
        # Hiển thị thông tin file
        st.info(f"📊 Kích thước: {len(uploaded_file.getvalue())} bytes")

with col2:
    st.header("⚙️ Tùy chọn xử lý")
    
    process_latex = st.checkbox("🧮 Xử lý công thức LaTeX", value=True)
    process_tikz = st.checkbox("🖼️ Compile TikZ thành hình ảnh", value=True)  
    process_tables = st.checkbox("📊 Xử lý bảng Markdown", value=True)
    
    st.divider()
    
    output_format = st.selectbox(
        "📄 Định dạng đầu ra",
        ["docx", "pdf", "html"],
        help="Chọn định dạng file muốn xuất ra"
    )
    
    # Advanced options
    with st.expander("🔧 Tùy chọn nâng cao"):
        tikz_dpi = st.slider("TikZ DPI (chất lượng hình ảnh)", 150, 600, 300)
        include_source = st.checkbox("Bao gồm source LaTeX trong comment", value=False)
        standalone_equations = st.checkbox("Equation riêng biệt (không inline)", value=True)

# Processing function
def process_document(uploaded_file, temp_dir, options):
    """Xử lý document hoàn chỉnh"""
    
    input_path = os.path.join(temp_dir, "input.docx")
    
    # Lưu file upload
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    # Bước 1: Đọc nội dung từ Word
    st.write("📖 Đang đọc nội dung từ file Word...")
    with open(input_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html_content = result.value if hasattr(result, 'value') else result.html
        warnings = result.messages if hasattr(result, 'messages') else result.warnings
    
    # Bước 2: Xử lý từng thành phần
    processed_content = html_content
    
    if options['process_latex']:
        st.write("🧮 Đang xử lý công thức LaTeX...")
        processed_content = process_latex_formulas(processed_content)
    
    if options['process_tikz']:
        st.write("🖼️ Đang compile TikZ diagrams...")
        processed_content = process_tikz_code(processed_content, temp_dir, options['tikz_dpi'])
    
    if options['process_tables']:
        st.write("📊 Đang xử lý bảng...")
        processed_content = process_markdown_tables(processed_content)
    
    # Bước 3: Tạo file Markdown tạm thời
    st.write("📝 Đang tạo Markdown tạm thời...")
    
    # Convert HTML to Markdown-like format for pandoc
    markdown_content = html_to_markdown_for_pandoc(processed_content)
    
    markdown_path = os.path.join(temp_dir, "processed.md")
    with open(markdown_path, "w", encoding="utf-8") as f:
        f.write(markdown_content)
    
    # Bước 4: Chuyển đổi bằng pandoc
    st.write(f"📄 Đang chuyển đổi sang {options['output_format'].upper()}...")
    
    output_path = os.path.join(temp_dir, f"output.{options['output_format']}")
    
    if options['output_format'] == "docx":
        # Word output với equation support
        pypandoc.convert_file(
            markdown_path, 
            'docx', 
            outputfile=output_path,
            extra_args=[
                '--standalone',
                '--mathml',  # Convert math to MathML for Word
                '--wrap=preserve'
            ]
        )
    elif options['output_format'] == "pdf":
        pypandoc.convert_file(
            markdown_path,
            'pdf',
            outputfile=output_path,
            extra_args=[
                '--pdf-engine=pdflatex',
                '--standalone'
            ]
        )
    else:  # html
        pypandoc.convert_file(
            markdown_path,
            'html',
            outputfile=output_path,
            extra_args=[
                '--standalone',
                '--mathml',
                '--css=style.css'
            ]
        )
    
    return output_path, processed_content, warnings

def html_to_markdown_for_pandoc(html_content):
    """Convert HTML to Markdown format that pandoc can process with math"""
    
    # Simple HTML to Markdown conversion
    content = html_content
    
    # Convert common HTML tags
    content = re.sub(r'<p>(.*?)</p>', r'\1\n\n', content, flags=re.DOTALL)
    content = re.sub(r'<br/?>', '\n', content)
    content = re.sub(r'<strong>(.*?)</strong>', r'**\1**', content)
    content = re.sub(r'<em>(.*?)</em>', r'*\1*', content)
    content = re.sub(r'<h1>(.*?)</h1>', r'# \1\n', content)
    content = re.sub(r'<h2>(.*?)</h2>', r'## \1\n', content)
    content = re.sub(r'<h3>(.*?)</h3>', r'### \1\n', content)
    
    # Handle images (from TikZ conversion)
    content = re.sub(r'<img src="([^"]+)"[^>]*>', r'![](\1)', content)
    
    # Clean up HTML entities and tags
    content = re.sub(r'<[^>]+>', '', content)  # Remove remaining HTML tags
    content = re.sub(r'&nbsp;', ' ', content)
    content = re.sub(r'&amp;', '&', content)
    content = re.sub(r'&lt;', '<', content)
    content = re.sub(r'&gt;', '>', content)
    
    # Clean up extra whitespace
    content = re.sub(r'\n\s*\n\s*\n', '\n\n', content)
    
    return content.strip()

# Nút xử lý
if uploaded_file is not None:
    if st.button("🚀 Bắt đầu xử lý", type="primary"):
        
        # Check dependencies
        deps = check_system_dependencies()
        missing_deps = [tool for tool, available in deps.items() if not available]
        
        if missing_deps and (process_tikz or output_format != 'html'):
            st.error(f"❌ Thiếu dependencies: {', '.join(missing_deps)}")
            st.info("💡 Để cài đặt:")
            st.code("sudo apt install pandoc texlive-latex-base texlive-pictures imagemagick")
            st.stop()
        
        with st.spinner("⏳ Đang xử lý file..."):
            try:
                # Tạo thư mục tạm thời
                temp_dir = tempfile.mkdtemp()
                
                # Tùy chọn xử lý
                options = {
                    'process_latex': process_latex,
                    'process_tikz': process_tikz,
                    'process_tables': process_tables,
                    'output_format': output_format,
                    'tikz_dpi': tikz_dpi,
                    'include_source': include_source,
                    'standalone_equations': standalone_equations
                }
                
                # Xử lý document
                output_path, processed_content, warnings = process_document(
                    uploaded_file, temp_dir, options
                )
                
                st.success("✅ Xử lý hoàn tất!")
                
                # Hiển thị kết quả
                st.header("📋 Xem trước kết quả")
                
                # Thống kê
                original_html = mammoth.convert_to_html(uploaded_file).value
                stats = get_math_statistics(original_html)
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("🧮 Công thức LaTeX", stats['total_math'])
                
                with col2:
                    st.metric("🖼️ TikZ diagrams", stats['tikz_diagrams'])
                
                with col3:
                    st.metric("📊 Bảng", stats['tables'])
                
                with col4:
                    file_size = os.path.getsize(output_path)
                    st.metric("📄 Kích thước output", f"{file_size/1024:.1f} KB")
                
                # Tabs hiển thị chi tiết
                tab1, tab2, tab3 = st.tabs(["📄 Preview", "📊 Chi tiết", "⚠️ Warnings"])
                
                with tab1:
                    if output_format == "html":
                        # Hiển thị HTML trực tiếp
                        with open(output_path, 'r', encoding='utf-8') as f:
                            html_output = f.read()
                        st.components.v1.html(html_output, height=600, scrolling=True)
                    elif output_format == "docx":
                        st.info("📄 File Word đã được tạo thành công! Tải về để xem.")
                        st.write("**Tính năng trong file Word:**")
                        st.write("- ✅ Công thức LaTeX → Equation thật có thể chỉnh sửa")
                        st.write("- ✅ TikZ → Hình ảnh PNG chất lượng cao") 
                        st.write("- ✅ Bảng Markdown → Bảng Word chuẩn")
                    else:  # PDF
                        st.info("📄 File PDF đã được tạo thành công! Tải về để xem.")
                
                with tab2:
                    st.subheader("📊 Thống kê chi tiết:")
                    st.json(stats)
                    
                    st.subheader("🔧 Tùy chọn đã sử dụng:")
                    st.json(options)
                
                with tab3:
                    if warnings:
                        for warning in warnings:
                            st.warning(str(warning))
                    else:
                        st.success("✅ Không có warning nào!")
                
                # Nút tải về
                with open(output_path, "rb") as file:
                    file_data = file.read()
                    
                mime_types = {
                    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'pdf': 'application/pdf',
                    'html': 'text/html'
                }
                
                st.download_button(
                    label=f"💾 Tải về file {output_format.upper()}",
                    data=file_data,
                    file_name=f"converted_{uploaded_file.name.split('.')[0]}.{output_format}",
                    mime=mime_types[output_format],
                    type="primary"
                )
                
                # Bonus: Tạo comparison view
                if st.checkbox("📊 So sánh trước/sau"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader("📄 Nội dung gốc")
                        original_html = mammoth.convert_to_html(uploaded_file).value
                        st.code(original_html[:1000] + "..." if len(original_html) > 1000 else original_html)
                    
                    with col2:
                        st.subheader("✨ Nội dung đã xử lý")
                        st.code(processed_content[:1000] + "..." if len(processed_content) > 1000 else processed_content)
                
                # Dọn dẹp file tạm thời
                shutil.rmtree(temp_dir)
                
            except Exception as e:
                st.error(f"❌ Lỗi khi xử lý: {str(e)}")
                
                # Debug information
                with st.expander("🔍 Debug info"):
                    import traceback
                    st.code(traceback.format_exc())
                
                st.write("💡 Gợi ý khắc phục:")
                st.write("- Kiểm tra file Word hợp lệ")
                st.write("- Đảm bảo có đầy đủ system dependencies")
                st.write("- Thử với file đơn giản hơn")
                st.write("- Tắt một số tùy chọn xử lý")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666;">
    <p>🚀 Full Version với đầy đủ tính năng | 📧 Hỗ trợ: support@example.com</p>
    <p>💡 Ứng dụng này sử dụng Streamlit, Pandoc, LaTeX và ImageMagick</p>
</div>
""", unsafe_allow_html=True)

# Hiển thị ví dụ nếu chưa tải file
if uploaded_file is None:
    st.header("📝 Demo tính năng đầy đủ")
    
    st.markdown("""
    <div class="success-box">
    <h4>🎯 Tính năng hoàn chỉnh của phiên bản này:</h4>
    <ol>
        <li>📁 <strong>Upload file Word</strong> chứa LaTeX, TikZ, bảng</li>
        <li>🧮 <strong>Chuyển LaTeX → Equation thật</strong> trong Word có thể chỉnh sửa</li>
        <li>🖼️ <strong>Compile TikZ → PNG</strong> chất lượng cao chèn vào Word</li>
        <li>📊 <strong>Bảng Markdown → Bảng Word</strong> chuẩn</li>
        <li>📄 <strong>Xuất Word/PDF/HTML</strong> hoàn chỉnh</li>
        <li>✨ <strong>Preview và so sánh</strong> trước/sau xử lý</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="feature-box">
        <h4>🧮 Input: LaTeX trong Word</h4>
        <code>Phương trình: $ax^2 + bx + c = 0$</code><br>
        <code>Display: $$\\int_0^1 x^2 dx = \\frac{1}{3}$$</code><br>
        <br>
        <h4>⬇️ Output: Equation thật trong Word</h4>
        <p>✅ Equation object có thể chỉnh sửa trong Word<br>
        ✅ MathML format chuẩn<br>
        ✅ Render đẹp trong Word</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="feature-box">
        <h4>🖼️ Input: TikZ Code</h4>
        <code>\\begin{tikzpicture}<br>
        \\draw (0,0) circle (1cm);<br>
        \\draw[->] (0,0) -- (1,0);<br>
        \\end{tikzpicture}</code><br>
        <br>
        <h4>⬇️ Output: Hình ảnh PNG</h4>
        <p>✅ Compile bằng pdflatex<br>
        ✅ Convert PDF → PNG chất lượng cao<br>
        ✅ Chèn vào đúng vị trí trong Word</p>
        </div>
        """, unsafe_allow_html=True)
