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
from convert_utils import process_latex_formulas, process_markdown_tables

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
    <h1>🧮 LaTeX to Word Converter</h1>
    <p>Chuyển đổi file Word chứa LaTeX, TikZ và Markdown thành Word hoàn chỉnh</p>
</div>
""", unsafe_allow_html=True)

# Sidebar với thông tin
with st.sidebar:
    st.header("📋 Hướng dẫn sử dụng")
    st.markdown("""
    **Ứng dụng này giúp bạn:**
    - ✅ Chuyển công thức LaTeX `$...$` thành Equation thật
    - ✅ Chuyển TikZ thành hình ảnh PNG  
    - ✅ Chuyển bảng Markdown thành bảng Word
    
    **Các bước thực hiện:**
    1. Tải lên file Word (.docx)
    2. Xem trước nội dung được xử lý
    3. Tải về file Word đã chuyển đổi
    """)
    
    st.header("🔧 Yêu cầu hệ thống")
    st.markdown("""
    - Pandoc
    - TeX Live (pdflatex)
    - ImageMagick
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
    process_tikz = st.checkbox("🖼️ Xử lý TikZ diagrams", value=True)  
    process_tables = st.checkbox("📊 Xử lý bảng Markdown", value=True)
    
    output_format = st.selectbox(
        "📄 Định dạng đầu ra",
        ["docx", "pdf", "html"],
        help="Chọn định dạng file muốn xuất ra"
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
                    html_content = result.html
                
                # Bước 2: Xử lý nội dung
                processed_content = html_content
                
                if process_latex:
                    st.write("🧮 Đang xử lý công thức LaTeX...")
                    processed_content = process_latex_formulas(processed_content)
                
                if process_tikz:
                    st.write("🖼️ Đang xử lý TikZ diagrams...")
                    processed_content = process_tikz_code(processed_content, temp_dir)
                
                if process_tables:
                    st.write("📊 Đang xử lý bảng...")
                    processed_content = process_markdown_tables(processed_content)
                
                # Bước 3: Chuyển đổi về Word
                st.write(f"📄 Đang chuyển đổi sang {output_format.upper()}...")
                
                # Lưu HTML tạm thời
                html_path = os.path.join(temp_dir, "processed.html")
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(processed_content)
                
                # Chuyển đổi bằng pandoc
                output_path = os.path.join(temp_dir, f"output.{output_format}")
                
                if output_format == "docx":
                    pypandoc.convert_file(
                        html_path, 
                        'docx', 
                        outputfile=output_path,
                        extra_args=['--standalone']
                    )
                elif output_format == "pdf":
                    pypandoc.convert_file(
                        html_path,
                        'pdf',
                        outputfile=output_path,
                        extra_args=['--pdf-engine=pdflatex']
                    )
                else:  # html
                    shutil.copy(html_path, output_path)
                
                st.success("✅ Xử lý hoàn tất!")
                
                # Hiển thị kết quả
                st.header("📋 Xem trước kết quả")
                
                # Tab để hiển thị nội dung
                tab1, tab2 = st.tabs(["🔍 Nội dung gốc", "✨ Nội dung đã xử lý"])
                
                with tab1:
                    st.code(html_content[:1000] + "..." if len(html_content) > 1000 else html_content, language="html")
                
                with tab2:
                    if output_format == "html":
                        st.components.v1.html(processed_content, height=400, scrolling=True)
                    else:
                        st.code(processed_content[:1000] + "..." if len(processed_content) > 1000 else processed_content, language="html")
                
                # Nút tải về
                with open(output_path, "rb") as file:
                    file_data = file.read()
                    
                st.download_button(
                    label=f"💾 Tải về file {output_format.upper()}",
                    data=file_data,
                    file_name=f"converted_document.{output_format}",
                    mime=f"application/{output_format}",
                    type="primary"
                )
                
                # Thống kê xử lý
                st.header("📊 Thống kê xử lý")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    latex_count = len(re.findall(r'\$+[^$]+\$+', html_content))
                    st.metric("🧮 Công thức LaTeX", latex_count)
                
                with col2:
                    tikz_count = len(re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', html_content, re.DOTALL))
                    st.metric("🖼️ TikZ diagrams", tikz_count)
                
                with col3:
                    table_count = len(re.findall(r'\|.*?\|', html_content))
                    st.metric("📊 Bảng phát hiện", table_count)
                
                # Dọn dẹp file tạm thời
                shutil.rmtree(temp_dir)
                
            except Exception as e:
                st.error(f"❌ Lỗi khi xử lý: {str(e)}")
                st.write("💡 Hãy đảm bảo:")
                st.write("- File Word hợp lệ")
                st.write("- Đã cài đặt pandoc và các công cụ cần thiết")
                st.write("- Cú pháp LaTeX và TikZ đúng")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666;">
    <p>🚀 Phát triển bởi AI Assistant | 📧 Hỗ trợ: support@example.com</p>
    <p>💡 Ứng dụng này sử dụng Streamlit, Pandoc, và LaTeX</p>
</div>
""", unsafe_allow_html=True)

# Hiển thị ví dụ nếu chưa tải file
if uploaded_file is None:
    st.header("📝 Ví dụ định dạng đầu vào")
    
    st.markdown("""
    <div class="feature-box">
    <h4>🧮 Công thức LaTeX</h4>
    <code>Phương trình bậc hai: $ax^2 + bx + c = 0$</code><br>
    <code>Tích phân: $\\int_0^1 x^2 dx = \\frac{1}{3}$</code>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="feature-box">
    <h4>🖼️ TikZ Diagram</h4>
    <code>
    \\begin{tikzpicture}<br>
    \\draw (0,0) circle (1cm);<br>
    \\draw (0,0) -- (1,0);<br>
    \\end{tikzpicture}
    </code>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="feature-box">
    <h4>📊 Bảng Markdown</h4>
    <code>
    | Tên | Điểm | Xếp loại |<br>
    |-----|------|----------|<br>
    | An  | 9.5  | Giỏi     |<br>
    | Bình| 8.0  | Khá      |
    </code>
    </div>
    """, unsafe_allow_html=True)
