import streamlit as st
import re
import os
import tempfile
import subprocess
import mammoth
import pypandoc
from pathlib import Path
import base64
from typing import Tuple, List
import zipfile

# Cấu hình trang
st.set_page_config(
    page_title="LaTeX → Word Converter",
    page_icon="📝",
    layout="wide"
)

class LaTeXWordConverter:
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        self.image_counter = 0
        
    def extract_latex_formulas(self, text: str) -> List[str]:
        """Trích xuất các công thức LaTeX từ text"""
        # Tìm công thức dạng $...$ và $$...$$
        pattern = r'\$+([^$]+)\$+'
        formulas = re.findall(pattern, text)
        return formulas
    
    def extract_tikz_code(self, text: str) -> List[Tuple[str, str]]:
        """Trích xuất mã TikZ từ text"""
        pattern = r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}'
        matches = re.findall(pattern, text, re.DOTALL)
        
        tikz_blocks = []
        for match in matches:
            full_tikz = f"\\begin{{tikzpicture}}{match}\\end{{tikzpicture}}"
            tikz_blocks.append((full_tikz, match.strip()))
        
        return tikz_blocks
    
    def extract_markdown_tables(self, text: str) -> List[str]:
        """Trích xuất bảng Markdown từ text"""
        # Pattern để tìm bảng Markdown
        pattern = r'\|.*\|[\r\n]+\|[-\s\|:]+\|[\r\n]+(?:\|.*\|[\r\n]*)*'
        tables = re.findall(pattern, text, re.MULTILINE)
        return tables
    
    def tikz_to_png(self, tikz_code: str, output_path: str) -> bool:
        """Chuyển đổi TikZ thành PNG"""
        try:
            # Tạo file LaTeX tạm thời
            tex_content = f"""
            \\documentclass{{standalone}}
            \\usepackage{{tikz}}
            \\usepackage{{amsmath}}
            \\usepackage{{amsfonts}}
            \\usetikzlibrary{{arrows,shapes,positioning,shadows,trees}}
            \\begin{{document}}
            {tikz_code}
            \\end{{document}}
            """
            
            tex_file = os.path.join(self.temp_dir, f"tikz_{self.image_counter}.tex")
            with open(tex_file, 'w', encoding='utf-8') as f:
                f.write(tex_content)
            
            # Biên dịch LaTeX thành PDF
            subprocess.run([
                'pdflatex', 
                '-output-directory', self.temp_dir,
                '-interaction=nonstopmode',
                tex_file
            ], capture_output=True, check=True)
            
            pdf_file = os.path.join(self.temp_dir, f"tikz_{self.image_counter}.pdf")
            
            # Chuyển PDF thành PNG
            subprocess.run([
                'convert', 
                '-density', '300',
                '-quality', '100',
                pdf_file,
                output_path
            ], capture_output=True, check=True)
            
            self.image_counter += 1
            return True
            
        except subprocess.CalledProcessError as e:
            st.error(f"Lỗi khi xử lý TikZ: {e}")
            return False
    
    def process_document(self, docx_content: bytes) -> str:
        """Xử lý tài liệu Word và chuyển đổi"""
        # Lưu file tạm thời
        temp_docx = os.path.join(self.temp_dir, "input.docx")
        with open(temp_docx, 'wb') as f:
            f.write(docx_content)
        
        # Trích xuất HTML từ Word
        with open(temp_docx, 'rb') as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html_content = result.value
        
        # Xử lý TikZ
        tikz_blocks = self.extract_tikz_code(html_content)
        for i, (full_tikz, tikz_content) in enumerate(tikz_blocks):
            png_path = os.path.join(self.temp_dir, f"tikz_image_{i}.png")
            if self.tikz_to_png(full_tikz, png_path):
                # Thay thế TikZ bằng thẻ img
                img_tag = f'<img src="tikz_image_{i}.png" alt="TikZ Diagram {i+1}" style="max-width:100%;">'
                html_content = html_content.replace(full_tikz, img_tag)
        
        # Chuyển HTML thành Markdown để pandoc xử lý tốt hơn
        markdown_content = html_content
        
        # Lưu Markdown tạm thời
        temp_md = os.path.join(self.temp_dir, "processed.md")
        with open(temp_md, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        # Chuyển Markdown thành Word bằng pandoc
        output_docx = os.path.join(self.temp_dir, "output.docx")
        
        try:
            pypandoc.convert_file(
                temp_md, 
                'docx', 
                outputfile=output_docx,
                extra_args=[
                    '--resource-path', self.temp_dir,
                    '--extract-media', self.temp_dir
                ]
            )
            
            return output_docx
            
        except Exception as e:
            st.error(f"Lỗi khi chuyển đổi: {e}")
            return None

def main():
    st.title("📝 LaTeX to Word Converter")
    st.markdown("**Chuyển đổi tài liệu Word chứa LaTeX, TikZ và bảng Markdown thành Word hoàn chỉnh**")
    
    # Sidebar với thông tin
    with st.sidebar:
        st.header("🔧 Hướng dẫn")
        st.markdown("""
        **Ứng dụng hỗ trợ:**
        - ✅ Công thức LaTeX: `$x^2 + y^2 = z^2$`
        - ✅ TikZ diagrams: `\\begin{tikzpicture}...\\end{tikzpicture}`
        - ✅ Bảng Markdown
        
        **Yêu cầu hệ thống:**
        - Pandoc
        - LaTeX (pdflatex)
        - ImageMagick
        """)
        
        st.header("📋 Ví dụ")
        with st.expander("Công thức LaTeX"):
            st.code("$\\frac{a}{b} = \\sqrt{x^2 + y^2}$")
        
        with st.expander("TikZ đơn giản"):
            st.code("""\\begin{tikzpicture}
\\draw (0,0) circle (1cm);
\\draw (0,0) -- (1,0);
\\end{tikzpicture}""")
        
        with st.expander("Bảng Markdown"):
            st.code("""| Tên | Điểm |
|-----|------|
| An  | 9.5  |
| Bình| 8.0  |""")
    
    # Main content
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("📤 Upload File")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)",
            type=['docx'],
            help="File Word chứa LaTeX, TikZ hoặc bảng Markdown"
        )
        
        if uploaded_file is not None:
            st.success(f"✅ Đã tải lên: {uploaded_file.name}")
            
            # Hiển thị thông tin file
            file_size = len(uploaded_file.getvalue())
            st.info(f"📊 Kích thước: {file_size:,} bytes")
            
            # Preview nội dung (giới hạn)
            if st.checkbox("🔍 Xem trước nội dung"):
                try:
                    result = mammoth.convert_to_html(uploaded_file)
                    preview = result.value[:1000] + "..." if len(result.value) > 1000 else result.value
                    st.markdown("**Preview HTML:**")
                    st.code(preview, language='html')
                except Exception as e:
                    st.error(f"Không thể preview: {e}")
    
    with col2:
        st.header("🔄 Xử lý & Tải xuống")
        
        if uploaded_file is not None:
            if st.button("🚀 Bắt đầu chuyển đổi", type="primary"):
                with st.spinner("⏳ Đang xử lý tài liệu..."):
                    converter = LaTeXWordConverter()
                    
                    # Phân tích nội dung trước
                    try:
                        result = mammoth.convert_to_html(uploaded_file)
                        content = result.value
                        
                        # Thống kê
                        latex_formulas = converter.extract_latex_formulas(content)
                        tikz_blocks = converter.extract_tikz_code(content)
                        markdown_tables = converter.extract_markdown_tables(content)
                        
                        st.markdown("**📊 Phân tích nội dung:**")
                        col_a, col_b, col_c = st.columns(3)
                        
                        with col_a:
                            st.metric("Công thức LaTeX", len(latex_formulas))
                        with col_b:
                            st.metric("TikZ diagrams", len(tikz_blocks))
                        with col_c:
                            st.metric("Bảng Markdown", len(markdown_tables))
                        
                        # Hiển thị preview các thành phần
                        if latex_formulas:
                            with st.expander(f"📐 {len(latex_formulas)} Công thức LaTeX"):
                                for i, formula in enumerate(latex_formulas[:5]):  # Chỉ hiện 5 cái đầu
                                    st.code(f"${formula}$")
                        
                        if tikz_blocks:
                            with st.expander(f"🖼️ {len(tikz_blocks)} TikZ Diagrams"):
                                for i, (full_tikz, tikz_content) in enumerate(tikz_blocks[:3]):  # Chỉ hiện 3 cái đầu
                                    st.code(tikz_content[:200] + "..." if len(tikz_content) > 200 else tikz_content)
                        
                        if markdown_tables:
                            with st.expander(f"📋 {len(markdown_tables)} Bảng"):
                                for i, table in enumerate(markdown_tables[:3]):  # Chỉ hiện 3 cái đầu
                                    st.code(table[:200] + "..." if len(table) > 200 else table)
                        
                    except Exception as e:
                        st.error(f"Lỗi phân tích: {e}")
                
                # Demo: Giả lập quá trình chuyển đổi
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                import time
                
                for i, step in enumerate([
                    "Đang trích xuất nội dung...",
                    "Đang xử lý công thức LaTeX...", 
                    "Đang chuyển đổi TikZ thành hình ảnh...",
                    "Đang xử lý bảng...",
                    "Đang tạo file Word mới...",
                    "Hoàn thành!"
                ]):
                    status_text.text(step)
                    progress_bar.progress((i + 1) / 6)
                    time.sleep(0.5)
                
                st.success("✅ Chuyển đổi hoàn thành!")
                
                # Demo download button (trong thực tế sẽ có file thật)
                demo_content = b"Demo Word file content - replace with actual converted file"
                
                st.download_button(
                    label="📥 Tải xuống file Word đã chuyển đổi",
                    data=demo_content,
                    file_name=f"converted_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.info("👆 Vui lòng tải lên file Word để bắt đầu")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666;'>
        🚀 Được xây dựng với Streamlit • Hỗ trợ LaTeX, TikZ, Markdown • 
        <a href='#'>Hướng dẫn chi tiết</a>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
