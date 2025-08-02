import streamlit as st
import re
import io
import os
import tempfile
import subprocess
from pathlib import Path
import base64

# Required imports (install via pip)
try:
    from docx import Document
    from docx.shared import Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import mammoth
    import markdown
    from PIL import Image
    import fitz  # PyMuPDF
except ImportError as e:
    st.error(f"Missing required packages. Please install: {e}")
    st.stop()

class WordDocumentProcessor:
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        
    def read_word_document(self, uploaded_file):
        """Đọc nội dung từ file Word"""
        try:
            # Sử dụng mammoth để đọc và giữ format tốt hơn
            result = mammoth.extract_raw_text(uploaded_file)
            return result.value
        except:
            # Fallback sang python-docx
            doc = Document(uploaded_file)
            content = []
            for paragraph in doc.paragraphs:
                content.append(paragraph.text)
            return '\n'.join(content)
    
    def find_latex_formulas(self, text):
        """Tìm và trích xuất các công thức LaTeX"""
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
                    'original': match.group(0),
                    'latex': match.group(1),
                    'start': match.start(),
                    'end': match.end()
                })
        
        # Sort by position để xử lý từ cuối lên đầu
        formulas.sort(key=lambda x: x['start'], reverse=True)
        return formulas
    
    def latex_to_mathml(self, latex_string):
        """Chuyển đổi LaTeX sang MathML (simplified)"""
        # Đây là implementation đơn giản, thực tế nên dùng latex2mathml
        # hoặc gọi API service như MathJax
        
        # Một số chuyển đổi cơ bản
        replacements = {
            r'\\pi': 'π',
            r'\\alpha': 'α',
            r'\\beta': 'β',
            r'\\gamma': 'γ',
            r'\\delta': 'δ',
            r'\\epsilon': 'ε',
            r'\\theta': 'θ',
            r'\\lambda': 'λ',
            r'\\mu': 'μ',
            r'\\sigma': 'σ',
            r'\\phi': 'φ',
            r'\\omega': 'ω',
            r'\\infty': '∞',
            r'\\sum': '∑',
            r'\\int': '∫',
            r'\\sqrt': '√',
            r'\^(\w+)': lambda m: f"^{m.group(1)}",
            r'_(\w+)': lambda m: f"_{m.group(1)}",
        }
        
        result = latex_string
        for pattern, replacement in replacements.items():
            if callable(replacement):
                result = re.sub(pattern, replacement, result)
            else:
                result = result.replace(pattern, replacement)
        
        return result
    
    def find_markdown_tables(self, text):
        """Tìm và parse bảng Markdown"""
        # Pattern cho bảng markdown
        table_pattern = r'(\|[^\n]+\|\n)*\|[-|\s:]+\|\n(\|[^\n]+\|\n?)+'
        tables = []
        
        for match in re.finditer(table_pattern, text, re.MULTILINE):
            table_text = match.group(0)
            # Parse bảng
            lines = table_text.strip().split('\n')
            
            if len(lines) >= 2:
                # Header
                header = [cell.strip() for cell in lines[0].split('|')[1:-1]]
                
                # Rows (bỏ qua dòng separator)
                rows = []
                for line in lines[2:]:
                    if line.strip():
                        row = [cell.strip() for cell in line.split('|')[1:-1]]
                        rows.append(row)
                
                tables.append({
                    'original': table_text,
                    'header': header,
                    'rows': rows,
                    'start': match.start(),
                    'end': match.end()
                })
        
        tables.sort(key=lambda x: x['start'], reverse=True)
        return tables
    
    def find_tikz_code(self, text):
        """Tìm và trích xuất mã TikZ"""
        pattern = r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}'
        tikz_blocks = []
        
        for match in re.finditer(pattern, text, re.DOTALL):
            tikz_blocks.append({
                'original': match.group(0),
                'code': match.group(1).strip(),
                'start': match.start(),
                'end': match.end()
            })
        
        tikz_blocks.sort(key=lambda x: x['start'], reverse=True)
        return tikz_blocks
    
    def compile_tikz_to_png(self, tikz_code):
        """Biên dịch TikZ thành PNG"""
        try:
            # Tạo file LaTeX tạm thời
            latex_content = f"""
\\documentclass{{standalone}}
\\usepackage{{tikz}}
\\begin{{document}}
\\begin{{tikzpicture}}
{tikz_code}
\\end{{tikzpicture}}
\\end{{document}}
"""
            
            # Tạo file tạm
            tex_file = os.path.join(self.temp_dir, 'temp.tex')
            with open(tex_file, 'w', encoding='utf-8') as f:
                f.write(latex_content)
            
            # Biên dịch với pdflatex
            result = subprocess.run([
                'pdflatex', '-output-directory', self.temp_dir, tex_file
            ], capture_output=True, text=True, cwd=self.temp_dir)
            
            if result.returncode != 0:
                st.warning(f"LaTeX compilation error: {result.stderr}")
                return None
            
            # Chuyển PDF sang PNG
            pdf_file = os.path.join(self.temp_dir, 'temp.pdf')
            if os.path.exists(pdf_file):
                # Sử dụng PyMuPDF để chuyển PDF sang PNG
                doc = fitz.open(pdf_file)
                page = doc[0]
                mat = fitz.Matrix(2.0, 2.0)  # Zoom 2x
                pix = page.get_pixmap(matrix=mat)
                
                png_file = os.path.join(self.temp_dir, 'temp.png')
                pix.save(png_file)
                doc.close()
                
                return png_file
            
        except Exception as e:
            st.warning(f"Error compiling TikZ: {e}")
            return None
    
    def create_processed_document(self, original_text, formulas, tables, tikz_blocks):
        """Tạo document Word đã được xử lý"""
        doc = Document()
        
        # Xử lý text theo thứ tự từ đầu đến cuối
        current_pos = 0
        processed_text = original_text
        
        # Collect all elements to process
        all_elements = []
        
        for formula in formulas:
            all_elements.append({
                'type': 'formula',
                'start': formula['start'],
                'end': formula['end'],
                'data': formula
            })
        
        for table in tables:
            all_elements.append({
                'type': 'table',
                'start': table['start'],
                'end': table['end'],
                'data': table
            })
        
        for tikz in tikz_blocks:
            all_elements.append({
                'type': 'tikz',
                'start': tikz['start'],
                'end': tikz['end'],
                'data': tikz
            })
        
        # Sort by position
        all_elements.sort(key=lambda x: x['start'])
        
        current_pos = 0
        for element in all_elements:
            # Add text before element
            if element['start'] > current_pos:
                text_before = processed_text[current_pos:element['start']]
                if text_before.strip():
                    doc.add_paragraph(text_before.strip())
            
            # Process element
            if element['type'] == 'formula':
                # Add formula as formatted text (MathML insertion is complex)
                formula_text = self.latex_to_mathml(element['data']['latex'])
                p = doc.add_paragraph()
                run = p.add_run(f"[FORMULA: {formula_text}]")
                run.italic = True
                
            elif element['type'] == 'table':
                # Add table
                table_data = element['data']
                table = doc.add_table(rows=1+len(table_data['rows']), cols=len(table_data['header']))
                table.style = 'Table Grid'
                
                # Header
                for i, header in enumerate(table_data['header']):
                    table.cell(0, i).text = header
                    table.cell(0, i).paragraphs[0].runs[0].bold = True
                
                # Rows
                for row_idx, row in enumerate(table_data['rows']):
                    for col_idx, cell_text in enumerate(row):
                        if col_idx < len(table_data['header']):
                            table.cell(row_idx + 1, col_idx).text = cell_text
                
            elif element['type'] == 'tikz':
                # Add TikZ as image
                png_path = self.compile_tikz_to_png(element['data']['code'])
                if png_path and os.path.exists(png_path):
                    p = doc.add_paragraph()
                    run = p.add_run()
                    run.add_picture(png_path, width=Inches(4))
                else:
                    p = doc.add_paragraph()
                    run = p.add_run(f"[TikZ CODE: {element['data']['code'][:50]}...]")
                    run.italic = True
            
            current_pos = element['end']
        
        # Add remaining text
        if current_pos < len(processed_text):
            remaining_text = processed_text[current_pos:]
            if remaining_text.strip():
                doc.add_paragraph(remaining_text.strip())
        
        return doc
    
    def cleanup(self):
        """Dọn dẹp files tạm thời"""
        import shutil
        try:
            shutil.rmtree(self.temp_dir)
        except:
            pass

def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📝",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("""
    **Chức năng:**
    - Chuyển đổi công thức LaTeX ($...$, ${...}$) → Equation trong Word
    - Chuyển đổi bảng Markdown → Bảng thật trong Word  
    - Chuyển đổi mã TikZ → Ảnh PNG trong Word
    """)
    
    # Upload file
    uploaded_file = st.file_uploader(
        "📁 Upload file Word (.docx)", 
        type=['docx'],
        help="Chọn file Word chứa LaTeX, Markdown tables, và TikZ code"
    )
    
    if uploaded_file:
        processor = WordDocumentProcessor()
        
        try:
            # Đọc nội dung
            with st.spinner("Đang đọc file Word..."):
                content = processor.read_word_document(uploaded_file)
            
            # Preview nội dung gốc
            with st.expander("📖 Xem nội dung gốc"):
                st.text_area("Content", content, height=200)
            
            # Tìm các elements cần xử lý
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.subheader("🧮 Công thức LaTeX")
                formulas = processor.find_latex_formulas(content)
                st.write(f"Tìm thấy {len(formulas)} công thức")
                for i, formula in enumerate(formulas[:3]):  # Show first 3
                    st.code(formula['original'])
                    st.write(f"→ {processor.latex_to_mathml(formula['latex'])}")
            
            with col2:
                st.subheader("📊 Bảng Markdown")
                tables = processor.find_markdown_tables(content)
                st.write(f"Tìm thấy {len(tables)} bảng")
                for i, table in enumerate(tables[:2]):  # Show first 2
                    st.write(f"**Bảng {i+1}:**")
                    st.dataframe({
                        col: [row[j] if j < len(row) else '' for row in table['rows']]
                        for j, col in enumerate(table['header'])
                    })
            
            with col3:
                st.subheader("🎨 Mã TikZ") 
                tikz_blocks = processor.find_tikz_code(content)
                st.write(f"Tìm thấy {len(tikz_blocks)} khối TikZ")
                for i, tikz in enumerate(tikz_blocks[:2]):  # Show first 2
                    st.code(tikz['code'][:100] + "..." if len(tikz['code']) > 100 else tikz['code'])
            
            # Process button
            if st.button("🚀 Xử lý và tạo file Word mới", type="primary"):
                with st.spinner("Đang xử lý document..."):
                    processed_doc = processor.create_processed_document(
                        content, formulas, tables, tikz_blocks
                    )
                    
                    # Save to bytes
                    doc_io = io.BytesIO()
                    processed_doc.save(doc_io)
                    doc_io.seek(0)
                    
                    # Download button
                    st.success("✅ Xử lý hoàn tất!")
                    st.download_button(
                        label="📥 Tải xuống file Word đã xử lý",
                        data=doc_io.getvalue(),
                        file_name="processed_document.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    # Show stats
                    st.info(f"""
                    📈 **Thống kê xử lý:**
                    - Công thức LaTeX: {len(formulas)}
                    - Bảng Markdown: {len(tables)} 
                    - Khối TikZ: {len(tikz_blocks)}
                    """)
        
        except Exception as e:
            st.error(f"❌ Lỗi xử lý file: {e}")
        
        finally:
            processor.cleanup()
    
    # Instructions
    with st.expander("📋 Hướng dẫn sử dụng"):
        st.markdown("""
        ### 📝 Format đầu vào được hỗ trợ:
        
        **1. Công thức LaTeX:**
        ```
        Tính diện tích: $A = \\pi r^2$
        Hoặc: ${E = mc^2}$
        ```
        
        **2. Bảng Markdown:**
        ```
        | Tên | Tuổi | Điểm |
        |-----|------|------|
        | An  | 15   | 9.0  |
        | Bình| 16   | 8.5  |
        ```
        
        **3. Mã TikZ:**
        ```
        \\begin{tikzpicture}
        \\draw (0,0) circle (1cm);
        \\draw (0,0) -- (1,0);
        \\end{tikzpicture}
        ```
        
        ### ⚙️ Yêu cầu hệ thống:
        - Python packages: `python-docx`, `mammoth`, `markdown`, `PyMuPDF`, `Pillow`
        - LaTeX distribution (MiKTeX/TeX Live) cho TikZ
        - Package TikZ trong LaTeX
        """)
    
    # Technical notes
    with st.expander("🔧 Ghi chú kỹ thuật"):
        st.markdown("""
        ### 🚀 Cải tiến có thể thêm:
        
        1. **LaTeX → MathML:** Dùng `latex2mathml` hoặc MathJax API
        2. **Word Equations:** Chèn Office MathML thực sự
        3. **TikZ cải tiến:** Dùng `dvisvgm` cho vector graphics
        4. **Batch processing:** Xử lý nhiều files cùng lúc
        5. **Template support:** Hỗ trợ template Word có sẵn
        6. **OCR integration:** Đọc công thức từ ảnh scan
        
        ### 📚 Dependencies cần cài đặt:
        ```bash
        pip install streamlit python-docx mammoth markdown PyMuPDF Pillow
        
        # Cài LaTeX (Ubuntu/Debian)
        sudo apt-get install texlive-full
        
        # Hoặc Windows: tải MiKTeX
        ```
        """)

if __name__ == "__main__":
    main()
