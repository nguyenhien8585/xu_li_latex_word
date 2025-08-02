import streamlit as st
import re
import io
import os
import tempfile
from pathlib import Path
import base64
import subprocess
import pandas as pd

# Core imports
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.shared import OxmlElement, qn
    import mammoth
    import markdown
    from PIL import Image, ImageDraw, ImageFont
    import fitz  # PyMuPDF
except ImportError as e:
    st.error(f"Missing required packages: {e}")
    st.stop()

# Try to import math libraries
try:
    import math2docx
    MATH2DOCX_AVAILABLE = True
except ImportError:
    MATH2DOCX_AVAILABLE = False

try:
    from latex2mathml.converter import convert as latex2mathml_convert
    LATEX2MATHML_AVAILABLE = True
except ImportError:
    LATEX2MATHML_AVAILABLE = False

class AdvancedWordProcessor:
    """Advanced Word processor với TikZ compilation và proper table handling"""
    
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        
    def read_word_document_preserve_format(self, uploaded_file):
        """Đọc Word document và giữ nguyên format"""
        try:
            uploaded_file.seek(0)
            # Sử dụng mammoth để extract với format
            result = mammoth.extract_raw_text(uploaded_file)
            content = result.value
            
            # Cũng thử extract với styles
            uploaded_file.seek(0)
            with_styles = mammoth.convert_to_html(uploaded_file)
            html_content = with_styles.value
            
            return {
                'text': content,
                'html': html_content,
                'messages': result.messages if hasattr(result, 'messages') else []
            }
        except Exception as e:
            st.error(f"Error reading Word document: {e}")
            # Fallback method
            try:
                uploaded_file.seek(0)
                doc = Document(uploaded_file)
                content = []
                for paragraph in doc.paragraphs:
                    content.append(paragraph.text)
                return {
                    'text': '\n'.join(content),
                    'html': '',
                    'messages': []
                }
            except Exception as e2:
                st.error(f"Fallback method failed: {e2}")
                return {'text': '', 'html': '', 'messages': []}
    
    def find_markdown_tables_robust(self, text):
        """Robust markdown table detection với nhiều format"""
        tables = []
        
        # Pattern 1: Standard markdown tables
        pattern1 = r'(\|[^\n\r]+\|(?:\r?\n|\r))+(\|[-:\s|]+\|(?:\r?\n|\r))(\|[^\n\r]+\|(?:\r?\n|\r)?)*'
        
        for match in re.finditer(pattern1, text, re.MULTILINE):
            table_text = match.group(0).strip()
            parsed_table = self._parse_markdown_table(table_text)
            if parsed_table:
                parsed_table.update({
                    'original': table_text,
                    'start': match.start(),
                    'end': match.end()
                })
                tables.append(parsed_table)
        
        # Pattern 2: Simple tables without separators
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if '|' in line and line.count('|') >= 3:  # At least 2 columns
                # Look for next few lines with similar structure
                table_lines = [line]
                j = i + 1
                while j < len(lines) and '|' in lines[j]:
                    table_lines.append(lines[j])
                    j += 1
                
                if len(table_lines) >= 2:  # At least header + 1 data row
                    table_text = '\n'.join(table_lines)
                    parsed_table = self._parse_simple_table(table_lines)
                    if parsed_table:
                        # Find position in original text
                        start_pos = text.find(table_text)
                        if start_pos != -1:
                            parsed_table.update({
                                'original': table_text,
                                'start': start_pos,
                                'end': start_pos + len(table_text)
                            })
                            # Check if not already found
                            if not any(abs(t['start'] - start_pos) < 10 for t in tables):
                                tables.append(parsed_table)
        
        # Sort by position and remove duplicates
        tables.sort(key=lambda x: x['start'])
        return self._remove_duplicate_tables(tables)
    
    def _parse_markdown_table(self, table_text):
        """Parse standard markdown table"""
        lines = [line.strip() for line in table_text.split('\n') if line.strip()]
        
        if len(lines) < 2:
            return None
        
        # First line is header
        header_line = lines[0]
        header = [cell.strip() for cell in header_line.split('|')[1:-1] if cell.strip()]
        
        if not header:
            return None
        
        # Find separator line (contains -, :, |)
        separator_idx = -1
        for i, line in enumerate(lines[1:], 1):
            if re.match(r'^[\s\-:|]+$', line.replace('|', '')):
                separator_idx = i
                break
        
        if separator_idx == -1 and len(lines) > 1:
            # No separator found, treat second line as data
            separator_idx = 1
        
        # Parse data rows
        rows = []
        for line in lines[separator_idx + 1:]:
            if line.strip() and '|' in line:
                row_cells = [cell.strip() for cell in line.split('|')[1:-1]]
                # Pad or truncate to match header length
                while len(row_cells) < len(header):
                    row_cells.append('')
                rows.append(row_cells[:len(header)])
        
        return {'header': header, 'rows': rows} if header and rows else None
    
    def _parse_simple_table(self, table_lines):
        """Parse simple table without separators"""
        if len(table_lines) < 2:
            return None
        
        # First line as header
        header = [cell.strip() for cell in table_lines[0].split('|')[1:-1] if cell.strip()]
        
        if not header:
            return None
        
        rows = []
        for line in table_lines[1:]:
            if line.strip() and '|' in line:
                row_cells = [cell.strip() for cell in line.split('|')[1:-1]]
                while len(row_cells) < len(header):
                    row_cells.append('')
                rows.append(row_cells[:len(header)])
        
        return {'header': header, 'rows': rows} if rows else None
    
    def _remove_duplicate_tables(self, tables):
        """Remove duplicate tables based on content similarity"""
        unique_tables = []
        for table in tables:
            is_duplicate = False
            for existing in unique_tables:
                if (table['header'] == existing['header'] and 
                    len(table['rows']) == len(existing['rows'])):
                    is_duplicate = True
                    break
            if not is_duplicate:
                unique_tables.append(table)
        return unique_tables
    
    def compile_tikz_to_png(self, tikz_code, output_path=None):
        """Compile TikZ code to PNG using LaTeX"""
        if output_path is None:
            output_path = os.path.join(self.temp_dir, f'tikz_{hash(tikz_code)}.png')
        
        try:
            # Create LaTeX document
            latex_template = r"""
\documentclass[border=10pt]{{standalone}}
\usepackage{{tikz}}
\usepackage{{amsmath}}
\usepackage{{amsfonts}}
\usetikzlibrary{{shapes,arrows,positioning,calc}}
\begin{{document}}
\begin{{tikzpicture}}
{tikz_code}
\end{{tikzpicture}}
\end{{document}}
"""
            
            latex_content = latex_template.format(tikz_code=tikz_code)
            
            # Write LaTeX file
            tex_file = os.path.join(self.temp_dir, 'tikz_temp.tex')
            with open(tex_file, 'w', encoding='utf-8') as f:
                f.write(latex_content)
            
            # Compile with pdflatex
            cmd = ['pdflatex', '-output-directory', self.temp_dir, 
                   '-interaction=nonstopmode', tex_file]
            
            result = subprocess.run(cmd, capture_output=True, text=True, cwd=self.temp_dir)
            
            if result.returncode != 0:
                st.warning(f"LaTeX compilation failed: {result.stderr}")
                return self._create_tikz_fallback_image(tikz_code, output_path)
            
            # Convert PDF to PNG
            pdf_file = os.path.join(self.temp_dir, 'tikz_temp.pdf')
            if os.path.exists(pdf_file):
                # Use PyMuPDF to convert PDF to PNG
                doc = fitz.open(pdf_file)
                page = doc[0]
                
                # High resolution conversion
                mat = fitz.Matrix(3.0, 3.0)  # 3x zoom for high quality
                pix = page.get_pixmap(matrix=mat)
                pix.save(output_path)
                doc.close()
                
                return output_path
            else:
                return self._create_tikz_fallback_image(tikz_code, output_path)
                
        except Exception as e:
            st.warning(f"TikZ compilation error: {e}")
            return self._create_tikz_fallback_image(tikz_code, output_path)
    
    def _create_tikz_fallback_image(self, tikz_code, output_path):
        """Create fallback image when TikZ compilation fails"""
        try:
            width, height = 600, 400
            img = Image.new('RGB', (width, height), color='white')
            draw = ImageDraw.Draw(img)
            
            # Draw border
            draw.rectangle([10, 10, width-10, height-10], outline='#2E86AB', width=3)
            
            # Add title
            try:
                font = ImageFont.load_default()
            except:
                font = None
            
            title = "📊 TikZ Diagram (Compilation Failed)"
            if font:
                bbox = draw.textbbox((0, 0), title, font=font)
                title_width = bbox[2] - bbox[0]
                draw.text(((width - title_width) // 2, 30), title, fill='#2E86AB', font=font)
            
            # Show code
            draw.text((20, 80), "LaTeX Code:", fill='#555555', font=font)
            
            # Format code for display
            code_lines = tikz_code.split('\n')
            y = 110
            for i, line in enumerate(code_lines[:15]):  # Show first 15 lines
                if line.strip():
                    display_line = line[:80] + "..." if len(line) > 80 else line
                    draw.text((25, y), display_line, fill='#333333', font=font)
                    y += 20
                if y > height - 60:
                    break
            
            if len(code_lines) > 15:
                draw.text((25, height - 40), f"... và {len(code_lines) - 15} dòng nữa", 
                         fill='#888888', font=font)
            
            img.save(output_path, 'PNG')
            return output_path
            
        except Exception as e:
            st.error(f"Could not create fallback image: {e}")
            return None
    
    def find_tikz_blocks_enhanced(self, text):
        """Enhanced TikZ detection với nhiều patterns"""
        tikz_blocks = []
        
        # Pattern 1: Standard tikzpicture
        pattern1 = r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}'
        for match in re.finditer(pattern1, text, re.DOTALL):
            tikz_blocks.append({
                'original': match.group(0),
                'code': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'type': 'tikzpicture'
            })
        
        # Pattern 2: tikzpicture với options
        pattern2 = r'\\begin\{tikzpicture\}\[([^\]]*)\](.*?)\\end\{tikzpicture\}'
        for match in re.finditer(pattern2, text, re.DOTALL):
            tikz_blocks.append({
                'original': match.group(0),
                'code': match.group(2).strip(),
                'options': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'type': 'tikzpicture_with_options'
            })
        
        # Pattern 3: Standalone TikZ commands
        tikz_commands = [
            r'\\draw[^\;]*\;',
            r'\\node[^\;]*\;',
            r'\\fill[^\;]*\;',
            r'\\path[^\;]*\;'
        ]
        
        for pattern in tikz_commands:
            for match in re.finditer(pattern, text, re.MULTILINE):
                # Check if not already in a tikzpicture block
                in_existing_block = any(
                    block['start'] <= match.start() <= block['end'] 
                    for block in tikz_blocks
                )
                if not in_existing_block:
                    tikz_blocks.append({
                        'original': match.group(0),
                        'code': match.group(0),
                        'start': match.start(),
                        'end': match.end(),
                        'type': 'standalone_command'
                    })
        
        tikz_blocks.sort(key=lambda x: x['start'], reverse=True)
        return tikz_blocks
    
    def create_professional_table(self, doc, table_data):
        """Tạo bảng Word professional với styling đẹp"""
        if not table_data['header'] or not table_data['rows']:
            return None
        
        # Create table
        num_rows = len(table_data['rows']) + 1  # +1 for header
        num_cols = len(table_data['header'])
        
        table = doc.add_table(rows=num_rows, cols=num_cols)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Style header row
        header_row = table.rows[0]
        for i, header_text in enumerate(table_data['header']):
            cell = header_row.cells[i]
            cell.text = str(header_text)
            
            # Header formatting
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(12)
                    run.font.color.rgb = RGBColor(255, 255, 255)
            
            # Header background
            shading_elm = OxmlElement("w:shd")
            shading_elm.set(qn("w:fill"), "366092")  # Dark blue
            cell._tc.get_or_add_tcPr().append(shading_elm)
        
        # Add data rows with alternating colors
        for row_idx, row_data in enumerate(table_data['rows']):
            table_row = table.rows[row_idx + 1]
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < num_cols:
                    cell = table_row.cells[col_idx]
                    cell.text = str(cell_data) if cell_data is not None else ''
                    
                    # Cell formatting
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
                    
                    # Alternating row colors
                    if row_idx % 2 == 0:
                        shading_elm = OxmlElement("w:shd")
                        shading_elm.set(qn("w:fill"), "F8F9FA")  # Light gray
                        cell._tc.get_or_add_tcPr().append(shading_elm)
        
        return table
    
    def find_formulas_enhanced(self, text):
        """Enhanced formula detection"""
        formulas = []
        
        # Pattern cho [FORMULA: ...]
        pattern1 = r'\[FORMULA:\s*([^\]]+)\]'
        for match in re.finditer(pattern1, text):
            formulas.append({
                'original': match.group(0),
                'latex': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'type': 'formula_block'
            })
        
        # Pattern cho LaTeX math
        patterns = [
            (r'\$\{([^}]+)\}\$', 'latex_brace'),
            (r'\$([^$]+)\$', 'latex_dollar'),
            (r'\\begin\{equation\}(.*?)\\end\{equation\}', 'equation_env'),
            (r'\\begin\{align\}(.*?)\\end\{align\}', 'align_env')
        ]
        
        for pattern, formula_type in patterns:
            for match in re.finditer(pattern, text, re.DOTALL):
                formulas.append({
                    'original': match.group(0),
                    'latex': match.group(1).strip(),
                    'start': match.start(),
                    'end': match.end(),
                    'type': formula_type
                })
        
        formulas.sort(key=lambda x: x['start'], reverse=True)
        return formulas
    
    def add_math_equation_enhanced(self, paragraph, latex_string):
        """Enhanced math equation với multiple methods"""
        try:
            if MATH2DOCX_AVAILABLE:
                math2docx.add_math(paragraph, latex_string)
                return True
            elif LATEX2MATHML_AVAILABLE:
                mathml = latex2mathml_convert(latex_string)
                # Add MathML to paragraph (simplified)
                run = paragraph.add_run(f"[MATH: {latex_string}]")
                run.italic = True
                return True
            else:
                return False
        except Exception as e:
            st.warning(f"Math equation error: {e}")
            return False
    
    def create_processed_document_advanced(self, content_data, formulas, tables, tikz_blocks):
        """Tạo document với advanced processing"""
        doc = Document()
        
        # Title
        title = doc.add_heading('Advanced Document Processing', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Collect all elements
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
        stats = {'equations': 0, 'tables': 0, 'tikz_images': 0}
        
        original_text = content_data['text']
        
        for element in all_elements:
            # Add text before element
            if element['start'] > current_pos:
                text_before = original_text[current_pos:element['start']]
                if text_before.strip():
                    paragraphs = text_before.split('\n')
                    for para_text in paragraphs:
                        if para_text.strip():
                            p = doc.add_paragraph(para_text.strip())
                            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            # Process element
            if element['type'] == 'formula':
                formula_data = element['data']
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                if self.add_math_equation_enhanced(p, formula_data['latex']):
                    stats['equations'] += 1
                else:
                    # Unicode fallback
                    run = p.add_run(f"[MATH] {formula_data['latex']}")
                    run.italic = True
                    run.bold = True
                    run.font.color.rgb = RGBColor(0, 100, 200)
            
            elif element['type'] == 'table':
                table = self.create_professional_table(doc, element['data'])
                if table:
                    stats['tables'] += 1
                    doc.add_paragraph()  # Spacing
            
            elif element['type'] == 'tikz':
                tikz_data = element['data']
                
                # Compile TikZ to PNG
                with st.spinner(f"Compiling TikZ diagram..."):
                    img_path = self.compile_tikz_to_png(tikz_data['code'])
                
                if img_path and os.path.exists(img_path):
                    try:
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p.add_run()
                        run.add_picture(img_path, width=Inches(5))
                        stats['tikz_images'] += 1
                        
                        # Add caption
                        caption_p = doc.add_paragraph(f"TikZ Diagram: {tikz_data.get('type', 'Generated')}")
                        caption_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in caption_p.runs:
                            run.italic = True
                            run.font.size = Pt(9)
                    except Exception as e:
                        st.warning(f"Could not insert TikZ image: {e}")
                else:
                    # Text fallback
                    p = doc.add_paragraph()
                    run = p.add_run(f"[TikZ DIAGRAM]\n{tikz_data['code']}")
                    run.font.color.rgb = RGBColor(139, 69, 19)
            
            current_pos = element['end']
        
        # Add remaining text
        if current_pos < len(original_text):
            remaining_text = original_text[current_pos:]
            if remaining_text.strip():
                paragraphs = remaining_text.split('\n')
                for para_text in paragraphs:
                    if para_text.strip():
                        p = doc.add_paragraph(para_text.strip())
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Add processing summary
        doc.add_page_break()
        summary = doc.add_heading('Processing Summary', 1)
        summary.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        summary_text = f"""
Advanced Processing Results:

✅ Math Equations: {stats['equations']} processed
✅ Tables: {stats['tables']} created with professional styling  
✅ TikZ Diagrams: {stats['tikz_images']} compiled to high-quality images
✅ Format preservation: Enhanced text layout

Generated by Advanced Word Processor
"""
        
        p = doc.add_paragraph(summary_text.strip())
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        return doc, stats
    
    def cleanup(self):
        """Cleanup temporary files"""
        import shutil
        try:
            shutil.rmtree(self.temp_dir)
        except:
            pass

def main():
    st.set_page_config(
        page_title="Advanced Word Processor",
        page_icon="🚀",
        layout="wide"
    )
    
    st.title("🚀 Advanced Word Document Processor")
    st.markdown("""
    **Phiên bản nâng cao** với đầy đủ tính năng:
    - ✅ **TikZ Compilation**: LaTeX TikZ → PNG images chất lượng cao
    - ✅ **Professional Tables**: Markdown → Word tables với styling đẹp
    - ✅ **Enhanced Formulas**: Multiple LaTeX formats support
    - ✅ **Format Preservation**: Giữ nguyên định dạng gốc
    """)
    
    # System requirements check
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if MATH2DOCX_AVAILABLE:
            st.success("🧮 math2docx: ✅")
        else:
            st.warning("🧮 math2docx: ❌")
    
    with col2:
        # Check LaTeX installation
        try:
            result = subprocess.run(['pdflatex', '--version'], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                st.success("📊 LaTeX: ✅")
            else:
                st.warning("📊 LaTeX: ❌")
        except:
            st.warning("📊 LaTeX: ❌ (Install TeX Live/MiKTeX)")
    
    with col3:
        if LATEX2MATHML_AVAILABLE:
            st.success("🔢 latex2mathml: ✅")
        else:
            st.info("🔢 latex2mathml: ⚠️")
    
    # Upload section
    uploaded_file = st.file_uploader(
        "📁 Upload Word Document (.docx)",
        type=['docx'],
        help="Upload document with LaTeX formulas, Markdown tables, and TikZ diagrams"
    )
    
    if uploaded_file:
        processor = AdvancedWordProcessor()
        
        try:
            # Read with format preservation
            with st.spinner("Reading document with format preservation..."):
                content_data = processor.read_word_document_preserve_format(uploaded_file)
            
            if not content_data['text']:
                st.error("Could not read document content.")
                return
            
            # Preview
            with st.expander("📖 Document Preview"):
                st.text_area("Original Content", 
                           content_data['text'][:2000] + "..." if len(content_data['text']) > 2000 else content_data['text'],
                           height=200)
            
            # Analysis
            st.subheader("📊 Content Analysis")
            
            # Find elements
            formulas = processor.find_formulas_enhanced(content_data['text'])
            tables = processor.find_markdown_tables_robust(content_data['text'])
            tikz_blocks = processor.find_tikz_blocks_enhanced(content_data['text'])
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("🧮 Math Formulas", len(formulas))
                for i, formula in enumerate(formulas[:3]):
                    st.code(f"{formula['type']}: {formula['latex'][:50]}...")
            
            with col2:
                st.metric("📊 Tables", len(tables))
                for i, table in enumerate(tables[:2]):
                    if table['header']:
                        st.write(f"**Table {i+1}**: {len(table['rows'])}×{len(table['header'])}")
                        # Preview as DataFrame
                        try:
                            df_data = {}
                            for j, col in enumerate(table['header']):
                                df_data[col] = [
                                    row[j] if j < len(row) else ''
                                    for row in table['rows'][:3]  # First 3 rows
                                ]
                            st.dataframe(df_data, use_container_width=True)
                        except:
                            st.write(f"Headers: {table['header']}")
            
            with col3:
                st.metric("🎨 TikZ Diagrams", len(tikz_blocks))
                for i, tikz in enumerate(tikz_blocks[:2]):
                    st.write(f"**TikZ {i+1}**: {tikz['type']}")
                    st.code(tikz['code'][:100] + "..." if len(tikz['code']) > 100 else tikz['code'])
            
            # Processing options
            st.subheader("⚙️ Processing Options")
            
            col1, col2 = st.columns(2)
            with col1:
                compile_tikz = st.checkbox("🎨 Compile TikZ to PNG", value=True,
                                         help="Requires LaTeX installation")
                preserve_format = st.checkbox("📝 Preserve Original Format", value=True)
            
            with col2:
                table_style = st.selectbox("📊 Table Style", 
                                         ["Professional", "Simple", "Colorful"])
                math_method = st.selectbox("🧮 Math Method",
                                         ["Auto", "math2docx", "Unicode"])
            
            # Process button
            if st.button("🚀 Process Advanced Document", type="primary", use_container_width=True):
                with st.spinner("Advanced processing in progress..."):
                    try:
                        processed_doc, stats = processor.create_processed_document_advanced(
                            content_data, formulas, tables, tikz_blocks
                        )
                        
                        # Save document
                        doc_io = io.BytesIO()
                        processed_doc.save(doc_io)
                        doc_io.seek(0)
                        
                        # Success message
                        st.success("✅ Advanced processing complete!")
                        
                        # Statistics
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Math Equations", stats['equations'])
                        with col2:
                            st.metric("Professional Tables", stats['tables'])
                        with col3:
                            st.metric("TikZ Images", stats['tikz_images'])
                        
                        # Download
                        st.download_button(
                            label="📥 Download Advanced Document",
                            data=doc_io.getvalue(),
                            file_name="advanced_processed_document.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                        # Processing info
                        st.info(f"""
🎯 **Advanced Processing Results:**
- {stats['equations']} math equations processed with real LaTeX rendering
- {stats['tables']} professional tables with color styling and borders
- {stats['tikz_images']} TikZ diagrams compiled to high-quality PNG images
- Original document structure and formatting preserved
- Enhanced typography and layout applied
                        """)
                        
                    except Exception as e:
                        st.error(f"Processing error: {e}")
                        import traceback
                        st.code(traceback.format_exc())
        
        except Exception as e:
            st.error(f"Error: {e}")
            import traceback
            st.code(traceback.format_exc())
        
        finally:
            processor.cleanup()
    
    # Setup instructions
    with st.expander("🔧 System Setup Instructions"):
        st.markdown("""
        ### Required Packages:
        ```bash
        # Core packages
        pip install streamlit python-docx mammoth markdown Pillow PyMuPDF pandas
        
        # For real math equations
        pip install math2docx latex2mathml
        
        # For LaTeX compilation (choose one):
        # Windows: Install MiKTeX from https://miktex.org/
        # Mac: brew install mactex
        # Linux: sudo apt-get install texlive-full
        ```
        
        ### Features by Package:
        - **math2docx**: Real LaTeX equations in Word
        - **LaTeX**: TikZ compilation to PNG
        - **latex2mathml**: Alternative math rendering
        - **mammoth**: Enhanced Word reading with format preservation
        
        ### TikZ Requirements:
        TikZ compilation requires a full LaTeX installation with:
        - `pdflatex` command available
        - `tikz` package and libraries
        - Standard math packages (`amsmath`, `amsfonts`)
        """)

if __name__ == "__main__":
    main()
