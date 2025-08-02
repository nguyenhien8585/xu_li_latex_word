import streamlit as st
import re
import io
import os
import tempfile
from pathlib import Path
import subprocess

# Core imports
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.shared import OxmlElement, qn
    import mammoth
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
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

class FinalSolutionProcessor:
    """FINAL SOLUTION - Khắc phục hoàn toàn vấn đề math detection"""
    
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        
    def tikz2image(self, tikz_code, dpi=300, width=None, height=None):
        """FIXED TikZ compilation function with enhanced error handling"""
        tex_template = f"""
\\documentclass[tikz,border=10pt]{{standalone}}
\\usepackage{{amsmath,amssymb,amsfonts}}
\\usepackage{{tikz,tkz-euclide,tkz-tab,pgfplots}}
\\usepackage{{xcolor}}
\\pgfplotsset{{compat=newest}}
\\usetikzlibrary{{shapes,arrows,positioning,calc,patterns,decorations.pathreplacing,shadows,3d,intersections}}
\\begin{{document}}
{tikz_code}
\\end{{document}}
"""
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tex_path = os.path.join(tmpdir, "tikz_img.tex")
                pdf_path = os.path.join(tmpdir, "tikz_img.pdf")
                png_path = os.path.join(tmpdir, "tikz_img.png")
                
                # Write LaTeX file
                with open(tex_path, "w", encoding="utf-8") as f:
                    f.write(tex_template)
                
                # Compile with pdflatex - enhanced command
                proc = subprocess.run(
                    ["pdflatex", "-shell-escape", "-interaction=nonstopmode", "-output-directory", tmpdir, tex_path],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=120,  # Increased timeout
                    cwd=tmpdir
                )
                
                # Check if PDF was created
                if proc.returncode == 0 and os.path.exists(pdf_path):
                    st.success(f"✅ TikZ compiled successfully!")
                    
                    # Convert PDF to PNG with high quality
                    if PDF2IMAGE_AVAILABLE:
                        try:
                            imgs = convert_from_path(pdf_path, dpi=dpi, 
                                                   size=(width, height) if width and height else None)
                            imgs[0].save(png_path, "PNG", optimize=True)
                            st.info("Used pdf2image for conversion")
                        except Exception as e:
                            st.warning(f"pdf2image failed: {e}, using PyMuPDF fallback")
                            # Fallback to PyMuPDF
                            doc = fitz.open(pdf_path)
                            page = doc[0]
                            mat = fitz.Matrix(dpi/72, dpi/72)
                            pix = page.get_pixmap(matrix=mat, alpha=False)
                            pix.save(png_path)
                            doc.close()
                    else:
                        # Use PyMuPDF
                        doc = fitz.open(pdf_path)
                        page = doc[0]
                        mat = fitz.Matrix(dpi/72, dpi/72)
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        pix.save(png_path)
                        doc.close()
                        st.info("Used PyMuPDF for conversion")
                    
                    # Verify PNG was created
                    if os.path.exists(png_path) and os.path.getsize(png_path) > 1000:
                        with open(png_path, "rb") as f:
                            img_bytes = f.read()
                        st.success(f"✅ PNG created successfully! Size: {len(img_bytes)} bytes")
                        return img_bytes, None
                    else:
                        return None, "PNG file not created or too small"
                else:
                    # Compilation failed
                    error_msg = proc.stderr.decode("utf-8") if proc.stderr else "Unknown LaTeX error"
                    st.error(f"❌ TikZ compilation failed!")
                    st.code(f"LaTeX Error: {error_msg[:500]}...")
                    return None, error_msg
                
        except subprocess.TimeoutExpired:
            return None, "TikZ compilation timeout (120s)"
        except Exception as e:
            st.error(f"❌ TikZ compilation exception: {e}")
            return None, str(e)
    
    def read_document_with_raw_content(self, uploaded_file):
        """Đọc document với RAW content để bắt được math"""
        try:
            # Method 1: mammoth extract raw text
            uploaded_file.seek(0)
            result = mammoth.extract_raw_text(uploaded_file)
            raw_text = result.value
            
            # Method 2: mammoth extract with HTML
            uploaded_file.seek(0)
            html_result = mammoth.convert_to_html(uploaded_file)
            html_content = html_result.value
            
            # Method 3: python-docx for structure
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            
            # Extract ALL runs and their text to catch math
            all_content = []
            for paragraph in doc.paragraphs:
                para_content = ""
                for run in paragraph.runs:
                    para_content += run.text
                all_content.append(para_content)
            
            docx_text = '\n'.join(all_content)
            
            # Extract tables
            existing_tables = []
            for table in doc.tables:
                table_data = []
                for row in table.rows:
                    row_data = []
                    for cell in row.cells:
                        cell_text = cell.text.strip()
                        row_data.append(cell_text)
                    if any(cell.strip() for cell in row_data):
                        table_data.append(row_data)
                
                if table_data and len(table_data) >= 2:
                    existing_tables.append({
                        'header': table_data[0],
                        'rows': table_data[1:],
                        'source': 'existing'
                    })
            
            # Combine all text sources
            combined_text = f"{raw_text}\n\n{docx_text}"
            
            return {
                'raw_text': raw_text,
                'html_content': html_content,
                'docx_text': docx_text,
                'combined_text': combined_text,
                'existing_tables': existing_tables
            }
            
        except Exception as e:
            st.error(f"Error reading document: {e}")
            return {
                'raw_text': '', 
                'html_content': '', 
                'docx_text': '',
                'combined_text': '',
                'existing_tables': []
            }
    
    def detect_math_from_original_file(self, text_content):
        """CRITICAL: Detect math từ original file content"""
        
        # Show debug info
        st.write("**Debug - Searching for math in content:**")
        st.write(f"Content length: {len(text_content)}")
        
        # Show sample content
        if len(text_content) > 0:
            st.write("**Sample content (first 500 chars):**")
            st.code(text_content[:500])
        
        math_expressions = []
        
        # COMPREHENSIVE math patterns based on what we see in original files
        patterns = [
            # Standard LaTeX patterns
            (r'\$([^$\n]+)\$', 'inline_math'),
            (r'\$\$([^$]+)\$\$', 'display_math'),
            (r'\\\[([^\]]+)\\\]', 'display_math'),
            (r'\\\(([^)]+)\\\)', 'inline_math'),
            
            # Formula blocks
            (r'\[FORMULA:\s*([^\]]+)\]', 'display_math'),
            (r'\[MATH:\s*([^\]]+)\]', 'inline_math'),
            
            # Word-specific patterns (from conversion)
            (r'\$\\mathit\{([^}]+)\}\$', 'inline_math'),
            (r'\$\\mathrm\{([^}]+)\}\$', 'inline_math'),
            (r'\$\\text\{([^}]+)\}\$', 'inline_math'),
            
            # Common math expressions
            (r'\$([A-Za-z]+)\s*=\s*([^$]+)\$', 'inline_math'),
            (r'\$\\frac\{[^}]+\}\{[^}]+\}\$', 'inline_math'),
            (r'\$\\sqrt\{[^}]+\}\$', 'inline_math'),
            (r'\$[A-Za-z]_\{[^}]+\}\$', 'inline_math'),
            (r'\$[A-Za-z]\^*\{[^}]+\}\$', 'inline_math'),
            
            # Greek letters and symbols
            (r'\$\\[a-zA-Z]+\$', 'inline_math'),
            (r'\$[αβγδεζηθικλμνξπρστυφχψω]+\$', 'inline_math'),
            
            # Numbers and simple expressions
            (r'\$\d+\$', 'inline_math'),
            (r'\$\d+[\+\-\*/]\d+\$', 'inline_math'),
            
            # Fractions in different formats
            (r'\$\d+/\d+\$', 'inline_math'),
            (r'\$\([^)]+\)/\([^)]+\)\$', 'inline_math'),
        ]
        
        # Search for each pattern
        total_found = 0
        for pattern, math_type in patterns:
            matches = list(re.finditer(pattern, text_content, re.IGNORECASE))
            for match in matches:
                content = match.group(1) if match.groups() else match.group(0)
                math_expressions.append({
                    'original': match.group(0),
                    'content': content,
                    'type': math_type,
                    'start': match.start(),
                    'end': match.end(),
                    'pattern': pattern
                })
                total_found += 1
        
        st.write(f"**Total math expressions found: {total_found}**")
        
        # Remove duplicates
        unique_expressions = []
        seen = set()
        for expr in math_expressions:
            key = (expr['start'], expr['end'])
            if key not in seen:
                seen.add(key)
                unique_expressions.append(expr)
        
        unique_expressions.sort(key=lambda x: x['start'])
        
        # Show found expressions
        if unique_expressions:
            st.write("**Found math expressions:**")
            for i, expr in enumerate(unique_expressions[:10]):  # Show first 10
                st.write(f"{i+1}. {expr['type']}: `{expr['original']}`")
        else:
            st.warning("No math expressions detected!")
        
        return unique_expressions
    
    def detect_markdown_tables_aggressive(self, text):
        """Aggressive table detection"""
        tables = []
        
        # Pattern 1: Standard markdown
        pattern1 = r'(\|[^\n\r]+\|(?:\r?\n|\r))+(\|[\s\-:|]+\|(?:\r?\n|\r))?(\|[^\n\r]+\|(?:\r?\n|\r)?)*'
        for match in re.finditer(pattern1, text, re.MULTILINE):
            table_text = match.group(0).strip()
            parsed = self._parse_table_aggressive(table_text)
            if parsed:
                parsed.update({
                    'original': table_text,
                    'start': match.start(),
                    'end': match.end()
                })
                tables.append(parsed)
        
        # Pattern 2: Simple space-separated
        lines = text.split('\n')
        i = 0
        while i < len(lines):
            if '|' in lines[i] and lines[i].count('|') >= 2:
                table_lines = [lines[i]]
                j = i + 1
                while j < len(lines) and ('|' in lines[j] or re.match(r'^[\s\-:|]+$', lines[j])):
                    table_lines.append(lines[j])
                    j += 1
                
                if len(table_lines) >= 2:
                    table_text = '\n'.join(table_lines)
                    parsed = self._parse_table_aggressive(table_text)
                    if parsed:
                        start_pos = text.find(table_text)
                        if start_pos != -1:
                            parsed.update({
                                'original': table_text,
                                'start': start_pos,
                                'end': start_pos + len(table_text)
                            })
                            tables.append(parsed)
                i = j
            else:
                i += 1
        
        return tables
    
    def _parse_table_aggressive(self, table_text):
        """Aggressive table parsing"""
        lines = [line.strip() for line in table_text.split('\n') if line.strip()]
        
        if len(lines) < 2:
            return None
        
        # Find header
        header_line = lines[0]
        if '|' not in header_line:
            return None
        
        header_parts = header_line.split('|')
        header = [part.strip() for part in header_parts if part.strip()]
        
        if len(header) < 2:
            return None
        
        # Skip separator if exists
        data_start = 1
        if len(lines) > 1 and re.match(r'^[\s\-:|]+$', lines[1].replace('|', '')):
            data_start = 2
        
        # Parse rows
        rows = []
        for line in lines[data_start:]:
            if '|' in line:
                row_parts = line.split('|')
                row = [part.strip() for part in row_parts if part.strip() or len(row_parts) > len(header)]
                
                # Adjust length
                while len(row) < len(header):
                    row.append('')
                row = row[:len(header)]
                
                if any(cell.strip() for cell in row):
                    rows.append(row)
        
        return {'header': header, 'rows': rows} if rows else None
    
    def detect_tikz_comprehensive(self, text):
        """Comprehensive TikZ detection"""
        tikz_blocks = []
        
        patterns = [
            r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}',
            r'\\begin\{tikzpicture\}\[.*?\].*?\\end\{tikzpicture\}',
        ]
        
        for pattern in patterns:
            for match in re.finditer(pattern, text, re.DOTALL):
                code = re.sub(r'\\begin\{tikzpicture\}(?:\[.*?\])?', '', match.group(0))
                code = re.sub(r'\\end\{tikzpicture\}', '', code).strip()
                
                if code:
                    tikz_blocks.append({
                        'original': match.group(0),
                        'code': code,
                        'start': match.start(),
                        'end': match.end()
                    })
        
        return tikz_blocks
    
    def create_math_paragraph(self, doc, content, is_display=False):
        """Create paragraph with math - FIXED centering and line breaks"""
        # ALWAYS create new paragraph for math (both inline and display)
        para = doc.add_paragraph()
        
        # Set alignment based on type
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add spacing for display math
        if is_display:
            para.paragraph_format.space_before = Pt(12)
            para.paragraph_format.space_after = Pt(12)
            para.paragraph_format.line_spacing = 1.2
        else:
            # Even inline math gets some spacing for visibility
            para.paragraph_format.space_before = Pt(3)
            para.paragraph_format.space_after = Pt(3)
        
        # Try math2docx first
        if MATH2DOCX_AVAILABLE:
            try:
                cleaned = self._clean_latex(content)
                math2docx.add_math(para, cleaned)
                return True
            except Exception as e:
                st.warning(f"math2docx failed: {e}")
        
        # Enhanced Unicode fallback with better formatting
        unicode_math = self._latex_to_unicode(content)
        run = para.add_run(unicode_math)
        run.font.name = 'Cambria Math'
        run.font.size = Pt(14) if is_display else Pt(12)
        run.font.color.rgb = RGBColor(0, 51, 102)
        run.bold = True if is_display else False
        
        # Add visual brackets for better distinction
        if is_display:
            # Add some visual enhancement for display math
            run.font.size = Pt(16)
            run.bold = True
        
        return True
    
    def _clean_latex(self, latex):
        """Clean LaTeX for processing"""
        cleaned = latex.strip()
        
        replacements = {
            r'\\mathit\{([^}]+)\}': r'\1',
            r'\\mathrm\{([^}]+)\}': r'\1',
            r'\\text\{([^}]+)\}': r'\1',
            r'\s+': ' ',
        }
        
        for pattern, replacement in replacements.items():
            cleaned = re.sub(pattern, replacement, cleaned)
        
        return cleaned
    
    def _latex_to_unicode(self, latex):
        """Convert LaTeX to Unicode"""
        symbols = {
            r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ',
            r'\\pi': 'π', r'\\sigma': 'σ', r'\\theta': 'θ', r'\\lambda': 'λ',
            r'\\mu': 'μ', r'\\nu': 'ν', r'\\omega': 'ω', r'\\phi': 'φ',
            r'\\times': '×', r'\\div': '÷', r'\\pm': '±', r'\\infty': '∞',
            r'\\leq': '≤', r'\\geq': '≥', r'\\neq': '≠', r'\\approx': '≈',
            r'\\sqrt': '√', r'\\sum': '∑', r'\\int': '∫',
        }
        
        result = latex
        for pattern, replacement in symbols.items():
            result = re.sub(pattern, replacement, result)
        
        # Handle fractions
        result = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'(\1)/(\2)', result)
        
        # Handle superscripts/subscripts  
        result = re.sub(r'\^{([^}]+)}', lambda m: self._to_superscript(m.group(1)), result)
        result = re.sub(r'_{([^}]+)}', lambda m: self._to_subscript(m.group(1)), result)
        
        return result
    
    def _to_superscript(self, text):
        """Convert to superscript"""
        sup_map = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵',
                   '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', '+': '⁺', '-': '⁻'}
        return ''.join(sup_map.get(c, c) for c in text)
    
    def _to_subscript(self, text):
        """Convert to subscript"""
        sub_map = {'0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅',
                   '6': '₆', '7': '₇', '8': '₈', '9': '₉', '+': '₊', '-': '₋'}
        return ''.join(sub_map.get(c, c) for c in text)
    
    def create_professional_table(self, doc, table_data):
        """Create professional table"""
        header = table_data['header']
        rows = table_data['rows']
        
        table = doc.add_table(rows=len(rows) + 1, cols=len(header))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Header
        header_row = table.rows[0]
        for i, text in enumerate(header):
            cell = header_row.cells[i]
            cell.text = str(text)
            
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(11)
                    run.font.color.rgb = RGBColor(255, 255, 255)
            
            shading = OxmlElement("w:shd")
            shading.set(qn("w:fill"), "366092")
            cell._tc.get_or_add_tcPr().append(shading)
        
        # Data rows
        for row_idx, row_data in enumerate(rows):
            table_row = table.rows[row_idx + 1]
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < len(header):
                    cell = table_row.cells[col_idx]
                    cell.text = str(cell_data)
                    
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
                    
                    if row_idx % 2 == 0:
                        shading = OxmlElement("w:shd")
                        shading.set(qn("w:fill"), "F8F9FA")
                        cell._tc.get_or_add_tcPr().append(shading)
        
        return table
    
    def create_final_document(self, content_data, math_expressions, tables, tikz_blocks):
        """Create final corrected document"""
        doc = Document()
        
        # Title
        title = doc.add_heading('FINAL SOLUTION - Corrected Mathematical Document', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Use the most complete text source
        text_sources = [
            content_data.get('combined_text', ''),
            content_data.get('raw_text', ''),
            content_data.get('docx_text', '')
        ]
        
        # Choose the longest text (most complete)
        original_text = max(text_sources, key=len)
        
        if not original_text:
            st.error("No text content found!")
            return doc, {}
        
        # Collect all elements
        all_elements = []
        
        for expr in math_expressions:
            all_elements.append({
                'type': 'math',
                'start': expr['start'],
                'end': expr['end'],
                'data': expr
            })
        
        for table in tables:
            all_elements.append({
                'type': 'table',
                'start': table.get('start', 0),
                'end': table.get('end', 0),
                'data': table
            })
        
        for tikz in tikz_blocks:
            all_elements.append({
                'type': 'tikz',
                'start': tikz['start'],
                'end': tikz['end'],
                'data': tikz
            })
        
        all_elements.sort(key=lambda x: x['start'])
        
        # Process document
        current_pos = 0
        stats = {'math': 0, 'tables': 0, 'tikz': 0}
        
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
            if element['type'] == 'math':
                # FIXED: Always treat math as display (centered, on new line)
                if self.create_math_paragraph(doc, element['data']['content'], is_display=True):
                    stats['math'] += 1
                    
            elif element['type'] == 'table':
                table = self.create_professional_table(doc, element['data'])
                if table:
                    stats['tables'] += 1
                    doc.add_paragraph()  # Add spacing after table
                    
            elif element['type'] == 'tikz':
                st.write(f"🎨 Processing TikZ block {len([x for x in all_elements[:all_elements.index(element)+1] if x['type'] == 'tikz'])}...")
                
                # Enhanced TikZ processing with debug
                tikz_code = element['data']['code']
                st.code(f"TikZ Code Preview:\n{tikz_code[:200]}...")
                
                img_bytes, error = self.tikz2image(tikz_code, dpi=300)
                
                if img_bytes:
                    try:
                        # Save image temporarily
                        tikz_num = stats.get('tikz', 0) + 1
                        img_path = os.path.join(self.temp_dir, f'tikz_{tikz_num}.png')
                        with open(img_path, 'wb') as f:
                            f.write(img_bytes)
                        
                        # Add to document with enhanced formatting
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Add some spacing before image
                        p.paragraph_format.space_before = Pt(12)
                        p.paragraph_format.space_after = Pt(6)
                        
                        run = p.add_run()
                        run.add_picture(img_path, width=Inches(6))
                        
                        # Add professional caption
                        caption_p = doc.add_paragraph()
                        caption_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        caption_p.paragraph_format.space_after = Pt(12)
                        
                        caption_run = caption_p.add_run(f"Figure {tikz_num}: TikZ Diagram")
                        caption_run.italic = True
                        caption_run.font.size = Pt(10)
                        caption_run.font.color.rgb = RGBColor(100, 100, 100)
                        caption_run.font.name = 'Times New Roman'
                        
                        stats['tikz'] = tikz_num
                        st.success(f"✅ TikZ {tikz_num} successfully added to document!")
                        
                    except Exception as e:
                        st.error(f"❌ Error inserting TikZ image: {e}")
                        # Add error message to document
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p.add_run(f"[TikZ Image Insert Error: {str(e)}]")
                        run.italic = True
                        run.font.color.rgb = RGBColor(200, 0, 0)
                else:
                    st.error(f"❌ TikZ compilation failed: {error}")
                    # Add error message to document
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    error_text = f"[TikZ Compilation Failed: {error[:100]}...]"
                    run = p.add_run(error_text)
                    run.italic = True
                    run.font.color.rgb = RGBColor(200, 0, 0)
            
            current_pos = element['end']
        
        # Add remaining text
        if current_pos < len(original_text):
            remaining = original_text[current_pos:]
            if remaining.strip():
                paragraphs = remaining.split('\n')
                for para_text in paragraphs:
                    if para_text.strip():
                        p = doc.add_paragraph(para_text.strip())
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        return doc, stats
    
    def cleanup(self):
        """Cleanup"""
        import shutil
        try:
            shutil.rmtree(self.temp_dir)
        except:
            pass

def main():
    st.set_page_config(page_title="FINAL SOLUTION", page_icon="🎯", layout="wide")
    
    st.title("🎯 FINAL SOLUTION - Math Document Processor")
    st.markdown("""
    **GIẢI PHÁP CUỐI CÙNG** - Khắc phục hoàn toàn vấn đề math detection:
    - 🎯 **DEBUG MODE**: Hiển thị chi tiết quá trình detect math
    - 🎯 **MULTIPLE READ METHODS**: Đọc document bằng nhiều cách
    - 🎯 **COMPREHENSIVE PATTERNS**: Patterns toàn diện cho math detection
    - 🎯 **AGGRESSIVE TABLE DETECTION**: Bắt tất cả table formats
    - 🎯 **WORKING TIKZ**: TikZ compilation thành công
    """)
    
    uploaded_file = st.file_uploader("📁 Upload Word Document (.docx)", type=['docx'])
    
    if uploaded_file:
        processor = FinalSolutionProcessor()
        
        try:
            with st.spinner("Reading document with multiple methods..."):
                content_data = processor.read_document_with_raw_content(uploaded_file)
            
            # Show content sources
            st.subheader("📖 Content Sources Analysis")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.write("**Raw Text (mammoth)**")
                st.write(f"Length: {len(content_data.get('raw_text', ''))}")
                if content_data.get('raw_text'):
                    st.text_area("Raw preview", content_data['raw_text'][:300], height=100)
            
            with col2:
                st.write("**DocX Text (python-docx)**") 
                st.write(f"Length: {len(content_data.get('docx_text', ''))}")
                if content_data.get('docx_text'):
                    st.text_area("DocX preview", content_data['docx_text'][:300], height=100)
            
            with col3:
                st.write("**Combined Text**")
                st.write(f"Length: {len(content_data.get('combined_text', ''))}")
                if content_data.get('combined_text'):
                    st.text_area("Combined preview", content_data['combined_text'][:300], height=100)
            
            # Analyze content
            st.subheader("🔍 FINAL Analysis with Debug")
            
            # Choose best text source for analysis
            text_for_analysis = content_data.get('combined_text', '') or content_data.get('raw_text', '') or content_data.get('docx_text', '')
            
            if not text_for_analysis:
                st.error("No text content found in document!")
                return
            
            # Detect math with debug
            st.write("### 🧮 Math Detection (with debug)")
            math_expressions = processor.detect_math_from_original_file(text_for_analysis)
            
            # Detect tables
            st.write("### 📊 Table Detection")
            markdown_tables = processor.detect_markdown_tables_aggressive(text_for_analysis)
            all_tables = content_data['existing_tables'] + markdown_tables
            st.write(f"Found {len(all_tables)} tables ({len(content_data['existing_tables'])} existing + {len(markdown_tables)} markdown)")
            
            # Detect TikZ
            st.write("### 🎨 TikZ Detection")
            tikz_blocks = processor.detect_tikz_comprehensive(text_for_analysis)
            st.write(f"Found {len(tikz_blocks)} TikZ blocks")
            
            # Summary
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("🧮 Math Expressions", len(math_expressions))
            with col2:
                st.metric("📊 Tables", len(all_tables))
            with col3:
                st.metric("🎨 TikZ Blocks", len(tikz_blocks))
            
            # Process button
            if st.button("🎯 FINAL PROCESSING", type="primary", use_container_width=True):
                with st.spinner("Creating final corrected document..."):
                    try:
                        doc, stats = processor.create_final_document(
                            content_data, math_expressions, all_tables, tikz_blocks
                        )
                        
                        # Save
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        st.success("✅ FINAL SOLUTION Complete!")
                        
                        # Stats
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Math Processed", stats['math'])
                        with col2:
                            st.metric("Tables Created", stats['tables'])
                        with col3:
                            st.metric("TikZ Compiled", stats['tikz'])
                        
                        # Download
                        st.download_button(
                            label="📥 Download FINAL CORRECTED Document",
                            data=doc_io.getvalue(),
                            file_name="FINAL_SOLUTION_document.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                        st.info(f"""
🎯 **FINAL SOLUTION RESULTS:**
- {stats['math']} math expressions successfully processed
- {stats['tables']} professional tables created  
- {stats['tikz']} TikZ diagrams compiled
- All content preserved and enhanced
- Professional formatting applied
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

if __name__ == "__main__":
    main()
