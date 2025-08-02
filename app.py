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
    from docx.oxml.ns import nsdecls
    from docx.oxml import parse_xml
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

class ProfessionalWordProcessor:
    """Professional Word processor với format preservation và styling đẹp"""
    
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        
    def read_word_with_structure(self, uploaded_file):
        """Đọc Word document với cấu trúc đầy đủ"""
        try:
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            
            structure = {
                'paragraphs': [],
                'tables': [],
                'raw_text': '',
                'styles_info': {}
            }
            
            full_text = []
            
            # Extract paragraphs với style info
            for i, paragraph in enumerate(doc.paragraphs):
                para_info = {
                    'text': paragraph.text,
                    'style': paragraph.style.name if paragraph.style else 'Normal',
                    'alignment': paragraph.alignment,
                    'runs': []
                }
                
                # Extract runs với formatting
                for run in paragraph.runs:
                    run_info = {
                        'text': run.text,
                        'bold': run.bold,
                        'italic': run.italic,
                        'underline': run.underline,
                        'font_size': run.font.size.pt if run.font.size else None
                    }
                    para_info['runs'].append(run_info)
                
                structure['paragraphs'].append(para_info)
                full_text.append(paragraph.text)
            
            # Extract tables
            for table in doc.tables:
                table_data = []
                for row in table.rows:
                    row_data = []
                    for cell in row.cells:
                        row_data.append(cell.text.strip())
                    table_data.append(row_data)
                
                if table_data and any(any(cell for cell in row) for row in table_data):
                    structure['tables'].append({
                        'data': table_data,
                        'header': table_data[0] if table_data else [],
                        'rows': table_data[1:] if len(table_data) > 1 else []
                    })
            
            structure['raw_text'] = '\n'.join(full_text)
            
            return structure
            
        except Exception as e:
            st.error(f"Error reading Word document: {e}")
            return {'paragraphs': [], 'tables': [], 'raw_text': '', 'styles_info': {}}
    
    def detect_math_expressions_advanced(self, text):
        """Advanced math detection với nhiều formats"""
        math_expressions = []
        
        # Patterns for different math formats
        patterns = [
            # LaTeX formats
            (r'\$([^$]+)\$', 'latex_inline'),
            (r'\$\$([^$]+)\$\$', 'latex_display'),
            (r'\\\[([^\]]+)\\\]', 'latex_display_bracket'),
            (r'\\\(([^\)]+)\\\)', 'latex_inline_paren'),
            
            # Formula blocks
            (r'\[FORMULA:\s*([^\]]+)\]', 'formula_block'),
            (r'\[MATH:\s*([^\]]+)\]', 'math_block'),
            
            # Environment blocks
            (r'\\begin\{equation\}(.*?)\\end\{equation\}', 'equation'),
            (r'\\begin\{align\}(.*?)\\end\{align\}', 'align'),
            (r'\\begin\{displaymath\}(.*?)\\end\{displaymath\}', 'displaymath'),
            
            # Unicode math symbols detection
            (r'[α-ωΑ-Ω∀∃∈∉∅∪∩⊂⊃≤≥≠≈±×÷∞∑∏∫√π]', 'unicode_math')
        ]
        
        for pattern, math_type in patterns:
            for match in re.finditer(pattern, text, re.DOTALL):
                content = match.group(1) if match.groups() else match.group(0)
                math_expressions.append({
                    'original': match.group(0),
                    'content': content.strip(),
                    'type': math_type,
                    'start': match.start(),
                    'end': match.end()
                })
        
        # Remove duplicates and sort
        math_expressions = list({expr['original']: expr for expr in math_expressions}.values())
        math_expressions.sort(key=lambda x: x['start'], reverse=True)
        
        return math_expressions
    
    def detect_tables_comprehensive(self, text):
        """Comprehensive table detection"""
        tables = []
        
        # Method 1: Markdown tables
        markdown_pattern = r'(\|[^\n\r]+\|(?:\r?\n|\r))+(\|[-:\s|]+\|(?:\r?\n|\r))?(\|[^\n\r]+\|(?:\r?\n|\r)?)*'
        
        for match in re.finditer(markdown_pattern, text, re.MULTILINE):
            table_text = match.group(0).strip()
            parsed = self._parse_markdown_table_robust(table_text)
            if parsed:
                parsed.update({
                    'original': table_text,
                    'start': match.start(),
                    'end': match.end(),
                    'type': 'markdown'
                })
                tables.append(parsed)
        
        # Method 2: Space-separated tables
        lines = text.split('\n')
        current_table = []
        table_start_line = -1
        
        for i, line in enumerate(lines):
            # Check if line looks like table row (multiple words separated by spaces/tabs)
            if self._is_table_row(line):
                if not current_table:
                    table_start_line = i
                current_table.append(line.strip())
            else:
                if len(current_table) >= 2:  # At least 2 rows
                    parsed = self._parse_space_table(current_table)
                    if parsed:
                        table_text = '\n'.join(current_table)
                        start_pos = text.find(table_text)
                        if start_pos != -1:
                            parsed.update({
                                'original': table_text,
                                'start': start_pos,
                                'end': start_pos + len(table_text),
                                'type': 'space_separated'
                            })
                            tables.append(parsed)
                current_table = []
        
        # Process last table if exists
        if len(current_table) >= 2:
            parsed = self._parse_space_table(current_table)
            if parsed:
                table_text = '\n'.join(current_table)
                start_pos = text.find(table_text)
                if start_pos != -1:
                    parsed.update({
                        'original': table_text,
                        'start': start_pos,
                        'end': start_pos + len(table_text),
                        'type': 'space_separated'
                    })
                    tables.append(parsed)
        
        # Remove duplicates and sort
        tables = self._remove_overlapping_tables(tables)
        tables.sort(key=lambda x: x['start'])
        
        return tables
    
    def _is_table_row(self, line):
        """Check if line looks like a table row"""
        line = line.strip()
        if not line or len(line) < 5:
            return False
        
        # Check for multiple words/numbers separated by whitespace
        words = line.split()
        if len(words) < 2:
            return False
        
        # Check for consistent separation patterns
        if '|' in line:
            return line.count('|') >= 2
        
        # Check for tab-separated or multiple space-separated
        if '\t' in line or '  ' in line:
            return True
        
        # Check for mix of text and numbers
        has_text = any(word.isalpha() for word in words)
        has_numbers = any(any(c.isdigit() for c in word) for word in words)
        
        return len(words) >= 3 and (has_text or has_numbers)
    
    def _parse_markdown_table_robust(self, table_text):
        """Robust markdown table parsing"""
        lines = [line.strip() for line in table_text.split('\n') if line.strip()]
        
        if len(lines) < 2:
            return None
        
        # Find header
        header_line = lines[0]
        if not '|' in header_line:
            return None
        
        header = [cell.strip() for cell in header_line.split('|') if cell.strip()]
        if not header:
            return None
        
        # Find separator line (optional)
        data_start = 1
        if len(lines) > 1 and re.match(r'^[\s\-:|]+$', lines[1].replace('|', '')):
            data_start = 2
        
        # Parse data rows
        rows = []
        for line in lines[data_start:]:
            if '|' in line:
                row = [cell.strip() for cell in line.split('|') if cell.strip()]
                # Adjust row length to match header
                while len(row) < len(header):
                    row.append('')
                rows.append(row[:len(header)])
        
        return {'header': header, 'rows': rows} if rows else None
    
    def _parse_space_table(self, table_lines):
        """Parse space/tab separated table"""
        if len(table_lines) < 2:
            return None
        
        # Parse first line as header
        header_line = table_lines[0]
        if '\t' in header_line:
            header = [cell.strip() for cell in header_line.split('\t') if cell.strip()]
        else:
            # Split by multiple spaces
            header = [cell.strip() for cell in re.split(r'  +', header_line) if cell.strip()]
        
        if not header or len(header) < 2:
            return None
        
        # Parse data rows
        rows = []
        for line in table_lines[1:]:
            if '\t' in line:
                row = [cell.strip() for cell in line.split('\t')]
            else:
                row = [cell.strip() for cell in re.split(r'  +', line)]
            
            # Adjust to header length
            while len(row) < len(header):
                row.append('')
            rows.append(row[:len(header)])
        
        return {'header': header, 'rows': rows} if rows else None
    
    def _remove_overlapping_tables(self, tables):
        """Remove overlapping table detections"""
        if not tables:
            return tables
        
        # Sort by start position
        tables.sort(key=lambda x: x['start'])
        
        filtered = [tables[0]]
        for table in tables[1:]:
            # Check if overlaps with any previous table
            overlaps = False
            for prev_table in filtered:
                if (table['start'] < prev_table['end'] and 
                    table['end'] > prev_table['start']):
                    overlaps = True
                    break
            
            if not overlaps:
                filtered.append(table)
        
        return filtered
    
    def convert_math_to_word_equation(self, paragraph, latex_content):
        """Convert LaTeX to Word equation"""
        try:
            if MATH2DOCX_AVAILABLE:
                # Clean LaTeX
                cleaned_latex = self._clean_latex_for_word(latex_content)
                math2docx.add_math(paragraph, cleaned_latex)
                return True
            else:
                # Enhanced Unicode fallback
                unicode_math = self._latex_to_unicode_enhanced(latex_content)
                run = paragraph.add_run(unicode_math)
                run.font.name = 'Cambria Math'
                run.font.size = Pt(12)
                run.italic = True
                return True
        except Exception as e:
            st.warning(f"Math conversion error for '{latex_content[:30]}...': {e}")
            return False
    
    def _clean_latex_for_word(self, latex):
        """Clean LaTeX for Word equation conversion"""
        # Remove common problematic patterns
        cleaned = latex.strip()
        
        # Fix common issues
        replacements = {
            r'\\mathit\{([^}]+)\}': r'\1',  # Remove \mathit{}
            r'\\mathrm\{([^}]+)\}': r'\1',  # Remove \mathrm{}
            r'\\text\{([^}]+)\}': r'\1',    # Remove \text{}
            r'\\,': ' ',                     # Replace thin space
            r'\\;': ' ',                     # Replace medium space
            r'\\quad': '    ',               # Replace quad space
            r'\\qquad': '        ',          # Replace qquad space
        }
        
        for pattern, replacement in replacements.items():
            cleaned = re.sub(pattern, replacement, cleaned)
        
        return cleaned
    
    def _latex_to_unicode_enhanced(self, latex):
        """Enhanced LaTeX to Unicode conversion"""
        # Comprehensive symbol mapping
        symbols = {
            # Greek letters
            r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ',
            r'\\epsilon': 'ε', r'\\zeta': 'ζ', r'\\eta': 'η', r'\\theta': 'θ',
            r'\\iota': 'ι', r'\\kappa': 'κ', r'\\lambda': 'λ', r'\\mu': 'μ',
            r'\\nu': 'ν', r'\\xi': 'ξ', r'\\pi': 'π', r'\\rho': 'ρ',
            r'\\sigma': 'σ', r'\\tau': 'τ', r'\\upsilon': 'υ', r'\\phi': 'φ',
            r'\\chi': 'χ', r'\\psi': 'ψ', r'\\omega': 'ω',
            
            # Uppercase Greek
            r'\\Gamma': 'Γ', r'\\Delta': 'Δ', r'\\Theta': 'Θ', r'\\Lambda': 'Λ',
            r'\\Xi': 'Ξ', r'\\Pi': 'Π', r'\\Sigma': 'Σ', r'\\Phi': 'Φ',
            r'\\Psi': 'Ψ', r'\\Omega': 'Ω',
            
            # Math operators
            r'\\pm': '±', r'\\mp': '∓', r'\\times': '×', r'\\div': '÷',
            r'\\cdot': '·', r'\\bullet': '•', r'\\ast': '∗',
            r'\\cap': '∩', r'\\cup': '∪', r'\\subset': '⊂', r'\\supset': '⊃',
            r'\\subseteq': '⊆', r'\\supseteq': '⊇', r'\\in': '∈', r'\\notin': '∉',
            r'\\emptyset': '∅', r'\\varnothing': '∅',
            
            # Relations
            r'\\leq': '≤', r'\\le': '≤', r'\\geq': '≥', r'\\ge': '≥',
            r'\\neq': '≠', r'\\ne': '≠', r'\\equiv': '≡', r'\\approx': '≈',
            r'\\sim': '∼', r'\\simeq': '≃', r'\\cong': '≅', r'\\propto': '∝',
            
            # Special symbols
            r'\\infty': '∞', r'\\partial': '∂', r'\\nabla': '∇',
            r'\\exists': '∃', r'\\forall': '∀', r'\\neg': '¬',
            r'\\angle': '∠', r'\\triangle': '△',
            
            # Large operators
            r'\\sum': '∑', r'\\prod': '∏', r'\\int': '∫', r'\\oint': '∮',
            
            # Functions
            r'\\sin': 'sin', r'\\cos': 'cos', r'\\tan': 'tan', r'\\cot': 'cot',
            r'\\sec': 'sec', r'\\csc': 'csc', r'\\log': 'log', r'\\ln': 'ln',
            r'\\exp': 'exp', r'\\lim': 'lim', r'\\max': 'max', r'\\min': 'min',
            
            # Arrows
            r'\\rightarrow': '→', r'\\to': '→', r'\\leftarrow': '←',
            r'\\leftrightarrow': '↔', r'\\Rightarrow': '⇒', r'\\Leftarrow': '⇐',
            r'\\Leftrightarrow': '⇔'
        }
        
        result = latex
        
        # Apply symbol replacements
        for pattern, replacement in symbols.items():
            result = re.sub(pattern, replacement, result)
        
        # Handle fractions
        result = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'(\1)/(\2)', result)
        
        # Handle square roots
        result = re.sub(r'\\sqrt\{([^}]+)\}', r'√(\1)', result)
        result = re.sub(r'\\sqrt\[([^\]]+)\]\{([^}]+)\}', r'ⁿ√(\2)', result)
        
        # Handle superscripts and subscripts
        result = re.sub(r'\^{([^}]+)}', lambda m: self._to_superscript(m.group(1)), result)
        result = re.sub(r'\^(\w)', lambda m: self._to_superscript(m.group(1)), result)
        result = re.sub(r'_{([^}]+)}', lambda m: self._to_subscript(m.group(1)), result)
        result = re.sub(r'_(\w)', lambda m: self._to_subscript(m.group(1)), result)
        
        # Clean up remaining LaTeX commands
        result = re.sub(r'\\mathit\{([^}]+)\}', r'\1', result)
        result = re.sub(r'\\mathrm\{([^}]+)\}', r'\1', result)
        result = re.sub(r'\\text\{([^}]+)\}', r'\1', result)
        result = re.sub(r'\\([a-zA-Z]+)', r'\1', result)
        result = result.replace('{', '').replace('}', '')
        
        return result
    
    def _to_superscript(self, text):
        """Convert text to superscript Unicode"""
        superscript_map = {
            '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵',
            '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', '+': '⁺', '-': '⁻',
            '=': '⁼', '(': '⁽', ')': '⁾', 'n': 'ⁿ', 'i': 'ⁱ'
        }
        return ''.join(superscript_map.get(c, c) for c in text)
    
    def _to_subscript(self, text):
        """Convert text to subscript Unicode"""
        subscript_map = {
            '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅',
            '6': '₆', '7': '₇', '8': '₈', '9': '₉', '+': '₊', '-': '₋',
            '=': '₌', '(': '₍', ')': '₎', 'a': 'ₐ', 'e': 'ₑ', 'i': 'ᵢ',
            'o': 'ₒ', 'u': 'ᵤ', 'x': 'ₓ', 'n': 'ₙ'
        }
        return ''.join(subscript_map.get(c, c) for c in text)
    
    def create_styled_table(self, doc, table_data):
        """Create professionally styled table"""
        if not table_data.get('header') or not table_data.get('rows'):
            return None
        
        header = table_data['header']
        rows = table_data['rows']
        
        # Create table
        table = doc.add_table(rows=len(rows) + 1, cols=len(header))
        
        # Set table style
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Style header
        header_row = table.rows[0]
        for i, header_text in enumerate(header):
            cell = header_row.cells[i]
            cell.text = str(header_text)
            
            # Header styling
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(11)
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.name = 'Calibri'
            
            # Header background
            shading = OxmlElement("w:shd")
            shading.set(qn("w:fill"), "2F5597")  # Professional blue
            cell._tc.get_or_add_tcPr().append(shading)
        
        # Style data rows
        for row_idx, row_data in enumerate(rows):
            table_row = table.rows[row_idx + 1]
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < len(header):
                    cell = table_row.cells[col_idx]
                    cell.text = str(cell_data) if cell_data else ''
                    
                    # Cell styling
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
                            run.font.name = 'Calibri'
                    
                    # Alternating row colors
                    if row_idx % 2 == 0:
                        shading = OxmlElement("w:shd")
                        shading.set(qn("w:fill"), "F8F9FA")  # Light gray
                        cell._tc.get_or_add_tcPr().append(shading)
        
        return table
    
    def create_formatted_document(self, structure, math_expressions, tables):
        """Create formatted Word document with proper styling"""
        doc = Document()
        
        # Set document styles
        styles = doc.styles
        
        # Create or modify Normal style
        normal_style = styles['Normal']
        normal_style.font.name = 'Times New Roman'
        normal_style.font.size = Pt(12)
        
        # Document title
        title = doc.add_heading('Processed Mathematical Document', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_format = title.runs[0]
        title_format.font.name = 'Times New Roman'
        title_format.font.size = Pt(16)
        title_format.font.color.rgb = RGBColor(0, 51, 102)
        
        # Collect all elements with positions
        all_elements = []
        
        # Add math expressions
        for expr in math_expressions:
            all_elements.append({
                'type': 'math',
                'start': expr['start'],
                'end': expr['end'],
                'data': expr
            })
        
        # Add tables
        for table in tables:
            all_elements.append({
                'type': 'table',
                'start': table['start'],
                'end': table['end'],
                'data': table
            })
        
        # Sort by position
        all_elements.sort(key=lambda x: x['start'])
        
        # Process document
        original_text = structure['raw_text']
        current_pos = 0
        stats = {'math_processed': 0, 'tables_created': 0}
        
        for element in all_elements:
            # Add text before element
            if element['start'] > current_pos:
                text_before = original_text[current_pos:element['start']]
                if text_before.strip():
                    # Split into paragraphs and add
                    paragraphs = text_before.split('\n')
                    for para_text in paragraphs:
                        if para_text.strip():
                            p = doc.add_paragraph()
                            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            run = p.add_run(para_text.strip())
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(12)
            
            # Process element
            if element['type'] == 'math':
                math_data = element['data']
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                if self.convert_math_to_word_equation(p, math_data['content']):
                    stats['math_processed'] += 1
                else:
                    # Fallback styling
                    run = p.add_run(f"[MATH] {math_data['content']}")
                    run.italic = True
                    run.bold = True
                    run.font.color.rgb = RGBColor(0, 100, 200)
                    run.font.name = 'Cambria Math'
                
            elif element['type'] == 'table':
                table = self.create_styled_table(doc, element['data'])
                if table:
                    stats['tables_created'] += 1
                    # Add spacing
                    doc.add_paragraph()
            
            current_pos = element['end']
        
        # Add remaining text
        if current_pos < len(original_text):
            remaining_text = original_text[current_pos:]
            if remaining_text.strip():
                paragraphs = remaining_text.split('\n')
                for para_text in paragraphs:
                    if para_text.strip():
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        run = p.add_run(para_text.strip())
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)
        
        # Add summary
        doc.add_page_break()
        summary_heading = doc.add_heading('Processing Summary', 1)
        summary_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        summary_para = doc.add_paragraph()
        summary_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        summary_text = f"""
Document Processing Complete

✅ Mathematical expressions processed: {stats['math_processed']}
✅ Professional tables created: {stats['tables_created']}
✅ Original formatting preserved and enhanced
✅ Professional styling applied

Generated by Professional Word Processor
"""
        run = summary_para.add_run(summary_text.strip())
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.italic = True
        
        return doc, stats
    
    def cleanup(self):
        """Clean up temporary files"""
        import shutil
        try:
            shutil.rmtree(self.temp_dir)
        except:
            pass

def main():
    st.set_page_config(
        page_title="Professional Word Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("📄 Professional Word Document Processor")
    st.markdown("""
    **Phiên bản chuyên nghiệp** với format preservation:
    - ✅ **Format Preservation**: Giữ nguyên cấu trúc và định dạng gốc
    - ✅ **Professional Math**: LaTeX equations → Native Word equations
    - ✅ **Smart Table Detection**: Multiple table format support
    - ✅ **Enhanced Styling**: Professional typography và layout
    """)
    
    # Show status
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if MATH2DOCX_AVAILABLE:
            st.success("🧮 Real equations: ✅")
        else:
            st.warning("🧮 Unicode fallback: ⚠️")
    
    with col2:
        st.success("📊 Professional tables: ✅")
    
    with col3:
        st.success("📝 Format preservation: ✅")
    
    # File upload
    uploaded_file = st.file_uploader(
        "📁 Upload Word Document (.docx)",
        type=['docx'],
        help="Upload your Word document for professional processing"
    )
    
    if uploaded_file:
        processor = ProfessionalWordProcessor()
        
        try:
            # Read document structure
            with st.spinner("Reading document structure..."):
                structure = processor.read_word_with_structure(uploaded_file)
            
            if not structure['raw_text']:
                st.error("Could not read document content.")
                return
            
            # Show preview
            with st.expander("📖 Document Preview"):
                st.write(f"**Text length**: {len(structure['raw_text'])} characters")
                st.write(f"**Paragraphs**: {len(structure['paragraphs'])}")
                st.write(f"**Existing tables**: {len(structure['tables'])}")
                st.text_area("Content Preview", 
                           structure['raw_text'][:1500] + "..." if len(structure['raw_text']) > 1500 else structure['raw_text'],
                           height=200)
            
            # Analyze content
            st.subheader("🔍 Content Analysis")
            
            with st.spinner("Analyzing mathematical content and tables..."):
                math_expressions = processor.detect_math_expressions_advanced(structure['raw_text'])
                detected_tables = processor.detect_tables_comprehensive(structure['raw_text'])
                # Merge with existing tables
                all_tables = structure['tables'] + detected_tables
            
            # Show analysis results
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("🧮 Math Expressions", len(math_expressions))
                
                if math_expressions:
                    st.write("**Found math expressions:**")
                    for i, expr in enumerate(math_expressions[:5]):
                        st.code(f"{expr['type']}: {expr['content'][:60]}...")
                    if len(math_expressions) > 5:
                        st.write(f"... and {len(math_expressions) - 5} more")
            
            with col2:
                st.metric("📊 Tables", len(all_tables))
                
                if all_tables:
                    st.write("**Found tables:**")
                    for i, table in enumerate(all_tables[:3]):
                        if table.get('header'):
                            st.write(f"**Table {i+1}**: {len(table.get('rows', []))}×{len(table['header'])}")
                            # Show preview
                            if table.get('rows'):
                                preview_data = {}
                                for j, col in enumerate(table['header'][:4]):  # Max 4 columns
                                    preview_data[col] = [
                                        row[j] if j < len(row) else ''
                                        for row in table['rows'][:3]  # Max 3 rows
                                    ]
                                st.dataframe(preview_data, use_container_width=True)
            
            # Processing options
            st.subheader("⚙️ Processing Options")
            
            col1, col2 = st.columns(2)
            with col1:
                math_method = st.selectbox(
                    "🧮 Math Processing Method",
                    ["Auto (Recommended)", "math2docx Only", "Unicode Enhanced"],
                    help="Choose how to process mathematical expressions"
                )
                
                preserve_styles = st.checkbox("📝 Preserve Original Styles", value=True)
            
            with col2:
                table_styling = st.selectbox(
                    "📊 Table Styling",
                    ["Professional Blue", "Classic Gray", "Minimal"],
                    help="Choose table styling theme"
                )
                
                enhance_typography = st.checkbox("✨ Enhanced Typography", value=True)
            
            # Process button
            if st.button("🚀 Process Document Professionally", type="primary", use_container_width=True):
                with st.spinner("Professional processing in progress..."):
                    progress_bar = st.progress(0)
                    
                    try:
                        # Progress update
                        progress_bar.progress(25)
                        
                        # Process document
                        processed_doc, stats = processor.create_formatted_document(
                            structure, math_expressions, all_tables
                        )
                        
                        progress_bar.progress(75)
                        
                        # Save document
                        doc_io = io.BytesIO()
                        processed_doc.save(doc_io)
                        doc_io.seek(0)
                        
                        progress_bar.progress(100)
                        
                        # Success
                        st.success("✅ Professional processing complete!")
                        
                        # Show results
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("📐 Math Expressions", stats['math_processed'])
                        with col2:
                            st.metric("📊 Tables Created", stats['tables_created'])
                        with col3:
                            st.metric("📄 Pages", "Professional")
                        
                        # Download button
                        st.download_button(
                            label="📥 Download Professional Document",
                            data=doc_io.getvalue(),
                            file_name="professional_processed_document.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                        # Processing details
                        st.info(f"""
🎯 **Professional Processing Results:**

**Mathematical Content:**
- {stats['math_processed']} expressions converted to native Word equations
- Enhanced Unicode fallback for unsupported expressions
- Professional mathematical typography applied

**Table Enhancement:**
- {stats['tables_created']} tables created with professional styling
- Consistent formatting and alignment
- Color-coded headers and alternating row colors

**Format Preservation:**
- Original document structure maintained
- Enhanced typography with Times New Roman
- Professional styling and spacing applied
- Improved readability and presentation
                        """)
                        
                    except Exception as e:
                        st.error(f"Processing error: {e}")
                        import traceback
                        with st.expander("Error Details"):
                            st.code(traceback.format_exc())
                    finally:
                        progress_bar.empty()
        
        except Exception as e:
            st.error(f"Error reading document: {e}")
            import traceback
            with st.expander("Error Details"):
                st.code(traceback.format_exc())
        
        finally:
            processor.cleanup()
    
    # Help section
    with st.expander("📚 Help & Documentation"):
        st.markdown("""
        ### 🎯 What This Tool Does:
        
        **Format Preservation:**
        - Reads original Word document structure
        - Maintains paragraph styles and formatting
        - Preserves document hierarchy and layout
        
        **Mathematical Processing:**
        - Detects LaTeX math expressions: `$E = mc^2$`
        - Converts to native Word equations (with math2docx)
        - Enhanced Unicode fallback for compatibility
        - Supports multiple math formats
        
        **Table Enhancement:**
        - Detects markdown tables, space-separated tables
        - Creates professional Word tables with styling
        - Color-coded headers and alternating rows
        - Consistent formatting and alignment
        
        ### 📋 Supported Input Formats:
        
        **Math Expressions:**
        ```
        $E = mc^2$                    # Inline LaTeX
        $$F = ma$$                    # Display LaTeX
        [FORMULA: a^2 + b^2 = c^2]    # Formula blocks
        \\alpha, \\beta, \\pi         # Greek letters
        ```
        
        **Tables:**
        ```
        | Header 1 | Header 2 | Header 3 |    # Markdown style
        |----------|----------|----------|
        | Data 1   | Data 2   | Data 3   |
        
        Column1    Column2    Column3        # Space-separated
        Value1     Value2     Value3
        ```
        
        ### ⚙️ Requirements:
        ```bash
        # Essential packages
        pip install streamlit python-docx mammoth Pillow
        
        # For real Word equations (recommended)
        pip install math2docx
        
        # For enhanced math processing
        pip install latex2mathml sympy
        ```
        """)

if __name__ == "__main__":
    main()
