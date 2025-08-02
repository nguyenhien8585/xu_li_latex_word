import streamlit as st
import re
import io
import os
import tempfile
from pathlib import Path
import base64

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

# Try to import math2docx for real equations
try:
    import math2docx
    MATH2DOCX_AVAILABLE = True
    st.success("✅ math2docx available - Real equations enabled!")
except ImportError:
    MATH2DOCX_AVAILABLE = False
    st.warning("⚠️ math2docx not available - Using Unicode fallback")

class EnhancedWordProcessor:
    """Enhanced Word processor with real math equations and improved tables"""
    
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        
    def read_word_document(self, uploaded_file):
        """Đọc nội dung từ file Word"""
        try:
            uploaded_file.seek(0)
            result = mammoth.extract_raw_text(uploaded_file)
            return result.value
        except Exception as e:
            st.error(f"Error reading Word document: {e}")
            try:
                uploaded_file.seek(0)
                doc = Document(uploaded_file)
                content = []
                for paragraph in doc.paragraphs:
                    content.append(paragraph.text)
                return '\n'.join(content)
            except Exception as e2:
                st.error(f"Fallback method also failed: {e2}")
                return ""
    
    def find_formula_blocks(self, text):
        """Tìm và trích xuất các block [FORMULA: ...] từ file test"""
        # Pattern cho [FORMULA: ...]
        formula_pattern = r'\[FORMULA:\s*([^\]]+)\]'
        formulas = []
        
        for match in re.finditer(formula_pattern, text):
            formula_content = match.group(1).strip()
            formulas.append({
                'original': match.group(0),
                'latex': formula_content,
                'start': match.start(),
                'end': match.end()
            })
        
        # Cũng tìm format LaTeX thông thường
        latex_patterns = [
            r'\$\{([^}]+)\}\$',  # ${...}$
            r'\$([^$]+)\$'       # $...$
        ]
        
        for pattern in latex_patterns:
            for match in re.finditer(pattern, text):
                formulas.append({
                    'original': match.group(0),
                    'latex': match.group(1).strip(),
                    'start': match.start(),
                    'end': match.end()
                })
        
        formulas.sort(key=lambda x: x['start'], reverse=True)
        return formulas
    
    def clean_latex_formula(self, latex_string):
        """Làm sạch công thức LaTeX cho math2docx"""
        # Remove extra spaces and newlines
        cleaned = re.sub(r'\s+', ' ', latex_string.strip())
        
        # Fix common LaTeX issues
        replacements = {
            # Fix spacing around operators
            r'\s*=\s*': '=',
            r'\s*\+\s*': '+',
            r'\s*-\s*': '-',
            r'\s*\*\s*': '*',
            
            # Fix brackets
            r'\\left\s*\(': r'\\left(',
            r'\\right\s*\)': r'\\right)',
            r'\\left\s*\[': r'\\left[',
            r'\\right\s*\]': r'\\right]',
            
            # Common symbols
            r'\\cdot': r'\\cdot ',
            r'\\times': r'\\times ',
            r'\\div': r'\\div ',
            
            # Fix subscripts and superscripts spacing
            r'([a-zA-Z0-9])\s*\^': r'\1^',
            r'([a-zA-Z0-9])\s*_': r'\1_',
        }
        
        for pattern, replacement in replacements.items():
            cleaned = re.sub(pattern, replacement, cleaned)
        
        return cleaned
    
    def add_real_math_equation(self, paragraph, latex_string):
        """Thêm equation thật sử dụng math2docx"""
        try:
            if MATH2DOCX_AVAILABLE:
                # Clean the LaTeX first
                clean_latex = self.clean_latex_formula(latex_string)
                
                # Add equation using math2docx
                math2docx.add_math(paragraph, clean_latex)
                return True
            else:
                return False
        except Exception as e:
            st.warning(f"Error adding math equation '{latex_string[:50]}...': {e}")
            return False
    
    def latex_to_unicode_enhanced(self, latex_string):
        """Enhanced Unicode conversion with more symbols"""
        # All Unicode replacements from previous version plus more
        replacements = {
            # Greek letters (lowercase)
            r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ',
            r'\\epsilon': 'ε', r'\\varepsilon': 'ε', r'\\zeta': 'ζ', r'\\eta': 'η', 
            r'\\theta': 'θ', r'\\vartheta': 'ϑ', r'\\iota': 'ι', r'\\kappa': 'κ',
            r'\\lambda': 'λ', r'\\mu': 'μ', r'\\nu': 'ν', r'\\xi': 'ξ',
            r'\\pi': 'π', r'\\varpi': 'ϖ', r'\\rho': 'ρ', r'\\varrho': 'ϱ',
            r'\\sigma': 'σ', r'\\varsigma': 'ς', r'\\tau': 'τ', r'\\upsilon': 'υ',
            r'\\phi': 'φ', r'\\varphi': 'φ', r'\\chi': 'χ', r'\\psi': 'ψ', r'\\omega': 'ω',
            
            # Greek letters (uppercase)
            r'\\Gamma': 'Γ', r'\\Delta': 'Δ', r'\\Theta': 'Θ', r'\\Lambda': 'Λ',
            r'\\Xi': 'Ξ', r'\\Pi': 'Π', r'\\Sigma': 'Σ', r'\\Upsilon': 'Υ',
            r'\\Phi': 'Φ', r'\\Chi': 'Χ', r'\\Psi': 'Ψ', r'\\Omega': 'Ω',
            
            # Mathematical operators
            r'\\pm': '±', r'\\mp': '∓', r'\\times': '×', r'\\div': '÷',
            r'\\cdot': '·', r'\\bullet': '•', r'\\ast': '∗', r'\\star': '⋆',
            r'\\circ': '∘', r'\\cap': '∩', r'\\cup': '∪',
            r'\\sqcap': '⊓', r'\\sqcup': '⊔', r'\\vee': '∨', r'\\wedge': '∧',
            
            # Relations
            r'\\leq': '≤', r'\\le': '≤', r'\\geq': '≥', r'\\ge': '≥',
            r'\\neq': '≠', r'\\ne': '≠', r'\\equiv': '≡', r'\\approx': '≈',
            r'\\sim': '∼', r'\\simeq': '≃', r'\\cong': '≅', r'\\propto': '∝',
            r'\\prec': '≺', r'\\succ': '≻', r'\\preceq': '⪯', r'\\succeq': '⪰',
            r'\\subset': '⊂', r'\\supset': '⊃', r'\\subseteq': '⊆', r'\\supseteq': '⊇',
            r'\\in': '∈', r'\\notin': '∉', r'\\ni': '∋', r'\\notni': '∌',
            
            # Arrows
            r'\\leftarrow': '←', r'\\gets': '←', r'\\rightarrow': '→', r'\\to': '→',
            r'\\leftrightarrow': '↔', r'\\Leftarrow': '⇐', r'\\Rightarrow': '⇒',
            r'\\Leftrightarrow': '⇔', r'\\mapsto': '↦', r'\\longmapsto': '⟼',
            r'\\uparrow': '↑', r'\\downarrow': '↓', r'\\updownarrow': '↕',
            r'\\Uparrow': '⇑', r'\\Downarrow': '⇓', r'\\Updownarrow': '⇕',
            
            # Special symbols
            r'\\infty': '∞', r'\\partial': '∂', r'\\nabla': '∇', r'\\emptyset': '∅',
            r'\\varnothing': '∅', r'\\exists': '∃', r'\\nexists': '∄', r'\\forall': '∀',
            r'\\neg': '¬', r'\\lnot': '¬', r'\\therefore': '∴', r'\\because': '∵',
            r'\\angle': '∠', r'\\measuredangle': '∡', r'\\sphericalangle': '∢',
            r'\\triangle': '△', r'\\square': '□', r'\\blacksquare': '■',
            r'\\diamond': '◊', r'\\blacklozenge': '⧫', r'\\clubsuit': '♣',
            r'\\diamondsuit': '♢', r'\\heartsuit': '♡', r'\\spadesuit': '♠',
            
            # Large operators
            r'\\sum': '∑', r'\\prod': '∏', r'\\coprod': '∐', r'\\int': '∫',
            r'\\oint': '∮', r'\\iint': '∬', r'\\iiint': '∭', r'\\iiiint': '⨌',
            r'\\bigcap': '⋂', r'\\bigcup': '⋃', r'\\bigwedge': '⋀', r'\\bigvee': '⋁',
            r'\\bigoplus': '⊕', r'\\bigotimes': '⊗', r'\\bigodot': '⊙',
            
            # Functions
            r'\\sin': 'sin', r'\\cos': 'cos', r'\\tan': 'tan', r'\\cot': 'cot',
            r'\\sec': 'sec', r'\\csc': 'csc', r'\\sinh': 'sinh', r'\\cosh': 'cosh',
            r'\\tanh': 'tanh', r'\\coth': 'coth', r'\\arcsin': 'arcsin',
            r'\\arccos': 'arccos', r'\\arctan': 'arctan', r'\\arccot': 'arccot',
            r'\\log': 'log', r'\\ln': 'ln', r'\\lg': 'lg', r'\\exp': 'exp',
            r'\\max': 'max', r'\\min': 'min', r'\\sup': 'sup', r'\\inf': 'inf',
            r'\\lim': 'lim', r'\\limsup': 'lim sup', r'\\liminf': 'lim inf',
            r'\\gcd': 'gcd', r'\\lcm': 'lcm', r'\\det': 'det', r'\\dim': 'dim',
            r'\\ker': 'ker', r'\\hom': 'hom', r'\\arg': 'arg', r'\\deg': 'deg',
            
            # Miscellaneous
            r'\\hbar': 'ℏ', r'\\ell': 'ℓ', r'\\wp': '℘', r'\\Re': 'ℜ', r'\\Im': 'ℑ',
            r'\\aleph': 'ℵ', r'\\beth': 'ℶ', r'\\gimel': 'ℷ', r'\\daleth': 'ℸ',
        }
        
        result = latex_string
        for pattern, replacement in replacements.items():
            result = re.sub(pattern, replacement, result)
        
        # Handle fractions
        result = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'(\1)/(\2)', result)
        
        # Handle square roots
        result = re.sub(r'\\sqrt\{([^}]+)\}', r'√(\1)', result)
        result = re.sub(r'\\sqrt\[([^]]+)\]\{([^}]+)\}', r'ⁿ√(\2)', result)
        
        # Handle superscripts and subscripts with better formatting
        result = re.sub(r'\^{([^}]+)}', r'ᵖᵒʷ(\1)', result)  # Temporary placeholder
        result = re.sub(r'\^(\w)', r'ᵖᵒʷ\1', result)
        result = re.sub(r'_{([^}]+)}', r'ₛᵤᵦ(\1)', result)  # Temporary placeholder
        result = re.sub(r'_(\w)', r'ₛᵤᵦ\1', result)
        
        # Convert to actual superscript/subscript Unicode
        superscript_map = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', 
                          '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', '+': '⁺', '-': '⁻', 
                          '=': '⁼', '(': '⁽', ')': '⁾', 'n': 'ⁿ', 'i': 'ⁱ'}
        subscript_map = {'0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅',
                        '6': '₆', '7': '₇', '8': '₈', '9': '₉', '+': '₊', '-': '₋',
                        '=': '₌', '(': '₍', ')': '₎', 'a': 'ₐ', 'e': 'ₑ', 'i': 'ᵢ',
                        'o': 'ₒ', 'u': 'ᵤ', 'x': 'ₓ', 'n': 'ₙ'}
        
        # Replace superscripts
        def replace_superscript(match):
            content = match.group(1) if '(' in match.group(0) else match.group(0).replace('ᵖᵒʷ', '')
            return ''.join(superscript_map.get(c, c) for c in content)
        
        result = re.sub(r'ᵖᵒʷ\(([^)]+)\)', replace_superscript, result)
        result = re.sub(r'ᵖᵒʷ(\w)', replace_superscript, result)
        
        # Replace subscripts
        def replace_subscript(match):
            content = match.group(1) if '(' in match.group(0) else match.group(0).replace('ₛᵤᵦ', '')
            return ''.join(subscript_map.get(c, c) for c in content)
        
        result = re.sub(r'ₛᵤᵦ\(([^)]+)\)', replace_subscript, result)
        result = re.sub(r'ₛᵤᵦ(\w)', replace_subscript, result)
        
        # Handle brackets
        result = result.replace(r'\left(', '(').replace(r'\right)', ')')
        result = result.replace(r'\left[', '[').replace(r'\right]', ']')
        result = result.replace(r'\left\{', '{').replace(r'\right\}', '}')
        result = result.replace(r'\{', '{').replace(r'\}', '}')
        
        # Clean up remaining LaTeX commands
        result = re.sub(r'\\([a-zA-Z]+)', r'\1', result)
        
        return result
    
    def find_markdown_tables_enhanced(self, text):
        """Enhanced markdown table detection with better parsing"""
        # More robust table pattern
        table_pattern = r'(\|[^\n\r]+\|(?:\r?\n|\r))+\|[-:\s|]+\|(?:\r?\n|\r)(\|[^\n\r]+\|(?:\r?\n|\r)?)*'
        tables = []
        
        for match in re.finditer(table_pattern, text, re.MULTILINE):
            table_text = match.group(0).strip()
            lines = [line.strip() for line in table_text.split('\n') if line.strip()]
            
            if len(lines) >= 2:
                # Parse header
                header_line = lines[0]
                header = [cell.strip() for cell in header_line.split('|')[1:-1] if cell.strip()]
                
                # Skip separator line (lines[1])
                # Parse data rows
                rows = []
                for line in lines[2:]:
                    if line.strip() and '|' in line:
                        row_cells = [cell.strip() for cell in line.split('|')[1:-1]]
                        # Ensure row has same number of columns as header
                        while len(row_cells) < len(header):
                            row_cells.append('')
                        rows.append(row_cells[:len(header)])  # Truncate if too many columns
                
                if header and rows:  # Only add if we have both header and data
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
    
    def create_enhanced_table(self, doc, table_data):
        """Tạo bảng Word với styling đẹp"""
        if not table_data['header'] or not table_data['rows']:
            return None
            
        # Create table
        table = doc.add_table(rows=1+len(table_data['rows']), cols=len(table_data['header']))
        
        # Set table style
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Style header row
        header_row = table.rows[0]
        for i, header_text in enumerate(table_data['header']):
            cell = header_row.cells[i]
            cell.text = header_text
            
            # Header formatting
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(11)
                    run.font.color.rgb = RGBColor(255, 255, 255)  # White text
            
            # Cell background color (dark blue)
            shading_elm = OxmlElement("w:shd")
            shading_elm.set(qn("w:fill"), "4472C4")  # Blue background
            cell._tc.get_or_add_tcPr().append(shading_elm)
        
        # Add data rows
        for row_idx, row_data in enumerate(table_data['rows']):
            table_row = table.rows[row_idx + 1]
            for col_idx, cell_text in enumerate(row_data):
                if col_idx < len(table_data['header']):
                    cell = table_row.cells[col_idx]
                    cell.text = cell_text
                    
                    # Data cell formatting
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
                    
                    # Alternating row colors
                    if row_idx % 2 == 0:
                        shading_elm = OxmlElement("w:shd")
                        shading_elm.set(qn("w:fill"), "F2F2F2")  # Light gray
                        cell._tc.get_or_add_tcPr().append(shading_elm)
        
        return table
    
    def create_tikz_placeholder_enhanced(self, tikz_code):
        """Tạo ảnh placeholder TikZ với styling đẹp hơn"""
        try:
            width, height = 500, 400
            img = Image.new('RGB', (width, height), color='white')
            draw = ImageDraw.Draw(img)
            
            # Draw gradient-like background
            for i in range(height):
                color_val = int(255 - (i / height) * 20)  # Slight gradient
                draw.line([(0, i), (width, i)], fill=(color_val, color_val, color_val))
            
            # Draw border with rounded effect
            border_color = (70, 130, 180)  # Steel blue
            draw.rectangle([5, 5, width-5, height-5], outline=border_color, width=3)
            draw.rectangle([8, 8, width-8, height-8], outline=border_color, width=1)
            
            # Add title
            try:
                font_large = ImageFont.load_default()
                font_small = ImageFont.load_default()
            except:
                font_large = font_small = None
            
            # Title
            title = "📊 TikZ Diagram"
            title_bbox = draw.textbbox((0, 0), title, font=font_large)
            title_width = title_bbox[2] - title_bbox[0]
            draw.text(((width - title_width) // 2, 20), title, fill=(70, 130, 180), font=font_large)
            
            # Code preview
            draw.text((20, 60), "LaTeX Code:", fill=(100, 100, 100), font=font_small)
            
            # Format code nicely
            code_lines = tikz_code.replace('\\', '\n\\').split('\n')
            y_offset = 90
            for i, line in enumerate(code_lines[:12]):  # Show first 12 lines
                if line.strip():
                    display_line = line[:70] + "..." if len(line) > 70 else line
                    draw.text((25, y_offset), display_line, fill=(60, 60, 60), font=font_small)
                    y_offset += 20
                    
                if y_offset > height - 40:
                    break
            
            # Add footer
            if len(code_lines) > 12:
                draw.text((25, height - 30), f"... và {len(code_lines) - 12} dòng nữa", 
                         fill=(150, 150, 150), font=font_small)
            
            # Save image
            img_path = os.path.join(self.temp_dir, f'tikz_enhanced_{hash(tikz_code)}.png')
            img.save(img_path, 'PNG', quality=95)
            return img_path
            
        except Exception as e:
            st.warning(f"Could not create enhanced TikZ placeholder: {e}")
            return None
    
    def create_processed_document_enhanced(self, original_text, formulas, tables, tikz_blocks):
        """Tạo document Word với equations và tables thật"""
        doc = Document()
        
        # Add title
        title = doc.add_heading('Processed Document', 0)
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
        equations_added = 0
        tables_added = 0
        
        for element in all_elements:
            # Add text before element
            if element['start'] > current_pos:
                text_before = original_text[current_pos:element['start']]
                if text_before.strip():
                    p = doc.add_paragraph(text_before.strip())
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            # Process element
            if element['type'] == 'formula':
                formula_data = element['data']
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Try to add real equation first
                equation_added = self.add_real_math_equation(p, formula_data['latex'])
                
                if equation_added:
                    equations_added += 1
                else:
                    # Fallback to enhanced Unicode
                    unicode_formula = self.latex_to_unicode_enhanced(formula_data['latex'])
                    run = p.add_run(f"[MATH] {unicode_formula}")
                    run.italic = True
                    run.bold = True
                    run.font.size = Pt(12)
                    run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
                
            elif element['type'] == 'table':
                # Add enhanced table
                table = self.create_enhanced_table(doc, element['data'])
                if table:
                    tables_added += 1
                    # Add some spacing
                    doc.add_paragraph()
                
            elif element['type'] == 'tikz':
                # Add enhanced TikZ placeholder
                img_path = self.create_tikz_placeholder_enhanced(element['data']['code'])
                if img_path and os.path.exists(img_path):
                    try:
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p.add_run()
                        run.add_picture(img_path, width=Inches(5))
                    except Exception as e:
                        # Fallback to styled text
                        p = doc.add_paragraph()
                        run = p.add_run(f"[TikZ DIAGRAM]\n{element['data']['code']}")
                        run.italic = True
                        run.font.color.rgb = RGBColor(139, 69, 19)  # Brown color
                else:
                    p = doc.add_paragraph()
                    run = p.add_run(f"[TikZ DIAGRAM]\n{element['data']['code']}")
                    run.italic = True
                    run.font.color.rgb = RGBColor(139, 69, 19)
            
            current_pos = element['end']
        
        # Add remaining text
        if current_pos < len(original_text):
            remaining_text = original_text[current_pos:]
            if remaining_text.strip():
                p = doc.add_paragraph(remaining_text.strip())
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Add summary at the end
        doc.add_page_break()
        summary = doc.add_heading('Processing Summary', 1)
        summary.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        summary_text = f"""
        📊 Document Processing Complete
        
        • Total formulas processed: {len(formulas)}
        • Real equations added: {equations_added}
        • Tables created: {tables_added}
        • TikZ diagrams: {len(tikz_blocks)}
        
        Generated by Enhanced Word Processor
        """
        
        p = doc.add_paragraph(summary_text.strip())
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
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
        page_title="Enhanced Word Processor",
        page_icon="🔬",
        layout="wide"
    )
    
    st.title("🔬 Enhanced Word Document Processor")
    st.markdown("""
    **Professional Version** ⚡
    - ✅ **Real Math Equations** (LaTeX → Word Equations)
    - ✅ **Enhanced Tables** (Markdown → Styled Word Tables)
    - ✅ **Formula Detection** ([FORMULA: ...] format support)
    - ✅ **Professional Styling** (Colors, borders, alignment)
    """)
    
    # Show math2docx status
    col1, col2 = st.columns(2)
    with col1:
        if MATH2DOCX_AVAILABLE:
            st.success("🧮 Real equations enabled (math2docx)")
        else:
            st.info("📝 Unicode fallback mode")
            with st.expander("How to enable real equations"):
                st.code("pip install math2docx")
    
    with col2:
        st.info("📊 Enhanced tables & styling enabled")
    
    # Upload file
    uploaded_file = st.file_uploader(
        "📁 Upload Word Document (.docx)", 
        type=['docx'],
        help="Upload your Word document with LaTeX formulas, [FORMULA: ...] blocks, and Markdown tables"
    )
    
    if uploaded_file:
        processor = EnhancedWordProcessor()
        
        try:
            # Read content
            with st.spinner("Reading Word document..."):
                content = processor.read_word_document(uploaded_file)
            
            if not content:
                st.error("Could not read document content. Please check the file.")
                return
            
            # Preview original content
            with st.expander("📖 Original Content Preview"):
                st.text_area("Content", content[:2000] + "..." if len(content) > 2000 else content, height=200)
            
            # Find elements
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.subheader("🧮 Math Formulas")
                formulas = processor.find_formula_blocks(content)
                st.write(f"Found {len(formulas)} formulas")
                
                for i, formula in enumerate(formulas[:3]):
                    st.code(formula['original'])
                    if MATH2DOCX_AVAILABLE:
                        st.write(f"→ Real equation: {formula['latex']}")
                    else:
                        unicode_preview = processor.latex_to_unicode_enhanced(formula['latex'])
                        st.write(f"→ Unicode: {unicode_preview}")
            
            with col2:
                st.subheader("📊 Tables") 
                tables = processor.find_markdown_tables_enhanced(content)
                st.write(f"Found {len(tables)} tables")
                
                for i, table in enumerate(tables[:2]):
                    if table['header'] and table['rows']:
                        st.write(f"**Table {i+1}:** {len(table['rows'])} rows × {len(table['header'])} cols")
                        try:
                            df_data = {}
                            for j, col_name in enumerate(table['header']):
                                df_data[col_name] = [
                                    row[j] if j < len(row) else '' 
                                    for row in table['rows']
                                ]
                            st.dataframe(df_data, use_container_width=True)
                        except Exception as e:
                            st.write(f"Preview: {table['header']}")
            
            with col3:
                st.subheader("🎨 TikZ Diagrams")
                tikz_blocks = processor.find_tikz_code(content)
                st.write(f"Found {len(tikz_blocks)} TikZ blocks")
                
                for i, tikz in enumerate(tikz_blocks[:2]):
                    st.code(tikz['code'][:100] + "..." if len(tikz['code']) > 100 else tikz['code'])
            
            # Process button
            if st.button("🚀 Process Document", type="primary", use_container_width=True):
                with st.spinner("Processing document with enhanced features..."):
                    try:
                        processed_doc = processor.create_processed_document_enhanced(
                            content, formulas, tables, tikz_blocks
                        )
                        
                        # Save to bytes
                        doc_io = io.BytesIO()
                        processed_doc.save(doc_io)
                        doc_io.seek(0)
                        
                        # Success message
                        st.success("✅ Processing complete!")
                        
                        # Download button
                        st.download_button(
                            label="📥 Download Enhanced Word Document",
                            data=doc_io.getvalue(),
                            file_name="enhanced_document.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                        # Show detailed stats
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Math Formulas", len(formulas))
                        with col2:
                            st.metric("Tables Created", len(tables))
                        with col3:
                            st.metric("TikZ Diagrams", len(tikz_blocks))
                        
                        # Processing details
                        st.info(f"""
                        📈 **Processing Summary:**
                        - {len(formulas)} formulas → {'Real equations' if MATH2DOCX_AVAILABLE else 'Unicode symbols'}
                        - {len(tables)} tables → Styled Word tables with colors & borders
                        - {len(tikz_blocks)} TikZ → Enhanced placeholder images
                        - Professional formatting applied
                        """)
                        
                    except Exception as e:
                        st.error(f"Processing error: {e}")
                        st.code(str(e))
        
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
            st.code(str(e))
        
        finally:
            processor.cleanup()
    
    # Instructions section
    with st.expander("📋 Usage Instructions"):
        st.markdown("""
        ### 📝 Supported Input Formats:
        
        **1. Formula Formats:**
        ```
        [FORMULA: A = \pi r^2]        # Your test file format
        $E = mc^2$                    # Standard LaTeX
        ${F = ma}$                    # Alternative LaTeX
        ```
        
        **2. Enhanced Tables:**
        ```
        | Header 1 | Header 2 | Header 3 |
        |----------|----------|----------|
        | Cell 1   | Cell 2   | Cell 3   |
        | Data 1   | Data 2   | Data 3   |
        ```
        
        **3. TikZ Diagrams:**
        ```
        \\begin{tikzpicture}
        \\draw (0,0) circle (1cm);
        \\end{tikzpicture}
        ```
        
        ### 🌟 Enhanced Features:
        - **Real Equations**: LaTeX → Native Word equations (with math2docx)
        - **Styled Tables**: Professional colors, borders, alternating rows
        - **Smart Detection**: Handles [FORMULA: ...] format from your test
        - **Enhanced Unicode**: 200+ mathematical symbols
        - **Professional Layout**: Centered equations, justified text
        - **Processing Summary**: Detailed statistics report
        """)
    
    # Installation instructions
    with st.expander("⚙️ Installation for Real Equations"):
        st.markdown("""
        ### For REAL Math Equations (Recommended):
        
        ```bash
        # Install math2docx for real equations
        pip install math2docx
        
        # Full requirements
        pip install streamlit python-docx mammoth markdown Pillow PyMuPDF math2docx
        ```
        
        ### For Enhanced Unicode Only:
        ```bash
        # Minimal installation (Unicode fallback)
        pip install streamlit python-docx mammoth markdown Pillow PyMuPDF
        ```
        
        **Note**: math2docx enables real LaTeX → Word equation conversion!
        """)

if __name__ == "__main__":
    main()
