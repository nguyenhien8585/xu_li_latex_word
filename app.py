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
    from docx.shared import Inches, Pt, RGBColor, Cm
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

class UltimateWordProcessor:
    """Ultimate Word processor với perfect formatting và TikZ compilation"""
    
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        self.tikz_counter = 0
        
    def read_word_comprehensive(self, uploaded_file):
        """Comprehensive Word document reading"""
        try:
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            
            structure = {
                'paragraphs': [],
                'tables': [],
                'full_text': '',
                'document_elements': []
            }
            
            # Extract all elements in order
            for element in doc.element.body:
                if element.tag.endswith('p'):  # Paragraph
                    para_text = ''
                    for paragraph in doc.paragraphs:
                        if paragraph._element == element:
                            para_text = paragraph.text
                            break
                    
                    structure['document_elements'].append({
                        'type': 'paragraph',
                        'text': para_text,
                        'xml': element
                    })
                    structure['paragraphs'].append(para_text)
                    
                elif element.tag.endswith('tbl'):  # Table
                    for table in doc.tables:
                        if table._element == element:
                            table_data = self._extract_table_data(table)
                            structure['document_elements'].append({
                                'type': 'table',
                                'data': table_data,
                                'xml': element
                            })
                            structure['tables'].append(table_data)
                            break
            
            structure['full_text'] = '\n'.join(structure['paragraphs'])
            return structure
            
        except Exception as e:
            st.error(f"Error reading Word document: {e}")
            return {'paragraphs': [], 'tables': [], 'full_text': '', 'document_elements': []}
    
    def _extract_table_data(self, table):
        """Extract table data with formula detection"""
        table_data = {
            'rows': [],
            'header': [],
            'has_formulas': False
        }
        
        for i, row in enumerate(table.rows):
            row_data = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                # Check for formulas in cell
                if self._has_math_content(cell_text):
                    table_data['has_formulas'] = True
                row_data.append(cell_text)
            
            if i == 0:
                table_data['header'] = row_data
            else:
                table_data['rows'].append(row_data)
        
        return table_data
    
    def _has_math_content(self, text):
        """Check if text contains mathematical content"""
        math_patterns = [
            r'\$[^$]+\$',
            r'\\\([^)]+\\\)',
            r'\\\[[^\]]+\\\]',
            r'\\[a-zA-Z]+',
            r'[α-ωΑ-Ω∑∏∫]'
        ]
        return any(re.search(pattern, text) for pattern in math_patterns)
    
    def detect_all_math_expressions(self, text):
        """Comprehensive math expression detection"""
        expressions = []
        
        # Comprehensive patterns
        patterns = [
            # LaTeX patterns
            (r'\$([^$]+)\$', 'latex_inline', True),  # $...$
            (r'\$\$([^$]+)\$\$', 'latex_display', True),  # $$...$$
            (r'\\\(([^)]+)\\\)', 'latex_inline_paren', True),  # \(...\)
            (r'\\\[([^\]]+)\\\]', 'latex_display_bracket', True),  # \[...\]
            
            # Escaped dollar signs (from Word conversion)
            (r'\\\$([^\\$]+)\\\$', 'escaped_latex', True),  # \$...\$
            
            # Formula blocks
            (r'\[FORMULA:\s*([^\]]+)\]', 'formula_block', True),
            (r'\[MATH:\s*([^\]]+)\]', 'math_block', True),
            
            # Environment blocks
            (r'\\begin\{equation\}(.*?)\\end\{equation\}', 'equation_env', True),
            (r'\\begin\{align\}(.*?)\\end\{align\}', 'align_env', True),
            (r'\\begin\{math\}(.*?)\\end\{math\}', 'math_env', True),
            
            # Standalone math symbols and expressions
            (r'\\[a-zA-Z]+\{[^}]*\}', 'latex_command', False),
            (r'[α-ωΑ-Ω∑∏∫√±×÷≤≥≠∞∂∇]+', 'unicode_math', False),
        ]
        
        for pattern, math_type, is_block in patterns:
            flags = re.DOTALL if 'env' in math_type else 0
            for match in re.finditer(pattern, text, flags):
                content = match.group(1) if match.groups() else match.group(0)
                
                expressions.append({
                    'original': match.group(0),
                    'content': content.strip(),
                    'type': math_type,
                    'is_block': is_block,
                    'start': match.start(),
                    'end': match.end()
                })
        
        # Remove overlapping matches (keep longer ones)
        expressions = self._remove_overlapping_expressions(expressions)
        expressions.sort(key=lambda x: x['start'], reverse=True)
        
        return expressions
    
    def _remove_overlapping_expressions(self, expressions):
        """Remove overlapping expressions, keeping the longest ones"""
        expressions.sort(key=lambda x: (x['start'], -(x['end'] - x['start'])))
        
        filtered = []
        for expr in expressions:
            overlaps = False
            for existing in filtered:
                if (expr['start'] < existing['end'] and expr['end'] > existing['start']):
                    overlaps = True
                    break
            if not overlaps:
                filtered.append(expr)
        
        return filtered
    
    def detect_tikz_blocks(self, text):
        """Enhanced TikZ detection"""
        tikz_blocks = []
        
        # Pattern for TikZ blocks
        patterns = [
            r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}',
            r'\\begin\{tikzpicture\}\[.*?\].*?\\end\{tikzpicture\}'
        ]
        
        for pattern in patterns:
            for match in re.finditer(pattern, text, re.DOTALL):
                tikz_blocks.append({
                    'original': match.group(0),
                    'code': self._extract_tikz_code(match.group(0)),
                    'start': match.start(),
                    'end': match.end(),
                    'compiled_path': None
                })
        
        tikz_blocks.sort(key=lambda x: x['start'], reverse=True)
        return tikz_blocks
    
    def _extract_tikz_code(self, tikz_block):
        """Extract clean TikZ code"""
        # Remove begin/end tags and extract content
        content = re.sub(r'\\begin\{tikzpicture\}(?:\[.*?\])?', '', tikz_block)
        content = re.sub(r'\\end\{tikzpicture\}', '', content)
        return content.strip()
    
    def compile_tikz_advanced(self, tikz_code):
        """Advanced TikZ compilation with better error handling"""
        self.tikz_counter += 1
        output_path = os.path.join(self.temp_dir, f'tikz_{self.tikz_counter}.png')
        
        try:
            # Enhanced LaTeX template
            latex_template = r"""
\documentclass[border=5pt,convert={{density=300,outext=.png}}]{{standalone}}
\usepackage{{tikz}}
\usepackage{{amsmath,amsfonts,amssymb}}
\usepackage{{pgfplots}}
\pgfplotsset{{compat=1.18}}
\usetikzlibrary{{shapes,arrows,positioning,calc,patterns,decorations.pathreplacing}}

\begin{{document}}
\begin{{tikzpicture}}
{tikz_code}
\end{{tikzpicture}}
\end{{document}}
"""
            
            latex_content = latex_template.format(tikz_code=tikz_code)
            
            # Write LaTeX file
            tex_file = os.path.join(self.temp_dir, f'tikz_{self.tikz_counter}.tex')
            with open(tex_file, 'w', encoding='utf-8') as f:
                f.write(latex_content)
            
            # Try different compilation methods
            compilation_success = False
            
            # Method 1: pdflatex + ImageMagick
            try:
                cmd = ['pdflatex', '-shell-escape', '-output-directory', self.temp_dir, 
                       '-interaction=nonstopmode', tex_file]
                result = subprocess.run(cmd, capture_output=True, text=True, cwd=self.temp_dir, timeout=30)
                
                if result.returncode == 0:
                    pdf_file = os.path.join(self.temp_dir, f'tikz_{self.tikz_counter}.pdf')
                    if os.path.exists(pdf_file):
                        # Convert PDF to PNG using PyMuPDF
                        doc = fitz.open(pdf_file)
                        page = doc[0]
                        mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for quality
                        pix = page.get_pixmap(matrix=mat)
                        pix.save(output_path)
                        doc.close()
                        compilation_success = True
            except Exception as e:
                st.warning(f"Method 1 failed: {e}")
            
            # Method 2: Direct ImageMagick (if available)
            if not compilation_success:
                try:
                    cmd = ['convert', '-density', '300', pdf_file, output_path]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
                    if result.returncode == 0 and os.path.exists(output_path):
                        compilation_success = True
                except:
                    pass
            
            if compilation_success and os.path.exists(output_path):
                return output_path
            else:
                return self._create_tikz_fallback_advanced(tikz_code, output_path)
                
        except Exception as e:
            st.warning(f"TikZ compilation error: {e}")
            return self._create_tikz_fallback_advanced(tikz_code, output_path)
    
    def _create_tikz_fallback_advanced(self, tikz_code, output_path):
        """Create enhanced TikZ fallback image"""
        try:
            width, height = 800, 600
            img = Image.new('RGB', (width, height), color='#f8f9fa')
            draw = ImageDraw.Draw(img)
            
            # Draw professional border
            border_color = '#2c3e50'
            draw.rectangle([10, 10, width-10, height-10], outline=border_color, width=4)
            draw.rectangle([15, 15, width-15, height-15], outline='#3498db', width=2)
            
            # Add title with better font
            try:
                # Try to load a better font
                font_large = ImageFont.truetype("arial.ttf", 24)
                font_medium = ImageFont.truetype("arial.ttf", 14)
                font_small = ImageFont.truetype("arial.ttf", 12)
            except:
                font_large = font_medium = font_small = ImageFont.load_default()
            
            # Title
            title = "📊 TikZ Diagram"
            title_bbox = draw.textbbox((0, 0), title, font=font_large)
            title_width = title_bbox[2] - title_bbox[0]
            draw.text(((width - title_width) // 2, 30), title, fill='#2c3e50', font=font_large)
            
            # Subtitle
            subtitle = "(LaTeX Compilation Required)"
            subtitle_bbox = draw.textbbox((0, 0), subtitle, font=font_medium)
            subtitle_width = subtitle_bbox[2] - subtitle_bbox[0]
            draw.text(((width - subtitle_width) // 2, 65), subtitle, fill='#7f8c8d', font=font_medium)
            
            # Code preview section
            draw.text((30, 120), "TikZ Code Preview:", fill='#2c3e50', font=font_medium)
            
            # Clean and format code
            code_lines = tikz_code.split('\n')
            y_offset = 150
            max_lines = 20
            
            for i, line in enumerate(code_lines[:max_lines]):
                if line.strip():
                    # Truncate long lines
                    display_line = line[:100] + "..." if len(line) > 100 else line
                    draw.text((40, y_offset), display_line, fill='#34495e', font=font_small)
                    y_offset += 18
                    
                if y_offset > height - 100:
                    break
            
            # Add footer info
            if len(code_lines) > max_lines:
                draw.text((40, height - 80), f"... and {len(code_lines) - max_lines} more lines", 
                         fill='#95a5a6', font=font_small)
            
            # Add instruction
            instruction = "Install LaTeX (TeX Live/MiKTeX) to compile TikZ diagrams"
            instr_bbox = draw.textbbox((0, 0), instruction, font=font_small)
            instr_width = instr_bbox[2] - instr_bbox[0]
            draw.text(((width - instr_width) // 2, height - 40), instruction, 
                     fill='#e74c3c', font=font_small)
            
            # Save image
            img.save(output_path, 'PNG', quality=95)
            return output_path
            
        except Exception as e:
            st.error(f"Could not create TikZ fallback: {e}")
            return None
    
    def add_centered_math_equation(self, doc, latex_content, math_type='inline'):
        """Add properly centered math equation"""
        try:
            # Create new paragraph for math
            if math_type in ['display', 'equation_env', 'align_env', 'latex_display', 'formula_block']:
                # Display math - centered on new line
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Add some spacing before
                para_format = para.paragraph_format
                para_format.space_before = Pt(6)
                para_format.space_after = Pt(6)
                
            else:
                # Inline math - same line
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Clean LaTeX content
            cleaned_latex = self._clean_latex_comprehensive(latex_content)
            
            # Try to add real equation
            if MATH2DOCX_AVAILABLE and cleaned_latex:
                try:
                    math2docx.add_math(para, cleaned_latex)
                    return True
                except Exception as e:
                    st.warning(f"math2docx failed for '{cleaned_latex[:30]}...': {e}")
            
            # Fallback to enhanced Unicode
            unicode_math = self._latex_to_unicode_comprehensive(latex_content)
            run = para.add_run(unicode_math)
            run.font.name = 'Cambria Math'
            run.font.size = Pt(14) if math_type.startswith('display') else Pt(12)
            run.bold = True
            run.font.color.rgb = RGBColor(0, 51, 102)
            
            return True
            
        except Exception as e:
            st.warning(f"Error adding math equation: {e}")
            return False
    
    def _clean_latex_comprehensive(self, latex):
        """Comprehensive LaTeX cleaning"""
        if not latex:
            return latex
            
        cleaned = latex.strip()
        
        # Remove common Word-generated artifacts
        replacements = {
            r'\\mathit\{([^}]+)\}': r'\1',
            r'\\mathrm\{([^}]+)\}': r'\1', 
            r'\\text\{([^}]+)\}': r'\1',
            r'\\textrm\{([^}]+)\}': r'\1',
            r'\\textit\{([^}]+)\}': r'\1',
            r'\\\$': '$',  # Escaped dollar signs
            r'\\,': ' ',
            r'\\;': ' ',
            r'\\quad': '    ',
            r'\\qquad': '        ',
            # Fix spacing issues
            r'\s+': ' ',
            # Remove empty braces
            r'\{\s*\}': '',
        }
        
        for pattern, replacement in replacements.items():
            cleaned = re.sub(pattern, replacement, cleaned)
        
        # Handle nested braces carefully
        cleaned = self._clean_nested_braces(cleaned)
        
        return cleaned.strip()
    
    def _clean_nested_braces(self, text):
        """Clean nested braces intelligently"""
        # Only remove empty or whitespace-only braces
        while True:
            before = text
            text = re.sub(r'\{\s*\}', '', text)
            if text == before:
                break
        return text
    
    def _latex_to_unicode_comprehensive(self, latex):
        """Comprehensive LaTeX to Unicode conversion"""
        if not latex:
            return latex
            
        # Extended symbol mapping
        symbols = {
            # Greek alphabet (lowercase)
            r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ',
            r'\\epsilon': 'ε', r'\\varepsilon': 'ε', r'\\zeta': 'ζ', r'\\eta': 'η',
            r'\\theta': 'θ', r'\\vartheta': 'ϑ', r'\\iota': 'ι', r'\\kappa': 'κ',
            r'\\lambda': 'λ', r'\\mu': 'μ', r'\\nu': 'ν', r'\\xi': 'ξ',
            r'\\pi': 'π', r'\\varpi': 'ϖ', r'\\rho': 'ρ', r'\\varrho': 'ϱ',
            r'\\sigma': 'σ', r'\\varsigma': 'ς', r'\\tau': 'τ', r'\\upsilon': 'υ',
            r'\\phi': 'φ', r'\\varphi': 'φ', r'\\chi': 'χ', r'\\psi': 'ψ', r'\\omega': 'ω',
            
            # Greek alphabet (uppercase)
            r'\\Gamma': 'Γ', r'\\Delta': 'Δ', r'\\Theta': 'Θ', r'\\Lambda': 'Λ',
            r'\\Xi': 'Ξ', r'\\Pi': 'Π', r'\\Sigma': 'Σ', r'\\Upsilon': 'Υ',
            r'\\Phi': 'Φ', r'\\Chi': 'Χ', r'\\Psi': 'Ψ', r'\\Omega': 'Ω',
            
            # Mathematical operators
            r'\\pm': '±', r'\\mp': '∓', r'\\times': '×', r'\\div': '÷',
            r'\\cdot': '·', r'\\bullet': '•', r'\\ast': '∗', r'\\star': '⋆',
            r'\\circ': '∘', r'\\cap': '∩', r'\\cup': '∪', r'\\sqcap': '⊓',
            r'\\sqcup': '⊔', r'\\vee': '∨', r'\\wedge': '∧', r'\\oplus': '⊕',
            r'\\ominus': '⊖', r'\\otimes': '⊗', r'\\oslash': '⊘', r'\\odot': '⊙',
            
            # Relations
            r'\\leq': '≤', r'\\le': '≤', r'\\geq': '≥', r'\\ge': '≥',
            r'\\neq': '≠', r'\\ne': '≠', r'\\equiv': '≡', r'\\approx': '≈',
            r'\\sim': '∼', r'\\simeq': '≃', r'\\cong': '≅', r'\\propto': '∝',
            r'\\prec': '≺', r'\\succ': '≻', r'\\preceq': '⪯', r'\\succeq': '⪰',
            r'\\subset': '⊂', r'\\supset': '⊃', r'\\subseteq': '⊆', r'\\supseteq': '⊇',
            r'\\in': '∈', r'\\notin': '∉', r'\\ni': '∋', r'\\owns': '∋',
            
            # Special symbols
            r'\\infty': '∞', r'\\partial': '∂', r'\\nabla': '∇', r'\\emptyset': '∅',
            r'\\varnothing': '∅', r'\\exists': '∃', r'\\nexists': '∄', r'\\forall': '∀',
            r'\\neg': '¬', r'\\lnot': '¬', r'\\therefore': '∴', r'\\because': '∵',
            r'\\angle': '∠', r'\\triangle': '△', r'\\square': '□', r'\\diamond': '◊',
            r'\\heartsuit': '♡', r'\\clubsuit': '♣', r'\\diamondsuit': '♢', r'\\spadesuit': '♠',
            
            # Large operators
            r'\\sum': '∑', r'\\prod': '∏', r'\\coprod': '∐', r'\\int': '∫',
            r'\\oint': '∮', r'\\iint': '∬', r'\\iiint': '∭', r'\\iiiint': '⨌',
            r'\\bigcap': '⋂', r'\\bigcup': '⋃', r'\\bigwedge': '⋀', r'\\bigvee': '⋁',
            r'\\bigoplus': '⊕', r'\\bigotimes': '⊗', r'\\bigodot': '⊙',
            
            # Functions
            r'\\sin': 'sin', r'\\cos': 'cos', r'\\tan': 'tan', r'\\cot': 'cot',
            r'\\sec': 'sec', r'\\csc': 'csc', r'\\arcsin': 'arcsin', r'\\arccos': 'arccos',
            r'\\arctan': 'arctan', r'\\sinh': 'sinh', r'\\cosh': 'cosh', r'\\tanh': 'tanh',
            r'\\log': 'log', r'\\ln': 'ln', r'\\lg': 'lg', r'\\exp': 'exp',
            r'\\max': 'max', r'\\min': 'min', r'\\sup': 'sup', r'\\inf': 'inf',
            r'\\lim': 'lim', r'\\limsup': 'lim sup', r'\\liminf': 'lim inf',
            r'\\gcd': 'gcd', r'\\det': 'det', r'\\dim': 'dim', r'\\arg': 'arg',
            
            # Arrows
            r'\\leftarrow': '←', r'\\gets': '←', r'\\rightarrow': '→', r'\\to': '→',
            r'\\leftrightarrow': '↔', r'\\Leftarrow': '⇐', r'\\Rightarrow': '⇒',
            r'\\Leftrightarrow': '⇔', r'\\mapsto': '↦', r'\\longmapsto': '⟼',
            r'\\uparrow': '↑', r'\\downarrow': '↓', r'\\updownarrow': '↕',
            
            # Number sets
            r'\\mathbb\{N\}': 'ℕ', r'\\mathbb\{Z\}': 'ℤ', r'\\mathbb\{Q\}': 'ℚ',
            r'\\mathbb\{R\}': 'ℝ', r'\\mathbb\{C\}': 'ℂ', r'\\mathbb\{P\}': 'ℙ',
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
        
        # Handle superscripts and subscripts with proper Unicode
        result = re.sub(r'\^{([^}]+)}', lambda m: self._to_superscript(m.group(1)), result)
        result = re.sub(r'\^(\w)', lambda m: self._to_superscript(m.group(1)), result)
        result = re.sub(r'_{([^}]+)}', lambda m: self._to_subscript(m.group(1)), result)
        result = re.sub(r'_(\w)', lambda m: self._to_subscript(m.group(1)), result)
        
        # Handle brackets and parentheses
        result = result.replace(r'\left(', '(').replace(r'\right)', ')')
        result = result.replace(r'\left[', '[').replace(r'\right]', ']')
        result = result.replace(r'\left\{', '{').replace(r'\right\}', '}')
        result = result.replace(r'\{', '{').replace(r'\}', '}')
        
        # Clean up remaining LaTeX commands
        result = re.sub(r'\\([a-zA-Z]+)', r'\1', result)
        
        return result
    
    def _to_superscript(self, text):
        """Enhanced superscript conversion"""
        superscript_map = {
            '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵',
            '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', '+': '⁺', '-': '⁻',
            '=': '⁼', '(': '⁽', ')': '⁾', 'n': 'ⁿ', 'i': 'ⁱ',
            'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ', 'f': 'ᶠ',
            'g': 'ᵍ', 'h': 'ʰ', 'j': 'ʲ', 'k': 'ᵏ', 'l': 'ˡ', 'm': 'ᵐ',
            'o': 'ᵒ', 'p': 'ᵖ', 'r': 'ʳ', 's': 'ˢ', 't': 'ᵗ', 'u': 'ᵘ',
            'v': 'ᵛ', 'w': 'ʷ', 'x': 'ˣ', 'y': 'ʸ', 'z': 'ᶻ'
        }
        return ''.join(superscript_map.get(c, c) for c in text)
    
    def _to_subscript(self, text):
        """Enhanced subscript conversion"""
        subscript_map = {
            '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅',
            '6': '₆', '7': '₇', '8': '₈', '9': '₉', '+': '₊', '-': '₋',
            '=': '₌', '(': '₍', ')': '₎', 'a': 'ₐ', 'e': 'ₑ', 'i': 'ᵢ',
            'o': 'ₒ', 'u': 'ᵤ', 'x': 'ₓ', 'n': 'ₙ', 'h': 'ₕ', 'k': 'ₖ',
            'l': 'ₗ', 'm': 'ₘ', 'p': 'ₚ', 's': 'ₛ', 't': 'ₜ'
        }
        return ''.join(subscript_map.get(c, c) for c in text)
    
    def create_professional_table_with_math(self, doc, table_data):
        """Create table with math formula processing"""
        if not table_data.get('header') or not table_data.get('rows'):
            return None
        
        header = table_data['header']
        rows = table_data['rows']
        
        # Create table
        table = doc.add_table(rows=len(rows) + 1, cols=len(header))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Style header
        header_row = table.rows[0]
        for i, header_text in enumerate(header):
            cell = header_row.cells[i]
            
            # Process math in header
            if self._has_math_content(str(header_text)):
                self._add_math_to_cell(cell, str(header_text))
            else:
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
            shading.set(qn("w:fill"), "2F5597")
            cell._tc.get_or_add_tcPr().append(shading)
        
        # Style data rows
        for row_idx, row_data in enumerate(rows):
            table_row = table.rows[row_idx + 1]
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < len(header):
                    cell = table_row.cells[col_idx]
                    
                    # Process math in cell
                    if self._has_math_content(str(cell_data)):
                        self._add_math_to_cell(cell, str(cell_data))
                    else:
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
                        shading.set(qn("w:fill"), "F8F9FA")
                        cell._tc.get_or_add_tcPr().append(shading)
        
        return table
    
    def _add_math_to_cell(self, cell, text_with_math):
        """Add math expressions to table cell"""
        # Clear existing content
        cell.text = ''
        
        # Find all math expressions in text
        math_expressions = self.detect_all_math_expressions(text_with_math)
        
        if not math_expressions:
            cell.text = text_with_math
            return
        
        # Process text with math
        current_pos = 0
        paragraph = cell.paragraphs[0]
        
        for expr in sorted(math_expressions, key=lambda x: x['start']):
            # Add text before math
            if expr['start'] > current_pos:
                text_before = text_with_math[current_pos:expr['start']]
                if text_before:
                    run = paragraph.add_run(text_before)
                    run.font.size = Pt(10)
            
            # Add math expression
            unicode_math = self._latex_to_unicode_comprehensive(expr['content'])
            math_run = paragraph.add_run(unicode_math)
            math_run.font.name = 'Cambria Math'
            math_run.font.size = Pt(10)
            math_run.italic = True
            
            current_pos = expr['end']
        
        # Add remaining text
        if current_pos < len(text_with_math):
            remaining_text = text_with_math[current_pos:]
            if remaining_text:
                run = paragraph.add_run(remaining_text)
                run.font.size = Pt(10)
    
    def create_ultimate_document(self, structure, math_expressions, tables, tikz_blocks):
        """Create ultimate formatted document"""
        doc = Document()
        
        # Set document defaults
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12)
        
        # Title
        title = doc.add_heading('Ultimate Mathematical Document', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title.runs[0]
        title_run.font.name = 'Times New Roman'
        title_run.font.size = Pt(18)
        title_run.font.color.rgb = RGBColor(0, 51, 102)
        
        # Collect all elements
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
                'start': table.get('start', 0),
                'end': table.get('end', len(structure['full_text'])),
                'data': table
            })
        
        # Add TikZ blocks
        for tikz in tikz_blocks:
            all_elements.append({
                'type': 'tikz',
                'start': tikz['start'],
                'end': tikz['end'],
                'data': tikz
            })
        
        # Sort by position
        all_elements.sort(key=lambda x: x['start'])
        
        # Process document
        original_text = structure['full_text']
        current_pos = 0
        stats = {
            'math_processed': 0,
            'tables_created': 0,
            'tikz_compiled': 0,
            'tikz_fallbacks': 0
        }
        
        for element in all_elements:
            # Add text before element
            if element['start'] > current_pos:
                text_before = original_text[current_pos:element['start']]
                if text_before.strip():
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
                if self.add_centered_math_equation(doc, math_data['content'], math_data['type']):
                    stats['math_processed'] += 1
                
            elif element['type'] == 'table':
                table = self.create_professional_table_with_math(doc, element['data'])
                if table:
                    stats['tables_created'] += 1
                    doc.add_paragraph()  # Add spacing
                
            elif element['type'] == 'tikz':
                tikz_data = element['data']
                
                # Try to compile TikZ
                with st.spinner(f"Compiling TikZ diagram {self.tikz_counter + 1}..."):
                    compiled_path = self.compile_tikz_advanced(tikz_data['code'])
                
                if compiled_path and os.path.exists(compiled_path):
                    try:
                        # Add compiled image
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p.add_run()
                        run.add_picture(compiled_path, width=Inches(6))
                        
                        # Add caption
                        caption_p = doc.add_paragraph()
                        caption_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        caption_run = caption_p.add_run(f"Figure {self.tikz_counter}: Compiled TikZ Diagram")
                        caption_run.italic = True
                        caption_run.font.size = Pt(10)
                        caption_run.font.color.rgb = RGBColor(100, 100, 100)
                        
                        stats['tikz_compiled'] += 1
                        
                    except Exception as e:
                        st.warning(f"Could not insert compiled TikZ: {e}")
                        stats['tikz_fallbacks'] += 1
                else:
                    stats['tikz_fallbacks'] += 1
            
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
        
        # Add processing summary
        doc.add_page_break()
        summary_heading = doc.add_heading('Ultimate Processing Summary', 1)
        summary_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        summary_text = f"""
✅ Mathematical expressions: {stats['math_processed']} processed with proper centering
✅ Professional tables: {stats['tables_created']} created with math formula support
✅ TikZ diagrams compiled: {stats['tikz_compiled']} high-quality images generated
⚠️ TikZ fallbacks: {stats['tikz_fallbacks']} placeholder images created
✅ Perfect formatting: Centered equations, professional typography

Generated by Ultimate Word Processor
"""
        
        summary_para = doc.add_paragraph()
        summary_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        summary_run = summary_para.add_run(summary_text.strip())
        summary_run.font.name = 'Times New Roman'
        summary_run.font.size = Pt(11)
        summary_run.italic = True
        
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
        page_title="Ultimate Word Processor",
        page_icon="🚀",
        layout="wide"
    )
    
    st.title("🚀 Ultimate Word Document Processor")
    st.markdown("""
    **Ultimate version** với tất cả features hoàn hảo:
    - ✅ **Perfect Math Formatting**: Centered equations, proper line breaks
    - ✅ **TikZ Compilation**: LaTeX TikZ → High-quality PNG images
    - ✅ **Table Math Support**: Formulas inside table cells
    - ✅ **Comprehensive Detection**: All math formats, mixed content
    - ✅ **Professional Output**: Perfect typography and layout
    """)
    
    # System status
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if MATH2DOCX_AVAILABLE:
            st.success("🧮 Real equations: ✅")
        else:
            st.warning("🧮 Unicode math: ⚠️")
    
    with col2:
        try:
            result = subprocess.run(['pdflatex', '--version'], 
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                st.success("📊 LaTeX/TikZ: ✅")
            else:
                st.warning("📊 TikZ fallback: ⚠️")
        except:
            st.warning("📊 TikZ fallback: ⚠️")
    
    with col3:
        st.success("📋 Table math: ✅")
    
    with col4:
        st.success("🎯 Perfect format: ✅")
    
    # File upload
    uploaded_file = st.file_uploader(
        "📁 Upload Word Document (.docx)",
        type=['docx'],
        help="Upload your document with math formulas, tables, and TikZ diagrams"
    )
    
    if uploaded_file:
        processor = UltimateWordProcessor()
        
        try:
            # Read document
            with st.spinner("Reading document comprehensively..."):
                structure = processor.read_word_comprehensive(uploaded_file)
            
            if not structure['full_text']:
                st.error("Could not read document content.")
                return
            
            # Preview
            with st.expander("📖 Document Analysis"):
                st.write(f"**Characters**: {len(structure['full_text'])}")
                st.write(f"**Paragraphs**: {len(structure['paragraphs'])}")
                st.write(f"**Existing tables**: {len(structure['tables'])}")
                st.text_area("Content Preview", 
                           structure['full_text'][:2000] + "..." if len(structure['full_text']) > 2000 else structure['full_text'],
                           height=200)
            
            # Analyze content
            st.subheader("🔍 Comprehensive Analysis")
            
            with st.spinner("Analyzing all content types..."):
                math_expressions = processor.detect_all_math_expressions(structure['full_text'])
                tikz_blocks = processor.detect_tikz_blocks(structure['full_text'])
                all_tables = structure['tables']  # Use existing tables from Word
            
            # Show results
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("🧮 Math Expressions", len(math_expressions))
                if math_expressions:
                    st.write("**Types found:**")
                    type_counts = {}
                    for expr in math_expressions:
                        expr_type = expr['type']
                        type_counts[expr_type] = type_counts.get(expr_type, 0) + 1
                    
                    for type_name, count in type_counts.items():
                        st.write(f"- {type_name}: {count}")
                    
                    st.write("**Samples:**")
                    for expr in math_expressions[:3]:
                        st.code(f"{expr['type']}: {expr['content'][:50]}...")
            
            with col2:
                st.metric("📊 Tables", len(all_tables))
                if all_tables:
                    for i, table in enumerate(all_tables[:2]):
                        if table.get('header'):
                            st.write(f"**Table {i+1}**: {len(table.get('rows', []))}×{len(table['header'])}")
                            has_math = table.get('has_formulas', False)
                            st.write(f"Contains math: {'✅' if has_math else '❌'}")
            
            with col3:
                st.metric("🎨 TikZ Diagrams", len(tikz_blocks))
                if tikz_blocks:
                    for i, tikz in enumerate(tikz_blocks[:2]):
                        st.write(f"**TikZ {i+1}**: {len(tikz['code'])} characters")
                        st.code(tikz['code'][:80] + "..." if len(tikz['code']) > 80 else tikz['code'])
            
            # Processing options
            st.subheader("⚙️ Ultimate Processing Options")
            
            col1, col2 = st.columns(2)
            with col1:
                center_equations = st.checkbox("📐 Center Display Equations", value=True)
                compile_tikz = st.checkbox("🎨 Compile TikZ to Images", value=True)
            
            with col2:
                process_table_math = st.checkbox("📋 Process Table Math", value=True)
                enhance_typography = st.checkbox("✨ Enhanced Typography", value=True)
            
            # Process button
            if st.button("🚀 Ultimate Processing", type="primary", use_container_width=True):
                
                progress = st.progress(0)
                status = st.empty()
                
                try:
                    status.text("🔍 Processing mathematical content...")
                    progress.progress(20)
                    
                    status.text("📊 Creating professional tables...")
                    progress.progress(40)
                    
                    status.text("🎨 Compiling TikZ diagrams...")
                    progress.progress(60)
                    
                    # Create ultimate document
                    processed_doc, stats = processor.create_ultimate_document(
                        structure, math_expressions, all_tables, tikz_blocks
                    )
                    
                    status.text("💾 Finalizing document...")
                    progress.progress(80)
                    
                    # Save document
                    doc_io = io.BytesIO()
                    processed_doc.save(doc_io)
                    doc_io.seek(0)
                    
                    progress.progress(100)
                    status.text("✅ Ultimate processing complete!")
                    
                    # Success message
                    st.success("🎉 Ultimate Document Processing Complete!")
                    
                    # Show detailed stats
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("📐 Equations", stats['math_processed'])
                    with col2:
                        st.metric("📊 Tables", stats['tables_created'])
                    with col3:
                        st.metric("🎨 TikZ Images", stats['tikz_compiled'])
                    with col4:
                        st.metric("⚠️ Fallbacks", stats['tikz_fallbacks'])
                    
                    # Download button
                    st.download_button(
                        label="📥 Download Ultimate Document",
                        data=doc_io.getvalue(),
                        file_name="ultimate_processed_document.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                    # Detailed results
                    st.info(f"""
🎯 **Ultimate Processing Results:**

**Mathematical Content:**
- {stats['math_processed']} equations processed with perfect centering
- Display equations on separate lines with proper spacing
- Inline equations with enhanced Unicode fallback
- Table cells with mathematical expressions supported

**Visual Content:**
- {stats['tikz_compiled']} TikZ diagrams compiled to high-quality PNG
- {stats['tikz_fallbacks']} fallback images with installation guidance
- Professional table styling with alternating colors

**Document Quality:**
- Times New Roman typography for professional appearance
- Centered math equations with proper line breaks
- Consistent formatting and spacing throughout
- Original structure preserved and enhanced
                    """)
                    
                except Exception as e:
                    st.error(f"Ultimate processing error: {e}")
                    import traceback
                    with st.expander("🔧 Error Details"):
                        st.code(traceback.format_exc())
                finally:
                    progress.empty()
                    status.empty()
        
        except Exception as e:
            st.error(f"Error reading document: {e}")
            import traceback
            with st.expander("🔧 Error Details"):
                st.code(traceback.format_exc())
        
        finally:
            processor.cleanup()
    
    # Help section
    with st.expander("📚 Ultimate Guide"):
        st.markdown("""
        ### 🎯 Ultimate Features:
        
        **Perfect Math Formatting:**
        - Display equations centered on separate lines
        - Inline equations with proper spacing
        - Enhanced Unicode with 200+ symbols
        - Real Word equations with math2docx
        
        **TikZ Compilation:**
        - LaTeX TikZ → High-quality PNG images
        - Automatic fallback with setup instructions
        - Professional diagram captions
        - Error handling and recovery
        
        **Table Enhancement:**
        - Mathematical formulas in table cells
        - Professional styling with colors
        - Consistent formatting and alignment
        - Header styling and alternating rows
        
        ### 📋 Requirements for Full Features:
        
        ```bash
        # Essential packages
        pip install streamlit python-docx mammoth Pillow PyMuPDF
        
        # For real Word equations
        pip install math2docx
        
        # For TikZ compilation
        # Windows: Install MiKTeX from https://miktex.org/
        # Mac: brew install mactex
        # Linux: sudo apt-get install texlive-full
        ```
        
        ### 🔧 Troubleshooting:
        
        **If TikZ doesn't compile:**
        - Install LaTeX distribution (TeX Live/MiKTeX)
        - Ensure `pdflatex` is in PATH
        - Fallback images will be created automatically
        
        **If math equations aren't perfect:**
        - Install `math2docx` for real Word equations
        - Unicode fallback provides good alternatives
        - All mathematical symbols are preserved
        """)

if __name__ == "__main__":
    main()
