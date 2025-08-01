#!/usr/bin/env python3
"""
Streamlit LaTeX to Word Converter
Beautiful interface for converting LaTeX expressions in Word documents
"""

import streamlit as st
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from io import BytesIO
from datetime import datetime
import traceback

# Page configuration
st.set_page_config(
    page_title="LaTeX to Word Converter",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 2rem;
        color: white;
    }
    .main-header h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: bold;
    }
    .main-header p {
        margin: 0.5rem 0 0 0;
        font-size: 1.1rem;
        opacity: 0.9;
    }
    .feature-card {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        margin: 1rem 0;
    }
    .success-box {
        background: linear-gradient(90deg, #56ab2f 0%, #a8e6cf 100%);
        padding: 1rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin: 1rem 0;
    }
    .warning-box {
        background: linear-gradient(90deg, #f093fb 0%, #f5576c 100%);
        padding: 1rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin: 1rem 0;
    }
    .stButton > button {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
    }
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

def convert_latex_symbols(text):
    """Enhanced LaTeX to Unicode conversion"""
    if not text:
        return text
    
    result = text
    
    # Prime notation patterns
    prime_patterns = [
        (r'([A-Za-z_]+)\^\{prime\}', r"\1'"),  # A^{prime} -> A'
        (r'([A-Za-z_]+)\s*\(prime\)', r"\1'"), # A(prime) -> A'
        (r'\(prime\)', "'"),                    # (prime) -> '
        (r'\^prime', "'"),                      # ^prime -> '
        (r'prime', "'"),                        # prime -> '
    ]
    
    for pattern, replacement in prime_patterns:
        result = re.sub(pattern, replacement, result)
    
    # Superscript with braces: A^{2} -> A²
    superscript_map = {
        '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', 
        '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
        '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾'
    }
    
    # Handle ^{...}
    while '^{' in result:
        pos = result.find('^{')
        if pos == -1:
            break
        
        # Find base
        base_start = pos - 1
        while base_start >= 0 and (result[base_start].isalnum() or result[base_start] in '()[]_'):
            base_start -= 1
        base_start += 1
        base = result[base_start:pos]
        
        # Find content in braces
        brace_start = pos + 2
        brace_count = 1
        brace_end = brace_start
        
        while brace_end < len(result) and brace_count > 0:
            if result[brace_end] == '{':
                brace_count += 1
            elif result[brace_end] == '}':
                brace_count -= 1
            brace_end += 1
        
        if brace_count == 0:
            content = result[brace_start:brace_end-1]
            # Convert to superscript if possible
            if len(content) == 1 and content in superscript_map:
                replacement = base + superscript_map[content]
            elif content.isdigit() and len(content) <= 3:
                sup_text = ''.join(superscript_map.get(c, c) for c in content)
                replacement = base + sup_text
            else:
                replacement = base + "^(" + content + ")"
            result = result[:base_start] + replacement + result[brace_end:]
        else:
            break
    
    # Subscript: A_{n} -> Aₙ
    subscript_map = {
        '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
        '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
        '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
        'a': 'ₐ', 'e': 'ₑ', 'i': 'ᵢ', 'o': 'ₒ', 'u': 'ᵤ',
        'n': 'ₙ', 'r': 'ᵣ', 's': 'ₛ', 't': 'ₜ', 'x': 'ₓ'
    }
    
    # Handle _{...}
    while '_{' in result:
        pos = result.find('_{')
        if pos == -1:
            break
        
        # Find base
        base_start = pos - 1
        while base_start >= 0 and (result[base_start].isalnum() or result[base_start] in '()[]'):
            base_start -= 1
        base_start += 1
        base = result[base_start:pos]
        
        # Find content in braces
        brace_start = pos + 2
        brace_count = 1
        brace_end = brace_start
        
        while brace_end < len(result) and brace_count > 0:
            if result[brace_end] == '{':
                brace_count += 1
            elif result[brace_end] == '}':
                brace_count -= 1
            brace_end += 1
        
        if brace_count == 0:
            content = result[brace_start:brace_end-1]
            # Convert to subscript if possible
            if len(content) <= 3:
                sub_text = ''.join(subscript_map.get(c, c) for c in content)
                replacement = base + sub_text
            else:
                replacement = base + "_(" + content + ")"
            result = result[:base_start] + replacement + result[brace_end:]
        else:
            break
    
    # Mathematical symbols
    symbols = {
        '\\pi': 'π', '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
        '\\theta': 'θ', '\\lambda': 'λ', '\\mu': 'μ', '\\sigma': 'σ', '\\phi': 'φ',
        '\\omega': 'ω', '\\Omega': 'Ω', '\\Phi': 'Φ', '\\Theta': 'Θ', '\\Delta': 'Δ',
        '\\in': '∈', '\\subset': '⊂', '\\cup': '∪', '\\cap': '∩', '\\emptyset': '∅',
        '\\leq': '≤', '\\geq': '≥', '\\neq': '≠', '\\approx': '≈', '\\equiv': '≡',
        '\\infty': '∞', '\\pm': '±', '\\mp': '∓', '\\times': '×', '\\div': '÷',
        '\\cdot': '·', '\\bullet': '•', '\\circ': '∘',
        '\\mathbb{R}': 'ℝ', '\\mathbb{Z}': 'ℤ', '\\mathbb{Q}': 'ℚ', 
        '\\mathbb{N}': 'ℕ', '\\mathbb{C}': 'ℂ', '\\mathbb{P}': 'ℙ',
        '\\sin': 'sin', '\\cos': 'cos', '\\tan': 'tan', '\\cot': 'cot',
        '\\sec': 'sec', '\\csc': 'csc', '\\arcsin': 'arcsin', '\\arccos': 'arccos',
        '\\lim': 'lim', '\\sup': 'sup', '\\inf': 'inf', '\\max': 'max', '\\min': 'min',
        '\\sum': '∑', '\\prod': '∏', '\\int': '∫', '\\oint': '∮',
        '\\partial': '∂', '\\nabla': '∇', '\\exists': '∃', '\\forall': '∀',
        '\\therefore': '∴', '\\because': '∵', '\\QED': '∎'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Handle parentheses and brackets
    result = result.replace('\\left(', '(').replace('\\right)', ')')
    result = result.replace('\\left[', '[').replace('\\right]', ']')
    result = result.replace('\\left{', '{').replace('\\right}', '}')
    result = result.replace('\\lbrack', '[').replace('\\rbrack', ']')
    result = result.replace('\\{', '{').replace('\\}', '}')
    
    # Handle fractions: \frac{a}{b} -> (a)/(b)
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        try:
            # Find numerator
            num_start = start + 6
            brace_count = 1
            num_end = num_start
            
            while num_end < len(result) and brace_count > 0:
                if result[num_end] == '{':
                    brace_count += 1
                elif result[num_end] == '}':
                    brace_count -= 1
                num_end += 1
            
            numerator = result[num_start:num_end-1]
            
            # Find denominator
            if num_end < len(result) and result[num_end] == '{':
                denom_start = num_end + 1
                brace_count = 1
                denom_end = denom_start
                
                while denom_end < len(result) and brace_count > 0:
                    if result[denom_end] == '{':
                        brace_count += 1
                    elif result[denom_end] == '}':
                        brace_count -= 1
                    denom_end += 1
                
                denominator = result[denom_start:denom_end-1]
                fraction = f"({numerator})/({denominator})"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        except:
            break
    
    # Handle square root: \sqrt{x} -> √(x)
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
        
        try:
            content_start = start + 6
            brace_count = 1
            content_end = content_start
            
            while content_end < len(result) and brace_count > 0:
                if result[content_end] == '{':
                    brace_count += 1
                elif result[content_end] == '}':
                    brace_count -= 1
                content_end += 1
            
            content = result[content_start:content_end-1]
            sqrt_result = f"√({content})"
            result = result[:start] + sqrt_result + result[content_end:]
        except:
            break
    
    # Clean up remaining backslashes
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Find mathematical expressions in text"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Handle ${...}$ format
            if i < len(text) and text[i] == '{':
                i += 1
                brace_count = 1
                content_start = i
                
                while i < len(text) and brace_count > 0:
                    if text[i] == '{':
                        brace_count += 1
                    elif text[i] == '}':
                        brace_count -= 1
                    i += 1
                
                if i < len(text) and text[i] == '$':
                    end = i + 1
                    content = text[content_start:i-1]
                    if content.strip():
                        expressions.append((start, end, content.strip()))
                    i += 1
                else:
                    i += 1
            else:
                # Handle $...$ format
                content_start = i
                
                while i < len(text) and text[i] != '$':
                    i += 1
                
                if i < len(text):
                    end = i + 1
                    content = text[content_start:i]
                    if content.strip():
                        expressions.append((start, end, content.strip()))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_equation_object(paragraph, latex_text, method='auto'):
    """Create equation object in Word"""
    unicode_text = convert_latex_symbols(latex_text)
    
    if method == 'omml' or method == 'auto':
        try:
            # OMML equation object
            safe_text = unicode_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            # Determine equation type for better formatting
            if "'" in unicode_text and any(c.isalpha() for c in unicode_text):
                # Prime notation
                omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_text}</m:t>
                    </m:r>
                </m:oMath>'''
            elif any(c in unicode_text for c in ['²', '³', '⁴', '⁵', '⁶', '⁷', '⁸', '⁹', '₁', '₂', '₃', '₄', '₅']):
                # Superscript/subscript
                omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_text}</m:t>
                    </m:r>
                </m:oMath>'''
            else:
                # General math
                omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_text}</m:t>
                    </m:r>
                </m:oMath>'''
            
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return 'omml'
        except Exception as e:
            if method == 'omml':
                raise e
    
    if method == 'styled' or method == 'auto':
        try:
            # Styled text approach
            run = paragraph.add_run(unicode_text)
            run.font.name = 'Cambria Math'
            run.font.size = Pt(12)
            run.italic = True
            run.font.color.rgb = RGBColor(0, 32, 96)
            return 'styled'
        except Exception as e:
            if method == 'styled':
                raise e
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_text)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)
        return 'fallback'
    except:
        run = paragraph.add_run(f"[{latex_text}]")
        return 'error'

def process_text_with_math(paragraph, text, options):
    """Process text containing mathematical expressions"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0, {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    last_pos = 0
    math_count = 0
    method_stats = {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    for start, end, latex_content in expressions:
        # Add text before math
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Create equation
        method_used = create_equation_object(paragraph, latex_content, options.get('equation_method', 'auto'))
        method_stats[method_used] += 1
        math_count += 1
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count, method_stats

def format_question_answer(paragraph, text, options):
    """Format questions and answers with Vietnamese support"""
    paragraph.clear()
    
    # Question pattern
    question_patterns = [
        r'^(Câu\s+\d+[\.:])\s*(.*)',
        r'^(Question\s+\d+[\.:])\s*(.*)',
        r'^(\d+[\.:])\s*(.*)'
    ]
    
    for pattern in question_patterns:
        match = re.match(pattern, text, re.IGNORECASE)
        if match:
            question_part = match.group(1)
            content_part = match.group(2)
            
            # Bold question number
            run_q = paragraph.add_run(question_part)
            run_q.bold = True
            run_q.font.size = Pt(12)
            run_q.font.name = 'Times New Roman'
            run_q.font.color.rgb = RGBColor(0, 0, 0)
            
            if content_part:
                paragraph.add_run(" ")
                process_text_with_math(paragraph, content_part, options)
            
            return True
    
    # Answer pattern
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Bold answer letter
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content, options)
        
        return True
    
    # Not Q&A format
    process_text_with_math(paragraph, text, options)
    return False

def process_document(doc, options):
    """Process Word document and convert LaTeX to equation objects"""
    new_doc = Document()
    
    stats = {
        'paragraphs': 0,
        'tables': 0,
        'math_expressions': 0,
        'questions_formatted': 0,
        'method_stats': {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0},
        'processing_log': []
    }
    
    # Process paragraphs
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        stats['paragraphs'] += 1
        
        if not text:
            new_doc.add_paragraph()
            continue
        
        new_para = new_doc.add_paragraph()
        
        try:
            if options.get('format_qa', True) and format_question_answer(new_para, text, options):
                stats['questions_formatted'] += 1
                stats['processing_log'].append(f"Para {i+1}: Q&A formatted")
            else:
                math_count, method_stats = process_text_with_math(new_para, text, options)
                stats['math_expressions'] += math_count
                for method, count in method_stats.items():
                    stats['method_stats'][method] += count
                
                if math_count > 0:
                    stats['processing_log'].append(f"Para {i+1}: {math_count} equations converted")
        except Exception as e:
            # Fallback to plain text
            new_para.clear()
            run = new_para.add_run(text)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
            stats['processing_log'].append(f"Para {i+1}: Error, used plain text - {str(e)}")
    
    # Process tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        try:
            num_rows = len(table.rows)
            num_cols = len(table.rows[0].cells) if table.rows else 0
            
            new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
            
            table_math = 0
            for r in range(num_rows):
                for c in range(min(len(table.rows[r].cells), num_cols)):
                    try:
                        original_cell = table.rows[r].cells[c]
                        new_cell = new_table.rows[r].cells[c]
                        
                        cell_text = original_cell.text.strip()
                        if cell_text:
                            new_cell.text = ""
                            cell_para = new_cell.paragraphs[0] if new_cell.paragraphs else new_cell.add_paragraph()
                            cell_para.clear()
                            
                            if not (options.get('format_qa', True) and format_question_answer(cell_para, cell_text, options)):
                                cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text, options)
                                table_math += cell_math_count
                                for method, count in cell_method_stats.items():
                                    stats['method_stats'][method] += count
                            
                            cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except Exception:
                        continue
            
            stats['math_expressions'] += table_math
            
            # Format table
            if options.get('format_tables', True):
                try:
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    # Style header row
                    if new_table.rows:
                        header_row = new_table.rows[0]
                        for cell in header_row.cells:
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.bold = True
                                    run.font.color.rgb = RGBColor(255, 255, 255)
                            
                            # Blue header background
                            try:
                                shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="4472C4"/>'
                                shading = parse_xml(shading_xml)
                                cell._tc.get_or_add_tcPr().append(shading)
                            except:
                                pass
                except Exception:
                    pass
            
            stats['tables'] += 1
            if table_math > 0:
                stats['processing_log'].append(f"Table {table_idx+1}: {table_math} equations converted")
                
        except Exception as e:
            stats['processing_log'].append(f"Table {table_idx+1}: Error - {str(e)}")
            continue
    
    return new_doc, stats

# Main Streamlit App
def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📚 LaTeX to Word Converter</h1>
        <p>Chuyển đổi công thức LaTeX thành Equation Objects thực sự trong Word</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("## ⚙️ Cài đặt")
        
        # Processing options
        st.markdown("### 🔧 Tùy chọn xử lý")
        format_qa = st.checkbox("📝 Định dạng câu hỏi & đáp án", value=True, 
                               help="Định dạng 'Câu 1:', 'A.', 'B.', etc.")
        format_tables = st.checkbox("📊 Định dạng bảng", value=True,
                                  help="Tạo header màu xanh và borders")
        
        equation_method = st.selectbox(
            "🧮 Phương pháp tạo equation",
            ["auto", "omml", "styled"],
            index=0,
            help="auto: Tự động chọn tốt nhất\nomml: OMML equation objects\nstyled: Text có styling đẹp"
        )
        
        st.markdown("### 🎯 Preview LaTeX")
        test_latex = st.text_input("Test công thức:", placeholder="u_{n}")
        if test_latex:
            try:
                converted = convert_latex_symbols(test_latex)
                st.code(f"${test_latex}$ → {converted}")
            except Exception:
                st.error("Lỗi chuyển đổi")
        
        st.markdown("---")
        st.markdown("### 📖 Hướng dẫn")
        st.info("""
        **Các định dạng được hỗ trợ:**
        - `$x^2$` → x²
        - `$u_{n}$` → uₙ  
        - `$A^{prime}$` → A'
        - `$\\frac{1}{2}$` → (1)/(2)
        - `$\\pi$` → π
        - `$\\sqrt{x}$` → √(x)
        
        **File hỗ trợ:** .docx
        """)
        
        st.markdown("### 🔬 Thống kê nâng cao")
        show_debug = st.checkbox("Hiện log xử lý", value=False)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("## 📁 Upload File Word")
        uploaded_file = st.file_uploader(
            "Chọn file .docx để xử lý",
            type=["docx"],
            help="Upload file Word chứa công thức LaTeX ($...$) hoặc (${...}$)"
        )
        
        if uploaded_file:
            st.success(f"✅ Đã upload: **{uploaded_file.name}**")
            
            # Preview section
            with st.expander("👀 Xem trước nội dung", expanded=False):
                try:
                    doc = Document(uploaded_file)
                    st.info(f"📄 {len(doc.paragraphs)} đoạn văn | 📊 {len(doc.tables)} bảng")
                    
                    # Show first few paragraphs
                    st.markdown("**Nội dung đầu:**")
                    preview_count = 0
                    for para in doc.paragraphs:
                        if para.text.strip() and preview_count < 5:
                            text = para.text[:100] + "..." if len(para.text) > 100 else para.text
                            st.text(f"{preview_count + 1}. {text}")
                            preview_count += 1
                    
                    # Show LaTeX expressions found
                    all_text = "\n".join([para.text for para in doc.paragraphs])
                    expressions = find_math_expressions(all_text)
                    if expressions:
                        st.markdown("**🧮 Công thức LaTeX tìm thấy:**")
                        unique_expressions = list(set([expr[2] for expr in expressions[:10]]))
                        for expr in unique_expressions:
                            converted = convert_latex_symbols(expr)
                            st.code(f"${expr}$ → {converted}")
                        if len(expressions) > 10:
                            st.info(f"... và {len(expressions) - 10} công thức khác")
                    else:
                        st.warning("❌ Không tìm thấy công thức LaTeX ($...$)")
                        
                except Exception as e:
                    st.error(f"❌ Lỗi đọc file: {e}")
    
    with col2:
        st.markdown("## 🎯 Tính năng chính")
        
        features = [
            ("🧮", "Equation Objects", "Tạo equation objects thực sự trong Word"),
            ("📝", "Q&A Formatting", "Định dạng câu hỏi và đáp án tự động"),  
            ("📊", "Table Formatting", "Bảng đẹp với header màu xanh"),
            ("🔤", "Unicode Support", "Hỗ trợ ký hiệu toán học Unicode"),
            ("🎨", "Professional Style", "Font và styling chuyên nghiệp"),
            ("⚡", "Fast Processing", "Xử lý nhanh chóng và ổn định")
        ]
        
        for icon, title, desc in features:
            st.markdown(f"""
            <div class="feature-card">
                <strong>{icon} {title}</strong><br>
                <small>{desc}</small>
            </div>
            """, unsafe_allow_html=True)
    
    # Processing section
    if uploaded_file:
        st.markdown("---")
        
        if st.button("🚀 Bắt đầu chuyển đổi", type="primary", use_container_width=True):
            options = {
                'format_qa': format_qa,
                'format_tables': format_tables,
                'equation_method': equation_method
            }
            
            with st.spinner("⏳ Đang xử lý file..."):
                try:
                    # Process document
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document(doc, options)
                    
                    # Save to buffer
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    # Generate filename
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_converted_{timestamp}.docx"
                    
                    # Success message
                    st.markdown("""
                    <div class="success-box">
                        <h3>🎉 Chuyển đổi thành công!</h3>
                        <p>File của bạn đã được xử lý và sẵn sàng tải về</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Statistics
                    st.markdown("## 📊 Thống kê xử lý")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.markdown("""
                        <div class="metric-card">
                            <h2 style="color: #667eea; margin: 0;">{}</h2>
                            <p style="margin: 0;">Đoạn văn</p>
                        </div>
                        """.format(stats['paragraphs']), unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown("""
                        <div class="metric-card">
                            <h2 style="color: #28a745; margin: 0;">{}</h2>
                            <p style="margin: 0;">Công thức</p>
                        </div>
                        """.format(stats['math_expressions']), unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown("""
                        <div class="metric-card">
                            <h2 style="color: #ffc107; margin: 0;">{}</h2>
                            <p style="margin: 0;">Câu hỏi</p>
                        </div>
                        """.format(stats['questions_formatted']), unsafe_allow_html=True)
                    
                    with col4:
                        st.markdown("""
                        <div class="metric-card">
                            <h2 style="color: #17a2b8; margin: 0;">{}</h2>
                            <p style="margin: 0;">Bảng</p>
                        </div>
                        """.format(stats['tables']), unsafe_allow_html=True)
                    
                    # Method breakdown
                    if stats['math_expressions'] > 0:
                        st.markdown("### 🔧 Chi tiết phương pháp")
                        method_col1, method_col2, method_col3 = st.columns(3)
                        
                        with method_col1:
                            st.metric("OMML Objects", stats['method_stats']['omml'], 
                                    help="Equation objects thực sự")
                        with method_col2:
                            st.metric("Styled Text", stats['method_stats']['styled'], 
                                    help="Text với styling đẹp")  
                        with method_col3:
                            st.metric("Fallback", stats['method_stats']['fallback'] + stats['method_stats']['error'],
                                    help="Phương pháp dự phòng")
                    
                    # Debug information
                    if show_debug and stats['processing_log']:
                        with st.expander("🔍 Log xử lý chi tiết"):
                            for log_entry in stats['processing_log'][-20:]:  # Show last 20 entries
                                st.text(log_entry)
                    
                    # Download button
                    st.download_button(
                        "📥 Tải file đã chuyển đổi",
                        data=buffer.getvalue(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        type="primary"
                    )
                    
                    # Additional info
                    if stats['math_expressions'] > 0:
                        omml_count = stats['method_stats']['omml']
                        if omml_count > 0:
                            st.success(f"✨ Đã tạo {omml_count} equation objects thực sự!")
                        st.info(f"💡 Tổng cộng {stats['math_expressions']} công thức toán học đã được chuyển đổi")
                    
                    if stats['questions_formatted'] > 0:
                        st.info(f"📝 Đã định dạng {stats['questions_formatted']} câu hỏi/đáp án")
                        
                except Exception as e:
                    st.markdown("""
                    <div class="warning-box">
                        <h3>❌ Lỗi xử lý</h3>
                        <p>Có lỗi xảy ra khi xử lý file</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.error(f"Chi tiết lỗi: {str(e)}")
                    
                    if show_debug:
                        st.text("Stack trace:")
                        st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
