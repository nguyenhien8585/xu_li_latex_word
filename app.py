import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
from io import BytesIO
from datetime import datetime

# Cấu hình trang
st.set_page_config(
    page_title="Xử lý LaTeX & Format trong Word",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

def convert_latex_to_unicode(text):
    """Chuyển đổi LaTeX thành Unicode - approach đơn giản"""
    if not text:
        return text
    
    result = text
    
    # Dictionary cho ký hiệu toán học
    latex_symbols = {
        'pi': 'π',
        'alpha': 'α',
        'beta': 'β',
        'gamma': 'γ',
        'delta': 'δ',
        'theta': 'θ',
        'lambda': 'λ',
        'mu': 'μ',
        'sigma': 'σ',
        'phi': 'φ',
        'omega': 'ω',
        'in': '∈',
        'notin': '∉',
        'subset': '⊂',
        'supset': '⊃',
        'cup': '∪',
        'cap': '∩',
        'emptyset': '∅',
        'leq': '≤',
        'geq': '≥',
        'neq': '≠',
        'approx': '≈',
        'equiv': '≡',
        'infty': '∞',
        'pm': '±',
        'mp': '∓'
    }
    
    # Replace LaTeX symbols
    for latex, unicode_char in latex_symbols.items():
        result = result.replace('\\' + latex, unicode_char)
    
    # Replace number sets
    number_sets = {
        'mathbb{Z}': 'ℤ',
        'mathbb{R}': 'ℝ',
        'mathbb{Q}': 'ℚ',
        'mathbb{N}': 'ℕ',
        'mathbb{C}': 'ℂ'
    }
    
    for latex, unicode_char in number_sets.items():
        result = result.replace('\\' + latex, unicode_char)
    
    # Handle fractions: \frac{a}{b} -> (a)/(b)
    while 'frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        # Find numerator
        num_start = start + 6  # after \frac{
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
            
            # Replace with (a)/(b)
            fraction = f"({numerator})/({denominator})"
            result = result[:start] + fraction + result[denom_end:]
        else:
            break
    
    # Handle sqrt: \sqrt{x} -> √(x)
    while 'sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
            
        content_start = start + 6  # after \sqrt{
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
    
    # Handle superscripts: x^2 -> x²
    superscript_map = {
        '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
        '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'
    }
    
    i = 0
    new_result = ""
    while i < len(result):
        if i < len(result) - 2 and result[i+1] == '^' and result[i+2].isdigit():
            base = result[i]
            power = result[i+2]
            new_result += base + superscript_map.get(power, power)
            i += 3
        else:
            new_result += result[i]
            i += 1
    result = new_result
    
    # Handle subscripts: x_2 -> x₂
    subscript_map = {
        '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
        '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉'
    }
    
    i = 0
    new_result = ""
    while i < len(result):
        if i < len(result) - 2 and result[i+1] == '_' and result[i+2].isdigit():
            base = result[i]
            sub = result[i+2]
            new_result += base + subscript_map.get(sub, sub)
            i += 3
        else:
            new_result += result[i]
            i += 1
    result = new_result
    
    # Clean up remaining backslashes
    result = result.replace('\\', '')
    
    return result

def find_latex_expressions(text):
    """Tìm LaTeX expressions trong text - đơn giản"""
    expressions = []
    
    # Find $...$ patterns
    i = 0
    while i < len(text):
        if text[i] == '$':
            # Start of LaTeX
            start = i
            i += 1
            
            # Find closing $
            while i < len(text) and text[i] != '$':
                i += 1
            
            if i < len(text):  # Found closing $
                end = i + 1
                latex_content = text[start+1:i]
                if latex_content.strip():
                    expressions.append((start, end, latex_content.strip()))
                i += 1
            else:
                i += 1
        else:
            i += 1
    
    return expressions

def create_formatted_run(paragraph, text, is_equation=False):
    """Tạo run với formatting"""
    if is_equation:
        unicode_text = convert_latex_to_unicode(text)
        run = paragraph.add_run(unicode_text)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 100, 0)
        run.font.size = Pt(12)
        run.font.name = 'Cambria Math'
    else:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
    
    return run

def process_text_with_latex(paragraph, text):
    """Xử lý text có LaTeX"""
    latex_count = 0
    
    # Find LaTeX expressions
    expressions = find_latex_expressions(text)
    
    if not expressions:
        create_formatted_run(paragraph, text)
        return 0
    
    # Process text with LaTeX
    last_pos = 0
    
    for start, end, latex_content in expressions:
        # Add text before LaTeX
        if start > last_pos:
            before_text = text[last_pos:start]
            create_formatted_run(paragraph, before_text)
        
        # Add LaTeX as equation
        create_formatted_run(paragraph, latex_content, is_equation=True)
        latex_count += 1
        
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining_text = text[last_pos:]
        create_formatted_run(paragraph, remaining_text)
    
    return latex_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    original_text = text
    
    # Clear paragraph
    paragraph.clear()
    
    # Check for question pattern: Câu X.
    question_match = re.match(r'^(Câu\s+\d+[\.:])\s*(.*)', text, re.IGNORECASE)
    if question_match:
        question_part = question_match.group(1)
        content_part = question_match.group(2)
        
        # Bold question number
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Process content with LaTeX
        if content_part:
            paragraph.add_run(" ")
            process_text_with_latex(paragraph, content_part)
        
        return True
    
    # Check for answer pattern: A., B., C., D.
    answer_match = re.match(r'^([A-Da-d])[\.\)]\s*(.*)', text)
    if answer_match:
        answer_letter = answer_match.group(1).upper() + '.'
        answer_content = answer_match.group(2)
        
        # Format answer letter (bold, blue)
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Process answer content with LaTeX
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_latex(paragraph, answer_content)
        
        return True
    
    # Not a question or answer, process normally
    process_text_with_latex(paragraph, original_text)
    return False

def process_table_cell(cell, text, options):
    """Xử lý content trong cell của table"""
    if not text:
        return 0
    
    latex_count = 0
    
    # Clear cell
    cell.text = ""
    
    # Get or create paragraph
    if len(cell.paragraphs) == 0:
        paragraph = cell.add_paragraph()
    else:
        paragraph = cell.paragraphs[0]
        paragraph.clear()
    
    # Format as Q&A or process LaTeX
    if options.get('format_questions', True):
        if format_question_answer(paragraph, text):
            return latex_count
    
    # Process with LaTeX
    if options.get('convert_equations', True):
        latex_count = process_text_with_latex(paragraph, text)
    else:
        create_formatted_run(paragraph, text)
    
    # Basic formatting
    for run in paragraph.runs:
        if not run.font.name:
            run.font.name = 'Times New Roman'
        if not run.font.size:
            run.font.size = Pt(11)
    
    return latex_count

def format_table_structure(table):
    """Format structure của table"""
    try:
        if not table or len(table.rows) == 0:
            return False
        
        # Center table
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header row
        header_row = table.rows[0]
        for cell in header_row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(12)
                    run.font.name = 'Times New Roman'
            
            # Blue background for header
            try:
                shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                shading = parse_xml(shading_xml)
                cell._tc.get_or_add_tcPr().append(shading)
            except:
                pass
        
        # Format data rows
        for i in range(1, len(table.rows)):
            row = table.rows[i]
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        return True
        
    except Exception:
        return False

def process_document(doc, options):
    """Xử lý document chính"""
    try:
        new_doc = Document()
        
        # Counters
        latex_total = 0
        questions_formatted = 0
        tables_formatted = 0
        
        debug_log = []
        debug_log.append(f"Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
        
        # Process paragraphs
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text:
                new_doc.add_paragraph()
                continue
            
            # Create new paragraph
            new_para = new_doc.add_paragraph()
            
            # Copy alignment
            try:
                if para.alignment:
                    new_para.alignment = para.alignment
            except:
                pass
            
            # Format as Q&A or process normally
            if options.get('format_questions', True):
                if format_question_answer(new_para, text):
                    questions_formatted += 1
                    debug_log.append(f"P{i}: Formatted as Q&A")
                    continue
            
            # Process with LaTeX
            if options.get('convert_equations', True):
                latex_count = process_text_with_latex(new_para, text)
                latex_total += latex_count
                if latex_count > 0:
                    debug_log.append(f"P{i}: Processed {latex_count} LaTeX expressions")
            else:
                create_formatted_run(new_para, text)
        
        # Process tables
        for i, table in enumerate(doc.tables):
            try:
                if not table.rows:
                    debug_log.append(f"Table {i}: Empty, skipped")
                    continue
                
                num_rows = len(table.rows)
                num_cols = len(table.rows[0].cells) if table.rows else 0
                
                debug_log.append(f"Table {i}: {num_rows}x{num_cols}")
                
                # Create new table
                new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
                
                table_latex = 0
                
                # Copy content
                for r in range(num_rows):
                    row = table.rows[r]
                    for c in range(min(len(row.cells), num_cols)):
                        try:
                            original_cell = row.cells[c]
                            new_cell = new_table.rows[r].cells[c]
                            
                            cell_text = original_cell.text.strip()
                            if cell_text:
                                latex_count = process_table_cell(new_cell, cell_text, options)
                                table_latex += latex_count
                        
                        except Exception as e:
                            debug_log.append(f"Error in cell [{r}][{c}]: {str(e)}")
                
                # Format table structure
                if options.get('format_tables', True):
                    if format_table_structure(new_table):
                        tables_formatted += 1
                        debug_log.append(f"Table {i}: Formatted successfully, {table_latex} LaTeX processed")
                
                latex_total += table_latex
                
            except Exception as e:
                debug_log.append(f"Error in table {i}: {str(e)}")
        
        return new_doc, {
            'latex_total': latex_total,
            'questions_formatted': questions_formatted,
            'tables_formatted': tables_formatted,
            'total_paragraphs': len(doc.paragraphs),
            'total_tables': len(doc.tables),
            'debug_log': debug_log
        }
        
    except Exception as e:
        error_doc = Document()
        error_doc.add_paragraph(f"Lỗi xử lý: {str(e)}")
        
        return error_doc, {
            'latex_total': 0,
            'questions_formatted': 0,
            'tables_formatted': 0,
            'total_paragraphs': 0,
            'total_tables': 0,
            'debug_log': [f"ERROR: {str(e)}"]
        }

def main():
    # Header
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor - v7.0</h1>
        <p style="color: white; margin: 10px 0 0 0;">Xử lý LaTeX, Format Q&A và Bảng - Phiên bản ổn định</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar options
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Chuyển đổi LaTeX", value=True,
                                       help="$x^2$ → x², $\\pi$ → π")
        
        format_questions = st.checkbox("📝 Format Q&A", value=True,
                                      help="Câu 1., A., B., C., D.")
        
        format_tables = st.checkbox("📊 Format bảng", value=True,
                                   help="Header màu xanh, căn giữa")
        
        show_debug = st.checkbox("🔍 Debug info", value=False)
        
        st.markdown("---")
        st.markdown("### 📝 LaTeX hỗ trợ")
        st.code("""
$x^2$ → x²
$\\pi$ → π  
$\\pm$ → ±
$[6;8]$ → [6;8]
$\\frac{a}{b}$ → (a)/(b)
$\\sqrt{x}$ → √(x)
        """)
    
    # Main content
    st.markdown("### 📁 Tải file")
    
    uploaded_file = st.file_uploader(
        "Chọn file Word (.docx)",
        type=["docx"],
        help="File có LaTeX, câu hỏi, bảng"
    )
    
    if uploaded_file:
        st.success(f"✅ File: {uploaded_file.name} ({uploaded_file.size/1024:.1f} KB)")
        
        # Preview
        with st.expander("👀 Preview"):
            try:
                doc = Document(uploaded_file)
                st.info(f"📄 {len(doc.paragraphs)} paragraphs | 📊 {len(doc.tables)} tables")
                
                # Show first few paragraphs
                for i, para in enumerate(doc.paragraphs[:3]):
                    if para.text.strip():
                        preview = para.text[:60] + "..." if len(para.text) > 60 else para.text
                        st.text(f"{i+1}. {preview}")
                
                # Show table info
                if doc.tables:
                    st.text(f"First table: {len(doc.tables[0].rows)} rows")
                    
            except Exception as e:
                st.warning(f"Preview error: {e}")
        
        # Process button
        if st.button("🚀 Xử lý file", type="primary", use_container_width=True):
            options = {
                'convert_equations': convert_equations,
                'format_questions': format_questions,
                'format_tables': format_tables
            }
            
            with st.spinner("⏳ Đang xử lý..."):
                try:
                    # Process
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document(doc, options)
                    
                    # Save to buffer
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    # Generate filename
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    name_base = uploaded_file.name.rsplit('.', 1)[0]
                    new_filename = f"{name_base}_processed_{timestamp}.docx"
                    
                    # Show results
                    st.markdown("### 📊 Kết quả")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("🔢 LaTeX", stats['latex_total'])
                    with col2:
                        st.metric("📝 Q&A", stats['questions_formatted'])
                    with col3:
                        st.metric("📊 Tables", stats['tables_formatted'])
                    
                    # Debug info
                    if show_debug:
                        with st.expander("🔍 Debug Log"):
                            for line in stats['debug_log']:
                                st.text(line)
                    
                    # Download
                    st.markdown("### 📥 Tải file")
                    st.download_button(
                        label="⬇️ Tải file đã xử lý",
                        data=buffer.getvalue(),
                        file_name=new_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                    if stats['latex_total'] > 0 or stats['questions_formatted'] > 0:
                        st.success("🎉 Xử lý thành công!")
                    else:
                        st.warning("⚠️ Không tìm thấy LaTeX hoặc Q&A")
                        
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666;">
        <p>📚 LaTeX Word Processor v7.0 | Stable & Reliable</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
