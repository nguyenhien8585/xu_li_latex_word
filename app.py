import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="LaTeX Word Processor",
    page_icon="📚",
    layout="wide"
)

def convert_latex_symbols(text):
    """Chuyển đổi ký hiệu LaTeX thành Unicode"""
    if not text:
        return text
    
    result = text
    
    # Xử lý superscript có dấu ngoặc: A^{prime} -> A'
    while True:
        pos = result.find('^{')
        if pos == -1:
            break
        
        # Tìm ký tự trước ^{
        base_start = pos - 1
        while base_start >= 0 and result[base_start].isalnum():
            base_start -= 1
        base_start += 1
        base = result[base_start:pos]
        
        # Tìm nội dung trong {}
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
            if content == 'prime':
                replacement = base + "'"
            else:
                replacement = base + "^(" + content + ")"
            result = result[:base_start] + replacement + result[brace_end:]
        else:
            break
    
    # Các ký hiệu toán học cơ bản
    symbols = {
        '\\pi': 'π',
        '\\alpha': 'α',
        '\\beta': 'β',
        '\\gamma': 'γ',
        '\\delta': 'δ',
        '\\theta': 'θ',
        '\\lambda': 'λ',
        '\\mu': 'μ',
        '\\sigma': 'σ',
        '\\phi': 'φ',
        '\\omega': 'ω',
        '\\in': '∈',
        '\\subset': '⊂',
        '\\cup': '∪',
        '\\cap': '∩',
        '\\emptyset': '∅',
        '\\leq': '≤',
        '\\geq': '≥',
        '\\neq': '≠',
        '\\approx': '≈',
        '\\infty': '∞',
        '\\pm': '±',
        '\\mp': '∓',
        '\\mathbb{R}': 'ℝ',
        '\\mathbb{Z}': 'ℤ',
        '\\mathbb{Q}': 'ℚ',
        '\\mathbb{N}': 'ℕ',
        '\\mathbb{C}': 'ℂ'
    }
    
    for latex_sym, unicode_sym in symbols.items():
        result = result.replace(latex_sym, unicode_sym)
    
    # Xử lý left, right parentheses
    result = result.replace('\\left(', '(')
    result = result.replace('\\right)', ')')
    result = result.replace('\\left[', '[')
    result = result.replace('\\right]', ']')
    result = result.replace('\\left{', '{')
    result = result.replace('\\right}', '}')
    
    # Xử lý frac
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        # Tìm tử số
        num_start = start + 6
        brace_count = 1
        num_end = num_start
        
        while num_end < len(result) and brace_count > 0:
            if result[num_end] == '{':
                brace_count += 1
            elif result[num_end] == '}':
                brace_count -= 1
            num_end += 1
        
        if brace_count > 0:
            break
            
        numerator = result[num_start:num_end-1]
        
        # Tìm mẫu số
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
            
            if brace_count == 0:
                denominator = result[denom_start:denom_end-1]
                fraction = f"({numerator})/({denominator})"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        else:
            break
    
    # Xử lý sqrt
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
            
        content_start = start + 6
        brace_count = 1
        content_end = content_start
        
        while content_end < len(result) and brace_count > 0:
            if result[content_end] == '{':
                brace_count += 1
            elif result[content_end] == '}':
                brace_count -= 1
            content_end += 1
        
        if brace_count == 0:
            content = result[content_start:content_end-1]
            sqrt_result = f"√({content})"
            result = result[:start] + sqrt_result + result[content_end:]
        else:
            break
    
    # Xử lý superscript đơn giản: x^2 -> x²
    superscript_digits = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    
    i = 0
    new_result = ""
    while i < len(result):
        if (i < len(result) - 2 and 
            result[i+1] == '^' and 
            result[i+2].isdigit() and
            result[i].isalnum()):
            base = result[i]
            power = result[i+2]
            new_result += base + superscript_digits.get(power, power)
            i += 3
        else:
            new_result += result[i]
            i += 1
    
    result = new_result
    
    # Loại bỏ backslash còn lại
    result = result.replace('\\', '')
    
    return result

def find_math_expressions(text):
    """Tìm các biểu thức toán học trong text"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Kiểm tra nếu là ${...}$
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
                    latex_content = text[content_start:i-1]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip()))
                    i += 1
                else:
                    i += 1
            else:
                # Format $...$ thông thường
                content_start = i
                
                while i < len(text) and text[i] != '$':
                    i += 1
                
                if i < len(text):
                    end = i + 1
                    latex_content = text[content_start:i]
                    if latex_content.strip():
                        expressions.append((start, end, latex_content.strip()))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_math_run(paragraph, latex_text):
    """Tạo run cho công thức toán học"""
    unicode_text = convert_latex_symbols(latex_text)
    run = paragraph.add_run(unicode_text)
    run.italic = True
    run.font.color.rgb = RGBColor(0, 100, 0)
    run.font.size = Pt(12)
    run.font.name = 'Cambria Math'
    return run

def process_text_with_math(paragraph, text):
    """Xử lý text có công thức toán học"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        return 0
    
    last_pos = 0
    math_count = 0
    
    for start, end, latex_content in expressions:
        # Thêm text trước công thức
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'
        
        # Thêm công thức toán học
        create_math_run(paragraph, latex_content)
        math_count += 1
        
        last_pos = end
    
    # Thêm text còn lại
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
    
    return math_count

def format_question_answer(paragraph, text):
    """Format câu hỏi và đáp án"""
    paragraph.clear()
    
    # Kiểm tra pattern câu hỏi
    question_pattern = r'^(Câu\s+\d+[\.:])\s*(.*)'
    match = re.match(question_pattern, text, re.IGNORECASE)
    if match:
        question_part = match.group(1)
        content_part = match.group(2)
        
        # In đậm phần câu hỏi
        run_q = paragraph.add_run(question_part)
        run_q.bold = True
        run_q.font.size = Pt(12)
        run_q.font.name = 'Times New Roman'
        
        # Xử lý phần nội dung
        if content_part:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, content_part)
        
        return True
    
    # Kiểm tra pattern đáp án
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Format chữ cái đáp án
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(0, 112, 192)
        run_l.font.size = Pt(12)
        run_l.font.name = 'Times New Roman'
        
        # Xử lý nội dung đáp án
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content)
        
        return True
    
    # Không phải Q&A, xử lý bình thường
    process_text_with_math(paragraph, text)
    return False

def is_markdown_table(text):
    """Kiểm tra xem text có phải markdown table không"""
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False
    
    pipe_count = 0
    for line in lines:
        if '|' in line:
            pipe_count += 1
    
    return pipe_count >= 2

def parse_table_from_text(text):
    """Parse text thành table data"""
    lines = text.strip().split('\n')
    table_rows = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Bỏ qua dòng separator
        if re.match(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$', line):
            continue
        
        if '|' in line:
            # Loại bỏ | đầu và cuối
            line = line.strip('|')
            cells = [cell.strip() for cell in line.split('|')]
            if cells:
                table_rows.append(cells)
    
    return table_rows

def create_word_table(doc, table_data, options):
    """Tạo Word table từ data"""
    if not table_data:
        return 0
    
    math_count = 0
    max_cols = max(len(row) for row in table_data)
    
    # Tạo table
    word_table = doc.add_table(rows=len(table_data), cols=max_cols)
    
    # Điền dữ liệu
    for r, row_data in enumerate(table_data):
        for c, cell_text in enumerate(row_data):
            if c < max_cols and cell_text:
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Tạo paragraph trong cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Xử lý Q&A hoặc math
                if not format_question_answer(cell_para, cell_text):
                    math_count += process_text_with_math(cell_para, cell_text)
                
                # Format cell
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(11)
    
    # Format table structure
    if options.get('format_tables', True):
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header
        if word_table.rows:
            header_row = word_table.rows[0]
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                
                # Màu nền xanh
                try:
                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                    shading = parse_xml(shading_xml)
                    cell._tc.get_or_add_tcPr().append(shading)
                except:
                    pass
    
    return math_count

def process_document_content(doc, options):
    """Xử lý nội dung document"""
    new_doc = Document()
    
    stats = {
        'math_expressions': 0,
        'questions_formatted': 0,
        'word_tables': 0,
        'markdown_tables': 0,
        'debug_log': []
    }
    
    stats['debug_log'].append(f"Bắt đầu xử lý: {len(doc.paragraphs)} paragraphs")
    
    # Xử lý paragraphs
    i = 0
    while i < len(doc.paragraphs):
        para = doc.paragraphs[i]
        text = para.text.strip()
        
        if not text:
            new_doc.add_paragraph()
            i += 1
            continue
        
        # Kiểm tra markdown table
        if '|' in text and options.get('convert_markdown', True):
            # Collect multiple lines for table
            table_lines = []
            j = i
            
            while j < len(doc.paragraphs):
                current_text = doc.paragraphs[j].text.strip()
                if not current_text:
                    break
                if '|' in current_text:
                    table_lines.append(current_text)
                    j += 1
                else:
                    break
            
            # Check if it's a valid table
            combined_text = '\n'.join(table_lines)
            if is_markdown_table(combined_text):
                table_data = parse_table_from_text(combined_text)
                if table_data:
                    math_count = create_word_table(new_doc, table_data, options)
                    stats['math_expressions'] += math_count
                    stats['markdown_tables'] += 1
                    stats['debug_log'].append(f"Converted markdown table: {len(table_data)} rows")
                
                i = j
                continue
        
        # Xử lý paragraph thông thường
        new_para = new_doc.add_paragraph()
        
        if format_question_answer(new_para, text):
            stats['questions_formatted'] += 1
        else:
            math_count = process_text_with_math(new_para, text)
            stats['math_expressions'] += math_count
        
        i += 1
    
    # Xử lý Word tables
    for table_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        num_rows = len(table.rows)
        num_cols = len(table.rows[0].cells) if table.rows else 0
        
        # Tạo table mới
        new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
        
        table_math = 0
        for r in range(num_rows):
            for c in range(min(len(table.rows[r].cells), num_cols)):
                original_cell = table.rows[r].cells[c]
                new_cell = new_table.rows[r].cells[c]
                
                cell_text = original_cell.text.strip()
                if cell_text:
                    new_cell.text = ""
                    if new_cell.paragraphs:
                        cell_para = new_cell.paragraphs[0]
                        cell_para.clear()
                    else:
                        cell_para = new_cell.add_paragraph()
                    
                    if not format_question_answer(cell_para, cell_text):
                        table_math += process_text_with_math(cell_para, cell_text)
                    
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        stats['math_expressions'] += table_math
        
        # Format table
        if options.get('format_tables', True):
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if new_table.rows:
                header_row = new_table.rows[0]
                for cell in header_row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                    
                    try:
                        shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>'
                        shading = parse_xml(shading_xml)
                        cell._tc.get_or_add_tcPr().append(shading)
                    except:
                        pass
            
            stats['word_tables'] += 1
    
    return new_doc, stats

def main():
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 LaTeX Word Processor v7.0</h1>
        <p style="color: white; margin: 10px 0 0 0;">LaTeX + Q&A + Markdown Tables → Word</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn")
        
        convert_equations = st.checkbox("🔢 Chuyển LaTeX", value=True)
        format_questions = st.checkbox("📝 Format Q&A", value=True)  
        format_tables = st.checkbox("📊 Format bảng", value=True)
        convert_markdown = st.checkbox("📋 Markdown Tables", value=True)
        show_debug = st.checkbox("🔍 Debug", value=False)
        
        st.markdown("### 📝 Ví dụ")
        st.code("""
$A^{prime}$ → A'
$\\left(x\\right)$ → (x)  
${B^{prime}}$ → B'
$x^2$ → x²
$\\frac{a}{b}$ → (a)/(b)

| Col1 | Col2 |
|------|------|
| Data | Data |
        """)
    
    # Main area
    st.markdown("### 📁 Upload File")
    
    uploaded_file = st.file_uploader("Chọn file Word (.docx)", type=["docx"])
    
    if uploaded_file:
        st.success(f"✅ {uploaded_file.name}")
        
        with st.expander("Preview"):
            try:
                doc = Document(uploaded_file)
                st.info(f"📄 {len(doc.paragraphs)} paragraphs | 📊 {len(doc.tables)} tables")
                
                for i, para in enumerate(doc.paragraphs[:5]):
                    if para.text.strip():
                        preview_text = para.text[:80] + "..." if len(para.text) > 80 else para.text
                        st.text(f"{i+1}. {preview_text}")
            except Exception as e:
                st.error(f"Preview error: {e}")
        
        if st.button("🚀 Xử lý", type="primary", use_container_width=True):
            options = {
                'convert_equations': convert_equations,
                'format_questions': format_questions,
                'format_tables': format_tables,
                'convert_markdown': convert_markdown
            }
            
            with st.spinner("Đang xử lý..."):
                try:
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document_content(doc, options)
                    
                    # Save
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_processed_{timestamp}.docx"
                    
                    # Results
                    st.markdown("### 📊 Kết quả")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("🔢 Math", stats['math_expressions'])
                    col2.metric("📝 Q&A", stats['questions_formatted'])
                    col3.metric("📊 Word Tables", stats['word_tables'])
                    col4.metric("📋 Markdown", stats['markdown_tables'])
                    
                    if show_debug:
                        with st.expander("Debug"):
                            for log in stats['debug_log']:
                                st.text(log)
                    
                    # Download
                    st.download_button(
                        "⬇️ Tải file",
                        buffer.getvalue(),
                        filename,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                    total_processed = (stats['math_expressions'] + stats['questions_formatted'] + 
                                     stats['word_tables'] + stats['markdown_tables'])
                    
                    if total_processed > 0:
                        st.success("🎉 Hoàn thành!")
                    else:
                        st.warning("⚠️ Không tìm thấy nội dung cần xử lý")
                
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")

if __name__ == "__main__":
    main()
