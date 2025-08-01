import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.table import Table
import re
from io import BytesIO
import tempfile
import os
from datetime import datetime

# Cấu hình trang
st.set_page_config(
    page_title="Xử lý LaTeX & Format trong Word",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

def create_equation_run(paragraph, latex_expr):
    """Tạo run với formatting đặc biệt cho equation thay vì OMML"""
    try:
        # Approach mới: Format equation như italic, bold, với màu khác biệt
        # Thay vì convert sang OMML phức tạp
        
        # Làm sạch LaTeX expression
        clean_expr = latex_expr.strip()
        
        # Xử lý một số pattern phổ biến
        # \frac{a}{b} -> (a)/(b)
        clean_expr = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'(\1)/(\2)', clean_expr)
        
        # \sqrt{x} -> √(x)
        clean_expr = re.sub(r'\\sqrt\{([^}]+)\}', r'√(\1)', clean_expr)
        
        # x^2 -> x²
        clean_expr = re.sub(r'([a-zA-Z0-9]+)\^([0-9]+)', lambda m: m.group(1) + ''.join('⁰¹²³⁴⁵⁶⁷⁸⁹'[int(d)] for d in m.group(2)), clean_expr)
        
        # \pi -> π
        clean_expr = clean_expr.replace('\\pi', 'π')
        
        # \in -> ∈
        clean_expr = clean_expr.replace('\\in', '∈')
        
        # \mathbb{Z} -> ℤ, \mathbb{R} -> ℝ, etc.
        clean_expr = clean_expr.replace('\\mathbb{Z}', 'ℤ')
        clean_expr = clean_expr.replace('\\mathbb{R}', 'ℝ')
        clean_expr = clean_expr.replace('\\mathbb{Q}', 'ℚ')
        clean_expr = clean_expr.replace('\\mathbb{N}', 'ℕ')
        
        # Loại bỏ các ký tự LaTeX còn lại
        clean_expr = clean_expr.replace('\\', '')
        
        # Tạo run với format đặc biệt
        run = paragraph.add_run(clean_expr)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 100, 0)  # Màu xanh lá đậm
        run.font.size = Pt(12)
        run.font.name = 'Cambria Math'  # Font toán học
        
        return True
        
    except Exception as e:
        return False

def format_question_and_answers(paragraph):
    """Format câu hỏi và đáp án trắc nghiệm"""
    try:
        text = paragraph.text.strip()
        if not text:
            return False
        
        # Backup text gốc
        original_text = text
        
        # Clear paragraph
        paragraph.clear()
        
        # Pattern 1: Câu số. (có thể có **Câu** hoặc không)
        question_patterns = [
            r'^(\*\*Câu\s+\d+[\.\:]?\*\*\.?)\s*(.*)',  # **Câu 1:** hoặc **Câu 1.**
            r'^(Câu\s+\d+[\.\:])\s*(.*)',              # Câu 1: hoặc Câu 1.
        ]
        
        for pattern in question_patterns:
            match = re.match(pattern, text, re.IGNORECASE)
            if match:
                question_part = match.group(1).replace('*', '')  # Loại bỏ **
                content_part = match.group(2)
                
                # Format question part (in đậm)
                run_q = paragraph.add_run(question_part)
                run_q.bold = True
                run_q.font.size = Pt(12)
                run_q.font.name = 'Times New Roman'
                
                # Format content part
                if content_part:
                    run_c = paragraph.add_run(" " + content_part)
                    run_c.font.size = Pt(12)
                    run_c.font.name = 'Times New Roman'
                
                return True
        
        # Pattern 2: Đáp án A., B., C., D.
        answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
        match = re.match(answer_pattern, text)
        if match:
            answer_letter = match.group(1).upper() + '.'
            answer_content = match.group(2)
            
            # Format answer letter (in đậm, màu xanh)
            run_l = paragraph.add_run(answer_letter)
            run_l.bold = True
            run_l.font.color.rgb = RGBColor(0, 112, 192)  # Xanh dương
            run_l.font.size = Pt(12)
            run_l.font.name = 'Times New Roman'
            
            # Format answer content
            if answer_content:
                run_a = paragraph.add_run(" " + answer_content)
                run_a.font.size = Pt(12)
                run_a.font.name = 'Times New Roman'
            
            return True
        
        # Nếu không match pattern nào, khôi phục text gốc
        paragraph.add_run(original_text)
        return False
        
    except Exception as e:
        # Nếu có lỗi, khôi phục text gốc
        paragraph.clear()
        paragraph.add_run(original_text if 'original_text' in locals() else text)
        return False

def format_table_properly(table):
    """Format bảng đúng cách"""
    try:
        if not table or len(table.rows) == 0:
            return False
        
        # Set table alignment
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Format header row (first row)
        header_row = table.rows[0]
        for cell in header_row.cells:
            # Bold and center header
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(12)
                    run.font.name = 'Times New Roman'
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add blue background to header
            try:
                shading = parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="D9E2F3"/>')
                cell._tc.get_or_add_tcPr().append(shading)
            except:
                pass
        
        # Format data rows
        for i in range(1, len(table.rows)):
            row = table.rows[i]
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(11)
                        run.font.name = 'Times New Roman'
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        return True
        
    except Exception as e:
        return False

def find_latex_in_text(text):
    """Tìm tất cả LaTeX expressions trong text"""
    expressions = []
    
    # Pattern tìm $...$ và ${...}$
    patterns = [
        r'\$\{([^}]+)\}\$',  # ${...}$
        r'\$([^$]+)\$',      # $...$
    ]
    
    for pattern in patterns:
        for match in re.finditer(pattern, text):
            start, end = match.span()
            expr = match.group(1)
            
            # Kiểm tra không overlap với expressions đã tìm
            overlap = False
            for existing_start, existing_end, _, _ in expressions:
                if not (end <= existing_start or start >= existing_end):
                    overlap = True
                    break
            
            if not overlap and expr.strip():
                expressions.append((start, end, expr.strip(), 'latex'))
    
    return sorted(expressions, key=lambda x: x[0])

def process_paragraph_with_latex(paragraph, text):
    """Xử lý paragraph có chứa LaTeX"""
    try:
        # Tìm LaTeX expressions
        latex_exprs = find_latex_in_text(text)
        
        if not latex_exprs:
            # Không có LaTeX, chỉ add text thường
            paragraph.add_run(text)
            return 0
        
        # Có LaTeX, xử lý từng phần
        last_end = 0
        processed_count = 0
        
        for start, end, expr, expr_type in latex_exprs:
            # Add text trước LaTeX
            if start > last_end:
                before_text = text[last_end:start]
                if before_text:
                    paragraph.add_run(before_text)
            
            # Add LaTeX expression
            if create_equation_run(paragraph, expr):
                processed_count += 1
            else:
                # Fallback: add as formatted text
                run = paragraph.add_run(f"[{expr}]")
                run.italic = True
                run.bold = True
            
            last_end = end
        
        # Add text còn lại
        if last_end < len(text):
            remaining = text[last_end:]
            if remaining:
                paragraph.add_run(remaining)
        
        return processed_count
        
    except Exception as e:
        # Fallback: add original text
        paragraph.add_run(text)
        return 0

def process_document_comprehensive(doc, options):
    """Xử lý document một cách toàn diện"""
    try:
        new_doc = Document()
        
        # Counters
        latex_processed = 0
        questions_formatted = 0
        tables_formatted = 0
        paragraphs_cleaned = 0
        
        # Debug info
        debug_log = []
        
        debug_log.append(f"Bắt đầu xử lý {len(doc.paragraphs)} paragraphs và {len(doc.tables)} tables")
        
        # Process paragraphs
        for i, para in enumerate(doc.paragraphs):
            try:
                text = para.text.strip()
                if not text:
                    new_doc.add_paragraph()
                    continue
                
                debug_log.append(f"P{i}: {text[:50]}...")
                
                # Tạo paragraph mới
                new_para = new_doc.add_paragraph()
                
                # Copy basic formatting
                try:
                    if para.alignment:
                        new_para.alignment = para.alignment
                except:
                    pass
                
                # Clean text if requested
                if options.get('clean_text', False):
                    original_text = text
                    # Simple cleaning
                    text = ' '.join(text.split())
                    if text != original_text:
                        paragraphs_cleaned += 1
                
                # Try to format as question/answer first
                formatted_as_qa = False
                if options.get('format_questions', True):
                    # Set text vào paragraph trước để test
                    new_para.add_run(text)
                    if format_question_and_answers(new_para):
                        questions_formatted += 1
                        formatted_as_qa = True
                        debug_log.append(f"  -> Formatted as Q&A")
                
                # If not formatted as Q&A, process LaTeX
                if not formatted_as_qa:
                    new_para.clear()
                    if options.get('convert_equations', True):
                        processed = process_paragraph_with_latex(new_para, text)
                        latex_processed += processed
                        if processed > 0:
                            debug_log.append(f"  -> Processed {processed} LaTeX expressions")
                    else:
                        new_para.add_run(text)
                
                # Apply basic formatting
                if options.get('fix_formatting', True):
                    for run in new_para.runs:
                        if not run.font.name:
                            run.font.name = 'Times New Roman'
                        if not run.font.size:
                            run.font.size = Pt(12)
                
            except Exception as e:
                debug_log.append(f"  -> ERROR in paragraph {i}: {str(e)}")
                new_doc.add_paragraph(text if 'text' in locals() else "")
        
        # Process tables
        debug_log.append(f"Xử lý {len(doc.tables)} tables...")
        for i, table in enumerate(doc.tables):
            try:
                debug_log.append(f"Table {i}: {len(table.rows)}x{len(table.rows[0].cells) if table.rows else 0}")
                
                # Copy table structure
                new_table = new_doc.add_table(rows=len(table.rows), cols=len(table.rows[0].cells) if table.rows else 0)
                
                # Copy content
                for r in range(len(table.rows)):
                    for c in range(len(table.rows[r].cells)):
                        try:
                            original_cell = table.rows[r].cells[c]
                            new_cell = new_table.rows[r].cells[c]
                            new_cell.text = original_cell.text
                        except:
                            pass
                
                # Format table
                if options.get('format_tables', True):
                    if format_table_properly(new_table):
                        tables_formatted += 1
                        debug_log.append(f"  -> Table {i} formatted successfully")
                    else:
                        debug_log.append(f"  -> Table {i} formatting failed")
                
            except Exception as e:
                debug_log.append(f"  -> ERROR in table {i}: {str(e)}")
        
        return new_doc, {
            'processed_count': latex_processed,
            'error_count': 0,  # We don't track errors in this version
            'cleaned_paragraphs': paragraphs_cleaned,
            'formatted_questions': questions_formatted,
            'formatted_tables': tables_formatted,
            'total_paragraphs': len(doc.paragraphs),
            'total_tables': len(doc.tables),
            'debug_info': debug_log
        }
        
    except Exception as e:
        # Create error document
        error_doc = Document()
        error_doc.add_paragraph(f"Lỗi xử lý document: {str(e)}")
        
        return error_doc, {
            'processed_count': 0,
            'error_count': 1,
            'cleaned_paragraphs': 0,
            'formatted_questions': 0,
            'formatted_tables': 0,
            'total_paragraphs': 0,
            'total_tables': 0,
            'debug_info': [f"CRITICAL ERROR: {str(e)}"]
        }

def main():
    # Header
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0;">📚 Xử lý LaTeX & Format trong Word - v6.0</h1>
        <p style="color: white; margin: 10px 0 0 0;">Chuyển đổi LaTeX, Format câu hỏi trắc nghiệm & Xử lý bảng</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Tùy chọn xử lý")
        
        convert_equations = st.checkbox("🔢 Chuyển đổi công thức LaTeX", value=True,
                                       help="Chuyển $...$, ${...}$ thành ký hiệu toán học Unicode")
        
        format_questions = st.checkbox("📝 Format câu hỏi trắc nghiệm", value=True,
                                      help="In đậm câu hỏi, tô màu xanh các đáp án A, B, C, D")
        
        format_tables = st.checkbox("📊 Format bảng", value=True,
                                   help="Căn chỉnh bảng, in đậm header, tô màu nền")
        
        clean_text = st.checkbox("🧹 Làm sạch văn bản", value=True,
                                help="Chuẩn hóa khoảng trắng")
        
        fix_formatting = st.checkbox("🎨 Sửa định dạng cơ bản", value=True,
                                    help="Chuẩn hóa font chữ Times New Roman 12pt")
        
        show_debug = st.checkbox("🔍 Hiển thị debug info", value=False,
                                help="Xem chi tiết quá trình xử lý")
        
        st.markdown("---")
        st.markdown("### 📖 Cách sử dụng")
        st.markdown("""
        1. Tải file Word (.docx)
        2. Chọn các tùy chọn xử lý
        3. Nhấn "Xử lý file"
        4. Tải file đã được format
        
        **LaTeX hỗ trợ:**
        - `$x^2$` → x²
        - `$\\frac{a}{b}$` → (a)/(b)
        - `$\\sqrt{x}$` → √(x)
        - `$\\pi$` → π
        """)
    
    # Main content
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.markdown("### 📁 Tải file Word")
        
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)",
            type=["docx"],
            help="File chứa LaTeX, câu hỏi trắc nghiệm, bảng"
        )
        
        if uploaded_file:
            st.success(f"✅ Đã tải: **{uploaded_file.name}** ({uploaded_file.size/1024:.1f} KB)")
            
            # Preview
            with st.expander("👀 Preview nội dung"):
                try:
                    doc = Document(uploaded_file)
                    st.info(f"📄 {len(doc.paragraphs)} đoạn văn | 📊 {len(doc.tables)} bảng")
                    
                    for i, para in enumerate(doc.paragraphs[:3]):
                        if para.text.strip():
                            preview = para.text[:80] + "..." if len(para.text) > 80 else para.text
                            st.text(f"{i+1}. {preview}")
                            
                except Exception as e:
                    st.warning(f"Lỗi preview: {e}")
            
            # Process button
            if st.button("🚀 Xử lý file", type="primary", use_container_width=True):
                options = {
                    'convert_equations': convert_equations,
                    'format_questions': format_questions,
                    'format_tables': format_tables,
                    'clean_text': clean_text,
                    'fix_formatting': fix_formatting
                }
                
                with st.spinner("⏳ Đang xử lý..."):
                    try:
                        # Load and process
                        doc = Document(uploaded_file)
                        new_doc, stats = process_document_comprehensive(doc, options)
                        
                        # Save to buffer
                        buffer = BytesIO()
                        new_doc.save(buffer)
                        buffer.seek(0)
                        
                        # Generate filename
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        name_without_ext = uploaded_file.name.rsplit('.', 1)[0]
                        new_filename = f"{name_without_ext}_processed_{timestamp}.docx"
                        
                        # Results
                        st.markdown("### 📊 Kết quả")
                        
                        cols = st.columns(4)
                        with cols[0]:
                            st.metric("🔢 LaTeX", stats['processed_count'])
                        with cols[1]:
                            st.metric("📝 Câu hỏi", stats['formatted_questions'])
                        with cols[2]:
                            st.metric("📊 Bảng", stats['formatted_tables'])
                        with cols[3]:
                            st.metric("🧹 Làm sạch", stats['cleaned_paragraphs'])
                        
                        # Debug info
                        if show_debug and stats.get('debug_info'):
                            with st.expander("🔍 Debug Log"):
                                for line in stats['debug_info']:
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
                        
                        # Success message
                        if stats['processed_count'] > 0 or stats['formatted_questions'] > 0 or stats['formatted_tables'] > 0:
                            st.success("🎉 Xử lý thành công!")
                        else:
                            st.warning("⚠️ Không tìm thấy nội dung cần xử lý")
                            
                    except Exception as e:
                        st.error(f"❌ Lỗi: {str(e)}")
    
    with col2:
        st.markdown("### 💡 Tips")
        
        with st.expander("🧪 Test LaTeX"):
            st.code("""
Thử với nội dung này:

Câu 1. Phương trình $x^2 = 4$ có nghiệm là
A. $x = \\pm 2$
B. $x = \\frac{4}{2}$  
C. $x = \\sqrt{4}$
D. Tất cả đều đúng
            """)
        
        with st.expander("🎯 Ví dụ format"):
            st.markdown("""
            **Trước:**
            ```
            Câu 1. Giải $x^2=4$
            A. x=2
            B. x=-2
            ```
            
            **Sau:**
            - **Câu 1.** Giải x² = 4
            - **A.** x = 2 (màu xanh)
            - **B.** x = -2 (màu xanh)
            """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 10px;">
        <p>📚 LaTeX & Word Processor v6.0 | 🔧 Simplified & Reliable</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
