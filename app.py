import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file, use_word_equation=False, math_font="Cambria Math"):
    """
    Xử lý file Word đã tải lên - Phiên bản nâng cao với tùy chọn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            progress_math = st.progress(0)
            
            total_paragraphs = len(doc.paragraphs)
            for i, paragraph in enumerate(doc.paragraphs):
                try:
                    if process_paragraph_math_advanced(paragraph, use_word_equation, math_font):
                        math_count += 1
                    
                    # Cập nhật progress
                    progress_math.progress(min(int((i + 1) / total_paragraphs * 100), 100))
                    
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi ở đoạn {i+1}: {str(e)}")
                    continue
            
            progress_math.empty()
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
    # Sidebar với hướng dẫn
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        **Bước 1:** Tải lên file Word (.docx)
        
        **Bước 2:** Ứng dụng sẽ tự động:
        - Chuyển `$công thứcimport streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
 → Unicode Math
        - Chuyển `${công thức}import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
 → Unicode Math  
        - Chuyển bảng Markdown → Bảng Word
        
        **Bước 3:** Tải xuống file đã xử lý
        
        ---
        
        **✨ Công thức LaTeX được hỗ trợ:**
        - `$x^2 + y^2 = z^2import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
 → x² + y² = z²
        - `$\\alpha + \\betaimport streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
 → α + β
        - `$\\frac{a}{b}import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
 → (a)/(b)
        - `$\\sqrt{x}import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
 → √(x)
        - `$\\sum, \\int, \\inftyimport streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
 → ∑, ∫, ∞
        - `$\\leq, \\geq, \\neqimport streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
 → ≤, ≥, ≠
        
        **📊 Bảng Markdown:**
        ```
        | Cột 1 | Cột 2 |
        |-------|-------|
        | A     | B     |
        | C     | D     |
        ```
        
        ---
        
        **🔧 Khắc phục sự cố:**
        
        **File không mở được?**
        - Đảm bảo file gốc .docx hợp lệ
        - Đóng file trong Word trước khi upload
        - Thử với file Word khác để test
        
        **Công thức không chuyển?**
        - Đảm bảo format đúng: `$...import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    

        - Kiểm tra ký tự đặc biệt
        - Sử dụng dấu \\\\ cho LaTeX
        
        **Bảng không hiển thị?**
        - Kiểm tra format Markdown đúng
        - Cần có dòng separator `|---|`
        - Ít nhất 2 dòng dữ liệu
        """)
        
        st.markdown("---")
        st.markdown("**🚀 Phiên bản:** 3.0 (Advanced)")
        st.markdown("**✨ Tính năng mới:** Word Equation thực sự + Unicode safe mode")
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📤 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)",
            type=['docx'],
            help="Tải lên file Word chứa công thức LaTeX, [MATH: ...] và/hoặc bảng Markdown"
        )
        
        # Thêm tùy chọn xử lý
        st.subheader("⚙️ Tùy chọn xử lý")
        col_opt1, col_opt2 = st.columns(2)
        
        with col_opt1:
            math_format = st.radio(
                "🧮 Định dạng công thức output:",
                options=["Unicode Text (An toàn)", "Word Equation (Thực sự)"],
                help="Unicode: An toàn, không corrupt file. Word Equation: Thực sự nhưng có thể rủi ro"
            )
        
        with col_opt2:
            font_option = st.selectbox(
                "🔤 Font cho công thức:",
                options=["Cambria Math", "Times New Roman", "Arial", "Calibri"],
                help="Font hiển thị cho các ký hiệu toán học"
            )
        
        if uploaded_file is not None:
            # Validate file trước khi xử lý
            is_valid, message = validate_docx_file(uploaded_file)
            
            if not is_valid:
                st.error(f"❌ {message}")
                st.info("💡 Vui lòng tải lên file Word (.docx) hợp lệ")
                return
            
            # Hiển thị thông tin file
            file_details = {
                "Tên file": uploaded_file.name,
                "Kích thước": f"{uploaded_file.size:,} bytes ({uploaded_file.size/1024/1024:.1f} MB)",
                "Loại file": uploaded_file.type,
                "Trạng thái": "✅ Hợp lệ"
            }
            
            st.subheader("📄 Thông tin file")
            col_info1, col_info2 = st.columns(2)
            
            with col_info1:
                st.write(f"**{list(file_details.keys())[0]}:** {list(file_details.values())[0]}")
                st.write(f"**{list(file_details.keys())[1]}:** {list(file_details.values())[1]}")
            
            with col_info2:
                st.write(f"**{list(file_details.keys())[2]}:** {list(file_details.values())[2]}")
                st.write(f"**{list(file_details.keys())[3]}:** {list(file_details.values())[3]}")
            
            # Xử lý file
            if st.button("🔄 Xử lý file", type="primary", use_container_width=True):
                
                # Progress bar
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    status_text.text("🔍 Đang phân tích file...")
                    progress_bar.progress(20)
                    
                    # Xác định tùy chọn xử lý
                    use_word_eq = (math_format == "Word Equation (Thực sự)")
                    
                    doc, math_count, table_count = process_word_document(
                        uploaded_file, 
                        use_word_equation=use_word_eq,
                        math_font=font_option
                    )
                    
                    if doc is not None:
                        progress_bar.progress(80)
                        status_text.text("💾 Đang tạo file mới...")
                        
                        # Tạo file download
                        processed_file = save_document(doc, uploaded_file.name)
                        
                        if processed_file is not None:
                            progress_bar.progress(100)
                            status_text.text("✅ Hoàn thành!")
                            
                            st.success("🎉 Xử lý file thành công!")
                            
                            # Hiển thị kết quả
                            col_result1, col_result2, col_result3 = st.columns(3)
                            with col_result1:
                                st.metric("Công thức đã chuyển đổi", math_count)
                            with col_result2:
                                st.metric("Bảng đã chuyển đổi", table_count)
                            with col_result3:
                                st.metric("Kích thước file mới", f"{len(processed_file)/1024:.1f} KB")
                            
                            # Tạo tên file mới
                            base_name = uploaded_file.name.rsplit('.', 1)[0]
                            new_filename = f"{base_name}_processed.docx"
                            
                            # Nút download
                            st.download_button(
                                label="📥 Tải xuống file đã xử lý",
                                data=processed_file,
                                file_name=new_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="secondary",
                                use_container_width=True
                            )
                            
                            if math_count > 0 or table_count > 0:
                                st.info("💡 File đã được xử lý thành công! Các công thức LaTeX đã được chuyển thành Unicode và bảng Markdown đã thành bảng Word.")
                            else:
                                st.info("ℹ️ Không tìm thấy công thức LaTeX hoặc bảng Markdown nào trong file.")
                                
                            # Hiển thị preview (nếu có thay đổi)
                            if math_count > 0 or table_count > 0:
                                with st.expander("🔍 Xem trước các thay đổi"):
                                    if math_count > 0:
                                        st.write("**Công thức đã chuyển đổi:**")
                                        st.write("- `$x^2import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
    # Sidebar với hướng dẫn
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        **Bước 1:** Tải lên file Word (.docx)
        
        **Bước 2:** Ứng dụng sẽ tự động:
        - Chuyển `$công thức$` → Equation Word
        - Chuyển `${công thức}$` → Equation Word  
        - Chuyển bảng Markdown → Bảng Word
        
        **Bước 3:** Tải xuống file đã xử lý
        
        ---
        
        **Ví dụ công thức LaTeX:**
        - `$x^2 + y^2 = z^2$`
        - `$\\frac{a}{b} + \\sqrt{c}$`
        - `$\\alpha + \\beta = \\gamma$`
        
        **Ví dụ bảng Markdown:**
        ```
        | Cột 1 | Cột 2 |
        |-------|-------|
        | A     | B     |
        | C     | D     |
        ```
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📤 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)",
            type=['docx'],
            help="Tải lên file Word chứa công thức LaTeX và/hoặc bảng Markdown"
        )
        
 → x²")
                                        st.write("- `$\\alpha + \\betaimport streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
    # Sidebar với hướng dẫn
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        **Bước 1:** Tải lên file Word (.docx)
        
        **Bước 2:** Ứng dụng sẽ tự động:
        - Chuyển `$công thức$` → Equation Word
        - Chuyển `${công thức}$` → Equation Word  
        - Chuyển bảng Markdown → Bảng Word
        
        **Bước 3:** Tải xuống file đã xử lý
        
        ---
        
        **Ví dụ công thức LaTeX:**
        - `$x^2 + y^2 = z^2$`
        - `$\\frac{a}{b} + \\sqrt{c}$`
        - `$\\alpha + \\beta = \\gamma$`
        
        **Ví dụ bảng Markdown:**
        ```
        | Cột 1 | Cột 2 |
        |-------|-------|
        | A     | B     |
        | C     | D     |
        ```
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📤 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)",
            type=['docx'],
            help="Tải lên file Word chứa công thức LaTeX và/hoặc bảng Markdown"
        )
        
 → α + β") 
                                        st.write("- `$\\frac{a}{b}import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
    # Sidebar với hướng dẫn
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        **Bước 1:** Tải lên file Word (.docx)
        
        **Bước 2:** Ứng dụng sẽ tự động:
        - Chuyển `$công thức$` → Equation Word
        - Chuyển `${công thức}$` → Equation Word  
        - Chuyển bảng Markdown → Bảng Word
        
        **Bước 3:** Tải xuống file đã xử lý
        
        ---
        
        **Ví dụ công thức LaTeX:**
        - `$x^2 + y^2 = z^2$`
        - `$\\frac{a}{b} + \\sqrt{c}$`
        - `$\\alpha + \\beta = \\gamma$`
        
        **Ví dụ bảng Markdown:**
        ```
        | Cột 1 | Cột 2 |
        |-------|-------|
        | A     | B     |
        | C     | D     |
        ```
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📤 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)",
            type=['docx'],
            help="Tải lên file Word chứa công thức LaTeX và/hoặc bảng Markdown"
        )
        
 → (a)/(b)")
                                        st.write("- `$\\sqrt{x}import streamlit as st
import re
import io
import tempfile
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

def convert_latex_to_unicode(latex_text):
    """
    Chuyển đổi công thức LaTeX thành Unicode text (thay vì Word equation)
    Cách này an toàn hơn và không gây corrupt file
    """
    # Loại bỏ dấu $ ở đầu và cuối
    latex_text = latex_text.strip('${}')
    
    # Các chuyển đổi cơ bản từ LaTeX sang Unicode
    conversions = {
        # Greek letters
        r'\\alpha': 'α',
        r'\\beta': 'β', 
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\sigma': 'σ',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Pi': 'Π',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        
        # Math symbols
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\infty': '∞',
        r'\\leq': '≤',
        r'\\geq': '≥',
        r'\\neq': '≠',
        r'\\pm': '±',
        r'\\mp': '∓',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\cdot': '·',
        r'\\subset': '⊂',
        r'\\supset': '⊃',
        r'\\in': '∈',
        r'\\notin': '∉',
        r'\\cup': '∪',
        r'\\cap': '∩',
        r'\\emptyset': '∅',
        r'\\partial': '∂',
        r'\\nabla': '∇',
        r'\\forall': '∀',
        r'\\exists': '∃',
        r'\\therefore': '∴',
        r'\\because': '∵',
        r'\\approx': '≈',
        r'\\sim': '∼',
        r'\\equiv': '≡',
        r'\\prec': '≺',
        r'\\succ': '≻',
        
        # Arrows
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\leftrightarrow': '↔',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
        r'\\Leftrightarrow': '⇔',
        
        # Superscript numbers (simple cases)
        r'\^0': '⁰',
        r'\^1': '¹',
        r'\^2': '²', 
        r'\^3': '³',
        r'\^4': '⁴',
        r'\^5': '⁵',
        r'\^6': '⁶',
        r'\^7': '⁷',
        r'\^8': '⁸',
        r'\^9': '⁹',
        
        # Subscript numbers (simple cases)
        r'_0': '₀',
        r'_1': '₁',
        r'_2': '₂',
        r'_3': '₃',
        r'_4': '₄',
        r'_5': '₅',
        r'_6': '₆',
        r'_7': '₇',
        r'_8': '₈',
        r'_9': '₉',
        
        # Fractions (simple ones)
        r'\\frac\{1\}\{2\}': '½',
        r'\\frac\{1\}\{3\}': '⅓',
        r'\\frac\{2\}\{3\}': '⅔',
        r'\\frac\{1\}\{4\}': '¼',
        r'\\frac\{3\}\{4\}': '¾',
        r'\\frac\{1\}\{5\}': '⅕',
        r'\\frac\{1\}\{6\}': '⅙',
        r'\\frac\{1\}\{8\}': '⅛',
        
        # Root symbols
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        
        # Complex fractions (general case)
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        
        # Exponents with braces
        r'\^\{([^}]+)\}': r'^(\1)',
        
        # Subscripts with braces  
        r'_\{([^}]+)\}': r'_(\1)',
    }
    
    result = latex_text
    for pattern, replacement in conversions.items():
        result = re.sub(pattern, replacement, result)
    
    return result

def process_paragraph_math(paragraph):
    """
    Xử lý công thức toán học trong một đoạn văn
    Sử dụng Unicode thay vì Word equation để tránh corrupt file
    """
    # Tìm các công thức trong format $...$ hoặc ${...}$
    math_pattern = r'\$\{?([^$]+)\}?\

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    try:
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        if len(lines) < 2:
            return None
        
        # Lấy header - loại bỏ | ở đầu và cuối
        header_line = lines[0].strip('|').strip()
        headers = [cell.strip() for cell in header_line.split('|')]
        
        # Kiểm tra dòng separator
        if len(lines) < 3:
            return [headers] if headers else None
        
        # Lấy dữ liệu
        data = [headers]
        for line in lines[2:]:
            if line.strip() and '|' in line:
                # Loại bỏ | ở đầu và cuối
                row_line = line.strip('|').strip()
                row = [cell.strip() for cell in row_line.split('|')]
                
                if row and any(cell.strip() for cell in row):  # Có ít nhất 1 cell không rỗng
                    # Đảm bảo số cột bằng header
                    while len(row) < len(headers):
                        row.append('')
                    row = row[:len(headers)]
                    data.append(row)
        
        return data if len(data) > 1 else None
        
    except Exception as e:
        st.warning(f"Lỗi phân tích bảng: {str(e)}")
        return None

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word - Phiên bản an toàn
    """
    try:
        # Pattern đơn giản hơn để tìm bảng Markdown
        table_pattern = r'(\|[^\n]*\|[\s]*\n\|[-\s|:]*\|[\s]*\n(?:\|[^\n]*\|[\s]*\n?)*)'
        
        tables_processed = 0
        
        # Xử lý từng paragraph
        paragraphs_to_process = []
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text
            if '|' in text and '-' in text:  # Có khả năng chứa bảng
                matches = list(re.finditer(table_pattern, text, re.MULTILINE))
                if matches:
                    paragraphs_to_process.append((i, paragraph, matches))
        
        # Xử lý các paragraph có bảng
        for i, paragraph, matches in reversed(paragraphs_to_process):
            for match in reversed(matches):  # Xử lý từ cuối lên để không ảnh hưởng index
                table_text = match.group(1)
                table_data = parse_markdown_table(table_text)
                
                if table_data and len(table_data) > 1:
                    try:
                        # Tạo bảng mới
                        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                        
                        # Sử dụng style có sẵn
                        try:
                            table.style = 'Table Grid'
                        except:
                            pass  # Nếu style không có, bỏ qua
                        
                        # Điền dữ liệu
                        for row_idx, row_data in enumerate(table_data):
                            if row_idx < len(table.rows):
                                row = table.rows[row_idx]
                                for col_idx, cell_data in enumerate(row_data):
                                    if col_idx < len(row.cells):
                                        cell = row.cells[col_idx]
                                        cell.text = str(cell_data)
                                        
                                        # Định dạng header
                                        if row_idx == 0 and cell.paragraphs:
                                            for run in cell.paragraphs[0].runs:
                                                run.bold = True
                        
                        # Thay thế text bảng bằng placeholder
                        original_text = paragraph.text
                        new_text = (original_text[:match.start()] + 
                                  f"\n[TABLE_{tables_processed}]\n" + 
                                  original_text[match.end():])
                        paragraph.text = new_text
                        
                        # Di chuyển bảng đến sau paragraph
                        paragraph._element.addnext(table._element)
                        
                        tables_processed += 1
                        
                    except Exception as e:
                        st.warning(f"Không thể tạo bảng: {str(e)}")
                        continue
        
        # Loại bỏ các placeholder
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Loại bỏ tất cả placeholder
            new_text = re.sub(r'\[TABLE_\d+\]', '', text)
            if new_text != text:
                paragraph.text = new_text.strip()
        
        return tables_processed
        
    except Exception as e:
        st.error(f"Lỗi xử lý bảng: {str(e)}")
        return 0

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên - Phiên bản an toàn
    """
    try:
        # Tạo temporary file để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Đọc file Word
            doc = Document(tmp_file_path)
            
            # Đếm số thay đổi
            math_count = 0
            table_count = 0
            
            # Xử lý công thức toán học
            st.info("🔄 Đang xử lý công thức toán học...")
            for paragraph in doc.paragraphs:
                try:
                    if process_paragraph_math(paragraph):
                        math_count += 1
                except Exception as e:
                    st.warning(f"Bỏ qua công thức lỗi: {str(e)}")
                    continue
            
            # Xử lý bảng Markdown
            st.info("📊 Đang xử lý bảng Markdown...")
            try:
                table_count = replace_markdown_tables_with_word_tables(doc)
            except Exception as e:
                st.warning(f"Lỗi xử lý bảng: {str(e)}")
                table_count = 0
            
            return doc, math_count, table_count
            
        finally:
            # Xóa temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        st.error("Vui lòng kiểm tra:")
        st.error("- File có phải là .docx hợp lệ không?")
        st.error("- File có bị password protect không?")
        st.error("- Kích thước file có quá lớn không?")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download - Phiên bản an toàn
    """
    try:
        # Sử dụng temporary file để đảm bảo file được tạo đúng
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            doc.save(tmp_file.name)
            
            # Đọc file đã lưu thành bytes
            with open(tmp_file.name, 'rb') as f:
                file_bytes = f.read()
            
            # Xóa temporary file
            try:
                os.unlink(tmp_file.name)
            except:
                pass
            
            return file_bytes
            
    except Exception as e:
        st.error(f"Lỗi khi lưu file: {str(e)}")
        return None

def validate_docx_file(uploaded_file):
    """
    Kiểm tra file Word có hợp lệ không
    """
    try:
        # Kiểm tra extension
        if not uploaded_file.name.lower().endswith('.docx'):
            return False, "File phải có định dạng .docx"
        
        # Kiểm tra kích thước (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False, "File quá lớn (max 50MB)"
        
        # Thử đọc file
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            doc = Document(uploaded_file)
            uploaded_file.seek(0)  # Reset lại để xử lý sau
            return True, "File hợp lệ"
        except Exception as e:
            return False, f"File không đọc được: {str(e)}"
            
    except Exception as e:
        return False, f"Lỗi validate: {str(e)}"

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
    # Sidebar với hướng dẫn
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        **Bước 1:** Tải lên file Word (.docx)
        
        **Bước 2:** Ứng dụng sẽ tự động:
        - Chuyển `$công thức$` → Equation Word
        - Chuyển `${công thức}$` → Equation Word  
        - Chuyển bảng Markdown → Bảng Word
        
        **Bước 3:** Tải xuống file đã xử lý
        
        ---
        
        **Ví dụ công thức LaTeX:**
        - `$x^2 + y^2 = z^2$`
        - `$\\frac{a}{b} + \\sqrt{c}$`
        - `$\\alpha + \\beta = \\gamma$`
        
        **Ví dụ bảng Markdown:**
        ```
        | Cột 1 | Cột 2 |
        |-------|-------|
        | A     | B     |
        | C     | D     |
        ```
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📤 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)",
            type=['docx'],
            help="Tải lên file Word chứa công thức LaTeX và/hoặc bảng Markdown"
        )
        
 → √(x)")
                                    
                                    if table_count > 0:
                                        st.write("**Bảng Markdown đã chuyển đổi:**")
                                        st.write("- Chuyển từ text sang bảng Word thực")
                                        st.write("- Header được in đậm tự động")
                                        st.write("- Áp dụng style Table Grid")
                        else:
                            st.error("❌ Không thể tạo file mới. Vui lòng thử lại.")
                    else:
                        st.error("❌ Không thể xử lý file. Vui lòng kiểm tra file và thử lại.")
                
                except Exception as e:
                    st.error(f"❌ Lỗi không mong muốn: {str(e)}")
                    st.info("💡 Vui lòng thử:")
                    st.info("- Tải lại trang và thử lại")  
                    st.info("- Kiểm tra file có mở trong Word không")
                    st.info("- Sử dụng file Word khác để test")
                
                finally:
                    # Xóa progress bar
                    progress_bar.empty()
                    status_text.empty()
    
    with col2:
        st.header("🛠️ Tính năng")
        
        features = [
            {
                "icon": "🧮",
                "title": "Chuyển đổi công thức",
                "desc": "Tự động chuyển $...$ và ${...}$ thành Equation Word"
            },
            {
                "icon": "📊", 
                "title": "Chuyển đổi bảng",
                "desc": "Chuyển bảng Markdown thành bảng Word đẹp mắt"
            },
            {
                "icon": "⚡",
                "title": "Xử lý nhanh",
                "desc": "Xử lý tự động và nhanh chóng"
            },
            {
                "icon": "💾",
                "title": "Tải xuống dễ dàng", 
                "desc": "Tải file đã xử lý với một click"
            }
        ]
        
        for feature in features:
            with st.container():
                st.markdown(f"""
                <div style="
                    border: 1px solid #e0e0e0;
                    border-radius: 10px;
                    padding: 15px;
                    margin: 10px 0;
                    background-color: #f9f9f9;
                ">
                    <h4>{feature['icon']} {feature['title']}</h4>
                    <p style="margin: 0; color: #666;">{feature['desc']}</p>
                </div>
                """, unsafe_allow_html=True)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; margin-top: 50px;">
        <p>🚀 <strong>Word Document Processor v3.0</strong> - Hỗ trợ Word Equation thực sự!</p>
        <p>📧 Xử lý LaTeX, [MATH: ...], và bảng Markdown trong tài liệu Word</p>
        <p>✨ <strong>Mới:</strong> Chế độ Word Equation thực sự + Unicode an toàn</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
    
    # Lấy text từ paragraph
    full_text = paragraph.text
    
    # Tìm tất cả công thức
    matches = list(re.finditer(math_pattern, full_text))
    
    if not matches:
        return False
    
    try:
        # Thay thế các công thức bằng Unicode
        new_text = full_text
        for match in reversed(matches):  # Reverse để không ảnh hưởng index
            latex_content = match.group(1)
            unicode_math = convert_latex_to_unicode(latex_content)
            new_text = new_text[:match.start()] + unicode_math + new_text[match.end():]
        
        # Cập nhật text của paragraph
        paragraph.text = new_text
        
        # Định dạng toán học (italic)
        for run in paragraph.runs:
            if any(char in run.text for char in 'αβγδεπσθλμωΓΔΘΛΠΣΩ∑∫∞≤≥≠±×÷'):
                run.italic = True
        
        return True
        
    except Exception as e:
        st.warning(f"Không thể xử lý công thức: {match.group(1)}")
        return False

def parse_markdown_table(text):
    """
    Phân tích bảng Markdown và trả về dữ liệu bảng
    """
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Lấy header
    header_line = lines[0]
    headers = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Bỏ qua dòng separator (dòng thứ 2)
    if len(lines) < 3:
        return [headers]
    
    # Lấy dữ liệu
    data = [headers]
    for line in lines[2:]:
        if line.strip():
            row = [cell.strip() for cell in line.split('|') if cell.strip()]
            if row:
                # Đảm bảo số cột bằng header
                while len(row) < len(headers):
                    row.append('')
                row = row[:len(headers)]
                data.append(row)
    
    return data

def replace_markdown_tables_with_word_tables(doc):
    """
    Tìm và thay thế bảng Markdown bằng bảng Word
    """
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|[^|\n]+\|[\n\r]+\|[-\s|:]+\|[\n\r]+(?:\|[^|\n]*\|[\n\r]*)+)'
    
    paragraphs_to_remove = []
    tables_to_add = []
    
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text
        matches = list(re.finditer(table_pattern, text, re.MULTILINE))
        
        for match in matches:
            table_text = match.group(1)
            table_data = parse_markdown_table(table_text)
            
            if table_data and len(table_data) > 1:
                # Lưu thông tin để thêm bảng
                tables_to_add.append({
                    'position': i,
                    'data': table_data,
                    'paragraph': paragraph
                })
                
                # Thay thế text trong paragraph
                new_text = text[:match.start()] + "[TABLE_PLACEHOLDER]" + text[match.end():]
                paragraph.text = new_text
    
    # Thêm bảng vào document
    for table_info in reversed(tables_to_add):  # Reverse để không ảnh hưởng index
        paragraph = table_info['paragraph']
        data = table_info['data']
        
        # Tạo bảng mới
        table = doc.add_table(rows=len(data), cols=len(data[0]))
        table.style = 'Table Grid'
        
        # Điền dữ liệu
        for row_idx, row_data in enumerate(data):
            row = table.rows[row_idx]
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < len(row.cells):
                    row.cells[col_idx].text = cell_data
                    # Định dạng header
                    if row_idx == 0:
                        row.cells[col_idx].paragraphs[0].runs[0].bold = True
        
        # Thay thế placeholder
        paragraph.text = paragraph.text.replace("[TABLE_PLACEHOLDER]", "")
        
        # Di chuyển bảng đến đúng vị trí
        table._element.getparent().remove(table._element)
        paragraph._element.addnext(table._element)

def process_word_document(uploaded_file):
    """
    Xử lý file Word đã tải lên
    """
    try:
        # Đọc file Word
        doc = Document(uploaded_file)
        
        # Đếm số thay đổi
        math_count = 0
        table_count = 0
        
        # Xử lý công thức toán học
        for paragraph in doc.paragraphs:
            if process_paragraph_math(paragraph):
                math_count += 1
        
        # Xử lý bảng Markdown
        replace_markdown_tables_with_word_tables(doc)
        
        # Đếm bảng (đơn giản)
        table_pattern = r'(\|[^|\n]+\|[\n\r]+\|[-\s|:]+\|[\n\r]+(?:\|[^|\n]*\|[\n\r]*)+)'
        for paragraph in doc.paragraphs:
            table_count += len(re.findall(table_pattern, paragraph.text, re.MULTILINE))
        
        return doc, math_count, table_count
        
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {str(e)}")
        return None, 0, 0

def save_document(doc, filename):
    """
    Lưu document và trả về bytes để download
    """
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# Streamlit App
def main():
    st.set_page_config(
        page_title="Word Document Processor",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🔄 Word Document Processor")
    st.markdown("**Chuyển đổi công thức LaTeX và bảng Markdown trong file Word**")
    
    # Sidebar với hướng dẫn
    with st.sidebar:
        st.header("📋 Hướng dẫn sử dụng")
        st.markdown("""
        **Bước 1:** Tải lên file Word (.docx)
        
        **Bước 2:** Ứng dụng sẽ tự động:
        - Chuyển `$công thức$` → Equation Word
        - Chuyển `${công thức}$` → Equation Word  
        - Chuyển bảng Markdown → Bảng Word
        
        **Bước 3:** Tải xuống file đã xử lý
        
        ---
        
        **Ví dụ công thức LaTeX:**
        - `$x^2 + y^2 = z^2$`
        - `$\\frac{a}{b} + \\sqrt{c}$`
        - `$\\alpha + \\beta = \\gamma$`
        
        **Ví dụ bảng Markdown:**
        ```
        | Cột 1 | Cột 2 |
        |-------|-------|
        | A     | B     |
        | C     | D     |
        ```
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📤 Tải lên file Word")
        uploaded_file = st.file_uploader(
            "Chọn file Word (.docx)",
            type=['docx'],
            help="Tải lên file Word chứa công thức LaTeX và/hoặc bảng Markdown"
        )
        
        if uploaded_file is not None:
            # Hiển thị thông tin file
            file_details = {
                "Tên file": uploaded_file.name,
                "Kích thước": f"{uploaded_file.size:,} bytes",
                "Loại file": uploaded_file.type
            }
            
            st.subheader("📄 Thông tin file")
            for key, value in file_details.items():
                st.write(f"**{key}:** {value}")
            
            # Xử lý file
            if st.button("🔄 Xử lý file", type="primary"):
                with st.spinner("Đang xử lý file..."):
                    doc, math_count, table_count = process_word_document(uploaded_file)
                    
                    if doc is not None:
                        st.success("✅ Xử lý file thành công!")
                        
                        # Hiển thị kết quả
                        col_result1, col_result2 = st.columns(2)
                        with col_result1:
                            st.metric("Công thức đã chuyển đổi", math_count)
                        with col_result2:
                            st.metric("Bảng đã chuyển đổi", table_count)
                        
                        # Tạo file download
                        processed_file = save_document(doc, uploaded_file.name)
                        
                        # Tạo tên file mới
                        base_name = uploaded_file.name.rsplit('.', 1)[0]
                        new_filename = f"{base_name}_processed.docx"
                        
                        # Nút download
                        st.download_button(
                            label="📥 Tải xuống file đã xử lý",
                            data=processed_file,
                            file_name=new_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="secondary"
                        )
                        
                        st.info("💡 File đã được xử lý và sẵn sàng tải xuống!")
    
    with col2:
        st.header("🛠️ Tính năng")
        
        features = [
            {
                "icon": "🧮",
                "title": "Chuyển đổi công thức",
                "desc": "Tự động chuyển $...$ và ${...}$ thành Equation Word"
            },
            {
                "icon": "📊", 
                "title": "Chuyển đổi bảng",
                "desc": "Chuyển bảng Markdown thành bảng Word đẹp mắt"
            },
            {
                "icon": "⚡",
                "title": "Xử lý nhanh",
                "desc": "Xử lý tự động và nhanh chóng"
            },
            {
                "icon": "💾",
                "title": "Tải xuống dễ dàng", 
                "desc": "Tải file đã xử lý với một click"
            }
        ]
        
        for feature in features:
            with st.container():
                st.markdown(f"""
                <div style="
                    border: 1px solid #e0e0e0;
                    border-radius: 10px;
                    padding: 15px;
                    margin: 10px 0;
                    background-color: #f9f9f9;
                ">
                    <h4>{feature['icon']} {feature['title']}</h4>
                    <p style="margin: 0; color: #666;">{feature['desc']}</p>
                </div>
                """, unsafe_allow_html=True)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; margin-top: 50px;">
        <p>🚀 <strong>Word Document Processor</strong> - Công cụ xử lý file Word với LaTeX và Markdown</p>
        <p>📧 Phát triển để hỗ trợ việc chuyển đổi công thức và bảng trong tài liệu Word</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
