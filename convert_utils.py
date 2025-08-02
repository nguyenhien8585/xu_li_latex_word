import re
import html
from typing import List, Dict, Tuple

def process_latex_formulas(content: str) -> str:
    """
    Xử lý và chuẩn hóa các công thức LaTeX trong nội dung cho pandoc
    
    Args:
        content (str): Nội dung HTML chứa công thức LaTeX
        
    Returns:
        str: Nội dung đã xử lý công thức LaTeX
    """
    
    # Patterns cho các loại công thức LaTeX khác nhau
    patterns = {
        'inline_single': r'\$([^$\n]+)\

def clean_latex_formula(latex_code: str) -> str:
    """
    Làm sạch và chuẩn hóa mã LaTeX
    
    Args:
        latex_code (str): Mã LaTeX thô
        
    Returns:
        str: Mã LaTeX đã được làm sạch
    """
    
    # Loại bỏ khoảng trắng thừa
    cleaned = latex_code.strip()
    
    # Loại bỏ các ký tự HTML entities
    cleaned = html.unescape(cleaned)
    
    # Chuẩn hóa các lệnh LaTeX phổ biến
    replacements = {
        r'\\displaystyle\s*': '',  # Bỏ displaystyle
        r'\\textstyle\s*': '',     # Bỏ textstyle
        r'\\scriptstyle\s*': '',   # Bỏ scriptstyle
        r'\\scriptscriptstyle\s*': '',  # Bỏ scriptscriptstyle
        r'\s+': ' ',               # Chuẩn hóa khoảng trắng
        r'\\\\': r'\\\\',          # Giữ line breaks
    }
    
    for pattern, replacement in replacements.items():
        cleaned = re.sub(pattern, replacement, cleaned)
    
    return cleaned.strip()

def process_markdown_tables(content: str) -> str:
    """
    Tìm và xử lý các bảng Markdown trong nội dung
    
    Args:
        content (str): Nội dung HTML chứa bảng Markdown
        
    Returns:
        str: Nội dung đã xử lý bảng
    """
    
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|.*?\|(?:\r?\n\|.*?\|)*)'
    
    # Tìm tất cả các bảng
    tables = re.findall(table_pattern, content, re.MULTILINE)
    
    processed_content = content
    
    for table_text in tables:
        # Phân tích bảng Markdown
        parsed_table = parse_markdown_table(table_text)
        
        if parsed_table:
            # Chuyển thành HTML table
            html_table = convert_table_to_html(parsed_table)
            
            # Thay thế trong nội dung
            processed_content = processed_content.replace(table_text, html_table)
    
    return processed_content

def parse_markdown_table(table_text: str) -> Dict:
    """
    Phân tích bảng Markdown thành cấu trúc dữ liệu
    
    Args:
        table_text (str): Văn bản bảng Markdown
        
    Returns:
        Dict: Cấu trúc bảng đã phân tích
    """
    
    lines = table_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Dòng đầu tiên là header
    header_line = lines[0]
    header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Dòng thứ hai là separator (|-----|-----|)
    if len(lines) < 3:
        return None
    
    separator_line = lines[1]
    alignments = parse_table_alignments(separator_line)
    
    # Các dòng còn lại là data
    data_rows = []
    for line in lines[2:]:
        if line.strip():
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                data_rows.append(cells)
    
    return {
        'headers': header_cells,
        'alignments': alignments,
        'rows': data_rows
    }

def parse_table_alignments(separator_line: str) -> List[str]:
    """
    Phân tích canh lề của bảng từ dòng separator
    
    Args:
        separator_line (str): Dòng separator (|-----|:----:|----:|)
        
    Returns:
        List[str]: Danh sách alignment ('left', 'center', 'right')
    """
    
    cells = [cell.strip() for cell in separator_line.split('|') if cell.strip()]
    alignments = []
    
    for cell in cells:
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        else:
            alignments.append('left')
    
    return alignments

def convert_table_to_html(table_data: Dict) -> str:
    """
    Chuyển đổi cấu trúc bảng thành HTML
    
    Args:
        table_data (Dict): Dữ liệu bảng đã phân tích
        
    Returns:
        str: HTML table
    """
    
    headers = table_data['headers']
    alignments = table_data.get('alignments', ['left'] * len(headers))
    rows = table_data['rows']
    
    # Tạo HTML table
    html_parts = ['<table border="1" style="border-collapse: collapse; width: 100%;">']
    
    # Header
    html_parts.append('<thead>')
    html_parts.append('<tr>')
    for i, header in enumerate(headers):
        align = alignments[i] if i < len(alignments) else 'left'
        style = f'text-align: {align}; padding: 8px; background-color: #f2f2f2;'
        html_parts.append(f'<th style="{style}">{html.escape(header)}</th>')
    html_parts.append('</tr>')
    html_parts.append('</thead>')
    
    # Body
    html_parts.append('<tbody>')
    for row in rows:
        html_parts.append('<tr>')
        for i, cell in enumerate(row):
            align = alignments[i] if i < len(alignments) else 'left'
            style = f'text-align: {align}; padding: 8px; border: 1px solid #ddd;'
            # Xử lý LaTeX trong cell nếu có
            processed_cell = process_cell_content(cell)
            html_parts.append(f'<td style="{style}">{processed_cell}</td>')
        html_parts.append('</tr>')
    html_parts.append('</tbody>')
    
    html_parts.append('</table>')
    
    return '\n'.join(html_parts)

def process_cell_content(cell_content: str) -> str:
    """
    Xử lý nội dung trong ô bảng (có thể chứa LaTeX)
    
    Args:
        cell_content (str): Nội dung ô
        
    Returns:
        str: Nội dung đã xử lý
    """
    
    # Escape HTML
    content = html.escape(cell_content)
    
    # Xử lý LaTeX inline nếu có
    latex_pattern = r'\$([^$]+)\$'
    matches = re.findall(latex_pattern, content)
    
    for match in matches:
        original = f'${match}$'
        # Giữ nguyên LaTeX cho pandoc xử lý
        content = content.replace(html.escape(original), original)
    
    return content

def extract_math_environments(content: str) -> List[Dict]:
    """
    Trích xuất tất cả các môi trường toán học từ nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        List[Dict]: Danh sách các môi trường toán học
    """
    
    environments = []
    
    # Các pattern cho môi trường toán học
    math_patterns = {
        'equation': r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        'align': r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        'gather': r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        'multline': r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',
        'displaymath': r'\\\[(.*?)\\\]',
        'inline': r'\$([^$]+)\$',
        'display': r'\$\$(.*?)\$\$'
    }
    
    for env_type, pattern in math_patterns.items():
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            environments.append({
                'type': env_type,
                'content': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'full_match': match.group(0)
            })
    
    # Sắp xếp theo vị trí xuất hiện
    environments.sort(key=lambda x: x['start'])
    
    return environments

def validate_latex_syntax(latex_code: str) -> Tuple[bool, List[str]]:
    """
    Kiểm tra cú pháp LaTeX cơ bản
    
    Args:
        latex_code (str): Mã LaTeX cần kiểm tra
        
    Returns:
        Tuple[bool, List[str]]: (Hợp lệ?, Danh sách lỗi)
    """
    
    errors = []
    
    # Kiểm tra cặp ngoặc
    braces = {'open': 0, 'close': 0}
    brackets = {'open': 0, 'close': 0}
    parens = {'open': 0, 'close': 0}
    
    for char in latex_code:
        if char == '{':
            braces['open'] += 1
        elif char == '}':
            braces['close'] += 1
        elif char == '[':
            brackets['open'] += 1
        elif char == ']':
            brackets['close'] += 1
        elif char == '(':
            parens['open'] += 1
        elif char == ')':
            parens['close'] += 1
    
    if braces['open'] != braces['close']:
        errors.append(f"Không cân bằng ngoặc nhọn: {braces['open']} mở, {braces['close']} đóng")
    
    if brackets['open'] != brackets['close']:
        errors.append(f"Không cân bằng ngoặc vuông: {brackets['open']} mở, {brackets['close']} đóng")
    
    if parens['open'] != parens['close']:
        errors.append(f"Không cân bằng ngoặc tròn: {parens['open']} mở, {parens['close']} đóng")
    
    # Kiểm tra các lệnh LaTeX cơ bản
    invalid_commands = re.findall(r'\\[a-zA-Z]*[^a-zA-Z\s\\{]', latex_code)
    if invalid_commands:
        errors.append(f"Lệnh LaTeX không hợp lệ: {', '.join(set(invalid_commands))}")
    
    return len(errors) == 0, errors

def get_math_statistics(content: str) -> Dict:
    """
    Thống kê các thành phần toán học trong nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        Dict: Thống kê chi tiết
    """
    
    stats = {
        'inline_formulas': len(re.findall(r'\$[^$]+\$', content)),
        'display_formulas': len(re.findall(r'\$\$[^$]+\$\$', content)),
        'equation_environments': len(re.findall(r'\\begin\{equation\}.*?\\end\{equation\}', content, re.DOTALL)),
        'align_environments': len(re.findall(r'\\begin\{align\}.*?\\end\{align\}', content, re.DOTALL)),
        'tables': len(re.findall(r'\|.*?\|', content)),
        'tikz_diagrams': len(re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', content, re.DOTALL))
    }
    
    stats['total_math'] = stats['inline_formulas'] + stats['display_formulas'] + stats['equation_environments'] + stats['align_environments']
    
    return stats,           # $x^2$
        'inline_double': r'\$\$([^$]+)\$\

def clean_latex_formula(latex_code: str) -> str:
    """
    Làm sạch và chuẩn hóa mã LaTeX
    
    Args:
        latex_code (str): Mã LaTeX thô
        
    Returns:
        str: Mã LaTeX đã được làm sạch
    """
    
    # Loại bỏ khoảng trắng thừa
    cleaned = latex_code.strip()
    
    # Loại bỏ các ký tự HTML entities
    cleaned = html.unescape(cleaned)
    
    # Chuẩn hóa các lệnh LaTeX phổ biến
    replacements = {
        r'\\displaystyle\s*': '',  # Bỏ displaystyle
        r'\\textstyle\s*': '',     # Bỏ textstyle
        r'\\scriptstyle\s*': '',   # Bỏ scriptstyle
        r'\\scriptscriptstyle\s*': '',  # Bỏ scriptscriptstyle
        r'\s+': ' ',               # Chuẩn hóa khoảng trắng
        r'\\\\': r'\\\\',          # Giữ line breaks
    }
    
    for pattern, replacement in replacements.items():
        cleaned = re.sub(pattern, replacement, cleaned)
    
    return cleaned.strip()

def process_markdown_tables(content: str) -> str:
    """
    Tìm và xử lý các bảng Markdown trong nội dung
    
    Args:
        content (str): Nội dung HTML chứa bảng Markdown
        
    Returns:
        str: Nội dung đã xử lý bảng
    """
    
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|.*?\|(?:\r?\n\|.*?\|)*)'
    
    # Tìm tất cả các bảng
    tables = re.findall(table_pattern, content, re.MULTILINE)
    
    processed_content = content
    
    for table_text in tables:
        # Phân tích bảng Markdown
        parsed_table = parse_markdown_table(table_text)
        
        if parsed_table:
            # Chuyển thành HTML table
            html_table = convert_table_to_html(parsed_table)
            
            # Thay thế trong nội dung
            processed_content = processed_content.replace(table_text, html_table)
    
    return processed_content

def parse_markdown_table(table_text: str) -> Dict:
    """
    Phân tích bảng Markdown thành cấu trúc dữ liệu
    
    Args:
        table_text (str): Văn bản bảng Markdown
        
    Returns:
        Dict: Cấu trúc bảng đã phân tích
    """
    
    lines = table_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Dòng đầu tiên là header
    header_line = lines[0]
    header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Dòng thứ hai là separator (|-----|-----|)
    if len(lines) < 3:
        return None
    
    separator_line = lines[1]
    alignments = parse_table_alignments(separator_line)
    
    # Các dòng còn lại là data
    data_rows = []
    for line in lines[2:]:
        if line.strip():
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                data_rows.append(cells)
    
    return {
        'headers': header_cells,
        'alignments': alignments,
        'rows': data_rows
    }

def parse_table_alignments(separator_line: str) -> List[str]:
    """
    Phân tích canh lề của bảng từ dòng separator
    
    Args:
        separator_line (str): Dòng separator (|-----|:----:|----:|)
        
    Returns:
        List[str]: Danh sách alignment ('left', 'center', 'right')
    """
    
    cells = [cell.strip() for cell in separator_line.split('|') if cell.strip()]
    alignments = []
    
    for cell in cells:
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        else:
            alignments.append('left')
    
    return alignments

def convert_table_to_html(table_data: Dict) -> str:
    """
    Chuyển đổi cấu trúc bảng thành HTML
    
    Args:
        table_data (Dict): Dữ liệu bảng đã phân tích
        
    Returns:
        str: HTML table
    """
    
    headers = table_data['headers']
    alignments = table_data.get('alignments', ['left'] * len(headers))
    rows = table_data['rows']
    
    # Tạo HTML table
    html_parts = ['<table border="1" style="border-collapse: collapse; width: 100%;">']
    
    # Header
    html_parts.append('<thead>')
    html_parts.append('<tr>')
    for i, header in enumerate(headers):
        align = alignments[i] if i < len(alignments) else 'left'
        style = f'text-align: {align}; padding: 8px; background-color: #f2f2f2;'
        html_parts.append(f'<th style="{style}">{html.escape(header)}</th>')
    html_parts.append('</tr>')
    html_parts.append('</thead>')
    
    # Body
    html_parts.append('<tbody>')
    for row in rows:
        html_parts.append('<tr>')
        for i, cell in enumerate(row):
            align = alignments[i] if i < len(alignments) else 'left'
            style = f'text-align: {align}; padding: 8px; border: 1px solid #ddd;'
            # Xử lý LaTeX trong cell nếu có
            processed_cell = process_cell_content(cell)
            html_parts.append(f'<td style="{style}">{processed_cell}</td>')
        html_parts.append('</tr>')
    html_parts.append('</tbody>')
    
    html_parts.append('</table>')
    
    return '\n'.join(html_parts)

def process_cell_content(cell_content: str) -> str:
    """
    Xử lý nội dung trong ô bảng (có thể chứa LaTeX)
    
    Args:
        cell_content (str): Nội dung ô
        
    Returns:
        str: Nội dung đã xử lý
    """
    
    # Escape HTML
    content = html.escape(cell_content)
    
    # Xử lý LaTeX inline nếu có
    latex_pattern = r'\$([^$]+)\$'
    matches = re.findall(latex_pattern, content)
    
    for match in matches:
        original = f'${match}$'
        # Giữ nguyên LaTeX cho pandoc xử lý
        content = content.replace(html.escape(original), original)
    
    return content

def extract_math_environments(content: str) -> List[Dict]:
    """
    Trích xuất tất cả các môi trường toán học từ nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        List[Dict]: Danh sách các môi trường toán học
    """
    
    environments = []
    
    # Các pattern cho môi trường toán học
    math_patterns = {
        'equation': r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        'align': r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        'gather': r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        'multline': r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',
        'displaymath': r'\\\[(.*?)\\\]',
        'inline': r'\$([^$]+)\$',
        'display': r'\$\$(.*?)\$\$'
    }
    
    for env_type, pattern in math_patterns.items():
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            environments.append({
                'type': env_type,
                'content': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'full_match': match.group(0)
            })
    
    # Sắp xếp theo vị trí xuất hiện
    environments.sort(key=lambda x: x['start'])
    
    return environments

def validate_latex_syntax(latex_code: str) -> Tuple[bool, List[str]]:
    """
    Kiểm tra cú pháp LaTeX cơ bản
    
    Args:
        latex_code (str): Mã LaTeX cần kiểm tra
        
    Returns:
        Tuple[bool, List[str]]: (Hợp lệ?, Danh sách lỗi)
    """
    
    errors = []
    
    # Kiểm tra cặp ngoặc
    braces = {'open': 0, 'close': 0}
    brackets = {'open': 0, 'close': 0}
    parens = {'open': 0, 'close': 0}
    
    for char in latex_code:
        if char == '{':
            braces['open'] += 1
        elif char == '}':
            braces['close'] += 1
        elif char == '[':
            brackets['open'] += 1
        elif char == ']':
            brackets['close'] += 1
        elif char == '(':
            parens['open'] += 1
        elif char == ')':
            parens['close'] += 1
    
    if braces['open'] != braces['close']:
        errors.append(f"Không cân bằng ngoặc nhọn: {braces['open']} mở, {braces['close']} đóng")
    
    if brackets['open'] != brackets['close']:
        errors.append(f"Không cân bằng ngoặc vuông: {brackets['open']} mở, {brackets['close']} đóng")
    
    if parens['open'] != parens['close']:
        errors.append(f"Không cân bằng ngoặc tròn: {parens['open']} mở, {parens['close']} đóng")
    
    # Kiểm tra các lệnh LaTeX cơ bản
    invalid_commands = re.findall(r'\\[a-zA-Z]*[^a-zA-Z\s\\{]', latex_code)
    if invalid_commands:
        errors.append(f"Lệnh LaTeX không hợp lệ: {', '.join(set(invalid_commands))}")
    
    return len(errors) == 0, errors

def get_math_statistics(content: str) -> Dict:
    """
    Thống kê các thành phần toán học trong nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        Dict: Thống kê chi tiết
    """
    
    stats = {
        'inline_formulas': len(re.findall(r'\$[^$]+\$', content)),
        'display_formulas': len(re.findall(r'\$\$[^$]+\$\$', content)),
        'equation_environments': len(re.findall(r'\\begin\{equation\}.*?\\end\{equation\}', content, re.DOTALL)),
        'align_environments': len(re.findall(r'\\begin\{align\}.*?\\end\{align\}', content, re.DOTALL)),
        'tables': len(re.findall(r'\|.*?\|', content)),
        'tikz_diagrams': len(re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', content, re.DOTALL))
    }
    
    stats['total_math'] = stats['inline_formulas'] + stats['display_formulas'] + stats['equation_environments'] + stats['align_environments']
    
    return stats,         # $x^2$
        'display_single': r'\\\[([^\]]+)\\\]',       # \[x^2\]
        'equation_env': r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',  # \begin{equation}
        'align_env': r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',           # \begin{align}
        'gather_env': r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',        # \begin{gather}
        'multline_env': r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',  # \begin{multline}
    }
    
    processed_content = content
    
    # Xử lý từng loại pattern theo thứ tự ưu tiên
    for pattern_name, pattern in patterns.items():
        matches = list(re.finditer(pattern, processed_content, re.DOTALL))
        
        # Xử lý từ cuối về đầu để không ảnh hưởng đến index
        for match in reversed(matches):
            latex_code = match.group(1).strip()
            
            # Làm sạch và chuẩn hóa công thức
            cleaned_latex = clean_latex_formula(latex_code)
            
            # Tạo công thức chuẩn cho pandoc
            if pattern_name in ['display_single', 'equation_env', 'align_env', 'gather_env', 'multline_env']:
                # Công thức hiển thị (display math) - pandoc format
                standardized = f"\n$\n{cleaned_latex}\n$\n"
            elif pattern_name == 'inline_double':
                # $...$ cũng là display math
                standardized = f"\n$\n{cleaned_latex}\n$\n"
            else:
                # Công thức inline
                standardized = f"${cleaned_latex}$"
            
            # Thay thế trong nội dung
            start, end = match.span()
            processed_content = processed_content[:start] + standardized + processed_content[end:]
    
    return processed_content

def clean_latex_formula(latex_code: str) -> str:
    """
    Làm sạch và chuẩn hóa mã LaTeX
    
    Args:
        latex_code (str): Mã LaTeX thô
        
    Returns:
        str: Mã LaTeX đã được làm sạch
    """
    
    # Loại bỏ khoảng trắng thừa
    cleaned = latex_code.strip()
    
    # Loại bỏ các ký tự HTML entities
    cleaned = html.unescape(cleaned)
    
    # Chuẩn hóa các lệnh LaTeX phổ biến
    replacements = {
        r'\\displaystyle\s*': '',  # Bỏ displaystyle
        r'\\textstyle\s*': '',     # Bỏ textstyle
        r'\\scriptstyle\s*': '',   # Bỏ scriptstyle
        r'\\scriptscriptstyle\s*': '',  # Bỏ scriptscriptstyle
        r'\s+': ' ',               # Chuẩn hóa khoảng trắng
        r'\\\\': r'\\\\',          # Giữ line breaks
    }
    
    for pattern, replacement in replacements.items():
        cleaned = re.sub(pattern, replacement, cleaned)
    
    # Xử lý các trường hợp đặc biệt
    cleaned = handle_special_latex_commands(cleaned)
    
    return cleaned.strip()

def handle_special_latex_commands(latex_code: str) -> str:
    """
    Xử lý các lệnh LaTeX đặc biệt
    
    Args:
        latex_code (str): Mã LaTeX
        
    Returns:
        str: Mã LaTeX đã xử lý
    """
    
    # Thay thế các lệnh không chuẩn
    special_replacements = {
        r'\\text\{([^}]+)\}': r'\\textrm{\1}',  # text -> textrm
        r'\\mathrm\{([^}]+)\}': r'\\textrm{\1}',  # mathrm -> textrm cho pandoc
        r'\\operatorname\{([^}]+)\}': r'\\textrm{\1}',  # operatorname
    }
    
    processed = latex_code
    for pattern, replacement in special_replacements.items():
        processed = re.sub(pattern, replacement, processed)
    
    return processed

def process_markdown_tables(content: str) -> str:
    """
    Tìm và xử lý các bảng Markdown trong nội dung
    
    Args:
        content (str): Nội dung HTML chứa bảng Markdown
        
    Returns:
        str: Nội dung đã xử lý bảng
    """
    
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|.*?\|(?:\r?\n\|.*?\|)*)'
    
    # Tìm tất cả các bảng
    tables = re.findall(table_pattern, content, re.MULTILINE)
    
    processed_content = content
    
    for table_text in tables:
        # Phân tích bảng Markdown
        parsed_table = parse_markdown_table(table_text)
        
        if parsed_table:
            # Chuyển thành Markdown table chuẩn cho pandoc
            markdown_table = convert_table_to_markdown(parsed_table)
            
            # Thay thế trong nội dung
            processed_content = processed_content.replace(table_text, markdown_table)
    
    # Xử lý HTML tables nếu có
    processed_content = process_html_tables(processed_content)
    
    return processed_content

def parse_markdown_table(table_text: str) -> Dict:
    """
    Phân tích bảng Markdown thành cấu trúc dữ liệu
    
    Args:
        table_text (str): Văn bản bảng Markdown
        
    Returns:
        Dict: Cấu trúc bảng đã phân tích
    """
    
    lines = [line.strip() for line in table_text.strip().split('\n') if line.strip()]
    if len(lines) < 2:
        return None
    
    # Dòng đầu tiên là header
    header_line = lines[0]
    header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Dòng thứ hai là separator (|-----|-----|)
    if len(lines) < 3:
        return None
    
    separator_line = lines[1]
    alignments = parse_table_alignments(separator_line)
    
    # Các dòng còn lại là data
    data_rows = []
    for line in lines[2:]:
        if line.strip() and '|' in line:
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                data_rows.append(cells)
    
    return {
        'headers': header_cells,
        'alignments': alignments,
        'rows': data_rows
    }

def parse_table_alignments(separator_line: str) -> List[str]:
    """
    Phân tích canh lề của bảng từ dòng separator
    
    Args:
        separator_line (str): Dòng separator (|-----|:----:|----:|)
        
    Returns:
        List[str]: Danh sách alignment ('left', 'center', 'right')
    """
    
    cells = [cell.strip() for cell in separator_line.split('|') if cell.strip()]
    alignments = []
    
    for cell in cells:
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        else:
            alignments.append('left')
    
    return alignments

def convert_table_to_markdown(table_data: Dict) -> str:
    """
    Chuyển đổi cấu trúc bảng thành Markdown chuẩn cho pandoc
    
    Args:
        table_data (Dict): Dữ liệu bảng đã phân tích
        
    Returns:
        str: Markdown table
    """
    
    headers = table_data['headers']
    alignments = table_data.get('alignments', ['left'] * len(headers))
    rows = table_data['rows']
    
    # Tạo header row
    markdown_lines = []
    header_row = '| ' + ' | '.join(headers) + ' |'
    markdown_lines.append(header_row)
    
    # Tạo separator row
    separator_cells = []
    for align in alignments:
        if align == 'center':
            separator_cells.append(':---:')
        elif align == 'right':
            separator_cells.append('---:')
        else:
            separator_cells.append('---')
    
    separator_row = '| ' + ' | '.join(separator_cells) + ' |'
    markdown_lines.append(separator_row)
    
    # Tạo data rows
    for row in rows:
        # Đảm bảo số cột đúng
        padded_row = row + [''] * (len(headers) - len(row))
        padded_row = padded_row[:len(headers)]
        
        # Xử lý LaTeX trong cells
        processed_cells = [process_cell_content(cell) for cell in padded_row]
        data_row = '| ' + ' | '.join(processed_cells) + ' |'
        markdown_lines.append(data_row)
    
    return '\n' + '\n'.join(markdown_lines) + '\n'

def process_html_tables(content: str) -> str:
    """
    Chuyển đổi HTML tables thành Markdown tables
    
    Args:
        content (str): Nội dung chứa HTML tables
        
    Returns:
        str: Nội dung với Markdown tables
    """
    
    # Simple HTML table to Markdown conversion
    # Tìm các bảng HTML đơn giản
    table_pattern = r'<table[^>]*>(.*?)</table>'
    tables = re.findall(table_pattern, content, re.DOTALL | re.IGNORECASE)
    
    for table_html in tables:
        try:
            markdown_table = html_table_to_markdown(table_html)
            if markdown_table:
                full_table_html = f'<table{table_html}</table>'
                content = content.replace(full_table_html, markdown_table)
        except:
            continue  # Skip nếu không parse được
    
    return content

def html_table_to_markdown(table_html: str) -> str:
    """
    Chuyển đổi HTML table thành Markdown
    
    Args:
        table_html (str): Nội dung HTML table
        
    Returns:
        str: Markdown table
    """
    
    # Extract rows
    rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table_html, re.DOTALL | re.IGNORECASE)
    
    if not rows:
        return ""
    
    markdown_rows = []
    
    for i, row in enumerate(rows):
        # Extract cells (th hoặc td)
        cells = re.findall(r'<(?:th|td)[^>]*>(.*?)</(?:th|td)>', row, re.DOTALL | re.IGNORECASE)
        
        if not cells:
            continue
        
        # Clean cell content
        cleaned_cells = []
        for cell in cells:
            # Remove HTML tags
            cleaned = re.sub(r'<[^>]+>', '', cell)
            cleaned = html.unescape(cleaned).strip()
            cleaned_cells.append(cleaned)
        
        # Create markdown row
        markdown_row = '| ' + ' | '.join(cleaned_cells) + ' |'
        markdown_rows.append(markdown_row)
        
        # Add separator after header row
        if i == 0:
            separator = '| ' + ' | '.join(['---'] * len(cleaned_cells)) + ' |'
            markdown_rows.append(separator)
    
    return '\n' + '\n'.join(markdown_rows) + '\n' if markdown_rows else ""

def process_cell_content(cell_content: str) -> str:
    """
    Xử lý nội dung trong ô bảng (có thể chứa LaTeX)
    
    Args:
        cell_content (str): Nội dung ô
        
    Returns:
        str: Nội dung đã xử lý
    """
    
    # Escape pipe characters trong cell
    content = cell_content.replace('|', '\\|')
    
    # Xử lý LaTeX inline nếu có (giữ nguyên cho pandoc)
    # Không cần escape vì pandoc sẽ xử lý
    
    # Remove line breaks trong cell
    content = content.replace('\n', ' ').replace('\r', ' ')
    
    # Normalize whitespace
    content = re.sub(r'\s+', ' ', content).strip()
    
    return content

def extract_math_environments(content: str) -> List[Dict]:
    """
    Trích xuất tất cả các môi trường toán học từ nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        List[Dict]: Danh sách các môi trường toán học
    """
    
    environments = []
    
    # Các pattern cho môi trường toán học
    math_patterns = {
        'equation': r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        'align': r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        'gather': r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        'multline': r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',
        'displaymath': r'\\\[(.*?)\\\]',
        'inline': r'\$([^$\n]+)\

def clean_latex_formula(latex_code: str) -> str:
    """
    Làm sạch và chuẩn hóa mã LaTeX
    
    Args:
        latex_code (str): Mã LaTeX thô
        
    Returns:
        str: Mã LaTeX đã được làm sạch
    """
    
    # Loại bỏ khoảng trắng thừa
    cleaned = latex_code.strip()
    
    # Loại bỏ các ký tự HTML entities
    cleaned = html.unescape(cleaned)
    
    # Chuẩn hóa các lệnh LaTeX phổ biến
    replacements = {
        r'\\displaystyle\s*': '',  # Bỏ displaystyle
        r'\\textstyle\s*': '',     # Bỏ textstyle
        r'\\scriptstyle\s*': '',   # Bỏ scriptstyle
        r'\\scriptscriptstyle\s*': '',  # Bỏ scriptscriptstyle
        r'\s+': ' ',               # Chuẩn hóa khoảng trắng
        r'\\\\': r'\\\\',          # Giữ line breaks
    }
    
    for pattern, replacement in replacements.items():
        cleaned = re.sub(pattern, replacement, cleaned)
    
    return cleaned.strip()

def process_markdown_tables(content: str) -> str:
    """
    Tìm và xử lý các bảng Markdown trong nội dung
    
    Args:
        content (str): Nội dung HTML chứa bảng Markdown
        
    Returns:
        str: Nội dung đã xử lý bảng
    """
    
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|.*?\|(?:\r?\n\|.*?\|)*)'
    
    # Tìm tất cả các bảng
    tables = re.findall(table_pattern, content, re.MULTILINE)
    
    processed_content = content
    
    for table_text in tables:
        # Phân tích bảng Markdown
        parsed_table = parse_markdown_table(table_text)
        
        if parsed_table:
            # Chuyển thành HTML table
            html_table = convert_table_to_html(parsed_table)
            
            # Thay thế trong nội dung
            processed_content = processed_content.replace(table_text, html_table)
    
    return processed_content

def parse_markdown_table(table_text: str) -> Dict:
    """
    Phân tích bảng Markdown thành cấu trúc dữ liệu
    
    Args:
        table_text (str): Văn bản bảng Markdown
        
    Returns:
        Dict: Cấu trúc bảng đã phân tích
    """
    
    lines = table_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Dòng đầu tiên là header
    header_line = lines[0]
    header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Dòng thứ hai là separator (|-----|-----|)
    if len(lines) < 3:
        return None
    
    separator_line = lines[1]
    alignments = parse_table_alignments(separator_line)
    
    # Các dòng còn lại là data
    data_rows = []
    for line in lines[2:]:
        if line.strip():
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                data_rows.append(cells)
    
    return {
        'headers': header_cells,
        'alignments': alignments,
        'rows': data_rows
    }

def parse_table_alignments(separator_line: str) -> List[str]:
    """
    Phân tích canh lề của bảng từ dòng separator
    
    Args:
        separator_line (str): Dòng separator (|-----|:----:|----:|)
        
    Returns:
        List[str]: Danh sách alignment ('left', 'center', 'right')
    """
    
    cells = [cell.strip() for cell in separator_line.split('|') if cell.strip()]
    alignments = []
    
    for cell in cells:
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        else:
            alignments.append('left')
    
    return alignments

def convert_table_to_html(table_data: Dict) -> str:
    """
    Chuyển đổi cấu trúc bảng thành HTML
    
    Args:
        table_data (Dict): Dữ liệu bảng đã phân tích
        
    Returns:
        str: HTML table
    """
    
    headers = table_data['headers']
    alignments = table_data.get('alignments', ['left'] * len(headers))
    rows = table_data['rows']
    
    # Tạo HTML table
    html_parts = ['<table border="1" style="border-collapse: collapse; width: 100%;">']
    
    # Header
    html_parts.append('<thead>')
    html_parts.append('<tr>')
    for i, header in enumerate(headers):
        align = alignments[i] if i < len(alignments) else 'left'
        style = f'text-align: {align}; padding: 8px; background-color: #f2f2f2;'
        html_parts.append(f'<th style="{style}">{html.escape(header)}</th>')
    html_parts.append('</tr>')
    html_parts.append('</thead>')
    
    # Body
    html_parts.append('<tbody>')
    for row in rows:
        html_parts.append('<tr>')
        for i, cell in enumerate(row):
            align = alignments[i] if i < len(alignments) else 'left'
            style = f'text-align: {align}; padding: 8px; border: 1px solid #ddd;'
            # Xử lý LaTeX trong cell nếu có
            processed_cell = process_cell_content(cell)
            html_parts.append(f'<td style="{style}">{processed_cell}</td>')
        html_parts.append('</tr>')
    html_parts.append('</tbody>')
    
    html_parts.append('</table>')
    
    return '\n'.join(html_parts)

def process_cell_content(cell_content: str) -> str:
    """
    Xử lý nội dung trong ô bảng (có thể chứa LaTeX)
    
    Args:
        cell_content (str): Nội dung ô
        
    Returns:
        str: Nội dung đã xử lý
    """
    
    # Escape HTML
    content = html.escape(cell_content)
    
    # Xử lý LaTeX inline nếu có
    latex_pattern = r'\$([^$]+)\$'
    matches = re.findall(latex_pattern, content)
    
    for match in matches:
        original = f'${match}$'
        # Giữ nguyên LaTeX cho pandoc xử lý
        content = content.replace(html.escape(original), original)
    
    return content

def extract_math_environments(content: str) -> List[Dict]:
    """
    Trích xuất tất cả các môi trường toán học từ nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        List[Dict]: Danh sách các môi trường toán học
    """
    
    environments = []
    
    # Các pattern cho môi trường toán học
    math_patterns = {
        'equation': r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        'align': r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        'gather': r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        'multline': r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',
        'displaymath': r'\\\[(.*?)\\\]',
        'inline': r'\$([^$]+)\$',
        'display': r'\$\$(.*?)\$\$'
    }
    
    for env_type, pattern in math_patterns.items():
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            environments.append({
                'type': env_type,
                'content': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'full_match': match.group(0)
            })
    
    # Sắp xếp theo vị trí xuất hiện
    environments.sort(key=lambda x: x['start'])
    
    return environments

def validate_latex_syntax(latex_code: str) -> Tuple[bool, List[str]]:
    """
    Kiểm tra cú pháp LaTeX cơ bản
    
    Args:
        latex_code (str): Mã LaTeX cần kiểm tra
        
    Returns:
        Tuple[bool, List[str]]: (Hợp lệ?, Danh sách lỗi)
    """
    
    errors = []
    
    # Kiểm tra cặp ngoặc
    braces = {'open': 0, 'close': 0}
    brackets = {'open': 0, 'close': 0}
    parens = {'open': 0, 'close': 0}
    
    for char in latex_code:
        if char == '{':
            braces['open'] += 1
        elif char == '}':
            braces['close'] += 1
        elif char == '[':
            brackets['open'] += 1
        elif char == ']':
            brackets['close'] += 1
        elif char == '(':
            parens['open'] += 1
        elif char == ')':
            parens['close'] += 1
    
    if braces['open'] != braces['close']:
        errors.append(f"Không cân bằng ngoặc nhọn: {braces['open']} mở, {braces['close']} đóng")
    
    if brackets['open'] != brackets['close']:
        errors.append(f"Không cân bằng ngoặc vuông: {brackets['open']} mở, {brackets['close']} đóng")
    
    if parens['open'] != parens['close']:
        errors.append(f"Không cân bằng ngoặc tròn: {parens['open']} mở, {parens['close']} đóng")
    
    # Kiểm tra các lệnh LaTeX cơ bản
    invalid_commands = re.findall(r'\\[a-zA-Z]*[^a-zA-Z\s\\{]', latex_code)
    if invalid_commands:
        errors.append(f"Lệnh LaTeX không hợp lệ: {', '.join(set(invalid_commands))}")
    
    return len(errors) == 0, errors

def get_math_statistics(content: str) -> Dict:
    """
    Thống kê các thành phần toán học trong nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        Dict: Thống kê chi tiết
    """
    
    stats = {
        'inline_formulas': len(re.findall(r'\$[^$]+\$', content)),
        'display_formulas': len(re.findall(r'\$\$[^$]+\$\$', content)),
        'equation_environments': len(re.findall(r'\\begin\{equation\}.*?\\end\{equation\}', content, re.DOTALL)),
        'align_environments': len(re.findall(r'\\begin\{align\}.*?\\end\{align\}', content, re.DOTALL)),
        'tables': len(re.findall(r'\|.*?\|', content)),
        'tikz_diagrams': len(re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', content, re.DOTALL))
    }
    
    stats['total_math'] = stats['inline_formulas'] + stats['display_formulas'] + stats['equation_environments'] + stats['align_environments']
    
    return stats,
        'display': r'\$\$(.*?)\$\

def clean_latex_formula(latex_code: str) -> str:
    """
    Làm sạch và chuẩn hóa mã LaTeX
    
    Args:
        latex_code (str): Mã LaTeX thô
        
    Returns:
        str: Mã LaTeX đã được làm sạch
    """
    
    # Loại bỏ khoảng trắng thừa
    cleaned = latex_code.strip()
    
    # Loại bỏ các ký tự HTML entities
    cleaned = html.unescape(cleaned)
    
    # Chuẩn hóa các lệnh LaTeX phổ biến
    replacements = {
        r'\\displaystyle\s*': '',  # Bỏ displaystyle
        r'\\textstyle\s*': '',     # Bỏ textstyle
        r'\\scriptstyle\s*': '',   # Bỏ scriptstyle
        r'\\scriptscriptstyle\s*': '',  # Bỏ scriptscriptstyle
        r'\s+': ' ',               # Chuẩn hóa khoảng trắng
        r'\\\\': r'\\\\',          # Giữ line breaks
    }
    
    for pattern, replacement in replacements.items():
        cleaned = re.sub(pattern, replacement, cleaned)
    
    return cleaned.strip()

def process_markdown_tables(content: str) -> str:
    """
    Tìm và xử lý các bảng Markdown trong nội dung
    
    Args:
        content (str): Nội dung HTML chứa bảng Markdown
        
    Returns:
        str: Nội dung đã xử lý bảng
    """
    
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|.*?\|(?:\r?\n\|.*?\|)*)'
    
    # Tìm tất cả các bảng
    tables = re.findall(table_pattern, content, re.MULTILINE)
    
    processed_content = content
    
    for table_text in tables:
        # Phân tích bảng Markdown
        parsed_table = parse_markdown_table(table_text)
        
        if parsed_table:
            # Chuyển thành HTML table
            html_table = convert_table_to_html(parsed_table)
            
            # Thay thế trong nội dung
            processed_content = processed_content.replace(table_text, html_table)
    
    return processed_content

def parse_markdown_table(table_text: str) -> Dict:
    """
    Phân tích bảng Markdown thành cấu trúc dữ liệu
    
    Args:
        table_text (str): Văn bản bảng Markdown
        
    Returns:
        Dict: Cấu trúc bảng đã phân tích
    """
    
    lines = table_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Dòng đầu tiên là header
    header_line = lines[0]
    header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Dòng thứ hai là separator (|-----|-----|)
    if len(lines) < 3:
        return None
    
    separator_line = lines[1]
    alignments = parse_table_alignments(separator_line)
    
    # Các dòng còn lại là data
    data_rows = []
    for line in lines[2:]:
        if line.strip():
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                data_rows.append(cells)
    
    return {
        'headers': header_cells,
        'alignments': alignments,
        'rows': data_rows
    }

def parse_table_alignments(separator_line: str) -> List[str]:
    """
    Phân tích canh lề của bảng từ dòng separator
    
    Args:
        separator_line (str): Dòng separator (|-----|:----:|----:|)
        
    Returns:
        List[str]: Danh sách alignment ('left', 'center', 'right')
    """
    
    cells = [cell.strip() for cell in separator_line.split('|') if cell.strip()]
    alignments = []
    
    for cell in cells:
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        else:
            alignments.append('left')
    
    return alignments

def convert_table_to_html(table_data: Dict) -> str:
    """
    Chuyển đổi cấu trúc bảng thành HTML
    
    Args:
        table_data (Dict): Dữ liệu bảng đã phân tích
        
    Returns:
        str: HTML table
    """
    
    headers = table_data['headers']
    alignments = table_data.get('alignments', ['left'] * len(headers))
    rows = table_data['rows']
    
    # Tạo HTML table
    html_parts = ['<table border="1" style="border-collapse: collapse; width: 100%;">']
    
    # Header
    html_parts.append('<thead>')
    html_parts.append('<tr>')
    for i, header in enumerate(headers):
        align = alignments[i] if i < len(alignments) else 'left'
        style = f'text-align: {align}; padding: 8px; background-color: #f2f2f2;'
        html_parts.append(f'<th style="{style}">{html.escape(header)}</th>')
    html_parts.append('</tr>')
    html_parts.append('</thead>')
    
    # Body
    html_parts.append('<tbody>')
    for row in rows:
        html_parts.append('<tr>')
        for i, cell in enumerate(row):
            align = alignments[i] if i < len(alignments) else 'left'
            style = f'text-align: {align}; padding: 8px; border: 1px solid #ddd;'
            # Xử lý LaTeX trong cell nếu có
            processed_cell = process_cell_content(cell)
            html_parts.append(f'<td style="{style}">{processed_cell}</td>')
        html_parts.append('</tr>')
    html_parts.append('</tbody>')
    
    html_parts.append('</table>')
    
    return '\n'.join(html_parts)

def process_cell_content(cell_content: str) -> str:
    """
    Xử lý nội dung trong ô bảng (có thể chứa LaTeX)
    
    Args:
        cell_content (str): Nội dung ô
        
    Returns:
        str: Nội dung đã xử lý
    """
    
    # Escape HTML
    content = html.escape(cell_content)
    
    # Xử lý LaTeX inline nếu có
    latex_pattern = r'\$([^$]+)\$'
    matches = re.findall(latex_pattern, content)
    
    for match in matches:
        original = f'${match}$'
        # Giữ nguyên LaTeX cho pandoc xử lý
        content = content.replace(html.escape(original), original)
    
    return content

def extract_math_environments(content: str) -> List[Dict]:
    """
    Trích xuất tất cả các môi trường toán học từ nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        List[Dict]: Danh sách các môi trường toán học
    """
    
    environments = []
    
    # Các pattern cho môi trường toán học
    math_patterns = {
        'equation': r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        'align': r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        'gather': r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        'multline': r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',
        'displaymath': r'\\\[(.*?)\\\]',
        'inline': r'\$([^$]+)\$',
        'display': r'\$\$(.*?)\$\$'
    }
    
    for env_type, pattern in math_patterns.items():
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            environments.append({
                'type': env_type,
                'content': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'full_match': match.group(0)
            })
    
    # Sắp xếp theo vị trí xuất hiện
    environments.sort(key=lambda x: x['start'])
    
    return environments

def validate_latex_syntax(latex_code: str) -> Tuple[bool, List[str]]:
    """
    Kiểm tra cú pháp LaTeX cơ bản
    
    Args:
        latex_code (str): Mã LaTeX cần kiểm tra
        
    Returns:
        Tuple[bool, List[str]]: (Hợp lệ?, Danh sách lỗi)
    """
    
    errors = []
    
    # Kiểm tra cặp ngoặc
    braces = {'open': 0, 'close': 0}
    brackets = {'open': 0, 'close': 0}
    parens = {'open': 0, 'close': 0}
    
    for char in latex_code:
        if char == '{':
            braces['open'] += 1
        elif char == '}':
            braces['close'] += 1
        elif char == '[':
            brackets['open'] += 1
        elif char == ']':
            brackets['close'] += 1
        elif char == '(':
            parens['open'] += 1
        elif char == ')':
            parens['close'] += 1
    
    if braces['open'] != braces['close']:
        errors.append(f"Không cân bằng ngoặc nhọn: {braces['open']} mở, {braces['close']} đóng")
    
    if brackets['open'] != brackets['close']:
        errors.append(f"Không cân bằng ngoặc vuông: {brackets['open']} mở, {brackets['close']} đóng")
    
    if parens['open'] != parens['close']:
        errors.append(f"Không cân bằng ngoặc tròn: {parens['open']} mở, {parens['close']} đóng")
    
    # Kiểm tra các lệnh LaTeX cơ bản
    invalid_commands = re.findall(r'\\[a-zA-Z]*[^a-zA-Z\s\\{]', latex_code)
    if invalid_commands:
        errors.append(f"Lệnh LaTeX không hợp lệ: {', '.join(set(invalid_commands))}")
    
    return len(errors) == 0, errors

def get_math_statistics(content: str) -> Dict:
    """
    Thống kê các thành phần toán học trong nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        Dict: Thống kê chi tiết
    """
    
    stats = {
        'inline_formulas': len(re.findall(r'\$[^$]+\$', content)),
        'display_formulas': len(re.findall(r'\$\$[^$]+\$\$', content)),
        'equation_environments': len(re.findall(r'\\begin\{equation\}.*?\\end\{equation\}', content, re.DOTALL)),
        'align_environments': len(re.findall(r'\\begin\{align\}.*?\\end\{align\}', content, re.DOTALL)),
        'tables': len(re.findall(r'\|.*?\|', content)),
        'tikz_diagrams': len(re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', content, re.DOTALL))
    }
    
    stats['total_math'] = stats['inline_formulas'] + stats['display_formulas'] + stats['equation_environments'] + stats['align_environments']
    
    return stats
    }
    
    for env_type, pattern in math_patterns.items():
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            environments.append({
                'type': env_type,
                'content': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'full_match': match.group(0)
            })
    
    # Sắp xếp theo vị trí xuất hiện
    environments.sort(key=lambda x: x['start'])
    
    return environments

def validate_latex_syntax(latex_code: str) -> Tuple[bool, List[str]]:
    """
    Kiểm tra cú pháp LaTeX cơ bản
    
    Args:
        latex_code (str): Mã LaTeX cần kiểm tra
        
    Returns:
        Tuple[bool, List[str]]: (Hợp lệ?, Danh sách lỗi)
    """
    
    errors = []
    
    # Kiểm tra cặp ngoặc
    braces = {'open': 0, 'close': 0}
    brackets = {'open': 0, 'close': 0}
    parens = {'open': 0, 'close': 0}
    
    for char in latex_code:
        if char == '{':
            braces['open'] += 1
        elif char == '}':
            braces['close'] += 1
        elif char == '[':
            brackets['open'] += 1
        elif char == ']':
            brackets['close'] += 1
        elif char == '(':
            parens['open'] += 1
        elif char == ')':
            parens['close'] += 1
    
    if braces['open'] != braces['close']:
        errors.append(f"Không cân bằng ngoặc nhọn: {braces['open']} mở, {braces['close']} đóng")
    
    if brackets['open'] != brackets['close']:
        errors.append(f"Không cân bằng ngoặc vuông: {brackets['open']} mở, {brackets['close']} đóng")
    
    if parens['open'] != parens['close']:
        errors.append(f"Không cân bằng ngoặc tròn: {parens['open']} mở, {parens['close']} đóng")
    
    # Kiểm tra các lệnh LaTeX cơ bản
    invalid_commands = re.findall(r'\\[a-zA-Z]*[^a-zA-Z\s\\{]', latex_code)
    if invalid_commands:
        errors.append(f"Lệnh LaTeX không hợp lệ: {', '.join(set(invalid_commands))}")
    
    return len(errors) == 0, errors

def get_math_statistics(content: str) -> Dict:
    """
    Thống kê các thành phần toán học trong nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        Dict: Thống kê chi tiết
    """
    
    stats = {
        'inline_formulas': len(re.findall(r'\$[^$\n]+\

def clean_latex_formula(latex_code: str) -> str:
    """
    Làm sạch và chuẩn hóa mã LaTeX
    
    Args:
        latex_code (str): Mã LaTeX thô
        
    Returns:
        str: Mã LaTeX đã được làm sạch
    """
    
    # Loại bỏ khoảng trắng thừa
    cleaned = latex_code.strip()
    
    # Loại bỏ các ký tự HTML entities
    cleaned = html.unescape(cleaned)
    
    # Chuẩn hóa các lệnh LaTeX phổ biến
    replacements = {
        r'\\displaystyle\s*': '',  # Bỏ displaystyle
        r'\\textstyle\s*': '',     # Bỏ textstyle
        r'\\scriptstyle\s*': '',   # Bỏ scriptstyle
        r'\\scriptscriptstyle\s*': '',  # Bỏ scriptscriptstyle
        r'\s+': ' ',               # Chuẩn hóa khoảng trắng
        r'\\\\': r'\\\\',          # Giữ line breaks
    }
    
    for pattern, replacement in replacements.items():
        cleaned = re.sub(pattern, replacement, cleaned)
    
    return cleaned.strip()

def process_markdown_tables(content: str) -> str:
    """
    Tìm và xử lý các bảng Markdown trong nội dung
    
    Args:
        content (str): Nội dung HTML chứa bảng Markdown
        
    Returns:
        str: Nội dung đã xử lý bảng
    """
    
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|.*?\|(?:\r?\n\|.*?\|)*)'
    
    # Tìm tất cả các bảng
    tables = re.findall(table_pattern, content, re.MULTILINE)
    
    processed_content = content
    
    for table_text in tables:
        # Phân tích bảng Markdown
        parsed_table = parse_markdown_table(table_text)
        
        if parsed_table:
            # Chuyển thành HTML table
            html_table = convert_table_to_html(parsed_table)
            
            # Thay thế trong nội dung
            processed_content = processed_content.replace(table_text, html_table)
    
    return processed_content

def parse_markdown_table(table_text: str) -> Dict:
    """
    Phân tích bảng Markdown thành cấu trúc dữ liệu
    
    Args:
        table_text (str): Văn bản bảng Markdown
        
    Returns:
        Dict: Cấu trúc bảng đã phân tích
    """
    
    lines = table_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Dòng đầu tiên là header
    header_line = lines[0]
    header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Dòng thứ hai là separator (|-----|-----|)
    if len(lines) < 3:
        return None
    
    separator_line = lines[1]
    alignments = parse_table_alignments(separator_line)
    
    # Các dòng còn lại là data
    data_rows = []
    for line in lines[2:]:
        if line.strip():
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                data_rows.append(cells)
    
    return {
        'headers': header_cells,
        'alignments': alignments,
        'rows': data_rows
    }

def parse_table_alignments(separator_line: str) -> List[str]:
    """
    Phân tích canh lề của bảng từ dòng separator
    
    Args:
        separator_line (str): Dòng separator (|-----|:----:|----:|)
        
    Returns:
        List[str]: Danh sách alignment ('left', 'center', 'right')
    """
    
    cells = [cell.strip() for cell in separator_line.split('|') if cell.strip()]
    alignments = []
    
    for cell in cells:
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        else:
            alignments.append('left')
    
    return alignments

def convert_table_to_html(table_data: Dict) -> str:
    """
    Chuyển đổi cấu trúc bảng thành HTML
    
    Args:
        table_data (Dict): Dữ liệu bảng đã phân tích
        
    Returns:
        str: HTML table
    """
    
    headers = table_data['headers']
    alignments = table_data.get('alignments', ['left'] * len(headers))
    rows = table_data['rows']
    
    # Tạo HTML table
    html_parts = ['<table border="1" style="border-collapse: collapse; width: 100%;">']
    
    # Header
    html_parts.append('<thead>')
    html_parts.append('<tr>')
    for i, header in enumerate(headers):
        align = alignments[i] if i < len(alignments) else 'left'
        style = f'text-align: {align}; padding: 8px; background-color: #f2f2f2;'
        html_parts.append(f'<th style="{style}">{html.escape(header)}</th>')
    html_parts.append('</tr>')
    html_parts.append('</thead>')
    
    # Body
    html_parts.append('<tbody>')
    for row in rows:
        html_parts.append('<tr>')
        for i, cell in enumerate(row):
            align = alignments[i] if i < len(alignments) else 'left'
            style = f'text-align: {align}; padding: 8px; border: 1px solid #ddd;'
            # Xử lý LaTeX trong cell nếu có
            processed_cell = process_cell_content(cell)
            html_parts.append(f'<td style="{style}">{processed_cell}</td>')
        html_parts.append('</tr>')
    html_parts.append('</tbody>')
    
    html_parts.append('</table>')
    
    return '\n'.join(html_parts)

def process_cell_content(cell_content: str) -> str:
    """
    Xử lý nội dung trong ô bảng (có thể chứa LaTeX)
    
    Args:
        cell_content (str): Nội dung ô
        
    Returns:
        str: Nội dung đã xử lý
    """
    
    # Escape HTML
    content = html.escape(cell_content)
    
    # Xử lý LaTeX inline nếu có
    latex_pattern = r'\$([^$]+)\$'
    matches = re.findall(latex_pattern, content)
    
    for match in matches:
        original = f'${match}$'
        # Giữ nguyên LaTeX cho pandoc xử lý
        content = content.replace(html.escape(original), original)
    
    return content

def extract_math_environments(content: str) -> List[Dict]:
    """
    Trích xuất tất cả các môi trường toán học từ nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        List[Dict]: Danh sách các môi trường toán học
    """
    
    environments = []
    
    # Các pattern cho môi trường toán học
    math_patterns = {
        'equation': r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        'align': r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        'gather': r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        'multline': r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',
        'displaymath': r'\\\[(.*?)\\\]',
        'inline': r'\$([^$]+)\$',
        'display': r'\$\$(.*?)\$\$'
    }
    
    for env_type, pattern in math_patterns.items():
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            environments.append({
                'type': env_type,
                'content': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'full_match': match.group(0)
            })
    
    # Sắp xếp theo vị trí xuất hiện
    environments.sort(key=lambda x: x['start'])
    
    return environments

def validate_latex_syntax(latex_code: str) -> Tuple[bool, List[str]]:
    """
    Kiểm tra cú pháp LaTeX cơ bản
    
    Args:
        latex_code (str): Mã LaTeX cần kiểm tra
        
    Returns:
        Tuple[bool, List[str]]: (Hợp lệ?, Danh sách lỗi)
    """
    
    errors = []
    
    # Kiểm tra cặp ngoặc
    braces = {'open': 0, 'close': 0}
    brackets = {'open': 0, 'close': 0}
    parens = {'open': 0, 'close': 0}
    
    for char in latex_code:
        if char == '{':
            braces['open'] += 1
        elif char == '}':
            braces['close'] += 1
        elif char == '[':
            brackets['open'] += 1
        elif char == ']':
            brackets['close'] += 1
        elif char == '(':
            parens['open'] += 1
        elif char == ')':
            parens['close'] += 1
    
    if braces['open'] != braces['close']:
        errors.append(f"Không cân bằng ngoặc nhọn: {braces['open']} mở, {braces['close']} đóng")
    
    if brackets['open'] != brackets['close']:
        errors.append(f"Không cân bằng ngoặc vuông: {brackets['open']} mở, {brackets['close']} đóng")
    
    if parens['open'] != parens['close']:
        errors.append(f"Không cân bằng ngoặc tròn: {parens['open']} mở, {parens['close']} đóng")
    
    # Kiểm tra các lệnh LaTeX cơ bản
    invalid_commands = re.findall(r'\\[a-zA-Z]*[^a-zA-Z\s\\{]', latex_code)
    if invalid_commands:
        errors.append(f"Lệnh LaTeX không hợp lệ: {', '.join(set(invalid_commands))}")
    
    return len(errors) == 0, errors

def get_math_statistics(content: str) -> Dict:
    """
    Thống kê các thành phần toán học trong nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        Dict: Thống kê chi tiết
    """
    
    stats = {
        'inline_formulas': len(re.findall(r'\$[^$]+\$', content)),
        'display_formulas': len(re.findall(r'\$\$[^$]+\$\$', content)),
        'equation_environments': len(re.findall(r'\\begin\{equation\}.*?\\end\{equation\}', content, re.DOTALL)),
        'align_environments': len(re.findall(r'\\begin\{align\}.*?\\end\{align\}', content, re.DOTALL)),
        'tables': len(re.findall(r'\|.*?\|', content)),
        'tikz_diagrams': len(re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', content, re.DOTALL))
    }
    
    stats['total_math'] = stats['inline_formulas'] + stats['display_formulas'] + stats['equation_environments'] + stats['align_environments']
    
    return stats, content)),
        'display_formulas': len(re.findall(r'\$\$[^$]+\$\

def clean_latex_formula(latex_code: str) -> str:
    """
    Làm sạch và chuẩn hóa mã LaTeX
    
    Args:
        latex_code (str): Mã LaTeX thô
        
    Returns:
        str: Mã LaTeX đã được làm sạch
    """
    
    # Loại bỏ khoảng trắng thừa
    cleaned = latex_code.strip()
    
    # Loại bỏ các ký tự HTML entities
    cleaned = html.unescape(cleaned)
    
    # Chuẩn hóa các lệnh LaTeX phổ biến
    replacements = {
        r'\\displaystyle\s*': '',  # Bỏ displaystyle
        r'\\textstyle\s*': '',     # Bỏ textstyle
        r'\\scriptstyle\s*': '',   # Bỏ scriptstyle
        r'\\scriptscriptstyle\s*': '',  # Bỏ scriptscriptstyle
        r'\s+': ' ',               # Chuẩn hóa khoảng trắng
        r'\\\\': r'\\\\',          # Giữ line breaks
    }
    
    for pattern, replacement in replacements.items():
        cleaned = re.sub(pattern, replacement, cleaned)
    
    return cleaned.strip()

def process_markdown_tables(content: str) -> str:
    """
    Tìm và xử lý các bảng Markdown trong nội dung
    
    Args:
        content (str): Nội dung HTML chứa bảng Markdown
        
    Returns:
        str: Nội dung đã xử lý bảng
    """
    
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|.*?\|(?:\r?\n\|.*?\|)*)'
    
    # Tìm tất cả các bảng
    tables = re.findall(table_pattern, content, re.MULTILINE)
    
    processed_content = content
    
    for table_text in tables:
        # Phân tích bảng Markdown
        parsed_table = parse_markdown_table(table_text)
        
        if parsed_table:
            # Chuyển thành HTML table
            html_table = convert_table_to_html(parsed_table)
            
            # Thay thế trong nội dung
            processed_content = processed_content.replace(table_text, html_table)
    
    return processed_content

def parse_markdown_table(table_text: str) -> Dict:
    """
    Phân tích bảng Markdown thành cấu trúc dữ liệu
    
    Args:
        table_text (str): Văn bản bảng Markdown
        
    Returns:
        Dict: Cấu trúc bảng đã phân tích
    """
    
    lines = table_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Dòng đầu tiên là header
    header_line = lines[0]
    header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Dòng thứ hai là separator (|-----|-----|)
    if len(lines) < 3:
        return None
    
    separator_line = lines[1]
    alignments = parse_table_alignments(separator_line)
    
    # Các dòng còn lại là data
    data_rows = []
    for line in lines[2:]:
        if line.strip():
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                data_rows.append(cells)
    
    return {
        'headers': header_cells,
        'alignments': alignments,
        'rows': data_rows
    }

def parse_table_alignments(separator_line: str) -> List[str]:
    """
    Phân tích canh lề của bảng từ dòng separator
    
    Args:
        separator_line (str): Dòng separator (|-----|:----:|----:|)
        
    Returns:
        List[str]: Danh sách alignment ('left', 'center', 'right')
    """
    
    cells = [cell.strip() for cell in separator_line.split('|') if cell.strip()]
    alignments = []
    
    for cell in cells:
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        else:
            alignments.append('left')
    
    return alignments

def convert_table_to_html(table_data: Dict) -> str:
    """
    Chuyển đổi cấu trúc bảng thành HTML
    
    Args:
        table_data (Dict): Dữ liệu bảng đã phân tích
        
    Returns:
        str: HTML table
    """
    
    headers = table_data['headers']
    alignments = table_data.get('alignments', ['left'] * len(headers))
    rows = table_data['rows']
    
    # Tạo HTML table
    html_parts = ['<table border="1" style="border-collapse: collapse; width: 100%;">']
    
    # Header
    html_parts.append('<thead>')
    html_parts.append('<tr>')
    for i, header in enumerate(headers):
        align = alignments[i] if i < len(alignments) else 'left'
        style = f'text-align: {align}; padding: 8px; background-color: #f2f2f2;'
        html_parts.append(f'<th style="{style}">{html.escape(header)}</th>')
    html_parts.append('</tr>')
    html_parts.append('</thead>')
    
    # Body
    html_parts.append('<tbody>')
    for row in rows:
        html_parts.append('<tr>')
        for i, cell in enumerate(row):
            align = alignments[i] if i < len(alignments) else 'left'
            style = f'text-align: {align}; padding: 8px; border: 1px solid #ddd;'
            # Xử lý LaTeX trong cell nếu có
            processed_cell = process_cell_content(cell)
            html_parts.append(f'<td style="{style}">{processed_cell}</td>')
        html_parts.append('</tr>')
    html_parts.append('</tbody>')
    
    html_parts.append('</table>')
    
    return '\n'.join(html_parts)

def process_cell_content(cell_content: str) -> str:
    """
    Xử lý nội dung trong ô bảng (có thể chứa LaTeX)
    
    Args:
        cell_content (str): Nội dung ô
        
    Returns:
        str: Nội dung đã xử lý
    """
    
    # Escape HTML
    content = html.escape(cell_content)
    
    # Xử lý LaTeX inline nếu có
    latex_pattern = r'\$([^$]+)\$'
    matches = re.findall(latex_pattern, content)
    
    for match in matches:
        original = f'${match}$'
        # Giữ nguyên LaTeX cho pandoc xử lý
        content = content.replace(html.escape(original), original)
    
    return content

def extract_math_environments(content: str) -> List[Dict]:
    """
    Trích xuất tất cả các môi trường toán học từ nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        List[Dict]: Danh sách các môi trường toán học
    """
    
    environments = []
    
    # Các pattern cho môi trường toán học
    math_patterns = {
        'equation': r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        'align': r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        'gather': r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        'multline': r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',
        'displaymath': r'\\\[(.*?)\\\]',
        'inline': r'\$([^$]+)\$',
        'display': r'\$\$(.*?)\$\$'
    }
    
    for env_type, pattern in math_patterns.items():
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            environments.append({
                'type': env_type,
                'content': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'full_match': match.group(0)
            })
    
    # Sắp xếp theo vị trí xuất hiện
    environments.sort(key=lambda x: x['start'])
    
    return environments

def validate_latex_syntax(latex_code: str) -> Tuple[bool, List[str]]:
    """
    Kiểm tra cú pháp LaTeX cơ bản
    
    Args:
        latex_code (str): Mã LaTeX cần kiểm tra
        
    Returns:
        Tuple[bool, List[str]]: (Hợp lệ?, Danh sách lỗi)
    """
    
    errors = []
    
    # Kiểm tra cặp ngoặc
    braces = {'open': 0, 'close': 0}
    brackets = {'open': 0, 'close': 0}
    parens = {'open': 0, 'close': 0}
    
    for char in latex_code:
        if char == '{':
            braces['open'] += 1
        elif char == '}':
            braces['close'] += 1
        elif char == '[':
            brackets['open'] += 1
        elif char == ']':
            brackets['close'] += 1
        elif char == '(':
            parens['open'] += 1
        elif char == ')':
            parens['close'] += 1
    
    if braces['open'] != braces['close']:
        errors.append(f"Không cân bằng ngoặc nhọn: {braces['open']} mở, {braces['close']} đóng")
    
    if brackets['open'] != brackets['close']:
        errors.append(f"Không cân bằng ngoặc vuông: {brackets['open']} mở, {brackets['close']} đóng")
    
    if parens['open'] != parens['close']:
        errors.append(f"Không cân bằng ngoặc tròn: {parens['open']} mở, {parens['close']} đóng")
    
    # Kiểm tra các lệnh LaTeX cơ bản
    invalid_commands = re.findall(r'\\[a-zA-Z]*[^a-zA-Z\s\\{]', latex_code)
    if invalid_commands:
        errors.append(f"Lệnh LaTeX không hợp lệ: {', '.join(set(invalid_commands))}")
    
    return len(errors) == 0, errors

def get_math_statistics(content: str) -> Dict:
    """
    Thống kê các thành phần toán học trong nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        Dict: Thống kê chi tiết
    """
    
    stats = {
        'inline_formulas': len(re.findall(r'\$[^$]+\$', content)),
        'display_formulas': len(re.findall(r'\$\$[^$]+\$\$', content)),
        'equation_environments': len(re.findall(r'\\begin\{equation\}.*?\\end\{equation\}', content, re.DOTALL)),
        'align_environments': len(re.findall(r'\\begin\{align\}.*?\\end\{align\}', content, re.DOTALL)),
        'tables': len(re.findall(r'\|.*?\|', content)),
        'tikz_diagrams': len(re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', content, re.DOTALL))
    }
    
    stats['total_math'] = stats['inline_formulas'] + stats['display_formulas'] + stats['equation_environments'] + stats['align_environments']
    
    return stats, content)),
        'equation_environments': len(re.findall(r'\\begin\{equation\*?\}.*?\\end\{equation\*?\}', content, re.DOTALL)),
        'align_environments': len(re.findall(r'\\begin\{align\*?\}.*?\\end\{align\*?\}', content, re.DOTALL)),
        'gather_environments': len(re.findall(r'\\begin\{gather\*?\}.*?\\end\{gather\*?\}', content, re.DOTALL)),
        'tables': len(re.findall(r'\|.*?\|', content)),
        'html_tables': len(re.findall(r'<table.*?</table>', content, re.DOTALL | re.IGNORECASE)),
        'tikz_diagrams': len(re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', content, re.DOTALL))
    }
    
    stats['total_math'] = (stats['inline_formulas'] + stats['display_formulas'] + 
                          stats['equation_environments'] + stats['align_environments'] +
                          stats['gather_environments'])
    
    return stats

def optimize_latex_for_pandoc(latex_code: str) -> str:
    """
    Tối ưu hóa LaTeX code cho pandoc
    
    Args:
        latex_code (str): Mã LaTeX
        
    Returns:
        str: Mã LaTeX tối ưu cho pandoc
    """
    
    # Các tối ưu hóa cho pandoc
    optimizations = {
        # Thay thế các lệnh không được hỗ trợ tốt
        r'\\boldsymbol\{([^}]+)\}': r'\\mathbf{\1}',
        r'\\bm\{([^}]+)\}': r'\\mathbf{\1}',
        r'\\varnothing': r'\\emptyset',
        r'\\epsilon': r'\\varepsilon',
        r'\\phi': r'\\varphi',
        
        # Chuẩn hóa spacing
        r'\\,': r'\\:',  # thin space
        r'\\;': r'\\:',  # thick space
        r'\\!': '',      # negative space
        
        # Chuẩn hóa delimiters  
        r'\\left\(': r'\\left(',
        r'\\right\)': r'\\right)',
    }
    
    optimized = latex_code
    for pattern, replacement in optimizations.items():
        optimized = re.sub(pattern, replacement, optimized)
    
    return optimized

def clean_latex_formula(latex_code: str) -> str:
    """
    Làm sạch và chuẩn hóa mã LaTeX
    
    Args:
        latex_code (str): Mã LaTeX thô
        
    Returns:
        str: Mã LaTeX đã được làm sạch
    """
    
    # Loại bỏ khoảng trắng thừa
    cleaned = latex_code.strip()
    
    # Loại bỏ các ký tự HTML entities
    cleaned = html.unescape(cleaned)
    
    # Chuẩn hóa các lệnh LaTeX phổ biến
    replacements = {
        r'\\displaystyle\s*': '',  # Bỏ displaystyle
        r'\\textstyle\s*': '',     # Bỏ textstyle
        r'\\scriptstyle\s*': '',   # Bỏ scriptstyle
        r'\\scriptscriptstyle\s*': '',  # Bỏ scriptscriptstyle
        r'\s+': ' ',               # Chuẩn hóa khoảng trắng
        r'\\\\': r'\\\\',          # Giữ line breaks
    }
    
    for pattern, replacement in replacements.items():
        cleaned = re.sub(pattern, replacement, cleaned)
    
    return cleaned.strip()

def process_markdown_tables(content: str) -> str:
    """
    Tìm và xử lý các bảng Markdown trong nội dung
    
    Args:
        content (str): Nội dung HTML chứa bảng Markdown
        
    Returns:
        str: Nội dung đã xử lý bảng
    """
    
    # Pattern để tìm bảng Markdown
    table_pattern = r'(\|.*?\|(?:\r?\n\|.*?\|)*)'
    
    # Tìm tất cả các bảng
    tables = re.findall(table_pattern, content, re.MULTILINE)
    
    processed_content = content
    
    for table_text in tables:
        # Phân tích bảng Markdown
        parsed_table = parse_markdown_table(table_text)
        
        if parsed_table:
            # Chuyển thành HTML table
            html_table = convert_table_to_html(parsed_table)
            
            # Thay thế trong nội dung
            processed_content = processed_content.replace(table_text, html_table)
    
    return processed_content

def parse_markdown_table(table_text: str) -> Dict:
    """
    Phân tích bảng Markdown thành cấu trúc dữ liệu
    
    Args:
        table_text (str): Văn bản bảng Markdown
        
    Returns:
        Dict: Cấu trúc bảng đã phân tích
    """
    
    lines = table_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    # Dòng đầu tiên là header
    header_line = lines[0]
    header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()]
    
    # Dòng thứ hai là separator (|-----|-----|)
    if len(lines) < 3:
        return None
    
    separator_line = lines[1]
    alignments = parse_table_alignments(separator_line)
    
    # Các dòng còn lại là data
    data_rows = []
    for line in lines[2:]:
        if line.strip():
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                data_rows.append(cells)
    
    return {
        'headers': header_cells,
        'alignments': alignments,
        'rows': data_rows
    }

def parse_table_alignments(separator_line: str) -> List[str]:
    """
    Phân tích canh lề của bảng từ dòng separator
    
    Args:
        separator_line (str): Dòng separator (|-----|:----:|----:|)
        
    Returns:
        List[str]: Danh sách alignment ('left', 'center', 'right')
    """
    
    cells = [cell.strip() for cell in separator_line.split('|') if cell.strip()]
    alignments = []
    
    for cell in cells:
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        else:
            alignments.append('left')
    
    return alignments

def convert_table_to_html(table_data: Dict) -> str:
    """
    Chuyển đổi cấu trúc bảng thành HTML
    
    Args:
        table_data (Dict): Dữ liệu bảng đã phân tích
        
    Returns:
        str: HTML table
    """
    
    headers = table_data['headers']
    alignments = table_data.get('alignments', ['left'] * len(headers))
    rows = table_data['rows']
    
    # Tạo HTML table
    html_parts = ['<table border="1" style="border-collapse: collapse; width: 100%;">']
    
    # Header
    html_parts.append('<thead>')
    html_parts.append('<tr>')
    for i, header in enumerate(headers):
        align = alignments[i] if i < len(alignments) else 'left'
        style = f'text-align: {align}; padding: 8px; background-color: #f2f2f2;'
        html_parts.append(f'<th style="{style}">{html.escape(header)}</th>')
    html_parts.append('</tr>')
    html_parts.append('</thead>')
    
    # Body
    html_parts.append('<tbody>')
    for row in rows:
        html_parts.append('<tr>')
        for i, cell in enumerate(row):
            align = alignments[i] if i < len(alignments) else 'left'
            style = f'text-align: {align}; padding: 8px; border: 1px solid #ddd;'
            # Xử lý LaTeX trong cell nếu có
            processed_cell = process_cell_content(cell)
            html_parts.append(f'<td style="{style}">{processed_cell}</td>')
        html_parts.append('</tr>')
    html_parts.append('</tbody>')
    
    html_parts.append('</table>')
    
    return '\n'.join(html_parts)

def process_cell_content(cell_content: str) -> str:
    """
    Xử lý nội dung trong ô bảng (có thể chứa LaTeX)
    
    Args:
        cell_content (str): Nội dung ô
        
    Returns:
        str: Nội dung đã xử lý
    """
    
    # Escape HTML
    content = html.escape(cell_content)
    
    # Xử lý LaTeX inline nếu có
    latex_pattern = r'\$([^$]+)\$'
    matches = re.findall(latex_pattern, content)
    
    for match in matches:
        original = f'${match}$'
        # Giữ nguyên LaTeX cho pandoc xử lý
        content = content.replace(html.escape(original), original)
    
    return content

def extract_math_environments(content: str) -> List[Dict]:
    """
    Trích xuất tất cả các môi trường toán học từ nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        List[Dict]: Danh sách các môi trường toán học
    """
    
    environments = []
    
    # Các pattern cho môi trường toán học
    math_patterns = {
        'equation': r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        'align': r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        'gather': r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        'multline': r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',
        'displaymath': r'\\\[(.*?)\\\]',
        'inline': r'\$([^$]+)\$',
        'display': r'\$\$(.*?)\$\$'
    }
    
    for env_type, pattern in math_patterns.items():
        matches = re.finditer(pattern, content, re.DOTALL)
        for match in matches:
            environments.append({
                'type': env_type,
                'content': match.group(1).strip(),
                'start': match.start(),
                'end': match.end(),
                'full_match': match.group(0)
            })
    
    # Sắp xếp theo vị trí xuất hiện
    environments.sort(key=lambda x: x['start'])
    
    return environments

def validate_latex_syntax(latex_code: str) -> Tuple[bool, List[str]]:
    """
    Kiểm tra cú pháp LaTeX cơ bản
    
    Args:
        latex_code (str): Mã LaTeX cần kiểm tra
        
    Returns:
        Tuple[bool, List[str]]: (Hợp lệ?, Danh sách lỗi)
    """
    
    errors = []
    
    # Kiểm tra cặp ngoặc
    braces = {'open': 0, 'close': 0}
    brackets = {'open': 0, 'close': 0}
    parens = {'open': 0, 'close': 0}
    
    for char in latex_code:
        if char == '{':
            braces['open'] += 1
        elif char == '}':
            braces['close'] += 1
        elif char == '[':
            brackets['open'] += 1
        elif char == ']':
            brackets['close'] += 1
        elif char == '(':
            parens['open'] += 1
        elif char == ')':
            parens['close'] += 1
    
    if braces['open'] != braces['close']:
        errors.append(f"Không cân bằng ngoặc nhọn: {braces['open']} mở, {braces['close']} đóng")
    
    if brackets['open'] != brackets['close']:
        errors.append(f"Không cân bằng ngoặc vuông: {brackets['open']} mở, {brackets['close']} đóng")
    
    if parens['open'] != parens['close']:
        errors.append(f"Không cân bằng ngoặc tròn: {parens['open']} mở, {parens['close']} đóng")
    
    # Kiểm tra các lệnh LaTeX cơ bản
    invalid_commands = re.findall(r'\\[a-zA-Z]*[^a-zA-Z\s\\{]', latex_code)
    if invalid_commands:
        errors.append(f"Lệnh LaTeX không hợp lệ: {', '.join(set(invalid_commands))}")
    
    return len(errors) == 0, errors

def get_math_statistics(content: str) -> Dict:
    """
    Thống kê các thành phần toán học trong nội dung
    
    Args:
        content (str): Nội dung cần phân tích
        
    Returns:
        Dict: Thống kê chi tiết
    """
    
    stats = {
        'inline_formulas': len(re.findall(r'\$[^$]+\$', content)),
        'display_formulas': len(re.findall(r'\$\$[^$]+\$\$', content)),
        'equation_environments': len(re.findall(r'\\begin\{equation\}.*?\\end\{equation\}', content, re.DOTALL)),
        'align_environments': len(re.findall(r'\\begin\{align\}.*?\\end\{align\}', content, re.DOTALL)),
        'tables': len(re.findall(r'\|.*?\|', content)),
        'tikz_diagrams': len(re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', content, re.DOTALL))
    }
    
    stats['total_math'] = stats['inline_formulas'] + stats['display_formulas'] + stats['equation_environments'] + stats['align_environments']
    
    return stats
