import re
import html
from typing import List, Dict, Tuple

def process_latex_formulas(content: str) -> str:
    """
    Xử lý và chuẩn hóa các công thức LaTeX trong nội dung
    
    Args:
        content (str): Nội dung HTML chứa công thức LaTeX
        
    Returns:
        str: Nội dung đã xử lý công thức LaTeX
    """
    
    # Patterns cho các loại công thức LaTeX khác nhau
    patterns = {
        'inline_single': r'\$([^$]+)\$',           # $x^2$
        'inline_double': r'\$\$([^$]+)\$\$',       # $$x^2$$
        'display_single': r'\\\[([^\]]+)\\\]',     # \[x^2\]
        'display_double': r'\\\[\\\[([^\]]+)\\\]\\\]', # \[\[x^2\]\]
        'equation_env': r'\\begin\{equation\}(.*?)\\end\{equation\}',  # \begin{equation}
        'align_env': r'\\begin\{align\}(.*?)\\end\{align\}',           # \begin{align}
    }
    
    processed_content = content
    
    # Xử lý từng loại pattern
    for pattern_name, pattern in patterns.items():
        matches = re.findall(pattern, processed_content, re.DOTALL)
        
        for match in matches:
            if isinstance(match, tuple):
                latex_code = match[0] if match[0] else match[1] if len(match) > 1 else ""
            else:
                latex_code = match
            
            # Làm sạch và chuẩn hóa công thức
            cleaned_latex = clean_latex_formula(latex_code)
            
            # Tạo công thức chuẩn cho pandoc
            if pattern_name in ['display_single', 'display_double', 'equation_env', 'align_env']:
                # Công thức hiển thị (display math)
                standardized = f"$$\n{cleaned_latex}\n$$"
            else:
                # Công thức inline
                standardized = f"${cleaned_latex}$"
            
            # Thay thế trong nội dung
            original = re.search(pattern, processed_content, re.DOTALL).group(0)
            processed_content = processed_content.replace(original, standardized, 1)
    
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
