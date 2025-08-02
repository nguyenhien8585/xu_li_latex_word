import re
import os
import subprocess
import tempfile
import uuid
from pathlib import Path

def process_tikz_code(html_content, temp_dir, dpi=300):
    """
    Tìm và xử lý các đoạn TikZ trong HTML, chuyển thành hình ảnh PNG
    
    Args:
        html_content (str): Nội dung HTML chứa TikZ
        temp_dir (str): Thư mục tạm thời để lưu file
        dpi (int): Độ phân giải hình ảnh
        
    Returns:
        str: HTML đã thay thế TikZ bằng thẻ img
    """
    
    # Pattern để tìm TikZ
    tikz_pattern = r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}'
    
    # Tìm tất cả các đoạn TikZ
    tikz_matches = re.findall(tikz_pattern, html_content, re.DOTALL)
    
    if not tikz_matches:
        return html_content
    
    # Tạo thư mục images để lưu hình ảnh
    images_dir = os.path.join(temp_dir, "images")
    os.makedirs(images_dir, exist_ok=True)
    
    processed_content = html_content
    
    for i, tikz_code in enumerate(tikz_matches):
        try:
            # Tạo tên file unique
            image_id = str(uuid.uuid4())[:8]
            image_name = f"tikz_{i}_{image_id}.png"
            image_path = os.path.join(images_dir, image_name)
            
            # Compile TikZ thành PNG
            success = compile_tikz_to_png(tikz_code, image_path, temp_dir, dpi)
            
            if success:
                # Thay thế TikZ bằng thẻ img trong HTML
                tikz_full = f"\\begin{{tikzpicture}}{tikz_code}\\end{{tikzpicture}}"
                # Sử dụng relative path cho pandoc
                relative_path = f"images/{image_name}"
                img_tag = f'![TikZ Diagram {i+1}]({relative_path})'
                processed_content = processed_content.replace(tikz_full, img_tag)
                
                print(f"✅ Đã xử lý TikZ {i+1}: {image_name}")
            else:
                print(f"❌ Lỗi xử lý TikZ {i+1}")
                # Giữ nguyên code TikZ nếu không compile được
                
        except Exception as e:
            print(f"❌ Lỗi xử lý TikZ {i+1}: {str(e)}")
            continue
    
    return processed_content

def compile_tikz_to_png(tikz_code, output_path, temp_dir, dpi=300):
    """
    Compile một đoạn TikZ thành file PNG
    
    Args:
        tikz_code (str): Mã TikZ (không bao gồm begin/end tikzpicture)
        output_path (str): Đường dẫn file PNG đầu ra
        temp_dir (str): Thư mục tạm thời
        dpi (int): Độ phân giải
        
    Returns:
        bool: True nếu thành công, False nếu lỗi
    """
    
    try:
        # Tạo nội dung LaTeX hoàn chỉnh
        latex_content = create_latex_document(tikz_code)
        
        # Tạo file .tex tạm thời
        tex_file = os.path.join(temp_dir, f"tikz_{uuid.uuid4().hex}.tex")
        with open(tex_file, 'w', encoding='utf-8') as f:
            f.write(latex_content)
        
        # Compile LaTeX thành PDF
        pdf_file = tex_file.replace('.tex', '.pdf')
        success = compile_latex_to_pdf(tex_file, temp_dir)
        
        if not success or not os.path.exists(pdf_file):
            print(f"❌ Lỗi compile LaTeX: {tex_file}")
            return False
        
        # Chuyển PDF thành PNG
        success = convert_pdf_to_png(pdf_file, output_path, dpi)
        
        # Dọn dẹp file tạm thời
        cleanup_temp_files(tex_file)
        
        return success
        
    except Exception as e:
        print(f"Lỗi compile TikZ: {str(e)}")
        return False

def create_latex_document(tikz_code):
    """
    Tạo document LaTeX hoàn chỉnh chứa TikZ
    
    Args:
        tikz_code (str): Mã TikZ
        
    Returns:
        str: Nội dung LaTeX hoàn chỉnh
    """
    
    latex_template = r"""
\documentclass[border=2pt,varwidth]{standalone}
\usepackage{tikz}
\usepackage{amsmath}
\usepackage{amsfonts}
\usepackage{amssymb}
\usepackage{xcolor}
\usetikzlibrary{shapes,arrows,positioning,patterns,decorations.pathreplacing,calc,intersections}

\begin{document}
\begin{tikzpicture}
""" + tikz_code + r"""
\end{tikzpicture}
\end{document}
"""
    
    return latex_template

def compile_latex_to_pdf(tex_file, work_dir):
    """
    Compile file LaTeX thành PDF
    
    Args:
        tex_file (str): Đường dẫn file .tex
        work_dir (str): Thư mục làm việc
        
    Returns:
        bool: True nếu thành công
    """
    
    try:
        # Chạy pdflatex với options để tránh interaction
        cmd = [
            'pdflatex', 
            '-interaction=nonstopmode',
            '-halt-on-error',
            '-output-directory', work_dir,
            tex_file
        ]
        
        result = subprocess.run(
            cmd, 
            capture_output=True, 
            text=True, 
            cwd=work_dir,
            timeout=30
        )
        
        # Check nếu PDF được tạo thành công
        pdf_file = tex_file.replace('.tex', '.pdf')
        success = result.returncode == 0 and os.path.exists(pdf_file)
        
        if not success:
            print(f"❌ pdflatex error: {result.stderr}")
            
        return success
        
    except subprocess.TimeoutExpired:
        print("❌ Timeout khi compile LaTeX")
        return False
    except FileNotFoundError:
        print("❌ Không tìm thấy pdflatex. Hãy cài đặt TeX Live.")
        return False
    except Exception as e:
        print(f"❌ Lỗi compile LaTeX: {str(e)}")
        return False

def convert_pdf_to_png(pdf_file, png_file, dpi=300):
    """
    Chuyển PDF thành PNG bằng ImageMagick
    
    Args:
        pdf_file (str): Đường dẫn file PDF
        png_file (str): Đường dẫn file PNG đầu ra
        dpi (int): Độ phân giải
        
    Returns:
        bool: True nếu thành công
    """
    
    try:
        # Sử dụng ImageMagick convert với options tối ưu
        cmd = [
            'convert',
            '-density', str(dpi),  # Độ phân giải
            '-quality', '100',     # Chất lượng cao
            '-background', 'white',
            '-alpha', 'remove',    # Loại bỏ transparency
            '-trim',               # Cắt bỏ khoảng trắng thừa
            '+repage',             # Reset page geometry
            pdf_file,
            png_file
        ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        success = result.returncode == 0 and os.path.exists(png_file)
        
        if not success:
            print(f"❌ ImageMagick error: {result.stderr}")
            
        return success
        
    except subprocess.TimeoutExpired:
        print("❌ Timeout khi convert PDF to PNG")
        return False
    except FileNotFoundError:
        print("❌ Không tìm thấy ImageMagick. Hãy cài đặt ImageMagick.")
        return False
    except Exception as e:
        print(f"❌ Lỗi convert PDF to PNG: {str(e)}")
        return False

def cleanup_temp_files(tex_file):
    """
    Dọn dẹp các file tạm thời sau khi compile
    
    Args:
        tex_file (str): Đường dẫn file .tex gốc
    """
    
    base_name = os.path.splitext(tex_file)[0]
    extensions = ['.tex', '.pdf', '.aux', '.log', '.fls', '.fdb_latexmk', '.synctex.gz']
    
    for ext in extensions:
        temp_file = base_name + ext
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except:
                pass  # Bỏ qua lỗi khi xóa file

def validate_tikz_code(tikz_code):
    """
    Kiểm tra tính hợp lệ cơ bản của mã TikZ
    
    Args:
        tikz_code (str): Mã TikZ cần kiểm tra
        
    Returns:
        bool: True nếu hợp lệ
    """
    
    # Kiểm tra các lệnh TikZ cơ bản
    basic_commands = [r'\\draw', r'\\node', r'\\fill', r'\\path', r'\\coordinate']
    
    # Ít nhất phải có một lệnh TikZ
    has_command = any(re.search(cmd, tikz_code) for cmd in basic_commands)
    
    # Kiểm tra cặp ngoặc
    open_braces = tikz_code.count('{')
    close_braces = tikz_code.count('}')
    balanced_braces = (open_braces == close_braces)
    
    # Kiểm tra semicolon (TikZ commands thường kết thúc bằng ;)
    has_semicolon = ';' in tikz_code
    
    return has_command and balanced_braces and has_semicolon

def test_tikz_system():
    """
    Test xem system có thể compile TikZ không
    
    Returns:
        dict: Kết quả test từng component
    """
    
    results = {
        'pdflatex': False,
        'imagemagick': False,
        'tikz_compile': False
    }
    
    # Test pdflatex
    try:
        result = subprocess.run(['pdflatex', '--version'], capture_output=True, timeout=5)
        results['pdflatex'] = result.returncode == 0
    except:
        pass
    
    # Test ImageMagick
    try:
        result = subprocess.run(['convert', '--version'], capture_output=True, timeout=5)
        results['imagemagick'] = result.returncode == 0
    except:
        pass
    
    # Test TikZ compile với simple example
    if results['pdflatex'] and results['imagemagick']:
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                simple_tikz = r'\draw (0,0) circle (1cm);'
                test_output = os.path.join(temp_dir, 'test.png')
                results['tikz_compile'] = compile_tikz_to_png(simple_tikz, test_output, temp_dir)
        except:
            pass
    
    return results

# Các TikZ templates phổ biến cho testing
TIKZ_TEMPLATES = {
    "circle": r"\draw (0,0) circle (1cm);",
    "rectangle": r"\draw (0,0) rectangle (2,1);",
    "arrow": r"\draw[->, thick] (0,0) -- (2,0);",
    "grid": r"\draw[step=0.5cm,gray,very thin] (-1,-1) grid (3,3);",
    "function": r"\draw[domain=0:4,smooth] plot (\x,{sin(\x r)});",
    "triangle": r"\draw (0,0) -- (1,0) -- (0.5,0.866) -- cycle;",
    "flowchart": r"""
        \node[rectangle,draw] (start) at (0,0) {Start};
        \node[rectangle,draw] (process) at (0,-1) {Process};
        \node[rectangle,draw] (end) at (0,-2) {End};
        \draw[->,thick] (start) -- (process);
        \draw[->,thick] (process) -- (end);
    """,
}

def get_tikz_template(template_name):
    """
    Lấy template TikZ có sẵn
    
    Args:
        template_name (str): Tên template
        
    Returns:
        str: Mã TikZ template
    """
    return TIKZ_TEMPLATES.get(template_name, "")

def list_available_templates():
    """
    Liệt kê các template TikZ có sẵn
    
    Returns:
        list: Danh sách tên template
    """
    return list(TIKZ_TEMPLATES.keys())

def compile_tikz_to_png(tikz_code, output_path, temp_dir):
    """
    Compile một đoạn TikZ thành file PNG
    
    Args:
        tikz_code (str): Mã TikZ (không bao gồm begin/end tikzpicture)
        output_path (str): Đường dẫn file PNG đầu ra
        temp_dir (str): Thư mục tạm thời
        
    Returns:
        bool: True nếu thành công, False nếu lỗi
    """
    
    try:
        # Tạo nội dung LaTeX hoàn chỉnh
        latex_content = create_latex_document(tikz_code)
        
        # Tạo file .tex tạm thời
        tex_file = os.path.join(temp_dir, f"tikz_{uuid.uuid4()}.tex")
        with open(tex_file, 'w', encoding='utf-8') as f:
            f.write(latex_content)
        
        # Compile LaTeX thành PDF
        pdf_file = tex_file.replace('.tex', '.pdf')
        success = compile_latex_to_pdf(tex_file, temp_dir)
        
        if not success or not os.path.exists(pdf_file):
            return False
        
        # Chuyển PDF thành PNG
        success = convert_pdf_to_png(pdf_file, output_path)
        
        # Dọn dẹp file tạm thời
        cleanup_temp_files(tex_file)
        
        return success
        
    except Exception as e:
        print(f"Lỗi compile TikZ: {str(e)}")
        return False

def create_latex_document(tikz_code):
    """
    Tạo document LaTeX hoàn chỉnh chứa TikZ
    
    Args:
        tikz_code (str): Mã TikZ
        
    Returns:
        str: Nội dung LaTeX hoàn chỉnh
    """
    
    latex_template = r"""
\documentclass[border=2pt,varwidth]{standalone}
\usepackage{tikz}
\usepackage{amsmath}
\usepackage{amsfonts}
\usepackage{amssymb}
\usetikzlibrary{shapes,arrows,positioning,patterns,decorations.pathreplacing}

\begin{document}
\begin{tikzpicture}
""" + tikz_code + r"""
\end{tikzpicture}
\end{document}
"""
    
    return latex_template

def compile_latex_to_pdf(tex_file, work_dir):
    """
    Compile file LaTeX thành PDF
    
    Args:
        tex_file (str): Đường dẫn file .tex
        work_dir (str): Thư mục làm việc
        
    Returns:
        bool: True nếu thành công
    """
    
    try:
        # Chạy pdflatex
        cmd = [
            'pdflatex', 
            '-interaction=nonstopmode',
            '-output-directory', work_dir,
            tex_file
        ]
        
        result = subprocess.run(
            cmd, 
            capture_output=True, 
            text=True, 
            cwd=work_dir,
            timeout=30
        )
        
        return result.returncode == 0
        
    except subprocess.TimeoutExpired:
        print("Timeout khi compile LaTeX")
        return False
    except FileNotFoundError:
        print("Không tìm thấy pdflatex. Hãy cài đặt TeX Live.")
        return False
    except Exception as e:
        print(f"Lỗi compile LaTeX: {str(e)}")
        return False

def convert_pdf_to_png(pdf_file, png_file):
    """
    Chuyển PDF thành PNG bằng ImageMagick
    
    Args:
        pdf_file (str): Đường dẫn file PDF
        png_file (str): Đường dẫn file PNG đầu ra
        
    Returns:
        bool: True nếu thành công
    """
    
    try:
        # Sử dụng ImageMagick convert
        cmd = [
            'convert',
            '-density', '300',  # Độ phân giải cao
            '-quality', '100',   # Chất lượng cao
            '-background', 'white',
            '-alpha', 'remove',
            pdf_file,
            png_file
        ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        return result.returncode == 0 and os.path.exists(png_file)
        
    except subprocess.TimeoutExpired:
        print("Timeout khi convert PDF to PNG")
        return False
    except FileNotFoundError:
        print("Không tìm thấy ImageMagick. Hãy cài đặt ImageMagick.")
        return False
    except Exception as e:
        print(f"Lỗi convert PDF to PNG: {str(e)}")
        return False

def cleanup_temp_files(tex_file):
    """
    Dọn dẹp các file tạm thời sau khi compile
    
    Args:
        tex_file (str): Đường dẫn file .tex gốc
    """
    
    base_name = os.path.splitext(tex_file)[0]
    extensions = ['.tex', '.pdf', '.aux', '.log', '.fls', '.fdb_latexmk']
    
    for ext in extensions:
        temp_file = base_name + ext
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except:
                pass  # Bỏ qua lỗi khi xóa file

def validate_tikz_code(tikz_code):
    """
    Kiểm tra tính hợp lệ cơ bản của mã TikZ
    
    Args:
        tikz_code (str): Mã TikZ cần kiểm tra
        
    Returns:
        bool: True nếu hợp lệ
    """
    
    # Kiểm tra các lệnh TikZ cơ bản
    basic_commands = [r'\\draw', r'\\node', r'\\fill', r'\\path', r'\\coordinate']
    
    # Ít nhất phải có một lệnh TikZ
    has_command = any(re.search(cmd, tikz_code) for cmd in basic_commands)
    
    # Kiểm tra cặp ngoặc
    open_braces = tikz_code.count('{')
    close_braces = tikz_code.count('}')
    balanced_braces = (open_braces == close_braces)
    
    return has_command and balanced_braces

# Các TikZ templates phổ biến
TIKZ_TEMPLATES = {
    "circle": r"\draw (0,0) circle (1cm);",
    "rectangle": r"\draw (0,0) rectangle (2,1);",
    "arrow": r"\draw[->, thick] (0,0) -- (2,0);",
    "grid": r"\draw[step=0.5cm,gray,very thin] (-1,-1) grid (3,3);",
    "function": r"\draw[domain=0:4] plot (\x,{sin(\x r)});",
}

def get_tikz_template(template_name):
    """
    Lấy template TikZ có sẵn
    
    Args:
        template_name (str): Tên template
        
    Returns:
        str: Mã TikZ template
    """
    return TIKZ_TEMPLATES.get(template_name, "")

def list_available_templates():
    """
    Liệt kê các template TikZ có sẵn
    
    Returns:
        list: Danh sách tên template
    """
    return list(TIKZ_TEMPLATES.keys())
