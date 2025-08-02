#!/usr/bin/env python3
"""
🧮 Script tạo file test có LaTeX và Markdown phức tạp
Để demo ứng dụng chuyển đổi
"""

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_latex_demo():
    """Tạo file demo với nhiều dạng LaTeX"""
    
    doc = Document()
    
    # Tiêu đề
    title = doc.add_heading('DEMO CHUYỂN ĐỔI LATEX VÀ MARKDOWN', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Thông tin
    info = doc.add_paragraph('File test để demo chuyển đổi công thức LaTeX và bảng Markdown')
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # PHẦN 1: Các dạng LaTeX
    doc.add_heading('PHẦN 1: CÁC DẠNG CÔNG THỨC LATEX', level=2)
    
    latex_examples = [
        "Công thức đơn giản: $x^2 + y^2 = r^2$",
        
        "Phân số: $\\frac{a}{b} = \\frac{c}{d}$ và ${\\frac{x+1}{x-1}}$",
        
        "Căn bậc hai: $\\sqrt{a^2 + b^2}$ và $\\sqrt[3]{x+y}$",
        
        "Tích phân: $\\int_0^1 x^2 dx = \\frac{1}{3}$",
        
        "Tổng: $\\sum_{i=1}^n i = \\frac{n(n+1)}{2}$",
        
        "Giới hạn: $\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1$",
        
        "Ký hiệu Hy Lạp: $\\alpha + \\beta = \\gamma$ và $\\pi \\approx 3.14159$",
        
        "So sánh: $a \\leq b \\geq c$ và $x \\neq y$",
        
        "Tập hợp: $A \\cup B \\cap C$ và $x \\in \\mathbb{R}$",
        
        "Ma trận: $\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}$"
    ]
    
    for example in latex_examples:
        doc.add_paragraph(f"• {example}")
    
    doc.add_paragraph()
    
    # PHẦN 2: Display Math
    doc.add_heading('PHẦN 2: CÔNG THỨC DISPLAY', level=2)
    
    display_examples = [
        "\\[\\int_0^{\\infty} e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}\\]",
        
        "\\[\\sum_{n=1}^{\\infty} \\frac{1}{n^2} = \\frac{\\pi^2}{6}\\]",
        
        "\\[f(x) = \\begin{cases} x^2 & \\text{if } x \\geq 0 \\\\ -x^2 & \\text{if } x < 0 \\end{cases}\\]"
    ]
    
    for example in display_examples:
        doc.add_paragraph(example)
        doc.add_paragraph()
    
    # PHẦN 3: Bảng với LaTeX
    doc.add_heading('PHẦN 3: BẢNG MARKDOWN VỚI LATEX', level=2)
    
    table_text = """
Bảng 1: Các hàm số cơ bản

| Hàm số | Đạo hàm | Tích phân |
|--------|---------|-----------|
| $x^n$ | $nx^{n-1}$ | $\\frac{x^{n+1}}{n+1}$ |
| $\\sin x$ | $\\cos x$ | $-\\cos x$ |
| $\\cos x$ | $-\\sin x$ | $\\sin x$ |
| $e^x$ | $e^x$ | $e^x$ |
| $\\ln x$ | $\\frac{1}{x}$ | $x\\ln x - x$ |

Giải thích: Các công thức này được sử dụng trong tính toán vi phân và tích phân cơ bản.
"""
    
    doc.add_paragraph(table_text)
    
    # PHẦN 4: Bảng thống kê
    doc.add_heading('PHẦN 4: BẢNG THỐNG KÊ', level=2)
    
    stats_table = """
Bảng 2: Thống kê điểm thi môn Toán

| Khoảng điểm | Số học sinh | Tần suất | Tần suất tích lũy |
|-------------|-------------|----------|-------------------|
| [0, 2) | 3 | 6% | 6% |
| [2, 4) | 8 | 16% | 22% |
| [4, 6) | 15 | 30% | 52% |
| [6, 8) | 18 | 36% | 88% |
| [8, 10] | 6 | 12% | 100% |

Công thức tính điểm trung bình: $\\bar{x} = \\frac{\\sum f_i x_i}{\\sum f_i}$

Với kết quả này, điểm trung bình là: $\\bar{x} \\approx 5.8$
"""
    
    doc.add_paragraph(stats_table)
    
    # PHẦN 5: Vật lý
    doc.add_heading('PHẦN 5: CÔNG THỨC VẬT LÝ', level=2)
    
    physics_content = """
Một số công thức Vật lý cơ bản:

1. Định luật Newton thứ hai: $F = ma$

2. Năng lượng động học: $E_k = \\frac{1}{2}mv^2$

3. Công thức Einstein: $E = mc^2$

4. Định luật Coulomb: $F = k\\frac{q_1 q_2}{r^2}$

5. Phương trình sóng: $v = f\\lambda$

Bảng đơn vị:

| Đại lượng | Ký hiệu | Đơn vị SI | Công thức |
|-----------|---------|-----------|-----------|
| Lực | $F$ | Newton (N) | $F = ma$ |
| Năng lượng | $E$ | Joule (J) | $E = \\frac{1}{2}mv^2$ |
| Công suất | $P$ | Watt (W) | $P = \\frac{E}{t}$ |
| Điện tích | $q$ | Coulomb (C) | $q = It$ |
| Điện trường | $\\vec{E}$ | N/C | $\\vec{E} = \\frac{\\vec{F}}{q}$ |

Ứng dụng: Tính toán trong bài tập về chuyển động và điện học.
"""
    
    doc.add_paragraph(physics_content)
    
    # PHẦN 6: Hóa học
    doc.add_heading('PHẦN 6: PHƯƠNG TRÌNH HÓA HỌC', level=2)
    
    chemistry_content = """
Các phản ứng hóa học với cân bằng:

1. Phản ứng cháy: $C_2H_5OH + 3O_2 \\rightarrow 2CO_2 + 3H_2O$

2. Phân ly axit: $HCl \\rightleftharpoons H^+ + Cl^-$

3. Cân bằng hóa học: $K_c = \\frac{[C]^c[D]^d}{[A]^a[B]^b}$

Bảng hằng số cân bằng:

| Phản ứng | Nhiệt độ (K) | $K_c$ | Ghi chú |
|----------|--------------|-------|---------|
| $N_2 + 3H_2 \\rightleftharpoons 2NH_3$ | 673 | $6.2 \\times 10^{-3}$ | Tổng hợp amoniac |
| $H_2 + I_2 \\rightleftharpoons 2HI$ | 731 | $5.4 \\times 10^{1}$ | Phản ứng đơn giản |
| $CO + H_2O \\rightleftharpoons CO_2 + H_2$ | 1100 | $1.0$ | Cân bằng hoàn hảo |

Công thức Van't Hoff: $\\ln K = -\\frac{\\Delta H}{RT} + \\frac{\\Delta S}{R}$
"""
    
    doc.add_paragraph(chemistry_content)
    
    # Lưu file
    doc.save('test_latex_markdown.docx')
    print("✅ Đã tạo file 'test_latex_markdown.docx'")
    print("🧮 File này chứa đầy đủ LaTeX và Markdown để test!")

def create_simple_demo():
    """Tạo file demo đơn giản"""
    
    doc = Document()
    
    title = doc.add_heading('DEMO ĐỐN GIẢN - LATEX VÀ MARKDOWN', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Công thức đơn giản
    doc.add_heading('CÔNG THỨC Cơ BẢN', level=2)
    
    simple_examples = [
        "Pythagorean: $a^2 + b^2 = c^2$",
        "Quadratic formula: $x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}$",
        "Area of circle: $A = \\pi r^2$",
        "Einstein: $E = mc^2$"
    ]
    
    for example in simple_examples:
        doc.add_paragraph(example)
    
    doc.add_paragraph()
    
    # Bảng đơn giản
    doc.add_heading('BẢNG ĐỐN GIẢN', level=2)
    
    simple_table = """
Bảng điểm số:

| Tên | Toán | Lý | Hóa |
|-----|------|----|----|
| An  | 8    | 7  | 9  |
| Bình| 9    | 8  | 7  |
| Chi | 7    | 9  | 8  |

Công thức tính điểm trung bình: $\\bar{x} = \\frac{x_1 + x_2 + x_3}{3}$
"""
    
    doc.add_paragraph(simple_table)
    
    # Lưu file
    doc.save('simple_latex_demo.docx')
    print("✅ Đã tạo file 'simple_latex_demo.docx'")
    print("📝 File demo đơn giản cho người mới bắt đầu!")

def create_edge_cases():
    """Tạo file test các trường hợp đặc biệt"""
    
    doc = Document()
    
    title = doc.add_heading('TEST EDGE CASES - LATEX VÀ MARKDOWN', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Edge cases
    edge_cases = [
        # LaTeX trong text
        "Câu có nhiều công thức: $x^2$ và $y^2$ trong cùng dòng với $z = \\sqrt{x^2 + y^2}$",
        
        # Các dạng bracket khác nhau
        "Dạng bracket: ${a \\over b}$ và $\\{x | x > 0\\}$",
        
        # LaTeX với ký tự đặc biệt
        "Ký tự đặc biệt: $a \\& b$ và $c \\% d$",
        
        # LaTeX liền nhau
        "$\\alpha$$\\beta$$\\gamma$ - ba ký tự liền nhau",
        
        # Bảng với ký tự đặc biệt
        """
| Ký hiệu | Ý nghĩa | Ví dụ |
|---------|---------|-------|
| $\\alpha$ | Alpha | $\\alpha = 30°$ |
| $\\beta$ | Beta | $\\beta \\geq 0$ |
| $\\gamma$ | Gamma | $\\gamma \\in [0,1]$ |
        """,
        
        # Bảng không cân đối
        """
| A | B | C |
|---|---|
| 1 | 2 |
| 3 | 4 | 5 | 6 |
        """,
        
        # LaTeX không hợp lệ (để test error handling)
        "LaTeX lỗi: $\\frac{a}{$ và $\\sqrt{$ - thiếu đóng ngoặc",
        
        # Bảng chỉ có header
        """
| Header 1 | Header 2 |
|----------|----------|
        """,
        
        # Text với $ nhưng không phải LaTeX
        "Giá tiền: $100 và $200 - không phải LaTeX",
        
        # Nested LaTeX
        "Nested: $\\frac{\\sqrt{a^2 + b^2}}{\\sin(\\alpha + \\beta)}$"
    ]
    
    for case in edge_cases:
        doc.add_paragraph(case)
        doc.add_paragraph()  # Khoảng cách
    
    # Lưu file
    doc.save('edge_cases_test.docx')
    print("✅ Đã tạo file 'edge_cases_test.docx'")
    print("🧪 File test các trường hợp đặc biệt!")

if __name__ == "__main__":
    print("🚀 Tạo các file test LaTeX và Markdown...")
    print()
    
    # File demo đầy đủ
    create_latex_demo()
    print()
    
    # File demo đơn giản
    create_simple_demo()
    print()
    
    # File test edge cases
    create_edge_cases()
    print()
    
    print("✨ Hoàn thành! Các file đã tạo:")
    print("   📐 test_latex_markdown.docx (Demo đầy đủ)")
    print("   📝 simple_latex_demo.docx (Demo đơn giản)")
    print("   🧪 edge_cases_test.docx (Test edge cases)")
    print()
    print("🔥 Upload các file này vào ứng dụng để test chuyển đổi!")
    print("⚡ Chạy: streamlit run app.py")
