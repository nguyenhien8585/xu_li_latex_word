#!/usr/bin/env python3
"""
🧮 Tạo file test đơn giản cho phiên bản ổn định
"""

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_simple_test():
    """Tạo file test đơn giản"""
    
    doc = Document()
    
    # Tiêu đề
    title = doc.add_heading('TEST FILE - LATEX VÀ MARKDOWN', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Thông tin
    info = doc.add_paragraph('File test để demo chuyển đổi LaTeX và Markdown table')
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Phần 1: LaTeX đơn giản
    doc.add_heading('PHẦN 1: CÔNG THỨC LATEX ĐỐN GIẢN', level=2)
    
    simple_latex = [
        "Công thức bậc hai: $x^2 + y^2 = r^2$",
        "Phân số: $\\frac{a}{b} = \\frac{c}{d}$",
        "Căn bậc hai: $\\sqrt{x + y}$",
        "Ký hiệu Hy Lạp: $\\alpha + \\beta = \\gamma$",
        "Tích phân: $\\int_0^1 x dx = \\frac{1}{2}$",
        "Tổng: $\\sum_{i=1}^n i = \\frac{n(n+1)}{2}$"
    ]
    
    for latex in simple_latex:
        doc.add_paragraph(f"• {latex}")
    
    doc.add_paragraph()
    
    # Phần 2: Bảng Markdown
    doc.add_heading('PHẦN 2: BẢNG MARKDOWN', level=2)
    
    table1_text = """
Bảng 1: Điểm thi các môn

| Tên | Toán | Lý | Hóa | Trung bình |
|-----|------|----|----|------------|
| An  | 8    | 7  | 9  | 8.0        |
| Bình| 9    | 8  | 7  | 8.0        |
| Chi | 7    | 9  | 8  | 8.0        |
| Dũng| 10   | 9  | 8  | 9.0        |

Công thức tính trung bình: $\\bar{x} = \\frac{x_1 + x_2 + x_3}{3}$
"""
    
    doc.add_paragraph(table1_text)
    
    # Phần 3: Bảng với LaTeX
    doc.add_heading('PHẦN 3: BẢNG CÓ LATEX', level=2)
    
    table2_text = """
Bảng 2: Các hàm số cơ bản

| Hàm số | Đạo hàm | Nguyên hàm |
|--------|---------|------------|
| $x^2$ | $2x$ | $\\frac{x^3}{3}$ |
| $\\sin x$ | $\\cos x$ | $-\\cos x$ |
| $e^x$ | $e^x$ | $e^x$ |
| $\\ln x$ | $\\frac{1}{x}$ | $x\\ln x - x$ |

Quy tắc tính đạo hàm: $(f \\cdot g)' = f' \\cdot g + f \\cdot g'$
"""
    
    doc.add_paragraph(table2_text)
    
    # Phần 4: Mix content
    doc.add_heading('PHẦN 4: NỘI DUNG PHỐI HỢP', level=2)
    
    mix_content = """
Trong toán học, phương trình bậc hai $ax^2 + bx + c = 0$ có nghiệm:

$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$

Bảng phân tích delta:

| Delta ($\\Delta$) | Số nghiệm | Loại nghiệm |
|-------------------|-----------|-------------|
| $\\Delta > 0$ | 2 | Phân biệt |
| $\\Delta = 0$ | 1 | Kép |
| $\\Delta < 0$ | 0 | Vô nghiệm |

Với $\\Delta = b^2 - 4ac$ là biệt thức của phương trình.
"""
    
    doc.add_paragraph(mix_content)
    
    # Lưu file
    doc.save('simple_test.docx')
    print("✅ Đã tạo file 'simple_test.docx'")
    print("📝 File test đơn giản và ổn định!")

def create_minimal_test():
    """Tạo file test tối thiểu"""
    
    doc = Document()
    
    title = doc.add_heading('MINIMAL TEST', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Chỉ có vài công thức và một bảng đơn giản
    minimal_content = """
Công thức đơn giản: $x^2 + 1 = 0$

Phân số: $\\frac{1}{2} + \\frac{1}{3} = \\frac{5}{6}$

Bảng đơn giản:

| A | B |
|---|---|
| 1 | 2 |
| 3 | 4 |

Căn bậc hai: $\\sqrt{4} = 2$
"""
    
    doc.add_paragraph(minimal_content)
    
    # Lưu file
    doc.save('minimal_test.docx')
    print("✅ Đã tạo file 'minimal_test.docx'")
    print("📝 File test tối thiểu để kiểm tra!")

if __name__ == "__main__":
    print("🚀 Tạo file test cho phiên bản ổn định...")
    print()
    
    # File test đơn giản
    create_simple_test()
    print()
    
    # File test tối thiểu
    create_minimal_test()
    print()
    
    print("✨ Hoàn thành! Files đã tạo:")
    print("   📄 simple_test.docx (test đầy đủ)")
    print("   📄 minimal_test.docx (test tối thiểu)")
    print()
    print("🔥 Chạy: streamlit run app_simple.py")
    print("📤 Upload file test và kiểm tra download!")
