#!/usr/bin/env python3
"""
📝 Script tạo file đề thi mẫu với LaTeX math và Markdown tables
Để test tính năng mới của ứng dụng
"""

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_advanced_math_exam():
    """Tạo đề thi có chứa công thức LaTeX và bảng Markdown"""
    
    doc = Document()
    
    # Tiêu đề
    title = doc.add_heading('ĐỀ THI TOÁN HỌC NÂNG CAO', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Thông tin đề thi
    info = doc.add_paragraph('Thời gian: 120 phút | Có sử dụng công thức LaTeX và bảng dữ liệu')
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # PHẦN I: TRẮC NGHIỆM VỚI LATEX
    doc.add_heading('PHẦN I: CÂU HỎI TRẮC NGHIỆM (8 câu)', level=2)
    
    math_questions = [
        {
            'num': 1,
            'question': r'Giới hạn $\lim_{x \to 0} \frac{\sin x}{x}$ có giá trị bằng:',
            'A': '0',
            'B': '1', 
            'C': '$+\infty$',
            'D': 'Không tồn tại'
        },
        {
            'num': 2,
            'question': r'Đạo hàm của hàm số $f(x) = e^{x^2}$ là:',
            'A': r'$e^{x^2}$',
            'B': r'$2xe^{x^2}$',
            'C': r'$x^2 e^{x^2}$',
            'D': r'$2e^{x^2}$'
        },
        {
            'num': 3,
            'question': r'Tích phân $\int_0^1 x^2 dx$ có giá trị:',
            'A': r'$\frac{1}{2}$',
            'B': r'$\frac{1}{3}$',
            'C': r'$\frac{2}{3}$',
            'D': '1'
        },
        {
            'num': 4,
            'question': r'Phương trình $x^2 + 4x + 4 = 0$ có nghiệm kép là:',
            'A': '$x = -2$',
            'B': '$x = 2$', 
            'C': '$x = 4$',
            'D': 'Vô nghiệm'
        },
        {
            'num': 5,
            'question': r'Công thức tính diện tích hình tròn có bán kính $r$ là:',
            'A': r'$S = \pi r$',
            'B': r'$S = \pi r^2$',
            'C': r'$S = 2\pi r$', 
            'D': r'$S = \frac{\pi r^2}{2}$'
        }
    ]
    
    for q in math_questions:
        # Câu hỏi
        question_text = f"Câu {q['num']}. {q['question']}\n"
        question_text += f"A. {q['A']}\n"
        question_text += f"B. {q['B']}\n" 
        question_text += f"C. {q['C']}\n"
        question_text += f"D. {q['D']}"
        
        doc.add_paragraph(question_text)
        doc.add_paragraph()
    
    # PHẦN II: CÂU HỎI VỚI BẢNG MARKDOWN
    doc.add_heading('PHẦN II: PHÂN TÍCH DỮ LIỆU (2 câu)', level=2)
    
    # Câu 6: Bảng thống kê
    doc.add_paragraph("""Câu 6. Cho bảng thống kê điểm thi của lớp 12A:

| Điểm | Số học sinh | Tần suất |
|------|-------------|----------|
| 0-2  | 2           | 6.7%     |
| 3-4  | 5           | 16.7%    |
| 5-6  | 10          | 33.3%    |
| 7-8  | 8           | 26.7%    |
| 9-10 | 5           | 16.7%    |

Điểm trung bình của lớp gần nhất với giá trị nào?
A. 5.8
B. 6.2
C. 6.5  
D. 7.0""")
    
    doc.add_paragraph()
    
    # Câu 7: Hàm số với LaTeX và bảng
    doc.add_paragraph(r"""Câu 7. Cho hàm số $y = x^3 - 3x^2 + 2$ có bảng biến thiên:

| x     | $-\infty$ | 0 | 2 | $+\infty$ |
|-------|-----------|---|---|-----------|
| y'    | +         | 0 | - | 0         |
| y     | $-\infty$ | 2 | -2| $+\infty$ |

Hàm số đạt cực đại tại:
A. $x = 0$
B. $x = 2$
C. $x = -2$ 
D. Không có cực đại""")
    
    doc.add_paragraph()
    
    # PHẦN III: TỰ LUẬN VỚI CÔNG THỨC PHỨC TẠP
    doc.add_heading('PHẦN III: CÂU HỎI TỰ LUẬN (2 câu)', level=2)
    
    # Câu 8
    doc.add_paragraph(r"""Câu 8. Cho tích phân $I = \int_0^{\pi/2} \sin^2 x \cos x dx$

a) Tính tích phân I bằng phương pháp đổi biến số.
b) Chứng minh rằng $I = \frac{1}{3}$""")
    
    doc.add_paragraph()
    
    # Câu 9  
    doc.add_paragraph(r"""Câu 9. Cho phương trình vi phân: $\frac{dy}{dx} + 2y = e^{-x}$

a) Giải phương trình vi phân tuyến tính cấp 1 này.
b) Tìm nghiệm riêng thỏa mãn điều kiện đầu $y(0) = 1$.
c) Vẽ đồ thị nghiệm trong khoảng $[0, 3]$.""")
    
    doc.add_paragraph()
    
    # Lưu file
    doc.save('advanced_math_exam.docx')
    print("✅ Đã tạo file 'advanced_math_exam.docx'")
    print("🧮 File này chứa công thức LaTeX và bảng Markdown!")

def create_physics_exam_with_formulas():
    """Tạo đề thi Vật lý với nhiều công thức"""
    
    doc = Document()
    
    title = doc.add_heading('ĐỀ THI VẬT LÝ 12 - ĐIỆN TỪHỌC', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    physics_questions = [
        {
            'num': 1,
            'question': r'Công thức tính lực Lorentz tác dụng lên điện tích $q$ chuyển động với vận tốc $\vec{v}$ trong từ trường $\vec{B}$ là:',
            'A': r'$\vec{F} = q\vec{v} \times \vec{B}$',
            'B': r'$\vec{F} = q\vec{E}$', 
            'C': r'$\vec{F} = q\vec{v} \cdot \vec{B}$',
            'D': r'$\vec{F} = \frac{q\vec{v}}{B}$'
        },
        {
            'num': 2,
            'question': r'Năng lượng điện từ trong mạch dao động LC được tính theo công thức:',
            'A': r'$W = \frac{1}{2}LI^2 + \frac{1}{2}\frac{Q^2}{C}$',
            'B': r'$W = LI^2 + \frac{Q^2}{C}$',
            'C': r'$W = \frac{LI^2}{2} - \frac{Q^2}{2C}$', 
            'D': r'$W = \sqrt{LI^2 + \frac{Q^2}{C}}$'
        },
        {
            'num': 3,
            'question': 'Cho bảng thông số của các sóng điện từ:',
            'table': """
| Loại sóng    | Tần số (Hz)      | Bước sóng (m) |
|--------------|------------------|---------------|
| Sóng radio   | $10^6 - 10^9$    | $300 - 0.3$   |
| Vi ba        | $10^9 - 10^{12}$ | $0.3 - 0.0003$|
| Hồng ngoại   | $10^{12}-10^{14}$| $0.0003-3×10^{-6}$|
| Ánh sáng     | $10^{14}-10^{15}$| $3×10^{-6}-3×10^{-7}$|

Sóng có tần số $f = 5 \times 10^{11}$ Hz thuộc loại nào?""",
            'A': 'Sóng radio',
            'B': 'Vi ba',
            'C': 'Hồng ngoại', 
            'D': 'Ánh sáng'
        }
    ]
    
    for q in physics_questions:
        question_text = f"Câu {q['num']}. {q['question']}\n"
        
        if 'table' in q:
            question_text += q['table'] + "\n"
        
        question_text += f"A. {q['A']}\n"
        question_text += f"B. {q['B']}\n"
        question_text += f"C. {q['C']}\n" 
        question_text += f"D. {q['D']}"
        
        doc.add_paragraph(question_text)
        doc.add_paragraph()
    
    # Câu tự luận
    doc.add_heading('PHẦN TỰ LUẬN:', level=2)
    
    doc.add_paragraph(r"""Câu 4. Một cuộn dây có độ tự cảm $L = 0.1$ H và tụ điện có điện dung $C = 10^{-6}$ F tạo thành mạch dao động LC.

a) Tính chu kỳ dao động riêng của mạch theo công thức $T = 2\pi\sqrt{LC}$.

b) Tại thời điểm $t = 0$, điện tích trên tụ điện là $Q_0 = 10^{-6}$ C. Viết phương trình:
   - Điện tích: $Q(t) = Q_0 \cos(\omega t + \phi)$  
   - Dòng điện: $I(t) = -\omega Q_0 \sin(\omega t + \phi)$

c) Tính năng lượng điện từ toàn phần của mạch.""")
    
    doc.save('physics_exam_formulas.docx')
    print("✅ Đã tạo file 'physics_exam_formulas.docx'")
    print("⚡ File Vật lý với công thức LaTeX phức tạp!")

def create_chemistry_exam():
    """Tạo đề thi Hóa học với công thức và bảng"""
    
    doc = Document()
    
    title = doc.add_heading('ĐỀ THI HÓA HỌC 12 - HÓA HỮU CƠ', level=1) 
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Câu hỏi với công thức hóa học
    chem_text = r"""Câu 1. Phản ứng đốt cháy hoàn toàn ethanol trong không khí:
$C_2H_5OH + 3O_2 \rightarrow 2CO_2 + 3H_2O$

Tính khối lượng $CO_2$ tạo thành khi đốt cháy 4.6g ethanol ($M_{C_2H_5OH} = 46$ g/mol)?
A. 4.4g
B. 8.8g  
C. 6.6g
D. 2.2g"""

    doc.add_paragraph(chem_text)
    doc.add_paragraph()
    
    # Câu với bảng dữ liệu
    table_text = """Câu 2. Cho bảng thế điện cực chuẩn:

| Cặp oxi hóa - khử | $E^0$ (V) |
|-------------------|-----------|
| $Cu^{2+}/Cu$      | +0.34     |
| $Zn^{2+}/Zn$      | -0.76     |
| $Ag^+/Ag$         | +0.80     |
| $Fe^{2+}/Fe$      | -0.44     |

Kim loại nào có tính khử mạnh nhất?
A. Cu
B. Zn
C. Ag  
D. Fe"""

    doc.add_paragraph(table_text)
    doc.add_paragraph()
    
    doc.save('chemistry_exam.docx')
    print("✅ Đã tạo file 'chemistry_exam.docx'")
    print("🧪 File Hóa học với công thức và bảng!")

if __name__ == "__main__":
    print("🚀 Tạo các file đề thi mẫu với LaTeX và Markdown...")
    print()
    
    # Tạo đề thi toán với LaTeX
    create_advanced_math_exam()
    print()
    
    # Tạo đề thi vật lý với công thức
    create_physics_exam_with_formulas()
    print()
    
    # Tạo đề thi hóa học
    create_chemistry_exam()
    print()
    
    print("✨ Hoàn thành! Các file đã tạo:")
    print("   📐 advanced_math_exam.docx (Toán học với LaTeX)")
    print("   ⚡ physics_exam_formulas.docx (Vật lý với công thức)")
    print("   🧪 chemistry_exam.docx (Hóa học với bảng)")
    print()
    print("🧠 Upload các file này vào ứng dụng để test tính năng LaTeX và Markdown!")
    print("🔬 Ứng dụng sẽ tự động nhận dạng và chuyển đổi công thức + bảng!")
