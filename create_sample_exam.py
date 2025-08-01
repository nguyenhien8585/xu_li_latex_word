#!/usr/bin/env python3
"""
📝 Script tạo file đề thi mẫu để test ứng dụng
Chạy: python create_sample_exam.py
"""

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_sample_exam():
    """Tạo file đề thi mẫu với đầy đủ các loại câu hỏi"""
    
    doc = Document()
    
    # Tiêu đề
    title = doc.add_heading('ĐỀ THI MẪU - TOÁN HỌC', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Thông tin đề thi
    info = doc.add_paragraph('Thời gian: 90 phút | Số câu: 15 câu')
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()  # Khoảng trống
    
    # PHẦN I: TRẮC NGHIỆM
    doc.add_heading('PHẦN I: CÂU HỎI TRẮC NGHIỆM (10 câu)', level=2)
    
    # Câu trắc nghiệm mẫu
    multiple_choice_questions = [
        {
            'num': 1,
            'question': 'Tập xác định của hàm số y = √(x-1) là:',
            'A': '[1; +∞)',
            'B': '(-∞; 1]', 
            'C': '(1; +∞)',
            'D': 'ℝ'
        },
        {
            'num': 2,
            'question': 'Giá trị của lim(x→∞) (2x+1)/(x-3) là:',
            'A': '0',
            'B': '1',
            'C': '2', 
            'D': '+∞'
        },
        {
            'num': 3,
            'question': 'Đạo hàm của hàm số f(x) = x³ - 2x + 1 là:',
            'A': '3x² - 2',
            'B': '3x² + 2',
            'C': 'x² - 2x',
            'D': '3x - 2'
        },
        {
            'num': 4,
            'question': 'Phương trình x² - 5x + 6 = 0 có nghiệm là:',
            'A': 'x = 1; x = 6',
            'B': 'x = 2; x = 3',
            'C': 'x = -2; x = -3',
            'D': 'Vô nghiệm'
        },
        {
            'num': 5,
            'question': 'Giá trị nhỏ nhất của hàm số y = x² - 4x + 5 trên ℝ là:',
            'A': '1',
            'B': '2',
            'C': '3',
            'D': '5'
        }
    ]
    
    for q in multiple_choice_questions:
        # Câu hỏi
        question_para = doc.add_paragraph()
        question_para.add_run(f"Câu {q['num']}. ").bold = True
        question_para.add_run(q['question'])
        
        # Đáp án
        for option in ['A', 'B', 'C', 'D']:
            option_para = doc.add_paragraph()
            option_para.add_run(f"{option}. ").bold = True
            option_para.add_run(q[option])
        
        doc.add_paragraph()  # Khoảng trống
    
    # PHẦN II: ĐÚNG/SAI
    doc.add_heading('PHẦN II: CÂU HỎI ĐÚNG/SAI (3 câu)', level=2)
    
    true_false_questions = [
        {
            'num': 6,
            'question': 'Hàm số y = sin(x) có chu kỳ là 2π.'
        },
        {
            'num': 7,
            'question': 'Đạo hàm của hàm số y = e^x là e^x.'
        },
        {
            'num': 8,
            'question': 'Tích phân ∫₀¹ x dx = 1/2.'
        }
    ]
    
    for q in true_false_questions:
        question_para = doc.add_paragraph()
        question_para.add_run(f"Câu {q['num']}. ").bold = True
        question_para.add_run(q['question'])
        doc.add_paragraph()
    
    # PHẦN III: TỰ LUẬN
    doc.add_heading('PHẦN III: CÂU HỎI TỰ LUẬN (2 câu)', level=2)
    
    essay_questions = [
        {
            'num': 9,
            'question': 'Cho hàm số y = x³ - 3x² + 2. Tìm cực trị của hàm số và vẽ đồ thị.'
        },
        {
            'num': 10,
            'question': 'Giải phương trình lượng giác: 2sin²x + 3cosx - 3 = 0 trên [0; 2π].'
        }
    ]
    
    for q in essay_questions:
        question_para = doc.add_paragraph()
        question_para.add_run(f"Câu {q['num']}. ").bold = True
        question_para.add_run(q['question'])
        doc.add_paragraph()
        doc.add_paragraph()  # Khoảng trống cho trả lời
    
    # Lưu file
    doc.save('example_input.docx')
    print("✅ Đã tạo file 'example_input.docx'")
    print("🚀 Có thể sử dụng file này để test ứng dụng!")

def create_complex_sample():
    """Tạo file đề thi phức tạp hơn để test edge cases"""
    
    doc = Document()
    
    # Tiêu đề
    title = doc.add_heading('ĐỀ THI PHỨC TẠP - VẬT LÝ', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Câu hỏi có format khác nhau
    complex_questions = [
        "Câu 1.Tốc độ ánh sáng trong chân không là bao nhiều?\nA.3×10⁸ m/s\nB.3×10⁶ m/s\nC.3×10⁵ m/s\nD.3×10⁷ m/s",
        
        "Câu 2. Định luật Ohm được phát biểu như thế nào?\n\nA. U = I × R\nB. P = U × I\nC. E = m × c²\nD. F = m × a",
        
        "Câu3.Đơn vị của cường độ dòng điện là gì?\nA.Ampe (A)\nB.Volt (V)\nC.Ohm (Ω)\nD.Watt (W)",
        
        "Câu 4. Trong dao động điều hòa, chu kỳ T liên hệ với tần số f theo công thức nào?\nA. T = f\nB. T = 1/f\nC. T = 2πf\nD. T = f/2π",
    ]
    
    for question in complex_questions:
        para = doc.add_paragraph(question)
        doc.add_paragraph()
    
    # Thêm một số câu đúng/sai
    doc.add_heading('PHẦN ĐÚNG/SAI:', level=2)
    
    tf_questions = [
        "Câu 5. Ánh sáng có tính chất sóng và hạt.",
        "Câu 6. Electron mang điện tích dương.",
        "Câu 7. Từ trường được tạo ra bởi dòng điện."
    ]
    
    for question in tf_questions:
        para = doc.add_paragraph(question)
        doc.add_paragraph()
    
    # Lưu file
    doc.save('complex_input.docx')
    print("✅ Đã tạo file 'complex_input.docx'")
    print("🧪 File này chứa các edge cases để test parser!")

if __name__ == "__main__":
    print("🚀 Tạo file đề thi mẫu...")
    print()
    
    # Tạo file đơn giản
    create_sample_exam()
    print()
    
    # Tạo file phức tạp  
    create_complex_sample()
    print()
    
    print("✨ Hoàn thành! Có thể upload các file .docx vào ứng dụng để test.")
    print("📁 Files được tạo:")
    print("   - example_input.docx (đề thi chuẩn)")
    print("   - complex_input.docx (test edge cases)")
