#!/usr/bin/env python3
"""
📝 Module xử lý đề thi Word thành bảng chuẩn hóa
Phân tích cấu trúc đề thi và tạo bảng Word chuyên nghiệp
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
import re
from io import BytesIO

def analyze_exam_structure(doc):
    """Phân tích cấu trúc đề thi để hiểu nội dung"""
    
    analysis = {
        'total_questions': 0,
        'multiple_choice': 0,
        'true_false': 0,
        'essay': 0,
        'sample_questions': [],
        'sections': []
    }
    
    current_section = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        
        if not text:
            continue
        
        # Nhận dạng phần đề thi
        section_match = re.match(r'(PHẦN|Phần)\s*(I+|1|2|3|II|III)', text, re.IGNORECASE)
        if section_match:
            current_section = {
                'name': text,
                'type': determine_section_type(text),
                'questions': []
            }
            analysis['sections'].append(current_section)
            continue
        
        # Nhận dạng câu hỏi
        question_match = re.match(r'(Câu|CÂU)\s*(\d+)', text, re.IGNORECASE)
        if question_match:
            analysis['total_questions'] += 1
            
            # Phân loại câu hỏi
            question_type = classify_question_type(text, para)
            
            if question_type == 'multiple_choice':
                analysis['multiple_choice'] += 1
            elif question_type == 'true_false':
                analysis['true_false'] += 1
            elif question_type == 'essay':
                analysis['essay'] += 1
            
            # Lưu mẫu câu hỏi
            if len(analysis['sample_questions']) < 5:
                parsed_question = parse_question(text, question_type)
                if parsed_question:
                    analysis['sample_questions'].append(parsed_question)
    
    return analysis

def determine_section_type(section_text):
    """Xác định loại phần đề thi"""
    section_lower = section_text.lower()
    
    if any(keyword in section_lower for keyword in ['trắc nghiệm', 'lựa chọn', 'multiple']):
        return 'multiple_choice'
    elif any(keyword in section_lower for keyword in ['đúng', 'sai', 'true', 'false']):
        return 'true_false'
    elif any(keyword in section_lower for keyword in ['tự luận', 'essay', 'viết']):
        return 'essay'
    else:
        return 'unknown'

def classify_question_type(text, paragraph=None):
    """Phân loại loại câu hỏi dựa trên nội dung"""
    
    # Kiểm tra có đáp án A, B, C, D không
    if re.search(r'[A-D]\.', text):
        return 'multiple_choice'
    
    # Kiểm tra câu đúng/sai
    if any(keyword in text.lower() for keyword in ['đúng', 'sai', 'true', 'false']):
        return 'true_false'
    
    # Mặc định là tự luận
    return 'essay'

def parse_question(text, question_type):
    """Phân tích chi tiết câu hỏi"""
    
    result = {
        'type': question_type,
        'number': '',
        'question': '',
        'raw_text': text
    }
    
    # Tách số câu hỏi
    number_match = re.match(r'(Câu|CÂU)\s*(\d+)', text, re.IGNORECASE)
    if number_match:
        result['number'] = number_match.group(2)
    
    if question_type == 'multiple_choice':
        # Phân tích câu trắc nghiệm
        mc_parsed = extract_choice_question(text)
        if mc_parsed:
            result.update(mc_parsed)
    
    elif question_type == 'true_false':
        # Phân tích câu đúng/sai
        tf_parsed = extract_true_false_question(text)
        if tf_parsed:
            result.update(tf_parsed)
    
    else:
        # Câu tự luận
        essay_parsed = extract_essay_question(text)
        if essay_parsed:
            result.update(essay_parsed)
    
    return result

def extract_choice_question(text):
    """Trích xuất câu hỏi trắc nghiệm với 4 đáp án A, B, C, D"""
    
    # Pattern phức tạp để bắt câu hỏi trắc nghiệm
    patterns = [
        # Pattern 1: Câu hỏi và đáp án trên cùng một dòng
        r'(Câu\s*\d+)\.(.*?)\s*A\.\s*(.*?)\s*B\.\s*(.*?)\s*C\.\s*(.*?)\s*D\.\s*(.*?)$',
        
        # Pattern 2: Câu hỏi trên dòng riêng
        r'(Câu\s*\d+)\.(.*?)(?=\s*A\.)',
        
        # Pattern 3: Flexible pattern
        r'(Câu\s*\d+)\.(.*?)(?:A\.\s*(.*?))?(?:B\.\s*(.*?))?(?:C\.\s*(.*?))?(?:D\.\s*(.*?))?'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if match:
            groups = match.groups()
            
            result = {
                'number': groups[0].replace('Câu', '').strip() if groups[0] else '',
                'question': groups[1].strip() if len(groups) > 1 and groups[1] else '',
                'A': groups[2].strip() if len(groups) > 2 and groups[2] else '',
                'B': groups[3].strip() if len(groups) > 3 and groups[3] else '',
                'C': groups[4].strip() if len(groups) > 4 and groups[4] else '',
                'D': groups[5].strip() if len(groups) > 5 and groups[5] else ''
            }
            
            # Kiểm tra có đủ thông tin không
            if result['question'] and any([result['A'], result['B'], result['C'], result['D']]):
                return result
    
    # Fallback: Tách thủ công
    lines = text.split('\n')
    result = {'number': '', 'question': '', 'A': '', 'B': '', 'C': '', 'D': ''}
    
    current_key = 'question'
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Tìm số câu hỏi
        if re.match(r'Câu\s*\d+', line, re.IGNORECASE):
            result['number'] = re.search(r'\d+', line).group()
            result['question'] = re.sub(r'Câu\s*\d+\.?\s*', '', line, flags=re.IGNORECASE)
            continue
        
        # Tìm đáp án
        for option in ['A', 'B', 'C', 'D']:
            if line.startswith(f'{option}.'):
                result[option] = line[2:].strip()
                current_key = option
                break
        else:
            # Nếu không phải đáp án mới, nối vào key hiện tại
            if current_key and result[current_key]:
                result[current_key] += ' ' + line
            elif current_key == 'question':
                result['question'] += ' ' + line
    
    return result if result['question'] else None

def extract_true_false_question(text):
    """Trích xuất câu hỏi đúng/sai"""
    
    result = {'number': '', 'question': '', 'type': 'true_false'}
    
    # Tìm số câu hỏi
    number_match = re.search(r'Câu\s*(\d+)', text, re.IGNORECASE)
    if number_match:
        result['number'] = number_match.group(1)
    
    # Lấy nội dung câu hỏi
    question_text = re.sub(r'Câu\s*\d+\.?\s*', '', text, flags=re.IGNORECASE)
    result['question'] = question_text.strip()
    
    return result

def extract_essay_question(text):
    """Trích xuất câu hỏi tự luận"""
    
    result = {'number': '', 'question': '', 'type': 'essay'}
    
    # Tìm số câu hỏi
    number_match = re.search(r'Câu\s*(\d+)', text, re.IGNORECASE)
    if number_match:
        result['number'] = number_match.group(1)
    
    # Lấy nội dung câu hỏi
    question_text = re.sub(r'Câu\s*\d+\.?\s*', '', text, flags=re.IGNORECASE)
    result['question'] = question_text.strip()
    
    return result

def create_professional_table(doc, title, headers, data):
    """Tạo bảng Word chuyên nghiệp với style đẹp"""
    
    # Thêm tiêu đề bảng
    title_para = doc.add_paragraph()
    title_run = title_para.add_run(title)
    title_run.font.size = Pt(14)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(31, 78, 121)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.space_after = Pt(12)
    
    # Tạo bảng
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Thiết lập độ rộng cột
    if len(headers) == 6:  # Bảng trắc nghiệm
        widths = [Inches(0.5), Inches(3.0), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5)]
    elif len(headers) == 3:  # Bảng đúng/sai hoặc tự luận
        widths = [Inches(0.5), Inches(4.5), Inches(1.5)]
    else:
        widths = [Inches(1.5)] * len(headers)
    
    for i, width in enumerate(widths[:len(headers)]):
        table.columns[i].width = width
    
    # Định dạng header
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = header_cells[i]
        cell.text = header
        
        # Định dạng text trong cell
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(11)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Màu nền header
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shading_elm = parse_xml(r'<w:shd {} w:fill="1f4e79"/>'.format(
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'))
        cell._tc.get_or_add_tcPr().append(shading_elm)
    
    # Thêm dữ liệu
    for row_data in data:
        row_cells = table.add_row().cells
        
        for i, cell_data in enumerate(row_data):
            if i < len(row_cells):
                cell = row_cells[i]
                cell.text = str(cell_data) if cell_data else ''
                
                # Định dạng text
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
                    
                    if i == 0:  # Cột STT
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    
    # Thêm border cho bảng
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            
            # Thêm border
            borders = parse_xml(r'''
                <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="1f4e79"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="1f4e79"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="1f4e79"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="1f4e79"/>
                </w:tcBorders>
            ''')
            tcPr.append(borders)
    
    # Thêm khoảng cách sau bảng
    doc.add_paragraph().space_after = Pt(18)
    
    return table

def process_docx(uploaded_file):
    """Xử lý file Word đề thi thành bảng chuẩn hóa"""
    
    # Đọc document gốc
    doc = Document(uploaded_file)
    
    # Tạo document mới
    new_doc = Document()
    
    # Thiết lập margins
    sections = new_doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Thêm header chính
    header = new_doc.add_heading('ĐỀ THI ĐÃ CHUẨN HÓA', level=1)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in header.runs:
        run.font.color.rgb = RGBColor(31, 78, 121)
        run.font.size = Pt(16)
    
    # Thêm thông tin ngày tạo
    date_para = new_doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in date_para.runs:
        run.font.italic = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(128, 128, 128)
    
    new_doc.add_paragraph()  # Khoảng trống
    
    # Phân tích cấu trúc
    analysis = analyze_exam_structure(doc)
    
    # Thu thập tất cả câu hỏi
    all_questions = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if re.match(r'Câu\s*\d+', text, re.IGNORECASE):
            question_type = classify_question_type(text)
            parsed = parse_question(text, question_type)
            if parsed:
                all_questions.append(parsed)
    
    # Phân loại câu hỏi
    multiple_choice_questions = [q for q in all_questions if q.get('type') == 'multiple_choice']
    true_false_questions = [q for q in all_questions if q.get('type') == 'true_false']
    essay_questions = [q for q in all_questions if q.get('type') == 'essay']
    
    # Tạo bảng cho câu hỏi trắc nghiệm
    if multiple_choice_questions:
        mc_headers = ['STT', 'Câu hỏi', 'A', 'B', 'C', 'D']
        mc_data = []
        
        for i, q in enumerate(multiple_choice_questions, 1):
            mc_data.append([
                str(i),
                q.get('question', ''),
                q.get('A', ''),
                q.get('B', ''),
                q.get('C', ''),
                q.get('D', '')
            ])
        
        create_professional_table(new_doc, "PHẦN I: CÂU HỎI TRẮC NGHIỆM", mc_headers, mc_data)
    
    # Tạo bảng cho câu hỏi đúng/sai
    if true_false_questions:
        tf_headers = ['STT', 'Câu hỏi', 'Đáp án']
        tf_data = []
        
        for i, q in enumerate(true_false_questions, 1):
            tf_data.append([
                str(i),
                q.get('question', ''),
                ''  # Để trống cho giáo viên điền
            ])
        
        create_professional_table(new_doc, "PHẦN II: CÂU HỎI ĐÚNG/SAI", tf_headers, tf_data)
    
    # Tạo bảng cho câu hỏi tự luận
    if essay_questions:
        essay_headers = ['STT', 'Câu hỏi', 'Ghi chú']
        essay_data = []
        
        for i, q in enumerate(essay_questions, 1):
            essay_data.append([
                str(i),
                q.get('question', ''),
                ''  # Để trống cho ghi chú
            ])
        
        create_professional_table(new_doc, "PHẦN III: CÂU HỎI TỰ LUẬN", essay_headers, essay_data)
    
    # Thống kê
    stats = {
        'total_processed': len(all_questions),
        'multiple_choice_processed': len(multiple_choice_questions),
        'true_false_processed': len(true_false_questions),
        'essay_processed': len(essay_questions),
        'processing_details': [
            f"Đã phân tích {len(all_questions)} câu hỏi tổng cộng",
            f"Trắc nghiệm: {len(multiple_choice_questions)} câu",
            f"Đúng/Sai: {len(true_false_questions)} câu",
            f"Tự luận: {len(essay_questions)} câu"
        ]
    }
    
    return new_doc, stats

# Import datetime if not imported
from datetime import datetime
