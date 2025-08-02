#!/usr/bin/env python3
"""
🧮 Ứng dụng chuyển LaTeX và Markdown trong Word
Streamlit App - Tự động chuyển $...$ thành Equation và bảng Markdown thành Table
"""

import streamlit as st
import sys
import os
from processor import (
    convert_docx_with_latex_and_tables, 
    post_process_with_word_com,
    detect_latex_patterns,
    detect_markdown_tables,
    HAS_WIN32COM
)
from docx import Document
from io import BytesIO
import traceback
from datetime import datetime

# Cấu hình trang
st.set_page_config(
    page_title="🧮 Chuyển LaTeX và Markdown trong Word",
    page_icon="🧮",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS tùy chỉnh
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2.5rem;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 2rem;
        color: white;
        box-shadow: 0 10px 30px rgba(0,0,0,0.3);
    }
    .main-header h1 {
        font-size: 2.8rem;
        font-weight: bold;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    .feature-card {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #667eea;
        margin: 1rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .analysis-result {
        background: linear-gradient(135deg, #56ab2f 0%, #a8e6cf 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        margin: 1rem 0;
        box-shadow: 0 6px 20px rgba(86, 171, 47, 0.3);
    }
    .latex-highlight {
        background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: #333;
        margin: 1rem 0;
        box-shadow: 0 6px 20px rgba(255, 154, 158, 0.3);
    }
    .table-highlight {
        background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: #333;
        margin: 1rem 0;
        box-shadow: 0 6px 20px rgba(168, 237, 234, 0.3);
    }
    .warning-box {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        margin: 1rem 0;
    }
    .success-box {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin: 1rem 0;
    }
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        text-align: center;
        border-top: 4px solid #667eea;
    }
    .win32-available {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        font-weight: bold;
    }
    .win32-unavailable {
        background: linear-gradient(135deg, #ffeaa7 0%, #fab1a0 100%);
        color: #2d3436;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        font-weight: bold;
    }
    .code-example {
        background: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 5px;
        padding: 1rem;
        font-family: 'Courier New', monospace;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

def analyze_content(doc):
    """Phân tích nội dung document để tìm LaTeX và Markdown"""
    
    analysis = {
        'total_paragraphs': 0,
        'latex_formulas': [],
        'markdown_tables': [],
        'latex_count': 0,
        'table_count': 0,
        'preview_content': []
    }
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        analysis['total_paragraphs'] += 1
        
        # Tìm LaTeX
        formulas = detect_latex_patterns(text)
        if formulas:
            analysis['latex_count'] += len(formulas)
            analysis['latex_formulas'].extend(formulas)
        
        # Tìm bảng Markdown
        tables = detect_markdown_tables(text)
        if tables:
            analysis['table_count'] += len(tables)
            analysis['markdown_tables'].extend(tables)
        
        # Lưu preview (3 đoạn đầu)
        if len(analysis['preview_content']) < 3:
            analysis['preview_content'].append({
                'text': text[:200] + "..." if len(text) > 200 else text,
                'has_latex': len(formulas) > 0,
                'has_tables': len(tables) > 0,
                'latex_count': len(formulas),
                'table_count': len(tables)
            })
    
    return analysis

def main():
    # Header chính
    st.markdown("""
    <div class="main-header">
        <h1>🧮 Chuyển LaTeX và Markdown trong Word</h1>
        <p>Tự động chuyển công thức $...$ thành Equation và bảng Markdown thành Table thực sự</p>
        <p><strong>⚡ Hỗ trợ Win32COM để tạo Equation chuẩn Word!</strong></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar thông tin
    with st.sidebar:
        st.markdown("## 🎯 Tính năng chính")
        
        # Hiển thị trạng thái Win32COM
        if HAS_WIN32COM:
            st.markdown("""
            <div class="win32-available">
            ✅ Win32COM khả dụng<br>
            Sẽ tạo Equation thực sự trong Word
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="win32-unavailable">
            ⚠️ Win32COM không khả dụng<br>
            LaTeX sẽ hiển thị dạng italic
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("### 🧮 LaTeX được hỗ trợ:")
        st.markdown("""
        - **Inline**: `$x^2 + y^2 = r^2$`
        - **Brackets**: `${a \\over b}$`  
        - **Parentheses**: `\\(\\sin x\\)`
        - **Display**: `\\[\\int_0^1 f(x)dx\\]`
        """)
        
        st.markdown("### 📊 Markdown Tables:")
        st.code("""
| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |
| Cell 3   | Cell 4   |
        """)
        
        st.markdown("### ✨ Ví dụ kết hợp:")
        st.code("""
Tính tích phân $\\int_0^1 x^2 dx$:

| Phương pháp | Kết quả |
|-------------|---------|
| Trực tiếp   | $\\frac{1}{3}$ |
| Số học      | 0.333... |
        """)
        
        st.markdown("---")
        st.markdown("### 🛠️ Yêu cầu hệ thống:")
        st.markdown("""
        - **Windows**: Để dùng Win32COM
        - **Microsoft Word**: Để tạo Equation
        - **Python 3.7+**: Môi trường chạy
        """)
        
        st.markdown("### 📦 Thư viện cần thiết:")
        st.code("""
pip install streamlit python-docx
pip install pywin32  # Windows only
pip install sympy latex2mathml  # Optional
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("## 📁 Tải lên file Word")
        
        uploaded_file = st.file_uploader(
            "Chọn file .docx chứa LaTeX và/hoặc bảng Markdown",
            type=["docx"],
            help="File Word có thể chứa công thức $...$ và bảng |...|..."
        )
        
        if uploaded_file:
            st.success(f"✅ Đã tải lên: **{uploaded_file.name}**")
            
            # Phân tích nội dung
            with st.expander("🔍 Phân tích nội dung", expanded=True):
                try:
                    doc = Document(uploaded_file)
                    analysis = analyze_content(doc)
                    
                    # Thống kê tổng quan
                    st.markdown(f"""
                    <div class="analysis-result">
                        <h3>📊 Tổng quan:</h3>
                        <p><strong>Tổng đoạn văn:</strong> {analysis['total_paragraphs']}</p>
                        <p><strong>Công thức LaTeX:</strong> {analysis['latex_count']}</p>
                        <p><strong>Bảng Markdown:</strong> {analysis['table_count']}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Chi tiết LaTeX và Tables
                    if analysis['latex_count'] > 0 or analysis['table_count'] > 0:
                        col_latex, col_table = st.columns(2)
                        
                        with col_latex:
                            if analysis['latex_count'] > 0:
                                st.markdown(f"""
                                <div class="latex-highlight">
                                    <h4>🧮 Công thức LaTeX</h4>
                                    <p><strong>Số lượng:</strong> {analysis['latex_count']}</p>
                                    <p>Sẽ chuyển thành Equation Word</p>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                # Hiển thị mẫu công thức
                                if analysis['latex_formulas']:
                                    st.markdown("**Mẫu công thức:**")
                                    for i, formula in enumerate(analysis['latex_formulas'][:3]):
                                        st.code(f"${formula['latex']}$")
                        
                        with col_table:
                            if analysis['table_count'] > 0:
                                st.markdown(f"""
                                <div class="table-highlight">
                                    <h4>📊 Bảng Markdown</h4>
                                    <p><strong>Số lượng:</strong> {analysis['table_count']}</p>
                                    <p>Sẽ chuyển thành Table Word</p>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                # Hiển thị mẫu bảng
                                if analysis['markdown_tables']:
                                    st.markdown("**Mẫu bảng:**")
                                    table = analysis['markdown_tables'][0]
                                    st.write(f"Headers: {', '.join(table['header'])}")
                                    st.write(f"Rows: {len(table['rows'])}")
                    
                    # Preview nội dung
                    if analysis['preview_content']:
                        st.markdown("### 👀 Preview nội dung:")
                        for i, preview in enumerate(analysis['preview_content']):
                            with st.expander(f"Đoạn {i+1}"):
                                st.write(preview['text'])
                                
                                features = []
                                if preview['has_latex']:
                                    features.append(f"🧮 {preview['latex_count']} LaTeX")
                                if preview['has_tables']:
                                    features.append(f"📊 {preview['table_count']} Table")
                                
                                if features:
                                    st.write(f"**Đặc điểm:** {', '.join(features)}")
                    
                except Exception as e:
                    st.error(f"Lỗi phân tích: {str(e)}")
                    st.code(traceback.format_exc())
    
    with col2:
        st.markdown("## ⚙️ Quy trình xử lý")
        
        process_steps = [
            ("1️⃣", "Phân tích nội dung", "Tìm LaTeX và Markdown"),
            ("2️⃣", "Tạo document mới", "Chuyển đổi cấu trúc"),
            ("3️⃣", "Xử lý LaTeX", "$ → Equation hoặc italic"),
            ("4️⃣", "Xử lý bảng", "| | → Table Word"),
            ("5️⃣", "Win32COM (nếu có)", "Tạo Equation thực sự"),
            ("6️⃣", "Xuất kết quả", "File Word hoàn chỉnh")
        ]
        
        for step, title, desc in process_steps:
            st.markdown(f"""
            <div class="feature-card">
                <strong>{step} {title}</strong><br>
                <small>{desc}</small>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("## 💡 Lưu ý")
        
        if HAS_WIN32COM:
            st.markdown("""
            <div class="win32-available">
            ✅ <strong>Equation thực sự</strong><br>
            Công thức sẽ được tạo dạng native Word Equation
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="win32-unavailable">
            ⚠️ <strong>Fallback mode</strong><br>
            LaTeX sẽ hiển thị dạng italic. Cài pywin32 để có Equation thực sự.
            </div>
            """, unsafe_allow_html=True)
    
    # Xử lý chuyển đổi
    if uploaded_file:
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            if st.button("🚀 Chuyển đổi LaTeX và Markdown", type="primary", use_container_width=True):
                with st.spinner("⏳ Đang xử lý LaTeX và Markdown..."):
                    try:
                        # Bước 1: Chuyển đổi cơ bản
                        st.info("🔄 Đang xử lý LaTeX và tạo bảng...")
                        output_doc, stats = convert_docx_with_latex_and_tables(uploaded_file)
                        
                        # Tạo buffer
                        output_stream = BytesIO()
                        output_doc.save(output_stream)
                        output_stream.seek(0)
                        
                        # Bước 2: Xử lý với Win32COM (nếu có)
                        if HAS_WIN32COM and stats['latex_count'] > 0:
                            st.info("⚡ Đang tạo Equation thực sự với Win32COM...")
                            processed_stream, com_success = post_process_with_word_com(output_stream)
                            
                            if com_success:
                                output_stream = processed_stream
                                st.success("✅ Đã tạo Equation thực sự!")
                            else:
                                st.warning("⚠️ Không thể tạo Equation, giữ format italic")
                        
                        # Tạo tên file
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_filename = f"latex_markdown_converted_{timestamp}.docx"
                        
                        # Thông báo thành công
                        st.markdown("""
                        <div class="success-box">
                            <h3>🎉 Chuyển đổi hoàn tất!</h3>
                            <p>LaTeX và Markdown đã được chuyển thành format Word chuẩn</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Thống kê kết quả
                        st.markdown("## 📊 Kết quả chuyển đổi")
                        
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #667eea;">{stats['processed_paragraphs']}</h2>
                                <p>Đoạn văn</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col2:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #e83e8c;">{stats['latex_count']}</h2>
                                <p>Công thức LaTeX</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col3:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #17a2b8;">{stats['table_count']}</h2>
                                <p>Bảng Markdown</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col4:
                            total_converted = stats['latex_count'] + stats['table_count']
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #28a745;">{total_converted}</h2>
                                <p>Tổng chuyển đổi</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        # Chi tiết kết quả
                        if stats['latex_count'] > 0:
                            if HAS_WIN32COM:
                                st.success(f"🧮 Đã chuyển {stats['latex_count']} công thức LaTeX thành Equation Word")
                            else:
                                st.info(f"🧮 Đã format {stats['latex_count']} công thức LaTeX dạng italic")
                        
                        if stats['table_count'] > 0:
                            st.success(f"📊 Đã chuyển {stats['table_count']} bảng Markdown thành Table Word")
                        
                        # Nút tải xuống
                        st.download_button(
                            label="📥 Tải xuống file Word đã chuyển đổi",
                            data=output_stream.getvalue(),
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            type="primary"
                        )
                        
                        # Hướng dẫn mở file
                        st.markdown("### 💡 Sau khi tải xuống:")
                        if HAS_WIN32COM and stats['latex_count'] > 0:
                            st.markdown("""
                            - ✅ **Công thức Equation**: Hiển thị chuẩn trong Word
                            - ✅ **Bảng**: Format chuyên nghiệp với border
                            - 🔧 **Chỉnh sửa**: Click đúp vào công thức để edit
                            """)
                        else:
                            st.markdown("""
                            - 📝 **LaTeX dạng italic**: Có thể chỉnh sửa thành Equation thủ công
                            - ✅ **Bảng**: Format chuyên nghiệp với border  
                            - 💡 **Tip**: Cài pywin32 để có Equation tự động
                            """)
                            
                    except Exception as e:
                        st.markdown("""
                        <div class="warning-box">
                            <h3>❌ Lỗi xử lý</h3>
                            <p>Có lỗi xảy ra trong quá trình chuyển đổi</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.error(f"Chi tiết lỗi: {str(e)}")
                        
                        with st.expander("🔍 Debug thông tin"):
                            st.code(traceback.format_exc())
                        
                        st.markdown("### 💡 Gợi ý khắc phục:")
                        st.markdown("""
                        - Kiểm tra file Word không bị khóa
                        - Đảm bảo công thức LaTeX đúng syntax
                        - Bảng Markdown cần có header separator (|---|)  
                        - Thử tắt antivirus tạm thời (cho Win32COM)
                        - Chạy với quyền Administrator (Windows)
                        """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 2rem;">
        <p>🧮 <strong>Ứng dụng chuyển LaTeX và Markdown trong Word</strong></p>
        <p>⚡ <strong>Hỗ trợ Win32COM để tạo Equation thực sự</strong></p>
        <p>Phát triển với ❤️ bằng Streamlit, python-docx và Win32COM</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
