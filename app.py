#!/usr/bin/env python3
"""
🧮 Ứng dụng chuyển LaTeX và Markdown - Phiên bản ổn định
Đảm bảo file Word có thể tải về và mở được
"""

import streamlit as st
import sys
import os
from processor_simple import convert_docx_simple, detect_latex_patterns, detect_markdown_tables
from docx import Document
from io import BytesIO
import traceback
from datetime import datetime

# Cấu hình trang
st.set_page_config(
    page_title="🧮 Chuyển LaTeX và Markdown (Ổn định)",
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
    .warning-box {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        margin: 1rem 0;
    }
    .stable-version {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

def analyze_content_simple(doc):
    """Phân tích nội dung document đơn giản"""
    
    analysis = {
        'total_paragraphs': 0,
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
        
        # Tìm bảng Markdown
        tables = detect_markdown_tables(text)
        if tables:
            analysis['table_count'] += len(tables)
        
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
        <h1>🧮 Chuyển LaTeX và Markdown</h1>
        <p>Phiên bản ổn định - Đảm bảo file Word có thể tải về và mở được</p>
        <p><strong>📝 LaTeX → Text có format | 📊 Markdown → Table Word</strong></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar thông tin
    with st.sidebar:
        st.markdown("## 🎯 Phiên bản ổn định")
        
        st.markdown("""
        <div class="stable-version">
        ✅ 100% tương thích<br>
        ✅ File Word mở được<br>
        ✅ Không cần Win32COM<br>
        ✅ Hoạt động mọi platform
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 🧮 LaTeX được chuyển thành:")
        st.markdown("""
        - **$x^2$** → **[x^(2)]** (italic, màu xanh)
        - **$\\frac{a}{b}$** → **[(a)/(b)]** 
        - **$\\alpha + \\beta$** → **[α + β]**
        - **$\\int_0^1 f(x)dx$** → **[∫_0^1 f(x)dx]**
        """)
        
        st.markdown("### 📊 Markdown Tables:")
        st.code("""
| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |
| $x^2$    | $y^2$    |
        """)
        
        st.markdown("### ✨ Ví dụ:")
        st.code("""
Input: $\\frac{a}{b} = \\frac{c}{d}$
Output: [(a)/(b) = (c)/(d)]

Input: $\\alpha + \\beta = \\pi$  
Output: [α + β = π]
        """)
        
        st.markdown("---")
        st.markdown("### 💡 Ưu điểm:")
        st.markdown("""
        - ✅ **File Word mở được 100%**
        - ✅ **Không cần thư viện phức tạp**  
        - ✅ **Hoạt động trên mọi hệ thống**
        - ✅ **Xử lý được LaTeX phức tạp**
        - ✅ **Bảng chuyên nghiệp**
        """)
        
        st.markdown("### 🔧 Sau khi tải xuống:")
        st.markdown("""
        - 📝 **Mở file trong Word**
        - ✏️ **Chỉnh sửa công thức thủ công nếu cần**
        - 📊 **Bảng đã format sẵn**
        - 🎨 **Có thể thay đổi style**
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
                    analysis = analyze_content_simple(doc)
                    
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
                                    <p>Chuyển thành text có format</p>
                                </div>
                                """, unsafe_allow_html=True)
                        
                        with col_table:
                            if analysis['table_count'] > 0:
                                st.markdown(f"""
                                <div class="table-highlight">
                                    <h4>📊 Bảng Markdown</h4>
                                    <p><strong>Số lượng:</strong> {analysis['table_count']}</p>
                                    <p>Chuyển thành Table Word</p>
                                </div>
                                """, unsafe_allow_html=True)
                    
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
            ("2️⃣", "Chuyển LaTeX", "$ → text có format"),
            ("3️⃣", "Tạo bảng Word", "| | → Table chuyên nghiệp"),
            ("4️⃣", "Xuất file", "Đảm bảo mở được"),
            ("5️⃣", "Tải xuống", "File Word hoàn chỉnh")
        ]
        
        for step, title, desc in process_steps:
            st.markdown(f"""
            <div class="feature-card">
                <strong>{step} {title}</strong><br>
                <small>{desc}</small>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("## 💡 Phương pháp")
        
        st.markdown("""
        <div class="stable-version">
        <strong>📝 LaTeX Processing:</strong><br>
        • Phát hiện tự động<br>
        • Chuyển thành Unicode<br>
        • Format italic + màu sắc<br>
        • Dễ đọc và chỉnh sửa<br><br>
        
        <strong>📊 Table Processing:</strong><br>
        • Nhận dạng Markdown<br>  
        • Tạo bảng Word native<br>
        • Header có màu nền<br>
        • Border chuyên nghiệp
        </div>
        """, unsafe_allow_html=True)
    
    # Xử lý chuyển đổi
    if uploaded_file:
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            if st.button("🚀 Chuyển đổi LaTeX và Markdown", type="primary", use_container_width=True):
                with st.spinner("⏳ Đang xử lý..."):
                    try:
                        # Xử lý chuyển đổi
                        st.info("🔄 Đang phân tích và chuyển đổi...")
                        output_doc, stats = convert_docx_simple(uploaded_file)
                        
                        # Tạo buffer an toàn
                        output_stream = BytesIO()
                        output_doc.save(output_stream)
                        output_stream.seek(0)
                        
                        # Kiểm tra file có được tạo không
                        file_size = len(output_stream.getvalue())
                        if file_size < 1000:  # File quá nhỏ, có thể bị lỗi
                            st.error("❌ File được tạo quá nhỏ, có thể bị lỗi")
                            return
                        
                        # Tạo tên file
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_filename = f"converted_stable_{timestamp}.docx"
                        
                        # Thông báo thành công
                        st.markdown("""
                        <div class="success-box">
                            <h3>🎉 Chuyển đổi thành công!</h3>
                            <p>File Word đã sẵn sàng để tải xuống và mở trong Microsoft Word</p>
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
                            file_size_mb = round(file_size / 1024 / 1024, 2)
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #28a745;">{file_size_mb}</h2>
                                <p>MB</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        # Chi tiết kết quả
                        if stats['latex_count'] > 0:
                            st.success(f"🧮 Đã chuyển {stats['latex_count']} công thức LaTeX thành text có format")
                        
                        if stats['table_count'] > 0:
                            st.success(f"📊 Đã tạo {stats['table_count']} bảng Word chuyên nghiệp")
                        
                        # Thông tin file
                        st.info(f"📄 Kích thước file: {file_size_mb} MB - Đảm bảo mở được trong Word")
                        
                        # Nút tải xuống
                        st.download_button(
                            label="📥 Tải xuống file Word (Ổn định 100%)",
                            data=output_stream.getvalue(),
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            type="primary"
                        )
                        
                        # Hướng dẫn sau khi tải
                        st.markdown("### 💡 Sau khi tải xuống:")
                        st.markdown("""
                        1. ✅ **Mở file trong Microsoft Word** - File sẽ mở được bình thường
                        2. 📝 **LaTeX hiển thị dạng [...]** - Có format italic và màu xanh  
                        3. 📊 **Bảng có header màu xanh** - Đã format chuyên nghiệp
                        4. ✏️ **Có thể chỉnh sửa bình thường** - Copy/paste, format như Word thường
                        5. 🔄 **Chuyển [LaTeX] thành Equation** - Thủ công nếu cần thiết
                        """)
                        
                        # Lưu ý
                        if 'error' in stats:
                            st.warning(f"⚠️ Có một số lỗi nhỏ: {stats['error']}")
                            
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
                        - Kiểm tra file Word không bị khóa hoặc corrupt
                        - Thử với file nhỏ hơn trước  
                        - Đảm bảo có quyền ghi trong thư mục download
                        - Refresh browser và thử lại
                        """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 2rem;">
        <p>🧮 <strong>Ứng dụng chuyển LaTeX và Markdown - Phiên bản ổn định</strong></p>
        <p>✅ <strong>Đảm bảo file Word mở được 100%</strong></p>
        <p>Phát triển với ❤️ bằng Streamlit và python-docx</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
