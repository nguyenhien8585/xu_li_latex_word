#!/usr/bin/env python3
"""
🧠 Ứng dụng chuyển đề thi Word thành bảng chuẩn hóa
Streamlit App - Phiên bản đơn giản
"""

import streamlit as st
import sys
import os
from processor import process_docx, analyze_exam_structure
from docx import Document
from io import BytesIO
import traceback
from datetime import datetime

# Cấu hình trang
st.set_page_config(
    page_title="🧠 Chuyển đề thi Word thành bảng chuẩn",
    page_icon="🧠",
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
    .info-box {
        background: #e3f2fd;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #2196f3;
        margin: 1rem 0;
        color: #1565c0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header chính
    st.markdown("""
    <div class="main-header">
        <h1>🧠 Chuyển đề thi Word thành bảng chuẩn</h1>
        <p>Tự động nhận dạng cấu trúc đề thi và chuẩn hóa thành format bảng chuyên nghiệp</p>
        <p><strong>📝 Phiên bản đơn giản - Hoạt động ổn định</strong></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar thông tin
    with st.sidebar:
        st.markdown("## 📋 Hướng dẫn sử dụng")
        
        st.markdown("""
        ### 🎯 Ứng dụng hỗ trợ:
        - **Phần I**: Câu hỏi trắc nghiệm (A, B, C, D)
        - **Phần II**: Câu hỏi đúng/sai
        - **Phần III**: Câu hỏi tự luận
        """)
        
        st.markdown("### 📝 Format nhận dạng:")
        st.code("""
Câu 1. Nội dung câu hỏi...
A. Đáp án A
B. Đáp án B  
C. Đáp án C
D. Đáp án D
        """)
        
        st.markdown("### ✨ Tính năng:")
        features = [
            "🔍 Tự động nhận dạng cấu trúc",
            "📊 Tạo bảng Word chuẩn hóa", 
            "🎨 Format chuyên nghiệp",
            "📥 Tải xuống ngay lập tức",
            "⚡ Xử lý nhanh chóng",
            "🧮 Phát hiện công thức LaTeX",
            "📊 Phát hiện bảng Markdown"
        ]
        
        for feature in features:
            st.markdown(f"- {feature}")
        
        st.markdown("---")
        st.markdown("### 📞 Hỗ trợ")
        st.info("Phiên bản đơn giản này ổn định và dễ cài đặt.")
        
        st.markdown("### 🛠️ Cài đặt:")
        st.code("""
pip install streamlit python-docx regex
streamlit run app.py
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("## 📁 Tải lên đề thi")
        
        uploaded_file = st.file_uploader(
            "Chọn file .docx đề thi cần chuẩn hóa",
            type=["docx"],
            help="File Word chứa đề thi với cấu trúc câu hỏi trắc nghiệm, đúng/sai, tự luận"
        )
        
        if uploaded_file:
            st.success(f"✅ Đã tải lên: **{uploaded_file.name}**")
            
            # Phân tích cấu trúc đề thi
            with st.expander("🔍 Phân tích cấu trúc đề thi", expanded=True):
                try:
                    doc = Document(uploaded_file)
                    analysis = analyze_exam_structure(doc)
                    
                    st.markdown(f"""
                    <div class="analysis-result">
                        <h3>📊 Kết quả phân tích:</h3>
                        <p><strong>Tổng số câu hỏi:</strong> {analysis['total_questions']}</p>
                        <p><strong>Câu trắc nghiệm:</strong> {analysis['multiple_choice']}</p>
                        <p><strong>Câu đúng/sai:</strong> {analysis['true_false']}</p>
                        <p><strong>Câu tự luận:</strong> {analysis['essay']}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Hiển thị thông tin về LaTeX và Markdown
                    if analysis.get('has_math_formulas', False) or analysis.get('has_markdown_tables', False):
                        col_math, col_table = st.columns(2)
                        
                        with col_math:
                            if analysis.get('has_math_formulas', False):
                                st.markdown(f"""
                                <div class="info-box">
                                    <h4>🧮 Công thức LaTeX</h4>
                                    <p>Phát hiện: {analysis.get('math_count', 0)} công thức</p>
                                    <p>Sẽ được giữ nguyên trong bảng</p>
                                </div>
                                """, unsafe_allow_html=True)
                        
                        with col_table:
                            if analysis.get('has_markdown_tables', False):
                                st.markdown(f"""
                                <div class="info-box">
                                    <h4>📊 Bảng Markdown</h4>
                                    <p>Phát hiện: {analysis.get('table_count', 0)} bảng</p>
                                    <p>Sẽ được mô tả trong text</p>
                                </div>
                                """, unsafe_allow_html=True)
                    
                    # Hiển thị preview một số câu hỏi
                    if analysis['sample_questions']:
                        st.markdown("### 👀 Preview câu hỏi:")
                        for i, q in enumerate(analysis['sample_questions'][:3]):
                            with st.expander(f"Câu {i+1}: {q['number']}"):
                                st.write(f"**Câu hỏi:** {q['question'][:100]}...")
                                if q['type'] == 'multiple_choice':
                                    st.write(f"**A:** {q.get('A', 'N/A')[:50]}...")
                                    st.write(f"**B:** {q.get('B', 'N/A')[:50]}...")
                                    st.write(f"**C:** {q.get('C', 'N/A')[:50]}...")
                                    st.write(f"**D:** {q.get('D', 'N/A')[:50]}...")
                                
                except Exception as e:
                    st.error(f"Lỗi phân tích: {str(e)}")
                    st.code(traceback.format_exc())
    
    with col2:
        st.markdown("## 🎯 Quy trình xử lý")
        
        process_steps = [
            ("1️⃣", "Tải lên đề thi", "Upload file .docx"),
            ("2️⃣", "Phân tích cấu trúc", "Nhận dạng câu hỏi"),
            ("3️⃣", "Chuẩn hóa format", "Tạo bảng Word"),
            ("4️⃣", "Tải xuống kết quả", "File Word hoàn chỉnh")
        ]
        
        for step, title, desc in process_steps:
            st.markdown(f"""
            <div class="feature-card">
                <strong>{step} {title}</strong><br>
                <small>{desc}</small>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("## 💡 Lưu ý")
        st.markdown("""
        <div class="info-box">
        <strong>Phiên bản đơn giản này:</strong><br>
        • Ổn định và dễ cài đặt<br>
        • Hỗ trợ đầy đủ tính năng cơ bản<br>
        • Phát hiện LaTeX và Markdown<br>
        • Không cần thư viện phức tạp
        </div>
        """, unsafe_allow_html=True)
    
    # Xử lý chuyển đổi
    if uploaded_file:
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            if st.button("🚀 Chuyển đổi thành bảng chuẩn hóa", type="primary", use_container_width=True):
                with st.spinner("⏳ Đang xử lý đề thi..."):
                    try:
                        # Xử lý file
                        output_doc, stats = process_docx(uploaded_file)
                        
                        # Tạo buffer để tải xuống
                        output_stream = BytesIO()
                        output_doc.save(output_stream)
                        output_stream.seek(0)
                        
                        # Tạo tên file
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_filename = f"de_chuan_hoa_{timestamp}.docx"
                        
                        # Thông báo thành công
                        st.markdown("""
                        <div class="success-box">
                            <h3>🎉 Chuyển đổi thành công!</h3>
                            <p>Đề thi đã được chuẩn hóa thành format bảng chuyên nghiệp</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Thống kê kết quả
                        st.markdown("## 📊 Thống kê kết quả")
                        
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #667eea;">{stats['total_processed']}</h2>
                                <p>Câu đã xử lý</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col2:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #28a745;">{stats['multiple_choice_processed']}</h2>
                                <p>Trắc nghiệm</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col3:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #ffc107;">{stats['true_false_processed']}</h2>
                                <p>Đúng/Sai</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col4:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #17a2b8;">{stats['essay_processed']}</h2>
                                <p>Tự luận</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        # Hiển thị thông tin về LaTeX và Markdown nếu có
                        if stats.get('math_formulas', 0) > 0 or stats.get('markdown_tables', 0) > 0:
                            st.markdown("### 🔍 Phát hiện thêm:")
                            enhanced_info = []
                            if stats.get('math_formulas', 0) > 0:
                                enhanced_info.append(f"🧮 {stats['math_formulas']} công thức LaTeX")
                            if stats.get('markdown_tables', 0) > 0:
                                enhanced_info.append(f"📊 {stats['markdown_tables']} bảng Markdown")
                            
                            for info in enhanced_info:
                                st.info(info)
                        
                        # Nút tải xuống
                        st.download_button(
                            label="📥 Tải xuống file Word đã chuẩn hóa",
                            data=output_stream.getvalue(),
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            type="primary"
                        )
                        
                        # Chi tiết xử lý
                        if stats.get('processing_details'):
                            with st.expander("📋 Chi tiết xử lý"):
                                for detail in stats['processing_details']:
                                    st.text(detail)
                        
                        # Thông báo thành công
                        success_messages = []
                        if stats['multiple_choice_processed'] > 0:
                            success_messages.append(f"✅ Đã xử lý {stats['multiple_choice_processed']} câu trắc nghiệm")
                        if stats['true_false_processed'] > 0:
                            success_messages.append(f"✅ Đã xử lý {stats['true_false_processed']} câu đúng/sai")
                        if stats['essay_processed'] > 0:
                            success_messages.append(f"✅ Đã xử lý {stats['essay_processed']} câu tự luận")
                        
                        for msg in success_messages:
                            st.success(msg)
                            
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
                        - Kiểm tra format câu hỏi có đúng không
                        - Đảm bảo file Word không bị lỗi
                        - Thử lại với file khác
                        - Sử dụng file test mẫu để kiểm tra
                        """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 2rem;">
        <p>🧠 <strong>Ứng dụng chuyển đề thi Word thành bảng chuẩn</strong></p>
        <p>📝 <strong>Phiên bản đơn giản - Ổn định và dễ sử dụng</strong></p>
        <p>Phát triển với ❤️ bằng Streamlit và Python</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
