#!/usr/bin/env python3
"""
🧠 Ứng dụng chuyển đề thi Word thành bảng chuẩn hóa
Streamlit App - Hỗ trợ LaTeX Math và Markdown Tables
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
    page_title="🧠 Chuyển đề thi Word thành bảng chuẩn (LaTeX + Markdown)",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS tùy chỉnh với theme mới
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
    .math-highlight {
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
    .new-feature {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header chính
    st.markdown("""
    <div class="main-header">
        <h1>🧠 Chuyển đề thi Word thành bảng chuẩn</h1>
        <p>Tự động nhận dạng cấu trúc đề thi và chuẩn hóa thành format bảng chuyên nghiệp</p>
        <p><strong>🆕 Hỗ trợ LaTeX Math & Markdown Tables!</strong></p>
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
        
        st.markdown("### 🆕 Tính năng mới:")
        st.markdown("""
        <div class="new-feature">
        🧮 LaTeX Math: $x^2 + y^2 = r^2$
        </div>
        <div class="new-feature">
        📊 Markdown Tables: | A | B | C |
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 📝 Format nhận dạng:")
        st.code("""
Câu 1. Tính giới hạn $\\lim_{x \\to 0} \\frac{\\sin x}{x}$
A. 0
B. 1  
C. $+\\infty$
D. Không tồn tại
        """)
        
        st.markdown("### 📊 Bảng Markdown:")
        st.code("""
| Điểm | Số HS | Tần suất |
|------|-------|----------|
| 8-10 | 15    | 30%      |
| 6-8  | 25    | 50%      |
| 0-6  | 10    | 20%      |
        """)
        
        st.markdown("### ✨ Tính năng:")
        features = [
            "🔍 Tự động nhận dạng cấu trúc",
            "🧮 Chuyển LaTeX thành equation Word",
            "📊 Chuyển Markdown table thành bảng Word", 
            "🎨 Format chuyên nghiệp",
            "📥 Tải xuống ngay lập tức",
            "⚡ Xử lý nhanh chóng"
        ]
        
        for feature in features:
            st.markdown(f"- {feature}")
        
        st.markdown("---")
        st.markdown("### 📞 Hỗ trợ")
        st.info("Ứng dụng tự động nhận dạng và chuyển đổi công thức LaTeX + bảng Markdown.")
        
        st.markdown("### 🧪 File test:")
        st.markdown("""
        Chạy `python create_enhanced_test.py` để tạo file test có:
        - Công thức LaTeX 
        - Bảng Markdown
        - Mixed content
        """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("## 📁 Tải lên đề thi")
        
        uploaded_file = st.file_uploader(
            "Chọn file .docx đề thi cần chuẩn hóa (hỗ trợ LaTeX & Markdown)",
            type=["docx"],
            help="File Word chứa đề thi với công thức LaTeX ($...$) và bảng Markdown (|...|)"
        )
        
        if uploaded_file:
            st.success(f"✅ Đã tải lên: **{uploaded_file.name}**")
            
            # Phân tích cấu trúc đề thi
            with st.expander("🔍 Phân tích cấu trúc đề thi", expanded=True):
                try:
                    doc = Document(uploaded_file)
                    analysis = analyze_exam_structure(doc)
                    
                    # Hiển thị thống kê cơ bản
                    st.markdown(f"""
                    <div class="analysis-result">
                        <h3>📊 Kết quả phân tích:</h3>
                        <p><strong>Tổng số câu hỏi:</strong> {analysis['total_questions']}</p>
                        <p><strong>Câu trắc nghiệm:</strong> {analysis['multiple_choice']}</p>
                        <p><strong>Câu đúng/sai:</strong> {analysis['true_false']}</p>
                        <p><strong>Câu tự luận:</strong> {analysis['essay']}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Hiển thị thống kê LaTeX và Markdown
                    col_math, col_table = st.columns(2)
                    
                    with col_math:
                        if analysis.get('has_math_formulas', False):
                            st.markdown(f"""
                            <div class="math-highlight">
                                <h4>🧮 Công thức LaTeX</h4>
                                <p><strong>Số công thức:</strong> {analysis.get('math_count', 0)}</p>
                                <p>✅ Sẽ được chuyển thành equation Word</p>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.info("🧮 Không phát hiện công thức LaTeX")
                    
                    with col_table:
                        if analysis.get('has_markdown_tables', False):
                            st.markdown(f"""
                            <div class="table-highlight">
                                <h4>📊 Bảng Markdown</h4>
                                <p><strong>Số bảng:</strong> {analysis.get('table_count', 0)}</p>
                                <p>✅ Sẽ được chuyển thành bảng Word</p>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.info("📊 Không phát hiện bảng Markdown")
                    
                    # Hiển thị preview một số câu hỏi
                    if analysis['sample_questions']:
                        st.markdown("### 👀 Preview câu hỏi:")
                        for i, q in enumerate(analysis['sample_questions'][:3]):
                            with st.expander(f"Câu {i+1}: {q['number']}"):
                                # Hiển thị câu hỏi với formatting
                                question_preview = q['question'][:200]
                                if len(q['question']) > 200:
                                    question_preview += "..."
                                
                                st.write(f"**Câu hỏi:** {question_preview}")
                                
                                # Hiển thị đặc điểm đặc biệt
                                features = []
                                if q.get('has_math', False):
                                    features.append("🧮 Có LaTeX")
                                if q.get('has_tables', False):
                                    features.append("📊 Có bảng")
                                
                                if features:
                                    st.write(f"**Đặc điểm:** {', '.join(features)}")
                                
                                if q['type'] == 'multiple_choice':
                                    for option in ['A', 'B', 'C', 'D']:
                                        if q.get(option):
                                            preview = q[option][:50]
                                            if len(q[option]) > 50:
                                                preview += "..."
                                            st.write(f"**{option}:** {preview}")
                
                except Exception as e:
                    st.error(f"Lỗi phân tích: {str(e)}")
                    st.code(traceback.format_exc())
    
    with col2:
        st.markdown("## 🎯 Quy trình xử lý")
        
        process_steps = [
            ("1️⃣", "Tải lên đề thi", "Upload file .docx"),
            ("2️⃣", "Phân tích cấu trúc", "Nhận dạng câu hỏi + LaTeX + Markdown"),
            ("3️⃣", "Chuyển đổi nâng cao", "LaTeX → Equation, Markdown → Table"),
            ("4️⃣", "Chuẩn hóa format", "Tạo bảng Word chuyên nghiệp"),
            ("5️⃣", "Tải xuống kết quả", "File Word hoàn chỉnh")
        ]
        
        for step, title, desc in process_steps:
            st.markdown(f"""
            <div class="feature-card">
                <strong>{step} {title}</strong><br>
                <small>{desc}</small>
            </div>
            """, unsafe_allow_html=True)
        
        # Thêm thông tin về tính năng mới
        st.markdown("## 🆕 Tính năng mới")
        
        st.markdown("""
        <div class="new-feature">
        <h4>🧮 LaTeX Math Support</h4>
        <p>• Inline: $x^2 + y^2$</p>
        <p>• Display: $$\\int_0^1 f(x)dx$$</p>
        <p>• Greek: α, β, γ, π, Σ</p>
        <p>• Fractions: \\frac{a}{b}</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="new-feature">
        <h4>📊 Markdown Tables</h4>
        <p>• Auto-detect table structure</p>
        <p>• Convert to Word native tables</p>  
        <p>• Support LaTeX in cells</p>
        <p>• Professional formatting</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Xử lý chuyển đổi
    if uploaded_file:
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            if st.button("🚀 Chuyển đổi thành bảng chuẩn hóa (LaTeX + Markdown)", type="primary", use_container_width=True):
                with st.spinner("⏳ Đang xử lý đề thi với LaTeX và Markdown..."):
                    try:
                        # Xử lý file
                        output_doc, stats = process_docx(uploaded_file)
                        
                        # Tạo buffer để tải xuống
                        output_stream = BytesIO()
                        output_doc.save(output_stream)
                        output_stream.seek(0)
                        
                        # Tạo tên file
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_filename = f"de_chuan_hoa_enhanced_{timestamp}.docx"
                        
                        # Thông báo thành công
                        st.markdown("""
                        <div class="success-box">
                            <h3>🎉 Chuyển đổi thành công!</h3>
                            <p>Đề thi đã được chuẩn hóa với hỗ trợ LaTeX và Markdown</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Thống kê kết quả
                        st.markdown("## 📊 Thống kê kết quả")
                        
                        col1, col2, col3, col4, col5 = st.columns(5)
                        
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
                        
                        with col5:
                            total_enhanced = stats.get('math_formulas', 0) + stats.get('markdown_tables', 0)
                            st.markdown(f"""
                            <div class="metric-card">
                                <h2 style="color: #e83e8c;">{total_enhanced}</h2>
                                <p>LaTeX + Tables</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        # Hiển thị chi tiết về LaTeX và Markdown
                        if stats.get('math_formulas', 0) > 0 or stats.get('markdown_tables', 0) > 0:
                            st.markdown("### 🆕 Tính năng nâng cao đã xử lý:")
                            
                            enhanced_col1, enhanced_col2 = st.columns(2)
                            
                            with enhanced_col1:
                                if stats.get('math_formulas', 0) > 0:
                                    st.markdown(f"""
                                    <div class="math-highlight">
                                        <h4>🧮 Công thức LaTeX</h4>
                                        <p><strong>{stats['math_formulas']}</strong> công thức đã chuyển thành equation Word</p>
                                        <p>✅ Hiển thị chuyên nghiệp trong bảng</p>
                                    </div>
                                    """, unsafe_allow_html=True)
                            
                            with enhanced_col2:
                                if stats.get('markdown_tables', 0) > 0:
                                    st.markdown(f"""
                                    <div class="table-highlight">
                                        <h4>📊 Bảng Markdown</h4>
                                        <p><strong>{stats['markdown_tables']}</strong> bảng đã chuyển thành bảng Word</p>
                                        <p>✅ Format chuẩn với border và style</p>
                                    </div>
                                    """, unsafe_allow_html=True)
                        
                        # Nút tải xuống
                        st.download_button(
                            label="📥 Tải xuống file Word đã chuẩn hóa (LaTeX + Markdown)",
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
                        if stats.get('math_formulas', 0) > 0:
                            success_messages.append(f"🧮 Đã chuyển đổi {stats['math_formulas']} công thức LaTeX")
                        if stats.get('markdown_tables', 0) > 0:
                            success_messages.append(f"📊 Đã chuyển đổi {stats['markdown_tables']} bảng Markdown")
                        
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
                        - Đảm bảo công thức LaTeX đúng syntax: $x^2$, không phải $x^2
                        - Bảng Markdown cần có header separator: |---|---|
                        - Thử lại với file khác
                        """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 2rem;">
        <p>🧠 <strong>Ứng dụng chuyển đề thi Word thành bảng chuẩn</strong></p>
        <p>🆕 <strong>Enhanced với LaTeX Math & Markdown Tables</strong></p>
        <p>Phát triển với ❤️ bằng Streamlit và Python</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
