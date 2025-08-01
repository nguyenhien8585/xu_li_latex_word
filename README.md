# 🧠 Ứng dụng chuyển đề thi Word thành bảng chuẩn

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://your-app-url.streamlit.app)

## 📋 Tổng quan

Ứng dụng web **Streamlit** tự động chuyển đổi đề thi Word thành format bảng chuẩn hóa, chuyên nghiệp. Hỗ trợ nhận dạng và xử lý 3 loại câu hỏi:

- **Phần I**: Câu hỏi trắc nghiệm (A, B, C, D)
- **Phần II**: Câu hỏi đúng/sai  
- **Phần III**: Câu hỏi tự luận

## ✨ Tính năng chính

- 🔍 **Tự động nhận dạng** cấu trúc đề thi
- 📊 **Tạo bảng Word** chuẩn hóa, chuyên nghiệp
- 🎨 **Format đẹp** với màu sắc và typography chuẩn
- ⚡ **Xử lý nhanh** với giao diện thân thiện
- 📥 **Tải xuống ngay** file Word kết quả
- 📱 **Responsive** trên mọi thiết bị

## 🚀 Cách chạy local

### 1. Clone repository

```bash
git clone https://github.com/your-username/de-chuyen-bang-word.git
cd de-chuyen-bang-word
```

### 2. Cài đặt dependencies

```bash
pip install -r requirements.txt
```

### 3. Chạy ứng dụng

```bash
streamlit run app.py
```

### 4. Truy cập ứng dụng

Mở trình duyệt và truy cập: `http://localhost:8501`

## 🌐 Deploy trên Streamlit Cloud

### 1. Chuẩn bị repository

- Fork repository này về GitHub của bạn
- Đảm bảo có đầy đủ files: `app.py`, `processor.py`, `requirements.txt`

### 2. Deploy

1. Truy cập [share.streamlit.io](https://share.streamlit.io)
2. Đăng nhập bằng GitHub
3. Click **"New app"**
4. Chọn repository và branch
5. Nhập đường dẫn: `app.py`
6. Click **"Deploy!"**

### 3. Chia sẻ

Ứng dụng sẽ có URL dạng: `https://your-app-name.streamlit.app`

## 📂 Cấu trúc project

```
de-chuyen-bang-word/
│
├── app.py              # Streamlit app chính
├── processor.py        # Module xử lý Word -> bảng  
├── requirements.txt    # Thư viện cần thiết
├── README.md          # File này
│
└── example/           # Thư mục mẫu
    ├── input.docx     # File đề thi mẫu
    └── output.docx    # File kết quả mẫu
```

## 📝 Format đề thi hỗ trợ

### Câu hỏi trắc nghiệm

```
Câu 1. Nội dung câu hỏi ở đây...
A. Đáp án A
B. Đáp án B  
C. Đáp án C
D. Đáp án D
```

### Câu hỏi đúng/sai

```
Câu 5. Phát biểu nào sau đây là đúng?
```

### Câu hỏi tự luận

```
Câu 10. Trình bày về...
```

## 🔧 Cấu hình nâng cao

### Tùy chỉnh style bảng

Trong file `processor.py`, bạn có thể chỉnh sửa:

- Màu sắc header: `RGBColor(31, 78, 121)`
- Font size: `Pt(11)`
- Độ rộng cột: `Inches(1.5)`

### Thêm loại câu hỏi mới

1. Cập nhật function `classify_question_type()`
2. Thêm parser trong `parse_question()`
3. Tạo bảng mới trong `process_docx()`

## 🐛 Xử lý lỗi thường gặp

### Lỗi: "No module named 'docx'"

```bash
pip install python-docx
```

### Lỗi: "File format not supported"

- Đảm bảo file là `.docx` (không phải `.doc`)
- Kiểm tra file không bị corrupt

### Lỗi: "No questions found"

- Kiểm tra format câu hỏi có đúng không
- Đảm bảo có text "Câu 1", "Câu 2"...

## 📊 Kết quả mẫu

Ứng dụng tạo ra file Word với:

- **Header** chuyên nghiệp với tiêu đề và ngày tạo
- **Bảng trắc nghiệm** với 6 cột (STT, Câu hỏi, A, B, C, D)
- **Bảng đúng/sai** với 3 cột (STT, Câu hỏi, Đáp án)
- **Bảng tự luận** với 3 cột (STT, Câu hỏi, Ghi chú)
- **Styling** đẹp với màu sắc và border chuẩn

## 🤝 Đóng góp

1. Fork repository
2. Tạo feature branch: `git checkout -b feature/AmazingFeature`
3. Commit changes: `git commit -m 'Add AmazingFeature'`
4. Push to branch: `git push origin feature/AmazingFeature`
5. Open Pull Request

## 📝 Changelog

### v1.0.0 (2024-01-01)
- 🎉 Phiên bản đầu tiên
- ✅ Hỗ trợ 3 loại câu hỏi
- 🎨 UI/UX chuyên nghiệp
- 📊 Bảng Word chuẩn hóa

## 📞 Hỗ trợ

- 🐛 **Bug reports**: [GitHub Issues](https://github.com/your-username/de-chuyen-bang-word/issues)
- 💡 **Feature requests**: [GitHub Discussions](https://github.com/your-username/de-chuyen-bang-word/discussions)
- 📧 **Email**: your-email@example.com

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

---

⭐ **Star this repo** nếu bạn thấy hữu ích!

Phát triển với ❤️ bằng [Streamlit](https://streamlit.io) và [Python](https://python.org)
