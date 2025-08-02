# 🚀 Hướng dẫn Deploy nhanh trên GitHub và Streamlit Cloud

## 📋 Checklist trước khi deploy

- [ ] Đã có tài khoản GitHub
- [ ] Đã có tài khoản Streamlit Cloud (miễn phí)
- [ ] Các file đã sẵn sàng: `app.py`, `requirements.txt`, `README.md`

## 🎯 Bước 1: Tạo GitHub Repository

### 1.1. Tạo repository mới
```bash
# Trên GitHub.com
1. Click "New repository"
2. Đặt tên: "word-document-processor" 
3. Chọn "Public"
4. Check "Add a README file"
5. Click "Create repository"
```

### 1.2. Upload files
```bash
# Cách 1: Qua GitHub Web Interface
1. Click "uploading an existing file"
2. Drag & drop hoặc chọn files:
   - app.py
   - requirements.txt
   - README.md
   - .streamlit/config.toml (tùy chọn)
3. Commit files

# Cách 2: Qua Git command line
git clone https://github.com/your-username/word-document-processor.git
cd word-document-processor
# Copy các files vào folder
git add .
git commit -m "Initial commit: Word processor app"
git push origin main
```

## 🌐 Bước 2: Deploy trên Streamlit Cloud

### 2.1. Truy cập Streamlit Cloud
1. Mở [share.streamlit.io](https://share.streamlit.io)
2. Click "Sign in with GitHub"
3. Authorize Streamlit access

### 2.2. Tạo app mới
1. Click "New app"
2. **Repository**: Chọn `your-username/word-document-processor`
3. **Branch**: `main`
4. **Main file path**: `app.py`
5. **App URL** (tùy chọn): `word-processor` hoặc để trống
6. Click "Deploy!"

### 2.3. Chờ deploy
- Quá trình mất 2-5 phút
- Streamlit sẽ cài đặt dependencies từ `requirements.txt`
- Nếu có lỗi, kiểm tra logs và sửa

## ✅ Bước 3: Kiểm tra và sử dụng

### 3.1. Kiểm tra app
URL của bạn sẽ có dạng:
```
https://word-processor-[random].streamlit.app
hoặc
https://your-username-word-document-proces-[random].streamlit.app
```

### 3.2. Test các tính năng
- [ ] Upload file Word (.docx)
- [ ] Xử lý công thức LaTeX
- [ ] Xử lý bảng Markdown  
- [ ] Download file đã xử lý

## 🔧 Bước 4: Tùy chỉnh và cập nhật

### 4.1. Cập nhật code
```bash
# Local changes
git add .
git commit -m "Update: [mô tả thay đổi]"
git push origin main

# Streamlit sẽ tự động redeploy sau ~1 phút
```

### 4.2. Quản lý app
- Truy cập [share.streamlit.io](https://share.streamlit.io)
- Vào "My apps" để xem, edit, delete apps
- Xem logs để debug

## 🎨 Bước 5: Tùy chỉnh giao diện (tùy chọn)

### 5.1. Custom theme
Tạo file `.streamlit/config.toml`:
```toml
[theme]
primaryColor = "#FF6B6B"          # Màu chính
backgroundColor = "#FFFFFF"        # Màu nền
secondaryBackgroundColor = "#F0F2F6"  # Màu nền phụ
textColor = "#262730"             # Màu chữ
```

### 5.2. Custom domain (Pro plan)
- Upgrade lên Streamlit Cloud Pro
- Thêm custom domain trong settings

## 🚨 Troubleshooting

### Lỗi thường gặp:

#### 1. ModuleNotFoundError
**Lỗi**: `ModuleNotFoundError: No module named 'docx'`
**Giải pháp**: Kiểm tra `requirements.txt` có `python-docx`

#### 2. File upload lỗi
**Lỗi**: File quá lớn
**Giải pháp**: Thêm vào `.streamlit/config.toml`:
```toml
[server]
maxUploadSize = 200
```

#### 3. Memory error
**Lỗi**: Out of memory
**Giải pháp**: Optimize code hoặc upgrade plan

#### 4. Deploy fail
**Lỗi**: Build failed
**Giải pháp**: 
- Kiểm tra syntax lỗi trong `app.py`
- Xem logs chi tiết trên Streamlit Cloud
- Đảm bảo `requirements.txt` đúng format

## 📊 Monitoring và Analytics

### Xem usage stats:
1. Streamlit Cloud dashboard
2. GitHub traffic (Settings > Insights)
3. Custom analytics (Google Analytics)

### Performance tips:
- Cache expensive operations với `@st.cache_data`
- Optimize file size
- Minimize dependencies

## 🔐 Bảo mật

### Khuyến nghị:
- Không hardcode API keys
- Sử dụng Streamlit secrets cho sensitive data
- Regular security updates

### Streamlit Secrets:
1. Vào app settings trên Streamlit Cloud
2. Thêm secrets trong format TOML
3. Access qua `st.secrets["key"]`

## 📱 Chia sẻ app

### URL để chia sẻ:
```
https://your-app-name.streamlit.app
```

### Embed trong website:
```html
<iframe src="https://your-app-name.streamlit.app" 
        width="100%" height="600px">
</iframe>
```

### QR Code:
Tạo QR code trỏ đến URL app để chia sẻ dễ dàng

---

## 🎉 Hoàn thành!

Chúc mừng! Bạn đã successfully deploy ứng dụng Word Document Processor.

**Next steps:**
- Share URL với users
- Collect feedback
- Iterate và improve
- Monitor usage và performance

**Need help?** 
- GitHub Issues: [Link to your repo issues]
- Streamlit Community: [community.streamlit.io](https://community.streamlit.io)
- Documentation: [docs.streamlit.io](https://docs.streamlit.io)
