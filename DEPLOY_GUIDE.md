# 🚀 Hướng dẫn Deploy Word Converter

## 📋 Tổng quan các cách deploy

| Phương pháp | Độ khó | Chi phí | Thời gian setup | URL public |
|-------------|--------|---------|-----------------|------------|
| Streamlit Cloud | ⭐ | Miễn phí | 5 phút | ✅ |
| GitHub Codespaces | ⭐⭐ | Miễn phí* | 2 phút | ❌ |
| Local | ⭐⭐⭐ | Miễn phí | 5 phút | ❌ |
| Heroku | ⭐⭐⭐⭐ | Có phí | 15 phút | ✅ |

*GitHub Codespaces: 60 giờ/tháng miễn phí

## 🎯 Cách 1: Streamlit Cloud (Khuyến nghị)

### Bước 1: Chuẩn bị repository
```bash
# Tạo repository mới trên GitHub
# Upload tất cả files: app.py, requirements.txt, README.md, .streamlit/config.toml
```

### Bước 2: Deploy trên Streamlit Cloud
1. Truy cập: https://share.streamlit.io
2. Đăng nhập bằng GitHub
3. Nhấn "New app"
4. Chọn repository của bạn
5. Main file path: `app.py`
6. Nhấn "Deploy!"

### Bước 3: Kiểm tra
- App sẽ có URL dạng: `https://yourapp.streamlit.app`
- Chia sẻ URL này cho người khác sử dụng

### ✅ Ưu điểm:
- Hoàn toàn miễn phí
- Auto-deploy khi push code mới
- SSL certificate tự động
- Không cần server management

## 🔧 Cách 2: GitHub Codespaces

### Bước 1: Tạo Codespace
1. Vào repository trên GitHub
2. Nhấn "Code" → "Codespaces" → "Create codespace on main"
3. Đợi environment setup (2-3 phút)

### Bước 2: Chạy app
```bash
# Trong terminal của Codespace
pip install -r requirements.txt
streamlit run app.py
```

### Bước 3: Truy cập
- Codespace sẽ tự động mở browser với app
- URL dạng: `https://abc123-8501.app.github.dev`

### ✅ Ưu điểm:
- Setup nhanh
- Môi trường development đầy đủ
- Tích hợp với GitHub

### ⚠️ Hạn chế:
- Chỉ chạy khi Codespace active
- 60 giờ miễn phí/tháng

## 💻 Cách 3: Local Development

### Bước 1: Clone repository
```bash
git clone https://github.com/yourusername/word-converter.git
cd word-converter
```

### Bước 2: Setup environment
```bash
# Tạo virtual environment (optional nhưng khuyến nghị)
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Bước 3: Chạy app
```bash
streamlit run app.py
```

### Bước 4: Truy cập
- Mở browser: http://localhost:8501

## 🌐 Cách 4: Heroku (Advanced)

### Bước 1: Chuẩn bị files
Tạo thêm các files:

**Procfile:**
```
web: sh setup.sh && streamlit run app.py --server.port=$PORT --server.address=0.0.0.0
```

**setup.sh:**
```bash
mkdir -p ~/.streamlit/
echo "\
[server]\n\
headless = true\n\
port = $PORT\n\
enableCORS = false\n\
\n\
" > ~/.streamlit/config.toml
```

**runtime.txt:**
```
python-3.9.18
```

### Bước 2: Deploy lên Heroku
```bash
# Install Heroku CLI
# Login
heroku login

# Tạo app
heroku create your-word-converter

# Deploy
git add .
git commit -m "Deploy to Heroku"
git push heroku main
```

## 📊 So sánh chi tiết

### Performance:
- **Streamlit Cloud**: Tốt, server được optimize
- **Codespaces**: Rất tốt, container riêng
- **Local**: Tùy thuộc máy local
- **Heroku**: Tốt, nhưng có giới hạn memory

### Scalability:
- **Streamlit Cloud**: Auto-scale, handle nhiều user
- **Codespaces**: 1 user/session
- **Local**: 1 user/session
- **Heroku**: Scale được nhưng có phí

### Security:
- **Streamlit Cloud**: HTTPS mặc định
- **Codespaces**: HTTPS, private by default
- **Local**: HTTP, chỉ local access
- **Heroku**: HTTPS mặc định

## 🔐 Bảo mật và Cấu hình

### Environment Variables (nếu cần):
```bash
# Trong Streamlit Cloud
# Settings → Advanced → Secrets
# Thêm vào secrets.toml:

[database]
host = "your-host"
password = "your-password"
```

### File size limits:
```toml
# .streamlit/config.toml
[server]
maxUploadSize = 10  # MB
```

### CORS configuration:
```toml
[server]
enableCORS = false
enableXsrfProtection = false
```

## 🐛 Troubleshooting Deploy

### Lỗi thường gặp:

**1. "ModuleNotFoundError"**
```bash
# Kiểm tra requirements.txt có đầy đủ không
pip freeze > requirements.txt
```

**2. "Port binding error"**
```bash
# Đảm bảo app listen trên port từ environment
port = int(os.environ.get("PORT", 8501))
```

**3. "File upload error"**
```bash
# Kiểm tra file size limit
# Tăng maxUploadSize trong config.toml
```

**4. "Memory error"**
```bash
# Optimize code để dùng ít memory hơn
# Hoặc upgrade plan (Heroku)
```

## 📈 Monitoring và Analytics

### Streamlit Cloud:
- Built-in analytics trong dashboard
- View count, user metrics
- Error tracking

### Custom analytics:
```python
# Thêm vào app.py
import streamlit as st

# Track usage
if 'visit_count' not in st.session_state:
    st.session_state.visit_count = 0
st.session_state.visit_count += 1
```

## 🔄 Auto-deployment

### GitHub Actions (cho advanced users):
```yaml
# .github/workflows/deploy.yml
name: Deploy to Streamlit Cloud
on:
  push:
    branches: [ main ]
jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
    - uses: actions/checkout@v2
    - name: Setup Python
      uses: actions/setup-python@v2
      with:
        python-version: '3.9'
    - name: Install dependencies
      run: pip install -r requirements.txt
    - name: Run tests
      run: python -m pytest tests/
```

## 💡 Tips và Best Practices

1. **Code Organization:**
   - Tách logic xử lý ra module riêng
   - Sử dụng caching cho heavy operations
   - Error handling đầy đủ

2. **User Experience:**
   - Progress bars cho operations lâu
   - Clear error messages
   - Input validation

3. **Performance:**
   - Sử dụng st.cache_data cho functions nặng
   - Optimize file processing
   - Lazy loading

4. **Security:**
   - Validate file uploads
   - Sanitize user inputs
   - Set appropriate file size limits

## 📞 Support

Nếu gặp vấn đề trong quá trình deploy:

1. Kiểm tra logs trong platform dashboard
2. Verify tất cả dependencies trong requirements.txt
3. Test local trước khi deploy
4. Tạo issue trên GitHub repository

---

**Lưu ý:** Hướng dẫn này được update thường xuyên. Check lại repository để có version mới nhất.
