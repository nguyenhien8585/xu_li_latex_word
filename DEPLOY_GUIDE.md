# 🚀 Hướng dẫn triển khai ứng dụng lên Streamlit Cloud

## 📋 Checklist trước khi deploy

- [ ] Đã có tài khoản GitHub
- [ ] Đã test ứng dụng local thành công
- [ ] Có đầy đủ files: `app.py`, `processor.py`, `requirements.txt`
- [ ] Code đã commit và push lên GitHub

## 🔧 Bước 1: Chuẩn bị GitHub Repository

### 1.1 Tạo repository mới

```bash
# Tạo thư mục project
mkdir de-chuyen-bang-word
cd de-chuyen-bang-word

# Khởi tạo git
git init
```

### 1.2 Thêm files

Copy các files sau vào thư mục:
- `app.py` (Streamlit app chính)
- `processor.py` (Module xử lý)
- `requirements.txt` (Dependencies)
- `README.md` (Hướng dẫn)

### 1.3 Commit và push

```bash
# Thêm files
git add .
git commit -m "🎉 Initial commit - Exam processor app"

# Tạo repository trên GitHub và link
git remote add origin https://github.com/USERNAME/de-chuyen-bang-word.git
git branch -M main
git push -u origin main
```

## 🌐 Bước 2: Deploy trên Streamlit Cloud

### 2.1 Truy cập Streamlit Cloud

1. Mở [share.streamlit.io](https://share.streamlit.io)
2. Click **"Sign up"** hoặc **"Sign in"**
3. Chọn **"Continue with GitHub"**
4. Authorize Streamlit truy cập GitHub

### 2.2 Tạo app mới

1. Click **"New app"**
2. Chọn repository: `USERNAME/de-chuyen-bang-word`
3. Branch: `main`
4. Main file path: `app.py`
5. App URL (optional): `de-chuyen-bang-word` hoặc tên khác

### 2.3 Advanced settings (nếu cần)

```toml
# .streamlit/config.toml (tùy chọn)
[theme]
primaryColor = "#667eea"
backgroundColor = "#ffffff"
secondaryBackgroundColor = "#f0f2f6"
textColor = "#262730"

[server]
maxUploadSize = 200
```

### 2.4 Deploy!

Click **"Deploy!"** và chờ vài phút.

## ✅ Bước 3: Kiểm tra deployment

### 3.1 Logs

Streamlit Cloud sẽ hiển thị logs real-time:

```
Building...
✓ requirements.txt found
✓ Installing dependencies...
✓ Starting Streamlit app...
✓ App is live!
```

### 3.2 Troubleshooting thường gặp

#### Lỗi: "ModuleNotFoundError"

**Nguyên nhân**: Thiếu thư viện trong `requirements.txt`

**Giải pháp**:
```txt
streamlit>=1.28.0
python-docx>=0.8.11
lxml>=4.9.0
```

#### Lỗi: "File not found"

**Nguyên nhân**: Đường dẫn file không đúng

**Giải pháp**: Đảm bảo import đúng:
```python
from processor import process_docx  # Không phải ./processor
```

#### Lỗi: "Memory limit exceeded"

**Nguyên nhân**: File upload quá lớn

**Giải pháp**: Thêm config:
```toml
[server]
maxUploadSize = 200
```

## 🔄 Bước 4: Cập nhật ứng dụng

### 4.1 Chỉnh sửa code

```bash
# Sửa files trên local
git add .
git commit -m "✨ Update: new features"
git push
```

### 4.2 Auto-deploy

Streamlit Cloud tự động deploy khi detect changes trên GitHub.

### 4.3 Manual redeploy

1. Vào Streamlit Cloud dashboard
2. Chọn app
3. Click **"Reboot app"**

## 📊 Bước 5: Monitoring và Analytics

### 5.1 App metrics

Streamlit Cloud cung cấp:
- Visitor count
- Resource usage
- Error logs
- Performance metrics

### 5.2 Custom analytics (optional)

Thêm Google Analytics:

```python
# Trong app.py
import streamlit.components.v1 as components

# Google Analytics tracking
ga_code = """
<script async src="https://www.googletagmanager.com/gtag/js?id=GA_MEASUREMENT_ID"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', 'GA_MEASUREMENT_ID');
</script>
"""
components.html(ga_code, height=0)
```

## 🔐 Bước 6: Bảo mật và Performance

### 6.1 Environment variables

Tạo file `.streamlit/secrets.toml`:

```toml
# Secrets (không commit file này!)
API_KEY = "your-secret-key"
DATABASE_URL = "your-db-url"
```

Sử dụng trong code:
```python
import streamlit as st

api_key = st.secrets["API_KEY"]
```

### 6.2 Caching

Tối ưu performance:

```python
@st.cache_data
def process_large_file(file_content):
    # Heavy processing here
    return result

@st.cache_resource
def load_model():
    # Load ML model once
    return model
```

### 6.3 Rate limiting

```python
import time
from datetime import datetime, timedelta

# Simple rate limiting
if 'last_request' not in st.session_state:
    st.session_state.last_request = datetime.now()

time_diff = datetime.now() - st.session_state.last_request
if time_diff < timedelta(seconds=5):
    st.warning("Vui lòng chờ 5 giây giữa các requests")
    st.stop()
```

## 🎯 Bước 7: Domain tùy chỉnh (Pro)

### 7.1 Upgrade Streamlit Cloud

1. Upgrade to paid plan
2. Add custom domain
3. Configure DNS

### 7.2 Alternative: Vercel/Netlify

Deploy Streamlit app trên các platform khác:

```yaml
# vercel.json
{
  "builds": [
    {
      "src": "app.py",
      "use": "@vercel/python"
    }
  ],
  "routes": [
    {
      "src": "/(.*)",
      "dest": "app.py"
    }
  ]
}
```

## 📱 Bước 8: PWA (Progressive Web App)

Biến ứng dụng thành PWA:

```python
# Thêm manifest.json
manifest = {
    "name": "Exam Processor",
    "short_name": "ExamApp",
    "description": "Convert exam Word to table",
    "start_url": "/",
    "display": "standalone",
    "background_color": "#ffffff",
    "theme_color": "#667eea",
    "icons": [
        {
            "src": "icon-192.png",
            "sizes": "192x192",
            "type": "image/png"
        }
    ]
}

# Service worker
sw_js = """
const CACHE_NAME = 'exam-app-v1';
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(['/']))
  );
});
"""
```

## 🔍 Bước 9: SEO và Sharing

### 9.1 Meta tags

```python
st.set_page_config(
    page_title="🧠 Chuyển đề thi Word thành bảng",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://github.com/username/repo',
        'Report a bug': 'https://github.com/username/repo/issues',
        'About': "Ứng dụng chuyển đề thi Word thành bảng chuẩn hóa"
    }
)
```

### 9.2 Open Graph

```html
<!-- Thêm vào HTML head -->
<meta property="og:title" content="Exam Word Processor">
<meta property="og:description" content="Convert exam Word files to standardized tables">
<meta property="og:image" content="https://your-app.streamlit.app/preview.png">
<meta property="og:url" content="https://your-app.streamlit.app">
```

## 🎉 Hoàn thành!

Sau khi hoàn thành các bước trên, bạn sẽ có:

✅ Ứng dụng web hoạt động trên cloud  
✅ URL công khai để chia sẻ  
✅ Auto-deployment từ GitHub  
✅ Monitoring và analytics  
✅ Bảo mật cơ bản  

### 📞 Support

- **GitHub Issues**: [Link to your repo issues]
- **Streamlit Docs**: [docs.streamlit.io](https://docs.streamlit.io)
- **Community**: [discuss.streamlit.io](https://discuss.streamlit.io)

---

🚀 **Chúc mừng bạn đã deploy thành công ứng dụng!**
