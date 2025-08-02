#!/usr/bin/env python3
"""
Script test để debug lỗi mammoth và kiểm tra API
"""

import mammoth
import sys
import os

def test_mammoth_api():
    """Test API của mammoth để xem cấu trúc return object"""
    
    print("🔍 Testing mammoth API...")
    print(f"📦 Mammoth version: {mammoth.__version__}")
    
    # Tạo sample docx content (minimal)
    sample_html = "<p>Test content</p>"
    
    print(f"🧪 Mammoth attributes:")
    
    # Kiểm tra available methods
    methods = [attr for attr in dir(mammoth) if not attr.startswith('_')]
    print(f"📋 Available methods: {methods}")
    
    print("\n" + "="*50)
    
    # Nếu có file test
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        if os.path.exists(test_file):
            test_real_file(test_file)
        else:
            print(f"❌ File not found: {test_file}")
    else:
        print("💡 Usage: python test_mammoth.py <docx_file>")
        print("   Để test với file Word thực tế")

def test_real_file(docx_path):
    """Test với file Word thực tế"""
    
    print(f"\n🔍 Testing with real file: {docx_path}")
    
    try:
        with open(docx_path, "rb") as docx_file:
            print("📖 Converting with mammoth...")
            result = mammoth.convert_to_html(docx_file)
            
            print(f"📋 Result type: {type(result)}")
            print(f"📋 Result attributes: {dir(result)}")
            
            # Test different possible attributes
            possible_attrs = ['html', 'value', 'content', 'text']
            
            for attr in possible_attrs:
                if hasattr(result, attr):
                    content = getattr(result, attr)
                    print(f"✅ Found .{attr}: {type(content)} (length: {len(str(content))})")
                    
                    # Show preview
                    preview = str(content)[:200] + "..." if len(str(content)) > 200 else str(content)
                    print(f"   Preview: {repr(preview)}")
                else:
                    print(f"❌ No .{attr} attribute")
            
            # Test messages/warnings
            message_attrs = ['messages', 'warnings', 'errors']
            for attr in message_attrs:
                if hasattr(result, attr):
                    messages = getattr(result, attr)
                    print(f"✅ Found .{attr}: {type(messages)} (count: {len(messages)})")
                    for msg in messages:
                        print(f"   Message: {msg}")
                else:
                    print(f"❌ No .{attr} attribute")
                    
    except Exception as e:
        print(f"❌ Error: {e}")
        print(f"   Type: {type(e)}")
        import traceback
        traceback.print_exc()

def show_sample_usage():
    """Hiển thị cách sử dụng mammoth đúng"""
    
    print("\n" + "="*50)
    print("📚 CORRECT MAMMOTH USAGE:")
    print("="*50)
    
    code_sample = '''
import mammoth

# Method 1: Từ file path
with open("document.docx", "rb") as docx_file:
    result = mammoth.convert_to_html(docx_file)
    html = result.value      # ← HTML content
    messages = result.messages  # ← Warnings/errors

# Method 2: Từ bytes (Streamlit upload)
uploaded_bytes = uploaded_file.getvalue()
result = mammoth.convert_to_html(io.BytesIO(uploaded_bytes))
html = result.value
messages = result.messages

# Method 3: With options
result = mammoth.convert_to_html(
    docx_file,
    convert_image=mammoth.images.img_element(lambda image: {
        "src": f"images/{image.filename}"
    })
)
'''
    
    print(code_sample)

if __name__ == "__main__":
    test_mammoth_api()
    show_sample_usage()
    
    print("\n🔧 DEBUGGING TIPS:")
    print("="*30)
    print("1. Check mammoth version: pip show mammoth")
    print("2. Upgrade mammoth: pip install --upgrade mammoth") 
    print("3. Try different file formats")
    print("4. Check file permissions")
    print("5. Test with simple Word file first")
