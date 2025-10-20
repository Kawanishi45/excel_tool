"""
最小限のGemini APIテスト
"""

import os
import sys
from dotenv import load_dotenv

load_dotenv()

api_key = os.environ.get('GOOGLE_API_KEY')
print(f"API Key loaded: {api_key[:20]}...")

print("Importing google.generativeai...")
import google.generativeai as genai

print("Configuring API...")
genai.configure(api_key=api_key)

print("Creating model...")
model = genai.GenerativeModel('gemini-1.5-flash')

print("Generating content...")
sys.stdout.flush()

try:
    response = model.generate_content(
        "Say 'Hello!'",
        request_options={'timeout': 30}  # 30秒のタイムアウト
    )
    print(f"\n✓ Response: {response.text}")
except Exception as e:
    print(f"\n✗ Error: {e}")
    import traceback
    traceback.print_exc()
