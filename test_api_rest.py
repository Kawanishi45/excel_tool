"""
REST APIで直接Gemini APIをテスト
"""

import os
import requests
import json
from dotenv import load_dotenv

load_dotenv()

api_key = os.environ.get('GOOGLE_API_KEY')
print(f"API Key loaded: {api_key[:20]}...")

# REST APIエンドポイント (gemini-2.0-flashを使用)
url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"

# リクエストボディ
payload = {
    "contents": [{
        "parts": [{
            "text": "Say 'Hello!'"
        }]
    }]
}

headers = {
    "Content-Type": "application/json"
}

print("Sending request to Gemini API...")

try:
    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()

    result = response.json()

    if 'candidates' in result and len(result['candidates']) > 0:
        text = result['candidates'][0]['content']['parts'][0]['text']
        print(f"\n✓ Response: {text}")
        print("\n✅ AI API connection successful!")
    else:
        print(f"\n✗ Unexpected response format: {json.dumps(result, indent=2)}")

except requests.exceptions.Timeout:
    print("\n✗ Request timed out after 30 seconds")
except requests.exceptions.RequestException as e:
    print(f"\n✗ Request error: {e}")
except Exception as e:
    print(f"\n✗ Error: {e}")
    import traceback
    traceback.print_exc()
