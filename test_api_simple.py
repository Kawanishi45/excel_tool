"""
Google Gemini API接続の簡易テスト
"""

import os
from dotenv import load_dotenv
import google.generativeai as genai

# 環境変数を読み込む
load_dotenv()

api_key = os.environ.get('GOOGLE_API_KEY')

if not api_key:
    print("❌ GOOGLE_API_KEY が .env ファイルに設定されていません")
    exit(1)

print(f"✓ APIキーを取得しました: {api_key[:10]}...")

# Gemini APIを設定
genai.configure(api_key=api_key)

print("✓ Gemini APIを設定しました")

# モデルを取得
try:
    model = genai.GenerativeModel('gemini-1.5-flash')
    print("✓ モデル 'gemini-1.5-flash' を取得しました")
except Exception as e:
    print(f"❌ モデル取得エラー: {e}")
    exit(1)

# シンプルなテストプロンプト
test_prompt = "Hello! Please respond with 'API connection successful!' if you can read this."

print("\n--- テストプロンプトを送信中 ---")
print(f"プロンプト: {test_prompt}\n")

try:
    response = model.generate_content(test_prompt)
    print("✓ レスポンスを受信しました\n")
    print("--- AIからのレスポンス ---")
    print(response.text)
    print("\n✅ AI API接続テスト成功！")
except Exception as e:
    print(f"❌ API呼び出しエラー: {e}")
    exit(1)
