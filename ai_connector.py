"""
AI連携モジュール
「構造（IDアンカー画像）」と「テキスト（JSON指示書）」をAIに渡し、最終的なMermaidコードを生成させる。
"""
import os
import json
from dotenv import load_dotenv
import google.generativeai as genai
from PIL import Image

# 環境変数を読み込み
load_dotenv()


def generate_mermaid_code(json_path, image_path):
    """
    JSON指示書とIDアンカー画像からMermaidコードを生成する

    Args:
        json_path (str): instructions.jsonのパス
        image_path (str): anchor_image.pngのパス

    Returns:
        str: 生成されたMermaidコード（クリーンな形式）
    """
    # プロンプトと画像を準備
    prompt_text, image_object = build_prompt(json_path, image_path)

    # Gemini APIを使用してMermaidコードを生成
    raw_response = _call_gemini_api(prompt_text, image_object)

    # Markdownコードブロックを除去してクリーンなMermaidコードを抽出
    clean_code = _extract_mermaid_code(raw_response)

    return clean_code


def build_prompt(json_path, image_path):
    """
    AIへのプロンプトと画像オブジェクトを生成する

    Args:
        json_path (str): instructions.jsonのパス
        image_path (str): anchor_image.pngのパス

    Returns:
        tuple: (prompt_text, image_object)
    """
    # JSONファイルを読み込み
    with open(json_path, 'r', encoding='utf-8') as f:
        json_data = json.load(f)

    # 画像を読み込み
    image_object = Image.open(image_path)

    # プロンプトテンプレートを構築
    prompt_text = f"""あなたは、提供された画像とJSONデータからMermaidフローチャートを生成するシステムアーキテクトです。

以下の2つの情報を提供します。

【情報1：フローチャートの構造画像（ID付き）】
※ 添付された画像を参照してください

* この画像は、図形の「つながり（矢印）」と「配置」を示しています。
* 各図形には `node_XXX` というIDが振られています。
* あなたのタスクは、この画像から「IDとIDのつながり（矢印）」と「矢印に付随する分岐ラベル（Yes/Noなど）」を正確に読み取ることです。

【情報2：図形の詳細データ（JSON）】
```json
{json.dumps(json_data, ensure_ascii=False, indent=2)}
```

* これは、画像内の各IDに対応する「正式なテキスト」と「図形の種類」のリストです。

【タスク】

1. 【情報1】の画像の「ID間のつながり」を視覚的に解析してください。
2. 【情報1】の画像の矢印の近くにある「分岐ラベル（"Yes", "No", "OK", "NG"など）」を読み取ってください。これらは【情報2】のJSONには含まれていません。
3. 【情報2】のJSONを使い、各IDを「正式なテキスト」と「図形の種類」にマッピングしてください。
4. この情報を組み合わせて、完全なMermaid記法（`graph TD`）のコードを生成してください。

【Mermaid生成ルール】

* **グラフ方向:** 常に `graph TD` （上から下）を使用します。
* **ノード定義 (必須):**
    * `id["テキスト"]` (標準の四角形: auto_shape)
    * `id{{"テキスト"}}` (ひし形: decision - 注意: ダブルブレース)
    * `id(["テキスト"])` (角丸四角形: terminator)
    * ※ 【情報2】の `shape_type` を参考に、Mermaidの適切な括弧（`[]`, `{{}}`, `()`）を使い分けてください。
* **つながり:**
    * `id1 --> id2` (標準の矢印)
* **分岐のラベル (最重要):**
    * ひし形からの矢印には、【情報1】の画像から読み取った分岐ラベルを必ず付与してください。
    * **書式:** `id1 -->|"ラベル"| id2`
    * **例:** `node_003 -->|"Yes"| node_004`
    * **例:** `node_003 -->|"No"| node_005`

生成したMermaidコードのみを、Markdownコードブロック（```mermaid ... ```）で出力してください。
"""

    return prompt_text, image_object


def _call_gemini_api(prompt_text, image_object):
    """
    Gemini APIを呼び出してMermaidコードを生成する

    Args:
        prompt_text (str): プロンプトテキスト
        image_object (PIL.Image): 画像オブジェクト

    Returns:
        str: APIからの生のレスポンス
    """
    # APIキーを取得
    api_key = os.environ.get('GOOGLE_API_KEY')
    if not api_key:
        raise ValueError(
            "GOOGLE_API_KEY not found in environment variables. "
            "Please create a .env file with your API key."
        )

    # Gemini APIを設定
    genai.configure(api_key=api_key)

    # マルチモーダルモデルを使用
    model = genai.GenerativeModel('gemini-1.5-flash')

    # コンテンツを生成
    response = model.generate_content([prompt_text, image_object])

    return response.text


def _extract_mermaid_code(raw_response):
    """
    AIのレスポンスからMermaidコードブロックを抽出する

    Args:
        raw_response (str): AIからの生のレスポンス

    Returns:
        str: クリーンなMermaidコード
    """
    # Markdownコードブロックを除去
    lines = raw_response.split('\n')
    mermaid_lines = []
    in_code_block = False

    for line in lines:
        # コードブロックの開始
        if line.strip().startswith('```mermaid'):
            in_code_block = True
            continue
        # コードブロックの終了
        if line.strip() == '```' and in_code_block:
            in_code_block = False
            continue
        # コードブロック内の行を追加
        if in_code_block:
            mermaid_lines.append(line)

    # コードブロックが見つからない場合は、全体を返す
    if not mermaid_lines:
        return raw_response.strip()

    return '\n'.join(mermaid_lines)


if __name__ == "__main__":
    # テスト用コード
    import sys

    if len(sys.argv) >= 3:
        json_path = sys.argv[1]
        image_path = sys.argv[2]

        print("Testing ai_connector...")
        print("=" * 60)
        print(f"JSON: {json_path}")
        print(f"Image: {image_path}")
        print("=" * 60)

        try:
            mermaid_code = generate_mermaid_code(json_path, image_path)

            print("\nGenerated Mermaid Code:")
            print("=" * 60)
            print(mermaid_code)
            print("=" * 60)

        except Exception as e:
            print(f"\n✗ Error: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("Usage: python ai_connector.py <json_path> <image_path>")
        print("Example: python ai_connector.py output/instructions.json output/anchor_image.png")
