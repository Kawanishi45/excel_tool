# Excel-Mermaid変換ツール 開発タスクリスト (dev\_tasks.md)

## \!\! 最重要注意事項 \!\!

このタスクリストを上から順に実行してください。

1.  **進捗の更新:** タスクを一つ完了するたびに、このMarkdownファイルの該当するチェックボックスを [ ] から `[x]` に更新し、ファイルを保存します。
2.  **問題解決の厳守:** タスクがうまくいかない時、エラーが出た時、決して思いつきの修正を試みません。それは早計であり愚かです。代わりに、以下の問題解決ステップを一つずつ、ゆっくりと進めます。
    1.  「今のタスクの目的は何か」を振り返る。
    2.  「発生している事象（エラーメッセージ、意図しない出力）」を俯瞰的に理解する。
    3.  「原因はどこにあるか」を調査する（コードの読み返し、変数の確認）。
    4.  「原因を取り除くための修正案」を立て、それを実行し、再度確認する。

-----

## フェーズ0: プロジェクト初期設定

  - [x] **実装:** `excel-to-mermaid` という名前でプロジェクトディレクトリを作成します。
  - [x] **実装:** `cd excel-to-mermaid` でディレクトリに移動し、`git init` を実行します。
  - [x] **実装:** `.gitignore` ファイルを作成し、`__pycache__/`, `*.pyc`, `.env`, `*.xlsx`, `*.png`, `*.json`, `*.md` (ただし `README.md` と `dev_tasks.md` は除く), `venv/`, `dist/`, `build/` などを追加します。
  - [x] **実装:** `requirements.txt` ファイルを作成し、以下の初期ライブラリを記述します。
    ```
    xlwings
    Pillow
    mss
    python-dotenv
    google-generativeai
    ```
  - [x] **実装:** `python -m venv venv` で仮想環境を構築し、アクティベートします。
  - [x] **実装:** `pip install -r requirements.txt` を実行し、ライブラリをインストールします。
  - [x] **動作確認:**
      * **実行:** `pip list` を実行します。
      * **確認 (OK条件):** `xlwings`, `Pillow`, `mss`, `google-generativeai` などがリストに含まれていること。
  - [x] **Git:** `git add .` `git commit -m "feat: Initial project setup, venv, and dependencies"`

-----

## フェーズ1: モジュール1 (Excel解析 - `excel_parser.py`)

  - [x] **実装:** `excel_parser.py` ファイルを作成します。
  - [x] **実装:** `xlwings` をインポートし、Excelファイルパスとシート名を引数に取り、`sheet` オブジェクトを取得するメイン関数 `def parse_excel_shapes(file_path, sheet_name):` の骨格を作成します。
  - [x] **実装:** (仕様書 4.1.1) `sheet` オブジェクトを受け取り、全シェイプをループ処理し、`temp_id`, `text`, `position` (dict), `shape_type` を含む辞書のリスト `all_shapes` を作成して返す内部関数 `_get_all_shapes(sheet)` を実装します。
  - [x] **実装:** (仕様書 4.1.2) `all_shapes` リストを受け取り、`container_shapes` (ノード候補) と `text_shapes` (ラベル候補) の2つのリストに分類する内部関数 `_classify_shapes(all_shapes)` を実装します。
      * *ヒント: `container_shapes` は `msoTextBox` や `msoConnector` 以外、`text_shapes` は `text` が空でないか `msoTextBox` のもの。*
  - [x] **実装:** (仕様書 4.1.3) `container_shapes` と `text_shapes` を受け取り、座標マッピングを行う内部関数 `_map_text_to_containers(container_shapes, text_shapes)` を実装します。
      * **Logic A:** まず、コンテナ自身が有効なテキストを持っているか（`container['text']` が空でない）をチェックします。
      * **Logic B:** Logic AがFalseの場合のみ、`text_shapes` を走査し、テキストシェイプの**中心座標**がコンテナの `position` 範囲内に含まれるかを判定します。
      * **紐付け:** 含まれるテキストシェイプを発見した場合、`container['text']` をそのテキストで上書きし、そのテキストシェイプは以降の検索対象から除外します（重複防止）。
  - [x] **実装:** `parse_excel_shapes` のメインロジックを完成させます。`_get_all_shapes` -\> `_classify_shapes` -\> `_map_text_to_containers` の順で呼び出し、最終的にテキストがマッピングされた `container_shapes` のリストを返すようにします。
  - [x] **動作確認:**
      * **準備:** 以下の3つのノードと矢印を持つ `test_chart_simple.xlsx` を手動で作成します。
        1.  「開始」というテキストが *内部に* ある角丸四角形。
        2.  「処理A」というテキストが *内部に* ある四角形。
        3.  「処理B」という四角形（枠線あり）と、その *上に* 配置された「処理Bのラベル」というテキストボックス（枠線なし）。
      * **実行:** `test_parser.py` という一時ファイルを作成し、`excel_parser.parse_excel_shapes('test_chart_simple.xlsx', 'Sheet1')` を呼び出し、戻り値（`mapped_containers`）を `print()` します。
      * **確認 (OK条件):**
        1.  コンソールに3つの要素を持つリストが出力されること。
        2.  1番目の要素が `text: "開始"` を持っていること。
        3.  2番目の要素が `text: "処理A"` を持っていること。
        4.  3番目の要素が `text: "処理Bのラベル"` を持っていること（座標マッピングが成功）。
  - [x] **Git:** `git add excel_parser.py test_parser.py test_chart_simple.xlsx` `git commit -m "feat(parser): Implement Module 1 for Excel shape and text mapping"`

-----

## フェーズ2: モジュール2 (資材生成 - `asset_generator.py`)

  - [x] **実装:** `asset_generator.py` ファイルを作成します。`Pillow (PIL)` (Image, ImageDraw, ImageFont) と `mss`, `json` をインポートします。
  - [x] **実装:** (仕様書 4.2.1) モジュール1の出力 (`mapped_containers`) を受け取り、`id` (例: `node_001`) を割り振り、仕様書通りのJSON (`instructions.json`) を生成して保存する関数 `generate_json_instructions(mapped_containers, output_path)` を実装します。この関数は、JSONデータ（Pythonのリスト）も返すようにします。
  - [x] **実装:** (仕様書 4.2.2.a) Excelファイルパスとシート名を引数に取り、対象領域のスクリーンショットを `temp_original.png` として保存する関数 `_get_chart_screenshot(file_path, sheet_name)` を実装します。（`xlwings` の `range.to_png` または `mss` を使用）
  - [x] **実装:** (仕様書 4.2.2.b/c/d) `_get_chart_screenshot` で得た画像パスと、`generate_json_instructions` で得たJSONデータを引数に取り、IDアンカー画像を生成する関数 `generate_anchor_image(original_image_path, instructions_json, output_path)` を実装します。
      * `Pillow` で元画像を開き、`ImageDraw` を準備します。
      * `instructions_json` をループし、各要素の `position` に基づいて `draw.rectangle()` で白い四角形（マスキング）を描画します。
      * 同じ位置の中心に `draw.text()` で `id` を描画します。
      * **重要:** この処理は、JSONに含まれるノードのみを対象とし、矢印や「Yes/No」ラベル（JSONに含まれないテキスト）はマスキング *しない* ようにします。
      * `output_path` (`anchor_image.png`) に保存します。
  - [x] **実装:** これら二つの関数を呼び出すメイン関数 `generate_assets(mapped_containers, excel_file, sheet_name, json_out_path, image_out_path)` を実装します。
  - [x] **動作確認:**
      * **実行:** `test_asset_gen.py` を作成し、`test_chart_simple.xlsx` と `excel_parser` の結果を使って `asset_generator.generate_assets(...)` を呼び出します。
      * **確認 (OK条件):**
        1.  `instructions.json` が生成されること。中身に `id`, `text`, `shape_type`, `position` が含まれていることを確認。
        2.  `anchor_image.png` が生成されること。
        3.  `anchor_image.png` を目視で確認し、元のExcelの図形テキスト部分が白く塗りつぶされ、代わりに `node_001`, `node_002`... とIDが描画されていること。
        4.  図形間の「矢印」が、マスキングされずに *残っている* こと。
  - [x] **Git:** `git add asset_generator.py test_asset_gen.py` `git commit -m "feat(assets): Implement Module 2 for JSON and Anchor Image generation"`

-----

## フェーズ3: モジュール3 (AI連携 - `ai_connector.py`)

  - [x] **実装:** `ai_connector.py` ファイルを作成します。`google.generativeai` (または `openai`), `os`, `json`, `PIL.Image` をインポートし、`load_dotenv()` で環境変数を読み込みます。
  - [x] **実装:** (仕様書 4.3.2) `instructions.json` ファイルパスと `anchor_image.png` ファイルパスを引数に取り、仕様書 (4.3.2) 通りの**プロンプト文字列**と**画像オブジェクト**を生成する関数 `build_prompt(json_path, image_path)` を実装します。
      * *ヒント: JSONを読み込み、f-stringを使ってプロンプトテンプレートに埋め込みます。画像は `PIL.Image.open()` で開きます。*
  - [x] **実装:** (仕様書 4.3) `build_prompt` の結果（プロンプト文字列と画像オブジェクト）を受け取り、マルチモーダルAI (Gemini) のAPIを呼び出す関数 `generate_mermaid_code(prompt_text, image_object)` を実装します。
      * APIキーは `os.environ.get('GOOGLE_API_KEY')` などで取得します。
  - [x] **実装:** AIからのレスポンス（`response.text`）を受け取り、Markdownのコードブロック (` mermaid ...  `) を除去し、Mermaidコード本体のみをクリーンなテキストとして返す内部関数 `_extract_mermaid_code(raw_response)` を実装し、`generate_mermaid_code` はこのクリーンなコードを返すようにします。
  - [x] **動作確認:**
      * **準備:**
        1.  `.env` ファイルを作成し、`GOOGLE_API_KEY=...` (または使用するAIのキー) を設定します。
        2.  フェーズ2で生成済みの `instructions.json` と `anchor_image.png` を用意します。
      * **実行:** `test_ai_connector.py` を作成し、`ai_connector.build_prompt(...)` と `ai_connector.generate_mermaid_code(...)` を順に呼び出し、最終的なMermaidコードを `print()` します。
      * **確認 (OK条件):**
        1.  コンソールに `graph TD ...` で始まるMermaidコードが出力されること。
        2.  `instructions.json` に含まれていた `id` と `text` が、Mermaidのノード定義（例: `node_001(["開始"])`）として含まれていること。
        3.  `anchor_image.png` に描画されていた矢印が、Mermaidの接続（例: `node_001 --> node_002`）として含まれていること。
  - [x] **Git:** `git add ai_connector.py test_ai_connector.py .env.example` `git commit -m "feat(ai): Implement Module 3 for AI prompt generation and API call"`
  - [x] **クリーンアップ:** `rm test_parser.py test_asset_gen.py test_ai_connector.py` `git rm test_chart_simple.xlsx` `git commit -m "chore: Remove temporary test files"`

-----

## フェーズ4: メイン処理 (`main.py`) と E2Eテスト

  - [ ] **実装:** `main.py` ファイルを作成します。`argparse` と、作成した3つのモジュール (`excel_parser`, `asset_generator`, `ai_connector`) をインポートします。
  - [ ] **実装:** `argparse` を設定し、`--file` (必須), `--sheet` (必須), `--output` (デフォルト: `output.md`) のコマンドライン引数を受け取れるようにします。
  - [ ] **実装:** `main` 関数内で、以下のオーケストレーションを実行します。
    1.  `excel_parser.parse_excel_shapes(file, sheet)` を呼び出します。
    2.  `asset_generator.generate_assets(...)` を呼び出します。（中間ファイル名は `output/instructions.json`, `output/anchor_image.png` のように固定でOK）
    3.  `ai_connector.build_prompt(...)` を呼び出します。
    4.  `ai_connector.generate_mermaid_code(...)` を呼び出し、最終的な `mermaid_code` を取得します。
    5.  `mermaid_code` を `mermaid\n ... \n` で囲み、`--output` で指定されたパスにファイルとして書き込みます。
  - [ ] **動作確認 (E2Eテスト):**
      * **準備:**
        1.  仕様書 (6.) にあるような、分岐 (ひし形) と "Yes"/"No" ラベル（テキストボックスで矢印の横に配置）を含む、より複雑な `test_complex_chart.xlsx` を手動で作成します。
      * **実行:** コマンドラインから `python main.py --file test_complex_chart.xlsx --sheet Sheet1 --output test_complex_output.md` を実行します。
      * **確認 (OK条件):**
        1.  `test_complex_output.md` が生成されること。
        2.  `anchor_image.png` を目視確認し、「Yes」「No」のラベルがマスキングされずに *残っている* ことを確認します。
        3.  `test_complex_output.md` の中身をMermaid Live EditorやVSCodeプレビューで確認します。
        4.  元のExcelのフローチャート（ノードのテキスト、形状、矢印）が正しく再現されていること。
        5.  **最重要:** ひし形からの矢印に `-->|"Yes"|` や `-->|"No"|` のように、AIがアンカー画像の分岐ラベルを正しく読み取り、Mermaidに反映していること。
  - [ ] **Git:** `git add main.py test_complex_chart.xlsx` `git commit -m "feat(main): Implement main orchestration script and pass E2E test"`

-----

## フェーズ5: 仕上げ

  - [ ] **実装:** `README.md` を作成します。ツールの目的、必要なAPIキー (`.env.example` を参照)、インストール方法 (`pip install -r requirements.txt`)、実行方法 (`python main.py --file ... --sheet ...`) を簡潔に記載します。
  - [ ] **実装:** `pip freeze > requirements.txt` を実行し、`requirements.txt` の内容を正確なバージョンに更新します。
  - [ ] **Git:** `git add README.md requirements.txt` `git commit -m "docs: Add README and finalize requirements.txt"`
  - [ ] **完了:** 全てのタスクが完了しました。