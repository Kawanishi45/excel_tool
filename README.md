# Excel-Mermaid変換ツール

Microsoft Excelで作成されたフローチャートを、AI技術を活用して高精度なMermaid記法のテキストデータに変換するツールです。

## 概要

このツールは、Excelの複雑なフローチャート構造（図形とテキストが分離されているケース）に対応するため、以下のハイブリッドアプローチを採用しています：

1. **[Python] データ抽出とマッピング**: Excelファイルから全シェイプの座標情報とテキスト情報を抽出し、座標演算で紐付け
2. **[Python] AI用資材生成**: ID付きアンカー画像とJSON指示書を生成
3. **[AI] 構造認識**: AIは構造（矢印）と分岐ラベル（Yes/No）の認識に集中
4. **[出力] Mermaidコード生成**: 高精度なMermaid記法のフローチャートを出力

## 必要な環境

- Python 3.10以上
- Microsoft Excel（xlwingsによる自動操作のため）
- Google Gemini API Key

## セットアップ

### 1. リポジトリのクローン

```bash
git clone <repository-url>
cd excel_tool
```

### 2. 仮想環境の構築

```bash
python -m venv venv
source venv/bin/activate  # Windowsの場合: venv\\Scripts\\activate
```

### 3. 依存ライブラリのインストール

```bash
pip install -r requirements.txt
```

### 4. API キーの設定

`.env` ファイルを作成し、Google Gemini APIキーを設定します：

```bash
cp .env.example .env
```

`.env` ファイルを編集し、以下を設定：

```
GOOGLE_API_KEY=あなたのAPIキー
```

**APIキーの取得**: https://makersuite.google.com/app/apikey

## 使用方法

### 基本的な使い方

```bash
python main.py --file <Excelファイル> --sheet <シート名> --output <出力ファイル>
```

### 実行例

```bash
python main.py --file test_chart_simple.xlsx --sheet Sheet1 --output output.md
```

### オプション

- `--file` (必須): 変換するExcelファイルのパス
- `--sheet` (必須): 対象のシート名
- `--output` (オプション): 出力ファイル名（デフォルト: `output.md`）
- `--keep-intermediate`: 中間ファイル（JSON、画像）を保持する

## 出力の確認

生成されたMermaidコードは以下の方法で確認できます：

1. **Mermaid Live Editor**: https://mermaid.live
2. **VS Code**: Markdown Preview Enhanced 拡張機能
3. **GitHub/GitLab**: Markdownファイル内のMermaidコードブロックは自動レンダリングされます

## プロジェクト構成

```
excel_tool/
├── main.py                 # メインスクリプト（オーケストレーション）
├── excel_parser.py         # モジュール1: Excel解析・座標マッピング
├── asset_generator.py      # モジュール2: AI用資材生成
├── ai_connector.py         # モジュール3: AI連携・Mermaidコード生成
├── requirements.txt        # 依存ライブラリ一覧
├── .env.example           # 環境変数テンプレート
├── README.md              # このファイル
└── output/                # 中間ファイル・出力先ディレクトリ
```

## トラブルシューティング

### Q1. `GOOGLE_API_KEY not found` エラーが出る

`.env` ファイルが作成されているか、APIキーが正しく設定されているか確認してください。

### Q2. Excelファイルが開けない

- xlwingsはMicrosoft Excelがインストールされている必要があります
- Excelファイルのパスが正しいか確認してください
- Excelファイルが既に開かれている場合は閉じてください

### Q3. 図形のテキストが正しく認識されない

- 座標マッピングのロジックは、テキストボックスの中心座標がコンテナ図形内に含まれるかで判定しています
- テキストボックスが図形から大きくずれている場合は、手動で調整してください

### Q4. APIリクエストがタイムアウトする

- ネットワーク接続を確認してください
- デフォルトのタイムアウトは60秒です。`ai_connector.py`の`timeout`パラメータで調整可能です

## ライセンス

MIT License

## 開発者向け情報

詳細な開発仕様や実装アプローチについては以下のドキュメントを参照してください：

- `dev_spec.md`: 詳細な機能仕様書
- `dev_approach.md`: ハイブリッドアプローチの技術詳細
- `dev_tasks.md`: 開発タスクリスト

## サポート

問題が発生した場合は、GitHubのIssuesにて報告してください。
