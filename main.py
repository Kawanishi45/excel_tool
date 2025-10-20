"""
メイン処理
ExcelフローチャートをMermaid記法に変換するツールのエントリーポイント
"""
import argparse
import os
import sys

# 自作モジュールをインポート
import excel_parser
import asset_generator
import ai_connector


def main():
    """メイン処理のオーケストレーション"""
    # コマンドライン引数を解析
    parser = argparse.ArgumentParser(
        description="Convert Excel flowchart to Mermaid diagram"
    )
    parser.add_argument(
        "--file",
        required=True,
        help="Path to Excel file (*.xlsx)"
    )
    parser.add_argument(
        "--sheet",
        required=True,
        help="Sheet name to process"
    )
    parser.add_argument(
        "--output",
        default="output.md",
        help="Output markdown file path (default: output.md)"
    )
    parser.add_argument(
        "--keep-intermediate",
        action="store_true",
        help="Keep intermediate files (JSON and anchor image)"
    )

    args = parser.parse_args()

    # ファイルの存在確認
    if not os.path.exists(args.file):
        print(f"✗ Error: File not found: {args.file}")
        sys.exit(1)

    print("=" * 70)
    print("Excel to Mermaid Converter")
    print("=" * 70)
    print(f"Input file: {args.file}")
    print(f"Sheet name: {args.sheet}")
    print(f"Output file: {args.output}")
    print("=" * 70)

    # 中間ファイルの保存先
    intermediate_dir = "output"
    os.makedirs(intermediate_dir, exist_ok=True)

    json_path = os.path.join(intermediate_dir, "instructions.json")
    image_path = os.path.join(intermediate_dir, "anchor_image.png")

    try:
        # ステップ1: Excel解析
        print("\n[Step 1/4] Parsing Excel shapes...")
        mapped_containers = excel_parser.parse_excel_shapes(args.file, args.sheet)
        print(f"✓ Parsed {len(mapped_containers)} shapes")

        # ステップ2: 資材生成
        print("\n[Step 2/4] Generating AI input assets...")
        json_data, _ = asset_generator.generate_assets(
            mapped_containers,
            args.file,
            args.sheet,
            json_path,
            image_path
        )
        print(f"✓ Generated JSON: {json_path}")
        print(f"✓ Generated image: {image_path}")

        # ステップ3: AI連携
        print("\n[Step 3/4] Calling AI to generate Mermaid code...")
        print("Note: This requires GOOGLE_API_KEY in .env file")

        # APIキーの確認
        if not os.environ.get('GOOGLE_API_KEY'):
            print("\n⚠ Warning: GOOGLE_API_KEY not found!")
            print("Please create a .env file with your Google Gemini API key.")
            print("Example: cp .env.example .env")
            print("\nFor now, skipping AI generation step.")
            print("You can manually use the generated files:")
            print(f"  - JSON: {json_path}")
            print(f"  - Image: {image_path}")

            # ダミーのMermaidコードを生成
            mermaid_code = _generate_dummy_mermaid(json_data)
            print("\n✓ Generated dummy Mermaid code (without AI)")

        else:
            mermaid_code = ai_connector.generate_mermaid_code(json_path, image_path)
            print("✓ Mermaid code generated successfully")

        # ステップ4: Markdownファイルに保存
        print("\n[Step 4/4] Saving to output file...")
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write("# Flowchart (Generated from Excel)\n\n")
            f.write("```mermaid\n")
            f.write(mermaid_code)
            f.write("\n```\n")

        print(f"✓ Saved to: {args.output}")

        # 中間ファイルの削除（オプション）
        if not args.keep_intermediate:
            print("\nCleaning up intermediate files...")
            if os.path.exists(json_path):
                os.remove(json_path)
                print(f"  - Removed: {json_path}")
            if os.path.exists(image_path):
                os.remove(image_path)
                print(f"  - Removed: {image_path}")
        else:
            print("\nIntermediate files kept:")
            print(f"  - JSON: {json_path}")
            print(f"  - Image: {image_path}")

        # 完了
        print("\n" + "=" * 70)
        print("✓ Conversion complete!")
        print("=" * 70)
        print(f"\nYou can preview the result at: {args.output}")
        print("Recommended tools:")
        print("  - Mermaid Live Editor: https://mermaid.live")
        print("  - VS Code: Markdown Preview Enhanced extension")

    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def _generate_dummy_mermaid(json_data):
    """
    APIキーがない場合のダミーMermaidコード生成

    Args:
        json_data (list): JSON指示書データ

    Returns:
        str: ダミーのMermaidコード
    """
    lines = ["graph TD"]

    for node in json_data:
        node_id = node["id"]
        text = node["text"]

        # シンプルなノード定義（四角形のみ）
        lines.append(f'    {node_id}["{text}"]')

    # 順番に接続（ダミー）
    for i in range(len(json_data) - 1):
        lines.append(f'    {json_data[i]["id"]} --> {json_data[i+1]["id"]}')

    return '\n'.join(lines)


if __name__ == "__main__":
    main()
