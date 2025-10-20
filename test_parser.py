"""
Excel解析モジュールのテストスクリプト
"""
import excel_parser

# テスト対象のExcelファイル
TEST_FILE = "test_chart_simple.xlsx"
SHEET_NAME = "Sheet1"

def main():
    print(f"Testing excel_parser with file: {TEST_FILE}")
    print(f"Sheet: {SHEET_NAME}")
    print("=" * 50)

    try:
        # Excel解析を実行
        mapped_containers = excel_parser.parse_excel_shapes(TEST_FILE, SHEET_NAME)

        print(f"\nTotal containers found: {len(mapped_containers)}")
        print("=" * 50)

        # 結果を表示
        for idx, container in enumerate(mapped_containers, 1):
            print(f"\n[Container {idx}]")
            print(f"  Temp ID: {container['temp_id']}")
            print(f"  Text: '{container['text']}'")
            print(f"  Shape Type: {container['shape_type']}")
            print(f"  Position: top={container['position']['top']:.1f}, "
                  f"left={container['position']['left']:.1f}, "
                  f"width={container['position']['width']:.1f}, "
                  f"height={container['position']['height']:.1f}")

        # 検証条件を確認
        print("\n" + "=" * 50)
        print("Verification:")

        if len(mapped_containers) == 3:
            print("✓ Found 3 containers (as expected)")
        else:
            print(f"✗ Expected 3 containers, but found {len(mapped_containers)}")

        # 各コンテナのテキストを確認
        expected_texts = ["開始", "処理A", "処理Bのラベル"]
        for idx, expected_text in enumerate(expected_texts):
            if idx < len(mapped_containers):
                actual_text = mapped_containers[idx]['text']
                if expected_text in actual_text:
                    print(f"✓ Container {idx+1} has text containing '{expected_text}'")
                else:
                    print(f"✗ Container {idx+1} expected text '{expected_text}', but got '{actual_text}'")

    except FileNotFoundError:
        print(f"\n✗ Error: File '{TEST_FILE}' not found.")
        print("Please create the test Excel file according to the task specification:")
        print("  1. A rounded rectangle with text '開始' inside")
        print("  2. A rectangle with text '処理A' inside")
        print("  3. A rectangle (with border) and a text box '処理Bのラベル' placed on top")
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
