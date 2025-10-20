"""
資材生成モジュールのテストスクリプト
"""
import excel_parser
import asset_generator
import os
from PIL import Image, ImageDraw


def main():
    TEST_FILE = "test_chart_simple.xlsx"
    SHEET_NAME = "Sheet1"

    print("Testing asset_generator...")
    print("=" * 60)

    # Step 1: Parse Excel
    print("\n[Step 1] Parsing Excel file...")
    mapped_containers = excel_parser.parse_excel_shapes(TEST_FILE, SHEET_NAME)
    print(f"✓ Parsed {len(mapped_containers)} containers")

    for idx, container in enumerate(mapped_containers, 1):
        print(f"  Container {idx}: '{container['text']}'")

    # Step 2: Generate JSON instructions
    print("\n[Step 2] Generating JSON instructions...")
    os.makedirs("output", exist_ok=True)
    json_data = asset_generator.generate_json_instructions(
        mapped_containers,
        "output/instructions.json"
    )

    print(f"✓ Generated JSON with {len(json_data)} nodes")

    # Step 3: Create a dummy screenshot for testing
    print("\n[Step 3] Creating dummy screenshot for testing...")
    # 800x600の白い画像を作成
    dummy_img = Image.new('RGB', (800, 600), color='lightgray')
    draw = ImageDraw.Draw(dummy_img)

    # いくつかの図形を描画（テスト用）
    # ダミーの図形を描いて、実際のExcelレイアウトを模擬
    for node in json_data:
        pos = node["position"]
        x1 = pos["left"]
        y1 = pos["top"]
        x2 = x1 + pos["width"]
        y2 = y1 + pos["height"]

        # 図形を描画（青色の四角形）
        draw.rectangle([x1, y1, x2, y2], fill="lightblue", outline="blue", width=3)

        # テキストを描画
        draw.text((x1 + 10, y1 + 10), node["text"], fill="black")

    dummy_screenshot_path = "output/temp_dummy_screenshot.png"
    dummy_img.save(dummy_screenshot_path)
    print(f"✓ Dummy screenshot created: {dummy_screenshot_path}")

    # Step 4: Generate anchor image
    print("\n[Step 4] Generating anchor image...")
    asset_generator.generate_anchor_image(
        dummy_screenshot_path,
        json_data,
        "output/anchor_image.png"
    )

    # Cleanup
    if os.path.exists(dummy_screenshot_path):
        os.remove(dummy_screenshot_path)

    # Verification
    print("\n" + "=" * 60)
    print("Verification:")

    if os.path.exists("output/instructions.json"):
        print("✓ instructions.json created")
        import json
        with open("output/instructions.json", 'r', encoding='utf-8') as f:
            data = json.load(f)
            print(f"  - Contains {len(data)} nodes")
            for node in data:
                print(f"    * {node['id']}: '{node['text']}'")
    else:
        print("✗ instructions.json not found")

    if os.path.exists("output/anchor_image.png"):
        print("✓ anchor_image.png created")
        img = Image.open("output/anchor_image.png")
        print(f"  - Size: {img.size}")
    else:
        print("✗ anchor_image.png not found")

    print("\n" + "=" * 60)
    print("✓ Asset generation test complete!")


if __name__ == "__main__":
    main()
