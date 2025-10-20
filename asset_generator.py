"""
資材生成モジュール
モジュール1のマッピング結果に基づき、AIへの入力となる「JSON指示書」と「IDアンカー画像」を生成する。
"""
import json
import os
from PIL import Image, ImageDraw, ImageFont
import mss
import mss.tools


def generate_assets(mapped_containers, excel_file, sheet_name, json_out_path, image_out_path):
    """
    AIインプット資材を生成するメイン関数

    Args:
        mapped_containers (list): マッピング済みのコンテナ図形リスト
        excel_file (str): Excelファイルのパス
        sheet_name (str): シート名
        json_out_path (str): JSON出力パス
        image_out_path (str): 画像出力パス

    Returns:
        tuple: (json_data, image_path)
    """
    # JSON指示書を生成
    json_data = generate_json_instructions(mapped_containers, json_out_path)

    # スクリーンショットを取得
    screenshot_path = _get_chart_screenshot(excel_file, sheet_name)

    # IDアンカー画像を生成
    generate_anchor_image(screenshot_path, json_data, image_out_path)

    # 一時ファイルを削除
    if os.path.exists(screenshot_path):
        os.remove(screenshot_path)

    return json_data, image_out_path


def generate_json_instructions(mapped_containers, output_path):
    """
    モジュール1の出力からJSON指示書を生成する

    Args:
        mapped_containers (list): マッピング済みのコンテナ図形リスト
        output_path (str): JSON出力パス

    Returns:
        list: JSONデータ（Pythonのリスト形式）
    """
    json_data = []

    for idx, container in enumerate(mapped_containers, 1):
        node_id = f"node_{idx:03d}"

        node_entry = {
            "id": node_id,
            "text": container["text"],
            "shape_type": container["shape_type"],
            "position": container["position"]
        }

        json_data.append(node_entry)

    # JSONファイルとして保存
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)

    print(f"✓ JSON instructions saved: {output_path}")

    return json_data


def _get_chart_screenshot(file_path, sheet_name):
    """
    Excelファイルの指定シートのスクリーンショットを取得

    注意: このバージョンでは、Excelを手動で開いて全画面表示した状態で
    自動的にスクリーンショットを取得します。
    より高度な実装では、xlwingsを使ってExcelを自動操作することも可能です。

    Args:
        file_path (str): Excelファイルのパス
        sheet_name (str): シート名

    Returns:
        str: 保存されたスクリーンショットのパス
    """
    # 一時保存パス
    temp_path = "temp_screenshot.png"

    # mssを使用して画面全体のスクリーンショットを取得
    with mss.mss() as sct:
        # プライマリモニターの全画面をキャプチャ
        monitor = sct.monitors[1]  # 1 = プライマリモニター
        screenshot = sct.grab(monitor)

        # PIL Imageに変換して保存
        img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
        img.save(temp_path)

    print(f"✓ Screenshot captured: {temp_path}")
    print(f"  Note: Please ensure Excel file '{file_path}' (Sheet: '{sheet_name}') is visible on screen")

    return temp_path


def generate_anchor_image(original_image_path, instructions_json, output_path):
    """
    元のスクリーンショットにID情報を重ねた「IDアンカー画像」を生成

    Args:
        original_image_path (str): 元画像のパス
        instructions_json (list): JSON指示書データ
        output_path (str): 出力画像のパス
    """
    # 元画像を開く
    img = Image.open(original_image_path)
    draw = ImageDraw.Draw(img)

    # フォント設定（システムフォントを使用）
    try:
        # macOSの場合
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 24)
    except:
        try:
            # その他のシステム
            font = ImageFont.truetype("arial.ttf", 24)
        except:
            # デフォルトフォント
            font = ImageFont.load_default()

    # 各シェイプに対してマスキングとID描画を実行
    for node in instructions_json:
        pos = node["position"]
        node_id = node["id"]

        # 座標を取得
        left = pos["left"]
        top = pos["top"]
        width = pos["width"]
        height = pos["height"]

        # マスキング領域を定義
        x1, y1 = left, top
        x2, y2 = left + width, top + height

        # 白色で塗りつぶし（マスキング）
        draw.rectangle([x1, y1, x2, y2], fill="white", outline="black", width=2)

        # IDテキストを中心に描画
        # テキストのサイズを取得
        bbox = draw.textbbox((0, 0), node_id, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        # 中心座標を計算
        text_x = left + (width - text_width) / 2
        text_y = top + (height - text_height) / 2

        # テキストを描画
        draw.text((text_x, text_y), node_id, fill="black", font=font)

    # 画像を保存
    img.save(output_path)
    print(f"✓ Anchor image saved: {output_path}")


if __name__ == "__main__":
    # テスト用コード
    import excel_parser

    TEST_FILE = "test_chart_simple.xlsx"
    SHEET_NAME = "Sheet1"

    print("Testing asset_generator...")
    print("=" * 60)

    # Step 1: Parse Excel
    mapped_containers = excel_parser.parse_excel_shapes(TEST_FILE, SHEET_NAME)
    print(f"✓ Parsed {len(mapped_containers)} containers")

    # Step 2: Generate assets
    print("\nGenerating assets...")
    json_data, image_path = generate_assets(
        mapped_containers,
        TEST_FILE,
        SHEET_NAME,
        "output/instructions.json",
        "output/anchor_image.png"
    )

    print("\n" + "=" * 60)
    print("✓ Asset generation complete!")
    print(f"  JSON: output/instructions.json")
    print(f"  Image: output/anchor_image.png")
