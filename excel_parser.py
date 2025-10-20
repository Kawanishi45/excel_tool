"""
Excel解析モジュール
Excelファイルの複雑な構造を解析し、「どの図形が、どのテキストを持つか」を正確に特定する。

注意: macOS版xlwingsではシェイプのテキスト取得に制限があるため、
XML解析を使用してシェイプの情報を取得する。
"""
import zipfile
import xml.etree.ElementTree as ET


# Excel DrawingML名前空間
NAMESPACES = {
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
}


def parse_excel_shapes(file_path, sheet_name):
    """
    指定されたExcelファイルの指定シートから、すべてのシェイプ情報を抽出し、
    座標ベースで「コンテナ図形」と「テキスト」を紐付ける。

    Args:
        file_path (str): Excelファイルのパス
        sheet_name (str): 処理対象のシート名

    Returns:
        list: テキスト情報がマッピングされたコンテナ図形のリスト
    """
    # XMLから全シェイプ情報を取得
    all_shapes = _get_all_shapes_from_xml(file_path)

    # シェイプを役割ごとに分類
    container_shapes, text_shapes = _classify_shapes(all_shapes)

    # 座標マッピングを実行
    mapped_containers = _map_text_to_containers(container_shapes, text_shapes)

    return mapped_containers


def _get_all_shapes_from_xml(file_path):
    """
    ExcelファイルのXMLから全シェイプをループ処理し、必要な情報を抽出する。

    Args:
        file_path (str): Excelファイルのパス

    Returns:
        list: 全シェイプの情報を含む辞書のリスト
    """
    all_shapes = []

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        # drawingファイルを取得
        drawing_files = [name for name in zip_ref.namelist()
                        if 'xl/drawings/drawing' in name and name.endswith('.xml')]

        for drawing_file in drawing_files:
            content = zip_ref.read(drawing_file)
            root = ET.fromstring(content)

            # シェイプ要素を抽出 (sp: shape, txSp: text shape, cxnSp: connector shape)
            shape_elements = (
                root.findall('.//xdr:sp', NAMESPACES) +
                root.findall('.//xdr:txSp', NAMESPACES) +
                root.findall('.//xdr:cxnSp', NAMESPACES)
            )

            for idx, shape_elem in enumerate(shape_elements):
                temp_id = f"temp_{idx:03d}"

                # テキスト情報を取得
                text = _extract_text_from_shape(shape_elem)

                # 座標情報を取得
                position = _extract_position_from_shape(shape_elem, root)

                # シェイプタイプを判定
                shape_type = _determine_shape_type(shape_elem)

                all_shapes.append({
                    "temp_id": temp_id,
                    "text": text,
                    "position": position,
                    "shape_type": shape_type,
                    "_xml_element": shape_elem  # デバッグ用
                })

    return all_shapes


def _extract_text_from_shape(shape_elem):
    """
    シェイプ要素からテキストを抽出する。

    Args:
        shape_elem: XMLシェイプ要素

    Returns:
        str: 抽出されたテキスト
    """
    text_parts = []

    # テキスト要素を探す (a:t タグ)
    for t_elem in shape_elem.findall('.//a:t', NAMESPACES):
        if t_elem.text:
            text_parts.append(t_elem.text)

    return ''.join(text_parts)


def _extract_position_from_shape(shape_elem, root):
    """
    シェイプ要素から座標情報を抽出する。

    Args:
        shape_elem: XMLシェイプ要素
        root: XMLルート要素

    Returns:
        dict: 座標情報 {top, left, width, height}
    """
    # シェイプの親要素を探す (アンカー情報を含む)
    parent = None
    for anchor_type in ['xdr:twoCellAnchor', 'xdr:oneCellAnchor', 'xdr:absoluteAnchor']:
        for anchor in root.findall(f'.//{anchor_type}', NAMESPACES):
            if shape_elem in list(anchor.iter()):
                parent = anchor
                break
        if parent is not None:
            break

    if parent is None:
        return {"top": 0, "left": 0, "width": 0, "height": 0}

    # twoCellAnchorの場合
    from_elem = parent.find('.//xdr:from', NAMESPACES)
    to_elem = parent.find('.//xdr:to', NAMESPACES)

    if from_elem is not None and to_elem is not None:
        # 開始位置
        from_col = int(from_elem.findtext('xdr:col', '0', NAMESPACES))
        from_col_off = int(from_elem.findtext('xdr:colOff', '0', NAMESPACES))
        from_row = int(from_elem.findtext('xdr:row', '0', NAMESPACES))
        from_row_off = int(from_elem.findtext('xdr:rowOff', '0', NAMESPACES))

        # 終了位置
        to_col = int(to_elem.findtext('xdr:col', '0', NAMESPACES))
        to_col_off = int(to_elem.findtext('xdr:colOff', '0', NAMESPACES))
        to_row = int(to_elem.findtext('xdr:row', '0', NAMESPACES))
        to_row_off = int(to_elem.findtext('xdr:rowOff', '0', NAMESPACES))

        # EMU (English Metric Units) をポイントに変換
        # 1 EMU = 1/914400 inch, 1 point = 1/72 inch
        # よって、1 point = 12700 EMU
        EMU_PER_POINT = 12700

        # 概算の座標計算（セル幅/高さは仮に72ポイントとする簡易計算）
        CELL_WIDTH = 72  # ポイント
        CELL_HEIGHT = 18  # ポイント

        left = from_col * CELL_WIDTH + from_col_off / EMU_PER_POINT
        top = from_row * CELL_HEIGHT + from_row_off / EMU_PER_POINT
        width = (to_col - from_col) * CELL_WIDTH + (to_col_off - from_col_off) / EMU_PER_POINT
        height = (to_row - from_row) * CELL_HEIGHT + (to_row_off - from_row_off) / EMU_PER_POINT

        return {
            "top": top,
            "left": left,
            "width": width,
            "height": height
        }

    # oneCellAnchor or absoluteAnchorの場合
    ext_elem = parent.find('.//xdr:ext', NAMESPACES)
    if ext_elem is not None:
        cx = int(ext_elem.get('cx', '0'))
        cy = int(ext_elem.get('cy', '0'))

        EMU_PER_POINT = 12700
        width = cx / EMU_PER_POINT
        height = cy / EMU_PER_POINT

        # 開始位置
        if from_elem is not None:
            from_col = int(from_elem.findtext('xdr:col', '0', NAMESPACES))
            from_col_off = int(from_elem.findtext('xdr:colOff', '0', NAMESPACES))
            from_row = int(from_elem.findtext('xdr:row', '0', NAMESPACES))
            from_row_off = int(from_elem.findtext('xdr:rowOff', '0', NAMESPACES))

            CELL_WIDTH = 72
            CELL_HEIGHT = 18

            left = from_col * CELL_WIDTH + from_col_off / EMU_PER_POINT
            top = from_row * CELL_HEIGHT + from_row_off / EMU_PER_POINT
        else:
            left = 0
            top = 0

        return {
            "top": top,
            "left": left,
            "width": width,
            "height": height
        }

    return {"top": 0, "left": 0, "width": 0, "height": 0}


def _determine_shape_type(shape_elem):
    """
    シェイプ要素からタイプを判定する。

    Args:
        shape_elem: XMLシェイプ要素

    Returns:
        str: シェイプタイプ
    """
    # XMLタグ名からタイプを判定
    tag = shape_elem.tag.split('}')[-1] if '}' in shape_elem.tag else shape_elem.tag

    if tag == 'sp':
        return 'auto_shape'
    elif tag == 'txSp':
        return 'text_box'
    elif tag == 'cxnSp':
        return 'connector'
    else:
        return 'unknown'


def _classify_shapes(all_shapes):
    """
    全シェイプを「コンテナ図形」と「テキスト図形」に分類する。

    Args:
        all_shapes (list): 全シェイプの情報リスト

    Returns:
        tuple: (container_shapes, text_shapes)
    """
    container_shapes = []
    text_shapes = []

    for shape in all_shapes:
        shape_type = shape["shape_type"]

        # コネクタ（矢印）とテキストボックスは除外し、それ以外をコンテナとする
        if shape_type not in ['text_box', 'connector']:
            container_shapes.append(shape)

        # テキストを持つ、またはテキストボックスのシェイプをtext_shapesに追加
        if shape["text"].strip() != "" or shape_type == 'text_box':
            text_shapes.append(shape)

    return container_shapes, text_shapes


def _map_text_to_containers(container_shapes, text_shapes):
    """
    座標マッピング処理：コンテナ図形とテキスト図形を座標で紐付ける。

    Args:
        container_shapes (list): コンテナ図形のリスト
        text_shapes (list): テキスト図形のリスト

    Returns:
        list: テキストがマッピングされたコンテナ図形のリスト
    """
    # テキストシェイプのコピーを作成（重複防止のため）
    remaining_text_shapes = text_shapes.copy()

    for container in container_shapes:
        # ロジックA: 自己テキスト優先
        # コンテナ自身が有効なテキストを持っているか確認
        if container["text"].strip() != "":
            # 既にテキストを持っている場合はスキップ
            continue

        # ロジックB: 包含判定
        # コンテナの座標範囲を計算
        parent_x1 = container["position"]["left"]
        parent_y1 = container["position"]["top"]
        parent_x2 = parent_x1 + container["position"]["width"]
        parent_y2 = parent_y1 + container["position"]["height"]

        # テキストシェイプを走査
        for text_shape in remaining_text_shapes[:]:  # コピーをループ
            # 自分自身はスキップ
            if text_shape["temp_id"] == container["temp_id"]:
                continue

            # テキストシェイプの中心座標を計算
            child_center_x = text_shape["position"]["left"] + text_shape["position"]["width"] / 2
            child_center_y = text_shape["position"]["top"] + text_shape["position"]["height"] / 2

            # 包含判定：子の中心が親の範囲内にあるか
            if (parent_x1 < child_center_x < parent_x2 and
                parent_y1 < child_center_y < parent_y2):
                # 紐付け：コンテナのテキストを更新
                container["text"] = text_shape["text"]
                # 使用済みテキストシェイプを除外（重複防止）
                remaining_text_shapes.remove(text_shape)
                break  # 最初にマッチしたものを採用

    return container_shapes


if __name__ == "__main__":
    # テスト用コード
    import sys
    if len(sys.argv) >= 3:
        file_path = sys.argv[1]
        sheet_name = sys.argv[2]
        result = parse_excel_shapes(file_path, sheet_name)
        print("=== Parsed Shapes ===")
        for shape in result:
            print(f"ID: {shape['temp_id']}, Text: '{shape['text']}', Type: {shape['shape_type']}")
