
def main(doc, exclude_layers=[]):
    """LWPOLYLINをLINEに変換し、他の種類のエンティティは操作・削除しない。
    また、 exclude_layers に指定したレイヤーのポリラインは分解しない。"""

    msp = doc.modelspace()

    # polyline を line へ変換
    for polyline in msp.query("LWPOLYLINE"):
        if polyline.dxf.layer in exclude_layers:
            continue

        points = [(pt[0], pt[1]) for pt in polyline.get_points()]

        # 2025/10/29 @本間 refactor
        attribs = {
            'layer': polyline.dxf.layer,
            'color': polyline.dxf.color,
            'linetype': polyline.dxf.linetype
        }

        # 2025/10/29 @本間 refactor
        # 各連続する点のペアに対してLINEを作成
        for i in range(len(points) - 1):
            msp.add_line(points[i], points[i+1], dxfattribs=attribs)

        # 2025/10/29 @本間 feat
        # 閉じたポリラインなら最後と最初を結ぶ
        # get_points() は閉じたポリラインでも 最初の点を最後に重複して返さない仕様なので この操作が必要
        if polyline.closed:
            msp.add_line(points[-1], points[0], dxfattribs=attribs)

        # 2025/10/29 @本間 comment_out:
        """
        # 各連続する点のペアに対してLINEを作成
        for i in range(len(points) - 1):
            p_dxf = polyline.dxf
            if len(points) == 2:
                msp.add_line(points[i], points[i+1], dxfattribs={
                        'layer': p_dxf.layer, 'color': p_dxf.color, 'linetype': p_dxf.linetype, 'handle': p_dxf.handle})
            else:
                msp.add_line(points[i], points[i+1], dxfattribs={
                        'layer': p_dxf.layer, 'color': p_dxf.color, 'linetype': p_dxf.linetype})
        """

        # polyline を削除
        msp.delete_entity(polyline)

    return doc


if __name__ == "__main__":
    pass
