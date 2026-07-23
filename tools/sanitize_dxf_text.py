import re

def main(doc):
    # モデルスペースを取得
    msp = doc.modelspace()

    # TEXTエンティティの内容をクリーンアップ
    cleaned_texts = []
    for ent in msp.query('TEXT'):
        raw_text = ent.dxf.text
        # 半角スペース、全角スペース、TAB、改行を除去
        cleaned = re.sub(r'[ \u3000\t\r\n]+', '', raw_text)
        # デバッグ用
        cleaned_texts.append(cleaned)
        # TEXTエンティティの内容を更新
        ent.dxf.text = cleaned

    return doc

if __name__ == "__main__":
    pass