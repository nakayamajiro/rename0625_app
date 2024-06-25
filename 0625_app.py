import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFilter, ImageEnhance
import sys
import pyocr
import pyocr.builders
import pdf2image
import fitz  # PyMuPDF
import os

# OCRツールの利用可能性を確認
tools = pyocr.get_available_tools()
if len(tools) == 0:
    print("OCRツールが見つかりません")
    sys.exit(1)

tool = tools[0]

# スクリプトが格納されているディレクトリの絶対パスを取得
script_directory = os.path.dirname(os.path.abspath(__file__))

# 'pdf' フォルダへの相対パスを指定
relative_directory = 'pdf'
pdf_folder_path = os.path.join(script_directory, relative_directory)

# 赤枠の座標
red_boxes = [
    (500, 973, 2490, 303),   # 例として x=500, y=973, width=2490, height=303 番組名
    #(3750, 1170, 1850, 200),  # 2番目の赤枠　日付
    #(5847, 376, 677, 114)    # 3番目の赤枠　　　フォーマットNo.
]

# 全角文字および括弧付き数字を半角文字に変換する関数
def zenkaku_to_hankaku(text):
    zenkaku = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ（）．，－／：；？！＃＆％＠"
    hankaku = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz().,-/:;?!#&%@"
    trans_table = str.maketrans(zenkaku, hankaku)
    text = text.translate(trans_table)

    # 括弧付き数字の変換
    bracketed_nums = {
        '①': '1', '②': '2', '③': '3', '④': '4', '⑤': '5',
        '⑥': '6', '⑦': '7', '⑧': '8', '⑨': '9', '⑩': '10',
        '⑪': '11', '⑫': '12', '⑬': '13', '⑭': '14', '⑮': '15',
        '⑯': '16', '⑰': '17', '⑱': '18', '⑲': '19', '⑳': '20'
    }
    for zenkaku_num, hankaku_num in bracketed_nums.items():
        text = text.replace(zenkaku_num, hankaku_num)

    return text.strip()  # 文字列の先頭と末尾の空白を削除

# PDFファイルを処理する関数
def process_pdf(pdf_path, lang):
    # PDFを画像オブジェクトに変換（一枚目のみ）
    images = pdf2image.convert_from_path(pdf_path, dpi=600, fmt='jpg', first_page=0, last_page=1)
    
    # テキスト結果を格納するリスト
    ocr_results = []

    # 画像オブジェクトから特定部分のテキストを取得
    image = images[0]
    draw = ImageDraw.Draw(image)
    
    for j, box in enumerate(red_boxes):
        # 画像の特定部分をクロップ
        cropped_image = image.crop((
            box[0],
            box[1],
            box[0] + box[2],
            box[1] + box[3]
        ))

        # 前処理: RGBカラー変換
        cropped_image = cropped_image.convert('RGB')
        
        # 前処理: コントラストの強調
        enhancer = ImageEnhance.Contrast(cropped_image)
        cropped_image = enhancer.enhance(2.0)
        
        # 前処理: バイナリ化
        threshold = 128
        cropped_image = cropped_image.point(lambda p: p > threshold and 255)
        
        # 前処理: ノイズ除去
        cropped_image = cropped_image.filter(ImageFilter.MedianFilter(size=3))

        # 元の画像に赤枠を描画
        draw.rectangle(
            [
                (box[0], box[1]),
                (box[0] + box[2], box[1] + box[3])
            ],
            outline="red",
            width=5
        )

        # 保存用のパスを生成
        cropped_image_path = f"/Users/ichiyanagishouba1/Desktop/rename/cropped_image_{j}.jpg"
        
        # クロップした画像を保存（高解像度を保持）
        cropped_image.save(cropped_image_path, dpi=(600, 600))

        # Tesseractの設定を調整
        custom_config = r'--oem 3 --psm 6'  # OCR Engine Mode 3とPage Segmentation Mode 6

        # クロップした部分のテキストを取得
        txt = tool.image_to_string(
            Image.open(cropped_image_path),  # 画像のファイルパスから画像を直接読み込む
            lang=lang,
            builder=pyocr.builders.TextBuilder(tesseract_layout=6)
        )
        
        # 全角文字および括弧付き数字を半角文字に変換
        txt = zenkaku_to_hankaku(txt)
        
        # 半角・全角スペースを削除
        txt = txt.replace(' ', '').replace('　', '')

        txt = txt.replace('?', '')
        
        # OCR結果をリストに追加
        ocr_results.append(txt)
        
        #print(f"赤枠部分{j+1}の抽出されたテキスト:")
        print(txt)

    # OCR結果を出力する例
    for idx, result in enumerate(ocr_results):
        #print(f"OCR結果 {idx + 1}: {result}")
        if result == '日曜報道丁HEPRIME':
            result = '日曜報道 THE PRIME'
        elif result == 'FNN工iュve。News。イット!':
            result = 'FNN Live News イット!'
        elif result == 'Liュve。News。イット!第1部':
            result = 'FNN Live News イット! 第1部'
        elif result == 'FNN工ive。News。days':
            result = 'FNN Live News days'
        elif result == 'ほかぼか':
            result = 'ぽかぽか'
        elif result == 'PNN土1ve。NNewSc':
            result = 'FNN Live News α'
        elif result == 'hNN土1ュve。NewSo':
            result = 'FNN Live News α'  
        elif result == 'めさざさまし8':
            result = 'めざまし8'
        elif result == 'Liュve。News。イット!第2部':
            result = 'FNN Live News イット! イット第2部'
            
    print(result)
    
    os.remove(cropped_image_path)

    # 日付を抽出する。
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    rect = fitz.Rect(450, 100, 650, 200)
    text = page.get_text("text", clip=rect)
    text = text.replace(' ', '').replace('　', '')
    text = text.replace('/', '月')
    #print("指定した領域の抽出されたテキスト:")
    print(text)

    # PDFファイルの名前を変更
# PDFファイルの名前を変更

    old_name = os.path.basename(pdf_path)
    new_name = f"{result}_{text}_{old_name}"
    new_path = os.path.join(pdf_folder_path, new_name)
    os.rename(pdf_path, new_path)
    print(f"ファイル名を変更しました: {old_name} -> {new_name}")

# 言語設定（日本語）
lang = 'jpn'

# tkinterのメインウィンドウを作成
root = tk.Tk()
root.title("PDF処理アプリ")

# PDF処理ボタンクリック時の処理
def on_process_pdf():
    try:
        # PDFフォルダ内のすべてのPDFファイルを処理
        for filename in os.listdir(pdf_folder_path):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(pdf_folder_path, filename)
                print(f"処理中のファイル: {pdf_path}")
                process_pdf(pdf_path, lang)
        
        messagebox.showinfo("完了", "PDF処理が完了しました。")

    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました: {str(e)}")

# PDF処理ボタンを作成
process_button = tk.Button(root, text="PDF処理を実行", command=on_process_pdf)
process_button.pack(pady=20)

# tkinterのメインループを開始
root.mainloop()
