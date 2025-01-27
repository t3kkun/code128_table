import pandas as pd
from barcode import Code128
from barcode.writer import ImageWriter
import os
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from datetime import datetime
# TO USE, RUN "pip install python-barcode pandas, reportlab"

# フォントの登録
font_path = "NotoSansCJKjp-VF.ttf"  # 日本語フォントファイルのパス
pdfmetrics.registerFont(TTFont("NotoSans", font_path))

# 日本語用スタイルの作成
japanese_style = ParagraphStyle(
    name="Japanese",
    fontName="NotoSans",
    fontSize=18,
    leading=12,
    alignment=1,  # 中央寄せ
)

# CSVファイルを読み込む
csv_file = "codes.csv"
df = pd.read_csv(csv_file)

# 出力日付と時刻を取得（システム時）
current_datetime = datetime.now().strftime("%Y/%m/%d %H:%M")

# テーブルデータの初期化
header_row = ['Name', 'Record1', 'Recode2', 'Barcode1', 'Barcode2']

# バーコード画像のディレクトリ
barcode_dir = "output_images"
os.makedirs(barcode_dir, exist_ok=True)

# A列とB列のデータを取得
a_column = df.columns[0]
b_column = df.columns[1]

# カスタムImageWriterの設定
options = {
    'module_width': 0.3,  # 各バーの幅
    'module_height': 15.0,  # 各バーの高さ
    'quiet_zone': 6.5,  # バーコードの周囲の空白領域
    'font_size': 10,  # フォントサイズ
    'text_distance': 5.0,  # テキストとバーコードの間の距離
    'background': 'white',  # 背景色
    'foreground': 'black',  # 前景色
    'write_text': True,  # バーコードの下にテキストを表示
}

# 各コードをPNG画像として保存します
for idx, row in df.iterrows():
    a_code = str(row[a_column]).strip()  # 前後のスペースを除去
    b_code = str(row[b_column]).strip()  # 前後のスペースを除去

    if a_code.isdigit() and (len(a_code) == 6 or len(a_code) == 8):
        # A列のコードをバーコードにして保存
        barcode = Code128(a_code, writer=ImageWriter())
        output_file = os.path.join(barcode_dir, f'{a_code}_1')
        barcode.save(output_file, options=options)  # 拡張子は自動的に追加されます
        print(f'Saved {output_file}.png')
    else:
        print(f'Invalid A code at index {idx}: {a_code}')

    if b_code.isdigit() and (len(b_code) == 6 or len(b_code) == 8):
        # B列のコードをバーコードにして保存
        barcode = Code128(b_code, writer=ImageWriter())
        output_file = os.path.join(barcode_dir, f'{a_code}_2')
        barcode.save(output_file, options=options)  # 拡張子は自動的に追加されます
        print(f'Saved {output_file}.png')
    else:
        print(f'Invalid B code at index {idx}: {b_code}')


records = []
for _, row in df.iterrows():
    name = str(row['C列'])  # Name
    a_code = str(row['A列'])  # Record1
    b_code = str(row['B列'])  # Recode2

    id_barcode_file = os.path.join(barcode_dir, f"{a_code}_1.png")
    pass_barcode_file = os.path.join(barcode_dir, f"{a_code}_2.png")

    # 画像が存在しない場合のチェック
    if not os.path.exists(id_barcode_file) or not os.path.exists(pass_barcode_file):
        raise FileNotFoundError(f"バーコード画像が見つかりません: {id_barcode_file}, {pass_barcode_file}")

    # レコードを追加
    records.append([
        Paragraph(name, japanese_style),  # Name
        Paragraph(a_code, japanese_style),  # Record1
        Paragraph(b_code, japanese_style),  # Record2
        Image(id_barcode_file, width=120, height=60),  # Barcode1
        Image(pass_barcode_file, width=120, height=60),  # Barcode2
    ])

# レコードをページ単位で分割
max_rows_per_page = 10  # 1ページあたりの行数
page_data = []

while records:
    page = records[:max_rows_per_page]  # 最大行を取得
    # 空白行を追加して満たす
    while len(page) < max_rows_per_page:
        page.append([' \n\n\n\n ', ' ', ' ', ' ', ' '])
    page_data.append(page)
    records = records[max_rows_per_page:]

# PDF出力
output_pdf = "output.pdf"
doc = SimpleDocTemplate(output_pdf, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)

story = []

# 各ページを作成
for page in page_data:
    # テーブルデータを作成（ヘッダーを含む）
    table_data = [header_row] + page

    # テーブルを作成
    table = Table(table_data, colWidths=[130, 80, 90, 130, 130])

    # テーブルスタイルを設定
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # ヘッダ背景色
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),      # ヘッダ文字色
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),             # 中央寄せ
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),            # 上下中央寄せ
        ('FONTNAME', (0, 0), (-1, -1), 'NotoSans'),        # 日本語対応フォント
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),     # グリッド線
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),            # 下余白
        ('TOPPADDING', (0, 0), (-1, -1), 6),               # 上余白
        ('LEFTPADDING', (0, 0), (-1, -1), 6),              # 左余白
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),             # 右余白
    ]))

    story.append(table)

# 出力日付を描画する関数
def add_header(canvas, doc):
    canvas.saveState()
    canvas.setFont("NotoSans", 16)
    canvas.drawString(A4[0] - 200, A4[1] - 30, f"出力日: {current_datetime}")  # 右上に日付を描画
    canvas.restoreState()

# PDFを生成
doc.build(story, onFirstPage=add_header, onLaterPages=add_header)

print(f"PDFが生成されました: {output_pdf}")
