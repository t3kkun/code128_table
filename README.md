# code128_table
generate a table(name, barcode_1, barcode_2...) (PDF)

# バーコード表作成プログラム
CSVファイルを作成後実行すればすぐ規定フォーマットのバーコード表が作成されます

## 機能
- Code128にて数値からバーコードを生成
- A4のPDFにて表を自動生成
  
## 動作環境
- Python 3.x
- ライブラリはrequirement.txtを参照

## 使用方法
ご用意いただくものは2つです．
- 日本語フォント
- データの入ったCSVファイル(codes.csv)
  
[NotoSansCJKjp-VF.ttf](https://raw.githubusercontent.com/notofonts/noto-cjk/main/Sans/Variable/TTF/NotoSansCJKjp-VF.ttf)を想定．

他のフォントでも動くと思いますが，font_pathを変更するか"NotoSansCJKjp-VF.ttf"へリネームしていただければと思います

codes.csvの書き方はサンプルファイルを参照してください．2行目から変換されます．

## ライセンス
- MIT License
