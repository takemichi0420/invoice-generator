# invoice-generator
納品書（.xlsx）を集計して、書式付き請求書を自動生成するPythonツールです。

特長
	主な機能
	•	指定フォルダ内の納品書（.xlsx）ファイルを自動で検索
	•	品目・単価・金額などの明細を自動集計
	•	小計・税計算を含む請求書を出力（テンプレート形式）
	•	出力ファイル名に日付と取引先名を付与
	•	複数シート・複数ファイルに対応
	•	書式付きのExcel請求書を生成（交互配色・通貨書式など）

Python version
	3.13.3

ライブラリー
	•	pandas:2.2.3
	•	openpyxl:3.1.5
    •	numpy:2.2.5

ディレクトリ構成
invoice-generator/
├── input/             # 納品書ファイル置き場（.xlsx）
├── output/            # 自動生成される請求書
├── log/               # 売上ログファイル（任意）
├── template/          # 請求書テンプレート
├── temp/              # 一時処理用ファイル
├── src/
│   └── generate_invoice.py  # メインスクリプト
├── requirements.txt
├── .gitignore
└── README.md

・実行コマンド
python src/generate_invoice.py \
  --client_name "XXX" \
  --delivery_date_cell "E1"

  	•	--client_name：ファイル名・宛名・売上一覧の一致文字列
	•	--delivery_date_cell：納品日のセル位置（例：E1）

補足
	•	空白や表記ゆれにも対応（NFKC正規化）
	•	不要シートは自動削除
	•	金額・単価には通貨フォーマットが自動適用されます
