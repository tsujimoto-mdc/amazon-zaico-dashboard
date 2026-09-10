# Amazon納品分 ZAICO在庫数 自動確認ダッシュボード

Amazonの納品プランExcelと、ZAICOのダウンロードCSVをアップロードして
ボタンを1回押すだけで、「本社ZAICO / 3FZAICO」列を追記したExcelが
ダウンロードできる社内向けWebツールです。

**元のデスクトップ版ツール（`zaico_tool.py`）と、やっている処理はまったく同じです。**
違いは「フォルダにファイルを置く」代わりに「ブラウザでアップロードする」点だけです。

---

## 1. フォルダ構成

```
amazon-zaico-dashboard/
├── app.py                 ← アプリ本体（画面）
├── zaico_core.py          ← 転記ロジック（元 zaico_tool.py の業務部分）
├── requirements.txt       ← 必要なライブラリ一覧
├── secrets_example.toml   ← Secrets設定のひな形
├── .gitignore
└── README.md
```

テンプレートファイルなどは不要です。アップロードされた納品プランExcelを
そのまま加工して返します。

---

## 2. デプロイの流れ（初回のみ）

### ステップ1: GitHubリポジトリを作成する
1. GitHub（[github.com](https://github.com)）でアカウントがなければ作成
2. 「New repository」で新規リポジトリを作成
   - **在庫・商品情報を扱うため、必ず Private（非公開）を選択してください**
3. このフォルダ一式（`app.py`, `zaico_core.py`, `requirements.txt`, `.gitignore`,
   `secrets_example.toml`, `README.md`）をアップロード
   - **`secrets_example.toml` はアップロードしてOK**（値はひな形なので問題ありません）
   - 実際の納品プランExcel・ZAICO CSV・生成後のファイルはアップロードしないでください
     （`.gitignore` で `*.xlsx` / `*.csv` を除外済みです）

### ステップ2: Streamlit Community Cloud に登録する
1. [share.streamlit.io](https://share.streamlit.io) にアクセスし、GitHubアカウントで連携
2. 「Create app」→ 作成したリポジトリ・ブランチ・`app.py` を選択してデプロイ

### ステップ3: 合言葉を設定する（任意）
社内限定URLだけで運用するなら不要です。合言葉を掛けたい場合のみ:

1. デプロイしたアプリの管理画面 → 「Settings」→「Secrets」を開く
2. 次の1行を貼り付けて保存

```toml
APP_PASSWORD = "実際の合言葉"
```

`APP_PASSWORD` を設定しなければ、合言葉入力なしでそのまま使えます。

---

## 3. 使い方

1. アプリのURLにアクセス（合言葉を設定していれば入力）
2. **① Amazon納品プランExcel(.xlsx)** をアップロード（複数まとめて処理できます）
3. **② ZAICOダウンロードCSV(.csv)** を1つアップロード
4. **③「🚀 実行」** を押す（処理状況がリアルタイムで表示されます）
5. **④ ダウンロード** から、`〇〇_ZAICO数量記載.xlsx` を保存
   （複数ファイルを処理した場合はZIPでまとめてダウンロードもできます）

画面には確認用の一覧表も表示され、在庫合計が納品数量に満たないSKUや、
ZAICOに型番が見つからなかったSKUが色付きで分かります
（**この判定は画面表示だけのもので、ダウンロードするExcelの中身には影響しません**）。

---

## 4. 処理の中身（元ツールと同じ）

- 納品プランExcelの **「SKU」列** と、ZAICO CSVの **「型番」列** を突き合わせ
- ZAICO CSVで保管場所が **「本社在庫」** の数量を合計 → **「本社ZAICO」列**
- ZAICO CSVで保管場所が **「★3F EC★」** の数量を合計 → **「3FZAICO」列**
- 同じSKUがCSVに複数行あれば数量を合算
- ZAICO CSVに該当SKUが無ければ **0** を記載し、画面に警告一覧を表示
- 「数量」列の右隣に2列追加するだけで、既存の内容・書式は変更しません

運用ルール（保管場所名・列名など）が変わったときは、
`zaico_core.py` 冒頭の「業務ルール設定」だけ直せばOKです。

---

## 5. ローカルで動かす場合（任意）

```powershell
cd amazon-zaico-dashboard
pip install -r requirements.txt
streamlit run app.py
```

`.streamlit/secrets.toml` に `APP_PASSWORD` を書けば、ローカルでも合言葉が有効になります
（このファイルは Git 管理対象外です）。

---

## 6. 運用・セキュリティ上の注意点

- **リポジトリは必ず Private にしてください**（在庫・商品情報を扱うため）。
- アップロードしたファイル・生成したExcelはサーバーに保存されず、その場のメモリ上のみで
  処理されます。ブラウザを閉じる／再実行するとリセットされるので、必要なファイルは
  必ずダウンロードして保存してください。
- Streamlit Community Cloud の無料枠にはスリープ機能があります。しばらく使われないと
  自動休止し、次回アクセス時に再起動（数十秒）が必要になる場合があります。

---

## 7. 困ったときは

- **「SKU」「商品名」のヘッダー行が見つかりませんでした** と出る
  → 納品プランExcelの1シート目に、`SKU` と `商品名` が並ぶ見出し行があるか確認してください。
- **ZAICOのCSVに必要な列「型番」が見つかりません** と出る
  → ZAICOのダウンロード設定で「型番」「保管場所」「数量」列が含まれているCSVか確認してください。
- **特定SKUの数量がいつも0になる**
  → ZAICO CSVの「型番」と、納品プランExcelの「SKU」が完全一致しているか（表記ゆれがないか）確認してください。
