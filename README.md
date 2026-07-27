# ローカルガイドブック 販売サイト

PDFガイドブック（ローカルガイドブック）を販売するための静的サイトです。
ビルド不要・HTML/CSS/JSのみなので、GitHub Pages などにそのまま公開できます。

## 公開前に必ずやること

すべて `js/config.js` を編集するだけで反映されます。HTML/CSSは基本的に触る必要はありません。

1. **商品情報を入力する**
   - `guidebookTitle` / `guidebookSubtitle` / `guidebookDescription`
   - `priceYen`（価格）
   - `contents`（収録内容のリスト）
2. **Stripe Payment Link を発行して設定する**
   - Stripeダッシュボード → 「Payment Links」から、PDFガイドブック用の決済リンクを作成
   - 発行されたURLを `stripePaymentLink` に貼り付け
   - デジタル商品の受け渡しは、Stripeの「決済完了後の案内」設定でダウンロードリンクを表示するか、
     決済完了メールでPDFのダウンロードURLを送る運用にしてください
3. **`comingSoon` を `false` に変更する**
   - PDFが完成し、Stripeリンクの設定も終わったら `false` にすると「購入する」ボタンが有効になります
4. **Instagram / 連絡先を設定する**
   - `instagramHandle` / `instagramUrl` / `contactEmail`
5. **特定商取引法に基づく表記（`legal.html`）を必ず埋める**
   - `config.legal` 内の項目（事業者名・運営責任者・所在地・電話番号など）を実情報に書き換えてください
   - 個人で販売する場合の所在地・電話番号の扱いについては、消費者庁のガイドラインを確認の上、
     「請求があれば開示する」運用にするかどうかを判断してください
   - **この表記なしでの公開・販売は法令違反になり得るため、必ず対応してください**

## ファイル構成

```
index.html      販売ページ本体
legal.html      特定商取引法に基づく表記
css/style.css   スタイル
js/config.js    サイトの全文言・価格・リンクの設定（ここだけ編集すればOK）
js/main.js      販売ページの描画ロジック
js/legal.js     特定商取引法ページの描画ロジック
```

## ローカルで確認する

```bash
python3 -m http.server 8000
# http://localhost:8000 にアクセス
```

## 公開方法（例：GitHub Pages）

1. このブランチ、またはpublic用ブランチをGitHub Pagesの公開元に設定
2. 公開後のURLをInstagramのプロフィールリンクに設定
