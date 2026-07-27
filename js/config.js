/**
 * サイト設定ファイル
 * ここの値を書き換えるだけで、サイト全体の文言・価格・リンクが更新されます。
 * HTMLやCSSを直接触る必要はありません。
 */
window.SITE_CONFIG = {
  // ---- 販売者情報 ----
  sellerName: "Mr.sasuken",
  instagramHandle: "@sasuken", // EDIT ME: 実際のInstagramユーザー名に変更
  instagramUrl: "https://www.instagram.com/sasuken/", // EDIT ME: 実際のプロフィールURLに変更
  contactEmail: "your-email@example.com", // EDIT ME: 問い合わせ用メールアドレス

  // ---- 特定商取引法に基づく表記（legal.html用） ----
  // デジタルコンテンツの個人販売でも記載が必要です。必ず実情報に書き換えてください。
  legal: {
    name: "（例）〇〇〇〇（販売事業者名または個人名）", // EDIT ME
    representative: "（例）〇〇 〇〇", // EDIT ME
    address: "請求があった場合に遅滞なく開示します（個人情報保護のため、通常時は非表示にできます）", // EDIT ME: 開示方針に応じて修正
    phone: "請求があった場合に遅滞なく開示します", // EDIT ME
    extraFee: "インターネット接続料金等はお客様のご負担となります。", // EDIT ME
    paymentMethod: "クレジットカード決済（Stripe）",
    paymentTiming: "ご注文時に即時決済されます。",
    delivery: "決済完了後、速やかにダウンロード用リンクをご案内します。",
    refund: "デジタルコンテンツの性質上、購入後の返品・返金はお受けしておりません。",
  },

  // ---- 商品情報 ----
  guidebookTitle: "ローカルガイドブック", // EDIT ME: 正式な商品タイトルに変更（例：「大阪 裏路地さんぽガイド」）
  guidebookSubtitle: "地元民だけが知る、とっておきの街歩きガイド", // EDIT ME
  guidebookDescription:
    "観光サイトには載っていないローカルなお店・スポットだけを厳選して1冊にまとめたPDFガイドブックです。実際に自分の足で歩いて確かめた情報だけを掲載しています。",
  priceYen: 1980, // EDIT ME: 販売価格（税込・円）
  format: "PDF（スマホ・タブレット・PCで閲覧可能）",
  pageCountNote: "全〇〇ページ予定", // EDIT ME: ページ数が決まったら記載

  // ---- 商品ステータス ----
  // 準備中は true にしておくと「購入」ボタンが無効化され「近日公開」表示になります。
  // 販売開始の準備ができたら false に変更してください。
  comingSoon: true,

  // ---- 決済 ----
  // Stripe Payment Links で発行したURLに差し替えてください。
  // https://dashboard.stripe.com/payment-links から発行できます。
  stripePaymentLink: "https://buy.stripe.com/xxxxxxxxxxxxxxxx", // EDIT ME

  // ---- 目次・収録内容（自由に増減可） ----
  contents: [
    "地元民おすすめの飲食店リスト",
    "観光ガイドに載らない穴場スポット",
    "モデルコース（半日〜1日）",
    "アクセス・営業時間などの実用情報",
  ],

  // ---- よくある質問（自由に増減可） ----
  faq: [
    {
      q: "購入後、どうやってPDFを受け取れますか？",
      a: "Stripeでの決済完了後、決済完了ページおよびご登録のメールアドレス宛にダウンロード用のリンクをお送りします。",
    },
    {
      q: "スマホでも読めますか？",
      a: "はい。PDF形式のため、スマートフォン・タブレット・PCなど、PDFが閲覧できる端末であればご覧いただけます。",
    },
    {
      q: "返金は可能ですか？",
      a: "デジタルコンテンツの性質上、購入後の返金はお受けしておりません。あらかじめご了承の上ご購入ください。",
    },
  ],
};
