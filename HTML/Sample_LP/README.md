# サンプルランディングページ（HTML）

**レスポンシブ対応のダークテーマLP**です。
HTML/CSS のみで構成しております。

---

## デモ・公開URL


---

## 技術スタック

| 種別 | 技術 |
|------|------|
| マークアップ | HTML5（セマンティックな構造） |
| スタイル | CSS3（カスタムプロパティ・Flexbox・Grid・メディアクエリ） |
| フォント | Google Fonts（Noto Sans JP, Outfit） |

フレームワークなしの **Vanilla HTML/CSS** で実装しているため、軽量でカスタマイズしやすい構成です。

---

## 主な機能・構成

- **トップ（index.html）**
  - ヒーローセクション（キャッチコピー・CTA）
  - 特徴カード（スピード・シンプル・セキュリティへのリンク）
  - レスポンシブナビゲーション
- **特徴ページ**（feature-speed.html / feature-simple.html / feature-security.html）
  - 各テーマの詳細説明ページ
- **お問い合わせ**（contact.html）
  - 静的版の問い合わせページ
- **登録**（register.html）
  - 登録用のフォームページ
- **Sinatra 連携**（views/index.erb）
  - フォーム送信をサーバー側で受け取り、CSV に保存するテンプレート

デザインは **ダークテーマ** をベースに、アクセントカラー（インディゴ/パープル）とグラデーションで統一しています。

---

## ディレクトリ構成

```
HTML/Sample_LP
├── index.html          # トップページ
├── contact.html        # お問い合わせ（静的版）
├── register.html       # 登録ページ
├── feature-speed.html  # 特徴：スピード
├── feature-simple.html # 特徴：シンプル
├── feature-security.html # 特徴：セキュリティ
├── views/
│   └── index.erb       # 問い合わせフォーム付きLP用テンプレート（Sinatra 参照）
├── README.md           # 本ファイル
└── DEPLOY.md           # デプロイ手順（ポートフォリオ公開用）
```
