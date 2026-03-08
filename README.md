# Portfolio

個人で作成したサンプル・ツール群のリポジトリです。  
**言語名 / 動作名** の階層で整理し、各プロジェクトは独自の README で詳細を記載しています。

## リポジトリ構成
```
Portfolio/
├── Python/           # Python 製ツール
├── Ruby/             # Ruby（Sinatra）アプリ
├── HTML/             # 静的サイト・LP
├── Excel VBA/        # Excel VBA マクロ
├── scripts/          # CI・同期用スクリプト
├── .github/workflows/# GitHub Actions（Notion 同期）
└── README.md         # 本ファイル
```

## プロジェクト一覧

### Python

| 動作名 | 説明 | README |
|--------|------|--------|
| **Rename File** | 指定フォルダ内のファイルを「共通名＋4桁連番」で一括リネーム。拡張子は維持。`--dry-run` で事前確認可能。 | [python/Rename File/README.md](python/Rename%20File/README.md) |

### Ruby

| 動作名 | 説明 | README |
|--------|------|--------|
| **Form** | Sinatra 製の簡易サーバー。お問い合わせフォームの送信を受け取り、CSV に保存。HTML は `HTML/` フォルダを参照。 | [Ruby/Form/README.md](Ruby/Form/README.md) |

### HTML

| 動作名 | 説明 | README / デモ |
|--------|------|----------------|
| **Sample_LP** | レスポンシブ対応のダークテーマ LP。Vanilla HTML/CSS、特徴ページ・お問い合わせ・登録ページあり。Sinatra 連携用 ERB テンプレート含む。 | [HTML/Sample_LP/README.md](HTML/Sample_LP/README.md) · [サンプルページ](https://zesty-alpaca-59851b.netlify.app/) |

### Excel VBA

| 動作名 | 説明 | README |
|--------|------|--------|
| **Integration Shift** | シフト CSV の取り込み・勤務予定表の整形・色付け・出力。祝日 CSV の取得にも対応。 | [Excel VBA/Integration Shift/README.md](Excel%20VBA/Integration%20Shift/README.md) |
| **Export Lysithea** | シフト表の作成、現場シフト取込、リシテア取込用 Excel への転記。条件付き書式・祝日取得を含む一連の業務自動化。 | [Excel VBA/Export Lysithea/README.md](Excel%20VBA/Export%20Lysithea/README.md) |

## Notion 同期（GitHub Actions）
`main` ブランチへの push 時、**変更のあった「言語名/動作名」のみ** Notion のデータベースに同期します。

- **ワークフロー:** [.github/workflows/notion-sync.yml](.github/workflows/notion-sync.yml)
- **スクリプト:** `scripts/notion_sync.py`
- **対象外:** `.github/`、`scripts/`、`README.md` のみの変更では同期しない

Notion 側では `NOTION_TOKEN` と `NOTION_DATABASE_ID` のシークレット設定が必要です。

## 利用・開発について

各プロジェクトの実行方法・前提環境は、上記リンク先の README を参照してください。
