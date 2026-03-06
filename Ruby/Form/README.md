# サーバー（Ruby / Sinatra）

このフォルダにはアプリケーションのRubyコードのみ含まれます。  
HTML・テンプレート・静的ファイルは同じ階層の **HTML** フォルダを参照します。

## 起動方法

```bash
cd Ruby
bundle install
bundle exec ruby app.rb
```

ブラウザで http://localhost:4567 を開いてください。

## フォルダ構成

- `app.rb` … Sinatra アプリ（ルート・お問い合わせ処理）
- `Gemfile` / `Gemfile.lock` … 依存関係
- `inquiries.csv` … 送信時に自動作成（お問い合わせ保存先）

 views と public は **../HTML/views** と **../HTML/public** を参照しています。
