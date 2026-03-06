# 概要
Excel VBA を使って シフトCSVの取り込み、勤務予定表の整形、色付け、出力を行うツールです。

外部システム（このコードではoplusを想定）から出力したシフトデータをExcelに取り込み、月次の勤務予定表として見やすく加工して保存します。

あわせて、内閣府公開の祝日CSVを取得し、予定表作成時の祝日判定にも利用します。

# 主な機能
  - 新規予定表の作成
  - シフトCSVの読み込み
  - シフトデータの取り込み・整形
  - シフト・作業内容の色付け
  - 既存ブックまたは新規ブックへの出力
  - シフト表のクリア
  - 祝日一覧の取得

# マクロ一覧
### メイン操作
  - Create_Schedule
    - 新規予定表を作成
  - ExportsFile_Select
    - 読み込むCSVファイルパスを設定
  - Shift_Integration
    - シフトデータを取り込み、整形
  - change_Coloring
    - 予定表の色付け、凡例変換、再集計
  - SaveFile_Select
    - 保存先ファイルパスを設定
  - Book_OutPut
    - 勤務予定表を出力
  - Shift_Clear
    - シフト表および関連表示をクリア
  - Holiday_Get
    - 祝日一覧を取得

# 処理の基本フロー
  通常の運用は以下の順序を想定しています。
1. 新規予定表を作成
   - Create_Schedule
     - 既存のカレンダー領域をクリア
     - シフト領域をクリア
     - 指定した年・月に応じて日付・曜日を展開
     - 土日祝の色付けを実施

2. 読み込むCSVを指定
   - ExportsFile_Select
     - シフトCSVファイルを選択
     - 選択したパスを oplusFilePath に設定

3. シフトデータを統合
   - Shift_Integration
     - CSVを開いてシフト表データを読み込み
     - メインシートの対象位置に貼り付け
     - 罫線などの体裁を整える
     - 集計エリアをコピー展開

4. 色付けを実行
   - change_Coloring
     - シフト凡例に基づく色付け
     - 作業内容凡例に基づく色付け
     - シフトコードの変換
     - 再集計・体裁調整

5. 保存先を指定
   - SaveFile_Select
     - 出力先ブックを選択

6. 出力
   - Book_OutPut
     - 指定した既存ブックへシートコピー、または新規ブックとして保存
     - ボタンなどのShapeを削除
     - 先頭8行を削除して配布用体裁に調整
     - シート名を outputDay の値（例: 2026年3月）へ変更

# ブック構成
コード上、以下のシート名が固定で利用されています。

|定数|シート名|
|---|---|
|SHEETNAMECREATE|シフト表|
|SHEETNAMESHIFTROLE|凡例_シフト|
|SHEETNAMEWORKROLE|凡例_作業内容|
|SHEETNAMEPLANROLE|凡例_行事予定|
|SHEETNAMEHOLIDAY|祝日一覧|


# 各シートの役割

### シフト表
  メインの作業シートです。
  - 月次カレンダー表示
  - 取り込んだシフトデータの貼り付け
  - 集計エリアの表示
  - 各種 named range の基点

### 凡例_シフト
  シフトコード変換・色付けルール定義用のシートです。
  - 外部CSV上のシフト表記
  - Excel上の変換後表記
  - 背景色の設定（RGB）
  - 文字色の設定（RGB）
  - 太字指定

### 凡例_作業内容
作業内容セルの色付けルール定義用のシートです。

### 祝日一覧
内閣府の祝日CSVを取り込むシートです。

# 必要な named range
コード内では多数の named range を参照しています。少なくとも以下が必要です。

|名前|用途|
|---|---|
|startPosition|カレンダー開始位置|
|targetPaste|シフトCSV貼り付け開始位置|
|targetAggregation|集計表開始位置|
|oplusFilePath|読込対象CSVパス格納セル|
|saveFilePath|出力先ファイルパス格納セル|
|targetYear|作成対象年|
|targetMonth|作成対象月|
|outputDay|出力シート名に使う文字列|
|createPosition|出力時のシート挿入位置指定|
|oplusRole|シフト凡例の変換元列|
|shiftRole|シフト凡例の変換先列|
|shiftInteriorRed|シフト背景色R|
|shiftInteriorGreen|シフト背景色G|
|shiftInteriorBlue|シフト背景色B|
|shiftFontRed|シフト文字色R|
|shiftFontGreen|シフト文字色G|
|shiftFontBlue|シフト文字色B|
|shiftFontStyle|シフト文字スタイル|
|workRole|作業内容凡例の検索列|
|workInteriorRed|作業内容背景色R|
|workInteriorGreen|作業内容背景色G|
|workInteriorBlue|作業内容背景色B|
|workFontRed|作業内容文字色R|
|workFontGreen|作業内容文字色G|
|workFontBlue|作業内容文字色B|
|workFontStyle|作業内容文字スタイル|

# 入力データ仕様
### シフトCSV
  - 読み込むファイルはCSV
  - シート名はファイル名から拡張子を除いた名前を想定
  - 1行目3列目に対象月が入っている前提
  - 2行目以降をデータとして取り込み

### 祝日データ
以下のURLからCSVをダウンロードします。
  
内閣府 祝日一覧　[URLはこちら](https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv)

# 出力仕様
### 既存ブックへ出力
saveFilePath が設定されている場合、既存Excelブックを開いてシートコピーします。

createPosition の値で挿入位置が変わります。

|値|動作|
|---|---|
|1|先頭に挿入|
|2|末尾に挿入|
|0|未指定扱いで中断|

同名シートが存在する場合は、上書きするか確認メッセージが表示されます。

### 新規ブックへ出力
saveFilePath が空欄の場合は、新規ブックとして保存します。

保存時には以下を実施します。
  - Shape削除
  - 先頭8行削除
  - シート名変更
  - 指定先へ .xlsx 形式で保存

# 処理詳細
### 新規予定表作成
  - createNewSchedule
    - 画面更新停止
    - 各種シート情報の初期化
    - カレンダークリア
    - シフト領域クリア
    - 指定年月の日付をセット
    - 土日祝の色設定
    - 画面更新復帰

### シフト統合
  - Main
    - CSVパス確認
    - シフト領域初期化
    - CSV読込
    - メインシートへ貼付
    - 罫線調整
    - 集計エリア展開

### 色付け
  - changeColoring
   - シフト凡例によるセル色設定
   - 作業内容凡例によるセル色設定
   - シフト表記置換
   - 集計再作成
   - 体裁調整

### 祝日取得
  - HolidayGet
    - CSVダウンロード
    - 一時保存
    - 内容を 祝日一覧 シートへ貼付
    - 一時ファイル削除

# 動作環境
  - Windows版 Excel
  - VBA有効ブック（.xlsm）
  - FileSystemObject を利用
  - urlmon.dll を利用したファイルダウンロード

# 想定依存要素
  - Microsoft Scripting Runtime を参照設定している可能性あり（FileSystemObject 型を使用しているため）
  - 一部環境では late binding のみで動作しない箇所があるため、参照設定確認推奨

# 利用手順
  1.	マクロ有効ブックを開く
  2.	targetYear と targetMonth を設定する
  3.	Holiday_Get を実行して祝日一覧を更新する
  4.	Create_Schedule を実行する
  5.	ExportsFile_Select でシフトCSVを指定する
  6.	Shift_Integration を実行する
  7.	change_Coloring を実行する
  8.	必要に応じて SaveFile_Select で出力先を指定する
  9.	Book_OutPut を実行する
