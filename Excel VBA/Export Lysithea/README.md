# 概要
Excel VBA で動作する、シフト表作成・現場シフト取込・リシテア取込用データ作成を行うツールです。

主に以下の業務を自動化します。

  - 新規予定表の作成
  - 現場シフト表の取込とシフト表への反映
  - リシテア取込用シートの作成
  - リシテア取込Excelへの転記
  - 条件付き書式の再設定
  - 祝日データの取得

このブックは、複数のワークシートを前提に構成しています。

シフト表の作成から、外部シフト表の反映、最終的なリシテア連携用データの出力までを一連で処理します。

コードは以下の役割ごとにモジュール分割されています。

  - MainModule
    - ユーザーが実行する入口マクロ
  - modSetProcess
    - シートや行列位置などの初期設定
  - modMainProcess
    - メイン処理
  - modSubProcess
    - 実データ操作、転記、判定、書式設定などのサブ処理
  - modCommonProcess
    - 共通処理（画面更新制御、ファイル選択、URLデコードなど）
  - modGetProcess
    - 外部データ取得（祝日 CSV のダウンロード）
  - modConst
    - シート名・勤務区分・条件付き書式式などの定数定義
  - modVariable
    - 共通で使う構造体 wbInfo の定義

# 主な機能
### 1. 新規予定表の作成
  - 実行マクロ
    - createShift
  - 実施内容
    - 既存のカレンダー、シフト、出力領域をクリア
    - 対象月のカレンダーを生成
    - 土日・社休日に色付け
    - 適用単位が「1年」の場合はデフォルトシフトを自動生成
    - 所定労働時間上限を設定
    - 条件付き書式を再設定

### 2. 現場シフト表の取込
  - 実行マクロ
    - inputSiteShift
  - 実施内容
    - 指定 URL から対象ブックを開く
    - 対象月シートを取得
    - 氏名・日付を照合して現場シフトをシフト表へ反映
    - 記号や勤務区分を既定のシフト表記へ変換
  - 主な変換例
    - ★ / ● / ▲ / ■ / 夜勤 → 夜8
    - 休 / 休希望 → 休
    - 代休 → 代
    - / → 明
    - 年休 → 年休
    - 振休 → 日振 または 土祝振

### 3. リシテア取込Excel転記
  - 実行マクロ
    - exportLysitheaExcel
  - 実施内容
    - シフト表の対象者情報を「リシテア取込」シートへ展開
    - 日ごとの勤務区分・休日区分を判定
    - リシテア用コード体系に変換
    - 出力先 Excel ファイルを選択
    - 指定フォーマットのシートへ転記

### 4. 条件付き書式設定
  - 実行マクロ
    - setFormatCondition
  - 実施内容
    - 既存の条件付き書式を削除
    - 労働日数、休日日数、労働時間、夜勤、通し、休日の条件付き書式を再設定

### 5. 祝日・社休日取得
  - 実行マクロ
    - getHoliday
  - 実施内容
    - 内閣府の祝日 CSV をダウンロード
    - 社休日一覧 シートへ祝日一覧を貼り付け
    - シート内の社休日情報も追記して利用可能にする
  - 取得元
    - 内閣府祝日　[URLはこちら](https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv)

# 実行マクロ一覧
|マクロ名|説明|
|---|---|
|createShift|新規予定表を作成|
|inputSiteShift|現場シフト表を取り込み、シフト表へ反映|
|exportLysitheaExcel|リシテア取込Excelへ転記|
|setFormatCondition|条件付き書式を再設定|
|getHoliday|祝日・社休日一覧を更新|

# 前提となるシート構成
### VBAが動作するブック
  - リシテア取込
  - シフト表
  - マスタ
  - 条件
  - 労働時間チェック
  - 社休日一覧

## シフト取込ブック
  - 勤務予定
  - 現場シフトブック内の対象月シート（このコードでは「YYYY.M」形式で作成）

# 前提となる Named Range
  コード上、以下の named range を利用しています。

### 共通設定
  - targetFilePath
  - startDay
  - typeSystem
  - maxTime
  - maxDay
  - holidayList

### リシテア取込シート関連
  - exportStartShift
  - exportStartCalendar
  - exportEndCalendar
  - exportName
  - exportParsonalCode

### シフト表関連
  - shiftName
  - shiftStartCalendar
  - shiftEndCalendar
  - shiftParsonalCode
  - shiftWorkTime
  - shiftWorkDay

### 条件シート関連
  - conditionStartShift
  - conditionEndShift
  - conditionWorkTimePerday

### 労働時間チェック関連
  - checkStartCalendar
  - checkEndCalendar
  - checkWorkDay

### マスタ関連
  - masterShiftContent

実ブックでは、これらの named range が正しく設定されている必要があります。

# 処理フロー
### 新規予定表作成
  1.	画面更新停止
  2.	シートや変数情報を初期化
  3.	作成確認メッセージ表示
  4.	既存データ削除
  5.	適用単位判定（1か月 / 1年）
  6.	法定労働時間総枠設定
  7.	カレンダー生成
  8.	必要に応じて年間デフォルトシフト作成
  9.	条件付き書式設定
  10.	画面更新再開

### 現場シフト取込
  1.	画面更新停止
  2.	基本情報初期化
  3.	URL デコード
  4.	対象ブックを開く
  5.	対象月シートを取得
  6.	氏名・日付単位でシフトを照合
  7.	勤務記号を内部シフトへ変換
  8.	画面更新再開
  9.	完了メッセージ表示

### リシテア出力
  1.	適用単位判定
  2.	対象者一覧を出力シートへ展開
  3.	シフト表を走査し勤務区分を辞書へ格納
  4.	リシテア取込シートを生成
  5.	出力先 Excel を選択
  6.	指定レイアウトへデータ転記

# シフト変換ルールの概要
### 現場シフト → シフト表
  - 勤務系記号は主に 夜8 または 昼8 に変換
  - 休暇系記号は 休 / 代 / 明 / 年休 / 日振 / 土祝振 に変換
  - 灰色ハッチングは休日扱いとして判定する

### シフト表 → リシテア
  シフト表上の勤務時間・昼夜区分・適用単位に応じて、以下のようなコードへ変換します。
  
    例：
      ・月変8定
      ・月変8非
      ・年変8定
      ・年変8非
      ・変形休
      ・土祝振
      ・日振
      ・代休
      ・年休
      ・明

# 設計上の特徴
### 1. wbInfo による状態集約
  シート参照、行列位置、対象日、フラグ、辞書などを wbInfo に集約しています。

  各モジュールで共通の状態を利用できる構成です。

### 2. Dictionary を使った変換管理
  以下の辞書を利用しています。
  - classWorkDay
    - 日付ごとの休日区分
  - classWorkType
    - 日付ごとのデフォルト勤務区分
  - classWorkTranslate
    - 氏名 + 日付 単位のリシテア変換結果

### 3. 処理前後で Excel の描画・再計算を制御
  大量処理時の速度改善のため、処理前後で以下を制御しています。
  - ScreenUpdating
  - EnableEvents
  - DisplayAlerts
  - Calculation
  - Cursor

# 想定する利用手順
### 1.	社休日一覧 を更新する
  - getHoliday を実行
  - 
### 2.	新しい対象月を設定する
  - startDay を設定
  - typeSystem を設定（1か月 または 1年）

### 3.	新規予定表を作成する
  - createShift を実行

### 4.	必要に応じて現場シフトを反映する
  - targetFilePath に対象ファイル URL を設定
  - inputSiteShift を実行

### 5.	リシテア取込用データを作成・出力する
  - exportLysitheaExcel を実行
  - 出力先 Excel を選択

# 開発メモ
### 主な依存要素
  - Excel VBA
  - Scripting.Dictionary
  - ScriptControl
  - Windows API (URLDownloadToFile)
  - FileSystemObject

### 推奨環境
  - Windows 版 Microsoft Excel
  - マクロ有効ブック形式（.xlsm）
