# ファイル連番リネームツール

指定フォルダ内のファイルを **共通名＋4桁連番** でリネームします。拡張子はそのまま維持されます。

## 必要な環境

- Python 3.6 以上

## 使い方

```bash
# 共通名を指定して実行（-b または --base-name は必須）
python rename_files.py "C:\path\to\target\folder" -b "DSC_"

# カレントディレクトリ内を "DSC_" でリネーム
python rename_files.py -b "DSC_"

# ドライラン（変更せずにどう変わるかだけ表示）
python rename_files.py "C:\path\to\folder" -b "DSC_" --dry-run
```

## 例

共通名を `photo_` にした場合:

リネーム前:
- `photo_a.jpg`
- `photo_b.jpg`
- `document.pdf`

リネーム後:
- `photo_0001.jpg`
- `photo_0002.jpg`
- `photo_0003.pdf`

番号は **4桁**（0001, 0002, ... 9999）で付きます。ファイルは **名前の辞書順** で並べた順に割り当てられます。実行前に `--dry-run` で確認することを推奨します。
