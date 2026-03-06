#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
指定フォルダ内のファイルを、共通名＋4桁連番でリネームするスクリプト
例: 共通名_0001.jpg, 共通名_0002.png ...
拡張子は維持されます。
"""

import os
import sys
import argparse


def rename_files(
    target_dir: str,
    base_name: str,
    dry_run: bool = False,
) -> None:
    """
    対象フォルダ内のファイルを「共通名＋4桁数字」でリネームする。

    Args:
        target_dir: 対象フォルダのパス
        base_name: 全ファイル共通の名前（末尾に _0001, _0002 ... が付く）
        dry_run: True の場合は実際にリネームせず、変更内容のみ表示
    """
    if not os.path.isdir(target_dir):
        print(f"エラー: フォルダが見つかりません: {target_dir}")
        sys.exit(1)

    # ファイル一覧を取得（フォルダは除外）、名前でソート
    entries = [
        e for e in os.listdir(target_dir)
        if os.path.isfile(os.path.join(target_dir, e))
    ]
    entries.sort()

    if not entries:
        print("リネーム対象のファイルがありません。")
        return

    # 既存の連番ファイル名と衝突しないよう、まず一時名にリネームしてから本リネーム
    temp_suffix = ".tmp_renumber"
    renames_phase1 = []  # (old, temp)
    renames_phase2 = []  # (temp, new)

    for i, filename in enumerate(entries, start=1):
        _, ext = os.path.splitext(filename)
        num_str = f"{i:04d}"  # 4桁（0001, 0002, ...）
        new_name = f"{base_name}{num_str}{ext}"
        temp_name = f"{base_name}{num_str}{temp_suffix}{ext}"
        old_path = os.path.join(target_dir, filename)
        temp_path = os.path.join(target_dir, temp_name)
        new_path = os.path.join(target_dir, new_name)
        renames_phase1.append((old_path, temp_path))
        renames_phase2.append((temp_path, new_path))

    if dry_run:
        print("[ドライラン] 以下のようにリネームされます:\n")
        for i, filename in enumerate(entries, start=1):
            _, ext = os.path.splitext(filename)
            num_str = f"{i:04d}"
            print(f"  {filename}  →  {base_name}{num_str}{ext}")
        return

    # Phase 1: 一時名にリネーム
    for old_path, temp_path in renames_phase1:
        try:
            os.rename(old_path, temp_path)
        except OSError as e:
            print(f"エラー: {old_path} のリネームに失敗しました: {e}")
            sys.exit(1)

    # Phase 2: 正式な連番名にリネーム
    for temp_path, new_path in renames_phase2:
        try:
            os.rename(temp_path, new_path)
        except OSError as e:
            print(f"エラー: {temp_path} のリネームに失敗しました: {e}")
            sys.exit(1)

    print(f"完了: {len(entries)} 件のファイルをリネームしました。")


def main():
    parser = argparse.ArgumentParser(
        description="指定フォルダ内のファイルを「共通名＋4桁連番」でリネームします。"
    )
    parser.add_argument(
        "target_dir",
        nargs="?",
        default=".",
        help="リネーム対象のフォルダパス（省略時はカレントディレクトリ）",
    )
    parser.add_argument(
        "-b", "--base-name",
        required=True,
        help="全ファイル共通の名前（例: photo_ で photo_0001.jpg, photo_0002.png など）",
    )
    parser.add_argument(
        "-n", "--dry-run",
        action="store_true",
        help="実際にはリネームせず、変更予定のみ表示",
    )
    args = parser.parse_args()

    target_dir = os.path.abspath(args.target_dir)
    rename_files(target_dir, base_name=args.base_name, dry_run=args.dry_run)


if __name__ == "__main__":
    main()

