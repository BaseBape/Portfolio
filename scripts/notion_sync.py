"""
Notion 同期: リポジトリ構成「言語名/動作名/ソースコード」に合わせ、
更新のあった動作名のみ Notion のデータベースにページ作成・更新する。

必要な環境変数:
  NOTION_TOKEN, NOTION_DATABASE_ID, GITHUB_REPOSITORY
  CHANGED_FILES: 変更のあったファイルパス（改行区切り）。例: Python/notion_sync/notion_sync.py

Notion データベースに以下のプロパティが必要です:
  - Asset Name (タイトル): 動作名のみ（検索は GitHub Repo で行う）
  - Repository (リッチテキスト): リポジトリ名
  - Language (リッチテキスト): 言語名
  - Action Name (リッチテキスト): 動作名
  - Last Updated (日付)
  - Summary (リッチテキスト)
  - GitHub Repo (URL): その動作のフォルダへのリンク（検索に使用）
"""
import os
import datetime
import requests


def require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required env var: {name}")
    return value


def get_changed_actions() -> list[tuple[str, str]]:
    """
    環境変数 CHANGED_FILES（改行区切りパス）から
    「言語名/動作名」の一覧を抽出する。
    パスは「言語名/動作名/...」の形式を前提とする。
    """
    raw = os.getenv("CHANGED_FILES", "").strip()
    if not raw:
        return []

    actions: set[tuple[str, str]] = set()
    for path in raw.splitlines():
        path = path.strip()
        if not path or path.startswith("#"):
            continue
        parts = path.replace("\\", "/").split("/")
        if len(parts) >= 2:
            language, action_name = parts[0], parts[1]
            if language and action_name:
                actions.add((language, action_name))

    return sorted(actions)


def get_summary(language: str, action_name: str) -> str:
    """その動作フォルダ内の README.md の先頭100文字を返す。"""
    readme = os.path.join(language, action_name, "README.md")
    if not os.path.exists(readme):
        return ""
    try:
        with open(readme, "r", encoding="utf-8") as f:
            return f.read(100).strip()
    except Exception:
        return ""


def sync_action(
    *,
    token: str,
    database_id: str,
    repo_name: str,
    full_repo: str,
    server_url: str,
    ref_name: str,
    language: str,
    action_name: str,
) -> None:
    summary = get_summary(language, action_name)
    # フォルダへのリンク（ブランチ名がない場合は main を使用）
    branch = ref_name or "main"
    folder_path = f"{language}/{action_name}".replace(" ", "%20")
    github_url = f"{server_url}/{full_repo}/tree/{branch}/{folder_path}"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28",
    }

    # 既存ページ検索: GitHub Repo（フォルダURL）で検索
    search_data = {
        "filter": {"property": "GitHub Repo", "url": {"equals": github_url}}
    }

    res = requests.post(
        f"https://api.notion.com/v1/databases/{database_id}/query",
        headers=headers,
        json=search_data,
        timeout=30,
    )
    res.raise_for_status()
    pages = res.json().get("results", [])

    payload = {
        "properties": {
            "Asset Name": {"title": [{"text": {"content": action_name}}]},
            "Repository": {"rich_text": [{"text": {"content": repo_name}}]},
            "Language": {"rich_text": [{"text": {"content": language}}]},
            "Action Name": {"rich_text": [{"text": {"content": action_name}}]},
            "Last Updated": {"date": {"start": datetime.datetime.now().isoformat()}},
            "Summary": {"rich_text": [{"text": {"content": summary[:2000]}]}},
            "GitHub Repo": {"url": github_url},
        }
    }

    if pages:
        page_id = pages[0]["id"]
        r = requests.patch(
            f"https://api.notion.com/v1/pages/{page_id}",
            headers=headers,
            json=payload,
            timeout=30,
        )
        r.raise_for_status()
        print(f"Updated: {action_name}")
    else:
        create_payload = {"parent": {"database_id": database_id}, **payload}
        r = requests.post(
            "https://api.notion.com/v1/pages",
            headers=headers,
            json=create_payload,
            timeout=30,
        )
        r.raise_for_status()
        print(f"Created: {action_name}")


def main() -> None:
    token = require_env("NOTION_TOKEN")
    database_id = require_env("NOTION_DATABASE_ID")

    server_url = os.getenv("GITHUB_SERVER_URL", "https://github.com")
    full_repo = os.getenv("GITHUB_REPOSITORY", "")
    if not full_repo:
        raise RuntimeError("Missing required env var: GITHUB_REPOSITORY")

    repo_name = full_repo.split("/")[-1]
    ref_name = os.getenv("GITHUB_REF_NAME", "")  # ブランチ名

    actions = get_changed_actions()
    if not actions:
        print("No changed actions (CHANGED_FILES empty or no language/action paths). Exiting.")
        return

    print(f"Syncing {len(actions)} action(s) to Notion: {[f'{l}/{a}' for l, a in actions]}")

    for language, action_name in actions:
        sync_action(
            token=token,
            database_id=database_id,
            repo_name=repo_name,
            full_repo=full_repo,
            server_url=server_url,
            ref_name=ref_name,
            language=language,
            action_name=action_name,
        )


if __name__ == "__main__":
    main()
