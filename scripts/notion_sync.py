import os
import datetime
import requests


def require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required env var: {name}")
    return value


def main() -> None:
    # Required envs
    token = require_env("NOTION_TOKEN")
    database_id = require_env("NOTION_DATABASE_ID")

    # GitHub context (provided by Actions, we also pass explicitly from workflow)
    server_url = os.getenv("GITHUB_SERVER_URL", "https://github.com")
    full_repo = os.getenv("GITHUB_REPOSITORY", "")  # e.g. owner/repo
    if not full_repo:
        raise RuntimeError("Missing required env var: GITHUB_REPOSITORY")

    repo_name = full_repo.split("/")[-1]
    github_url = f"{server_url}/{full_repo}"

    # README head (first 100 chars)
    content = ""
    if os.path.exists("README.md"):
        with open("README.md", "r", encoding="utf-8") as f:
            content = f.read(100)

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28",
    }

    # Search existing page by Title property "Asset Name"
    search_data = {
        "filter": {
            "property": "Asset Name",
            "title": {"equals": repo_name},
        }
    }

    res = requests.post(
        f"https://api.notion.com/v1/databases/{database_id}/query",
        headers=headers,
        json=search_data,
        timeout=30,
    )
    res.raise_for_status()
    pages = res.json().get("results", [])

    # Update/Create payload
    payload = {
        "properties": {
            "Asset Name": {"title": [{"text": {"content": repo_name}}]},
            "Last Updated": {"date": {"start": datetime.datetime.now().isoformat()}},
            "Summary": {"rich_text": [{"text": {"content": content}}]},
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
        print(f"Updated Notion page for {repo_name}")
    else:
        create_payload = {"parent": {"database_id": database_id}, **payload}
        r = requests.post(
            "https://api.notion.com/v1/pages",
            headers=headers,
            json=create_payload,
            timeout=30,
        )
        r.raise_for_status()
        print(f"Created Notion page for {repo_name}")


if __name__ == "__main__":
    main()
