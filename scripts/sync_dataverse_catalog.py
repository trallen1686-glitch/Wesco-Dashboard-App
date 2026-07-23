"""Synchronize the Azure content catalog with Wesco Documents / Photos.

By default this script performs a read-only preview. Pass --apply to create
or update Dataverse records after the environment has been confirmed.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from auth import get_client  # noqa: E402


TABLE = "wes_documentsphotos"
ID_COLUMN = "wes_documentsphotosid"
NAME_COLUMN = "wes_documentsphotos_documentname"
URL_COLUMN = "wes_documentsphotos_filelink"
NOTES_COLUMN = "wes_documentsphotos_notes"
CATALOG_PATH = ROOT / "wesco-content-catalog.json"
MARKER = "[Wesco Hub Catalog]"


def notes_for(item: dict[str, object]) -> str:
    details = {
        "source": MARKER,
        "type": "file",
        "repositoryPath": item["path"],
        "githubUrl": item["githubUrl"],
        "category": item["category"],
        "extension": item["extension"],
        "sha256": item["sha256"],
        "sizeBytes": item["sizeBytes"],
    }
    return json.dumps(details, separators=(",", ":"), ensure_ascii=True)


def desired_records(catalog: dict[str, object]) -> list[dict[str, str]]:
    records = [
        {
            NAME_COLUMN: str(item["name"]),
            URL_COLUMN: str(item["azureUrl"]),
            NOTES_COLUMN: notes_for(item),
        }
        for item in catalog["files"]
    ]
    for folder in catalog["folders"]:
        folder_path = str(folder).strip("/")
        records.append(
            {
                NAME_COLUMN: f"Folder: {folder_path}",
                URL_COLUMN: f"{catalog['azureBaseUrl']}/{folder_path}/",
                NOTES_COLUMN: json.dumps(
                    {
                        "source": MARKER,
                        "type": "folder",
                        "repositoryPath": f"{folder_path}/",
                    },
                    separators=(",", ":"),
                    ensure_ascii=True,
                ),
            }
        )
    return records


def all_rows(client) -> list[dict[str, object]]:
    records = client.records.list(
        TABLE,
        select=[ID_COLUMN, NAME_COLUMN, URL_COLUMN, NOTES_COLUMN],
        top=5000,
    )
    return list(records)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Create and update Dataverse records. Default is preview only.",
    )
    args = parser.parse_args()

    catalog = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    desired = desired_records(catalog)
    client = get_client("dv-data")
    existing = all_rows(client)
    catalog_existing = [
        row
        for row in existing
        if MARKER in str(row.get(NOTES_COLUMN, ""))
    ]
    by_url = {
        str(row.get(URL_COLUMN, "")).rstrip("/").lower(): row
        for row in catalog_existing
        if row.get(URL_COLUMN)
    }

    creates: list[dict[str, str]] = []
    updates: list[tuple[str, dict[str, str]]] = []
    unchanged = 0
    for record in desired:
        key = record[URL_COLUMN].rstrip("/").lower()
        current = by_url.get(key)
        if not current:
            creates.append(record)
            continue
        changed = {
            column: value
            for column, value in record.items()
            if str(current.get(column, "")) != value
        }
        if changed:
            updates.append((str(current[ID_COLUMN]), changed))
        else:
            unchanged += 1

    print(
        f"Environment: {os.environ.get('DATAVERSE_URL', 'from .env')}\n"
        f"Table: {TABLE}\n"
        f"Desired catalog records: {len(desired)}\n"
        f"Create: {len(creates)}\n"
        f"Update: {len(updates)}\n"
        f"Unchanged: {unchanged}"
    )
    if not args.apply:
        print("Preview only. Re-run with --apply to write these changes.")
        return

    if creates:
        client.records.create(TABLE, creates)
    for record_id, changes in updates:
        client.records.update(TABLE, record_id, changes)
    print(
        f"Applied successfully: {len(creates)} created, "
        f"{len(updates)} updated."
    )


if __name__ == "__main__":
    main()
