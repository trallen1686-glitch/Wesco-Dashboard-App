"""Build the public Wesco Dashboard content catalog.

The generated JSON is deployed by Azure Static Web Apps and is also the
source used to synchronize the Dataverse template catalog.
"""

from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "wesco-content-catalog.json"
AZURE_BASE = "https://www.wescodashboardapp.com"
GITHUB_BASE = (
    "https://github.com/trallen1686-glitch/Wesco-Dashboard-App/blob/main"
)

EXCLUDED_PARTS = {
    ".git",
    ".github",
    ".vscode",
    ".claude",
    "plugins",
    "powerapps-review",
    "scripts",
    "solutions",
}
EXCLUDED_FILES = {
    ".env",
    ".gitignore",
    "wesco-content-catalog.json",
}
PUBLIC_SUFFIXES = {".html", ".json", ".md", ".png", ".xlsx"}


def is_public(path: Path) -> bool:
    relative = path.relative_to(ROOT)
    return (
        path.is_file()
        and path.name not in EXCLUDED_FILES
        and path.suffix.lower() in PUBLIC_SUFFIXES
        and not any(part in EXCLUDED_PARTS for part in relative.parts)
    )


def html_title(path: Path) -> str | None:
    if path.suffix.lower() != ".html":
        return None
    text = path.read_text(encoding="utf-8", errors="ignore")
    match = re.search(r"<title[^>]*>(.*?)</title>", text, re.I | re.S)
    if not match:
        return None
    return re.sub(r"\s+", " ", match.group(1)).strip()


def display_name(path: Path) -> str:
    title = html_title(path)
    if title:
        return title
    stem = path.stem.replace("_", " ").replace("-", " ")
    return re.sub(r"\s+", " ", stem).strip()


def category(path: Path) -> str:
    rel = path.relative_to(ROOT).as_posix()
    lower = rel.lower()
    if rel == "index.html":
        return "Hub"
    if lower.endswith((".png", ".xlsx")):
        return "Resource"
    if lower.endswith((".json", ".md")):
        return "Configuration"
    if "project" in lower:
        return "Projects"
    if "staff" in lower or "member" in lower or "team" in lower:
        return "Staff"
    if "subcontract" in lower:
        return "Subcontractors"
    if any(word in lower for word in ("estimate", "bid ", "proposal", "scope")):
        return "Estimating"
    if any(word in lower for word in ("purchase", "invoice", "contract", "po ")):
        return "Financials"
    return "Operations"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> None:
    paths = sorted(
        (path for path in ROOT.rglob("*") if is_public(path)),
        key=lambda path: path.relative_to(ROOT).as_posix().lower(),
    )
    files = []
    for path in paths:
        relative = path.relative_to(ROOT).as_posix()
        encoded = quote(relative, safe="/")
        files.append(
            {
                "name": display_name(path),
                "path": relative,
                "folder": str(Path(relative).parent).replace("\\", "/"),
                "category": category(path),
                "extension": path.suffix.lower().lstrip("."),
                "sizeBytes": path.stat().st_size,
                "sha256": sha256(path),
                "azureUrl": (
                    f"{AZURE_BASE}/" if relative == "index.html"
                    else f"{AZURE_BASE}/{encoded}"
                ),
                "githubUrl": f"{GITHUB_BASE}/{encoded}",
                "active": True,
            }
        )

    folders = sorted(
        {
            item["folder"]
            for item in files
            if item["folder"] not in {"", "."}
        }
    )
    payload = {
        "schemaVersion": 1,
        "generatedAtUtc": datetime.now(timezone.utc).isoformat(),
        "azureBaseUrl": AZURE_BASE,
        "githubRepository": (
            "https://github.com/trallen1686-glitch/Wesco-Dashboard-App"
        ),
        "folders": folders,
        "itemCount": len(files),
        "files": files,
    }
    OUTPUT.write_text(
        json.dumps(payload, indent=2, ensure_ascii=True) + "\n",
        encoding="utf-8",
    )
    print(f"Wrote {len(files)} public items to {OUTPUT.name}")


if __name__ == "__main__":
    main()
