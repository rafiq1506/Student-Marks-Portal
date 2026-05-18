from __future__ import annotations

import json
import os
import random
import re
import shutil
import smtplib
import ssl
import subprocess
import warnings
import time
import urllib.request
from io import BytesIO, StringIO
from email.message import EmailMessage
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse
from uuid import uuid4
import pandas as pd
from keep_alive import keep_alive
keep_alive()


warnings.filterwarnings(
    "ignore",
    message="Downcasting object dtype arrays on .fillna",
    category=FutureWarning,
)

ROOT = Path(__file__).resolve().parent

EXCEL_MIMES = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel"
]


DATA_DIR = ROOT / "data"
DEFAULT_DATA_FILE = ROOT / "data" / "student_data.xlsx"
FALLBACK_DATA_FILE = ROOT / "data" / "students.xlsx"
SOURCE_CONFIG_FILE = ROOT / "data" / "source_config.json"
PAPER_REGISTRY_FILE = ROOT / "data" / "papers.json"
PAPER_UPLOAD_DIR = ROOT / "data" / "papers"
SELECTED_PAPER_FILE = ROOT / "data" / "selected_paper.json"
OTP_TTL_SECONDS = 10 * 60
DRIVE_REGISTRY_FILENAME = os.getenv("GOOGLE_DRIVE_REGISTRY_FILENAME", "student_portal_papers.json")
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]

OTP_STORE: dict[str, dict[str, object]] = {}
DRIVE_SERVICE = None
DRIVE_PAPER_CACHE = []
DRIVE_PAPER_CACHE_TIME = 0
DRIVE_REGISTRY_CACHE: dict[str, object] | None = None
DRIVE_REGISTRY_CACHE_TIME = 0
PAPER_DATAFRAME_CACHE: dict[str, tuple[float, pd.DataFrame, str, dict[str, str]]] = {}

DRIVE_SYNC_SECONDS = int(
    os.getenv(
        "GOOGLE_DRIVE_SYNC_SECONDS",
        "300"
    )
)

DRIVE_REGISTRY_CACHE_SECONDS = int(os.getenv("GOOGLE_DRIVE_REGISTRY_CACHE_SECONDS", "60"))
PAPER_DATA_CACHE_SECONDS = int(os.getenv("PAPER_DATA_CACHE_SECONDS", "300"))

COLUMN_ALIASES = {
    "name": ["name", "student name", "student's name", "student_name"],
    "course_name": ["course name", "course", "program", "class"],
    "paper_name": ["paper name", "paper", "subject", "subject name"],
    "college_roll_number": [
        "college roll number",
        "college roll no",
        "college roll no.",
        "college_roll_number",
        "college_roll_no",
        "roll number",
        "roll no.",
        "roll number",
        "roll no",
    ],
    "exam_roll_number": [
        "exam roll number",
        "exam roll no",
        "exam roll no.",
        "exam_roll_number",
        "exam_roll_no",
    ],
    "email": ["email", "email id", "email_id", "student email", "mail"],
    "assignment_marks": ["assignment marks", "assignment_marks", "assignment", "assignment (4)"],
    "test_marks": ["test marks", "test_marks", "test", "test (4)"],
    "lectures_taken": [
        "lectures taken (theory+practical)",
        "lectures taken",
        "total lectures",
        "total lectures taken",
    ],
    "total_attendance": [
        "total lectures attended",
        "total attendance",
        "total attended",
        "total classes attended",
    ],
    "attendance_percentage": [
        "total percentage of attendance",
        "attendance percentage",
        "attendance %",
        "attendance",
        "percentage (%)",
        "percentage",
    ],
    "attendance_marks": [
        "attendance marks (2)",
        "attendance marks",
        "attendance_marks",
    ],
    "internal_marks": [
        "total marks of iaee/iees internals",
        "iaee/iees internals",
        "internal marks",
        "internals",
        "iaee internal marks",
        "iees internal marks",
        "total ia marks (10)",
        "total ia marks",
    ],
}

MONTH_NAMES = {
    "january",
    "february",
    "march",
    "april",
    "may",
    "june",
    "july",
    "august",
    "september",
    "october",
    "november",
    "december",
}


def get_drive_excel_files():

    global DRIVE_PAPER_CACHE
    global DRIVE_PAPER_CACHE_TIME

    now = time.time()

    if (
        DRIVE_PAPER_CACHE
        and
        (
            now
            -
            DRIVE_PAPER_CACHE_TIME
        )
        <
        DRIVE_SYNC_SECONDS
    ):

        return DRIVE_PAPER_CACHE

    service = google_drive_service()

    folder_id = os.getenv(
        "GOOGLE_DRIVE_FOLDER_ID"
    )

    mime_query = " or ".join(
        [
            f"mimeType='{m}'"
            for m in EXCEL_MIMES
        ]
    )

    query = (
        f"'{folder_id}' in parents "
        f"and trashed=false "
        f"and ({mime_query})"
    )

    result = (
        service.files()
        .list(
            q=query,
            fields=
            "files(id,name,createdTime)",
            supportsAllDrives=True
        )
        .execute()
    )

    files = result.get(
        "files",
        []
    )

    papers = []

    for f in files:

        try:

            excel_bytes = download_drive_file(
                f["id"]
            )

            frame = pd.read_excel(
                BytesIO(
                    excel_bytes
                ),
                dtype=object
            ).fillna("")

            paper_name = ""

            if "Paper Name" in frame.columns:

                for value in frame[
                    "Paper Name"
                ]:

                    value = clean_cell(
                        value
                    )

                    if value:
                        paper_name = value
                        break

            if not paper_name:
                paper_name = os.path.splitext(
                    f["name"]
                )[0]

        except Exception:

            paper_name = os.path.splitext(
                f["name"]
            )[0]

        papers.append(
            {
                "paper_id": f["id"],
                "paper_name": paper_name,
                "source": "google_drive",
                "file_name": f["name"]
            }
        )

    DRIVE_PAPER_CACHE = papers

    DRIVE_PAPER_CACHE_TIME = (
        time.time()
    )

    return papers

def clear_drive_cache():

    global DRIVE_PAPER_CACHE
    global DRIVE_PAPER_CACHE_TIME
    global DRIVE_REGISTRY_CACHE
    global DRIVE_REGISTRY_CACHE_TIME

    DRIVE_PAPER_CACHE = []

    DRIVE_PAPER_CACHE_TIME = 0
    DRIVE_REGISTRY_CACHE = None
    DRIVE_REGISTRY_CACHE_TIME = 0


def clear_paper_data_cache(paper_id: str = "") -> None:
    if paper_id:
        normalized = paper_id.strip().lower()
        for key in list(PAPER_DATAFRAME_CACHE):
            if key.startswith(f"{normalized}:"):
                PAPER_DATAFRAME_CACHE.pop(key, None)
        return
    PAPER_DATAFRAME_CACHE.clear()

def load_local_env() -> None:
    env_file = ROOT / ".env"
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_local_env()

try:
    import certifi

    os.environ.setdefault("SSL_CERT_FILE", certifi.where())
    os.environ.setdefault("REQUESTS_CA_BUNDLE", certifi.where())
except ImportError:
    pass


def load_env_file() -> None:
    env_path = ROOT / ".env"
    if not env_path.exists():
        return

    with env_path.open("r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            os.environ[key] = value


def smtp_is_configured() -> bool:
    return all(
        os.getenv(key)
        for key in ("SMTP_HOST", "SMTP_USER", "SMTP_PASS")
    )


def smtp_security_mode(port: int) -> str:
    configured = os.getenv("SMTP_SECURITY", "").strip().lower()
    if configured in {"ssl", "starttls", "none"}:
        return configured
    if port == 465:
        return "ssl"
    return "starttls"


def normalize_header(value: object) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("\n", " ")
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def clean_cell(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def clean_roll(value: object) -> str:
    return clean_cell(value).replace(" ", "")


def slugify(value: str) -> str:
    slug = normalize_header(value)
    slug = re.sub(r"[^a-z0-9]+", "-", slug).strip("-")
    return slug or f"paper-{int(time.time())}"


def google_drive_enabled() -> bool:
    return bool(os.getenv("GOOGLE_DRIVE_FOLDER_ID")) and bool(google_drive_auth_configured())


def google_drive_auth_configured() -> bool:
    return (
        os.getenv("GOOGLE_OAUTH_TOKEN_JSON")
        or os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
        or os.getenv("GOOGLE_DRIVE_CREDENTIALS_JSON")
        or os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    )


def google_drive_service():
    global DRIVE_SERVICE
    if DRIVE_SERVICE is not None:
        return DRIVE_SERVICE

    try:
        from google.oauth2.credentials import Credentials as OAuthCredentials
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from google_auth_httplib2 import AuthorizedHttp
        import httplib2
    except ImportError as exc:
        raise RuntimeError(
            "Google Drive storage is configured, but Google API libraries are missing. "
            "Run: pip install -r requirements.txt"
        ) from exc

    oauth_token_json = os.getenv("GOOGLE_OAUTH_TOKEN_JSON")
    credentials_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON") or os.getenv("GOOGLE_DRIVE_CREDENTIALS_JSON")
    credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if oauth_token_json:
        if oauth_token_json.strip().startswith("{"):
            payload = json.loads(oauth_token_json)
        else:
            with Path(oauth_token_json).expanduser().open("r", encoding="utf-8") as file:
                payload = json.load(file)
        if "installed" in payload or "web" in payload:
            raise RuntimeError(
                "GOOGLE_OAUTH_TOKEN_JSON contains an OAuth client JSON file, not an authorized token. "
                "Clear GOOGLE_OAUTH_TOKEN_JSON or generate a token JSON with refresh_token, client_id, and client_secret."
            )
        credentials = OAuthCredentials.from_authorized_user_info(payload, scopes=DRIVE_SCOPES)
    elif credentials_json:
        if credentials_json.strip().startswith("{"):
            payload = json.loads(credentials_json)
        else:
            with Path(credentials_json).expanduser().open("r", encoding="utf-8") as file:
                payload = json.load(file)
        credentials = service_account.Credentials.from_service_account_info(payload, scopes=DRIVE_SCOPES)
    elif credentials_path:
        credentials = service_account.Credentials.from_service_account_file(
            str(Path(credentials_path).expanduser()),
            scopes=DRIVE_SCOPES,
        )
    else:
        raise RuntimeError("Google Drive credentials are not configured.")

    disable_ssl_verify = os.getenv("GOOGLE_API_DISABLE_SSL_VERIFY", "").strip().lower() in {"1", "true", "yes"}
    if disable_ssl_verify:
        base_http = httplib2.Http(disable_ssl_certificate_validation=True)
    else:
        ca_certs = os.environ.get("SSL_CERT_FILE")
        if not ca_certs:
            try:
                import certifi

                ca_certs = certifi.where()
            except ImportError:
                ca_certs = None
        base_http = httplib2.Http(ca_certs=ca_certs)
    http = AuthorizedHttp(credentials, http=base_http)
    DRIVE_SERVICE = build("drive", "v3", http=http, cache_discovery=False)
    return DRIVE_SERVICE


def drive_folder_id() -> str:
    folder_id = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "").strip()
    if not folder_id:
        raise RuntimeError("GOOGLE_DRIVE_FOLDER_ID is required for Google Drive storage.")
    return folder_id


def drive_query_value(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def find_drive_file_by_name(name: str) -> str:
    service = google_drive_service()
    folder_id = drive_folder_id()
    query = (
        f"name = '{drive_query_value(name)}' and "
        f"'{drive_query_value(folder_id)}' in parents and trashed = false"
    )
    result = (
        service.files()
        .list(q=query, spaces="drive", fields="files(id, name)", pageSize=1, supportsAllDrives=True)
        .execute()
    )
    files = result.get("files", [])
    return str(files[0]["id"]) if files else ""


def download_drive_file(file_id: str) -> bytes:
    try:
        from googleapiclient.http import MediaIoBaseDownload
    except ImportError as exc:
        raise RuntimeError(
            "Google Drive storage is configured, but Google API libraries are missing. "
            "Run: pip install -r requirements.txt"
        ) from exc

    request = google_drive_service().files().get_media(fileId=file_id, supportsAllDrives=True)
    output = BytesIO()
    downloader = MediaIoBaseDownload(output, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return output.getvalue()


def upload_drive_file(
    *,
    name: str,
    content: bytes,
    mime_type: str,
    file_id: str = "",
) -> str:
    try:
        from googleapiclient.http import MediaIoBaseUpload
    except ImportError as exc:
        raise RuntimeError(
            "Google Drive storage is configured, but Google API libraries are missing. "
            "Run: pip install -r requirements.txt"
        ) from exc

    service = google_drive_service()
    media = MediaIoBaseUpload(BytesIO(content), mimetype=mime_type, resumable=False)
    if file_id:
        result = (
            service.files()
            .update(fileId=file_id, body={"name": name}, media_body=media, fields="id", supportsAllDrives=True)
            .execute()
        )
    else:
        result = (
            service.files()
            .create(
                body={"name": name, "parents": [drive_folder_id()]},
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            )
            .execute()
        )
    return str(result["id"])


def delete_drive_file(file_id: str) -> None:
    if file_id:
        google_drive_service().files().delete(fileId=file_id, supportsAllDrives=True).execute()


def read_drive_registry_payload() -> dict[str, object]:
    global DRIVE_REGISTRY_CACHE
    global DRIVE_REGISTRY_CACHE_TIME

    now = time.time()
    if (
        DRIVE_REGISTRY_CACHE is not None
        and (now - DRIVE_REGISTRY_CACHE_TIME) < DRIVE_REGISTRY_CACHE_SECONDS
    ):
        return dict(DRIVE_REGISTRY_CACHE)

    file_id = find_drive_file_by_name(DRIVE_REGISTRY_FILENAME)
    if not file_id:
        return {"papers": [], "selected_paper_id": ""}
    try:
        content = download_drive_file(file_id).decode("utf-8")
        payload = json.loads(content)
    except Exception:
        return {"papers": [], "selected_paper_id": ""}
    if not isinstance(payload, dict):
        return {"papers": [], "selected_paper_id": ""}
    payload["_drive_file_id"] = file_id
    DRIVE_REGISTRY_CACHE = dict(payload)
    DRIVE_REGISTRY_CACHE_TIME = now
    return payload


def write_drive_registry_payload(payload: dict[str, object]) -> None:
    global DRIVE_REGISTRY_CACHE
    global DRIVE_REGISTRY_CACHE_TIME

    file_id = str(payload.pop("_drive_file_id", "")) or find_drive_file_by_name(DRIVE_REGISTRY_FILENAME)
    content = json.dumps(payload, indent=2).encode("utf-8")
    saved_file_id = upload_drive_file(
        name=DRIVE_REGISTRY_FILENAME,
        content=content,
        mime_type="application/json",
        file_id=file_id,
    )
    cached_payload = dict(payload)
    if saved_file_id:
        cached_payload["_drive_file_id"] = saved_file_id
    DRIVE_REGISTRY_CACHE = cached_payload
    DRIVE_REGISTRY_CACHE_TIME = time.time()


def normalize_paper(item: dict[str, object]) -> dict[str, str]:
    source_type = str(item.get("source_type", item.get("storage", "excel"))).strip() or "excel"
    google_sheet_url = str(item.get("google_sheet_url", "")).strip()
    drive_file_id = str(item.get("drive_file_id", "")).strip()
    file_value = str(item.get("file", "")).strip()
    if source_type == "google_sheet" and not google_sheet_url:
        google_sheet_url = file_value
    return {
        "id": str(item.get("id", "")).strip(),
        "name": str(item.get("name", "")).strip(),
        "file": file_value,
        "uploaded_at": str(item.get("uploaded_at", "")).strip(),
        "source_type": source_type,
        "google_sheet_url": google_sheet_url,
        "drive_file_id": drive_file_id,
        "rows": str(item.get("rows", "")).strip(),
        "columns": item.get("columns", []) if isinstance(item.get("columns", []), list) else [],
    }


def read_paper_registry() -> list[dict[str, str]]:
    if google_drive_enabled():
        payload = read_drive_registry_payload()
        papers = payload.get("papers", [])
        if not isinstance(papers, list):
            return []
        return [
            paper
            for paper in (normalize_paper(item) for item in papers if isinstance(item, dict))
            if paper["id"] and (paper["file"] or paper["drive_file_id"] or paper["google_sheet_url"])
        ]

    if not PAPER_REGISTRY_FILE.exists():
        return []
    try:
        with PAPER_REGISTRY_FILE.open("r", encoding="utf-8") as file:
            payload = json.load(file)
    except (json.JSONDecodeError, OSError):
        return []
    papers = payload if isinstance(payload, list) else payload.get("papers", [])
    return [
        paper
        for paper in (normalize_paper(item) for item in papers if isinstance(item, dict))
        if paper["id"] and (paper["file"] or paper["drive_file_id"] or paper["google_sheet_url"])
    ]


def write_paper_registry(papers: list[dict[str, str]]) -> None:
    if google_drive_enabled():
        payload = read_drive_registry_payload()
        payload["papers"] = papers
        write_drive_registry_payload(payload)
        return

    DATA_DIR.mkdir(exist_ok=True)
    with PAPER_REGISTRY_FILE.open("w", encoding="utf-8") as file:
        json.dump({"papers": papers}, file, indent=2)


def read_selected_paper_id() -> str:
    if google_drive_enabled():
        return str(read_drive_registry_payload().get("selected_paper_id", "")).strip()

    if not SELECTED_PAPER_FILE.exists():
        return ""
    try:
        with SELECTED_PAPER_FILE.open("r", encoding="utf-8") as file:
            payload = json.load(file)
    except (json.JSONDecodeError, OSError):
        return ""
    return str(payload.get("paper_id", "")).strip()


def write_selected_paper_id(paper_id: str) -> None:
    if google_drive_enabled():
        payload = read_drive_registry_payload()
        payload["selected_paper_id"] = paper_id
        write_drive_registry_payload(payload)
        return

    DATA_DIR.mkdir(exist_ok=True)
    with SELECTED_PAPER_FILE.open("w", encoding="utf-8") as file:
        json.dump({"paper_id": paper_id}, file, indent=2)


def ensure_default_paper_registry() -> list[dict[str, str]]:
    papers = read_paper_registry()
    if google_drive_enabled():
        if papers:
            selected_id = read_selected_paper_id()
            if not selected_id or not any(p["id"] == selected_id for p in papers):
                write_selected_paper_id(papers[0]["id"])
            return papers

        drive_papers = []

        try:
            drive_files = get_drive_excel_files()

            for item in drive_files:
                drive_papers.append(
                    {
                        "id": item["paper_id"],
                        "name": item["paper_name"],
                        "file": item["file_name"],
                        "uploaded_at": "",
                        "source_type": "drive_excel",
                        "google_sheet_url": "",
                        "drive_file_id": item["paper_id"],
                        "rows": "",
                        "columns": [],
                    }
                )

            papers = drive_papers

            write_paper_registry(
                papers
            )

        except Exception as exc:
            print(
                "Drive sync failed:",
                exc
            )

        if papers:

            selected_id = (
                read_selected_paper_id()
            )

            if (
                    not selected_id
                    or not any(
                p["id"] == selected_id
                for p in papers
            )
            ):
                write_selected_paper_id(
                    papers[0]["id"]
                )

        return papers

    papers_by_file = {str(paper_file_path(paper)).lower(): paper for paper in papers}
    discovered = False

    if PAPER_UPLOAD_DIR.exists():
        for excel_file in sorted(PAPER_UPLOAD_DIR.glob("*.xls*")):
            if excel_file.name.startswith("~$"):
                continue
            resolved = str(excel_file.resolve()).lower()
            if resolved in papers_by_file:
                continue
            paper_name = default_paper_name_from_file(excel_file)
            paper_id = slugify(paper_name)
            existing_ids = {paper["id"] for paper in papers}
            if paper_id in existing_ids:
                paper_id = slugify(excel_file.stem)
            suffix = 2
            base_id = paper_id
            while paper_id in existing_ids:
                paper_id = f"{base_id}-{suffix}"
                suffix += 1
            paper = {
                "id": paper_id,
                "name": paper_name,
                "file": str(excel_file),
                "uploaded_at": "",
            }
            papers.append(paper)
            papers_by_file[resolved] = paper
            discovered = True

    if discovered:
        write_paper_registry(papers)

    if papers:
        selected_id = read_selected_paper_id()
        if not selected_id or not any(paper["id"] == selected_id for paper in papers):
            write_selected_paper_id(papers[0]["id"])
        return papers
    if DEFAULT_DATA_FILE.exists():
        paper_name = default_paper_name_from_file(DEFAULT_DATA_FILE)
        paper = {
            "id": slugify(paper_name),
            "name": paper_name,
            "file": str(DEFAULT_DATA_FILE),
            "uploaded_at": "",
        }
        write_paper_registry([paper])
        write_selected_paper_id(paper["id"])
        return [paper]
    return []


def paper_file_path(paper: dict[str, str]) -> Path:
    path = Path(paper.get("file", ""))
    if not path.is_absolute():
        path = ROOT / path
    return path


def default_paper_name_from_file(path: Path) -> str:
    try:
        frame = pd.read_excel(path, dtype=object).fillna("")
        normalized_columns = {normalize_header(col): col for col in frame.columns}
        paper_column = normalized_columns.get("paper name")
        if paper_column:
            for value in frame[paper_column]:
                name = clean_cell(value)
                if name:
                    return name
    except Exception:
        pass
    return path.stem.replace("_", " ").replace("-", " ").title()


def get_paper(paper_id: str) -> dict[str, str] | None:
    normalized = paper_id.strip().lower()
    for paper in ensure_default_paper_registry():
        if paper["id"].lower() == normalized:
            return paper
    return None


def paper_cache_key(paper: dict[str, str]) -> str:
    source_fingerprint = "|".join(
        [
            str(paper.get("source_type", "")),
            str(paper.get("file", "")),
            str(paper.get("uploaded_at", "")),
            str(paper.get("drive_file_id", "")),
            str(paper.get("google_sheet_url", "")),
        ]
    )
    return f"{paper.get('id', '').lower()}:{source_fingerprint}"


def paper_metadata_from_frame(frame: pd.DataFrame) -> dict[str, object]:
    return {
        "rows": len(students_from_frame(frame)),
        "columns": [str(col) for col in frame.columns],
    }


def update_paper_metadata(paper_id: str, metadata: dict[str, object]) -> None:
    papers = read_paper_registry()
    changed = False
    for paper in papers:
        if paper["id"] != paper_id:
            continue
        paper["rows"] = str(metadata.get("rows", ""))
        paper["columns"] = metadata.get("columns", [])
        changed = True
        break
    if changed:
        write_paper_registry(papers)


def cached_paper_metadata(paper: dict[str, str]) -> dict[str, object]:
    rows = str(paper.get("rows", "")).strip()
    raw_columns = paper.get("columns", [])
    columns = raw_columns if isinstance(raw_columns, list) else []
    return {
        "rows": int(rows) if rows.isdigit() else None,
        "columns": [str(col) for col in columns],
    }


def delete_paper(paper_id: str) -> None:
    papers = ensure_default_paper_registry()
    paper = next((item for item in papers if item["id"] == paper_id), None)
    if not paper:
        raise FileNotFoundError("Selected paper was not found.")

    remaining = [item for item in papers if item["id"] != paper_id]
    if paper.get("source_type") == "drive_excel":
        delete_drive_file(paper.get("drive_file_id", ""))
    elif paper.get("source_type") != "google_sheet":
        path = paper_file_path(paper)
        if path.exists() and path.parent.resolve() == PAPER_UPLOAD_DIR.resolve():
            path.unlink()
    write_paper_registry(remaining)

    selected_id = read_selected_paper_id()
    if selected_id == paper_id:
        write_selected_paper_id(remaining[0]["id"] if remaining else "")

    clear_drive_cache()


def data_file() -> Path:
    configured = os.getenv("STUDENT_DATA_FILE")
    if configured:
        return Path(configured)
    config = read_source_config()
    configured_excel = config.get("excel_file", "")
    if config.get("source_type") == "excel" and configured_excel:
        excel_path = Path(configured_excel)
        if excel_path.exists():
            return excel_path
    if DEFAULT_DATA_FILE.exists():
        return DEFAULT_DATA_FILE
    return FALLBACK_DATA_FILE


def display_source_name(source_type: str, source: str, google_sheet_url: str) -> str:
    if source_type == "google":
        parsed = urlparse(google_sheet_url or source)
        parts = [part for part in parsed.path.split("/") if part]
        if "d" in parts:
            index = parts.index("d")
            if index + 1 < len(parts):
                return f"Google Sheet {parts[index + 1][:10]}..."
        return parsed.netloc or "Google Sheet"

    name = Path(source or str(data_file())).name
    return name or "Excel workbook"


def read_source_config() -> dict[str, str]:
    if not SOURCE_CONFIG_FILE.exists():
        return {"source_type": "excel", "google_sheet_url": "", "excel_file": ""}
    try:
        with SOURCE_CONFIG_FILE.open("r", encoding="utf-8") as file:
            config = json.load(file)
    except (json.JSONDecodeError, OSError):
        return {"source_type": "excel", "google_sheet_url": "", "excel_file": ""}
    return {
        "source_type": str(config.get("source_type", "excel")),
        "google_sheet_url": str(config.get("google_sheet_url", "")),
        "excel_file": str(config.get("excel_file", "")),
    }


def write_source_config(
    source_type: str,
    google_sheet_url: str = "",
    excel_file: str | None = None,
) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    current = read_source_config()
    payload = {
        "source_type": source_type,
        "google_sheet_url": google_sheet_url,
        "excel_file": current.get("excel_file", "") if excel_file is None else excel_file,
        "updated_at": int(time.time()),
    }
    with SOURCE_CONFIG_FILE.open("w", encoding="utf-8") as file:
        json.dump(payload, file, indent=2)


def google_sheet_csv_url(url: str) -> str:
    parsed = urlparse(url)
    if "docs.google.com" not in parsed.netloc or "/spreadsheets/" not in parsed.path:
        return url

    match = re.search(r"/spreadsheets/d/([^/]+)", parsed.path)
    if not match:
        return url

    sheet_id = match.group(1)
    query = parse_qs(parsed.query)
    gid = query.get("gid", ["0"])[0]
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def google_sheet_csv_urls(url: str) -> list[str]:
    parsed = urlparse(url)
    if "docs.google.com" not in parsed.netloc or "/spreadsheets/" not in parsed.path:
        return [url]

    match = re.search(r"/spreadsheets/d/([^/]+)", parsed.path)
    if not match:
        return [url]

    sheet_id = match.group(1)
    query = parse_qs(parsed.query)
    gid = query.get("gid", ["0"])[0]
    urls = [f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"]

    sheet_name = query.get("sheet", [""])[0]
    if sheet_name:
        urls.append(
            f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote(sheet_name)}"
        )
    urls.append(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv")
    return list(dict.fromkeys(urls))


def is_ssl_certificate_error(exc: Exception) -> bool:
    reason = getattr(exc, "reason", exc)
    if isinstance(reason, ssl.SSLCertVerificationError):
        return True
    return "CERTIFICATE_VERIFY_FAILED" in str(exc)


def download_google_sheet_csv(csv_url: str) -> str:
    request = urllib.request.Request(
        csv_url,
        headers={"User-Agent": "Mozilla/5.0 StudentPortal/1.0"},
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            return response.read().decode("utf-8-sig")
    except Exception as exc:
        parsed = urlparse(csv_url)
        if "docs.google.com" not in parsed.netloc or not is_ssl_certificate_error(exc):
            raise

        fallback_context = ssl._create_unverified_context()
        with urllib.request.urlopen(request, timeout=20, context=fallback_context) as response:
            return response.read().decode("utf-8-sig")


def read_google_sheet_dataframe(url: str) -> tuple[pd.DataFrame, str]:
    errors: list[str] = []
    for csv_url in google_sheet_csv_urls(url):
        try:
            csv_text = download_google_sheet_csv(csv_url)
            if "<html" in csv_text[:500].lower() or "<!doctype html" in csv_text[:500].lower():
                raise ValueError("Google returned an HTML page instead of CSV")
            frame = pd.read_csv(StringIO(csv_text), dtype=str).fillna("")
            if frame.empty and len(frame.columns) == 0:
                raise ValueError("Google Sheet returned no columns")
            return frame, csv_url
        except Exception as exc:
            errors.append(f"{csv_url}: {type(exc).__name__}: {exc}")

    detail = " | ".join(errors[-2:]) if errors else "No Google Sheet URL could be read"
    raise ValueError(
        "Could not read Google Sheet as CSV. Make sure it is shared as 'Anyone with the link can view'. "
        f"Details: {detail}"
    )


def save_excel_bytes(excel_bytes: bytes) -> Path:
    DATA_DIR.mkdir(exist_ok=True)
    try:
        DEFAULT_DATA_FILE.write_bytes(excel_bytes)
        return DEFAULT_DATA_FILE
    except PermissionError:
        fallback = DATA_DIR / f"student_data_uploaded_{int(time.time())}.xlsx"
        fallback.write_bytes(excel_bytes)
        return fallback


def save_paper_excel_bytes(
    paper_name: str,
    excel_bytes: bytes,
    frame: pd.DataFrame | None = None,
) -> dict[str, str]:
    papers = ensure_default_paper_registry()
    frame = frame if frame is not None else pd.read_excel(BytesIO(excel_bytes), dtype=object).fillna("")
    metadata = paper_metadata_from_frame(frame)
    base_id = slugify(paper_name)
    paper_id = base_id
    existing_ids = {paper["id"] for paper in papers if paper["id"] != base_id}
    if paper_id in existing_ids:
        paper_id = f"{base_id}-{uuid4().hex[:6]}"
    filename = f"{paper_id}.xlsx"

    replaced = False
    existing_drive_file_id = ""
    for existing in papers:
        if existing["id"] == base_id or normalize_header(existing["name"]) == normalize_header(paper_name):
            existing_drive_file_id = existing.get("drive_file_id", "")
            break

    if google_drive_enabled():
        drive_file_id = upload_drive_file(
            name=filename,
            content=excel_bytes,
            mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            file_id=existing_drive_file_id,
        )
        paper = {
            "id": paper_id,
            "name": paper_name.strip(),
            "file": filename,
            "uploaded_at": str(int(time.time())),
            "source_type": "drive_excel",
            "google_sheet_url": "",
            "drive_file_id": drive_file_id,
            "rows": str(metadata["rows"]),
            "columns": metadata["columns"],
        }
    else:
        PAPER_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        path = PAPER_UPLOAD_DIR / filename
        path.write_bytes(excel_bytes)
        paper = {
            "id": paper_id,
            "name": paper_name.strip(),
            "file": str(path),
            "uploaded_at": str(int(time.time())),
            "source_type": "excel",
            "google_sheet_url": "",
            "drive_file_id": "",
            "rows": str(metadata["rows"]),
            "columns": metadata["columns"],
        }

    for index, existing in enumerate(papers):
        if existing["id"] == base_id or normalize_header(existing["name"]) == normalize_header(paper_name):
            old_path = paper_file_path(existing)
            if (
                not google_drive_enabled()
                and old_path.exists()
                and old_path != Path(paper["file"])
                and old_path.parent == PAPER_UPLOAD_DIR
            ):
                try:
                    old_path.unlink()
                except OSError:
                    pass
            papers[index] = paper
            replaced = True
            break
    if not replaced:
        papers.append(paper)
    write_paper_registry(papers)
    clear_drive_cache()
    clear_paper_data_cache(paper_id)
    return paper


def save_google_sheet_paper(
    paper_name: str,
    google_sheet_url: str,
    frame: pd.DataFrame | None = None,
) -> dict[str, str]:
    papers = ensure_default_paper_registry()
    metadata = paper_metadata_from_frame(frame) if frame is not None else {"rows": "", "columns": []}
    base_id = slugify(paper_name)
    paper_id = base_id
    existing_ids = {paper["id"] for paper in papers if paper["id"] != base_id}
    if paper_id in existing_ids:
        paper_id = f"{base_id}-{uuid4().hex[:6]}"

    paper = {
        "id": paper_id,
        "name": paper_name.strip(),
        "file": google_sheet_url,
        "uploaded_at": str(int(time.time())),
        "source_type": "google_sheet",
        "google_sheet_url": google_sheet_url,
        "drive_file_id": "",
        "rows": str(metadata["rows"]),
        "columns": metadata["columns"],
    }

    replaced = False
    for index, existing in enumerate(papers):
        if existing["id"] == base_id or normalize_header(existing["name"]) == normalize_header(paper_name):
            if existing.get("source_type") == "drive_excel":
                delete_drive_file(existing.get("drive_file_id", ""))
            elif existing.get("source_type") != "google_sheet":
                old_path = paper_file_path(existing)
                if old_path.exists() and old_path.parent == PAPER_UPLOAD_DIR:
                    try:
                        old_path.unlink()
                    except OSError:
                        pass
            papers[index] = paper
            replaced = True
            break
    if not replaced:
        papers.append(paper)
    write_paper_registry(papers)
    clear_paper_data_cache(paper_id)
    return paper


def load_dataframe() -> tuple[pd.DataFrame, str]:
    config = read_source_config()
    if config["source_type"] == "google" and config["google_sheet_url"]:
        frame, csv_url = read_google_sheet_dataframe(config["google_sheet_url"])
        return frame, csv_url

    source_file = data_file()
    if not source_file.exists():
        raise FileNotFoundError(
            f"Excel file not found at {source_file}. Put your workbook at data/student_data.xlsx."
        )
    frame = pd.read_excel(source_file, dtype=str).fillna("")
    return frame, str(source_file)


def load_paper_dataframe(paper_id: str) -> tuple[pd.DataFrame, str, dict[str, str]]:
    paper = get_paper(paper_id)
    if not paper:
        raise FileNotFoundError("Selected paper was not found.")

    cache_key = paper_cache_key(paper)
    cached = PAPER_DATAFRAME_CACHE.get(cache_key)
    if cached and (time.time() - cached[0]) < PAPER_DATA_CACHE_SECONDS:
        _, frame, source, cached_paper = cached
        return frame.copy(), source, dict(cached_paper)

    if paper.get("source_type") == "google_sheet":
        google_sheet_url = paper.get("google_sheet_url") or paper.get("file", "")
        frame, csv_url = read_google_sheet_dataframe(google_sheet_url)
        PAPER_DATAFRAME_CACHE[cache_key] = (time.time(), frame.copy(), csv_url, dict(paper))
        return frame, csv_url, paper

    if paper.get("source_type") == "drive_excel":
        drive_file_id = paper.get(
            "drive_file_id",
            ""
        )
        if not drive_file_id:
            raise FileNotFoundError(
                f"Google Drive file id was not found for {paper['name']}."
            )
        excel_bytes = download_drive_file(drive_file_id)

        frame = pd.read_excel(
            BytesIO(
                excel_bytes
            ),
            dtype=object
        ).fillna("")

        source = f"Google Drive: {paper.get('file', paper['name'])}"
        PAPER_DATAFRAME_CACHE[cache_key] = (time.time(), frame.copy(), source, dict(paper))
        return (
            frame,
            source,
            paper
        )

    source_file = paper_file_path(paper)
    if not source_file.exists():
        raise FileNotFoundError(f"Excel file not found for {paper['name']}.")
    frame = pd.read_excel(source_file, dtype=object).fillna("")
    PAPER_DATAFRAME_CACHE[cache_key] = (time.time(), frame.copy(), str(source_file), dict(paper))
    return frame, str(source_file), paper


def attendance_month_columns(frame: pd.DataFrame) -> list[dict[str, object]]:
    months: list[dict[str, object]] = []
    for column in frame.columns:
        label = str(column).replace("\n", " ").strip()
        normalized = normalize_header(label)
        match = re.match(r"^(.+?)\s*\((\d+(?:\.\d+)?)\)$", normalized)
        if not match:
            continue
        month = match.group(1).strip()
        if month not in MONTH_NAMES:
            continue
        total = float(match.group(2))
        months.append(
            {
                "column": column,
                "label": label,
                "month": month.title(),
                "total": int(total) if total.is_integer() else total,
            }
        )
    return months


def format_percentage(value: str) -> str:
    if value == "":
        return ""
    try:
        number = float(value)
    except ValueError:
        return value
    if 0 <= number <= 1:
        number *= 100
    return f"{number:.1f}".rstrip("0").rstrip(".")


def students_from_frame(frame: pd.DataFrame) -> list[dict[str, str]]:
    normalized_columns = {normalize_header(col): col for col in frame.columns}
    month_columns = attendance_month_columns(frame)

    column_map: dict[str, str] = {}
    for target, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in normalized_columns:
                column_map[target] = normalized_columns[alias]
                break

    required = ["name", "exam_roll_number", "email"]
    missing = [key for key in required if key not in column_map]
    if missing:
        readable = ", ".join(missing)
        raise ValueError(f"Missing required Excel columns: {readable}")

    students: list[dict[str, str]] = []
    for _, row in frame.iterrows():
        student = {}
        for target in COLUMN_ALIASES:
            source = column_map.get(target)
            student[target] = clean_cell(row[source]) if source else ""
        student["exam_roll_number"] = clean_roll(student["exam_roll_number"])
        student["college_roll_number"] = clean_roll(student["college_roll_number"])
        if not student["paper_name"]:
            student["paper_name"] = clean_cell(row[column_map["paper_name"]]) if "paper_name" in column_map else ""
        student["attendance_percentage"] = format_percentage(student["attendance_percentage"])
        student["attendance_months"] = [
            {
                "label": str(item["label"]),
                "month": str(item["month"]),
                "total": item["total"],
                "attended": clean_cell(row[item["column"]]),
            }
            for item in month_columns
        ]
        student["display_rows"] = [
            {"label": str(column).replace("\n", " ").strip(), "value": clean_cell(row[column])}
            for column in frame.columns
            if clean_cell(row[column]) != ""
        ]
        if student["exam_roll_number"]:
            students.append(student)
    return students


def load_students() -> list[dict[str, str]]:
    frame, _ = load_dataframe()
    return students_from_frame(frame)


def find_student(exam_roll_number: str, paper_id: str = "") -> dict[str, str] | None:
    normalized_roll = exam_roll_number.strip().lower()
    if paper_id:
        frame, _, paper = load_paper_dataframe(paper_id)
        students = students_from_frame(frame)
        if str(paper.get("rows", "")).strip() == "":
            update_paper_metadata(
                paper["id"],
                {"rows": len(students), "columns": [str(col) for col in frame.columns]},
            )
        for student in students:
            if not student.get("paper_name"):
                student["paper_name"] = paper["name"]
    else:
        students = load_students()
    for student in students:
        if student["exam_roll_number"].strip().lower() == normalized_roll:
            return student
    return None


def mask_email(email: str) -> str:
    if "@" not in email:
        return email[:2] + "***"
    user, domain = email.split("@", 1)
    visible = user[:2] if len(user) > 2 else user[:1]
    return f"{visible}{'*' * max(3, len(user) - len(visible))}@{domain}"


def send_email_otp(email: str, otp: str, name: str) -> tuple[bool, str]:
    if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
        return False, "student email is missing or invalid"
    if email.lower().endswith("@example.com"):
        return False, "student email is a placeholder example.com address"

    host = os.getenv("SMTP_HOST")
    user = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASS")
    sender = os.getenv("SMTP_FROM", user or "no-reply@student-portal.local")
    port = int(os.getenv("SMTP_PORT", "587"))
    security = smtp_security_mode(port)

    if not host or not user or not password:
        return False, "SMTP_HOST, SMTP_USER, or SMTP_PASS is missing"

    subject = "Your Login OTP"
    body = (
        f"Hello {name},\n\nYour OTP for the student portal is {otp}. "
        "It will expire in 10 minutes.\n\nThank you."
    )

    if os.name == "nt":
        curl_result = send_email_otp_with_curl(
            host=host,
            port=port,
            security=security,
            user=user,
            password=password,
            sender=sender,
            recipient=email,
            subject=subject,
            body=body,
        )
        if curl_result[0] or "curl.exe was not found" not in curl_result[1]:
            return curl_result
        if security == "starttls":
            return send_email_otp_with_powershell(
                host=host,
                port=port,
                user=user,
                password=password,
                sender=sender,
                recipient=email,
                subject=subject,
                body=body,
            )
        return curl_result

    smtp_result = send_email_otp_with_smtplib(
        host=host,
        port=port,
        security=security,
        user=user,
        password=password,
        sender=sender,
        recipient=email,
        subject=subject,
        body=body,
    )
    return smtp_result


def send_email_otp_with_smtplib(
    *,
    host: str,
    port: int,
    security: str,
    user: str,
    password: str,
    sender: str,
    recipient: str,
    subject: str,
    body: str,
) -> tuple[bool, str]:
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = recipient
    message.set_content(body)

    try:
        smtp_class = smtplib.SMTP_SSL if security == "ssl" else smtplib.SMTP
        with smtp_class(host, port, timeout=15) as smtp:
            if security == "starttls":
                smtp.starttls()
            smtp.login(user, password)
            smtp.send_message(message)
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"
    return True, "sent"


def send_email_otp_with_curl(
    *,
    host: str,
    port: int,
    security: str,
    user: str,
    password: str,
    sender: str,
    recipient: str,
    subject: str,
    body: str,
) -> tuple[bool, str]:
    curl = shutil.which("curl.exe") or shutil.which("curl")
    if not curl:
        return False, "curl.exe was not found for local Windows SMTP sending"

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = recipient
    message.set_content(body)

    scheme = "smtps" if security == "ssl" else "smtp"
    command = [
        curl,
        "--silent",
        "--show-error",
        "--tlsv1.2",
        *(["--ssl-no-revoke"] if os.name == "nt" else []),
        *(["--ssl-reqd"] if security == "starttls" else []),
        "--url",
        f"{scheme}://{host}:{port}",
        "--mail-from",
        sender,
        "--mail-rcpt",
        recipient,
        "--user",
        f"{user}:{password}",
        "--upload-file",
        "-",
    ]

    try:
        result = subprocess.run(
            command,
            input=message.as_string(),
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"

    if result.returncode != 0:
        details = (result.stderr or result.stdout or "curl SMTP send failed").strip()
        return False, details
    return True, "sent"


def send_email_otp_with_powershell(
    *,
    host: str,
    port: int,
    user: str,
    password: str,
    sender: str,
    recipient: str,
    subject: str,
    body: str,
) -> tuple[bool, str]:
    script = r"""
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$message = [System.Net.Mail.MailMessage]::new(
  $env:OTP_SMTP_FROM,
  $env:OTP_SMTP_TO,
  $env:OTP_SMTP_SUBJECT,
  $env:OTP_SMTP_BODY
)
$client = [System.Net.Mail.SmtpClient]::new($env:OTP_SMTP_HOST, [int]$env:OTP_SMTP_PORT)
$client.EnableSsl = $true
$client.Credentials = [System.Net.NetworkCredential]::new($env:OTP_SMTP_USER, $env:OTP_SMTP_PASS)
try {
  $client.Send($message)
}
catch {
  Write-Error $_.Exception.ToString()
  exit 1
}
finally {
  $message.Dispose()
  $client.Dispose()
}
"""
    env = {
        **os.environ,
        "OTP_SMTP_HOST": host,
        "OTP_SMTP_PORT": str(port),
        "OTP_SMTP_USER": user,
        "OTP_SMTP_PASS": password,
        "OTP_SMTP_FROM": sender,
        "OTP_SMTP_TO": recipient,
        "OTP_SMTP_SUBJECT": subject,
        "OTP_SMTP_BODY": body,
    }
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            env=env,
            text=True,
            timeout=30,
            check=False,
        )
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"

    if result.returncode != 0:
        details = (result.stderr or result.stdout or "PowerShell SMTP send failed").strip()
        return False, details
    return True, "sent"


def public_student(student: dict[str, str]) -> dict[str, str]:
    labels = {
        "name": "Name",
        "course_name": "Course Name",
        "paper_name": "Paper Name",
        "college_roll_number": "College Roll Number",
        "exam_roll_number": "Exam Roll Number",
        "lectures_taken": "Lectures Taken",
        "total_attendance": "Total Lectures Attended",
        "assignment_marks": "Assignment Marks",
        "test_marks": "Test Marks",
        "attendance_percentage": "Attendance Percentage",
        "attendance_marks": "Attendance Marks",
        "internal_marks": "IAEE/IEES Internals",
    }
    public = {key: student.get(key, "") for key in labels}
    public["attendance_months"] = student.get("attendance_months", [])
    public["display_rows"] = student.get("display_rows", [])
    return public


class StudentPortalHandler(SimpleHTTPRequestHandler):
    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def translate_path(self, path: str) -> str:
        path = unquote(path.split("?", 1)[0].split("#", 1)[0])
        if path == "/":
            path = "/index.html"
        return str((ROOT / path.lstrip("/")).resolve())

    def do_POST(self) -> None:
        if self.path == "/api/request-otp":
            self.handle_request_otp()
            return
        if self.path == "/api/verify-otp":
            self.handle_verify_otp()
            return
        if self.path == "/api/admin/source":
            self.handle_admin_set_google_source()
            return
        if self.path == "/api/admin/upload":
            self.handle_admin_upload_excel()
            return
        if self.path == "/api/admin/upload-google":
            self.handle_admin_upload_google_sheet()
            return
        if self.path == "/api/admin/paper/select":
            self.handle_admin_select_paper()
            return
        if self.path == "/api/admin/paper/delete":
            self.handle_admin_delete_paper()
            return
        self.send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)

    def do_GET(self) -> None:
        if self.path == "/api/admin/status":
            self.handle_admin_status()
            return
        if self.path == "/api/papers":
            self.handle_papers()
            return
        super().do_GET()

    def read_json(self) -> dict[str, object]:
        length = int(self.headers.get("Content-Length", "0"))
        payload = self.rfile.read(length).decode("utf-8")
        return json.loads(payload or "{}")

    def send_json(self, payload: dict[str, object], status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload).encode("utf-8")
        try:
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
            # Browsers may cancel stale dropdown requests during reloads.
            return

    def admin_authorized(self) -> bool:
        expected = os.getenv("ADMIN_PASSWORD", "admin123")
        received = self.headers.get("X-Admin-Password", "")
        return bool(expected) and received == expected

    def require_admin(self) -> bool:
        if self.admin_authorized():
            return True
        self.send_json({"error": "Invalid admin password."}, HTTPStatus.UNAUTHORIZED)
        return False

    def handle_admin_status(self) -> None:
        if not self.require_admin():
            return

        ensure_default_paper_registry()

        try:
            papers = []
            total_rows = 0
            selected_rows = 0
            all_columns: list[str] = []
            registry = ensure_default_paper_registry()
            selected_paper_id = read_selected_paper_id()
            selected_paper = next((item for item in registry if item["id"] == selected_paper_id), None)
            for paper in registry:
                metadata = cached_paper_metadata(paper)
                row_count = metadata["rows"]
                columns = metadata["columns"]
                if row_count is None:
                    row_count = 0
                total_rows += row_count
                if paper["id"] == selected_paper_id:
                    selected_rows = row_count
                for column in columns:
                    if column not in all_columns:
                        all_columns.append(column)
                papers.append(
                    {
                        "id": paper["id"],
                        "name": paper["name"],
                        "sourceType": paper.get("source_type", "excel"),
                        "source": paper.get("file", ""),
                        "rows": metadata["rows"],
                        "columns": columns,
                        "uploadedAt": paper.get("uploaded_at", ""),
                        "selected": paper["id"] == selected_paper_id,
                    }
                )
            config = read_source_config()
            self.send_json(
                {
                    "sourceType": "paper_uploads",
                    "sourceLabel": "Google Drive Papers" if google_drive_enabled() else "Paper-wise Excel Files",
                    "sourceName": selected_paper["name"] if selected_paper else "No paper selected",
                    "source": "Google Drive" if google_drive_enabled() else str(PAPER_UPLOAD_DIR),
                    "googleSheetUrl": config["google_sheet_url"],
                    "rows": selected_rows if selected_paper else total_rows,
                    "totalRows": total_rows,
                    "columns": all_columns,
                    "papers": papers,
                    "excelFile": "Google Drive" if google_drive_enabled() else str(PAPER_UPLOAD_DIR),
                }
            )
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_admin_select_paper(self) -> None:
        if not self.require_admin():
            return
        try:
            data = self.read_json()
            paper_id = str(data.get("paperId", "")).strip()
            if not paper_id or not get_paper(paper_id):
                self.send_json({"error": "Selected paper was not found."}, HTTPStatus.NOT_FOUND)
                return
            write_selected_paper_id(paper_id)
            self.handle_admin_status()
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_admin_delete_paper(self) -> None:
        if not self.require_admin():
            return
        try:
            data = self.read_json()
            paper_id = str(data.get("paperId", "")).strip()
            if not paper_id:
                self.send_json({"error": "Paper id is required."}, HTTPStatus.BAD_REQUEST)
                return
            delete_paper(paper_id)
            self.handle_admin_status()
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_papers(self) -> None:
        try:
            papers = []
            for paper in ensure_default_paper_registry():
                if paper.get("source_type") in {"google_sheet", "drive_excel"}:
                    papers.append({"id": paper["id"], "name": paper["name"]})
                    continue
                path = paper_file_path(paper)
                if path.exists():
                    papers.append({"id": paper["id"], "name": paper["name"]})
            self.send_json({"papers": papers})
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
            return
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_admin_set_google_source(self) -> None:
        if not self.require_admin():
            return
        try:
            data = self.read_json()
            source_type = str(data.get("sourceType", "")).strip().lower()
            google_sheet_url = str(data.get("googleSheetUrl", "")).strip()

            if source_type == "excel":
                write_source_config("excel", "")
            elif source_type == "google":
                if not google_sheet_url:
                    self.send_json({"error": "Google Sheet link is required."}, HTTPStatus.BAD_REQUEST)
                    return
                frame, _ = read_google_sheet_dataframe(google_sheet_url)
                students_from_frame(frame)
                write_source_config("google", google_sheet_url)
            else:
                self.send_json({"error": "sourceType must be excel or google."}, HTTPStatus.BAD_REQUEST)
                return

            self.handle_admin_status()
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_admin_upload_excel(self) -> None:
        if not self.require_admin():
            return
        try:
            content_type = self.headers.get("Content-Type", "")
            boundary_match = re.search(r"boundary=(.+)", content_type)
            if not boundary_match:
                self.send_json({"error": "Upload must be multipart/form-data."}, HTTPStatus.BAD_REQUEST)
                return

            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length)
            boundary = ("--" + boundary_match.group(1).strip().strip('"')).encode("utf-8")
            fields = self.extract_multipart_fields(body, boundary)
            paper_name = clean_cell(fields.get("paperName", ""))
            excel_bytes = fields.get("file", b"")
            if not paper_name:
                self.send_json({"error": "Paper name is required."}, HTTPStatus.BAD_REQUEST)
                return
            if not excel_bytes:
                self.send_json({"error": "No Excel file was uploaded."}, HTTPStatus.BAD_REQUEST)
                return

            frame = pd.read_excel(BytesIO(excel_bytes), dtype=object).fillna("")
            students_from_frame(frame)

            saved_paper = save_paper_excel_bytes(paper_name, excel_bytes, frame)
            write_selected_paper_id(saved_paper["id"])
            if not google_drive_enabled():
                write_source_config("excel", "", saved_paper["file"])
            self.handle_admin_status()
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_admin_upload_google_sheet(self) -> None:
        if not self.require_admin():
            return
        try:
            data = self.read_json()
            paper_name = clean_cell(data.get("paperName", ""))
            google_sheet_url = str(data.get("googleSheetUrl", "")).strip()
            if not paper_name:
                self.send_json({"error": "Paper name is required."}, HTTPStatus.BAD_REQUEST)
                return
            if not google_sheet_url:
                self.send_json({"error": "Google Sheet link is required."}, HTTPStatus.BAD_REQUEST)
                return

            frame, _ = read_google_sheet_dataframe(google_sheet_url)
            students_from_frame(frame)
            saved_paper = save_google_sheet_paper(paper_name, google_sheet_url, frame)
            write_selected_paper_id(saved_paper["id"])
            if not google_drive_enabled():
                write_source_config("google", google_sheet_url)
            self.handle_admin_status()
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def extract_uploaded_file(self, body: bytes, boundary: bytes) -> bytes:
        for part in body.split(boundary):
            if b'filename="' not in part:
                continue
            header_end = part.find(b"\r\n\r\n")
            if header_end == -1:
                continue
            content = part[header_end + 4 :]
            content = content.rstrip(b"\r\n")
            if content.endswith(b"--"):
                content = content[:-2].rstrip(b"\r\n")
            return content
        return b""

    def extract_multipart_fields(self, body: bytes, boundary: bytes) -> dict[str, object]:
        fields: dict[str, object] = {}
        for part in body.split(boundary):
            header_end = part.find(b"\r\n\r\n")
            if header_end == -1:
                continue
            header = part[:header_end].decode("utf-8", errors="ignore")
            name_match = re.search(r'name="([^"]+)"', header)
            if not name_match:
                continue
            name = name_match.group(1)
            content = part[header_end + 4 :]
            content = content.rstrip(b"\r\n")
            if content.endswith(b"--"):
                content = content[:-2].rstrip(b"\r\n")
            if 'filename="' in header:
                fields[name] = content
            else:
                fields[name] = content.decode("utf-8", errors="ignore").strip()
        return fields

    def handle_request_otp(self) -> None:
        try:
            data = self.read_json()
            paper_id = str(data.get("paperId", "")).strip()
            exam_roll_number = str(data.get("rollNumber", "")).strip()
            if not paper_id:
                self.send_json({"error": "Please choose a paper."}, HTTPStatus.BAD_REQUEST)
                return
            if not exam_roll_number:
                self.send_json({"error": "Please enter your exam roll number."}, HTTPStatus.BAD_REQUEST)
                return

            student = find_student(exam_roll_number, paper_id)
            if not student:
                self.send_json({"error": "No student found for this paper and exam roll number."}, HTTPStatus.NOT_FOUND)
                return

            otp = f"{random.randint(100000, 999999)}"
            email_sent, email_status = send_email_otp(student["email"], otp, student["name"])
            print(
                f"OTP request for exam roll {exam_roll_number}: "
                f"email={mask_email(student['email'])}, sent={email_sent}, status={email_status}"
            )
            otp_key = f"{paper_id.lower()}::{exam_roll_number.lower()}"
            OTP_STORE[otp_key] = {
                "otp": otp,
                "expires_at": time.time() + OTP_TTL_SECONDS,
            }

            response: dict[str, object] = {
                "message": "OTP sent successfully." if email_sent else "Demo OTP generated.",
                "emailMasked": mask_email(student["email"]),
                "emailSent": email_sent,
            }
            if not email_sent:
                response["emailStatus"] = email_status
                if smtp_is_configured():
                    OTP_STORE.pop(otp_key, None)
                    self.send_json(
                        {
                            "error": (
                                "OTP email could not be sent. Please check SMTP settings "
                                f"or the sender account. Mail server response: {email_status}"
                            ),
                            "emailMasked": mask_email(student["email"]),
                            "emailStatus": email_status,
                        },
                        HTTPStatus.BAD_GATEWAY,
                    )
                    return
                response["demoOtp"] = otp
            self.send_json(response)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_verify_otp(self) -> None:
        try:
            data = self.read_json()
            paper_id = str(data.get("paperId", "")).strip()
            exam_roll_number = str(data.get("rollNumber", "")).strip()
            otp = str(data.get("otp", "")).strip()
            otp_key = f"{paper_id.lower()}::{exam_roll_number.lower()}"
            record = OTP_STORE.get(otp_key)

            if not record:
                self.send_json({"error": "Please request an OTP first."}, HTTPStatus.BAD_REQUEST)
                return
            if time.time() > float(record["expires_at"]):
                OTP_STORE.pop(otp_key, None)
                self.send_json({"error": "OTP expired. Please request a new one."}, HTTPStatus.BAD_REQUEST)
                return
            if otp != record["otp"]:
                self.send_json({"error": "Invalid OTP. Please check and try again."}, HTTPStatus.UNAUTHORIZED)
                return

            student = find_student(exam_roll_number, paper_id)
            if not student:
                self.send_json({"error": "Student record was not found."}, HTTPStatus.NOT_FOUND)
                return

            OTP_STORE.pop(otp_key, None)
            self.send_json({"student": public_student(student)})
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)


def main() -> None:
    load_env_file()
    port = int(os.getenv("PORT", "8000"))
    host = os.getenv("HOST", "0.0.0.0")
    server = ThreadingHTTPServer((host, port), StudentPortalHandler)
    local_url = f"http://127.0.0.1:{port}"
    print(f"Student portal running on {host}:{port}")
    print(f"Open student page: {local_url}")
    print(f"Open admin panel: {local_url}/admin.html")
    print(f"Project folder: {ROOT}")
    print(f"Excel source: {data_file()}")
    print(f"SMTP configured: {'yes' if smtp_is_configured() else 'no'}")
    try:
        papers = ensure_default_paper_registry()
        if papers:
            print(f"Loaded papers: {len(papers)}")
            print(f"Paper preview: {', '.join(paper['name'] for paper in papers[:5])}")
        else:
            students = load_students()
            preview = ", ".join(
                f"{student['exam_roll_number']} -> {mask_email(student['email'])}"
                for student in students[:5]
            )
            print(f"Loaded students: {len(students)}")
            print(f"Email preview: {preview}")
    except Exception as exc:
        print(f"Startup data check skipped: {exc}")
    server.serve_forever()


if __name__ == "__main__":
    main()
