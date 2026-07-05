from __future__ import annotations

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from datetime import datetime
import json
import mimetypes
import os
import re
import tempfile
import traceback
from urllib.parse import quote
from urllib.parse import unquote

from report_engine import JOBS_DIR, build_report
from voc_product_runner import build_voc_product_report


ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
HISTORY_FILE = JOBS_DIR / "import_history.json"
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "300"))


def read_history() -> list[dict]:
    if not HISTORY_FILE.exists():
        return []
    try:
        data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def write_history(items: list[dict]) -> None:
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_FILE.write_text(json.dumps(items[:200], ensure_ascii=False, indent=2), encoding="utf-8")


def append_history(result: dict, saved: list[Path], metadata: dict[str, str] | None = None) -> dict:
    metadata = metadata or {}
    item = {
        "jobId": result.get("jobId"),
        "createdAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "asinCount": result.get("asinCount", 0),
        "modules": result.get("modules", []),
        "downloadUrl": result.get("downloadUrl", ""),
        "importer": metadata.get("importer", ""),
        "productCategory": metadata.get("productCategory", ""),
        "styleNumber": metadata.get("styleNumber", ""),
        "files": [
            {
                "name": path.name,
                "size": path.stat().st_size if path.exists() else 0,
            }
            for path in saved
        ],
        "asins": [record.get("竞对 ASIN") for record in result.get("records", []) if record.get("竞对 ASIN")],
    }
    history = read_history()
    history.insert(0, item)
    write_history(history)
    return item


def parse_multipart(body: bytes, content_type: str) -> tuple[list[tuple[str, bytes]], dict[str, str]]:
    match = re.search(r"boundary=([^;]+)", content_type)
    if not match:
        return [], {}
    boundary = ("--" + match.group(1).strip('"')).encode()
    files = []
    fields = {}
    for part in body.split(boundary):
        part = part.strip()
        if not part or part == b"--":
            continue
        if b"\r\n\r\n" not in part:
            continue
        head, data = part.split(b"\r\n\r\n", 1)
        data = data.rstrip(b"\r\n")
        if data.endswith(b"--"):
            data = data[:-2]
        headers = head.decode("utf-8", errors="ignore")
        name_match = re.search(r'name="([^"]*)"', headers)
        filename_match = re.search(r'filename="([^"]*)"', headers)
        if filename_match and filename_match.group(1):
            filename = Path(filename_match.group(1)).name
            files.append((filename, data))
        elif name_match and name_match.group(1):
            fields[name_match.group(1)] = data.decode("utf-8", errors="ignore").strip()
    return files, fields


def append_history(result: dict, saved: list[Path], metadata: dict[str, str] | None = None) -> dict:
    metadata = metadata or {}
    asins = list(result.get("asins") or [])
    for record in result.get("records", []):
        asin = record.get("竞对 ASIN") or record.get("ASIN")
        if not asin:
            for value in record.values():
                if isinstance(value, str):
                    match = re.search(r"B0[A-Z0-9]{8}", value.upper())
                    if match:
                        asin = match.group(0)
                        break
        if asin:
            asins.append(asin)
    item = {
        "jobId": result.get("jobId"),
        "createdAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "asinCount": result.get("asinCount", 0),
        "modules": result.get("modules", []),
        "downloadUrl": result.get("downloadUrl", ""),
        "importer": metadata.get("importer", ""),
        "productCategory": metadata.get("productCategory", ""),
        "styleNumber": metadata.get("styleNumber", ""),
        "reportMode": metadata.get("reportMode", ""),
        "reportType": metadata.get("reportType", ""),
        "llm": result.get("llm", {}),
        "files": [
            {
                "name": path.name,
                "size": path.stat().st_size if path.exists() else 0,
            }
            for path in saved
        ],
        "asins": asins,
    }
    history = read_history()
    history.insert(0, item)
    write_history(history)
    return item


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        return

    def send_json(self, payload, status=200):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = unquote(self.path.split("?", 1)[0])
        if path == "/":
            return self.serve_file(STATIC / "index.html")
        if path == "/api/history":
            return self.send_json({"items": read_history()})
        if path.startswith("/static/"):
            return self.serve_file(STATIC / path.replace("/static/", "", 1))
        if path.startswith("/download/"):
            _, _, job_id, filename = path.split("/", 3)
            return self.serve_file(JOBS_DIR / job_id / filename, download=True)
        self.send_error(404)

    def serve_file(self, path: Path, download=False):
        if not path.exists() or not path.is_file():
            self.send_error(404)
            return
        data = path.read_bytes()
        mime = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(data)))
        if download:
            fallback = re.sub(r"[^A-Za-z0-9_.-]+", "_", path.name)
            encoded = quote(path.name)
            self.send_header("Content-Disposition", f"attachment; filename=\"{fallback}\"; filename*=UTF-8''{encoded}")
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        if self.path != "/api/generate":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length > MAX_UPLOAD_MB * 1024 * 1024:
                return self.send_json({"error": f"上传文件过大，当前限制为 {MAX_UPLOAD_MB} MB"}, 413)
            body = self.rfile.read(length)
            files, fields = parse_multipart(body, self.headers.get("Content-Type", ""))
            if not files:
                return self.send_json({"error": "没有收到Excel文件"}, 400)
            if not fields.get("importer"):
                return self.send_json({"error": "请填写导入人"}, 400)
            temp_dir = Path(tempfile.mkdtemp(prefix="tank_report_upload_"))
            saved = []
            for filename, data in files:
                if not filename.lower().endswith((".xlsx", ".xlsm")):
                    continue
                target = temp_dir / filename
                target.write_bytes(data)
                saved.append(target)
            if not saved:
                return self.send_json({"error": "请上传 .xlsx 或 .xlsm 文件"}, 400)
            report_mode = fields.get("reportMode", "vocProduct")
            metadata = {
                "importer": fields.get("importer", ""),
                "productCategory": fields.get("productCategory", ""),
                "styleNumber": fields.get("styleNumber", ""),
                "reportMode": report_mode,
                "reportType": fields.get("reportType", "tank"),
                "llmEnabled": fields.get("llmEnabled", "false"),
                "llmModel": fields.get("llmModel", ""),
                "llmBaseUrl": fields.get("llmBaseUrl", ""),
            }
            llm_config = {
                "enabled": fields.get("llmEnabled", "false"),
                "protocol": fields.get("llmProtocol", "openai-compatible"),
                "baseUrl": fields.get("llmBaseUrl", ""),
                "model": fields.get("llmModel", ""),
                "apiKey": fields.get("llmApiKey", ""),
            }
            if report_mode == "competitor":
                result = build_report(saved)
            else:
                result = build_voc_product_report(saved, metadata, llm_config)
            result["importMeta"] = metadata
            result["historyItem"] = append_history(result, saved, result["importMeta"])
            self.send_json(result)
        except Exception as exc:
            traceback.print_exc()
            self.send_json({"error": str(exc)}, 500)


def main():
    port = int(os.environ.get("PORT", "8787"))
    host = os.environ.get("HOST", "0.0.0.0")
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
