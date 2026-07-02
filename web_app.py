from __future__ import annotations

import cgi
import html
import json
from functools import lru_cache
import os
import re
import shutil
import socket
import subprocess
import tempfile
import threading
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

from fill_sales import (
    BUSINESS_OWNERS,
    SUPPORTED_PREDICTION_SUFFIXES,
    SUPPORTED_SALES_SUFFIXES,
    process_sales_workbooks,
)


HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "8000"))
UPLOAD_MIME = "multipart/form-data"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PUBLIC_URL_ENV_VARS = ("PUBLIC_URL", "RENDER_EXTERNAL_URL", "APP_URL", "SITE_URL")
DEFAULT_OWNER = BUSINESS_OWNERS[0]
DATA_DIR = Path(os.environ.get("SALES_DATA_DIR", "/data/sales-upload" if Path("/data").exists() else "server_data"))
MASTER_SALES_PATH = DATA_DIR / "current_sales.xlsx"
LATEST_OUTPUT_PATH = DATA_DIR / "latest_generated.xlsx"
METADATA_PATH = DATA_DIR / "metadata.json"
BACKUP_DIR = DATA_DIR / "backups"
UPLOAD_DIR = DATA_DIR / "uploads"
REMOTE_STORAGE_ENDPOINT = os.environ.get("NETLIFY_BLOBS_ENDPOINT", "").strip().rstrip("/")
REMOTE_STORAGE_SECRET = os.environ.get("SALES_STORAGE_SECRET", "").strip()
REMOTE_MASTER_KEY = "current_sales.xlsx"
REMOTE_LATEST_KEY = "latest_generated.xlsx"
REMOTE_METADATA_KEY = "metadata.json"
REMOTE_CHUNK_SIZE = 2 * 1024 * 1024
STATE_LOCK = threading.Lock()
JOB_LOCK = threading.Lock()
JOB_STATUS = {
    "state": "idle",
    "owner": "",
    "message": "",
    "started_at": "",
    "finished_at": "",
    "updated_rows": "",
    "error": "",
}
OWNER_GUIDES = {
    "洪鸣": "读取预测文件中的 New part NO 匹配销售排单“客户机种”。",
    "李玎玲": "优先识别截图里的“机种名”和 6-10 月预测数量，自动匹配销售排单“客户机种”；数量单位 K 会自动换算成万 pcs。",
    "周文龙": "读取预测文件中的“品名”匹配销售排单“客户机种”。",
    "王永仁": "读取预测文件中的“子件描述”前缀匹配销售排单“客户机种”。",
    "叶振华": "读取预测文件中的“模组型号”或“料号”匹配销售排单“客户机种”。",
    "李海鹰": "读取预测文件中的“料号”匹配销售排单“客户机种”。",
}
OWNER_STATUS_LABELS = {
    "pending": "未上传",
    "running": "处理中",
    "done": "已完成",
    "error": "失败",
}


@lru_cache(maxsize=1)
def get_lan_ip() -> str:
    for iface in ("en0", "en1", "en2"):
        try:
            output = subprocess.check_output(["ifconfig", iface], text=True)
            for line in output.splitlines():
                line = line.strip()
                if line.startswith("inet ") and not line.startswith("inet 127."):
                    parts = line.split()
                    if len(parts) >= 2:
                        return parts[1]
        except Exception:
            pass

    backup_path: Optional[Path] = None
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            ip = sock.getsockname()[0]
            if ip and not ip.startswith("127."):
                return ip
    except OSError:
        pass

    return "127.0.0.1"


def normalize_base_url(value: str) -> str:
    value = value.strip()
    if not value:
        return ""
    if "://" not in value:
        value = f"https://{value.lstrip('/')}"
    return value.rstrip("/")


def get_public_share_url(handler: Optional[BaseHTTPRequestHandler] = None) -> Optional[str]:
    for env_name in PUBLIC_URL_ENV_VARS:
        raw_value = os.environ.get(env_name)
        if raw_value:
            return normalize_base_url(raw_value)

    if handler is not None:
        forwarded_host = handler.headers.get("X-Forwarded-Host")
        if forwarded_host:
            forwarded_proto = handler.headers.get("X-Forwarded-Proto", "https")
            host = forwarded_host.split(",")[0].strip()
            proto = forwarded_proto.split(",")[0].strip() or "https"
            if host:
                return normalize_base_url(f"{proto}://{host}")

        host_header = handler.headers.get("Host")
        if host_header:
            host = host_header.split(",")[0].strip()
            host_lc = host.lower()
            host_name = host_lc.split(":", 1)[0]
            is_local_host = host_lc.startswith("[::1]") or host_name in {
                "127.0.0.1",
                "localhost",
                "0.0.0.0",
                "::1",
            }
            if not is_local_host:
                scheme = handler.headers.get("X-Forwarded-Proto", "https").split(",")[0].strip() or "https"
                return normalize_base_url(f"{scheme}://{host}")

    return None


def uploaded_suffix(filename: str) -> str:
    return Path(str(filename).replace("\\", "/")).suffix.lower()


def content_disposition(filename: str) -> str:
    quoted = urllib.parse.quote(filename)
    return f'attachment; filename="result.xlsx"; filename*=UTF-8\'\'{quoted}'


def safe_owner(value: Optional[str]) -> str:
    value = urllib.parse.unquote(value or "").strip()
    return value if value in BUSINESS_OWNERS else DEFAULT_OWNER


def owner_link(owner: str) -> str:
    return f"/owner/{urllib.parse.quote(owner)}"


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


def remote_storage_enabled() -> bool:
    return bool(REMOTE_STORAGE_ENDPOINT and REMOTE_STORAGE_SECRET)


def remote_storage_url(key: str) -> str:
    return f"{REMOTE_STORAGE_ENDPOINT}?key={urllib.parse.quote(key)}"


def remote_request(
    method: str,
    key: str,
    data: bytes | None = None,
    content_type: str = "application/octet-stream",
) -> Optional[bytes]:
    if not remote_storage_enabled():
        return None

    headers = {"x-sales-storage-secret": REMOTE_STORAGE_SECRET}
    if data is not None:
        headers["Content-Type"] = content_type
    request = urllib.request.Request(
        remote_storage_url(key),
        data=data,
        headers=headers,
        method=method,
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return response.read()
    except urllib.error.HTTPError as exc:
        if exc.code == 404 and method == "GET":
            return None
        error_body = exc.read().decode("utf-8", errors="replace")[:500]
        raise RuntimeError(f"Netlify Blobs 请求失败：HTTP {exc.code} {error_body}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"连接 Netlify Blobs 失败：{exc.reason}") from exc


def remote_put_file(key: str, path: Path, content_type: str = XLSX_MIME) -> None:
    if not remote_storage_enabled() or not path.exists():
        return
    file_size = path.stat().st_size
    if file_size <= REMOTE_CHUNK_SIZE:
        remote_request("PUT", key, path.read_bytes(), content_type=content_type)
        return

    chunk_keys = []
    with path.open("rb") as source:
        index = 0
        while True:
            chunk = source.read(REMOTE_CHUNK_SIZE)
            if not chunk:
                break
            chunk_key = f"chunks/{key}/{index:05d}"
            remote_request("PUT", chunk_key, chunk, content_type="application/octet-stream")
            chunk_keys.append(chunk_key)
            index += 1

    remote_put_json(
        f"manifests/{key}.json",
        {
            "key": key,
            "content_type": content_type,
            "size": file_size,
            "chunk_size": REMOTE_CHUNK_SIZE,
            "chunks": chunk_keys,
            "updated_at": now_label(),
        },
    )


def remote_get_file(key: str, path: Path) -> bool:
    if not remote_storage_enabled():
        return False

    ensure_data_dir()
    temp_path = path.with_name(f".{path.name}.download")
    expected_size: Optional[int] = None
    manifest_data = remote_request("GET", f"manifests/{key}.json")
    if manifest_data is not None:
        manifest = json.loads(manifest_data.decode("utf-8"))
        chunk_keys = manifest.get("chunks") or []
        if not chunk_keys:
            return False
        expected_size = int(manifest.get("size") or 0) or None
        with temp_path.open("wb") as out:
            for chunk_key in chunk_keys:
                chunk = remote_request("GET", str(chunk_key))
                if chunk is None:
                    temp_path.unlink(missing_ok=True)
                    return False
                out.write(chunk)
        if expected_size is not None and temp_path.stat().st_size != expected_size:
            temp_path.unlink(missing_ok=True)
            raise RuntimeError("Netlify Blobs 里的 Excel 分片不完整，请重新上传共用销售排单。")
        if key.lower().endswith((".xlsx", ".xlsm")) and not zipfile.is_zipfile(temp_path):
            temp_path.unlink(missing_ok=True)
            raise RuntimeError("Netlify Blobs 里的 Excel 文件不完整，请重新上传共用销售排单。")
        temp_path.replace(path)
        return True

    data = remote_request("GET", key)
    if data is None:
        return False
    temp_path.write_bytes(data)
    if key.lower().endswith((".xlsx", ".xlsm")) and not zipfile.is_zipfile(temp_path):
        temp_path.unlink(missing_ok=True)
        raise RuntimeError("Netlify Blobs 里的 Excel 文件不完整，请重新上传共用销售排单。")
    temp_path.replace(path)
    return True


def remote_put_json(key: str, payload: dict) -> None:
    if not remote_storage_enabled():
        return
    remote_request(
        "PUT",
        key,
        json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
        content_type="application/json; charset=utf-8",
    )


def hydrate_remote_state() -> None:
    if not remote_storage_enabled():
        return
    ensure_data_dir()
    if not METADATA_PATH.exists():
        remote_get_file(REMOTE_METADATA_KEY, METADATA_PATH)
    if not MASTER_SALES_PATH.exists():
        remote_get_file(REMOTE_MASTER_KEY, MASTER_SALES_PATH)
    if MASTER_SALES_PATH.exists() and not LATEST_OUTPUT_PATH.exists():
        if not remote_get_file(REMOTE_LATEST_KEY, LATEST_OUTPUT_PATH):
            shutil.copy2(MASTER_SALES_PATH, LATEST_OUTPUT_PATH)


def sync_master_files_to_remote() -> None:
    if not remote_storage_enabled():
        return
    remote_put_file(REMOTE_MASTER_KEY, MASTER_SALES_PATH)
    remote_put_file(REMOTE_LATEST_KEY, LATEST_OUTPUT_PATH)


def load_metadata() -> dict:
    if remote_storage_enabled() and not METADATA_PATH.exists():
        hydrate_remote_state()
    if not METADATA_PATH.exists():
        return {}
    try:
        return json.loads(METADATA_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def default_owner_statuses() -> dict[str, dict[str, str]]:
    return {
        owner: {
            "state": "pending",
            "updated_rows": "",
            "updated_at": "",
            "uploaded_at": "",
            "prediction_name": "",
            "error": "",
        }
        for owner in BUSINESS_OWNERS
    }


def ensure_metadata_schema(metadata: dict) -> dict:
    statuses = metadata.get("owner_statuses")
    if not isinstance(statuses, dict):
        statuses = {}

    normalized = default_owner_statuses()
    for owner in BUSINESS_OWNERS:
        raw = statuses.get(owner)
        if isinstance(raw, dict):
            normalized[owner].update({key: str(value) for key, value in raw.items() if value is not None})
            if normalized[owner].get("state") not in OWNER_STATUS_LABELS:
                normalized[owner]["state"] = "pending"

    metadata["owner_statuses"] = normalized
    return metadata


def save_metadata(metadata: dict) -> None:
    ensure_data_dir()
    metadata = ensure_metadata_schema(metadata)
    METADATA_PATH.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    remote_put_json(REMOTE_METADATA_KEY, metadata)


def update_owner_status(owner: str, **values) -> None:
    with STATE_LOCK:
        metadata = ensure_metadata_schema(load_metadata())
        status = metadata["owner_statuses"].setdefault(owner, default_owner_statuses().get(owner, {}))
        for key, value in values.items():
            status[key] = "" if value is None else str(value)
        metadata["owner_statuses"][owner] = status
        save_metadata(metadata)


def get_job_status() -> dict:
    with JOB_LOCK:
        return dict(JOB_STATUS)


def set_job_status(**values) -> None:
    with JOB_LOCK:
        JOB_STATUS.update(values)


def job_is_running() -> bool:
    return get_job_status().get("state") == "running"


def now_label() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_filename(value: str, default: str = "销售排单.xlsx") -> str:
    name = os.path.basename(str(value).replace("\\", "/")).strip()
    return name or default


def master_sales_exists() -> bool:
    if not MASTER_SALES_PATH.exists() and remote_storage_enabled():
        hydrate_remote_state()
    if not MASTER_SALES_PATH.exists() or MASTER_SALES_PATH.stat().st_size <= 0:
        return False
    return zipfile.is_zipfile(MASTER_SALES_PATH)


def storage_description() -> str:
    if remote_storage_enabled():
        return "当前已启用 Netlify Blobs 远程保存；Render 重启后会自动恢复本周排单和业务状态。"
    if str(DATA_DIR).startswith("/data"):
        return "当前使用 Render 持久磁盘目录 /data，网站重启后仍会保留本周数据。"
    return "当前使用临时目录；若 Render 服务重启或重新部署，本周数据可能丢失，请在 Render 增加持久磁盘并挂载到 /data。"


def format_file_size(path: Path) -> str:
    if not path.exists():
        return "0 KB"
    size = path.stat().st_size
    if size >= 1024 * 1024:
        return f"{size / 1024 / 1024:.1f} MB"
    return f"{max(size / 1024, 1):.0f} KB"


def selected_owner_from_path(path: str, query: dict[str, list[str]]) -> str:
    if path.startswith("/owner/"):
        return safe_owner(path.removeprefix("/owner/"))
    if "owner" in query and query["owner"]:
        return safe_owner(query["owner"][0])
    return DEFAULT_OWNER


def render_page(
    message: str = "",
    error: str = "",
    handler: Optional[BaseHTTPRequestHandler] = None,
    selected_owner: Optional[str] = None,
) -> bytes:
    selected_owner = safe_owner(selected_owner)
    selected_owner_guide = OWNER_GUIDES.get(selected_owner, "按预测文件中的机种信息匹配销售排单“客户机种”。")
    if selected_owner == "李玎玲":
        upload_hint = "李玎玲页面建议上传完整截图，保留左侧“机种名”列和右侧 6-10 月数量列，系统会按表格结构去识别。"
    else:
        upload_hint = "预测信息优先使用表格文件，也可以上传清晰截图。"
    local_url = f"http://127.0.0.1:{PORT}/"
    public_url = get_public_share_url(handler)
    lan_ip = get_lan_ip() if public_url is None else ""
    lan_url = f"http://{lan_ip}:{PORT}/" if lan_ip and lan_ip != "127.0.0.1" else ""

    if public_url:
        access_note = f"公网地址：{public_url}/"
    elif lan_url:
        access_note = f"本机：{local_url} ｜ 同一 Wi-Fi：{lan_url}"
    else:
        access_note = f"本机：{local_url}"

    owner_nav = "\n".join(
        f'<a class="owner-link {"active" if owner == selected_owner else ""}" href="{owner_link(owner)}">{html.escape(owner)}</a>'
        for owner in BUSINESS_OWNERS
    )
    owner_options = "\n".join(
        f'<option value="{html.escape(owner)}" {"selected" if owner == selected_owner else ""}>{html.escape(owner)}</option>'
        for owner in BUSINESS_OWNERS
    )
    message_html = (
        f'<div class="notice success" role="status">{html.escape(message)}</div>' if message else ""
    )
    error_html = f'<div class="notice error" role="alert">{html.escape(error)}</div>' if error else ""
    metadata = ensure_metadata_schema(load_metadata())
    has_master = master_sales_exists()
    master_name = metadata.get("master_name", "未上传共用销售排单")
    master_time = metadata.get("master_uploaded_at", "")
    latest_owner = metadata.get("last_owner", "")
    latest_time = metadata.get("last_generated_at", "")
    latest_rows = metadata.get("last_updated_rows", "")
    latest_file = metadata.get("latest_name", metadata.get("master_name", "最新销售排单.xlsx"))
    master_status = "已启用" if has_master else "待上传"
    master_status_class = "ready" if has_master else "empty"
    download_latest_html = (
        '<a class="secondary button" href="/download/latest">下载当前最新版</a>' if has_master else ""
    )
    preview_latest_html = (
        '<a class="secondary button" href="/preview">在线预览/编辑</a>' if has_master else ""
    )
    state_text = (
        f"当前共用排单：{master_name}（{format_file_size(MASTER_SALES_PATH)}）"
        if has_master
        else "请先上传本周共用销售排单，然后各业务只上传自己的预测。"
    )
    detail_parts = []
    if master_time:
        detail_parts.append(f"初始化/替换时间：{master_time}")
    if latest_owner:
        detail_parts.append(f"最近回填：{latest_owner}，更新 {latest_rows} 行，{latest_time}")
    if latest_file and has_master:
        detail_parts.append(f"下载文件名：{latest_file}")
    detail_parts.append(storage_description())
    detail_text = " ｜ ".join(detail_parts) if detail_parts else "共用排单会在每位业务上传预测后自动保存为最新版。"
    owner_statuses = metadata.get("owner_statuses", default_owner_statuses())
    owner_status_card_items = []
    for owner in BUSINESS_OWNERS:
        status = owner_statuses.get(owner, {})
        state = status.get("state", "pending")
        row_text = f'{html.escape(status.get("updated_rows", ""))} 行' if status.get("updated_rows") else "等待预测"
        prediction_name = status.get("prediction_name") or "暂无上传记录"
        uploaded_at = status.get("uploaded_at") or status.get("updated_at") or ""
        done_at = status.get("updated_at") or ""
        error_text = status.get("error") or ""
        owner_status_card_items.append(
            f'<div class="owner-status {html.escape(state)}">'
            f'<strong>{html.escape(owner)}</strong>'
            f'<span>{html.escape(OWNER_STATUS_LABELS.get(state, "未上传"))}</span>'
            f'<small>最近文件：{html.escape(prediction_name)}</small>'
            f'<small>上传时间：{html.escape(uploaded_at or "未上传")}</small>'
            f'<small>回填结果：{row_text}{("，" + html.escape(done_at)) if done_at else ""}</small>'
            f'{f"<small>错误：{html.escape(error_text)}</small>" if error_text else ""}'
            f'</div>'
        )
    owner_status_cards = "\n".join(owner_status_card_items)
    job = get_job_status()
    job_state = job.get("state", "idle")
    refresh_url = owner_link(selected_owner)
    if job_state == "running":
        job_html = (
            f'<div class="notice working" role="status">正在处理：{html.escape(job.get("owner", ""))} 的预测，'
            f'开始时间：{html.escape(job.get("started_at", ""))}。处理完成前请不要重复上传。</div>'
        )
        refresh_meta = f'<meta http-equiv="refresh" content="15; url={html.escape(refresh_url)}">'
    elif job_state == "done":
        job_html = (
            f'<div class="notice success" role="status">最近处理完成：{html.escape(job.get("owner", ""))}，'
            f'更新 {html.escape(str(job.get("updated_rows", "")))} 行，'
            f'{html.escape(job.get("finished_at", ""))}。可以下载当前最新版。</div>'
        )
        refresh_meta = ""
    elif job_state == "error":
        job_html = (
            f'<div class="notice error" role="alert">最近处理失败：{html.escape(job.get("owner", ""))}，'
            f'{html.escape(job.get("error", ""))}</div>'
        )
        refresh_meta = ""
    else:
        job_html = ""
        refresh_meta = ""

    page = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  {refresh_meta}
  <title>销售排单回填工作台</title>
  <style>
    :root {{
      color-scheme: light;
      --paper: #f6f7f4;
      --panel: #ffffff;
      --line: #d8ded5;
      --text: #17211b;
      --muted: #5f6b63;
      --accent: #136f63;
      --accent-dark: #0c4f47;
      --warn: #8a5800;
      --danger: #b42318;
      --success: #126b45;
      --shadow: 0 16px 34px rgba(28, 43, 34, 0.08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      min-height: 100vh;
      color: var(--text);
      font-family: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", -apple-system, BlinkMacSystemFont, sans-serif;
      background:
        linear-gradient(90deg, rgba(19, 111, 99, 0.06) 1px, transparent 1px),
        linear-gradient(180deg, rgba(138, 88, 0, 0.05) 1px, transparent 1px),
        var(--paper);
      background-size: 28px 28px;
    }}
    a {{ color: inherit; }}
    .shell {{
      width: min(1180px, calc(100% - 28px));
      margin: 0 auto;
      padding: 22px 0 40px;
    }}
    .topline {{
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: flex-start;
      padding: 14px 0 18px;
      border-bottom: 1px solid var(--line);
    }}
    h1 {{
      margin: 0 0 6px;
      font-size: 26px;
      line-height: 1.25;
      letter-spacing: 0;
    }}
    .subtext {{
      margin: 0;
      color: var(--muted);
      line-height: 1.65;
      font-size: 14px;
    }}
    .access {{
      max-width: 430px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.6;
      text-align: right;
    }}
    .layout {{
      display: grid;
      grid-template-columns: 220px minmax(0, 1fr);
      gap: 18px;
      margin-top: 18px;
      align-items: start;
    }}
    .rail,
    .panel,
    .summary-item {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }}
    .rail {{
      padding: 12px;
      position: sticky;
      top: 14px;
    }}
    .rail-title {{
      padding: 4px 6px 10px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 700;
    }}
    .owner-list {{
      display: grid;
      gap: 6px;
    }}
    .owner-link {{
      display: flex;
      align-items: center;
      min-height: 38px;
      padding: 9px 10px;
      border-radius: 6px;
      text-decoration: none;
      color: var(--text);
      border: 1px solid transparent;
      font-weight: 700;
    }}
    .owner-link.active {{
      color: var(--accent-dark);
      background: #e9f3ee;
      border-color: #bcd8cc;
    }}
    .panel {{
      padding: 20px;
    }}
    .panel-head {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: flex-start;
      margin-bottom: 16px;
    }}
    h2 {{
      margin: 0 0 6px;
      font-size: 20px;
      line-height: 1.3;
      letter-spacing: 0;
    }}
    .badge {{
      display: inline-flex;
      align-items: center;
      min-height: 32px;
      padding: 6px 10px;
      border-radius: 6px;
      background: #f3ead8;
      color: var(--warn);
      font-weight: 800;
      white-space: nowrap;
    }}
    .badge.ready {{
      background: #e8f5ee;
      color: var(--success);
    }}
    .badge.empty {{
      background: #fff0ee;
      color: var(--danger);
    }}
    .notice {{
      margin: 0 0 14px;
      padding: 12px 14px;
      border-radius: 8px;
      border: 1px solid transparent;
      line-height: 1.6;
      font-weight: 700;
    }}
    .notice.success {{
      background: #e8f5ee;
      border-color: #bbdec9;
      color: var(--success);
    }}
    .notice.working {{
      background: #fff8e6;
      border-color: #efd596;
      color: var(--warn);
    }}
    .notice.error {{
      background: #fff0ee;
      border-color: #f3c3bd;
      color: var(--danger);
    }}
    form {{
      display: grid;
      gap: 16px;
    }}
    .field-grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }}
    label {{
      display: block;
      margin-bottom: 7px;
      font-weight: 800;
    }}
    select,
    input[type="file"] {{
      width: 100%;
      min-height: 44px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: #fff;
      padding: 10px 12px;
      font: inherit;
      color: var(--text);
    }}
    .hint {{
      margin-top: 7px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.55;
    }}
    .rule-panel {{
      display: grid;
      gap: 8px;
      padding: 12px 14px;
      border: 1px solid #cfe0d7;
      border-radius: 8px;
      background: #f4faf7;
      color: var(--accent-dark);
      line-height: 1.55;
      font-size: 14px;
      font-weight: 700;
    }}
    .rule-panel span {{
      color: var(--muted);
      font-size: 13px;
      font-weight: 600;
    }}
    .master-panel {{
      margin-top: 18px;
      padding: 18px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }}
    .master-card {{
      display: grid;
      gap: 10px;
      padding: 14px;
      border: 1px solid #cfe0d7;
      border-radius: 8px;
      background: #f8fbf8;
    }}
    .master-top {{
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 12px;
    }}
    .master-title {{
      margin: 0;
      font-size: 17px;
      font-weight: 900;
    }}
    .master-actions {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: end;
      margin-top: 12px;
    }}
    .master-actions .upload-field {{
      flex: 1 1 320px;
    }}
    .status-board {{
      display: grid;
      grid-template-columns: repeat(6, minmax(0, 1fr));
      gap: 10px;
      margin-top: 12px;
    }}
    .owner-status {{
      display: grid;
      gap: 4px;
      min-height: 92px;
      padding: 12px;
      border-radius: 8px;
      border: 1px solid var(--line);
      background: #fff;
    }}
    .owner-status strong {{
      font-size: 15px;
    }}
    .owner-status span {{
      display: inline-flex;
      width: fit-content;
      padding: 3px 7px;
      border-radius: 999px;
      font-size: 12px;
      font-weight: 900;
      color: var(--muted);
      background: #eef1ec;
    }}
    .owner-status small {{
      color: var(--muted);
      line-height: 1.35;
    }}
    .owner-status.running span {{
      color: var(--warn);
      background: #fff4d7;
    }}
    .owner-status.done span {{
      color: var(--success);
      background: #e8f5ee;
    }}
    .owner-status.error span {{
      color: var(--danger);
      background: #fff0ee;
    }}
    .actions {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
      padding-top: 4px;
    }}
    button,
    .button {{
      min-height: 44px;
      border: 0;
      border-radius: 7px;
      padding: 11px 16px;
      font: inherit;
      font-weight: 800;
      cursor: pointer;
    }}
    .primary {{
      color: #fff;
      background: var(--accent);
    }}
    .secondary {{
      color: var(--accent-dark);
      background: #e9f3ee;
      border: 1px solid #bdd9cd;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
    }}
    .summary {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-top: 14px;
    }}
    .summary-item {{
      padding: 14px;
      box-shadow: none;
    }}
    .summary-item strong {{
      display: block;
      margin-bottom: 5px;
      font-size: 15px;
    }}
    .summary-item span {{
      color: var(--muted);
      line-height: 1.55;
      font-size: 13px;
    }}
    @media (max-width: 820px) {{
      .topline,
      .panel-head {{
        display: block;
      }}
      .access {{
        margin-top: 10px;
        text-align: left;
      }}
      .layout,
      .field-grid,
      .summary,
      .status-board,
      .master-top {{
        grid-template-columns: 1fr;
        display: block;
      }}
      .rail {{
        position: static;
      }}
      .owner-list {{
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }}
    }}
  </style>
</head>
<body>
  <main class="shell">
    <section class="topline">
      <div>
        <h1>销售排单回填工作台</h1>
        <p class="subtext">网站维护一份共用销售排单。每周先上传/替换本周排单，随后每位业务只上传自己的预测，系统会按机种分别补充 6/29预估（7月）、6/29预估（8月）等每个月的预估数量和金额。</p>
      </div>
      <div class="access">{html.escape(access_note)}</div>
    </section>

    <section class="master-panel" aria-label="共用销售排单">
      <div class="master-card">
        <div class="master-top">
          <div>
            <p class="master-title">共用销售排单</p>
            <p class="subtext">{html.escape(state_text)}</p>
            <p class="hint">{html.escape(detail_text)}</p>
          </div>
          <div class="badge {master_status_class}">{html.escape(master_status)}</div>
        </div>
        <form class="master-actions" method="post" action="/sales-master" enctype="multipart/form-data">
          <div class="upload-field">
            <label for="master_sales_file">上传/替换本周共用销售排单</label>
            <input id="master_sales_file" name="sales_file" type="file" accept=".xlsx,.xlsm" required>
            <div class="hint">只有共用排单需要上传一次；各业务后续只上传预测。替换排单会从新文件重新开始。</div>
          </div>
          <button class="primary" type="submit">保存共用排单</button>
          {download_latest_html}
          {preview_latest_html}
        </form>
        <div class="status-board" aria-label="业务回填状态">
          {owner_status_cards}
        </div>
      </div>
    </section>

    <section class="layout">
      <aside class="rail" aria-label="业务担当">
        <div class="rail-title">业务担当页面</div>
        <nav class="owner-list">
          {owner_nav}
        </nav>
      </aside>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>{html.escape(selected_owner)} 的预测上传</h2>
            <p class="subtext">{html.escape(upload_hint)} 预测会写入上方共用销售排单，销售排单必须保留“业务担当”和“客户机种”。</p>
          </div>
          <div class="badge">无需登录</div>
        </div>
        {message_html}
        {error_html}
        {job_html}
        <form method="post" action="/generate" enctype="multipart/form-data">
          <div class="rule-panel">
            {html.escape(selected_owner_guide)}
            <span>当月会自动扣减同一客户机种已完成数量；后续月份按对应月份预估栏回填。每次回填都会保存为共用排单最新版。</span>
          </div>
          <div>
            <label for="business_owner">业务担当</label>
            <select id="business_owner" name="business_owner">
              {owner_options}
            </select>
            <div class="hint">只会回填销售排单中属于所选业务担当的行。</div>
          </div>
          <div>
            <label for="prediction_file">预测信息</label>
            <input id="prediction_file" name="prediction_file" type="file" accept=".xlsx,.jpg,.jpeg,.png" required>
            <div class="hint">支持 .xlsx、.jpg、.png；建议优先上传 .xlsx，清晰截图会按已适配规则识别。无需再上传销售排单。</div>
          </div>
          <div class="actions">
            <button class="primary" type="submit">开始回填</button>
            <a class="secondary button" href="{owner_link(selected_owner)}">刷新处理状态</a>
            {download_latest_html}
            {preview_latest_html}
          </div>
        </form>
      </section>
    </section>

    <section class="summary" aria-label="处理规则">
      <div class="summary-item">
        <strong>分人匹配</strong>
        <span>按业务担当读取 New part NO、机种名、品名、子件描述前缀、模组型号或料号。</span>
      </div>
      <div class="summary-item">
        <strong>担当隔离</strong>
        <span>每次只回填所选业务担当的销售排单行，便于多人分开上传。</span>
      </div>
      <div class="summary-item">
        <strong>扣已完成</strong>
        <span>当月客户总预测会扣掉销售排单中已完成数量，再回填剩余预测。</span>
      </div>
    </section>
  </main>
</body>
</html>
"""
    return page.encode("utf-8")


def write_html(handler: BaseHTTPRequestHandler, status: int, body: str, head: bool = False) -> None:
    encoded = body.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(encoded)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    if not head:
        handler.wfile.write(encoded)


def write_text(
    handler: BaseHTTPRequestHandler,
    status: int,
    body: str,
    content_type: str = "text/plain; charset=utf-8",
    head: bool = False,
) -> None:
    encoded = body.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", content_type)
    handler.send_header("Content-Length", str(len(encoded)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    if not head:
        handler.wfile.write(encoded)


def write_json(handler: BaseHTTPRequestHandler, status: int, payload: dict, head: bool = False) -> None:
    encoded = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(encoded)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    if not head:
        handler.wfile.write(encoded)


def save_upload(field: cgi.FieldStorage, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("wb") as out:
        shutil.copyfileobj(field.file, out)


def validate_excel_file(path: Path, label: str) -> None:
    suffix = path.suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        return
    if not path.exists() or path.stat().st_size <= 0:
        raise ValueError(f"{label}上传后为空文件，请重新选择原始 Excel 文件上传。")
    if not zipfile.is_zipfile(path):
        with path.open("rb") as f:
            head = f.read(16).hex(" ")
        raise ValueError(
            f"{label}不是标准 Excel 文件，网站无法读取。"
            f"请重新上传原始 .xlsx/.xlsm 文件（当前大小 {format_file_size(path)}，文件头 {head or '空'}）。"
        )
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        wb.close()
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"{label}无法被 Excel 解析：{exc}") from exc


def validate_prediction_file(path: Path, label: str) -> None:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        validate_excel_file(path, label)
        return

    if suffix in {".jpg", ".jpeg", ".png"}:
        if not path.exists() or path.stat().st_size <= 0:
            raise ValueError(f"{label}上传后为空文件，请重新选择原始图片上传。")
        with path.open("rb") as f:
            head = f.read(12)
        is_png = suffix == ".png" and head.startswith(b"\x89PNG\r\n\x1a\n")
        is_jpg = suffix in {".jpg", ".jpeg"} and head.startswith(b"\xff\xd8\xff")
        if not (is_png or is_jpg):
            raise ValueError(f"{label}不是标准 {suffix} 图片，请重新上传原始 jpg/png 截图。")
        return

    raise ValueError("预测信息请上传 .xlsx、.jpg 或 .png 文件。")


def first_upload(form: cgi.FieldStorage, name: str):
    field = form[name]
    if isinstance(field, list):
        return field[0]
    return field


def send_xlsx(handler: BaseHTTPRequestHandler, path: Path, download_name: str, head: bool = False) -> None:
    file_size = path.stat().st_size
    handler.send_response(200)
    handler.send_header("Content-Type", XLSX_MIME)
    handler.send_header("Content-Disposition", content_disposition(download_name))
    handler.send_header("Content-Length", str(file_size))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    if not head:
        with path.open("rb") as f:
            shutil.copyfileobj(f, handler.wfile)


def parse_cell_edit_value(raw_value: str):
    value = raw_value.strip()
    if value == "":
        return None
    if value.startswith("="):
        return value
    if re.fullmatch(r"-?\d+", value) and not re.match(r"-?0\d+", value):
        return int(value)
    if re.fullmatch(r"-?(?:\d+\.\d+|\d+\.|\.\d+)", value):
        return float(value)
    return raw_value


def render_preview_page(
    message: str = "",
    error: str = "",
    selected_sheet: str = "",
    rows_limit: int = 80,
) -> bytes:
    if not master_sales_exists():
        return render_page(error="还没有共用销售排单可预览，请先上传本周排单。")

    metadata = ensure_metadata_schema(load_metadata())
    message_html = f'<div class="notice success">{html.escape(message)}</div>' if message else ""
    error_html = f'<div class="notice error">{html.escape(error)}</div>' if error else ""
    rows_limit = max(20, min(rows_limit, 200))

    wb = load_workbook(MASTER_SALES_PATH, read_only=True, data_only=False)
    try:
        sheet_names = wb.sheetnames
        if selected_sheet not in sheet_names:
            selected_sheet = sheet_names[0]
        ws = wb[selected_sheet]
        max_row = min(ws.max_row, rows_limit)
        max_col = min(ws.max_column, 80)
        sheet_options = "\n".join(
            f'<option value="{html.escape(name)}" {"selected" if name == selected_sheet else ""}>{html.escape(name)}</option>'
            for name in sheet_names
        )
        header_cells = "".join(f"<th>{get_column_letter(col)}</th>" for col in range(1, max_col + 1))
        body_rows = []
        for row_idx in range(1, max_row + 1):
            cells = [f"<th>{row_idx}</th>"]
            for col_idx in range(1, max_col + 1):
                value = ws.cell(row_idx, col_idx).value
                display = "" if value is None else str(value)
                if len(display) > 80:
                    display = display[:77] + "..."
                coord = f"{get_column_letter(col_idx)}{row_idx}"
                cells.append(f'<td title="{html.escape(coord)}">{html.escape(display)}</td>')
            body_rows.append("<tr>" + "".join(cells) + "</tr>")
        table_html = "\n".join(body_rows)
    finally:
        wb.close()

    latest_name = metadata.get("latest_name") or metadata.get("master_name") or "当前最新版"
    page = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>在线预览/编辑</title>
  <style>
    :root {{
      --paper: #f6f7f4;
      --panel: #ffffff;
      --line: #d8ded5;
      --text: #17211b;
      --muted: #5f6b63;
      --accent: #136f63;
      --danger: #b42318;
      --success: #126b45;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      color: var(--text);
      font-family: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", -apple-system, BlinkMacSystemFont, sans-serif;
      background: var(--paper);
    }}
    .shell {{ width: min(1380px, calc(100% - 28px)); margin: 0 auto; padding: 22px 0 40px; }}
    .top {{ display: flex; justify-content: space-between; gap: 12px; align-items: flex-start; margin-bottom: 16px; }}
    h1 {{ margin: 0 0 6px; font-size: 24px; }}
    .subtext {{ margin: 0; color: var(--muted); line-height: 1.6; }}
    .panel {{ background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 16px; margin-bottom: 16px; }}
    .notice {{ margin: 0 0 14px; padding: 12px 14px; border-radius: 8px; line-height: 1.6; font-weight: 800; }}
    .notice.success {{ color: var(--success); background: #e8f5ee; border: 1px solid #bbdec9; }}
    .notice.error {{ color: var(--danger); background: #fff0ee; border: 1px solid #f3c3bd; }}
    form {{ display: flex; flex-wrap: wrap; gap: 10px; align-items: end; }}
    label {{ display: grid; gap: 6px; font-weight: 800; }}
    input, select {{ min-height: 40px; border: 1px solid var(--line); border-radius: 7px; padding: 8px 10px; font: inherit; }}
    button, .button {{ min-height: 40px; border: 0; border-radius: 7px; padding: 9px 14px; font: inherit; font-weight: 800; text-decoration: none; cursor: pointer; }}
    .primary {{ color: #fff; background: var(--accent); }}
    .secondary {{ color: #0c4f47; background: #e9f3ee; border: 1px solid #bdd9cd; display: inline-flex; align-items: center; }}
    .table-wrap {{ overflow: auto; max-height: 72vh; border: 1px solid var(--line); border-radius: 8px; background: #fff; }}
    table {{ border-collapse: collapse; min-width: 100%; font-size: 12px; }}
    th, td {{ border: 1px solid #d7ddd4; padding: 5px 7px; white-space: nowrap; max-width: 220px; overflow: hidden; text-overflow: ellipsis; }}
    th {{ position: sticky; top: 0; background: #e8efe9; z-index: 1; }}
    tr th:first-child {{ position: sticky; left: 0; z-index: 2; background: #e8efe9; }}
  </style>
</head>
<body>
  <main class="shell">
    <div class="top">
      <div>
        <h1>在线预览/编辑当前最新版</h1>
        <p class="subtext">当前文件：{html.escape(latest_name)}。预览最多显示前 {rows_limit} 行、前 80 列；修改会直接保存到网站当前共用表格。</p>
      </div>
      <a class="secondary button" href="/">返回上传页</a>
    </div>
    {message_html}
    {error_html}
    <section class="panel">
      <form method="get" action="/preview">
        <label>选择 Sheet
          <select name="sheet">{sheet_options}</select>
        </label>
        <label>预览行数
          <input name="rows" value="{rows_limit}" inputmode="numeric">
        </label>
        <button class="primary" type="submit">刷新预览</button>
        <a class="secondary button" href="/download/latest">下载当前最新版</a>
      </form>
    </section>
    <section class="panel">
      <form method="post" action="/edit-cell">
        <input type="hidden" name="sheet" value="{html.escape(selected_sheet)}">
        <label>单元格
          <input name="cell" placeholder="例如 BU25" required>
        </label>
        <label>新内容
          <input name="value" placeholder="留空则清空该单元格">
        </label>
        <button class="primary" type="submit">保存修改</button>
        <span class="subtext">提示：点击表格单元格时可在浏览器提示里看到坐标。</span>
      </form>
    </section>
    <div class="table-wrap">
      <table>
        <thead><tr><th>#</th>{header_cells}</tr></thead>
        <tbody>{table_html}</tbody>
      </table>
    </div>
  </main>
</body>
</html>
"""
    return page.encode("utf-8")


def handle_edit_cell(handler: BaseHTTPRequestHandler, head: bool = False) -> None:
    try:
        if job_is_running():
            raise ValueError("当前有预测正在后台回填，请处理完成后再编辑表格。")
        if not master_sales_exists():
            raise ValueError("还没有共用销售排单可编辑，请先上传本周排单。")

        content_length = int(handler.headers.get("Content-Length", "0") or "0")
        body = handler.rfile.read(content_length).decode("utf-8")
        form = urllib.parse.parse_qs(body, keep_blank_values=True)
        sheet = (form.get("sheet", [""])[0] or "").strip()
        cell = (form.get("cell", [""])[0] or "").strip().upper().replace(" ", "")
        raw_value = form.get("value", [""])[0]

        if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}", cell):
            raise ValueError("单元格格式不正确，请输入类似 BU25 的位置。")

        with STATE_LOCK:
            wb = load_workbook(MASTER_SALES_PATH)
            try:
                if sheet not in wb.sheetnames:
                    raise ValueError(f"没有找到 Sheet：{sheet}")
                ws = wb[sheet]
                coordinate_to_tuple(cell)
                ws[cell].value = parse_cell_edit_value(raw_value)
                wb.save(MASTER_SALES_PATH)
                shutil.copy2(MASTER_SALES_PATH, LATEST_OUTPUT_PATH)
                sync_master_files_to_remote()

                metadata = ensure_metadata_schema(load_metadata())
                metadata["last_manual_edit_at"] = now_label()
                metadata["last_manual_edit_cell"] = f"{sheet}!{cell}"
                save_metadata(metadata)
            finally:
                wb.close()

        write_html(
            handler,
            200,
            render_preview_page(message=f"已保存 {sheet}!{cell} 的修改。", selected_sheet=sheet).decode("utf-8"),
            head=head,
        )
    except Exception as exc:  # noqa: BLE001
        write_html(
            handler,
            400,
            render_preview_page(error=f"保存失败：{exc}").decode("utf-8"),
            head=head,
        )


def handle_sales_master(handler: BaseHTTPRequestHandler, head: bool = False) -> None:
    if handler.headers.get_content_type() != UPLOAD_MIME:
        write_html(
            handler,
            400,
            render_page(error="请使用网页表单上传共用销售排单。", handler=handler).decode("utf-8"),
            head=head,
        )
        return

    try:
        if job_is_running():
            raise ValueError("当前有预测正在后台回填，请处理完成后再替换共用销售排单。")

        form = cgi.FieldStorage(
            fp=handler.rfile,
            headers=handler.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": handler.headers.get("Content-Type", ""),
                "CONTENT_LENGTH": handler.headers.get("Content-Length", "0"),
            },
        )
        sales_field = first_upload(form, "sales_file")
        if not getattr(sales_field, "filename", None):
            raise ValueError("请先选择共用销售排单文件。")

        sales_suffix = uploaded_suffix(sales_field.filename)
        if sales_suffix not in SUPPORTED_SALES_SUFFIXES:
            raise ValueError("共用销售排单请上传 .xlsx 或 .xlsm 文件。")

        original_name = safe_filename(sales_field.filename)
        with STATE_LOCK:
            ensure_data_dir()
            if master_sales_exists():
                backup_name = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + MASTER_SALES_PATH.name
                backup_path = BACKUP_DIR / backup_name
                shutil.copy2(MASTER_SALES_PATH, backup_path)
            save_upload(sales_field, MASTER_SALES_PATH)
            validate_excel_file(MASTER_SALES_PATH, "共用销售排单")
            shutil.copy2(MASTER_SALES_PATH, LATEST_OUTPUT_PATH)
            sync_master_files_to_remote()
            save_metadata(
                {
                    "master_name": original_name,
                    "master_uploaded_at": now_label(),
                    "latest_name": f"{Path(original_name).stem}_当前最新版.xlsx",
                    "last_owner": "",
                    "last_generated_at": "",
                    "last_updated_rows": "",
                    "owner_statuses": default_owner_statuses(),
                }
            )

        write_html(
            handler,
            200,
            render_page(
                message=f"共用销售排单已保存：{original_name}。现在各业务只需要上传自己的预测。",
                handler=handler,
            ).decode("utf-8"),
            head=head,
        )
    except Exception as exc:  # noqa: BLE001
        if backup_path and backup_path.exists():
            try:
                shutil.copy2(backup_path, MASTER_SALES_PATH)
                shutil.copy2(backup_path, LATEST_OUTPUT_PATH)
            except Exception:
                pass
        write_html(
            handler,
            400,
            render_page(error=f"保存共用排单失败：{exc}", handler=handler).decode("utf-8"),
            head=head,
        )


def process_prediction_job(selected_owner: str, pred_path: Path) -> None:
    set_job_status(
        state="running",
        owner=selected_owner,
        message="正在回填共用销售排单",
        started_at=now_label(),
        finished_at="",
        updated_rows="",
        error="",
    )
    try:
        with tempfile.TemporaryDirectory(prefix="sales_job_") as tmp_dir:
            tmp_path = Path(tmp_dir)
            working_sales_path = tmp_path / "current_sales.xlsx"
            output_path = tmp_path / "generated.xlsx"

            with STATE_LOCK:
                if not master_sales_exists():
                    raise ValueError("共用销售排单不存在，请先重新上传本周排单。")
                ensure_data_dir()
                shutil.copy2(MASTER_SALES_PATH, working_sales_path)
                validate_excel_file(working_sales_path, "当前共用销售排单")
            validate_prediction_file(pred_path, f"{selected_owner} 的预测文件")

            summary = process_sales_workbooks(
                pred_path=pred_path,
                sales_path=working_sales_path,
                output_path=output_path,
                # The shared weekly schedule can contain hundreds of thousands
                # of formulas. Freezing them during a web upload makes the
                # request time out, so the website keeps formulas live.
                freeze_formulas=False,
                business_owner=selected_owner,
            )

            with STATE_LOCK:
                if not master_sales_exists():
                    raise ValueError("共用销售排单不存在，请先重新上传本周排单。")
                backup_name = datetime.now().strftime("%Y%m%d_%H%M%S") + f"_{selected_owner}_before.xlsx"
                shutil.copy2(MASTER_SALES_PATH, BACKUP_DIR / backup_name)
                shutil.copy2(output_path, MASTER_SALES_PATH)
                shutil.copy2(output_path, LATEST_OUTPUT_PATH)
                sync_master_files_to_remote()

                metadata = load_metadata()
                master_name = metadata.get("master_name", "销售排单.xlsx")
                stem = Path(master_name).stem if master_name else "销售排单"
                download_name = f"{stem}_{selected_owner}_已回填_当前最新版.xlsx"
                metadata.update(
                    {
                        "latest_name": download_name,
                        "last_owner": selected_owner,
                        "last_generated_at": now_label(),
                        "last_updated_rows": summary["updated_rows"],
                    }
                )
                metadata = ensure_metadata_schema(metadata)
                metadata["owner_statuses"][selected_owner].update(
                    {
                        "state": "done",
                        "updated_rows": str(summary["updated_rows"]),
                        "updated_at": now_label(),
                        "error": "",
                    }
                )
                save_metadata(metadata)

        set_job_status(
            state="done",
            owner=selected_owner,
            message="处理完成",
            finished_at=now_label(),
            updated_rows=summary["updated_rows"],
            error="",
        )
    except Exception as exc:  # noqa: BLE001
        set_job_status(
            state="error",
            owner=selected_owner,
            message="处理失败",
            finished_at=now_label(),
            updated_rows="",
            error=str(exc),
        )
        update_owner_status(
            selected_owner,
            state="error",
            updated_rows="",
            updated_at=now_label(),
            error=str(exc),
        )


def handle_generate(handler: BaseHTTPRequestHandler, head: bool = False) -> None:
    if handler.headers.get_content_type() != UPLOAD_MIME:
        write_html(
            handler,
            400,
            render_page(error="请使用网页表单上传预测信息。", handler=handler).decode("utf-8"),
            head=head,
        )
        return

    selected_owner = DEFAULT_OWNER
    try:
        form = cgi.FieldStorage(
            fp=handler.rfile,
            headers=handler.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": handler.headers.get("Content-Type", ""),
                "CONTENT_LENGTH": handler.headers.get("Content-Length", "0"),
            },
        )
        selected_owner = safe_owner(form.getfirst("business_owner", DEFAULT_OWNER))
        pred_field = first_upload(form, "prediction_file")
        if not getattr(pred_field, "filename", None):
            raise ValueError("请先选择预测信息文件。")

        pred_suffix = uploaded_suffix(pred_field.filename)
        if pred_suffix not in SUPPORTED_PREDICTION_SUFFIXES:
            raise ValueError("预测信息请上传 .xlsx、.jpg 或 .png 文件。")
        if not master_sales_exists():
            raise ValueError("请先在页面上方上传本周共用销售排单。")
        if job_is_running():
            raise ValueError("上一份预测还在处理，请稍后刷新页面，完成后再上传下一位业务预测。")

        ensure_data_dir()
        pred_name = datetime.now().strftime("%Y%m%d_%H%M%S") + f"_{selected_owner}{pred_suffix}"
        pred_path = UPLOAD_DIR / pred_name
        save_upload(pred_field, pred_path)
        validate_prediction_file(pred_path, f"{selected_owner} 的预测文件")
        update_owner_status(
            selected_owner,
            state="running",
            updated_rows="",
            uploaded_at=now_label(),
            updated_at="",
            prediction_name=safe_filename(pred_field.filename, "预测文件"),
            error="",
        )
        set_job_status(
            state="running",
            owner=selected_owner,
            message="正在回填共用销售排单",
            started_at=now_label(),
            finished_at="",
            updated_rows="",
            error="",
        )

        worker = threading.Thread(
            target=process_prediction_job,
            args=(selected_owner, pred_path),
            daemon=True,
        )
        worker.start()

        write_html(
            handler,
            202,
            render_page(
                message=f"{selected_owner} 的预测已上传，网站正在后台回填。请稍后刷新页面，完成后点击“下载当前最新版”。",
                handler=handler,
                selected_owner=selected_owner,
            ).decode("utf-8"),
            head=head,
        )
        return
    except Exception as exc:  # noqa: BLE001
        write_html(
            handler,
            400,
            render_page(error=f"生成失败：{exc}", handler=handler, selected_owner=selected_owner).decode("utf-8"),
            head=head,
        )


class SalesUploadHandler(BaseHTTPRequestHandler):
    server_version = "SalesUploadHTTP/4.0"

    def log_message(self, format: str, *args) -> None:  # noqa: A003
        print(f"{self.address_string()} - {format % args}")

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)
        if path in ("/", "/index.html") or path.startswith("/owner/"):
            owner = selected_owner_from_path(path, query)
            write_html(self, 200, render_page(handler=self, selected_owner=owner).decode("utf-8"))
            return

        if path in ("/generate", "/sales-master"):
            owner = selected_owner_from_path("/", query)
            write_html(self, 200, render_page(handler=self, selected_owner=owner).decode("utf-8"))
            return

        if path == "/preview":
            selected_sheet = query.get("sheet", [""])[0] if query.get("sheet") else ""
            try:
                rows_limit = int(query.get("rows", ["80"])[0])
            except ValueError:
                rows_limit = 80
            write_html(
                self,
                200,
                render_preview_page(selected_sheet=selected_sheet, rows_limit=rows_limit).decode("utf-8"),
            )
            return

        if path == "/healthz":
            write_text(self, 200, "ok")
            return

        if path == "/status":
            metadata = ensure_metadata_schema(load_metadata())
            write_json(
                self,
                200,
                {
                    "job": get_job_status(),
                    "master_exists": master_sales_exists(),
                    "master_name": metadata.get("master_name", ""),
                    "latest_name": metadata.get("latest_name", ""),
                    "last_owner": metadata.get("last_owner", ""),
                    "last_generated_at": metadata.get("last_generated_at", ""),
                    "last_updated_rows": metadata.get("last_updated_rows", ""),
                    "owner_statuses": metadata.get("owner_statuses", default_owner_statuses()),
                    "data_dir": str(DATA_DIR),
                    "storage": storage_description(),
                },
            )
            return

        if path == "/download/latest":
            if not master_sales_exists():
                write_html(
                    self,
                    404,
                    render_page(error="还没有共用销售排单可下载，请先上传本周排单。", handler=self).decode("utf-8"),
                )
                return
            metadata = load_metadata()
            download_name = metadata.get("latest_name") or metadata.get("master_name") or "销售排单_当前最新版.xlsx"
            send_xlsx(self, MASTER_SALES_PATH, download_name)
            return

        write_text(self, 404, "Not found")

    def do_HEAD(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)
        if path in ("/", "/index.html") or path.startswith("/owner/"):
            owner = selected_owner_from_path(path, query)
            write_html(self, 200, render_page(handler=self, selected_owner=owner).decode("utf-8"), head=True)
            return

        if path == "/healthz":
            write_text(self, 200, "ok", head=True)
            return

        if path == "/preview":
            write_html(self, 200, render_preview_page().decode("utf-8"), head=True)
            return

        if path == "/status":
            write_json(self, 200, {"ok": True}, head=True)
            return

        if path == "/download/latest":
            if not master_sales_exists():
                write_text(self, 404, "Not found", head=True)
                return
            metadata = load_metadata()
            download_name = metadata.get("latest_name") or metadata.get("master_name") or "销售排单_当前最新版.xlsx"
            send_xlsx(self, MASTER_SALES_PATH, download_name, head=True)
            return

        write_text(self, 404, "Not found", head=True)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/sales-master":
            handle_sales_master(self)
            return

        if parsed.path == "/edit-cell":
            handle_edit_cell(self)
            return

        if parsed.path != "/generate":
            write_text(self, 404, "Not found")
            return

        handle_generate(self)


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), SalesUploadHandler)
    public_url = get_public_share_url()
    print(f"销售排单网站已启动：http://0.0.0.0:{PORT}")
    print("现在不需要登录，按业务担当页面上传预测信息。")
    if public_url:
        print(f"公开访问：{public_url}/")
    else:
        lan_ip = get_lan_ip()
        print(f"本机访问：http://127.0.0.1:{PORT}")
        print(f"局域网访问：http://{lan_ip}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
