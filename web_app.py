from __future__ import annotations

import cgi
import html
import json
from functools import lru_cache
import os
import shutil
import socket
import subprocess
import tempfile
import threading
import urllib.parse
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional

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


def load_metadata() -> dict:
    if not METADATA_PATH.exists():
        return {}
    try:
        return json.loads(METADATA_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_metadata(metadata: dict) -> None:
    ensure_data_dir()
    METADATA_PATH.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")


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
    return MASTER_SALES_PATH.exists() and MASTER_SALES_PATH.stat().st_size > 0


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
    metadata = load_metadata()
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
    detail_text = " ｜ ".join(detail_parts) if detail_parts else "共用排单会在每位业务上传预测后自动保存为最新版。"
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
        </form>
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
            <input id="prediction_file" name="prediction_file" type="file" accept=".xlsx,.xlsm,.csv,.tsv,.txt,.png,.jpg,.jpeg,.webp,.heic,.tif,.tiff" required>
            <div class="hint">支持 Excel、CSV、TXT；截图入口已接入，系统 OCR 可用时会自动识别。无需再上传销售排单。</div>
          </div>
          <div class="actions">
            <button class="primary" type="submit">开始回填</button>
            <a class="secondary button" href="{owner_link(selected_owner)}">刷新处理状态</a>
            {download_latest_html}
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
                shutil.copy2(MASTER_SALES_PATH, BACKUP_DIR / backup_name)
            save_upload(sales_field, MASTER_SALES_PATH)
            shutil.copy2(MASTER_SALES_PATH, LATEST_OUTPUT_PATH)
            save_metadata(
                {
                    "master_name": original_name,
                    "master_uploaded_at": now_label(),
                    "latest_name": f"{Path(original_name).stem}_当前最新版.xlsx",
                    "last_owner": "",
                    "last_generated_at": "",
                    "last_updated_rows": "",
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
            raise ValueError("预测信息请上传 Excel、CSV、TXT 或图片文件。")
        if not master_sales_exists():
            raise ValueError("请先在页面上方上传本周共用销售排单。")
        if job_is_running():
            raise ValueError("上一份预测还在处理，请稍后刷新页面，完成后再上传下一位业务预测。")

        ensure_data_dir()
        pred_name = datetime.now().strftime("%Y%m%d_%H%M%S") + f"_{selected_owner}{pred_suffix}"
        pred_path = UPLOAD_DIR / pred_name
        save_upload(pred_field, pred_path)
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

        if path == "/healthz":
            write_text(self, 200, "ok")
            return

        if path == "/status":
            metadata = load_metadata()
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
