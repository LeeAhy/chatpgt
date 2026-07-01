from __future__ import annotations

import cgi
import html
from functools import lru_cache
import os
import shutil
import socket
import subprocess
import tempfile
import urllib.parse
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

    page = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
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
      .summary {{
        grid-template-columns: 1fr;
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
        <p class="subtext">选择业务担当，上传该担当的预测信息和销售排单。系统会按不同业务担当读取对应字段匹配“客户机种”，当月客户总预测会先扣掉已完成数据，再写入销售排单当前预留的空白数量/金额栏。</p>
      </div>
      <div class="access">{html.escape(access_note)}</div>
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
            <p class="subtext">{html.escape(upload_hint)} 销售排单必须保留“业务担当”和“客户机种”。</p>
          </div>
          <div class="badge">无需登录</div>
        </div>
        {message_html}
        {error_html}
        <form method="post" action="/generate" enctype="multipart/form-data">
          <div class="rule-panel">
            {html.escape(selected_owner_guide)}
            <span>当月会自动扣减同一客户机种已完成数量；扣完小于 0 时按 0 回填。</span>
          </div>
          <div>
            <label for="business_owner">业务担当</label>
            <select id="business_owner" name="business_owner">
              {owner_options}
            </select>
            <div class="hint">只会回填销售排单中属于所选业务担当的行。</div>
          </div>
          <div class="field-grid">
            <div>
              <label for="prediction_file">预测信息</label>
              <input id="prediction_file" name="prediction_file" type="file" accept=".xlsx,.xlsm,.csv,.tsv,.txt,.png,.jpg,.jpeg,.webp,.heic,.tif,.tiff" required>
              <div class="hint">支持 Excel、CSV、TXT；截图入口已接入，系统 OCR 可用时会自动识别。</div>
            </div>
            <div>
              <label for="sales_file">销售排单</label>
              <input id="sales_file" name="sales_file" type="file" accept=".xlsx,.xlsm" required>
              <div class="hint">系统会自动找销售排单里当前空白的数量/金额列。</div>
            </div>
          </div>
          <div class="actions">
            <button class="primary" type="submit">生成回填文件</button>
            <a class="secondary button" href="{owner_link(selected_owner)}">重置当前页面</a>
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


def save_upload(field: cgi.FieldStorage, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("wb") as out:
        shutil.copyfileobj(field.file, out)


def first_upload(form: cgi.FieldStorage, name: str):
    field = form[name]
    if isinstance(field, list):
        return field[0]
    return field


def handle_generate(handler: BaseHTTPRequestHandler, head: bool = False) -> None:
    if handler.headers.get_content_type() != UPLOAD_MIME:
        write_html(
            handler,
            400,
            render_page(error="请使用网页表单上传预测信息和销售排单。", handler=handler).decode("utf-8"),
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
        sales_field = first_upload(form, "sales_file")
        if not getattr(pred_field, "filename", None) or not getattr(sales_field, "filename", None):
            raise ValueError("请先选择预测信息和销售排单两个文件。")

        pred_suffix = uploaded_suffix(pred_field.filename)
        sales_suffix = uploaded_suffix(sales_field.filename)
        if pred_suffix not in SUPPORTED_PREDICTION_SUFFIXES:
            raise ValueError("预测信息请上传 Excel、CSV、TXT 或图片文件。")
        if sales_suffix not in SUPPORTED_SALES_SUFFIXES:
            raise ValueError("销售排单请上传 .xlsx 或 .xlsm 文件。")

        with tempfile.TemporaryDirectory(prefix="sales_upload_") as tmp_dir:
            tmp_path = Path(tmp_dir)
            pred_path = tmp_path / f"prediction{pred_suffix}"
            sales_path = tmp_path / f"sales{sales_suffix}"
            output_path = tmp_path / "generated.xlsx"

            save_upload(pred_field, pred_path)
            save_upload(sales_field, sales_path)

            summary = process_sales_workbooks(
                pred_path=pred_path,
                sales_path=sales_path,
                output_path=output_path,
                freeze_formulas=True,
                business_owner=selected_owner,
            )

            original_name = os.path.basename(str(sales_field.filename).replace("\\", "/"))
            stem = Path(original_name).stem if original_name else "销售排单"
            download_name = f"{stem}_{selected_owner}_已回填.xlsx"
            file_size = output_path.stat().st_size

            handler.send_response(200)
            handler.send_header("Content-Type", XLSX_MIME)
            handler.send_header("Content-Disposition", content_disposition(download_name))
            handler.send_header("Content-Length", str(file_size))
            handler.send_header("X-Updated-Rows", str(summary["updated_rows"]))
            if summary.get("warnings"):
                handler.send_header("X-Warnings", urllib.parse.quote("；".join(summary["warnings"])))
            if summary.get("missing_rows"):
                handler.send_header("X-Missing-Rows", str(len(summary["missing_rows"])))
            handler.send_header("Cache-Control", "no-store")
            handler.end_headers()
            if not head:
                with output_path.open("rb") as f:
                    shutil.copyfileobj(f, handler.wfile)
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

        if path == "/healthz":
            write_text(self, 200, "ok")
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

        write_text(self, 404, "Not found", head=True)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
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
