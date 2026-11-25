#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Word Formatter Pro — 本地 Web 界面（旁挂实现，不修改上游任何文件）。

启动：  python wfp_web.py
访问：  http://127.0.0.1:8878

设计约束：
- 只用标准库 + 上游已有依赖（python-docx / lxml / fonttools），不引入 Flask/FastAPI；
- 只监听 127.0.0.1，不对外暴露；
- 每个文件一次请求，便于前端显示真实进度；
- 处理串行化（LOCK），避免 COM 与临时文件竞争。
"""
from __future__ import annotations

import base64
import json
import os
import re
import socket
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

FROZEN = bool(getattr(sys, "frozen", False))
ROOT = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
RES = Path(getattr(sys, "_MEIPASS", ROOT))          # 只读资源（webui）在 onefile 模式下解压到这里
WEB = RES / "webui"

# 打包后 exe 可能被放在只读目录，输出与临时目录统一落到可写位置
def _writable(folder: Path) -> bool:
    try:
        folder.mkdir(parents=True, exist_ok=True)
        probe = folder / ".wfp_write_probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except Exception:
        return False


def _pick_data_home() -> Path:
    if not FROZEN:
        return ROOT
    local = os.environ.get("LOCALAPPDATA")
    candidates = []
    if local:
        candidates.append(Path(local) / "WordFormatter")
    candidates.append(ROOT / "WordFormatterData")
    candidates.append(Path(tempfile.gettempdir()) / "WordFormatter")
    for folder in candidates:
        if _writable(folder):
            return folder
    return Path(tempfile.gettempdir()) / "WordFormatter"


DATA_HOME = _pick_data_home()
OUT = DATA_HOME / "out"
if not FROZEN:
    sys.path.insert(0, str(ROOT))

# 本机 %TEMP% 对 Python 不可写，pywin32 的 gen_py 缓存和 COM 都会挂；
# 统一把临时目录重定向到项目内，必须在 import wfp_core（它会用到 win32com）之前。
TMP = DATA_HOME / "tmp"
try:
    TMP.mkdir(parents=True, exist_ok=True)
except Exception:
    TMP = Path(tempfile.gettempdir())
os.environ["TMP"] = str(TMP)
os.environ["TEMP"] = str(TMP)
tempfile.tempdir = str(TMP)

# PyInstaller 的 pyi_rth_win32comgenpy 钩子在本模块之前已按系统 TEMP 设定 gen_py 路径，
# 这里显式改回可写目录，否则 import win32com.client 会因建目录被拒而崩。
try:
    import win32com  # noqa: E402

    _gen_py = TMP / "gen_py" / f"{sys.version_info.major}.{sys.version_info.minor}"
    _gen_py.mkdir(parents=True, exist_ok=True)
    win32com.__gen_path__ = str(_gen_py)
except Exception:
    pass

from wfp_config import DEFAULT_CONFIG, FONT_SIZE_MAP, PRESET_FONT_OPTIONS  # noqa: E402
from wfp_core import (  # noqa: E402
    BLANK_LINE_MODE_OPTIONS,
    DEFAULT_BLANK_LINE_MODE,
    LegacyConversionUnavailable,
    WPSAppManager,
    WordProcessor,
    _initialize_com_for_thread,
    _uninitialize_com_for_thread,
)
from wfp_version import APP_NAME, __version__  # noqa: E402

PORT = int(os.environ.get("WFP_WEB_PORT", "8878"))
LOCK = threading.Lock()

# 预设方案：只覆盖与默认不同的键，其余由 normalize_format_options 补全
PRESETS = {
    "党政机关公文（默认）": {},
    "企业报告": {
        "body_font": "微软雅黑",
        "body_size": 10.5,
        "h1_font": "微软雅黑",
        "h1_size": 16,
        "h2_font": "微软雅黑",
        "h2_size": 14,
        "title_font": "微软雅黑",
        "title_size": 22,
        "line_spacing_unit": "multiple",
        "line_spacing_multiple": 1.5,
        "first_line_indent_chars": 0.0,
        "enable_table_formatting": True,
        "table_font": "微软雅黑",
        "table_size": 10.5,
        "margin_top": 2.54,
        "margin_bottom": 2.54,
        "margin_left": 3.18,
        "margin_right": 3.18,
    },
    "学术论文": {
        "body_font": "宋体",
        "body_size": 12,
        "h1_font": "黑体",
        "h1_size": 15,
        "h2_font": "黑体",
        "h2_size": 14,
        "title_font": "黑体",
        "title_size": 18,
        "line_spacing_unit": "multiple",
        "line_spacing_multiple": 1.5,
        "first_line_indent_chars": 2.0,
        "margin_top": 2.5,
        "margin_bottom": 2.5,
        "margin_left": 3.0,
        "margin_right": 3.0,
    },
}

# 高频键（前端手写控件），其余进「高级」自动生成
FREQUENT_KEYS = {
    "margin_top", "margin_bottom", "margin_left", "margin_right",
    "body_font", "body_size", "line_spacing", "line_spacing_unit",
    "line_spacing_multiple", "first_line_indent_chars",
    "title_font", "title_size", "h1_font", "h1_size", "h2_font", "h2_size",
    "enable_table_formatting", "table_auto_col_width", "table_border_mode",
    "enable_attachment_formatting", "blank_line_mode", "convert_auto_numbering",
    "enable_title_after_blank_line", "enable_title_after_spacing",
    "force_a4", "normalize_punctuation",
}


# 常见中西文对照：注册表里系统字体多用英文名
_FONT_FILE_ALIAS = {
    "simsun.ttc": "宋体", "simsunb.ttf": "宋体", "simsunextg.ttf": "宋体",
    "simfang.ttf": "仿宋", "simhei.ttf": "黑体", "simkai.ttf": "楷体",
    "simli.ttf": "隶书", "simyou.ttf": "幼圆",
    "msyh.ttc": "微软雅黑", "msyhbd.ttc": "微软雅黑", "msyhl.ttc": "微软雅黑",
}

_INSTALLED_CACHE: dict = {"at": 0.0, "fonts": None}


def _norm_font(name: str) -> str:
    s = str(name).strip().lower()
    s = re.sub(r"\s*\((truetype|opentype|truetype collection)\)\s*$", "", s)
    s = re.sub(r"[\s_\-]", "", s)
    s = re.sub(r"(gb2312|gbk|简体|简)$", "", s)
    return s


def installed_fonts() -> list:
    """枚举本机已安装字体（HKLM + HKCU），带 10 分钟缓存。"""
    now = time.time()
    if _INSTALLED_CACHE["fonts"] and now - _INSTALLED_CACHE["at"] < 600:
        return _INSTALLED_CACHE["fonts"]
    names: set = set()
    try:
        import winreg  # 仅 Windows 可用

        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            try:
                with winreg.OpenKey(root, path) as key:
                    idx = 0
                    while True:
                        try:
                            display, data, _ = winreg.EnumValue(key, idx)
                            idx += 1
                        except OSError:
                            break
                        names.add(_norm_font(display))
                        fname = str(data).split("\\")[-1].lower()
                        alias = _FONT_FILE_ALIAS.get(fname)
                        if alias:
                            names.add(_norm_font(alias))
            except OSError:
                continue
    except Exception:
        pass
    result = sorted(names)
    if result:
        _INSTALLED_CACHE["fonts"] = result
        _INSTALLED_CACHE["at"] = now
    return result


def config_meta() -> dict:
    defaults = dict(DEFAULT_CONFIG)
    advanced = {}
    for k, v in sorted(defaults.items()):
        if k in FREQUENT_KEYS:
            continue
        if isinstance(v, bool):
            t = "bool"
        elif isinstance(v, (int, float)):
            t = "number"
        else:
            t = "text"
        advanced[k] = {"type": t, "default": v}
    return {
        "app": APP_NAME,
        "version": __version__,
        "defaults": defaults,
        "advanced": advanced,
        "presets": PRESETS,
        "fonts": PRESET_FONT_OPTIONS,
        "fontSizes": FONT_SIZE_MAP,
        "blankLineModes": BLANK_LINE_MODE_OPTIONS,
        "defaultBlankLineMode": DEFAULT_BLANK_LINE_MODE,
        "defaultOut": str(OUT),
        "installedFonts": installed_fonts(),
    }


def safe_name(raw: str) -> str:
    name = Path(str(raw).replace("\\", "/")).name
    name = "".join(c for c in name if c not in '<>:"/\\|?*').strip() or "file.docx"
    return name


def format_one(payload: dict) -> dict:
    name = safe_name(payload.get("name") or "file.docx")
    raw = payload.get("data")
    if not raw:
        return {"ok": False, "error": "没有收到文件内容"}
    try:
        blob = base64.b64decode(raw, validate=False)
    except Exception:
        return {"ok": False, "error": "文件内容解码失败"}

    config = payload.get("config") or {}
    if not isinstance(config, dict):
        return {"ok": False, "error": "配置格式不正确"}

    # 上游的 normalize_format_options 只做迁移/校验，不补默认值，这里必须自己合并
    merged = dict(DEFAULT_CONFIG)
    merged.update({k: v for k, v in config.items() if v is not None})

    out_dir = Path(payload.get("outputDir") or OUT)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        return {"ok": False, "error": f"输出目录不可用：{exc}"}

    tmp_dir = Path(tempfile.mkdtemp(prefix="wfpweb_"))
    src = tmp_dir / name
    try:
        src.write_bytes(blob)
    except Exception as exc:
        return {"ok": False, "error": f"写入临时文件失败：{exc}"}

    dst = out_dir / f"{Path(name).stem}_排版.docx"
    logs: list[str] = []
    com_ok = _initialize_com_for_thread(logs.append)

    try:
        with LOCK:
            processor = WordProcessor(
                merged,
                logs.append,
                com_manager=WPSAppManager(logs.append),
            )
            processor.format_document(str(src), str(dst))
        return {
            "ok": True,
            "output": str(dst),
            "input": name,
            "size": dst.stat().st_size if dst.exists() else 0,
            "logs": logs[-40:],
        }
    except LegacyConversionUnavailable as exc:
        return {"ok": False, "error": f"旧格式转换不可用：{exc}", "logs": logs[-40:]}
    except Exception as exc:  # 单文件失败不拖垮整批
        return {
            "ok": False,
            "error": f"{type(exc).__name__}: {exc}",
            "detail": traceback.format_exc(limit=3),
            "logs": logs[-40:],
        }
    finally:
        _uninitialize_com_for_thread(com_ok, logs.append)
        try:
            src.unlink(missing_ok=True)
            tmp_dir.rmdir()
        except Exception:
            pass


class Handler(BaseHTTPRequestHandler):
    server_version = "wfp-web"
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        # --noconsole 打包后 sys.stderr 为 None，这里任何写入都会抛异常并断掉连接；
        # 桌面端不需要访问日志，直接静默。
        try:
            if not FROZEN and sys.stderr:
                sys.stderr.write("  %s\n" % (fmt % args))
        except Exception:
            pass

    def _send(self, code: int, body: bytes, ctype: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, obj, code: int = 200) -> None:
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def do_GET(self) -> None:
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            index = WEB / "index.html"
            if not index.exists():
                self._send(500, "webui/index.html 缺失".encode("utf-8"), "text/plain; charset=utf-8")
                return
            self._send(200, index.read_bytes(), "text/html; charset=utf-8")
            return
        if path == "/api/meta":
            self._json(config_meta())
            return
        if path == "/favicon.ico":
            self._send(204, b"", "image/x-icon")
            return
        self._send(404, b'{"ok":false,"error":"not found"}', "application/json; charset=utf-8")

    def do_POST(self) -> None:
        path = self.path.split("?")[0]
        length = int(self.headers.get("Content-Length") or 0)
        if length > 64 * 1024 * 1024:
            self._json({"ok": False, "error": "文件过大（上限 64 MB）"}, 413)
            return
        try:
            payload = json.loads(self.rfile.read(length) or b"{}")
        except Exception as exc:
            self._json({"ok": False, "error": f"请求解析失败：{exc}"}, 400)
            return

        if path == "/api/format":
            self._json(format_one(payload))
            return
        if path == "/api/open-dir":
            target = payload.get("path") or str(OUT)
            try:
                Path(target).mkdir(parents=True, exist_ok=True)
                os.startfile(target)  # noqa: S606 - 本地工具，用户主动触发
                self._json({"ok": True})
            except Exception as exc:
                self._json({"ok": False, "error": str(exc)})
            return
        self._json({"ok": False, "error": "not found"}, 404)


def pick_port(start: int, span: int = 20) -> int:
    """端口被占用时自动向后找一个空闲端口，避免二次启动直接崩。"""
    for port in range(start, start + span):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.25)
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start


def say(msg: str) -> None:
    """打包成 noconsole 后 stdout 是 None，print 会炸；统一走日志文件。"""
    try:
        if sys.stdout:
            print(msg, flush=True)
    except Exception:
        pass
    try:
        with open(DATA_HOME / "app.log", "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}\n")
    except Exception:
        pass


def find_edge() -> str | None:
    for path in (
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ):
        if Path(path).exists():
            return path
    return None


def window_disabled() -> bool:
    return os.environ.get("WFP_NO_WINDOW") == "1"


def run_app_window(url: str) -> bool:
    """用系统 Edge 的 app 模式开一个无地址栏窗口。进程存活期间阻塞，关窗即返回。"""
    if window_disabled():
        return False
    edge = find_edge()
    if not edge:
        say("未找到 Edge，退回默认浏览器")
        return False
    profile = DATA_HOME / "window-profile"
    profile.mkdir(parents=True, exist_ok=True)
    args = [
        edge,
        f"--app={url}",
        f"--user-data-dir={profile}",
        "--window-size=1200,860",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-features=Translate,msEdgeSplitScreen",
        # 由本进程（Python 子进程）拉起 Edge 时，Chromium 的渲染进程沙箱初始化会失败，
        # 表现为窗口出现但整片空白。实测只有 --no-sandbox 能恢复渲染。
        # 加载内容全部来自本机 127.0.0.1 的自有页面，不加载任何外部资源，风险可控。
        "--no-sandbox",
    ]
    try:
        proc = subprocess.Popen(args)
        say(f"窗口已启动 pid={proc.pid}")
        proc.wait()          # 关闭窗口 → Edge 该 profile 主进程退出
        say("窗口已关闭")
        return True
    except Exception as exc:
        say(f"窗口启动失败，退回浏览器：{exc}")
        return False


def main() -> int:
    DATA_HOME.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)
    port = pick_port(PORT)
    url = f"http://127.0.0.1:{port}/"
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    say(f"{APP_NAME} v{__version__}  {url}  输出={OUT}")

    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    try:
        if not run_app_window(url):
            if os.environ.get("WFP_NO_BROWSER") != "1":
                webbrowser.open(url)
            while True:                 # 浏览器模式：靠 Ctrl+C / 关进程退出
                time.sleep(3600)
    except KeyboardInterrupt:
        say("已停止")
    finally:
        httpd.shutdown()
        httpd.server_close()
    return 0



if __name__ == "__main__":
    raise SystemExit(main())
