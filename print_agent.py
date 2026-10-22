# -*- coding: utf-8 -*-
"""
明道云本地打印代理
==================
工作方式（watch 轮询「待打印队列」表）：
  明道云业务表按钮 → 工作流「获取记录打印文件」渲染 Word/Excel 并「新建记录」
  汇入待打印队列表 → 本代理轮询 V3 接口取附件 → 本地转 PDF → 静默打印 → 回写状态。
  需在 config.json 填好明道云连接参数（模板见 config.example.json）。

依赖（零 pip，仅标准库）：
  - Python 3.8+
  - Word/Excel → PDF：本机 Office 或 WPS（COM 调用，自动探测）
  - PDF 静默打印：SumatraPDF（https://www.sumatrapdfreader.org 下载 Portable 版，
    把 SumatraPDF.exe 放到本文件同目录，或安装到默认路径，均会被自动探测）
  - HTML → PDF：系统自带 Edge 或 Chrome 无头渲染（仅 --test 自检用）

用法：
  python print_agent.py            # 启动代理（轮询待打印队列）；首次运行自动进入配置向导，
                                   # 逐项填写明道云连接/字段/打印机并校验，保存 config.json
  python print_agent.py --test     # 自检：走完整管线打一张测试页到默认打印机
  python print_agent.py --config D:\\mdprint.json
  python print_agent.py --sumatra "D:\\tools\\SumatraPDF.exe"

分发给客户机：pyinstaller --onefile --name md-print-agent print_agent.py

明道云侧标准工作流（业务表任意，任务统一汇入「待打印队列」表）：
  自定义动作按钮触发
  → 节点1「获取记录打印文件」
      打印对象 = 按钮触发的工作表；打印模板 / 文件名 = 按业务自行配置
  → 节点2「新建记录」
      选择工作表 = 待打印队列；option = 新增一条记录
      名称 = 自定义记录名称；打印编号 = 留空(自动编号)
      打印状态 = 待打印；附件 = 节点1 输出的打印文件
「待打印队列」表字段规范：名称(标题/文本)、打印编号(自动编号，规则 print-自然数)、
  打印状态(单选：待打印/打印中/已打印/失败)、附件(附件字段)
"""
import argparse
import base64
import json
import os
import shutil
import ssl
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

VERSION = "0.2.0"
CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0
PRINT_LOCK = threading.Lock()  # ponytail: 全局串行打印；出现并发批量打印需求时升级为 queue.Queue + 单 worker
SUMATRA = None
# PyInstaller onefile 下 __file__ 指向解压临时目录；exe 场景以 exe 所在目录为基准（config.json / SumatraPDF 放 exe 旁）
BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent

SAMPLE_HTML = """<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  body { font-family: "Microsoft YaHei", sans-serif; padding: 40px; }
  h1 { border-bottom: 2px solid #333; padding-bottom: 8px; }
  table { border-collapse: collapse; margin-top: 16px; }
  td, th { border: 1px solid #999; padding: 6px 18px; }
</style></head><body>
<h1>明道云静默打印 · 自检页</h1>
<p>时间：__TIME__</p>
<p>如果你拿到这张纸，说明 油猴脚本 → 本地代理 → 系统打印队列 全链路已通。</p>
<table><tr><th>环节</th><th>组件</th></tr>
<tr><td>内容捕获</td><td>明道云静默打印助手.user.js</td></tr>
<tr><td>HTML→PDF</td><td>Edge/Chrome 无头模式</td></tr>
<tr><td>静默打印</td><td>SumatraPDF</td></tr></table>
</body></html>"""


def _first_existing(paths):
    for p in paths:
        try:
            if p and Path(p).exists():
                return str(p)
        except OSError:
            continue
    return None


def find_browser():
    """找到用于 HTML→PDF 的无头浏览器（Edge 优先，Chrome 兜底）"""
    here = BASE_DIR
    return _first_existing([
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        str(here / "msedge.exe"),
        shutil.which("msedge"),
        shutil.which("chrome"),
    ])


def find_sumatra(override=None):
    here = BASE_DIR
    local_app = os.environ.get("LOCALAPPDATA", "")
    return _first_existing([
        override,
        shutil.which("SumatraPDF") or shutil.which("SumatraPDF.exe"),
        str(here / "SumatraPDF.exe"),                 # Portable 版放脚本同目录
        str(here / "SumatraPDF" / "SumatraPDF.exe"),
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        os.path.join(local_app, "SumatraPDF", "SumatraPDF.exe") if local_app else None,
    ])


def list_printers():
    """[(名称, 是否默认)]；失败返回 []"""
    try:
        if os.name == "nt":
            # 中文 Windows 上 PowerShell 管道输出默认 GBK：强制其转 UTF-8，Python 侧按字节捕获显式解码，
            # 不依赖任何系统区域/编码设置，换机器不翻车
            out = subprocess.run(
                ["powershell", "-NoProfile", "-Command",
                 "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; "
                 "Get-CimInstance Win32_Printer | Select-Object Name,Default | ConvertTo-Json -Compress"],
                capture_output=True, timeout=30, creationflags=CREATE_NO_WINDOW)
            data = json.loads(out.stdout.decode("utf-8", errors="replace") or "[]")
            if isinstance(data, dict):  # 只有一台打印机时 ConvertTo-Json 返回单对象
                data = [data]
            return [(p["Name"], bool(p.get("Default"))) for p in data]
        # ponytail: 非Windows走lpstat，未实测；Windows是主路径
        out = subprocess.run(["lpstat", "-p"], capture_output=True, text=True, timeout=20)
        return [(line.split()[1], False) for line in out.stdout.splitlines()
                if line.startswith("printer ")]
    except Exception as e:
        print(f"打印机枚举失败: {e}")
        return []


def html_to_pdf(html_file: Path, out_pdf: Path):
    """用 Edge/Chrome 无头模式把 HTML 渲染为 PDF"""
    browser = find_browser()
    if not browser:
        raise RuntimeError("未找到 Edge/Chrome，无法把 HTML 转成 PDF")
    # 两代无头参数各试一次：新版本用 --headless=new，老版本只认 --headless
    for headless, no_header in (("--headless=new", "--no-pdf-header-footer"),
                                ("--headless", "--print-to-pdf-no-header")):
        out_pdf.unlink(missing_ok=True)
        # 独立 user-data-dir：否则 Edge 已在运行时，无头进程会把任务转交常驻实例导致拿不到 PDF
        profile = html_file.parent / "browser-profile"
        cmd = [browser, headless, no_header, "--disable-gpu", "--no-first-run",
               "--disable-extensions", "--virtual-time-budget=4000",
               f"--user-data-dir={profile}",
               f"--print-to-pdf={out_pdf}", html_file.as_uri()]
        try:
            subprocess.run(cmd, capture_output=True, timeout=90,
                           creationflags=CREATE_NO_WINDOW)
        except subprocess.TimeoutExpired:
            continue
        if out_pdf.exists() and out_pdf.stat().st_size > 0:
            return out_pdf
    raise RuntimeError("HTML → PDF 失败（无头浏览器未产出文件）")


def print_pdf(pdf: Path, printer=None, copies=1):
    """静默打印 PDF 到指定打印机（默认打印机则留空）"""
    if os.name != "nt":
        cmd = ["lp"] + (["-d", printer] if printer else []) + [str(pdf)] * max(1, copies)
        # ponytail: 非Windows用lp，未实测
        subprocess.run(cmd, check=True, timeout=60)
        return
    global SUMATRA
    if not SUMATRA:
        raise RuntimeError("未找到 SumatraPDF：请到 https://www.sumatrapdfreader.org "
                           "下载 Portable 版，把 SumatraPDF.exe 放到本脚本同目录")
    cmd = [SUMATRA]
    cmd += ["-print-to", printer] if printer else ["-print-to-default"]
    if copies and int(copies) > 1:
        cmd += ["-print-settings", f"copie={int(copies)}"]  # Sumatra 手册的拼写就是 copie
    cmd += ["-silent", "-exit-when-done", str(pdf)]
    r = subprocess.run(cmd, capture_output=True, timeout=120,
                       creationflags=CREATE_NO_WINDOW)
    if r.returncode != 0:
        raise RuntimeError(f"SumatraPDF 退出码 {r.returncode}：{r.stderr.decode(errors='ignore')[:200]}")


def do_print_job(job: dict) -> dict:
    """一个打印任务：{html | pdfBase64, printer?, copies?} → 出纸"""
    printer = job.get("printer") or None
    copies = int(job.get("copies") or 1)
    with tempfile.TemporaryDirectory(prefix="mdprint_") as td:
        td = Path(td)
        if job.get("pdfBase64"):
            pdf = td / "job.pdf"
            pdf.write_bytes(base64.b64decode(job["pdfBase64"]))
        elif job.get("html"):
            # ponytail: 外链资源须由油猴脚本捕获时内联；页面级 http 外链无头渲染时仍会尝试加载
            src = td / "job.html"
            src.write_text(job["html"], encoding="utf-8")
            pdf = td / "job.pdf"
            html_to_pdf(src, pdf)
        else:
            raise ValueError("任务体需要 html 或 pdfBase64 字段")
        print_pdf(pdf, printer, copies)
    return {"printer": printer or "默认打印机", "copies": copies}


# ==================== watch 模式：按钮直印 ====================

OFFICE_TO_PDF_PS = r"""
param($src, $out, $kind)
# Word/Excel -> PDF：优先 MS Office COM，WPS COM 兜底（KWPS=WPS文字 KET=WPS表格，接口兼容）
$ids = if ($kind -eq 'word') { @('Word.Application','KWPS.Application') } else { @('Excel.Application','KET.Application') }
$app = $null
foreach ($id in $ids) {
  try { $app = New-Object -ComObject $id; break } catch { }
}
if (-not $app) { Write-Error 'no COM engine'; exit 3 }
$app.Visible = $false
$app.DisplayAlerts = 0
try {
  if ($kind -eq 'word') {
    $doc = $app.Documents.Open($src, $false, $true)
    $doc.ExportAsFixedFormat($out, 17)   # 17 = wdExportFormatPDF
    $doc.Close($false)
  } else {
    $wb = $app.Workbooks.Open($src, 0, $true)
    $wb.ExportAsFixedFormat(0, $out)     # 0 = xlTypePDF
    $wb.Close($false)
  }
  exit 0
} finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
"""

OFFICE_EXTS = {".doc": "word", ".docx": "word", ".rtf": "word", ".wps": "word",
               ".xls": "excel", ".xlsx": "excel", ".et": "excel", ".csv": "excel"}


def office_to_pdf(src: Path, out_pdf: Path):
    """用本机 Office/WPS 把 Word/Excel 转成 PDF（PowerShell COM，零 pip）"""
    kind = OFFICE_EXTS.get(src.suffix.lower())
    if not kind:
        raise ValueError(f"不支持的文件类型: {src.name}")
    with tempfile.TemporaryDirectory(prefix="mdoffice_") as td:
        ps1 = Path(td) / "to_pdf.ps1"
        ps1.write_text(OFFICE_TO_PDF_PS, encoding="utf-8-sig")
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1),
             "-src", str(src), "-out", str(out_pdf), "-kind", kind],
            capture_output=True, timeout=180, creationflags=CREATE_NO_WINDOW)
    if r.returncode == 0 and out_pdf.exists() and out_pdf.stat().st_size > 0:
        return out_pdf
    raise RuntimeError(f"Office→PDF 失败 (exit {r.returncode}): "
                       f"{r.stderr.decode(errors='ignore')[:200]}")


class MingdaoClient:
    """明道云 V3 开放接口最小封装（鉴权：HAP-Appkey / HAP-Sign）"""

    def __init__(self, cfg):
        self.base = cfg["apiBase"].rstrip("/")
        self.headers = {"Content-Type": "application/json",
                        "HAP-Appkey": cfg["appKey"], "HAP-Sign": cfg["appSecret"]}
        self.ssl_ctx = ssl._create_unverified_context() if cfg.get("insecure") else None

    def _post(self, path, payload, method="POST"):
        req = urllib.request.Request(self.base + path, method=method,
                                     data=json.dumps(payload).encode("utf-8"),
                                     headers=self.headers)
        with urllib.request.urlopen(req, timeout=30, context=self.ssl_ctx) as r:
            data = json.loads(r.read().decode("utf-8"))
        if data.get("success") is False or data.get("error_code") not in (None, 0, 1):
            raise RuntimeError(f"V3 {path} 失败: {data.get('error_msg') or data}")
        return data.get("data") or data

    def rows_list(self, worksheet_id, filt, page_size=50):
        return self._post(f"/v3/app/worksheets/{worksheet_id}/rows/list",
                          {"filter": filt, "pageIndex": 1, "pageSize": page_size})

    def row_update(self, worksheet_id, row_id, control_id, value, is_select):
        fields = [{"id": control_id, "value": ([value] if is_select else value)}]
        # PATCH：实测 V3 更新记录用 PATCH（POST/PUT 均 405）；单选传文本值即可，服务端自动匹配选项
        # triggerWorkflow=False：回写状态绝不能触发工作流，否则渲染→打印→回写会死循环
        return self._post(f"/v3/app/worksheets/{worksheet_id}/rows/{row_id}",
                          {"fields": fields, "triggerWorkflow": False}, method="PATCH")

    def download(self, url) -> bytes:
        for headers in ({}, self.headers):  # downloadUrl 通常是预签名地址；私有部署可能仍要鉴权头
            req = urllib.request.Request(url, headers=headers)
            try:
                with urllib.request.urlopen(req, timeout=120, context=self.ssl_ctx) as r:
                    return r.read()
            except urllib.error.HTTPError as e:
                if e.code not in (401, 403) or headers:
                    raise
        raise RuntimeError("下载失败: " + url[:80])


def setup_wizard(path: str) -> dict:
    """首次运行配置向导：逐项引导填写 + 即时校验，保存 config.json 后直接进入 watch"""
    print("=" * 56)
    print("  首次运行配置向导（明道云打印代理）")
    print("  直接回车 = 使用 [方括号] 里的默认值")
    print("=" * 56)

    def ask(prompt, default=None, required=False):
        while True:
            tip = f" [{default}]" if default is not None and default != "" else ""
            v = input(f"{prompt}{tip}: ").strip()
            if v:
                return v
            if default is not None and default != "":
                return default
            if not required:
                return ""
            print("  [X] 此项必填，请输入")

    cfg = {}
    print("\n— 明道云连接（应用 → 设置 → API 密钥 页面获取）—")
    cfg["apiBase"] = ask("API 地址（私有部署填 https://你的域名）", "https://api.mingdao.com")
    cfg["appKey"] = ask("AppKey", required=True)
    cfg["appSecret"] = ask("Sign", required=True)

    print("\n— 工作表 —")
    auto_created = False
    if input("还没有任务表？可由代理自动创建「待打印队列」标准表 (y/n) [n]: ").strip().lower() == "y":
        ws_name = ask("表名", "待打印队列")
        try:
            md = MingdaoClient(cfg)
            r = md._post("/v3/app/worksheets", {
                "name": ws_name,
                "fields": [
                    {"name": "名称", "alias": "title", "type": "Text", "isTitle": True},
                    {"name": "打印编号", "alias": "num", "type": "AutoNumber"},
                    {"name": "打印状态", "alias": "status", "type": "Dropdown",
                     "options": [{"value": v} for v in ("待打印", "打印中", "已打印", "失败")]},
                    {"name": "附件", "alias": "file", "type": "Attachment"},
                ]})
            cfg["worksheetId"] = r["worksheetId"]
            cfg["fileControlId"] = "file"
            cfg["statusControlId"] = "status"
            print(f"  [OK] 已创建「{ws_name}」(ID: {cfg['worksheetId'][:14]}…)，字段别名已自动对齐")
            print("      编号规则默认纯数字递增；如需 print- 前缀，请在字段设置中调整自动编号规则")
            print("      剩余待办（需在明道云界面操作）：配置打印模板 + 按钮工作流（见文件头「标准工作流」）")
            auto_created = True
        except Exception as e:
            print(f"  [X] 创建失败: {e}")
            print("      改为手工填写已有工作表信息")
    if "worksheetId" not in cfg:
        cfg["worksheetId"] = ask("工作表 ID 或别名", required=True)
        cfg["fileControlId"] = ask("附件字段（打印文件写入的附件字段）", "file")
        cfg["statusControlId"] = ask("状态字段（打印状态流转字段）", "status")

    if not auto_created:  # 自动建表的选项即默认值，跳过询问
        print("\n— 状态值（须与工作表字段里的选项/文本一致）—")
        cfg["statusPending"] = ask("待打印", "待打印")
        cfg["statusPrinting"] = ask("打印中", "打印中")
        cfg["statusDone"] = ask("已打印", "已打印")
        cfg["statusFailed"] = ask("失败", "失败")

    print("\n— 打印机 —")
    printers = list_printers()
    for i, (name, is_default) in enumerate(printers, 1):
        print(f"  {i}. {name}" + ("（系统默认）" if is_default else ""))
    choice = input(f"选择编号 1-{len(printers)}（回车 = 系统默认打印机）: ").strip()
    cfg["printer"] = printers[int(choice) - 1][0] if choice.isdigit() and 1 <= int(choice) <= len(printers) else ""
    cfg["interval"] = int(ask("轮询间隔（秒）", "3"))
    cfg["insecure"] = False

    print("\n— 校验连接 —")
    try:
        md = MingdaoClient(cfg)
        probe = md.rows_list(cfg["worksheetId"], {"type": "group", "logic": "AND", "children": []}, 1)
        rows = probe.get("rows") or []
        if rows:
            sample = rows[0].get(cfg["statusControlId"])
            print(f"  [OK] 鉴权通过；状态字段类型: {'单选（传文本值即可）' if isinstance(sample, list) else '文本'}")
            print(f"  [OK] 附件字段: {'存在' if cfg['fileControlId'] in rows[0] else '样例记录中未见（可能为空，不影响）'}")
        else:
            print("  [OK] 鉴权通过（表内暂无记录，字段类型待首次任务时自动探测）")
    except Exception as e:
        print(f"  [X] 校验失败: {e}")
        if input("仍然保存这份配置吗？(y/n) [n]: ").strip().lower() != "y":
            raise SystemExit("已取消，配置未保存")

    Path(path).write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n[OK] 配置已保存 → {path}")
    return cfg


def load_watch_config(path: str) -> dict:
    p = Path(path)
    if not p.exists():
        cfg = setup_wizard(path)  # 首次运行：向导生成配置
    else:
        cfg = json.loads(p.read_text(encoding="utf-8"))
        required = ["apiBase", "appKey", "appSecret", "worksheetId", "fileControlId", "statusControlId"]
        missing = [k for k in required if not cfg.get(k)]
        if missing:
            raise SystemExit(f"config.json 缺少必填项: {', '.join(missing)}")
    # 默认值补全对向导/文件两条路径统一生效（向导自动建表分支不询问状态值）
    for k, v in {"statusPending": "待打印", "statusPrinting": "打印中", "statusDone": "已打印",
                 "statusFailed": "失败", "printer": "", "interval": 3, "insecure": False}.items():
        cfg.setdefault(k, v)
    return cfg


def run_watch(cfg: dict):
    md = MingdaoClient(cfg)
    ws = cfg["worksheetId"]
    file_ctl, status_ctl = cfg["fileControlId"], cfg["statusControlId"]

    # 探测状态字段是文本还是单选：单选筛选/更新必须用选项 key，文本直接用值
    is_select = False
    probe = md.rows_list(ws, {"type": "group", "logic": "AND", "children": []}, 1)
    rows = probe.get("rows") or []
    if rows:
        sample = rows[0].get(status_ctl)
        is_select = isinstance(sample, list)
        if is_select:
            print(f"[watch] 状态字段是单选（选项值直接填文本即可，服务端自动匹配）")
    pending_filter = {"type": "group", "logic": "AND", "children": [
        {"type": "condition", "field": status_ctl, "operator": "eq",
         "value": [cfg["statusPending"]]}]}

    print(f"[watch] 开始轮询 {cfg['apiBase']} 每 {cfg['interval']}s，待打印值=「{cfg['statusPending']}」，打印机={cfg['printer'] or '默认'}")
    while True:
        try:
            data = md.rows_list(ws, pending_filter)
            for row in data.get("rows") or []:
                process_print_row(md, cfg, row, is_select, file_ctl, status_ctl)
        except Exception as e:
            msg = f"[watch] 轮询出错: {e}"
            if isinstance(e, urllib.error.HTTPError):
                msg += f"\n  url={e.url}\n  headers={dict(e.headers)}\n  body={e.read().decode('utf-8', 'replace')[:300]}"
            print(msg)
        time.sleep(cfg["interval"])


def process_print_row(md, cfg, row, is_select, file_ctl, status_ctl):
    row_id = row.get("rowid") or row.get("rowId")
    files = row.get(file_ctl)
    if not row_id:
        print(f"[watch] 跳过无 rowid 的行: {list(row)[:8]}")
        return
    if not isinstance(files, list) or not files:
        # 附件还没写入（异步延迟），下一轮再捞
        return
    # 先占位「打印中」：处理耗时可能超过轮询间隔，防重复出纸
    md.row_update(cfg["worksheetId"], row_id, status_ctl, cfg["statusPrinting"], is_select)
    try:
        with tempfile.TemporaryDirectory(prefix="mdwatch_") as td:
            td = Path(td)
            for i, f in enumerate(files, 1):
                name = f.get("fileName") or f.get("file_name") or f.get("originalFileName") or f.get("name") or f"file{i}"
                url = f.get("downloadUrl") or f.get("url")
                if not url:
                    raise RuntimeError(f"附件 {name} 缺少下载地址: {f}")
                ext = Path(name).suffix.lower() or ".bin"
                src = td / f"{i}{ext}"
                src.write_bytes(md.download(url))
                if ext == ".pdf":
                    pdf = src
                elif ext in OFFICE_EXTS:
                    pdf = td / f"{i}.pdf"
                    office_to_pdf(src, pdf)
                else:
                    print(f"[watch] 跳过不支持的附件类型: {name}")
                    continue
                with PRINT_LOCK:
                    print_pdf(pdf, cfg["printer"] or None, 1)
                print(f"[watch] 已打印 {name} (row {row_id[:8]}…)")
        md.row_update(cfg["worksheetId"], row_id, status_ctl, cfg["statusDone"], is_select)
    except Exception as e:
        print(f"[watch] 打印失败 (row {row_id[:8]}…): {e}")
        try:
            md.row_update(cfg["worksheetId"], row_id, status_ctl, cfg["statusFailed"], is_select)
        except Exception as e2:
            print(f"[watch] 回写失败状态也出错: {e2}")


def main():
    global SUMATRA
    if sys.stdout and hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    ap = argparse.ArgumentParser(description="明道云本地打印代理（watch 轮询模式）")
    ap.add_argument("--sumatra", default=None, help="SumatraPDF.exe 路径")
    ap.add_argument("--test", action="store_true", help="自检：打一张测试页到默认打印机后退出")
    ap.add_argument("--watch", action="store_true", help="默认即 watch 模式，此参数仅为兼容旧命令保留")
    ap.add_argument("--config", default=str(BASE_DIR / "config.json"),
                    help="配置文件路径（默认同目录 config.json）")
    args = ap.parse_args()

    SUMATRA = find_sumatra(args.sumatra)

    if args.test:
        print("自检：HTML → PDF → 默认打印机 …")
        do_print_job({"html": SAMPLE_HTML.replace("__TIME__", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))})
        print("[OK] 测试页已发送。若未出纸：1) 确认 SumatraPDF 已就位 2) 查看打印机队列是否暂停")
        return

    cfg = load_watch_config(args.config)
    printers = ", ".join(n + ("(默认)" if d else "") for n, d in list_printers()) or "(未枚举到)"
    print("=" * 56)
    print(f"  明道云本地打印代理 v{VERSION}")
    print(f"  SumatraPDF: {SUMATRA or '未找到（静默打印不可用，见文件头说明）'}")
    print(f"  打印机   : {printers}")
    print(f"  watch    : {cfg['apiBase']} worksheet={cfg['worksheetId'][:12]}… 打印机={cfg['printer'] or '默认'}")
    print(f"  自检     : python {Path(sys.argv[0]).name} --test")
    print("=" * 56)

    try:
        run_watch(cfg)
    except KeyboardInterrupt:
        print("\n已停止")


if __name__ == "__main__":
    main()
