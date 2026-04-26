#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import Dict, Optional


MACRO_MODULE_NAME = "docx_to_odt_macro.py"
MACRO_FUNCTION_URL = "vnd.sun.star.script:docx_to_odt_macro.py$run_job?language=Python&location=user"

LEVEL_MACRO_MODULE_NAME = "libreoffice_generic_levels.py"
LEVEL_MACRO_SETUP_URL = "vnd.sun.star.script:libreoffice_generic_levels.py$setup_keyboard_bindings?language=Python&location=user"


def configure_stdout() -> None:
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass


def print_line(tag: str, message: str) -> None:
    print(f"[{tag}] {message}")


def get_soffice() -> Path:
    candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for path in candidates:
        if path.exists():
            return path
    raise RuntimeError("找不到 soffice.exe，請確認 LibreOffice 已安裝")


def default_profile_base() -> Path:
    return Path(tempfile.gettempdir()) / "docx-to-odt-lo-profile"


def profile_uri(path: Path) -> str:
    return path.resolve().as_uri()


def macro_dir(profile_base: Path) -> Path:
    return profile_base / "user" / "Scripts" / "python"


def jobs_dir(profile_base: Path) -> Path:
    return profile_base / "jobs"


def install_macro(profile_base: Path) -> Path:
    target_dir = macro_dir(profile_base)
    target_dir.mkdir(parents=True, exist_ok=True)
    target_file = target_dir / MACRO_MODULE_NAME
    target_file.write_text(MACRO_SOURCE, encoding="utf-8")
    return target_file


def delete_if_exists(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except Exception:
        pass


def get_lo_user_scripts_dir() -> Optional[Path]:
    """取得 LibreOffice 使用者 Python 巨集目錄（My Macros）。"""
    # Windows
    appdata = os.environ.get("APPDATA")
    if appdata:
        candidate = Path(appdata) / "LibreOffice" / "4" / "user" / "Scripts" / "python"
        # 確認 LO user 目錄本身存在（即 LO 曾啟動過）
        if candidate.parent.parent.parent.exists():
            return candidate
    return None


def install_level_macros() -> None:
    """將 Tab 升降級巨集安裝到 LibreOffice My Macros，並設定鍵盤快捷鍵。"""
    print("=" * 60)

    # 步驟 1：複製巨集檔案
    lo_scripts_dir = get_lo_user_scripts_dir()
    if lo_scripts_dir is None:
        print_line("FAIL", "找不到 LibreOffice 使用者目錄，請確認 LibreOffice 已安裝並至少啟動過一次")
        raise SystemExit(1)

    lo_scripts_dir.mkdir(parents=True, exist_ok=True)
    target = lo_scripts_dir / LEVEL_MACRO_MODULE_NAME
    target.write_text(LEVEL_MACRO_SOURCE, encoding="utf-8")
    print_line("OK", f"巨集已安裝至：{target}")

    # 步驟 2：透過 LO UNO 設定鍵盤快捷鍵
    # 需要把巨集也安裝到轉檔用的 profile（讓 run_macro 可以呼叫）
    profile_base = default_profile_base()
    level_macro_in_profile = macro_dir(profile_base) / LEVEL_MACRO_MODULE_NAME
    level_macro_in_profile.parent.mkdir(parents=True, exist_ok=True)
    level_macro_in_profile.write_text(LEVEL_MACRO_SOURCE, encoding="utf-8")

    soffice = get_soffice()
    p_uri = profile_uri(profile_base)
    env = _get_clean_env()

    cmd = [
        str(soffice),
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        f"-env:UserInstallation={p_uri}",
        LEVEL_MACRO_SETUP_URL,
    ]

    proc = subprocess.Popen(
        cmd,
        env=env,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    try:
        stdout, stderr = proc.communicate(timeout=60)
    except subprocess.TimeoutExpired:
        proc.terminate()
        stdout, stderr = "", ""

    if proc.returncode == 0:
        print_line("OK", "Tab/Shift+Tab 快捷鍵已設定至 LibreOffice Writer")
        _update_config({"tab_macro_installed": True, "last_install_date": time.strftime("%Y-%m-%d")})
    else:
        print_line("WARNING", f"快捷鍵設定可能未完成（exit={proc.returncode}）")
        print_line("INFO", "你可以在 LibreOffice Writer 中手動執行『工具 → 巨集 → 執行巨集』")
        print_line("INFO", f"選擇 {LEVEL_MACRO_MODULE_NAME} 中的 setup_keyboard_bindings 函式")

    print("=" * 60)
    print_line("INFO", "安裝完成。請重新啟動 LibreOffice Writer 後，")
    print_line("INFO", "在含有『通用_層級1~4』樣式的 ODT 中即可使用 Tab/Shift+Tab 升降級。")
    print("=" * 60)


def _update_config(new_data: dict) -> None:
    """更新技能目錄下的 config.json。"""
    config_path = Path(__file__).parent.parent / "config.json"
    data = {}
    if config_path.exists():
        try:
            data = json.loads(config_path.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    
    data.update(new_data)
    try:
        config_path.write_text(json.dumps(data, ensure_ascii=False, indent=4), encoding="utf-8")
    except Exception as e:
        print_line("WARNING", f"無法更新設定檔：{e}")


def build_job(input_docx: Path, output_odt: Path, profile_base: Path) -> Dict[str, str]:
    jid = uuid.uuid4().hex
    jdir = jobs_dir(profile_base)
    jdir.mkdir(parents=True, exist_ok=True)

    staging_odt = jdir / f"{jid}.staging.odt"
    temp_output = output_odt.with_name(output_odt.stem + ".__macro_tmp__.odt")
    job_path = jdir / f"{jid}.job.json"
    status_path = jdir / f"{jid}.status.json"
    starter_path = jdir / f"{jid}.starter.txt"

    starter_path.write_text("starter", encoding="utf-8")

    payload = {
        "job_id": jid,
        "input_docx": str(input_docx.resolve()),
        "staging_odt": str(staging_odt.resolve()),
        "output_odt": str(output_odt.resolve()),
        "temp_output_odt": str(temp_output.resolve()),
        "status_json": str(status_path.resolve()),
    }
    job_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "job_id": jid,
        "job_path": str(job_path),
        "status_path": str(status_path),
        "starter_path": str(starter_path),
        "temp_output_odt": str(temp_output),
        "staging_odt": str(staging_odt),
    }


def _get_clean_env() -> Dict[str, str]:
    env = os.environ.copy()
    env.pop("PYTHONHOME", None)
    env.pop("PYTHONPATH", None)
    return env


def convert_docx_to_staging_odt(input_docx: Path, staging_odt: Path, profile_base: Path) -> None:
    soffice = get_soffice()
    staging_odt.parent.mkdir(parents=True, exist_ok=True)
    delete_if_exists(staging_odt)

    cmd = [
        str(soffice),
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        f"-env:UserInstallation={profile_uri(profile_base)}",
        "--convert-to",
        "odt",
        "--outdir",
        str(staging_odt.parent),
        str(input_docx),
    ]

    proc = subprocess.run(
        cmd,
        env=_get_clean_env(),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if proc.returncode != 0:
        raise RuntimeError(f"soffice CLI 轉 ODT 失敗：{proc.stderr.strip() or proc.stdout.strip()}")

    default_odt = staging_odt.parent / f"{input_docx.stem}.odt"

    if default_odt.exists() and default_odt.resolve() != staging_odt.resolve():
        if staging_odt.exists():
            staging_odt.unlink()
        default_odt.replace(staging_odt)

    if not staging_odt.exists():
        raise RuntimeError("soffice CLI 未產生 staging ODT")


def wait_for_status(status_path: Path, proc: subprocess.Popen, timeout: int) -> Optional[dict]:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if status_path.exists():
            try:
                return json.loads(status_path.read_text(encoding="utf-8"))
            except Exception:
                return {"status": "fail", "message": "status 檔存在但無法解析"}
        if proc.poll() is not None:
            break
        time.sleep(1)
    return None


def run_macro(profile_base: Path, starter_path: Path, job_path: Path, status_path: Path, timeout: int) -> dict:
    soffice = get_soffice()
    p_uri = profile_uri(profile_base)

    env = _get_clean_env()
    env["DOCX_TO_ODT_JOB"] = str(job_path.resolve())

    strategies = [
        [
            str(soffice),
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={p_uri}",
            MACRO_FUNCTION_URL,
        ],
        [
            str(soffice),
            "--headless",
            "--writer",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={p_uri}",
            MACRO_FUNCTION_URL,
        ],
        [
            str(soffice),
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={p_uri}",
            str(starter_path),
            MACRO_FUNCTION_URL,
        ],
    ]

    failures = []

    for idx, cmd in enumerate(strategies, start=1):
        delete_if_exists(status_path)

        proc = subprocess.Popen(
            cmd,
            env=env,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        status = wait_for_status(status_path, proc, timeout)
        if status is not None:
            try:
                proc.wait(timeout=5)
            except Exception:
                try:
                    proc.terminate()
                except Exception:
                    pass
            return status

        try:
            stdout, stderr = proc.communicate(timeout=5)
        except subprocess.TimeoutExpired:
            try:
                proc.terminate()
            except Exception:
                pass
            try:
                stdout, stderr = proc.communicate(timeout=3)
            except Exception:
                stdout, stderr = "", ""

        failures.append(
            f"strategy {idx} exit={proc.returncode} stdout={stdout.strip()} stderr={stderr.strip()}"
        )

    raise RuntimeError("無法成功啟動 LibreOffice macro；" + " | ".join(failures))


def cleanup_job_files(job: Dict[str, str]) -> None:
    for key in ("job_path", "status_path", "starter_path", "temp_output_odt", "staging_odt"):
        if key in job:
            delete_if_exists(Path(job[key]))


def print_success(status: dict, output_odt: Path) -> None:
    print("=" * 60)
    print_line("OK", "conversion completed")
    print_line("OK", f"output: {output_odt}")
    print_line(
        "OK",
        f"hanging punctuation: disabled (styles={status.get('styles_changed', 0)}, paragraphs={status.get('paragraphs_changed', 0)})",
    )
    print_line(
        "OK",
        f"line numbering: outside (method={status.get('line_numbering_method', 'unknown')})",
    )
    if status.get("line_numbering_detail"):
        print_line("OK", f"line numbering detail: {status['line_numbering_detail']}")
    if status.get("list_fix_count", 0) > 0:
        print_line("OK", f"list style binding: fixed ({status.get('list_fix_detail', '')})")
    elif status.get("list_fix_detail", "").startswith("skip"):
        print_line("INFO", f"list style binding: {status.get('list_fix_detail', '')}")
    print("=" * 60)


def main() -> None:
    configure_stdout()

    parser = argparse.ArgumentParser(description="DOCX -> staging ODT -> LibreOffice internal Python macro -> ODT")
    parser.add_argument("input", nargs="?", help="輸入 DOCX 路徑（使用 --install-macros 時可省略）")
    parser.add_argument("--output", "-o", help="輸出 ODT 路徑（預設：同名 .odt）")
    parser.add_argument("--timeout", type=int, default=180, help="等待 macro 完成秒數")
    parser.add_argument(
        "--profile-base",
        help="LibreOffice 專用 profile 根目錄（預設：系統 TEMP 下的 docx-to-odt-lo-profile）",
    )
    parser.add_argument(
        "--delete-input",
        action="store_true",
        help="轉換成功後刪除輸入的 DOCX 檔案（適用於 MD 轉 ODT 的中間暫存檔清理）",
    )
    parser.add_argument(
        "--install-macros",
        action="store_true",
        help="將 Tab 升降級巨集安裝至 LibreOffice My Macros，並設定 Tab/Shift+Tab 快捷鍵（不執行轉檔）",
    )
    args = parser.parse_args()

    # 獨立安裝模式
    if args.install_macros:
        install_level_macros()
        raise SystemExit(0)

    if not args.input:
        print_line("FAIL", "請指定輸入 DOCX 路徑，或使用 --install-macros 安裝升降級巨集")
        raise SystemExit(1)

    input_docx = Path(args.input).resolve()
    if not input_docx.exists():
        print_line("FAIL", f"找不到輸入檔：{input_docx}")
        raise SystemExit(1)

    output_odt = Path(args.output).resolve() if args.output else input_docx.with_suffix(".odt")
    profile_base = Path(args.profile_base).resolve() if args.profile_base else default_profile_base().resolve()

    job: Dict[str, str] = {}

    try:
        install_macro(profile_base)
        job = build_job(input_docx, output_odt, profile_base)

        convert_docx_to_staging_odt(
            input_docx=input_docx,
            staging_odt=Path(job["staging_odt"]),
            profile_base=profile_base,
        )

        status = run_macro(
            profile_base=profile_base,
            starter_path=Path(job["starter_path"]),
            job_path=Path(job["job_path"]),
            status_path=Path(job["status_path"]),
            timeout=args.timeout,
        )

        if status.get("status") != "ok":
            delete_if_exists(output_odt)
            delete_if_exists(Path(job["temp_output_odt"]))
            msg = status.get("message", "macro 執行失敗")
            print_line("FAIL", msg)
            raise SystemExit(1)

        if not output_odt.exists():
            print_line("FAIL", "macro 回報成功，但最終 ODT 不存在")
            raise SystemExit(1)

        print_success(status, output_odt)
        
        # 成功後依參數決定是否刪除輸入檔
        if args.delete_input:
            try:
                if input_docx.exists():
                    input_docx.unlink()
                    print_line("INFO", f"已刪除輸入檔：{input_docx}")
            except Exception as e:
                print_line("WARNING", f"無法刪除輸入檔：{e}")

    except SystemExit:
        raise
    except Exception as e:
        delete_if_exists(output_odt)
        if job.get("temp_output_odt"):
            delete_if_exists(Path(job["temp_output_odt"]))
        print_line("FAIL", str(e))
        raise SystemExit(1)
    finally:
        try:
            cleanup_job_files(job)
        except Exception:
            pass


LEVEL_MACRO_SOURCE = '''# -*- coding: utf-8 -*-
# libreoffice_generic_levels.py
# Tab/Shift+Tab 升降級巨集：在 通用_層級1~4 段落開頭，同時改段落樣式與清單層級。
# 安裝後需重新啟動 LibreOffice Writer 方能生效。

import uno

STYLE_ORDER = [
    "通用_層級1",
    "通用_層級2",
    "通用_層級3",
    "通用_層級4",
]

STYLE_TO_LEVEL = {name: idx for idx, name in enumerate(STYLE_ORDER)}
LEVEL_TO_STYLE = {idx: name for idx, name in enumerate(STYLE_ORDER)}
MIN_LEVEL = 0
MAX_LEVEL = len(STYLE_ORDER) - 1


def _get_document():
    doc = XSCRIPTCONTEXT.getDocument()
    if doc is None or not doc.supportsService("com.sun.star.text.TextDocument"):
        return None
    return doc


def _get_paragraph_styles(doc):
    return doc.StyleFamilies.getByName("ParagraphStyles")


def _get_list_style_name_for_style(doc, style_name):
    try:
        style = _get_paragraph_styles(doc).getByName(style_name)
        name = getattr(style, "NumberingStyleName", "")
        return str(name).strip() if name else ""
    except Exception:
        return ""


def _is_at_paragraph_start(doc):
    """判斷游標是否在段落開頭（未選取多段時）。"""
    try:
        controller = doc.getCurrentController()
        vc = controller.getViewCursor()
        # 若有選取範圍，不視為「在段落開頭」
        if not vc.isCollapsed():
            return False
        # 取得游標位置的文字游標，確認是否在段落起始
        text_cursor = doc.Text.createTextCursorByRange(vc.getStart())
        text_cursor.gotoStartOfParagraph(False)
        start_pos = text_cursor.getStart()
        cur_pos = vc.getStart()
        # 比較兩者文字偏移量
        return start_pos.getOffset() == cur_pos.getOffset()
    except Exception:
        return True  # 無法判斷時，保守地允許執行


def _get_current_para_style(doc):
    """取得游標所在段落的樣式名稱。"""
    try:
        controller = doc.getCurrentController()
        vc = controller.getViewCursor()
        return str(vc.ParaStyleName).strip()
    except Exception:
        return ""


def _set_paragraph_level_and_style(doc, para, new_level):
    """同時設定段落樣式名稱與清單層級。"""
    target_style = LEVEL_TO_STYLE[new_level]
    try:
        para.ParaStyleName = target_style
    except Exception:
        pass
    try:
        para.NumberingLevel = new_level
    except Exception:
        pass
    try:
        para.NumberingIsNumber = True
    except Exception:
        pass
    list_style_name = _get_list_style_name_for_style(doc, target_style)
    if list_style_name:
        try:
            para.NumberingStyleName = list_style_name
        except Exception:
            pass
    try:
        para.ParaIsNumberingRestart = False
    except Exception:
        pass


def _get_current_paragraph(doc):
    """取得游標所在段落物件。"""
    try:
        controller = doc.getCurrentController()
        vc = controller.getViewCursor()
        text_cursor = doc.Text.createTextCursorByRange(vc.getStart())
        text_cursor.gotoStartOfParagraph(False)
        text_cursor.gotoEndOfParagraph(True)
        enum = text_cursor.createEnumeration()
        while enum.hasMoreElements():
            elem = enum.nextElement()
            if elem.supportsService("com.sun.star.text.Paragraph"):
                return elem
    except Exception:
        pass
    return None


def _dispatch(doc, command, args=()):
    """執行 UNO dispatch 命令。"""
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()
        dispatcher = smgr.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", ctx
        )
        frame = doc.getCurrentController().getFrame()
        dispatcher.executeDispatch(frame, command, "", 0, args)
    except Exception:
        pass


def _insert_tab_char(doc):
    """插入 Tab 字元（與原生 Tab 在段落中間的行為一致）。"""
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()
        dispatcher = smgr.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", ctx
        )
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = "Text"
        prop.Value = "\\t"
        frame = doc.getCurrentController().getFrame()
        dispatcher.executeDispatch(frame, ".uno:InsertText", "", 0, (prop,))
    except Exception:
        pass


def tab_demote_or_default(*args):
    """Tab 鍵主函式：在 通用_層級1~4 段落開頭降一級；否則插入 Tab 字元。"""
    doc = _get_document()
    if doc is None:
        return

    style_name = _get_current_para_style(doc)

    if style_name in STYLE_TO_LEVEL and _is_at_paragraph_start(doc):
        # 在目標樣式段落開頭：降一級
        current_level = STYLE_TO_LEVEL[style_name]
        new_level = current_level + 1
        if new_level > MAX_LEVEL:
            return  # 已是最下層，靜默略過
        para = _get_current_paragraph(doc)
        if para is not None:
            _set_paragraph_level_and_style(doc, para, new_level)
    else:
        # 其他情況：插入 Tab 字元
        _insert_tab_char(doc)


def shift_tab_promote_or_default(*args):
    """Shift+Tab 鍵主函式：在 通用_層級1~4 段落開頭升一級；否則執行原生 Shift+Tab。"""
    doc = _get_document()
    if doc is None:
        return

    style_name = _get_current_para_style(doc)

    if style_name in STYLE_TO_LEVEL and _is_at_paragraph_start(doc):
        # 在目標樣式段落開頭：升一級
        current_level = STYLE_TO_LEVEL[style_name]
        new_level = current_level - 1
        if new_level < MIN_LEVEL:
            return  # 已是最上層，靜默略過
        para = _get_current_paragraph(doc)
        if para is not None:
            _set_paragraph_level_and_style(doc, para, new_level)
    else:
        # 其他情況：執行原生 Shift+Tab（減少縮排）
        _dispatch(doc, ".uno:LeftPara")


def repair_generic_level_by_style(*args):
    """依目前段落樣式名稱，重新校正清單層級（修復工具）。"""
    doc = _get_document()
    if doc is None:
        return
    style_name = _get_current_para_style(doc)
    if style_name not in STYLE_TO_LEVEL:
        return
    level = STYLE_TO_LEVEL[style_name]
    para = _get_current_paragraph(doc)
    if para is not None:
        _set_paragraph_level_and_style(doc, para, level)


def setup_keyboard_bindings(*args):
    """透過 UNO API 設定 Tab/Shift+Tab 快捷鍵至本巨集的兩個函式。"""
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()

        supplier = smgr.createInstanceWithContext(
            "com.sun.star.ui.ModuleUIConfigurationManagerSupplier", ctx
        )
        cfg_mgr = supplier.getUIConfigurationManager("com.sun.star.text.TextDocument")
        accel_cfg = cfg_mgr.getShortCutManager()

        # TAB key: KeyCode=0x0009, Modifiers=0
        # Shift+TAB: KeyCode=0x0009, Modifiers=1 (SHIFT)
        tab_event = uno.createUnoStruct("com.sun.star.awt.KeyEvent")
        tab_event.KeyCode = 0x0009
        tab_event.Modifiers = 0

        shift_tab_event = uno.createUnoStruct("com.sun.star.awt.KeyEvent")
        shift_tab_event.KeyCode = 0x0009
        shift_tab_event.Modifiers = 1

        accel_cfg.setKeyEvent(
            tab_event,
            "macro:///libreoffice_generic_levels.tab_demote_or_default"
        )
        accel_cfg.setKeyEvent(
            shift_tab_event,
            "macro:///libreoffice_generic_levels.shift_tab_promote_or_default"
        )

        cfg_mgr.store()
    except Exception:
        pass


g_exportedScripts = (
    tab_demote_or_default,
    shift_tab_promote_or_default,
    repair_generic_level_by_style,
    setup_keyboard_bindings,
)
'''


MACRO_SOURCE = r'''# -*- coding: utf-8 -*-

import json
import os
import re
import shutil
import traceback
import zipfile
from pathlib import Path

import uno


def _write_status(status_path: Path, payload: dict) -> None:
    status_path.parent.mkdir(parents=True, exist_ok=True)
    status_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _delete_if_exists(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except Exception:
        pass


def _make_prop(name, value):
    prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    prop.Name = name
    prop.Value = value
    return prop


def _file_url(path: Path) -> str:
    return uno.systemPathToFileUrl(str(path.resolve()))


def _load_doc(desktop, path: Path):
    url = _file_url(path)
    strategies = [
        (),
        (_make_prop("Hidden", True),),
        (_make_prop("Hidden", True), _make_prop("ReadOnly", False)),
    ]
    last_error = None
    for props in strategies:
        try:
            doc = desktop.loadComponentFromURL(url, "_default", 0, props)
            if doc is not None:
                return doc
        except Exception as e:
            last_error = e
    if last_error is not None:
        raise RuntimeError(f"LibreOffice 無法載入文件：{last_error}")
    raise RuntimeError("LibreOffice 無法載入文件：loadComponentFromURL 回傳 None")


def _close_doc(doc) -> None:
    try:
        doc.close(True)
        return
    except Exception:
        pass
    try:
        doc.dispose()
    except Exception:
        pass


def _iter_paragraph_styles(styles):
    try:
        for name in styles.getElementNames():
            try:
                yield str(name), styles.getByName(name)
            except Exception:
                continue
        return
    except Exception:
        pass

    count = 0
    try:
        count = styles.getCount()
    except Exception:
        try:
            count = styles.Count
        except Exception:
            count = 0

    for i in range(count):
        try:
            yield str(i), styles.getByIndex(i)
        except Exception:
            continue


def _try_set_hanging_false(obj) -> bool:
    try:
        obj.ParaIsHangingPunctuation = False
        return True
    except Exception:
        pass
    try:
        obj.setPropertyValue("ParaIsHangingPunctuation", False)
        return True
    except Exception:
        return False


def _disable_hanging_punctuation(doc):
    styles_changed = 0
    paragraphs_changed = 0

    styles = doc.StyleFamilies.getByName("ParagraphStyles")
    for _, style in _iter_paragraph_styles(styles):
        if _try_set_hanging_false(style):
            styles_changed += 1

    enum = doc.Text.createEnumeration()
    while enum.hasMoreElements():
        elem = enum.nextElement()
        if _try_set_hanging_false(elem):
            paragraphs_changed += 1

    # ── 修正：禁選頁首/頁尾區域的行編號 ──
    try:
        page_styles = doc.StyleFamilies.getByName("PageStyles")
        for i in range(page_styles.getCount()):
            page_style = page_styles.getByIndex(i)
            # 1. 頁尾
            if hasattr(page_style, "FooterIsOn") and page_style.FooterIsOn:
                footer_text = page_style.FooterText
                if footer_text:
                    f_enum = footer_text.createEnumeration()
                    while f_enum.hasMoreElements():
                        f_par = f_enum.nextElement()
                        try:
                            f_par.ParaLineNumberCount = False
                        except Exception:
                            pass
            # 2. 頁首
            if hasattr(page_style, "HeaderIsOn") and page_style.HeaderIsOn:
                header_text = page_style.HeaderText
                if header_text:
                    h_enum = header_text.createEnumeration()
                    while h_enum.hasMoreElements():
                        h_par = h_enum.nextElement()
                        try:
                            h_par.ParaLineNumberCount = False
                        except Exception:
                            pass
    except Exception:
        pass

    return styles_changed, paragraphs_changed


def _fix_list_style_bindings(doc):
    """
    修復通用_層級1~4 的清單樣式綁定。
    將所有通用_層級1~4 綁定至 WWNum1，並處理正確的重設編號。
    """
    try:
        para_styles = doc.StyleFamilies.getByName("ParagraphStyles")
    except Exception:
        return 0, "skip: 無法取得 ParagraphStyles"

    # 1. 找父樣式的清單樣式名稱
    target_ls = None
    parent_name = '通用多層清單'
    if para_styles.hasByName(parent_name):
        try:
            target_ls = para_styles.getByName(parent_name).NumberingStyleName
        except Exception:
            pass

    if not target_ls:
        return 0, "skip: 找不到通用多層清單或其清單樣式"

    # 2. 設定子樣式的 NumberingStyleName
    display_to_level = {}
    internal_to_level = {}
    styles_fixed = 0

    for suffix, level in [('1', 0), ('2', 1), ('3', 2), ('4', 3)]:
        for variant in [f'通用_5f_層級{suffix}', f'通用_層級{suffix}']:
            if para_styles.hasByName(variant):
                try:
                    s = para_styles.getByName(variant)
                    s.NumberingStyleName = target_ls
                    dn = s.DisplayName if hasattr(s, 'DisplayName') else variant
                    display_to_level[dn] = level
                    internal_to_level[variant] = level
                    styles_fixed += 1
                except Exception:
                    pass
                break

    if styles_fixed == 0:
        return 0, "skip: 找不到通用_層級1~4"

    all_level_map = {}
    all_level_map.update(display_to_level)
    all_level_map.update(internal_to_level)

    # 3. 遍歷段落，設定 NumberingLevel 與 restart
    paras_fixed = 0
    restart_next = False

    enum = doc.Text.createEnumeration()
    while enum.hasMoreElements():
        elem = enum.nextElement()
        try:
            if elem.supportsService("com.sun.star.text.Paragraph"):
                text = elem.String.strip()
                if text in ["事實與理由", "理由", "訴之聲明"]:
                    restart_next = True
                else:
                    sn = elem.ParaStyleName
                    if sn in all_level_map:
                        level = all_level_map[sn]
                        elem.NumberingLevel = level
                        elem.NumberingIsNumber = True
                        if level == 0 and restart_next:
                            elem.ParaIsNumberingRestart = True
                            elem.NumberingStartValue = 1
                            restart_next = False
                        elif level == 0:
                            # Level 0 且沒遇到標題，確保不要重設
                            elem.ParaIsNumberingRestart = False
                            restart_next = False
                        elif level > 0:
                            # 內部層級，取消重設旗標
                            restart_next = False
                        paras_fixed += 1
                    else:
                        pass
        except Exception:
            pass

    total = styles_fixed + paras_fixed
    detail = f"styles={styles_fixed}, paragraphs={paras_fixed}, list_style={target_ls}"
    return total, detail


def _verify_hanging_disabled(doc):
    checked = 0
    still_true = 0

    try:
        styles = doc.StyleFamilies.getByName("ParagraphStyles")
        for _, style in _iter_paragraph_styles(styles):
            try:
                checked += 1
                if bool(style.ParaIsHangingPunctuation):
                    still_true += 1
            except Exception:
                pass
    except Exception:
        pass

    try:
        enum = doc.Text.createEnumeration()
        while enum.hasMoreElements():
            elem = enum.nextElement()
            try:
                checked += 1
                if bool(elem.ParaIsHangingPunctuation):
                    still_true += 1
            except Exception:
                pass
    except Exception:
        pass

    return checked, still_true


def _get_property_names(obj):
    names = set()
    try:
        info = obj.getPropertySetInfo()
        for p in info.getProperties():
            names.add(p.Name)
    except Exception:
        pass
    return names


def _try_set_line_numbering_uno(doc):
    candidates = []
    for attr_name in ("LineNumberingRules", "LineNumberingProperties"):
        try:
            obj = getattr(doc, attr_name)
            if obj is not None:
                candidates.append((attr_name, obj))
        except Exception:
            pass

    if not candidates:
        return False, "找不到 UNO 行編號設定物件"

    for attr_name, rules in candidates:
        prop_names = _get_property_names(rules)
        if not prop_names:
            continue

        try:
            if "IsOn" in prop_names:
                is_on = rules.getPropertyValue("IsOn")
                if not is_on:
                    rules.setPropertyValue("IsOn", True)
        except Exception:
            pass

        attempts = [
            ("Position", 3),
            ("NumberPosition", 3),
            ("Position", "OUTSIDE"),
            ("NumberPosition", "OUTSIDE"),
            ("Position", "outside"),
            ("NumberPosition", "outside"),
        ]

        for prop_name, value in attempts:
            if prop_name not in prop_names:
                continue
            try:
                rules.setPropertyValue(prop_name, value)
                try:
                    setattr(doc, attr_name, rules)
                except Exception:
                    pass
                return True, f"UNO 已設定 {prop_name}={value}"
            except Exception:
                continue

        return False, f"UNO 可存取行編號，但無法設定外側；可用屬性：{sorted(prop_names)}"

    return False, "未找到可寫入的 UNO 行編號位置屬性"


def _save_as_odt(doc, output_path: Path) -> None:
    props = (
        _make_prop("FilterName", "writer8"),
        _make_prop("Overwrite", True),
    )
    doc.storeAsURL(_file_url(output_path), props)


def _has_outside_line_numbering(text: str) -> bool:
    patterns = [
        r'<text:linenumbering-configuration\b[^>]*\btext:number-position="outside"',
        r'<text:linenumbering-configuration\b[^>]*\btext:position="outside"',
    ]
    for pattern in patterns:
        if re.search(pattern, text, flags=re.IGNORECASE):
            return True
    return False


def _normalize_line_numbering_tag(tag: str) -> str:
    tag2 = re.sub(r'\s+text:(?:number-)?position="[^"]*"', "", tag, flags=re.IGNORECASE)
    if tag2.endswith("/>"):
        return tag2[:-2].rstrip() + ' text:number-position="outside"/>'
    if tag2.endswith(">"):
        return tag2[:-1].rstrip() + ' text:number-position="outside">'
    return tag2


def _enforce_line_numbering_outside_xml(odt_path: Path):
    ln_config = (
        '<text:linenumbering-configuration '
        'text:number-lines="true" '
        'text:offset="1cm" '
        'style:num-format="1" '
        'text:count-empty-lines="true" '
        'text:count-in-floating-frames="false" '
        'text:restart-on-page="true" '
        'text:number-position="outside" '
        'text:increment="1"/>'
    )

    tmp_path = odt_path.with_suffix(".xmlfix.tmp")
    _delete_if_exists(tmp_path)
    odt_path.replace(tmp_path)

    changed = False
    verified = False
    found_styles = False

    try:
        with zipfile.ZipFile(tmp_path, "r") as zin, zipfile.ZipFile(odt_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename in ["styles.xml", "content.xml"]:
                    text = data.decode("utf-8", errors="replace")
                    # 修正懸尾標點符號
                    orig_text = text
                    text = re.sub(r'style:punctuation-wrap="hanging"', 'style:punctuation-wrap="simple"', text)
                    text = re.sub(r'style:hanging-punctuation="true"', 'style:hanging-punctuation="false"', text)
                    if text != orig_text:
                        changed = True

                    if item.filename == "styles.xml":
                        found_styles = True
                        if _has_outside_line_numbering(text):
                            verified = True
                        else:
                            if re.search(r"<text:linenumbering-configuration\b", text, flags=re.IGNORECASE):
                            new_text = re.sub(
                                r"<text:linenumbering-configuration\b[^>]*?/?>",
                                lambda m: _normalize_line_numbering_tag(m.group(0)),
                                text,
                                count=1,
                                flags=re.IGNORECASE,
                            )
                            if new_text != text:
                                text = new_text
                                changed = True
                        else:
                            inserted = False
                            for marker in ("</office:styles>", "</office:automatic-styles>", "</office:document-styles>"):
                                if marker in text:
                                    text = text.replace(marker, ln_config + "\n" + marker, 1)
                                    changed = True
                                    inserted = True
                                    break
                            if not inserted:
                                raise RuntimeError("styles.xml 中找不到可插入行編號設定的位置")

                            verified = _has_outside_line_numbering(text)
                    
                    data = text.encode("utf-8")

                zout.writestr(item, data)

        if not found_styles:
            raise RuntimeError("ODT 中找不到 styles.xml")
        if not verified:
            return False, "無法驗證最終 ODT 的行編號位置已為 outside"
        if changed:
            return True, "ODT XML fallback"
        return True, "UNO"

    except Exception:
        if odt_path.exists():
            try:
                odt_path.unlink()
            except Exception:
                pass
        if tmp_path.exists():
            tmp_path.replace(odt_path)
        raise
    finally:
        _delete_if_exists(tmp_path)


def _unify_list_style_xml(odt_path: Path):
    """
    XML 後處理：移除 content.xml 中自動段落樣式的 style:list-style-name 覆寫。
    只處理 parent 為 通用_層級1~4 的自動樣式，讓段落回歸繼承段落樣式的清單綁定。
    使用正則表達式避免依賴 lxml（LibreOffice 內建 Python 不保證有）。
    """
    target_parents = (
        '\u901a\u7528_5f_\u5c64\u7d1a1', '\u901a\u7528_5f_\u5c64\u7d1a2',
        '\u901a\u7528_5f_\u5c64\u7d1a3', '\u901a\u7528_5f_\u5c64\u7d1a4',
        '\u901a\u7528_\u5c64\u7d1a1', '\u901a\u7528_\u5c64\u7d1a2',
        '\u901a\u7528_\u5c64\u7d1a3', '\u901a\u7528_\u5c64\u7d1a4',
    )

    tmp_path = odt_path.with_suffix('.listfix.tmp')
    _delete_if_exists(tmp_path)
    odt_path.replace(tmp_path)

    cleaned = 0
    try:
        with zipfile.ZipFile(tmp_path, 'r') as zin, \
             zipfile.ZipFile(odt_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'content.xml':
                    text = data.decode('utf-8', errors='replace')
                    def _strip_list_attr(m):
                        nonlocal cleaned
                        tag = m.group(0)
                        parent_match = re.search(r'style:parent-style-name="([^"]*)"', tag)
                        if not parent_match:
                            return tag
                        parent = parent_match.group(1)
                        if parent not in target_parents:
                            return tag
                        new_tag = re.sub(r'\s+style:list-style-name="[^"]*"', '', tag)
                        if new_tag != tag:
                            cleaned += 1
                        return new_tag

                    text = re.sub(
                        r'<style:style\b[^>]*style:family="paragraph"[^>]*>',
                        _strip_list_attr,
                        text,
                    )
                    text = re.sub(
                        r'<style:style\b[^>]*style:parent-style-name="[^"]*"[^>]*style:family="paragraph"[^>]*>',
                        _strip_list_attr,
                        text,
                    )
                    data = text.encode('utf-8')
                zout.writestr(item, data)
        return cleaned
    except Exception:
        if odt_path.exists():
            try:
                odt_path.unlink()
            except Exception:
                pass
        if tmp_path.exists():
            tmp_path.replace(odt_path)
        raise
    finally:
        _delete_if_exists(tmp_path)


def _desktop():
    return XSCRIPTCONTEXT.getDesktop()


def run_job(*args):
    desktop = _desktop()
    status_path = None
    doc = None
    verify_doc = None

    try:
        job_env = os.environ.get("DOCX_TO_ODT_JOB", "").strip()
        if not job_env:
            raise RuntimeError("找不到 DOCX_TO_ODT_JOB 環境變數")

        job_path = Path(job_env)
        payload = json.loads(job_path.read_text(encoding="utf-8"))

        input_docx = Path(payload["input_docx"]).resolve()
        staging_odt = Path(payload["staging_odt"]).resolve()
        output_odt = Path(payload["output_odt"]).resolve()
        temp_output = Path(payload["temp_output_odt"]).resolve()
        status_path = Path(payload["status_json"]).resolve()

        _delete_if_exists(output_odt)
        _delete_if_exists(temp_output)

        if not staging_odt.exists():
            raise RuntimeError(f"找不到 staging ODT：{staging_odt}")

        doc = _load_doc(desktop, staging_odt)

        styles_changed, paragraphs_changed = _disable_hanging_punctuation(doc)
        if styles_changed == 0 and paragraphs_changed == 0:
            raise RuntimeError("懸尾修正失敗：未成功修改任何段落樣式或段落")

        list_fix_count, list_fix_detail = _fix_list_style_bindings(doc)

        ln_ok, ln_detail = _try_set_line_numbering_uno(doc)

        temp_output.parent.mkdir(parents=True, exist_ok=True)
        _save_as_odt(doc, temp_output)
        _close_doc(doc)
        doc = None

        verify_doc = _load_doc(desktop, temp_output)
        checked, still_true = _verify_hanging_disabled(verify_doc)
        _close_doc(verify_doc)
        verify_doc = None

        if checked == 0:
            raise RuntimeError("懸尾驗證失敗：重新開啟 ODT 後，無法讀取任何可驗證屬性")
        if still_true > 0:
            raise RuntimeError(f"懸尾驗證失敗：重新開啟 ODT 後，仍有 {still_true} 個樣式或段落維持懸尾")

        xml_ok, xml_method = _enforce_line_numbering_outside_xml(temp_output)
        if not xml_ok:
            raise RuntimeError("行編號修正失敗：最終 ODT 無法驗證為 outside")

        list_xml_cleaned = _unify_list_style_xml(temp_output)

        output_odt.parent.mkdir(parents=True, exist_ok=True)
        if output_odt.exists():
            output_odt.unlink()
        shutil.move(str(temp_output), str(output_odt))

        line_method = "UNO" if (ln_ok and xml_method == "UNO") else (
            "UNO + ODT XML fallback" if ln_ok else "ODT XML fallback"
        )

        _write_status(status_path, {
            "status": "ok",
            "styles_changed": styles_changed,
            "paragraphs_changed": paragraphs_changed,
            "line_numbering_method": line_method,
            "line_numbering_detail": ln_detail,
            "list_fix_count": list_fix_count,
            "list_fix_detail": list_fix_detail,
            "list_xml_cleaned": list_xml_cleaned,
            "output": str(output_odt),
            "input_docx": str(input_docx),
            "staging_odt": str(staging_odt),
        })

    except Exception as e:
        try:
            if status_path is not None:
                _write_status(status_path, {
                    "status": "fail",
                    "message": str(e),
                    "traceback": traceback.format_exc(),
                })
        except Exception:
            pass
        try:
            if "output_odt" in locals():
                _delete_if_exists(output_odt)
            if "temp_output" in locals():
                _delete_if_exists(temp_output)
        except Exception:
            pass
    finally:
        if doc is not None:
            _close_doc(doc)
        if verify_doc is not None:
            _close_doc(verify_doc)
        try:
            desktop.terminate()
        except Exception:
            pass


g_exportedScripts = (run_job,)
'''

if __name__ == "__main__":
    main()