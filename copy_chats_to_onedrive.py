# -*- coding: utf-8 -*-
"""
نقل/نسخ المحادثات إلى مجلد OneDrive.

ينسخ نسخة آمنة (لا يحذف الأصل) من:
- محادثات Cursor (قواعد البيانات + تصدير Markdown مقروء)
- واتساب / تليغرام / تيمز / ربط الهاتف إن وُجدت
- ملفات تصدير المحادثات على سطح المكتب والتحميلات
- أرشيف المحادثات داخل هذا المشروع

الاستخدام على ويندوز:
    python copy_chats_to_onedrive.py
    python copy_chats_to_onedrive.py --dry-run
    python copy_chats_to_onedrive.py --dest "D:\\تجربة"

هذه البيئة السحابية لا تصل إلى OneDrive على جهازك؛ شغّل الملف على ويندوز.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sqlite3
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple


SKIP_DIR_NAMES = {
    "cache",
    "cacheddata",
    "code cache",
    "gpucache",
    "service worker",
    "crashpad",
    "logs",
    "log",
    "cachedextensions",
    "cachedprofilesdata",
    "videohtml",
    "shadercache",
}

CHAT_FILE_RE = re.compile(
    r"(chat|whatsapp|telegram|teams|محادث|دردش|واتس|تليجرام|تليغرام)",
    re.IGNORECASE,
)


def eprint(*args: object) -> None:
    print(*args, file=sys.stderr)


def timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H%M")


def safe_name(text: str, fallback: str = "محادثة") -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", (text or "").strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ._")
    return (cleaned[:80] or fallback)


def folder_size_bytes(path: Path) -> int:
    total = 0
    if not path.exists():
        return 0
    if path.is_file():
        try:
            return path.stat().st_size
        except OSError:
            return 0
    for root, dirs, files in os.walk(path):
        dirs[:] = [d for d in dirs if d.lower() not in SKIP_DIR_NAMES]
        for name in files:
            try:
                total += (Path(root) / name).stat().st_size
            except OSError:
                continue
    return total


def format_size(num: int) -> str:
    value = float(num)
    for unit in ("بايت", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            if unit == "بايت":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024.0
    return f"{num} بايت"


def detect_onedrive(env: Optional[Dict[str, str]] = None, home: Optional[Path] = None) -> Optional[Path]:
    env = env if env is not None else dict(os.environ)
    home = home or Path.home()
    for key in ("OneDrive", "OneDriveConsumer", "OneDriveCommercial"):
        value = env.get(key)
        if value:
            path = Path(value)
            if path.exists():
                return path
    candidates = [
        home / "OneDrive",
        home / "OneDrive - Personal",
    ]
    for path in candidates:
        if path.exists():
            return path
    if home.exists():
        for child in home.iterdir():
            if child.is_dir() and child.name.lower().startswith("onedrive"):
                return child
    return None


def windows_appdata(env: Optional[Dict[str, str]] = None) -> Tuple[Path, Path, Path]:
    env = env if env is not None else dict(os.environ)
    home = Path(env.get("USERPROFILE") or Path.home())
    roaming = Path(env.get("APPDATA") or (home / "AppData" / "Roaming"))
    local = Path(env.get("LOCALAPPDATA") or (home / "AppData" / "Local"))
    return home, roaming, local


def glob_first(parent: Path, pattern: str) -> List[Path]:
    if not parent.exists():
        return []
    return sorted(parent.glob(pattern))


def discover_sources(
    env: Optional[Dict[str, str]] = None,
    extra_sources: Optional[Sequence[Path]] = None,
    project_dir: Optional[Path] = None,
) -> List[Dict[str, Any]]:
    env = env if env is not None else dict(os.environ)
    home, roaming, local = windows_appdata(env)
    sources: List[Dict[str, Any]] = []

    def add(key: str, title: str, path: Path, kind: str, note: str = "") -> None:
        if not path.exists():
            return
        sources.append(
            {
                "key": key,
                "title": title,
                "path": path,
                "kind": kind,
                "note": note,
                "size": folder_size_bytes(path),
            }
        )

    cursor_user = roaming / "Cursor" / "User"
    add("cursor_global", "Cursor — قاعدة المحادثات العامة", cursor_user / "globalStorage", "cursor", "state.vscdb")
    add("cursor_workspaces", "Cursor — محادثات المشاريع", cursor_user / "workspaceStorage", "cursor", "")
    cursor_home = home / ".cursor"
    if cursor_home.exists():
        for name in ("chats", "ai-tracking", "conversations"):
            p = cursor_home / name
            if p.exists():
                add(f"cursor_dot_{name.replace('-', '_')}", f"Cursor — {name}", p, "cursor", "")
        loose = [
            p
            for p in cursor_home.iterdir()
            if p.is_file() and CHAT_FILE_RE.search(p.name)
        ]
        if loose:
            sources.append(
                {
                    "key": "cursor_dot_files",
                    "title": "Cursor — ملفات محادثات",
                    "path": cursor_home,
                    "kind": "loose_files",
                    "note": "",
                    "size": sum(folder_size_bytes(p) for p in loose),
                    "files": loose,
                }
            )

    add("whatsapp_roaming", "واتساب (سطح المكتب)", roaming / "WhatsApp", "app", "")
    add("whatsapp_local", "واتساب (Local)", local / "WhatsApp", "app", "")
    for pkg in glob_first(local / "Packages", "5319275A.WhatsAppDesktop_*"):
        add("whatsapp_store", "واتساب من المتجر", pkg / "LocalState", "app", "")

    add("telegram", "تليغرام", roaming / "Telegram Desktop" / "tdata", "app", "")

    add("teams_classic", "تيمز (الإصدار القديم)", roaming / "Microsoft" / "Teams", "app", "")
    for pkg in glob_first(local / "Packages", "MSTeams_*"):
        add("teams_new", "تيمز (الجديد)", pkg / "LocalCache", "app", "")

    for pkg in glob_first(local / "Packages", "Microsoft.YourPhone_*"):
        add("phone_link", "ربط الهاتف / رسائل الجوال", pkg / "LocalState", "app", "")

    for folder_name, title in (("Desktop", "سطح المكتب"), ("Downloads", "التحميلات")):
        folder = home / folder_name
        if folder.exists():
            matches = [
                p
                for p in folder.iterdir()
                if p.is_file() and CHAT_FILE_RE.search(p.name)
            ]
            if matches:
                sources.append(
                    {
                        "key": f"loose_{folder_name.lower()}",
                        "title": f"ملفات محادثات من {title}",
                        "path": folder,
                        "kind": "loose_files",
                        "note": "",
                        "size": sum(folder_size_bytes(p) for p in matches),
                        "files": matches,
                    }
                )

    project_dir = project_dir or Path(__file__).resolve().parent
    archive = project_dir / "conversations"
    add("project_archive", "أرشيف محادثات هذا المشروع", archive, "archive", "")

    if extra_sources:
        for i, extra in enumerate(extra_sources, start=1):
            add(f"extra_{i}", f"مجلد إضافي {i}", Path(extra), "extra", "")

    return sources


def should_skip_dir(name: str) -> bool:
    return name.lower() in SKIP_DIR_NAMES


def copy_path(src: Path, dest: Path, dry_run: bool = False) -> Tuple[int, int]:
    """Copy file or folder. Returns (files_copied, bytes_copied)."""
    files = 0
    nbytes = 0
    if src.is_file():
        size = src.stat().st_size
        if not dry_run:
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
        return 1, size

    for root, dirs, filenames in os.walk(src):
        dirs[:] = [d for d in dirs if not should_skip_dir(d)]
        rel = Path(root).relative_to(src)
        target_root = dest / rel
        if not dry_run:
            target_root.mkdir(parents=True, exist_ok=True)
        for name in filenames:
            sfile = Path(root) / name
            dfile = target_root / name
            try:
                size = sfile.stat().st_size
            except OSError:
                continue
            if not dry_run:
                try:
                    shutil.copy2(sfile, dfile)
                except OSError:
                    continue
            files += 1
            nbytes += size
    return files, nbytes


def copy_source(source: Dict[str, Any], dest_root: Path, dry_run: bool = False) -> Tuple[int, int]:
    folder_name = safe_name(source["title"], source["key"])
    dest = dest_root / folder_name
    if source["kind"] == "loose_files":
        files = 0
        nbytes = 0
        if not dry_run:
            dest.mkdir(parents=True, exist_ok=True)
        for item in source.get("files") or []:
            f, n = copy_path(Path(item), dest / Path(item).name, dry_run=dry_run)
            files += f
            nbytes += n
        return files, nbytes

    if source["kind"] == "cursor":
        src = Path(source["path"])
        if src.name.lower() == "globalstorage":
            files = 0
            nbytes = 0
            if not dry_run:
                dest.mkdir(parents=True, exist_ok=True)
            for name in ("state.vscdb", "state.vscdb.backup"):
                candidate = src / name
                if candidate.exists():
                    f, n = copy_path(candidate, dest / name, dry_run=dry_run)
                    files += f
                    nbytes += n
            return files, nbytes
        if src.name.lower() == "workspacestorage":
            files = 0
            nbytes = 0
            if not src.exists():
                return 0, 0
            for ws in src.iterdir():
                if not ws.is_dir():
                    continue
                ws_dest = dest / ws.name
                for name in ("state.vscdb", "state.vscdb.backup", "workspace.json"):
                    candidate = ws / name
                    if candidate.exists():
                        f, n = copy_path(candidate, ws_dest / name, dry_run=dry_run)
                        files += f
                        nbytes += n
            return files, nbytes

    return copy_path(Path(source["path"]), dest, dry_run=dry_run)


def _json_load(raw: Any) -> Any:
    if raw is None:
        return None
    if isinstance(raw, bytes):
        try:
            raw = raw.decode("utf-8")
        except UnicodeDecodeError:
            raw = raw.decode("utf-8", errors="replace")
    if isinstance(raw, str):
        raw = raw.strip()
        if not raw:
            return None
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return None
    return raw


def _bubble_text(payload: Any) -> str:
    if not isinstance(payload, dict):
        return ""
    for key in ("text", "richText", "content"):
        value = payload.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
        if isinstance(value, dict):
            inner = value.get("text") or value.get("value")
            if isinstance(inner, str) and inner.strip():
                return inner.strip()
    return ""


def _bubble_role(payload: Any) -> str:
    if not isinstance(payload, dict):
        return "رسالة"
    type_val = payload.get("type")
    if type_val in (1, "1", "user", "human"):
        return "المستخدم"
    if type_val in (2, "2", "assistant", "ai"):
        return "المساعد"
    role = str(payload.get("role") or "").lower()
    if role in ("user", "human"):
        return "المستخدم"
    if role in ("assistant", "ai", "model"):
        return "المساعد"
    return "رسالة"


def open_sqlite_readonly(path: Path) -> Optional[sqlite3.Connection]:
    if not path.exists():
        return None
    uri = path.resolve().as_uri() + "?mode=ro"
    try:
        conn = sqlite3.connect(uri, uri=True)
        conn.row_factory = sqlite3.Row
        return conn
    except sqlite3.Error:
        try:
            conn = sqlite3.connect(str(path))
            conn.row_factory = sqlite3.Row
            return conn
        except sqlite3.Error:
            return None


def kv_rows(conn: sqlite3.Connection) -> Iterable[Tuple[str, Any]]:
    tables = [
        row[0]
        for row in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()
    ]
    for table in tables:
        cols = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
        col_l = [c.lower() for c in cols]
        if "key" in col_l and "value" in col_l:
            key_col = cols[col_l.index("key")]
            val_col = cols[col_l.index("value")]
            try:
                for key, value in conn.execute(
                    f'SELECT "{key_col}", "{val_col}" FROM "{table}"'
                ):
                    yield str(key), value
            except sqlite3.Error:
                continue


def export_cursor_db(db_path: Path, out_dir: Path, dry_run: bool = False) -> int:
    conn = open_sqlite_readonly(db_path)
    if conn is None:
        return 0
    exported = 0
    composers: Dict[str, Dict[str, Any]] = {}
    bubbles: Dict[str, List[Tuple[str, Dict[str, Any]]]] = {}
    try:
        for key, value in kv_rows(conn):
            data = _json_load(value)
            if key == "composer.composerData" and isinstance(data, dict):
                for item in data.get("allComposers") or []:
                    cid = str(item.get("composerId") or "")
                    if cid:
                        composers.setdefault(cid, {}).update(item)
            elif key == "composer.composerHeaders" and isinstance(data, dict):
                items = data.get("allComposers") or data.get("headers") or data.get("composers") or []
                if isinstance(data, list):
                    items = data
                for item in items:
                    if not isinstance(item, dict):
                        continue
                    cid = str(item.get("composerId") or item.get("id") or "")
                    if cid:
                        composers.setdefault(cid, {}).update(item)
            elif key.startswith("composerData:"):
                cid = key.split(":", 1)[1]
                if isinstance(data, dict):
                    composers.setdefault(cid, {}).update(data)
                    composers[cid].setdefault("composerId", cid)
            elif key.startswith("bubbleId:"):
                parts = key.split(":")
                if len(parts) >= 3:
                    cid = parts[1]
                    bubbles.setdefault(cid, []).append((parts[2], data if isinstance(data, dict) else {}))
            elif key in (
                "workbench.panel.aichat.view.aichat.chatdata",
                "aiService.prompts",
            ):
                legacy_dir = out_dir / "_legacy"
                if not dry_run:
                    legacy_dir.mkdir(parents=True, exist_ok=True)
                    (legacy_dir / f"{safe_name(key)}.json").write_text(
                        json.dumps(data, ensure_ascii=False, indent=2) if data is not None else str(value),
                        encoding="utf-8",
                    )
                exported += 1
    finally:
        conn.close()

    if not dry_run:
        out_dir.mkdir(parents=True, exist_ok=True)

    used_ids = set(composers) | set(bubbles)
    for cid in used_ids:
        meta = composers.get(cid) or {}
        title = safe_name(str(meta.get("name") or meta.get("title") or cid[:8]))
        created = meta.get("createdAt") or meta.get("createdAtMs") or ""
        lines = [
            f"# {title}",
            "",
            f"- المعرف: `{cid}`",
            f"- المصدر: `{db_path}`",
        ]
        if created:
            try:
                dt = datetime.fromtimestamp(int(created) / 1000)
                lines.append(f"- التاريخ: {dt.strftime('%Y-%m-%d %H:%M')}")
            except (ValueError, OSError, TypeError):
                lines.append(f"- التاريخ الخام: {created}")
        lines.append("")
        headers = meta.get("fullConversationHeadersOnly") or []
        order = [str(h.get("bubbleId")) for h in headers if isinstance(h, dict) and h.get("bubbleId")]
        raw_bubbles = bubbles.get(cid) or []
        by_id = {bid: payload for bid, payload in raw_bubbles}
        ordered: List[Tuple[str, Dict[str, Any]]] = []
        seen = set()
        for bid in order:
            if bid in by_id:
                ordered.append((bid, by_id[bid]))
                seen.add(bid)
        for bid, payload in raw_bubbles:
            if bid not in seen:
                ordered.append((bid, payload))
        if not ordered:
            name_only = str(meta.get("name") or "").strip()
            if not name_only:
                continue
            lines.append("_لا توجد رسائل نصية في قاعدة البيانات لهذه المحادثة._")
        for _bid, payload in ordered:
            role = _bubble_role(payload)
            text = _bubble_text(payload)
            if not text:
                continue
            lines.append(f"## {role}")
            lines.append("")
            lines.append(text)
            lines.append("")
        if not dry_run:
            out_file = out_dir / f"{title}_{cid[:8]}.md"
            out_file.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
        exported += 1
    return exported


def export_cursor_sources(sources: Sequence[Dict[str, Any]], dest_root: Path, dry_run: bool = False) -> int:
    readable = dest_root / "محادثات_Cursor_مقروءة"
    total = 0
    for source in sources:
        if source["kind"] != "cursor":
            continue
        src = Path(source["path"])
        dbs: List[Path] = []
        if src.is_file() and src.suffix.lower() in {".vscdb", ".db", ".sqlite"}:
            dbs.append(src)
        elif src.name.lower() == "globalstorage":
            for name in ("state.vscdb", "state.vscdb.backup"):
                if (src / name).exists():
                    dbs.append(src / name)
        elif src.name.lower() == "workspacestorage":
            dbs.extend(src.glob("*/state.vscdb"))
        else:
            dbs.extend(src.glob("**/state.vscdb"))
        for db in dbs:
            total += export_cursor_db(db, readable / safe_name(db.parent.name + "_" + db.name), dry_run=dry_run)
    return total


def write_report(
    dest_root: Path,
    onedrive: Optional[Path],
    results: Sequence[Dict[str, Any]],
    exported_md: int,
    dry_run: bool,
) -> Path:
    lines = [
        "تقرير نقل المحادثات إلى OneDrive",
        "=" * 50,
        f"التاريخ: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"التجارب فقط: {'نعم' if dry_run else 'لا'}",
        f"مجلد OneDrive: {onedrive or 'غير محدد'}",
        f"مجلد الحفظ: {dest_root}",
        "",
        "ملاحظة: تم النسخ دون حذف الأصل حتى تبقى البرامج تعمل.",
        "",
    ]
    copied_files = 0
    copied_bytes = 0
    for row in results:
        copied_files += int(row["files"])
        copied_bytes += int(row["bytes"])
        lines.append(
            f"- {row['title']}: {row['files']} ملف / {format_size(row['bytes'])}"
        )
    lines.append("")
    lines.append(f"المجموع: {copied_files} ملف / {format_size(copied_bytes)}")
    lines.append(f"محادثات Cursor المصدّرة كملفات مقروءة: {exported_md}")
    lines.append("")
    lines.append("بعد المزامنة ستظهر الملفات على أجهزة OneDrive الأخرى وعلى الويب:")
    lines.append("https://onedrive.live.com")
    report = dest_root / "تقرير_النقل.txt"
    if not dry_run:
        dest_root.mkdir(parents=True, exist_ok=True)
        report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return report


def run(
    dest: Optional[Path] = None,
    extra_sources: Optional[Sequence[Path]] = None,
    dry_run: bool = False,
    env: Optional[Dict[str, str]] = None,
    project_dir: Optional[Path] = None,
    keys: Optional[Sequence[str]] = None,
) -> Dict[str, Any]:
    env = env if env is not None else dict(os.environ)
    project_dir = project_dir or Path(__file__).resolve().parent
    onedrive = dest.parent.parent if dest is not None else detect_onedrive(env=env)
    if dest is None:
        detected = detect_onedrive(env=env)
        if detected is None:
            raise SystemExit(
                "لم يتم العثور على مجلد OneDrive.\n"
                "ثبّت OneDrive أو شغّل الأمر مع: --dest \"مسار_الحفظ\""
            )
        onedrive = detected
        dest = detected / "المحادثات" / timestamp()

    sources = discover_sources(env=env, extra_sources=extra_sources, project_dir=project_dir)
    if keys:
        wanted = set(keys)
        sources = [s for s in sources if s["key"] in wanted]
    if not sources:
        raise SystemExit("لم يتم العثور على مصادر محادثات لنقلها.")

    results = []
    if not dry_run:
        dest.mkdir(parents=True, exist_ok=True)
    for source in sources:
        files, nbytes = copy_source(source, dest, dry_run=dry_run)
        results.append(
            {
                "key": source["key"],
                "title": source["title"],
                "files": files,
                "bytes": nbytes,
            }
        )
    exported_md = export_cursor_sources(sources, dest, dry_run=dry_run)
    report = write_report(dest, onedrive, results, exported_md, dry_run=dry_run)
    return {
        "dest": dest,
        "onedrive": onedrive,
        "results": results,
        "exported_md": exported_md,
        "report": report,
        "sources": sources,
    }


def print_plan(sources: Sequence[Dict[str, Any]], dest: Path) -> None:
    print("=" * 60)
    print("نقل المحادثات إلى OneDrive")
    print("=" * 60)
    print(f"سيتم الحفظ في:\n  {dest}")
    print()
    print("المصادر الموجودة:")
    for i, source in enumerate(sources, start=1):
        print(f"  {i}. {source['title']}  ({format_size(source['size'])})")
    print()
    print("الأصل يبقى في مكانه. هذه نسخة محفوظة في ون درايف.")
    print()


def self_test() -> int:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        home = tmp_path / "home"
        onedrive = home / "OneDrive"
        roaming = home / "AppData" / "Roaming"
        desktop = home / "Desktop"
        project = tmp_path / "project"
        global_dir = roaming / "Cursor" / "User" / "globalStorage"
        ws_dir = roaming / "Cursor" / "User" / "workspaceStorage" / "abc123"
        wa = roaming / "WhatsApp" / "chats"
        archive = project / "conversations"
        for folder in (onedrive, global_dir, ws_dir, wa, desktop, archive):
            folder.mkdir(parents=True, exist_ok=True)

        db_path = global_dir / "state.vscdb"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE cursorDiskKV (key TEXT PRIMARY KEY, value BLOB)")
        cid = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
        composer = {
            "composerId": cid,
            "name": "نقل الخدمة",
            "createdAt": 1737316260000,
            "fullConversationHeadersOnly": [{"bubbleId": "b1"}, {"bubbleId": "b2"}],
        }
        conn.execute(
            "INSERT INTO cursorDiskKV VALUES (?, ?)",
            (f"composerData:{cid}", json.dumps(composer, ensure_ascii=False)),
        )
        conn.execute(
            "INSERT INTO cursorDiskKV VALUES (?, ?)",
            (f"bubbleId:{cid}:b1", json.dumps({"type": 1, "text": "انقل المحادثات الى ون درايف"}, ensure_ascii=False)),
        )
        conn.execute(
            "INSERT INTO cursorDiskKV VALUES (?, ?)",
            (f"bubbleId:{cid}:b2", json.dumps({"type": 2, "text": "تم النسخ إلى OneDrive"}, ensure_ascii=False)),
        )
        conn.commit()
        conn.close()
        (ws_dir / "workspace.json").write_text('{"folder":"C:\\\\work"}', encoding="utf-8")
        (wa / "chat.txt").write_text("مرحبا", encoding="utf-8")
        (desktop / "محادثة_واتساب.txt").write_text("نص الدردشة", encoding="utf-8")
        (archive / "جلسة_سابقة.md").write_text("# أرشيف\n", encoding="utf-8")

        env = {
            "USERPROFILE": str(home),
            "APPDATA": str(roaming),
            "LOCALAPPDATA": str(home / "AppData" / "Local"),
            "OneDrive": str(onedrive),
            "HOME": str(home),
        }
        dest = onedrive / "المحادثات" / "test-run"
        result = run(dest=dest, dry_run=False, env=env, project_dir=project)
        detected = detect_onedrive(env=env, home=home)
        if detected != onedrive:
            print("SELF-TEST FAILED: detect_onedrive")
            return 1
        dry_dest = onedrive / "المحادثات" / "dry-run"
        run(dest=dry_dest, dry_run=True, env=env, project_dir=project)
        if dry_dest.exists():
            print("SELF-TEST FAILED: dry-run wrote files")
            return 1
        report = dest / "تقرير_النقل.txt"
        copied_db = dest / "Cursor — قاعدة المحادثات العامة" / "state.vscdb"
        readable = list((dest / "محادثات_Cursor_مقروءة").rglob("*.md"))
        archive_copied = dest / "أرشيف محادثات هذا المشروع" / "جلسة_سابقة.md"
        desktop_copied = dest / "ملفات محادثات من سطح المكتب" / "محادثة_واتساب.txt"
        wa_copied = dest / "واتساب (سطح المكتب)" / "chats" / "chat.txt"

        checks = {
            "report": report.exists() and "OneDrive" in report.read_text(encoding="utf-8"),
            "db_copied": copied_db.exists(),
            "readable_md": any("انقل المحادثات" in p.read_text(encoding="utf-8") for p in readable),
            "archive": archive_copied.exists(),
            "desktop": desktop_copied.exists(),
            "whatsapp": wa_copied.exists(),
            "exported_count": result["exported_md"] >= 1,
        }
        failed = [name for name, ok in checks.items() if not ok]
        if failed:
            print("SELF-TEST FAILED:", ", ".join(failed))
            print("dest contents:")
            for p in dest.rglob("*"):
                print(" ", p.relative_to(dest))
            return 1
        print("SELF-TEST OK")
        print(f"dest={dest}")
        print(f"files={sum(r['files'] for r in result['results'])} exported_md={result['exported_md']}")
        return 0


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="نسخ المحادثات إلى OneDrive")
    parser.add_argument("--dest", type=Path, help="مجلد الحفظ (افتراضي: OneDrive/المحادثات/التاريخ)")
    parser.add_argument("--source", type=Path, action="append", default=[], help="مجلد إضافي لنسخه")
    parser.add_argument("--dry-run", action="store_true", help="عرض ما سيتم نسخه دون كتابة")
    parser.add_argument("--yes", action="store_true", help="بدون سؤال تأكيد")
    parser.add_argument("--self-test", action="store_true", help="اختبار داخلي")
    parser.add_argument("--only", help="مفاتيح المصادر مفصولة بفاصلة (مثال: cursor_global,whatsapp_roaming)")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    if args.self_test:
        return self_test()
    try:
        env = dict(os.environ)
        project_dir = Path(__file__).resolve().parent
        onedrive = detect_onedrive(env=env)
        dest = args.dest
        if dest is None:
            if onedrive is None:
                print("لم يتم العثور على مجلد OneDrive على هذا الجهاز.")
                print("ثبّت تطبيق OneDrive من مايكروسوفت، أو حدّد المجلد:")
                print('  python copy_chats_to_onedrive.py --dest "C:\\Users\\YOU\\OneDrive\\المحادثات"')
                return 2
            dest = onedrive / "المحادثات" / timestamp()
        keys = [k.strip() for k in args.only.split(",")] if args.only else None
        sources = discover_sources(env=env, extra_sources=args.source, project_dir=project_dir)
        if keys:
            sources = [s for s in sources if s["key"] in keys]
        if not sources:
            print("لم يتم العثور على محادثات لنقلها.")
            print("يمكنك تمرير مجلد يدوياً: --source \"D:\\مجلد_المحادثات\"")
            return 3
        print_plan(sources, dest)
        if not args.yes and not args.dry_run:
            try:
                answer = input("Enter للنسخ الآن، أو اكتب لا للإلغاء: ").strip()
            except EOFError:
                answer = ""
            if answer in {"لا", "n", "N", "no", "NO"}:
                print("تم الإلغاء.")
                return 0
        result = run(
            dest=dest,
            extra_sources=args.source,
            dry_run=args.dry_run,
            env=env,
            project_dir=project_dir,
            keys=keys,
        )
        print()
        if args.dry_run:
            print("تجربة بدون كتابة. لم يتم نسخ شيء.")
        else:
            print("تم الحفظ في:")
            print(f"  {result['dest']}")
            print(f"التقرير: {result['report']}")
        print(f"ملفات Cursor المقروءة: {result['exported_md']}")
        for row in result["results"]:
            print(f"  - {row['title']}: {row['files']} ملف / {format_size(row['bytes'])}")
        return 0
    except SystemExit as exc:
        if exc.code in (None, 0):
            return 0
        if isinstance(exc.code, str):
            print(exc.code)
            return 1
        return int(exc.code)
    except KeyboardInterrupt:
        print("\nتم الإلغاء.")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
