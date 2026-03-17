import csv
import html
import os
import plistlib
import re
import sqlite3
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable
from PySide6.QtCore import QObject, Qt, QThread, Signal, QUrl
from PySide6.QtGui import QAction, QColor, QDesktopServices, QIcon, QPixmap, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

# ── Palette ───────────────────────────────────────────────────────────────────
BG = "#0a0a0f"
GLASS_STR = "#18181f"
GLASS_HOV = "#22222c"
BORDER = "#2a2a38"
ACCENT = "#6c63ff"
ACCENT2 = "#a78bfa"
TEXT = "#f0f0ff"
TEXT_SEC = "#8888aa"
TEXT_DIM = "#44445a"
SUCCESS = "#34d399"
ERROR = "#f87171"
WARN = "#fbbf24"

# ── Extraction heuristics ─────────────────────────────────────────────────────
CALL_HISTORY_CANDIDATES = [
    ("HomeDomain", "Library/CallHistoryDB/CallHistory.storedata"),
    ("WirelessDomain", "Library/CallHistory/call_history.db"),
]
SMS_CANDIDATES = [
    ("HomeDomain", "Library/SMS/sms.db"),
]
PHOTO_PREFIX_KEYWORDS = (
    "dcim/",
    "media/dcim/",
    "photodata/",
    "media/photodata/",
    "applethumbnails/",
    "photo streamsdata/",
    "photos/",
)
PHOTO_EXTENSIONS = (
    ".jpg", ".jpeg", ".png", ".heic", ".heif", ".gif", ".bmp", ".tif", ".tiff",
    ".aae", ".mov", ".mp4", ".m4v", ".3gp", ".dng", ".avi", ".webp",
)
PHOTO_DOMAINS = {
    "mediadomain",
    "camerarolldomain",
}
VOICEMAIL_KEYWORDS = (
    "voicemail",
    "callrecording",
)
VOICEMAIL_AUDIO_EXTENSIONS = (
    ".amr", ".m4a", ".caf", ".aac", ".wav", ".mp3", ".3gp", ".m4b", ".m4r",
)
CONTACT_KEYWORDS = (
    "library/addressbook/",
    "addressbook.sqlitedb",
    "addressbookimages.sqlitedb",
    "addressbook-v22.abcddb",
    "contacts/",
)
CONTACT_DB_EXTENSIONS = (
    ".db", ".sqlite", ".sqlite3", ".sqlitedb", ".abcddb",
)
ZALO_KEYWORDS = (
    "zalo",
    "com.vng.zalo",
    "vn.com.vng.zalo",
    "net.zing.zalo",
    "com.zing.zalo",
)

ADVANCED_KEYCHAIN_HINTS = (
    "keychain",
    "password",
    "accounts",
    "credential",
    "security",
)
ADVANCED_WIFI_HINTS = (
    "wifi",
    "wi-fi",
    "wlan",
    "networkidentification",
    "systemconfiguration",
    "knownnetworks",
    "wireless",
)
ADVANCED_HEALTH_HINTS = (
    "health",
    "healthdb",
    "healthkit",
    "activitycache",
    "fitness",
    "mobility",
)
ADVANCED_PREVIEW_EXTENSIONS = {
    ".plist", ".db", ".sqlite", ".sqlite3", ".storedata", ".json", ".txt", ".log", ".xml"
}


# ── Icon helpers ──────────────────────────────────────────────────────────────
def _preferred_icon_names() -> tuple[str, ...]:
    if sys.platform == "darwin":
        return ("phone.icns", "phone.png", "phone.ico", "wicon.icns", "wicon.png", "wicon.ico")
    if os.name == "nt":
        return ("phone.ico", "phone.png", "phone.icns", "wicon.ico", "wicon.png", "wicon.icns")
    return ("phone.png", "phone.ico", "phone.icns", "wicon.png", "wicon.ico", "wicon.icns")



def _resource_markers() -> tuple[str, ...]:
    return (
        "icons",
        "splashscreen.png",
        "default_output.png",
        "novideothumb.jpg",
        "phone.png",
        "phone.ico",
        "phone.icns",
        "wicon.png",
        "wicon.ico",
        "wicon.icns",
        "call.png",
        "sms.png",
        "photo.png",
        "voice.png",
        "contact.png",
        "all.png",
    )



def _looks_like_resource_root(root: Path) -> bool:
    try:
        for marker in _resource_markers():
            p = root / marker
            if p.is_file() or p.is_dir():
                return True
    except Exception:
        return False
    return False



def resource_root() -> Path:
    exe_dir = Path(sys.argv[0]).resolve().parent
    if _looks_like_resource_root(exe_dir):
        return exe_dir

    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)

    return Path(__file__).resolve().parent



def app_icon() -> QIcon:
    root = resource_root()

    for name in _preferred_icon_names():
        f = root / name
        if f.is_file():
            return QIcon(str(f))

    for name in _preferred_icon_names():
        f = root / "icons" / name
        if f.is_file():
            return QIcon(str(f))

    return QIcon()


CATEGORY_ICON_FILES = {
    "call": ("call.png", "call_history.png", "phone.png", "phone.ico", "phone.icns"),
    "sms": ("sms.png", "message.png", "messages.png"),
    "photos": ("photo.png", "photos.png", "image.png"),
    "voicemail": ("voice.png", "voicemail.png", "mic.png"),
    "contacts": ("contact.png", "contacts.png", "addressbook.png", "person.png"),
    "advanced": ("all.png", "advanced.png", "shield.png", "database.png"),
    "all": ("all.png", "select_all.png"),
}


def resource_file_path(*names: str) -> Path | None:
    root = resource_root()
    for name in names:
        p = root / name
        if p.is_file():
            return p
    for name in names:
        p = root / "icons" / name
        if p.is_file():
            return p
    return None



def category_icon_path(key: str) -> Path | None:
    candidates = CATEGORY_ICON_FILES.get(key, ())
    return resource_file_path(*candidates)


# ── Backup discovery ──────────────────────────────────────────────────────────
def is_backup_dir(path: Path) -> bool:
    if not path or not path.is_dir():
        return False
    required = ("Manifest.db", "Manifest.plist", "Info.plist", "Status.plist")
    return all((path / name).is_file() for name in required)



def backup_root_candidates() -> list[Path]:
    home = Path.home()
    candidates: list[Path] = []

    if os.name == "nt":
        appdata = os.environ.get("APPDATA")
        if appdata:
            candidates.append(Path(appdata) / "Apple Computer" / "MobileSync" / "Backup")
        candidates.extend(
            [
                home / "Apple" / "MobileSync" / "Backup",
                home / "AppData" / "Roaming" / "Apple Computer" / "MobileSync" / "Backup",
                home / "AppData" / "Local" / "Apple" / "MobileSync" / "Backup",
            ]
        )
    elif sys.platform == "darwin":
        candidates.append(home / "Library" / "Application Support" / "MobileSync" / "Backup")
    else:
        candidates.extend(
            [
                home / ".config" / "MobileSync" / "Backup",
                home / "MobileSync" / "Backup",
            ]
        )

    deduped: list[Path] = []
    seen: set[str] = set()
    for item in candidates:
        key = str(item)
        if key not in seen:
            seen.add(key)
            deduped.append(item)
    return deduped



def discover_backups_in(root: Path, max_depth: int = 2) -> list[Path]:
    if not root or not root.exists():
        return []
    if is_backup_dir(root):
        return [root]

    found: list[Path] = []
    root = root.resolve()
    base_depth = len(root.parts)

    try:
        for child in root.rglob("*"):
            if not child.is_dir():
                continue
            depth = len(child.parts) - base_depth
            if depth > max_depth:
                continue
            if is_backup_dir(child):
                found.append(child)
    except Exception:
        return []
    return found



def latest_backup_from(paths: list[Path]) -> Path | None:
    valid = [p for p in paths if is_backup_dir(p)]
    if not valid:
        return None
    return max(valid, key=lambda p: p.stat().st_mtime)



def auto_find_latest_backup() -> Path | None:
    found: list[Path] = []
    for root in backup_root_candidates():
        if root.exists():
            found.extend(discover_backups_in(root, max_depth=2))
    return latest_backup_from(found)



def resolve_backup_folder(selected: str | Path) -> Path | None:
    if not selected:
        return auto_find_latest_backup()

    path = Path(selected).expanduser()
    if is_backup_dir(path):
        return path

    found = discover_backups_in(path, max_depth=2)
    if not found:
        return None
    return latest_backup_from(found)


# ── UI helpers ────────────────────────────────────────────────────────────────
def section_label(text: str) -> QLabel:
    lbl = QLabel(text)
    lbl.setStyleSheet(
        f"color: {ACCENT2}; font-size: 11px; font-weight: 700; letter-spacing: 0.5px;"
    )
    return lbl



def card_widget() -> QFrame:
    frame = QFrame()
    frame.setObjectName("card")
    frame.setStyleSheet(
        f"""
        QFrame#card {{
            background: {GLASS_STR};
            border: 1px solid {BORDER};
            border-radius: 16px;
        }}
        """
    )
    return frame



def divider() -> QFrame:
    line = QFrame()
    line.setFrameShape(QFrame.Shape.HLine)
    line.setFrameShadow(QFrame.Shadow.Plain)
    line.setFixedHeight(1)
    line.setStyleSheet(f"background: {BORDER}; border: none;")
    return line


class RowToggle(QWidget):
    toggled = Signal(bool)

    def __init__(self, icon_key: str, label_text: str, fallback_icon_text: str = "", parent: QWidget | None = None):
        super().__init__(parent)
        self._enabled = False
        self._icon_key = icon_key
        self._fallback_icon_text = fallback_icon_text

        self.checkbox = QCheckBox()
        self.checkbox.setCursor(Qt.CursorShape.PointingHandCursor)
        self.checkbox.setEnabled(False)
        self.checkbox.toggled.connect(self._on_toggled)

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(26, 26)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._set_icon()

        self.name_label = QLabel(label_text)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(12)
        layout.addWidget(self.icon_label, 0, Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(self.name_label, 1)
        layout.addWidget(self.checkbox)

        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.set_enabled(False)
        self._refresh_style()

    def _set_icon(self):
        icon_path = category_icon_path(self._icon_key)
        if icon_path and icon_path.is_file():
            pm = QPixmap(str(icon_path))
            if not pm.isNull():
                self.icon_label.setPixmap(pm.scaled(22, 22, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
                self.icon_label.setText("")
                self.icon_label.setStyleSheet("background: transparent;")
                return
        self.icon_label.setPixmap(QPixmap())
        self.icon_label.setText(self._fallback_icon_text)
        self.icon_label.setStyleSheet("font-size: 18px; background: transparent;")

    def mousePressEvent(self, event):
        if self._enabled:
            self.checkbox.toggle()
        super().mousePressEvent(event)

    def _on_toggled(self, checked: bool):
        self._refresh_style()
        self.toggled.emit(checked)

    def set_enabled(self, enabled: bool):
        self._enabled = enabled
        self.checkbox.setEnabled(enabled)
        if not enabled:
            self.checkbox.setChecked(False)
        self._refresh_style()

    def set_checked(self, checked: bool):
        self.checkbox.setChecked(checked)

    def is_checked(self) -> bool:
        return self.checkbox.isChecked()

    def _refresh_style(self):
        bg = GLASS_HOV if self.checkbox.isChecked() else GLASS_STR
        fg = TEXT if self.checkbox.isChecked() else (TEXT_SEC if self._enabled else TEXT_DIM)
        self.setStyleSheet(
            f"""
            QWidget {{
                background: {bg};
                border: none;
                border-radius: 14px;
            }}
            QLabel {{
                color: {fg};
                background: transparent;
                font-size: 14px;
            }}
            QCheckBox {{
                spacing: 0px;
                background: transparent;
            }}
            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
                border-radius: 9px;
                border: 1px solid {ACCENT2 if self.checkbox.isChecked() else TEXT_DIM};
                background: {'#a78bfa' if self.checkbox.isChecked() else 'transparent'};
            }}
            QCheckBox::indicator:disabled {{
                border: 1px solid {TEXT_DIM};
                background: transparent;
            }}
            """
        )


# ── Backend helpers ───────────────────────────────────────────────────────────
def ensure_pyiosbackup():
    try:
        from pyiosbackup import Backup  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("Run: pip install pyiosbackup") from exc



def classify_zalo(entry) -> bool:
    domain = (entry.domain or "").lower()
    rel = (entry.relative_path or "").lower()
    return any(key in domain or key in rel for key in ZALO_KEYWORDS)



def classify_photo(entry) -> bool:
    domain = (entry.domain or "").lower()
    rel = (entry.relative_path or "")
    rel_l = rel.lower()

    if domain in PHOTO_DOMAINS:
        if rel_l.startswith(PHOTO_PREFIX_KEYWORDS):
            return True
        if any(key in rel_l for key in PHOTO_PREFIX_KEYWORDS):
            return True
        if rel_l.endswith(PHOTO_EXTENSIONS):
            return True

    # Một số backup/app version có thể để ảnh trong app domain hoặc shared group,
    # miễn path vẫn rõ ràng là DCIM / PhotoData / file media phổ biến.
    if any(key in rel_l for key in PHOTO_PREFIX_KEYWORDS) and rel_l.endswith(PHOTO_EXTENSIONS):
        return True

    return False



def classify_voicemail(entry) -> bool:
    domain = (entry.domain or "").lower()
    rel = (entry.relative_path or "").lower()
    if any(key in rel for key in VOICEMAIL_KEYWORDS):
        return True
    return domain == "mediadomain" and "recordings" in rel and rel.endswith(VOICEMAIL_AUDIO_EXTENSIONS)


def classify_contacts(entry) -> bool:
    domain = (entry.domain or "").lower()
    rel = (entry.relative_path or "").replace("\\", "/").lower()
    combined = f"{domain}/{rel}"
    if any(key in combined for key in CONTACT_KEYWORDS):
        return True
    if rel.endswith(CONTACT_DB_EXTENSIONS) and ("addressbook" in rel or "/contacts/" in combined):
        return True
    return False


def advanced_data_group(domain: str, relative_path: str) -> str:
    domain_l = (domain or "").lower()
    rel_l = (relative_path or "").replace("\\", "/").lower()
    combined = f"{domain_l}/{rel_l}"

    if (
        "keychain" in combined
        or rel_l.endswith("keychain-backup.plist")
        or "password" in combined
        or ("accounts" in combined and ("preferences" in combined or "plist" in combined))
    ):
        return "Keychain / Passwords"

    if (
        any(hint in combined for hint in ADVANCED_WIFI_HINTS)
        or rel_l.endswith("com.apple.wifi.plist")
        or rel_l.endswith("networkinterfaces.plist")
        or rel_l.endswith("preferences.plist") and "systemconfiguration" in rel_l
    ):
        return "Wi-Fi / Network"

    if (
        "/health/" in combined
        or any(hint in combined for hint in ADVANCED_HEALTH_HINTS)
        or rel_l.endswith("healthdb.sqlite")
        or rel_l.endswith("healthdb_secure.sqlite")
    ):
        return "Health"

    return ""



def classify_advanced_data(entry) -> bool:
    return bool(advanced_data_group(entry.domain or "", entry.relative_path or ""))



def try_open_backup(folder: str, password: str):
    ensure_pyiosbackup()
    from pyiosbackup import Backup
    return Backup.from_path(Path(folder), password=password or "")



def write_entry(entry, dest_root: str | Path, preserve_domain: bool = False) -> Path:
    dest_root = Path(dest_root)
    if preserve_domain:
        dest = dest_root / entry.domain / entry.relative_path
    else:
        dest = dest_root / entry.relative_path
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(entry.read_bytes())
    return dest





def default_output_dir() -> Path:
    return Path.home() / "Downloads" / "Backup_decryptor"


def _safe_filename(name: str, default: str = "unknown") -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in (" ", "-", "_", ".") else "_" for ch in (name or "").strip())
    cleaned = cleaned.strip(" ._")
    return cleaned[:120] or default


def _apple_time_to_str(value) -> str:
    if value in (None, "", 0, "0"):
        return ""
    try:
        raw = int(value)
    except Exception:
        try:
            raw = float(value)
        except Exception:
            return str(value)

    abs_raw = abs(raw)
    if abs_raw > 10**15:
        seconds = raw / 1_000_000_000
    elif abs_raw > 10**12:
        seconds = raw / 1_000_000
    elif abs_raw > 10**10:
        seconds = raw / 1000
    else:
        seconds = raw

    try:
        dt = datetime(2001, 1, 1) + timedelta(seconds=seconds)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return str(value)


def _table_exists(conn: sqlite3.Connection, name: str) -> bool:
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (name,),
    ).fetchone()
    return row is not None


def _table_columns(conn: sqlite3.Connection, table: str) -> set[str]:
    try:
        return {row[1] for row in conn.execute(f"PRAGMA table_info('{table}')").fetchall()}
    except Exception:
        return set()


def _strip_typedstream_noise(text: str) -> str:
    if not text:
        return ""

    cleaned = str(text)
    replacements = (
        "[PhoneNumber/",
        "PhoneNumber/",
        "PhoneNumber",
        "NSNumber",
        "NSString",
        "NSArray",
        "NSDictionary",
        "NSObject",
        "NSMutableAttributedString",
        "NSAttributedString",
        "NSMutableString",
        "__kIMMessagePartAttributeName",
        "__kIMAttributeName",
        "kIMAttributeName",
        "AttributeName",
        "UValue",
        "IntegralValue",
        "TTime",
        "UHours",
        "DDScannerResult",
        "versionYdd",
        "version",
        "result",
    )
    for item in replacements:
        cleaned = cleaned.replace(item, " ")

    cleaned = re.sub(r"\b\d{0,4}ZNS\.objects\d*\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b\d{0,4}NS\.objects\d*\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b__?kIM[A-Za-z0-9_]*\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b[A-Za-z_]*AttributeName\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bNS[A-Za-z0-9_]+\b", " ", cleaned)
    cleaned = re.sub(r"(?<!https:)(?<!http:)\[", " ", cleaned)
    cleaned = cleaned.replace("]", " ")
    cleaned = re.sub(r"(?<=\d)\s*/\s*(?=$|[A-Za-z0-9_])", " ", cleaned)
    cleaned = re.sub(r"(^|\s)[_]{1,4}(?=\s|$)", " ", cleaned)
    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    return cleaned.strip(" _-:/")


def _inject_candidate_breaks(text: str) -> str:
    if not text:
        return ""

    cleaned = str(text).replace("\x00", " ")
    cleaned = cleaned.replace("\uFFFC", " ")

    # Add soft boundaries before common payloads that are often packed into
    # Apple attributedBody blobs so later extraction can split them cleanly.
    cleaned = re.sub(r"(?<![\s(])(?=https?://)", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"(?<![\s+])(?=\+?\d(?:[\d\s\-()]{5,}\d))", "\n", cleaned)

    # Separate common metadata markers from surrounding text.
    cleaned = re.sub(r"(?<=[A-Za-z0-9])(?=(?:UValue|IntegralValue|TTime|UHours|B\d|U\d)\b)", "\n", cleaned)
    cleaned = re.sub(r"(?<=\))(?=[A-Za-z0-9])", "\n", cleaned)

    return cleaned


def _is_sparse_digit_noise(text: str) -> bool:
    cleaned = _normalize_text(text)
    if not cleaned:
        return True
    if any(ch.isalpha() for ch in cleaned):
        return False

    digits = re.sub(r"\D", "", cleaned)
    if not digits:
        return False

    groups = re.findall(r"\d+", cleaned)
    if not groups:
        return False

    if re.fullmatch(r"(?:\d\s+){3,}\d", cleaned):
        return True
    if len(groups) >= 4 and max(len(g) for g in groups) <= 2:
        return True
    if len(digits) <= 5 and len(groups) >= 3:
        return True
    if len(digits) >= 5 and sum(1 for g in groups if len(g) == 1) >= 4:
        return True
    return False


def _clean_extracted_candidate(text: str) -> str:
    cleaned = _strip_typedstream_noise(_normalize_text(text))
    if not cleaned:
        return ""

    cleaned = cleaned.replace("\x00", " ").strip()
    cleaned = re.sub(r"^\+\s*([A-Za-z])(?=https?://)", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"^\+\s*(?=https?://)", "", cleaned, flags=re.IGNORECASE)

    url_match = re.search(r"https?://[^\s<>'\"]+", cleaned, flags=re.IGNORECASE)
    if url_match:
        url = url_match.group(0)
        url = re.sub(r"[A-Za-z]{1,2}$", "", url)
        url = url.rstrip(".,;)]}>")
        return url

    cleaned = re.sub(r"\s*([\-–—:])\s*", r" \1 ", cleaned)
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip(" _-:/")

    if _is_sparse_digit_noise(cleaned):
        return ""

    has_letters = any(ch.isalpha() for ch in cleaned)
    has_digits = any(ch.isdigit() for ch in cleaned)
    if has_letters and has_digits:
        return cleaned

    probe = cleaned
    probe = re.sub(r"^\+\s*([A-Za-z])(?=\d)", "", probe)
    probe = re.sub(r"^\+\s+(?=\d)", "+", probe)
    probe = re.sub(r"\s+[A-Za-z]{1,2}(?:\s+[A-Za-z]{1,2}){0,3}$", "", probe).strip()
    phone_matches = re.findall(r"\+?\d(?:[\d\s\-()]{5,}\d)", probe)
    if phone_matches:
        ranked_phones = sorted(
            phone_matches,
            key=lambda s: (len(re.sub(r"\D", "", s)), len(s)),
            reverse=True,
        )
        best_phone = re.sub(r"^\+\s+", "+", ranked_phones[0])
        if _is_sparse_digit_noise(best_phone):
            return ""
        return re.sub(r"\s{2,}", " ", best_phone).strip(" _-:/")

    cleaned = re.sub(r"^\+\s*[A-Za-z](?=https?://)", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"^\+\s*(?=https?://)", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"^\+\s*(?=\d)", "+", cleaned)
    if re.search(r"https?://", cleaned, flags=re.IGNORECASE):
        cleaned = re.sub(r"\s+[A-Za-z]{1,2}(?:\s+[A-Za-z]{1,2}){0,3}$", "", cleaned).strip()
    elif re.fullmatch(r"\+?\d[\d\s\-()]{5,}\d(?:\s+[A-Za-z]{1,3})?", cleaned):
        cleaned = re.sub(r"\s+[A-Za-z]{1,3}(?:\s+[A-Za-z]{1,3})?$", "", cleaned).strip()

    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    if _is_sparse_digit_noise(cleaned):
        return ""
    return cleaned.strip(" _-:/")


def _normalize_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        try:
            value = value.decode("utf-8")
        except Exception:
            value = value.decode("utf-8", errors="ignore")
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\uFFFC", "", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _is_truthy(value) -> bool:
    return str(value).strip().lower() in {"1", "true", "yes", "y"}


def _duration_to_str(value) -> str:
    try:
        total = int(float(value or 0))
    except Exception:
        total = 0
    if total <= 0:
        return "0s"
    hours, rem = divmod(total, 3600)
    minutes, seconds = divmod(rem, 60)
    if hours:
        return f"{hours}h {minutes:02d}m {seconds:02d}s"
    if minutes:
        return f"{minutes}m {seconds:02d}s"
    return f"{seconds}s"


def _candidate_score(text: str) -> tuple[int, int]:
    cleaned = _normalize_text(text)
    if not cleaned:
        return (-1000, 0)
    if _is_sparse_digit_noise(cleaned):
        return (-350, len(cleaned))

    lower = cleaned.lower()
    noise_terms = (
        "nsstring",
        "nsmutable",
        "nsdictionary",
        "nsdata",
        "nsnumber",
        "nsuuid",
        "nsobjects",
        "zns.objects",
        "nsarray",
        "nsvalue",
        "nsobject",
        "$objects",
        "$class",
        "foundation",
        "uikit",
        "nsattributes",
        "nsattachment",
        "streamtyped",
        "phonenumber/",
        "uvalue",
        "integralvalue",
        "ttime",
        "uhours",
        "ddscannerresult",
        "versionydd",
        "__kimmessagepartattributename",
        "__kimattributename",
        "kimattributename",
        "attributename",
        " il",
        " ii",
    )
    if any(term in lower for term in noise_terms):
        return (-500, len(cleaned))

    score = 0
    if any(ch.isspace() for ch in cleaned):
        score += 30
    if any(ch.isdigit() for ch in cleaned):
        score += 10
    if any(ch.isalpha() for ch in cleaned):
        score += 15
    if any(ch.isalpha() for ch in cleaned) and any(ch.isdigit() for ch in cleaned):
        score += 18
    if re.search(r"[A-Za-zÀ-ỹ]{2,}\s+[A-Za-zÀ-ỹ]{2,}", cleaned):
        score += 30
    if re.search(r"\b\d{8,20}\b\s*(?:[-–—:]\s*)?[A-Za-zÀ-ỹ]", cleaned):
        score += 40
    if len(cleaned) <= 500:
        score += 15
    if 4 <= len(cleaned) <= 12:
        score += 6
    if re.fullmatch(r"[A-Za-zÀ-ỹ0-9 _\-/\.\+:@#%&()\[\],;!?='\"]{1,600}", cleaned):
        score += 20

    weird = sum(1 for ch in cleaned if not (ch.isalnum() or ch.isspace() or ch in "._-/:@#%&()[],\'\"!?+=;*"))
    if weird:
        score -= weird * 4

    if cleaned.count("_") >= 2:
        score -= 20
    if sum(ch.isupper() for ch in cleaned) > max(8, len(cleaned) * 0.55):
        score -= 15

    for hint in ("otp", "code", "password", "verification", "http", "https", "khong", "ban ", "ma "):
        if hint in lower:
            score += 8

    return (score, len(cleaned))


def _is_probably_typedstream(raw: bytes) -> bool:
    return raw.startswith(b"\x04\x0bstreamtyped") or b"streamtyped" in raw[:64]


def _read_typedstream_length(raw: bytes, idx: int) -> tuple[int | None, int]:
    if idx >= len(raw):
        return (None, idx)
    first = raw[idx]
    idx += 1
    if first == 0x81:
        if idx + 2 > len(raw):
            return (None, idx)
        return (int.from_bytes(raw[idx:idx + 2], "little", signed=False), idx + 2)
    if first == 0x82:
        if idx + 4 > len(raw):
            return (None, idx)
        return (int.from_bytes(raw[idx:idx + 4], "little", signed=False), idx + 4)
    if first == 0x83:
        if idx + 8 > len(raw):
            return (None, idx)
        return (int.from_bytes(raw[idx:idx + 8], "little", signed=False), idx + 8)
    return (first, idx)


def _extract_typedstream_utf8_candidates(raw: bytes) -> list[str]:
    candidates: list[str] = []
    if not raw:
        return candidates

    for idx, byte in enumerate(raw):
        if byte != 0x2B:
            continue
        strlen, start = _read_typedstream_length(raw, idx + 1)
        if not strlen or strlen < 1 or strlen > 20000:
            continue
        end = start + strlen
        if end > len(raw):
            continue
        chunk = raw[start:end]
        try:
            decoded = chunk.decode("utf-8", errors="strict")
        except Exception:
            decoded = chunk.decode("utf-8", errors="ignore")
        cleaned = _normalize_text(decoded)
        if cleaned:
            candidates.append(cleaned)
    return candidates


def _extract_preferred_typedstream_text(raw: bytes) -> str:
    preferred: list[str] = []
    for candidate in _extract_typedstream_utf8_candidates(raw):
        lowered_raw = candidate.casefold()
        if any(term in lowered_raw for term in (
            'nsobject',
            'phonenumber/',
            'ddscannerresult',
            'streamtyped',
            'uvalue',
            'integralvalue',
            'ttime',
            'uhours',
            'versionydd',
        )):
            continue
        cleaned = _clean_extracted_candidate(candidate)
        if not cleaned:
            continue
        if _candidate_score(cleaned)[0] >= 35:
            preferred.append(cleaned)
    if not preferred:
        return ''
    preferred.sort(key=_candidate_score, reverse=True)
    return preferred[0]


def _extract_decoded_stream_candidates(raw: bytes) -> list[str]:
    candidates: list[str] = []
    try:
        decoded = raw.decode("utf-8", errors="ignore")
    except Exception:
        return candidates

    if not decoded:
        return candidates

    cleaned = decoded.replace("\x00", " ")
    cleaned = _inject_candidate_breaks(cleaned)
    cleaned = "".join(ch if (ch.isprintable() or ch in "\n\r\t") else " " for ch in cleaned)

    for match in re.finditer(r'https?://[^\s<>\'"]+', cleaned, flags=re.IGNORECASE):
        candidates.append(match.group(0))
    for match in re.finditer(r"\+?\d(?:[\d\s\-()]{5,}\d)", cleaned):
        candidates.append(match.group(0))

    parts = re.split(r"[\n\r\t]+", cleaned)
    for part in parts:
        part = _clean_extracted_candidate(part)
        if part:
            candidates.append(part)
    return candidates


def _extract_strings_from_obj(obj, sink: list[str]):
    if obj is None:
        return
    if isinstance(obj, str):
        sink.append(obj)
        return
    if isinstance(obj, bytes):
        try:
            sink.append(obj.decode("utf-8", errors="ignore"))
        except Exception:
            return
        return
    if isinstance(obj, dict):
        for value in obj.values():
            _extract_strings_from_obj(value, sink)
        return
    if isinstance(obj, (list, tuple, set)):
        for value in obj:
            _extract_strings_from_obj(value, sink)


def _extract_text_from_attributed_body(blob) -> str:
    if blob in (None, "", b""):
        return ""

    raw = bytes(blob) if isinstance(blob, memoryview) else blob
    if isinstance(raw, str):
        return _normalize_text(raw)

    if not isinstance(raw, (bytes, bytearray)):
        return ""

    raw = bytes(raw)
    candidates: list[str] = []

    try:
        parsed = plistlib.loads(raw)
        _extract_strings_from_obj(parsed, candidates)
    except Exception:
        pass

    if _is_probably_typedstream(raw):
        preferred = _extract_preferred_typedstream_text(raw)
        if preferred:
            return preferred
        candidates.extend(_extract_typedstream_utf8_candidates(raw))
        candidates.extend(_extract_decoded_stream_candidates(raw))
    else:
        candidates.extend(_extract_decoded_stream_candidates(raw))

    cleaned_candidates: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        cleaned = _clean_extracted_candidate(candidate)
        if not cleaned:
            continue
        if sum(ch.isalnum() for ch in cleaned) < 1:
            continue
        if cleaned not in seen:
            seen.add(cleaned)
            cleaned_candidates.append(cleaned)

    if not cleaned_candidates:
        return ""

    ranked = sorted(cleaned_candidates, key=_candidate_score, reverse=True)
    for candidate in ranked:
        if _candidate_score(candidate)[0] >= 0:
            return candidate
    return ""


def _message_display_text(row: sqlite3.Row | dict) -> str:
    subject = _normalize_text(row["subject"]) if "subject" in row.keys() else ""
    text_val = _normalize_text(row["text"]) if "text" in row.keys() else ""
    if text_val:
        return text_val

    if "attributedBody" in row.keys():
        from_blob = _extract_text_from_attributed_body(row["attributedBody"])
        if from_blob:
            return from_blob

    has_attachments = _is_truthy(row["cache_has_attachments"]) if "cache_has_attachments" in row.keys() else False
    if subject:
        return f"[Subject] {subject}"
    if has_attachments:
        return "[Attachment / rich content only]"
    return "[No readable text extracted from this row]"


def _conversation_label_for_row(row: sqlite3.Row | dict) -> str:
    chat_name = _normalize_text(row["chat_display_name"]) if "chat_display_name" in row.keys() else ""
    chat_identifier = _normalize_text(row["chat_identifier"]) if "chat_identifier" in row.keys() else ""
    contact = _normalize_text(row["contact"]) if "contact" in row.keys() else ""
    service = _normalize_text(row["service"]) if "service" in row.keys() else ""
    chat_service = _normalize_text(row["chat_service_name"]) if "chat_service_name" in row.keys() else ""

    primary = chat_name or chat_identifier or contact or "Unknown conversation"
    service_label = service or chat_service
    if service_label and service_label.lower() not in primary.lower():
        return f"{primary} · {service_label}"
    return primary


def _message_sort_key(msg: dict) -> tuple[str, int]:
    return (msg.get("datetime") or "9999-99-99 99:99:99", int(msg.get("rowid") or 0))


def _wrap_html_page(title: str, subtitle: str, body_html: str, extra_head: str = "") -> str:
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset=\"utf-8\">
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
<title>{html.escape(title)}</title>
<style>
:root {{
    --bg: #000000;
    --panel: #0f0f14;
    --panel-2: #15151c;
    --text: #f5f5f7;
    --muted: #8e8e93;
    --line: #25252b;
    --blue: #0a84ff;
    --bubble-them: #2c2c2e;
    --green: #30d158;
    --orange: #ff9f0a;
}}
* {{ box-sizing: border-box; }}
html, body {{ margin: 0; padding: 0; background: var(--bg); color: var(--text); font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Arial, sans-serif; }}
a {{ color: #9ecbff; text-decoration: none; }}
a:hover {{ text-decoration: underline; }}
.page {{ max-width: 1120px; margin: 0 auto; padding: 24px 18px 36px; }}
.hero {{ position: sticky; top: 0; z-index: 3; margin: -24px -18px 18px; padding: 18px 18px 14px; backdrop-filter: blur(18px); background: rgba(0, 0, 0, 0.82); border-bottom: 1px solid var(--line); }}
.hero h1 {{ margin: 0; font-size: 28px; font-weight: 700; }}
.hero p {{ margin: 6px 0 0; color: var(--muted); font-size: 13px; }}
.toolbar {{ display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-top: 10px; }}
.backlink {{ display: inline-flex; align-items: center; gap: 8px; padding: 8px 12px; border-radius: 999px; border: 1px solid var(--line); background: rgba(255,255,255,0.03); }}
.card-grid {{ display: grid; gap: 12px; }}
.card {{ display: block; padding: 14px 16px; border-radius: 18px; border: 1px solid var(--line); background: linear-gradient(180deg, rgba(255,255,255,0.03), rgba(255,255,255,0.015)); }}
.card h3 {{ margin: 0 0 8px; font-size: 16px; }}
.card .meta {{ color: var(--muted); font-size: 12px; margin-bottom: 8px; }}
.card .preview {{ color: var(--text); font-size: 14px; line-height: 1.45; white-space: pre-wrap; word-break: break-word; }}
.chat-shell {{ display: flex; flex-direction: column; gap: 12px; padding-top: 8px; }}
.day-label {{ align-self: center; padding: 6px 12px; border-radius: 999px; background: var(--panel-2); color: var(--muted); font-size: 12px; border: 1px solid var(--line); }}
.msg {{ display: flex; flex-direction: column; max-width: min(74%, 720px); }}
.msg.me {{ align-self: flex-end; }}
.msg.them {{ align-self: flex-start; }}
.bubble {{ padding: 11px 14px; border-radius: 22px; line-height: 1.45; font-size: 15px; white-space: pre-wrap; word-break: break-word; box-shadow: 0 8px 24px rgba(0,0,0,0.24); }}
.msg.me .bubble {{ background: var(--blue); color: #ffffff; border-bottom-right-radius: 8px; }}
.msg.them .bubble {{ background: var(--bubble-them); color: var(--text); border-bottom-left-radius: 8px; }}
.meta-row {{ margin-top: 5px; font-size: 12px; color: var(--muted); padding: 0 6px; }}
.call-list {{ display: grid; gap: 12px; }}
.call-item {{ padding: 14px 16px; border-radius: 18px; border: 1px solid var(--line); background: var(--panel); }}
.call-top {{ display: flex; justify-content: space-between; gap: 12px; align-items: center; flex-wrap: wrap; }}
.badge {{ display: inline-flex; align-items: center; gap: 6px; padding: 5px 10px; border-radius: 999px; font-size: 12px; border: 1px solid var(--line); }}
.badge.outgoing {{ background: rgba(10,132,255,0.16); color: #9ecbff; }}
.badge.incoming {{ background: rgba(48,209,88,0.14); color: #9bf0b5; }}
.badge.missed {{ background: rgba(255,159,10,0.16); color: #ffd08a; }}
.detail {{ margin-top: 8px; color: var(--muted); font-size: 13px; line-height: 1.45; }}
.kv {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 10px; margin-top: 10px; }}
.kv div {{ padding: 10px 12px; border-radius: 14px; background: rgba(255,255,255,0.03); border: 1px solid var(--line); }}
.kv b {{ display: block; margin-bottom: 4px; font-size: 12px; color: var(--muted); font-weight: 600; }}
@media (max-width: 780px) {{
    .page {{ padding: 18px 12px 28px; }}
    .hero {{ margin: -18px -12px 16px; padding: 14px 12px 12px; }}
    .msg {{ max-width: 88%; }}
}}
</style>
{extra_head}
</head>
<body>
<div class="page">
    <div class="hero">
        <h1>{html.escape(title)}</h1>
        <p>{html.escape(subtitle)}</p>
    </div>
    {body_html}
</div>
</body>
</html>
"""


def _write_text_file(path: Path, lines: list[str]):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("".join(lines), encoding="utf-8")


def _render_sms_conversation_page(title: str, subtitle: str, messages: list[dict], back_href: str) -> str:
    fragments: list[str] = [
        f'<div class="toolbar"><a class="backlink" href="{html.escape(back_href)}">← Back to conversations</a></div>',
        '<div class="chat-shell">',
    ]
    current_day = None
    for msg in messages:
        dt = msg.get("datetime") or ""
        day = dt[:10] if len(dt) >= 10 else "Unknown date"
        if day != current_day:
            fragments.append(f'<div class="day-label">{html.escape(day)}</div>')
            current_day = day
        side = "me" if msg.get("is_from_me") else "them"
        preview = html.escape(msg.get("text") or "")
        meta = " · ".join(part for part in [
            msg.get("time_only") or "Unknown time",
            msg.get("contact") or "",
            msg.get("service") or "",
        ] if part)
        fragments.append(
            f'<div class="msg {side}">'
            f'<div class="bubble">{preview}</div>'
            f'<div class="meta-row">{html.escape(meta)}</div>'
            f'</div>'
        )
    fragments.append("</div>")
    return _wrap_html_page(title, subtitle, "".join(fragments))


def _render_sms_index_page(title: str, subtitle: str, conversations: list[dict]) -> str:
    cards = ['<div class="card-grid">']
    for conv in conversations:
        cards.append(
            f'<a class="card" href="{html.escape(conv["href"])}">'
            f'<h3>{html.escape(conv["label"])}</h3>'
            f'<div class="meta">{html.escape(conv["meta"])}</div>'
            f'<div class="preview">{html.escape(conv["preview"])}</div>'
            f'</a>'
        )
    cards.append("</div>")
    return _wrap_html_page(title, subtitle, "".join(cards))


def export_sms_readable(sqlite_path: str | Path, out_dir: str | Path) -> int:
    sqlite_path = Path(sqlite_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(str(sqlite_path))
    conn.row_factory = sqlite3.Row
    try:
        if not _table_exists(conn, "message"):
            raise RuntimeError("sms.db does not contain a message table")

        message_cols = _table_columns(conn, "message")
        handle_cols = _table_columns(conn, "handle") if _table_exists(conn, "handle") else set()
        chat_cols = _table_columns(conn, "chat") if _table_exists(conn, "chat") else set()
        chat_join_exists = _table_exists(conn, "chat_message_join") and bool(chat_cols)

        select_parts = ["m.ROWID AS rowid"]
        joins: list[str] = []

        contact_expr = "'' AS contact"
        if "handle_id" in message_cols and handle_cols:
            joins.append("LEFT JOIN handle h ON h.ROWID = m.handle_id")
            handle_bits = [col for col in ("id", "uncanonicalized_id", "service") if col in handle_cols]
            if handle_bits:
                contact_expr = f"COALESCE({', '.join('h.' + col for col in handle_bits)}, '') AS contact"
        select_parts.append(contact_expr)

        if chat_join_exists:
            joins.append("LEFT JOIN chat_message_join cmj ON cmj.message_id = m.ROWID")
            joins.append("LEFT JOIN chat c ON c.ROWID = cmj.chat_id")
            if "chat_identifier" in chat_cols:
                select_parts.append("COALESCE(c.chat_identifier, '') AS chat_identifier")
            else:
                select_parts.append("'' AS chat_identifier")
            if "display_name" in chat_cols:
                select_parts.append("COALESCE(c.display_name, '') AS chat_display_name")
            else:
                select_parts.append("'' AS chat_display_name")
            if "service_name" in chat_cols:
                select_parts.append("COALESCE(c.service_name, '') AS chat_service_name")
            else:
                select_parts.append("'' AS chat_service_name")
        else:
            select_parts.extend([
                "'' AS chat_identifier",
                "'' AS chat_display_name",
                "'' AS chat_service_name",
            ])

        def col_or_blank(col: str, alias: str | None = None, default: str = "''") -> str:
            alias = alias or col
            if col in message_cols:
                return f"m.{col} AS {alias}"
            return f"{default} AS {alias}"

        for col in (
            "guid",
            "text",
            "subject",
            "service",
            "is_from_me",
            "date",
            "date_read",
            "date_delivered",
            "cache_has_attachments",
            "attributedBody",
        ):
            default = "X''" if col == "attributedBody" else "''"
            select_parts.append(col_or_blank(col, default=default))

        order_by = "m.ROWID"
        if "date" in message_cols:
            order_by = "m.date, m.ROWID"

        sql = f"SELECT {', '.join(select_parts)} FROM message m {' '.join(joins)} ORDER BY {order_by}"
        rows = conn.execute(sql).fetchall()

        csv_path = out_dir / "sms_messages.csv"
        txt_path = out_dir / "sms_messages.txt"
        html_index_path = out_dir / "sms_messages.html"
        conv_txt_dir = out_dir / "conversations"
        conv_html_dir = out_dir / "conversations_html"
        conv_txt_dir.mkdir(parents=True, exist_ok=True)
        conv_html_dir.mkdir(parents=True, exist_ok=True)

        conversations: dict[str, list[dict]] = {}

        with csv_path.open("w", encoding="utf-8-sig", newline="") as f_csv, txt_path.open("w", encoding="utf-8") as f_txt:
            writer = csv.writer(f_csv)
            writer.writerow([
                "rowid",
                "conversation",
                "datetime",
                "direction",
                "contact",
                "service",
                "text",
                "subject",
                "has_attachments",
            ])

            for row in rows:
                dt = _apple_time_to_str(row["date"])
                is_from_me = _is_truthy(row["is_from_me"])
                direction = "Me" if is_from_me else "Them"
                conversation = _conversation_label_for_row(row)
                contact = _normalize_text(row["contact"]) or "Unknown"
                service = _normalize_text(row["service"]) or _normalize_text(row["chat_service_name"]) or "SMS"
                subject = _normalize_text(row["subject"])
                text_val = _message_display_text(row)
                has_attachments = _is_truthy(row["cache_has_attachments"])

                writer.writerow([
                    row["rowid"],
                    conversation,
                    dt,
                    direction,
                    contact,
                    service,
                    text_val,
                    subject,
                    int(has_attachments),
                ])

                line = f"[{dt or 'Unknown time'}] {direction} | {conversation} | {contact} | {service}\n{text_val}\n\n"
                f_txt.write(line)

                time_only = dt[11:16] if len(dt) >= 16 else (dt or "Unknown time")
                conversations.setdefault(conversation, []).append(
                    {
                        "rowid": row["rowid"],
                        "datetime": dt,
                        "time_only": time_only,
                        "is_from_me": is_from_me,
                        "contact": contact,
                        "service": service,
                        "subject": subject,
                        "text": text_val,
                    }
                )

        written = 2
        conversation_cards: list[dict] = []

        for idx, (conversation, messages) in enumerate(sorted(conversations.items(), key=lambda item: (_message_sort_key(item[1][-1]) if item[1] else ("", 0))), start=1):
            messages.sort(key=_message_sort_key)
            slug = f"{idx:03d}_{_safe_filename(conversation, default='conversation')}"
            txt_lines: list[str] = []
            for msg in messages:
                direction = "Me" if msg["is_from_me"] else "Them"
                txt_lines.append(
                    f"[{msg['datetime'] or 'Unknown time'}] {direction} | {msg['contact']} | {msg['service']}\n{msg['text']}\n\n"
                )
            _write_text_file(conv_txt_dir / f"{slug}.txt", txt_lines)
            written += 1

            html_href = f"conversations_html/{slug}.html"
            html_doc = _render_sms_conversation_page(
                conversation,
                f"{len(messages)} message(s) · grouped by phone number / chat / service",
                messages,
                "../sms_messages.html",
            )
            (conv_html_dir / f"{slug}.html").write_text(html_doc, encoding="utf-8")
            written += 1

            last_msg = messages[-1]
            preview = last_msg["text"][:220]
            meta = " · ".join(part for part in [
                f"{len(messages)} msg",
                last_msg.get("datetime") or "Unknown time",
                last_msg.get("service") or "",
            ] if part)
            conversation_cards.append(
                {
                    "label": conversation,
                    "meta": meta,
                    "preview": preview,
                    "href": html_href,
                }
            )

        index_doc = _render_sms_index_page(
            "SMS & iMessage",
            "Readable export from sms.db. Conversations are split by phone number / chat / service instead of one long mixed thread.",
            conversation_cards,
        )
        html_index_path.write_text(index_doc, encoding="utf-8")
        written += 1

        return written
    finally:
        conn.close()


def _sqlite_table_names(conn: sqlite3.Connection) -> list[str]:
    return [row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").fetchall()]


def _sql_pick_column(table_alias: str, cols: set[str], aliases: tuple[str, ...], out_name: str, default: str = "''") -> str:
    for alias in aliases:
        if alias in cols:
            return f'{table_alias}."{alias}" AS {out_name}'
    return f"{default} AS {out_name}"


def _find_call_table(conn: sqlite3.Connection) -> str | None:
    tables = _sqlite_table_names(conn)
    preferred = [
        "ZCALLRECORD",
        "call",
        "Call",
        "call_record",
        "calls",
        "callhistory",
        "CallHistory",
    ]
    for name in preferred:
        if name in tables:
            return name
    for name in tables:
        if "call" in name.lower():
            return name
    return None


def _fetch_call_history_rows(conn: sqlite3.Connection) -> tuple[str, list[sqlite3.Row]]:
    table = _find_call_table(conn)
    if not table:
        raise RuntimeError("No supported call-history table was found in the SQLite file")

    cols = _table_columns(conn, table)
    select_parts = [
        _sql_pick_column("t", cols, ("ROWID", "rowid", "Z_PK"), "rowid", "NULL"),
        _sql_pick_column("t", cols, ("address", "ZADDRESS", "phone_number", "number", "remote_handle", "handle"), "address"),
        _sql_pick_column("t", cols, ("name", "localized_name", "display_name", "ZNAME"), "name"),
        _sql_pick_column("t", cols, ("date", "ZDATE", "timestamp", "start_date", "call_date"), "date", "0"),
        _sql_pick_column("t", cols, ("duration", "ZDURATION", "call_duration"), "duration", "0"),
        _sql_pick_column("t", cols, ("answered", "ZANSWERED"), "answered", "0"),
        _sql_pick_column("t", cols, ("originated", "ZORIGINATED", "outgoing"), "originated", "0"),
        _sql_pick_column("t", cols, ("flags", "ZFLAGS"), "flags", "0"),
        _sql_pick_column("t", cols, ("calltype", "call_type", "ZCALLTYPE"), "calltype"),
        _sql_pick_column("t", cols, ("service_provider", "service", "provider", "ZSERVICE_PROVIDER"), "service"),
        _sql_pick_column("t", cols, ("location", "ZLOCATION"), "location"),
        _sql_pick_column("t", cols, ("iso_country_code", "country_code", "ZISO_COUNTRY_CODE"), "country_code"),
        _sql_pick_column("t", cols, ("disconnected_cause", "ZDISCONNECTED_CAUSE"), "disconnect_cause"),
        _sql_pick_column("t", cols, ("read", "ZREAD"), "is_read"),
    ]

    sql = f'SELECT {", ".join(select_parts)} FROM "{table}" t ORDER BY date, rowid'
    rows = conn.execute(sql).fetchall()
    return table, rows


def _call_direction(row: sqlite3.Row | dict) -> str:
    originated = _is_truthy(row["originated"]) if "originated" in row.keys() else False
    answered = _is_truthy(row["answered"]) if "answered" in row.keys() else False
    try:
        duration = int(float(row["duration"] or 0))
    except Exception:
        duration = 0

    if originated:
        return "Outgoing"
    if not answered and duration <= 0:
        return "Missed"
    return "Incoming"


def _call_service_label(row: sqlite3.Row | dict) -> str:
    raw_service = _normalize_text(row["service"]) if "service" in row.keys() else ""
    if raw_service:
        return raw_service
    raw_type = _normalize_text(row["calltype"]) if "calltype" in row.keys() else ""
    known = {
        "8": "FaceTime Audio",
        "16": "FaceTime Video",
    }
    if raw_type in known:
        return known[raw_type]
    if raw_type:
        return f"Type {raw_type}"
    return "Call"


def _call_contact_label(row: sqlite3.Row | dict) -> str:
    name = _normalize_text(row["name"]) if "name" in row.keys() else ""
    address = _normalize_text(row["address"]) if "address" in row.keys() else ""
    if name and address and name != address:
        return f"{name} ({address})"
    return name or address or "Unknown caller"


def _render_call_conversation_page(title: str, subtitle: str, rows: list[dict], back_href: str) -> str:
    fragments: list[str] = [
        f'<div class="toolbar"><a class="backlink" href="{html.escape(back_href)}">← Back to calls</a></div>',
        '<div class="call-list">',
    ]
    for row in rows:
        direction = row["direction"].lower()
        badge_class = "outgoing" if direction == "outgoing" else ("missed" if direction == "missed" else "incoming")
        meta_top = row.get("datetime") or "Unknown time"
        details = []
        if row.get("service"):
            details.append(f'Service: {html.escape(row["service"])}')
        if row.get("location"):
            details.append(f'Location: {html.escape(row["location"])}')
        if row.get("country_code"):
            details.append(f'Country: {html.escape(row["country_code"])}')
        if row.get("disconnect_cause"):
            details.append(f'Disconnect cause: {html.escape(row["disconnect_cause"])}')
        detail_html = "<br>".join(details)
        fragments.append(
            f'<div class="call-item">'
            f'<div class="call-top"><div><strong>{html.escape(row["contact"])}</strong><div class="detail">{html.escape(meta_top)}</div></div>'
            f'<span class="badge {badge_class}">{html.escape(row["direction"])} · {html.escape(row["duration_text"])}</span></div>'
            f'<div class="kv">'
            f'<div><b>Address</b>{html.escape(row["address"] or row["contact"])}</div>'
            f'<div><b>Call Type</b>{html.escape(row["service"])}</div>'
            f'<div><b>Answered</b>{"Yes" if row["answered"] else "No"}</div>'
            f'</div>'
            f'{"<div class=\"detail\">" + detail_html + "</div>" if detail_html else ""}'
            f'</div>'
        )
    fragments.append("</div>")
    return _wrap_html_page(title, subtitle, "".join(fragments))


def _render_call_index_page(title: str, subtitle: str, cards: list[dict]) -> str:
    tiles = ['<div class="card-grid">']
    for card in cards:
        tiles.append(
            f'<a class="card" href="{html.escape(card["href"])}">'
            f'<h3>{html.escape(card["label"])}</h3>'
            f'<div class="meta">{html.escape(card["meta"])}</div>'
            f'<div class="preview">{html.escape(card["preview"])}</div>'
            f'</a>'
        )
    tiles.append("</div>")
    return _wrap_html_page(title, subtitle, "".join(tiles))


def export_call_history_readable(sqlite_path: str | Path, out_dir: str | Path) -> int:
    sqlite_path = Path(sqlite_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(str(sqlite_path))
    conn.row_factory = sqlite3.Row
    try:
        table_name, rows = _fetch_call_history_rows(conn)

        csv_path = out_dir / "call_history.csv"
        txt_path = out_dir / "call_history.txt"
        html_index_path = out_dir / "call_history.html"
        conv_txt_dir = out_dir / "conversations"
        conv_html_dir = out_dir / "conversations_html"
        conv_txt_dir.mkdir(parents=True, exist_ok=True)
        conv_html_dir.mkdir(parents=True, exist_ok=True)

        grouped: dict[str, list[dict]] = {}

        with csv_path.open("w", encoding="utf-8-sig", newline="") as f_csv, txt_path.open("w", encoding="utf-8") as f_txt:
            writer = csv.writer(f_csv)
            writer.writerow([
                "rowid",
                "datetime",
                "contact",
                "address",
                "direction",
                "duration_seconds",
                "duration_text",
                "service",
                "answered",
                "location",
                "country_code",
                "disconnect_cause",
                "source_table",
            ])

            for row in rows:
                dt = _apple_time_to_str(row["date"])
                contact = _call_contact_label(row)
                address = _normalize_text(row["address"]) if "address" in row.keys() else ""
                direction = _call_direction(row)
                service = _call_service_label(row)
                location = _normalize_text(row["location"]) if "location" in row.keys() else ""
                country_code = _normalize_text(row["country_code"]) if "country_code" in row.keys() else ""
                disconnect_cause = _normalize_text(row["disconnect_cause"]) if "disconnect_cause" in row.keys() else ""
                answered = _is_truthy(row["answered"]) if "answered" in row.keys() else False
                duration_text = _duration_to_str(row["duration"])

                writer.writerow([
                    row["rowid"],
                    dt,
                    contact,
                    address,
                    direction,
                    row["duration"] or 0,
                    duration_text,
                    service,
                    int(answered),
                    location,
                    country_code,
                    disconnect_cause,
                    table_name,
                ])

                line = f"[{dt or 'Unknown time'}] {direction} | {contact} | {service} | {duration_text}\n"
                if address:
                    line += f"Address: {address}\n"
                if location:
                    line += f"Location: {location}\n"
                if disconnect_cause:
                    line += f"Disconnect cause: {disconnect_cause}\n"
                line += "\n"
                f_txt.write(line)

                grouped.setdefault(contact, []).append(
                    {
                        "rowid": row["rowid"],
                        "datetime": dt,
                        "contact": contact,
                        "address": address,
                        "direction": direction,
                        "duration_text": duration_text,
                        "service": service,
                        "answered": answered,
                        "location": location,
                        "country_code": country_code,
                        "disconnect_cause": disconnect_cause,
                    }
                )

        written = 2
        cards: list[dict] = []

        sorted_items = sorted(
            grouped.items(),
            key=lambda item: ((item[1][-1]["datetime"] if item[1] else "") or "", item[0]),
        )

        for idx, (contact, entries) in enumerate(sorted_items, start=1):
            entries.sort(key=lambda item: ((item["datetime"] or "9999-99-99 99:99:99"), int(item["rowid"] or 0)))
            slug = f"{idx:03d}_{_safe_filename(contact, default='call_history')}"
            txt_lines: list[str] = []
            for entry in entries:
                txt_lines.append(
                    f"[{entry['datetime'] or 'Unknown time'}] {entry['direction']} | {entry['service']} | {entry['duration_text']}\n"
                    f"Address: {entry['address'] or entry['contact']}\n\n"
                )
            _write_text_file(conv_txt_dir / f"{slug}.txt", txt_lines)
            written += 1

            page = _render_call_conversation_page(
                contact,
                f"{len(entries)} call(s) · readable export from {table_name}",
                entries,
                "../call_history.html",
            )
            (conv_html_dir / f"{slug}.html").write_text(page, encoding="utf-8")
            written += 1

            last = entries[-1]
            preview = f'{last["direction"]} · {last["service"]} · {last["duration_text"]}'
            meta = " · ".join(part for part in [f"{len(entries)} call(s)", last.get("datetime") or "Unknown time"] if part)
            cards.append(
                {
                    "label": contact,
                    "meta": meta,
                    "preview": preview,
                    "href": f"conversations_html/{slug}.html",
                }
            )

        index_doc = _render_call_index_page(
            "Call History",
            f"Readable export from {table_name}. Calls are grouped by phone number / contact so you do not have to browse the raw database.",
            cards,
        )
        html_index_path.write_text(index_doc, encoding="utf-8")
        written += 1

        return written
    finally:
        conn.close()


def _db_value_to_text(column: str, value) -> str:
    if value is None:
        return ""
    if isinstance(value, memoryview):
        value = bytes(value)
    if isinstance(value, (bytes, bytearray)):
        raw = bytes(value)
        extracted = _extract_text_from_attributed_body(raw)
        if extracted:
            return extracted
        try:
            decoded = raw.decode("utf-8", errors="ignore")
            decoded = _normalize_text(decoded)
            if decoded and sum(ch.isprintable() for ch in decoded) >= max(4, int(len(decoded) * 0.8)):
                return decoded[:2000] + ("…" if len(decoded) > 2000 else "")
        except Exception:
            pass
        return f"<BLOB {len(raw)} bytes>"

    text = _normalize_text(value)
    if not text:
        return ""

    if re.search(r"(date|time|timestamp|created|modified|updated|received|start)", column, re.IGNORECASE):
        converted = _apple_time_to_str(value)
        if converted and converted != str(value):
            return converted

    return text[:2000] + ("…" if len(text) > 2000 else "")


def _sqlite_user_tables(conn: sqlite3.Connection) -> list[str]:
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
    ).fetchall()
    return [row[0] for row in rows]


def _render_db_table_preview_page(title: str, subtitle: str, rows: list[dict], back_href: str) -> str:
    fragments: list[str] = [
        f'<div class="toolbar"><a class="backlink" href="{html.escape(back_href)}">← Back to voicemail</a></div>',
        '<div class="call-list">',
    ]
    if not rows:
        fragments.append('<div class="call-item"><div class="detail">This table has no rows.</div></div>')
    for idx, row in enumerate(rows, start=1):
        kv = []
        for key, value in row.items():
            kv.append(f'<div><b>{html.escape(str(key))}</b>{html.escape(str(value))}</div>')
        fragments.append(
            f'<div class="call-item">'
            f'<div class="call-top"><div><strong>Row {idx}</strong></div></div>'
            f'<div class="kv">{"".join(kv)}</div>'
            f'</div>'
        )
    fragments.append('</div>')
    return _wrap_html_page(title, subtitle, ''.join(fragments))


def _render_voicemail_index_page(title: str, subtitle: str, audio_entries: list[dict], table_cards: list[dict], note: str = "") -> str:
    body: list[str] = []
    extra_head = """
<style>
.section-title { margin: 22px 0 10px; font-size: 18px; font-weight: 700; }
.notice { padding: 14px 16px; border-radius: 16px; border: 1px solid var(--line); background: rgba(255,255,255,0.03); color: var(--muted); margin-bottom: 14px; }
audio { width: 100%; margin-top: 12px; }
.pathline { color: var(--muted); font-size: 12px; margin-top: 8px; word-break: break-all; }
</style>
"""
    if note:
        body.append(f'<div class="notice">{html.escape(note)}</div>')
    body.append('<div class="section-title">Audio files</div>')
    if audio_entries:
        body.append('<div class="card-grid">')
        for entry in audio_entries:
            meta = " · ".join(part for part in [entry.get("modified") or "", entry.get("extension") or "", entry.get("size_text") or ""] if part)
            body.append(
                f'<div class="card">'
                f'<h3>{html.escape(entry["name"])}</h3>'
                f'<div class="meta">{html.escape(meta)}</div>'
                f'<div class="preview">{html.escape(entry["relative_path"])}</div>'
                f'<audio controls preload="none" src="{html.escape(entry["uri"])}"></audio>'
                f'<div class="pathline"><a href="{html.escape(entry["uri"])}">Open audio file</a></div>'
                f'</div>'
            )
        body.append('</div>')
    else:
        body.append('<div class="notice">No voicemail audio files were found in the extracted backup.</div>')

    body.append('<div class="section-title">Database previews</div>')
    if table_cards:
        body.append('<div class="card-grid">')
        for card in table_cards:
            body.append(
                f'<a class="card" href="{html.escape(card["href"])}">'
                f'<h3>{html.escape(card["label"])}</h3>'
                f'<div class="meta">{html.escape(card["meta"])}</div>'
                f'<div class="preview">{html.escape(card["preview"])}</div>'
                f'</a>'
            )
        body.append('</div>')
    else:
        body.append('<div class="notice">No readable SQLite voicemail tables were found. If you only see audio files, open them directly from the list above.</div>')

    return _wrap_html_page(title, subtitle, ''.join(body), extra_head=extra_head)


def _advanced_capability(group: str, preview_type: str) -> str:
    base = {
        "Keychain / Passwords": "Found in backup. Plaintext secrets may still be protected.",
        "Wi-Fi / Network": "Configuration data found. Some secrets may remain protected.",
        "Health": "Health databases/config files found. Schema preview is basic.",
    }.get(group, "File copied from backup.")
    if preview_type == "sqlite":
        return base + " SQLite tables were summarized."
    if preview_type == "plist":
        return base + " Top-level plist keys were summarized."
    if preview_type == "text":
        return base + " Plain text preview available."
    return base



def _advanced_file_preview(path: Path) -> tuple[str, str, str]:
    suffix = path.suffix.lower()
    if suffix in {".db", ".sqlite", ".sqlite3", ".storedata"}:
        try:
            conn = sqlite3.connect(str(path))
            try:
                tables = _sqlite_user_tables(conn)
                if not tables:
                    return ("SQLite file with no user tables detected", "sqlite", "Database opened successfully")
                bits = []
                for table in tables[:6]:
                    try:
                        count = conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
                    except Exception:
                        count = "?"
                    bits.append(f"{table} ({count})")
                more = " …" if len(tables) > 6 else ""
                return ("Tables: " + ", ".join(bits) + more, "sqlite", f"{len(tables)} table(s) detected")
            finally:
                conn.close()
        except Exception as exc:
            return (f"Database file copied; preview failed ({exc})", "binary", "Copied only")

    if suffix in {".plist", ".xml"}:
        try:
            obj = plistlib.loads(path.read_bytes())
            if isinstance(obj, dict):
                keys = [str(k) for k in list(obj.keys())[:8]]
                more = " …" if len(obj) > 8 else ""
                return ("Top-level keys: " + ", ".join(keys) + more, "plist", f"{len(obj)} key(s) detected")
            if isinstance(obj, list):
                return (f"Top-level array with {len(obj)} item(s)", "plist", "Parsed plist array")
            return (f"Top-level {type(obj).__name__}", "plist", "Parsed plist value")
        except Exception:
            text_sample = _normalize_text(path.read_text(encoding="utf-8", errors="ignore"))
            if text_sample:
                return (text_sample[:220] + ("…" if len(text_sample) > 220 else ""), "text", "UTF-8 preview available")
            return ("PLIST/XML file copied; preview not available", "binary", "Copied only")

    if suffix in {".json", ".txt", ".log"}:
        sample = _normalize_text(path.read_text(encoding="utf-8", errors="ignore"))
        if sample:
            return (sample[:220] + ("…" if len(sample) > 220 else ""), "text", "UTF-8 preview available")
        return ("Text-like file copied", "text", "Copied only")

    return ("Binary file copied", "binary", "Copied only")



def _render_advanced_data_index_page(title: str, subtitle: str, grouped_cards: dict[str, list[dict]], note: str = "") -> str:
    body: list[str] = []
    extra_head = """
<style>
.notice { padding: 14px 16px; border-radius: 16px; border: 1px solid var(--line); background: rgba(255,255,255,0.03); color: var(--muted); margin-bottom: 14px; }
.section-title { margin: 20px 0 10px; font-size: 18px; font-weight: 700; }
.pathline { color: var(--muted); font-size: 12px; margin-top: 8px; word-break: break-all; }
.capline { color: var(--text); font-size: 12px; margin-top: 8px; }
</style>
"""
    if note:
        body.append(f'<div class="notice">{html.escape(note)}</div>')
    if not grouped_cards:
        body.append('<div class="notice">No Keychain / Wi-Fi / Health files were matched. This often means the backup is unencrypted or that these datasets are not present in this backup.</div>')
    for group, cards in grouped_cards.items():
        body.append(f'<div class="section-title">{html.escape(group)}</div>')
        body.append('<div class="card-grid">')
        for card in cards:
            body.append(
                f'<div class="card">'
                f'<h3>{html.escape(card["name"])}</h3>'
                f'<div class="meta">{html.escape(card["meta"])}</div>'
                f'<div class="preview">{html.escape(card["preview"])}</div>'
                f'<div class="capline">{html.escape(card["capability"])}</div>'
                f'<div class="pathline">{html.escape(card["relative_path"])}</div>'
                f'</div>'
            )
        body.append('</div>')
    return _wrap_html_page(title, subtitle, ''.join(body), extra_head=extra_head)



def export_advanced_data_report(root_dir: str | Path, out_dir: str | Path) -> int:
    root_dir = Path(root_dir)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    csv_path = out_dir / "advanced_data_files.csv"
    txt_path = out_dir / "advanced_data_summary.txt"
    html_path = out_dir / "advanced_data.html"

    files = sorted((p for p in root_dir.rglob('*') if p.is_file()), key=lambda p: str(p).lower())
    grouped_cards: dict[str, list[dict]] = {}
    summary_lines: list[str] = []

    with csv_path.open('w', encoding='utf-8-sig', newline='') as f_csv:
        writer = csv.writer(f_csv)
        writer.writerow([
            'group', 'domain', 'relative_path', 'file_name', 'extension', 'size_bytes', 'modified_time', 'preview', 'capability'
        ])

        for path in files:
            rel = path.relative_to(root_dir)
            parts = rel.parts
            if len(parts) >= 2:
                domain = parts[0]
                rel_no_domain = '/'.join(parts[1:])
            else:
                domain = ''
                rel_no_domain = str(rel).replace('\\', '/')

            group = advanced_data_group(domain, rel_no_domain)
            if not group:
                continue

            preview, preview_type, preview_meta = _advanced_file_preview(path)
            capability = _advanced_capability(group, preview_type)
            stat = path.stat()
            modified = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            ext = path.suffix.lower() or '(no extension)'
            meta = ' · '.join(part for part in [domain or 'No domain', ext, _human_file_size(stat.st_size), modified] if part)
            card = {
                'name': path.name,
                'meta': meta,
                'preview': preview if preview else preview_meta,
                'capability': capability,
                'relative_path': f'{domain}/{rel_no_domain}' if domain else rel_no_domain,
            }
            grouped_cards.setdefault(group, []).append(card)
            writer.writerow([
                group,
                domain,
                rel_no_domain,
                path.name,
                ext,
                stat.st_size,
                modified,
                preview,
                capability,
            ])
            summary_lines.append(f'[{group}] {domain}/{rel_no_domain} | {preview_meta} | {capability}')

    for cards in grouped_cards.values():
        cards.sort(key=lambda item: item['name'].lower())

    note = (
        'Advanced Data tries to collect Keychain, Wi-Fi, and Health related files from the backup. '
        'This report is intentionally basic: it tells you what was found, what group it belongs to, and whether a lightweight preview was possible. '
        'It does not guarantee plaintext password extraction.'
    )
    html_path.write_text(
        _render_advanced_data_index_page(
            'Advanced Data',
            'Basic report for Keychain / Wi-Fi / Health related files found in the backup.',
            grouped_cards,
            note=note,
        ),
        encoding='utf-8',
    )
    txt_path.write_text(('\n'.join(summary_lines) if summary_lines else 'No matching advanced data files were found.') + '\n', encoding='utf-8')
    return 3


def _human_file_size(num_bytes: int) -> str:
    value = float(num_bytes or 0)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if value < 1024 or unit == "TB":
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024.0
    return f"{int(num_bytes)} B"





def _render_contacts_index_page(title: str, subtitle: str, contacts: list[dict], table_cards: list[dict], note: str = "") -> str:
    def render_value_html(text: str) -> str:
        text = _normalize_text(text)
        if not text:
            return ''
        pieces = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            href = ''
            display = line
            if line.startswith(('http://', 'https://')):
                href = line
            elif line.startswith('www.'):
                href = 'https://' + line
            if href:
                pieces.append(f'<a href="{html.escape(href)}">{html.escape(display)}</a>')
            else:
                pieces.append(html.escape(display))
        return '<br>'.join(pieces)

    def split_labeled_value(item: str) -> tuple[str, str]:
        item = _normalize_text(item)
        if not item:
            return ('', '')
        if item.startswith(('http://', 'https://', 'www.')):
            return ('', item)
        match = re.match(r'^([^:\n]{1,80}):\s*(.+)$', item, flags=re.DOTALL)
        if not match:
            return ('', item)
        label = _normalize_contact_label(match.group(1))
        value = _normalize_text(match.group(2))
        if not label or not value or '://' in label:
            return ('', item)
        return (label, value)

    body: list[str] = []
    extra_head = """
<style>
.section-title { margin: 22px 0 10px; font-size: 18px; font-weight: 700; }
.notice { padding: 14px 16px; border-radius: 16px; border: 1px solid var(--line); background: rgba(255,255,255,0.03); color: var(--muted); margin-bottom: 14px; }
.contacts-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 16px; }
.contact-card { border: 1px solid var(--line); background: rgba(255,255,255,0.03); border-radius: 22px; padding: 18px; }
.contact-top { display: flex; gap: 14px; align-items: flex-start; margin-bottom: 14px; }
.contact-avatar { width: 72px; height: 72px; border-radius: 50%; object-fit: cover; border: 1px solid var(--line); background: rgba(255,255,255,0.05); flex: 0 0 auto; }
.contact-avatar-fallback { width: 72px; height: 72px; border-radius: 50%; border: 1px solid var(--line); background: rgba(255,255,255,0.05); color: var(--muted); display: flex; align-items: center; justify-content: center; font-size: 26px; flex: 0 0 auto; }
.contact-name { font-size: 24px; font-weight: 800; line-height: 1.15; margin: 0; }
.contact-meta { margin-top: 6px; color: var(--muted); font-size: 13px; }
.contact-sections { display: grid; gap: 10px; }
.contact-row { border-top: 1px solid rgba(255,255,255,0.06); padding-top: 10px; }
.contact-label { font-size: 12px; color: var(--muted); letter-spacing: 0.02em; margin-bottom: 4px; }
.contact-value { color: var(--text); font-size: 15px; line-height: 1.5; white-space: pre-wrap; word-break: break-word; }
.contact-value a { color: #8ab4ff; text-decoration: none; }
.contact-value a:hover { text-decoration: underline; }
.small-muted { color: var(--muted); font-size: 12px; margin-top: 8px; }
</style>
"""
    if note:
        body.append(f'<div class="notice">{html.escape(note)}</div>')

    body.append('<div class="section-title">Contacts</div>')
    if contacts:
        body.append('<div class="contacts-grid">')
        for contact in contacts:
            avatar_rel = contact.get('avatar_rel_path') or ''
            if avatar_rel:
                avatar_html = f'<img class="contact-avatar" src="{html.escape(avatar_rel)}" alt="avatar">'
            else:
                initial = (contact.get('display_name') or '?').strip()[:1].upper() or '?'
                avatar_html = f'<div class="contact-avatar-fallback">{html.escape(initial)}</div>'

            meta_parts = [part for part in [contact.get('organization') or '', contact.get('job_title') or '', contact.get('source_db') or ''] if part]
            rows: list[str] = []

            for key, default_label in (
                ('phones', 'Phone'),
                ('emails', 'Email'),
                ('urls', 'URL'),
                ('addresses', 'Address'),
                ('extras', ''),
                ('birthday', 'Birthday'),
                ('note', 'Notes'),
                ('department', 'Department'),
            ):
                value = contact.get(key, '')
                if isinstance(value, list):
                    cleaned_values = [_normalize_text(item) for item in value if _normalize_text(item)]
                    for item in cleaned_values:
                        label, raw_value = split_labeled_value(item)
                        row_label = label or default_label or 'Other'
                        rendered_value = render_value_html(raw_value)
                        if not rendered_value:
                            continue
                        rows.append(
                            f'<div class="contact-row">'
                            f'<div class="contact-label">{html.escape(row_label)}</div>'
                            f'<div class="contact-value">{rendered_value}</div>'
                            f'</div>'
                        )
                else:
                    text = _normalize_text(value)
                    if not text:
                        continue
                    rendered_value = render_value_html(text)
                    if not rendered_value:
                        continue
                    rows.append(
                        f'<div class="contact-row">'
                        f'<div class="contact-label">{html.escape(default_label or "Other")}</div>'
                        f'<div class="contact-value">{rendered_value}</div>'
                        f'</div>'
                    )

            if not rows:
                rows.append('<div class="contact-row"><div class="contact-value">No readable phone / email / address fields found for this contact.</div></div>')

            body.append(
                f'<div class="contact-card">'
                f'<div class="contact-top">'
                f'{avatar_html}'
                f'<div>'
                f'<h3 class="contact-name">{html.escape(contact.get("display_name") or "Unnamed Contact")}</h3>'
                f'<div class="contact-meta">{html.escape(" · ".join(meta_parts))}</div>'
                f'</div>'
                f'</div>'
                f'<div class="contact-sections">{"".join(rows)}</div>'
                f'</div>'
            )
        body.append('</div>')
    else:
        body.append('<div class="contact-card"><div class="contact-value">No contacts could be parsed from the extracted contact database.</div></div>')

    if table_cards:
        body.append('<div class="section-title">Database table previews</div>')
        body.append('<div class="card-grid">')
        for card in table_cards:
            body.append(
                f'<a class="card" href="{html.escape(card["href"])}">'
                f'<h3>{html.escape(card["label"])}</h3>'
                f'<div class="meta">{html.escape(card["meta"])}</div>'
                f'<div class="preview">{html.escape(card["preview"])}</div>'
                f'</a>'
            )
        body.append('</div>')

    return _wrap_html_page(title, subtitle, ''.join(body), extra_head=extra_head)


def _render_contacts_table_preview_page(title: str, subtitle: str, rows: list[dict], back_href: str) -> str:
    fragments: list[str] = [
        f'<div class="toolbar"><a class="backlink" href="{html.escape(back_href)}">← Back to contacts</a></div>',
        '<div class="call-list">',
    ]
    if not rows:
        fragments.append('<div class="call-item"><div class="detail">This table has no rows.</div></div>')
    for idx, row in enumerate(rows, start=1):
        kv = []
        for key, value in row.items():
            kv.append(f'<div><b>{html.escape(str(key))}</b>{html.escape(str(value))}</div>')
        fragments.append(
            f'<div class="call-item">'
            f'<div class="call-top"><div><strong>Row {idx}</strong></div></div>'
            f'<div class="kv">{"".join(kv)}</div>'
            f'</div>'
        )
    fragments.append('</div>')
    return _wrap_html_page(title, subtitle, ''.join(fragments))


def _dedupe_preserve(values: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for value in values:
        cleaned = _normalize_text(value)
        if not cleaned:
            continue
        key = cleaned.casefold()
        if key not in seen:
            seen.add(key)
            out.append(cleaned)
    return out


def _contact_blank() -> dict:
    return {
        'record_id': None,
        'display_name': '',
        'organization': '',
        'job_title': '',
        'department': '',
        'birthday': '',
        'note': '',
        'phones': [],
        'emails': [],
        'urls': [],
        'addresses': [],
        'extras': [],
        'avatar_rel_path': '',
        'source_db': '',
    }


def _row_pick(row: sqlite3.Row, *needles: str) -> str:
    for key in row.keys():
        low = key.lower()
        for needle in needles:
            n = needle.lower()
            if low == n or low.endswith(n):
                value = _db_value_to_text(key, row[key])
                if value:
                    return value
    return ''


def _looks_like_contact_blob_noise(text: str) -> bool:
    text = _normalize_text(text)
    if not text:
        return False
    if text.startswith('<BLOB '):
        return True
    if len(text) > 100 and re.fullmatch(r'[A-Za-z0-9+/=\-_]+', text):
        return True
    if len(text) > 120:
        letters = sum(ch.isalpha() for ch in text)
        digits = sum(ch.isdigit() for ch in text)
        spaces = sum(ch.isspace() for ch in text)
        safe = letters + digits + spaces + sum(ch in '._-/:@,#;()[]{}+?&=%' for ch in text)
        if safe / max(1, len(text)) < 0.82:
            return True
    if 'addressbook.sqlitedb' in text.lower():
        return True
    return False


def _contact_is_image_bytes(raw: bytes) -> bool:
    if not raw:
        return False
    return (
        raw.startswith(b'\xff\xd8\xff')
        or raw.startswith(b'\x89PNG\r\n\x1a\n')
        or raw.startswith(b'GIF87a')
        or raw.startswith(b'GIF89a')
        or raw.startswith(b'RIFF') and b'WEBP' in raw[:16]
        or (len(raw) > 12 and raw[4:8] == b'ftyp')
    )


def _contact_image_extension(raw: bytes) -> str:
    if raw.startswith(b'\xff\xd8\xff'):
        return '.jpg'
    if raw.startswith(b'\x89PNG\r\n\x1a\n'):
        return '.png'
    if raw.startswith((b'GIF87a', b'GIF89a')):
        return '.gif'
    if raw.startswith(b'RIFF') and b'WEBP' in raw[:16]:
        return '.webp'
    if len(raw) > 12 and raw[4:8] == b'ftyp':
        return '.heic'
    return '.bin'


def _flatten_contact_plist(obj) -> str:
    if obj is None:
        return ''
    if isinstance(obj, dict):
        addr_order = ['Street', 'street', 'Address', 'address', 'ZIP', 'zip', 'PostalCode', 'postalCode', 'City', 'city', 'State', 'state', 'Country', 'country']
        lines = []
        for key in addr_order:
            value = obj.get(key)
            if value:
                lines.append(_normalize_text(value))
        if lines:
            return '\n'.join(line for line in lines if line)
        parts = []
        for key, value in obj.items():
            flat = _flatten_contact_plist(value)
            if not flat:
                continue
            if isinstance(value, (dict, list, tuple)):
                parts.append(flat)
            else:
                parts.append(f'{key}: {flat}')
        return '\n'.join(parts)
    if isinstance(obj, (list, tuple, set)):
        return '\n'.join(part for item in obj if (part := _flatten_contact_plist(item)))
    return _normalize_text(obj)


def _contact_value_from_raw(column: str, value) -> str:
    if value is None:
        return ''
    if isinstance(value, memoryview):
        value = bytes(value)
    if isinstance(value, (bytes, bytearray)):
        raw = bytes(value)
        if not raw:
            return ''
        if _contact_is_image_bytes(raw):
            return ''
        try:
            obj = plistlib.loads(raw)
            flattened = _normalize_text(_flatten_contact_plist(obj))
            if flattened and not _looks_like_contact_blob_noise(flattened):
                return flattened[:4000]
        except Exception:
            pass
        extracted = _extract_text_from_attributed_body(raw)
        if extracted and not _looks_like_contact_blob_noise(extracted):
            return extracted[:4000]
        for enc in ('utf-8', 'utf-16', 'utf-16le', 'utf-16be'):
            try:
                decoded = raw.decode(enc, errors='ignore')
            except Exception:
                continue
            decoded = _normalize_text(decoded)
            if decoded and not _looks_like_contact_blob_noise(decoded):
                return decoded[:4000]
        return ''
    text = _normalize_text(value)
    if not text or _looks_like_contact_blob_noise(text):
        return ''
    if re.search(r'(date|time|timestamp|created|modified|updated|received|start|birthday)', column, re.IGNORECASE):
        converted = _apple_time_to_str(value)
        if converted and converted != str(value):
            text = converted
    return text[:4000]


def _sanitize_contact_field(value: str) -> str:
    value = _normalize_text(value)
    if not value:
        return ''
    if _looks_like_contact_blob_noise(value):
        return ''
    return value[:4000]


def _normalize_contact_label(label: str, field: str = '') -> str:
    label = _sanitize_contact_field(label)
    if not label:
        return ''
    label = re.sub(r'^_+\$?!?<([^>]+)>!?\$_+$', r'\1', label)
    label = re.sub(r'^_+\$?!?<([^>]+)>!?\$_$', r'\1', label)
    label = re.sub(r'^_+\$?!(.*?)!\$_+$', r'\1', label)
    label = label.replace('_$!<', '').replace('>!$_', '').replace('_$!', '').replace('!$_', '')
    label = label.replace('kAB', '').replace('Label', '').replace('HOME', 'Home').replace('WORK', 'Work')
    label = re.sub(r'^[^A-Za-z]+', '', label)
    label = re.sub(r'\s+', ' ', label).strip(' :;,.\t\n\r')
    if not label:
        return ''
    if re.fullmatch(r'\d+(?:\.\d+)?', label):
        return ''
    generic = {
        'phone': '', 'email': '', 'url': '', 'urls': '', 'address': '', 'birthday': '', 'note': '',
        'homepage': 'homepage', 'website': 'website', 'home page': 'homepage',
        'mobile': 'mobile', 'iphone': 'iPhone', 'main': 'main', 'work': 'work', 'home': 'home', 'other': 'other',
        'facebook': 'Facebook', 'facebook messenger': 'Facebook Messenger', 'messenger': 'Messenger',
        'instagram': 'Instagram', 'twitter': 'Twitter', 'telegram': 'Telegram', 'zalo': 'Zalo',
        'linkedin': 'LinkedIn', 'signal': 'Signal', 'whatsapp': 'WhatsApp',
        'ringtone': 'Ringtone', 'text tone': 'Text Tone', 'sound': 'Sound',
    }
    low = label.casefold()
    if low in generic:
        return generic[low]
    if field == 'urls' and low in {'home', 'homepage', 'website', 'work', 'other'}:
        return 'homepage' if low in {'home', 'homepage'} else low
    if field == 'emails' and low in {'home', 'work', 'other', 'icloud', 'gmail'}:
        return low
    if field == 'phones' and low in {'home', 'work', 'other', 'mobile', 'main', 'iphone'}:
        return generic.get(low, low)
    return label[:120]


def _format_contact_birthday(value) -> str:
    if value in (None, '', 0, '0'):
        return ''
    if isinstance(value, str):
        txt = _normalize_text(value)
        if not txt:
            return ''
        if re.fullmatch(r'-?\d+(?:\.\d+)?', txt):
            try:
                value = float(txt)
            except Exception:
                return txt
        else:
            for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d %B %Y', '%d %b %Y'):
                try:
                    dt = datetime.strptime(txt, fmt)
                    return f'{dt.day} {dt.strftime("%B %Y")}' if dt.year != 1900 else f'{dt.day} {dt.strftime("%B")}'
                except Exception:
                    pass
            return txt
    try:
        raw = float(value)
    except Exception:
        return _normalize_text(value)
    abs_raw = abs(raw)
    if abs_raw > 10**15:
        seconds = raw / 1_000_000_000
    elif abs_raw > 10**12:
        seconds = raw / 1_000_000
    elif abs_raw > 10**10:
        seconds = raw / 1000
    else:
        seconds = raw
    try:
        dt = datetime(2001, 1, 1) + timedelta(seconds=seconds)
        if dt.year <= 1904:
            return f'{dt.day} {dt.strftime("%B")}'
        return f'{dt.day} {dt.strftime("%B %Y")}'
    except Exception:
        return _normalize_text(value)


def _append_contact_value(contact: dict, field: str, value: str, label: str = ''):
    value = _sanitize_contact_field(value)
    label = _normalize_contact_label(label, field)
    if not value:
        return
    if label and label.casefold() not in value.casefold():
        combined = f'{label}: {value}'
    else:
        combined = value
    contact.setdefault(field, []).append(combined)


def _append_contact_extra(contact: dict, title: str, value: str):
    title = _normalize_contact_label(title)
    value = _sanitize_contact_field(value)
    if not title or not value:
        return
    pair = f'{title}: {value}'
    existing = {item.casefold() for item in contact.setdefault('extras', [])}
    if pair.casefold() not in existing:
        contact['extras'].append(pair)



def _humanize_contact_service(service: str) -> str:
    service = _normalize_contact_label(service)
    if not service:
        return ''
    low = service.casefold()
    mapping = {
        'facebook': 'Facebook Messenger',
        'facebook messenger': 'Facebook Messenger',
        'messenger': 'Facebook Messenger',
        'instagram': 'Instagram',
        'telegram': 'Telegram',
        'whatsapp': 'WhatsApp',
        'signal': 'Signal',
        'zalo': 'Zalo',
        'viber': 'Viber',
        'skype': 'Skype',
        'linkedin': 'LinkedIn',
        'twitter': 'Twitter',
    }
    return mapping.get(low, service)


def _load_ab_label_lookup(conn: sqlite3.Connection) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for table_name in ('ABMultiValueLabel', 'ABMultiValueLabels'):
        if not _table_exists(conn, table_name):
            continue
        cols = list(_table_columns(conn, table_name))
        label_col = next((c for c in cols if 'label' in c.lower() or c.lower() in {'name', 'value'}), '')
        if not label_col:
            continue
        try:
            rows = conn.execute(f'SELECT rowid AS _rid, * FROM "{table_name}"').fetchall()
        except Exception:
            continue
        for row in rows:
            key = _normalize_text(row['_rid'])
            value = _contact_value_from_raw(label_col, row[label_col])
            if key and value:
                lookup[key] = value
        if lookup:
            break
    return lookup


def _load_ab_entry_key_lookup(conn: sqlite3.Connection) -> dict[str, str]:
    lookup: dict[str, str] = {}
    if not _table_exists(conn, 'ABMultiValueEntryKey'):
        return lookup
    try:
        rows = conn.execute('SELECT rowid AS _rid, * FROM ABMultiValueEntryKey').fetchall()
    except Exception:
        return lookup
    label_col = 'value'
    if rows:
        cols = rows[0].keys()
        label_col = next((c for c in cols if c != '_rid'), 'value')
    for row in rows:
        key = _normalize_text(row['_rid'])
        value = _normalize_contact_label(_contact_value_from_raw(label_col, row[label_col]))
        if key and value:
            lookup[key] = value
    return lookup


def _load_ab_multivalue_entries(conn: sqlite3.Connection, entry_key_lookup: dict[str, str]) -> dict[int, dict[str, list[str]]]:
    out: dict[int, dict[str, list[str]]] = {}
    if not _table_exists(conn, 'ABMultiValueEntry'):
        return out
    try:
        rows = conn.execute('SELECT parent_id, key, value FROM ABMultiValueEntry ORDER BY parent_id, key').fetchall()
    except Exception:
        return out
    for row in rows:
        try:
            parent_id = int(float(row['parent_id']))
        except Exception:
            continue
        key_id = _normalize_text(row['key'])
        key_name = entry_key_lookup.get(key_id, key_id or 'value')
        key_name = _normalize_contact_label(key_name) or key_name
        value = _contact_value_from_raw(key_name, row['value'])
        if not value:
            continue
        bucket = out.setdefault(parent_id, {})
        bucket.setdefault(key_name, []).append(value)
    return out


def _ab_entry_first(structured: dict[str, list[str]], *keys: str) -> str:
    if not structured:
        return ''
    lowered = {str(k).casefold(): v for k, v in structured.items()}
    for key in keys:
        values = lowered.get(key.casefold())
        if values:
            for value in values:
                cleaned = _sanitize_contact_field(value)
                if cleaned:
                    return cleaned
    return ''


def _address_from_ab_structured(structured: dict[str, list[str]]) -> str:
    if not structured:
        return ''
    lines = []
    street = _ab_entry_first(structured, 'Street', 'Address', 'street', 'address')
    city = _ab_entry_first(structured, 'City', 'city')
    state = _ab_entry_first(structured, 'State', 'state', 'Province', 'province')
    country = _ab_entry_first(structured, 'Country', 'country')
    postal = _ab_entry_first(structured, 'ZIP', 'zip', 'PostalCode', 'postalCode')
    if street:
        lines.append(street)
    city_line = ', '.join(part for part in (city, state) if part)
    if city_line:
        lines.append(city_line)
    region_line = ', '.join(part for part in (country, postal) if part)
    if region_line:
        lines.append(region_line)
    if lines:
        return '\n'.join(lines)
    flat_parts = []
    for key, values in structured.items():
        for value in values:
            value = _sanitize_contact_field(value)
            if value:
                flat_parts.append(value)
    return '\n'.join(flat_parts[:6])


def _decode_contact_tone_value(value: str) -> tuple[str, str]:
    value = _sanitize_contact_field(value)
    if not value:
        return ('', '')
    lower = value.casefold()
    if lower.startswith('texttone:'):
        text = value.split(':', 1)[1].strip()
        return ('Text Tone', text)
    if lower.startswith('itunes:'):
        text = value.split(':', 1)[1].strip()
        if re.fullmatch(r'[A-Fa-f0-9]{8,}', text):
            return ('', '')
        return ('Ringtone', text)
    return ('', '')


def _guess_contact_value_kind(prop: str, label: str, value: str) -> str:
    lower = value.lower()
    label_l = label.lower()
    if prop == '3' or re.fullmatch(r'\+?\d[\d\s\-()]{5,}', value):
        return 'phones'
    if prop == '4' or '@' in lower:
        return 'emails'
    if prop == '22' or lower.startswith(('http://', 'https://')):
        return 'urls'
    if prop == '5':
        return 'addresses'
    if any(token in label_l for token in ('home', 'work', 'address', 'street', 'city', 'zip', 'postal')):
        return 'addresses'
    if any(token in lower for token in ('http://', 'https://', 'www.')):
        return 'urls'
    if re.search(r'\b(vietnam|city|district|ward|street|road|avenue|apt|apartment|house|phường|quận|thành phố|tp\.|hcm|ho chi minh)\b', lower):
        return 'addresses'
    return ''


def _set_contact_note(contact: dict, value: str):
    value = _sanitize_contact_field(value)
    if not value:
        return
    current = _sanitize_contact_field(contact.get('note', ''))
    if current:
        if value.casefold() == current.casefold():
            return
        contact['note'] = (current + '\n' + value)[:4000]
    else:
        contact['note'] = value[:4000]


def _extract_contacts_abperson(conn: sqlite3.Connection, source_db: str) -> list[dict]:
    if not _table_exists(conn, 'ABPerson'):
        return []
    contacts: dict[int, dict] = {}
    rows = conn.execute('SELECT * FROM ABPerson ORDER BY COALESCE(Last, ""), COALESCE(First, ""), ROWID').fetchall()
    for row in rows:
        rid = int(row['ROWID']) if 'ROWID' in row.keys() else int(row[0])
        first = _contact_value_from_raw('first', row['First']) if 'First' in row.keys() else _row_pick(row, 'first', 'firstname')
        middle = _row_pick(row, 'middle', 'middlename')
        last = _contact_value_from_raw('last', row['Last']) if 'Last' in row.keys() else _row_pick(row, 'last', 'lastname')
        organization = _row_pick(row, 'organization', 'company')
        display_name = ' '.join(part for part in [first, middle, last] if part).strip() or organization or f'Contact {rid}'
        contact = _contact_blank()
        contact.update({
            'record_id': rid,
            'display_name': display_name,
            'organization': organization,
            'job_title': _row_pick(row, 'jobtitle', 'title'),
            'department': _row_pick(row, 'department'),
            'birthday': _format_contact_birthday(_row_pick(row, 'birthday')),
            'source_db': source_db,
        })
        _set_contact_note(contact, _row_pick(row, 'note'))
        for col in row.keys():
            low = col.lower()
            if low in {'url', 'website', 'homepage'}:
                _append_contact_value(contact, 'urls', _contact_value_from_raw(col, row[col]))
            elif low in {'phone', 'mobile', 'mainphone'}:
                _append_contact_value(contact, 'phones', _contact_value_from_raw(col, row[col]))
            elif low in {'email', 'emailaddress'}:
                _append_contact_value(contact, 'emails', _contact_value_from_raw(col, row[col]))
            elif any(tok in low for tok in ('street', 'city', 'state', 'country', 'zip', 'postal')):
                pass
        address_parts = []
        for key in ('address', 'street', 'city', 'state', 'zip', 'postalcode', 'country'):
            val = _row_pick(row, key)
            if val:
                address_parts.append(val)
        if address_parts:
            _append_contact_value(contact, 'addresses', '\n'.join(address_parts), 'home')
        contacts[rid] = contact

    label_lookup = _load_ab_label_lookup(conn)
    entry_key_lookup = _load_ab_entry_key_lookup(conn)
    entry_lookup = _load_ab_multivalue_entries(conn, entry_key_lookup)

    if _table_exists(conn, 'ABMultiValue'):
        mv_cols = _table_columns(conn, 'ABMultiValue')
        wanted = [col for col in ('record_id', 'property', 'label', 'value', 'UID') if col in mv_cols]
        if {'record_id', 'property'} <= mv_cols and 'UID' in mv_cols:
            sql = 'SELECT ' + ', '.join(f'"{c}"' for c in wanted) + ' FROM ABMultiValue'
            for row in conn.execute(sql).fetchall():
                try:
                    rid = int(float(row['record_id']))
                except Exception:
                    continue
                if rid not in contacts:
                    continue
                prop = _normalize_text(row['property']) if 'property' in row.keys() else ''
                label_id = _normalize_text(row['label']) if 'label' in row.keys() else ''
                label = label_lookup.get(label_id, _contact_value_from_raw('label', row['label']) if 'label' in row.keys() else '')
                value = _contact_value_from_raw('value', row['value']) if 'value' in row.keys() else ''
                uid = 0
                try:
                    uid = int(float(row['UID']))
                except Exception:
                    uid = 0
                structured = entry_lookup.get(uid, {})

                if prop == '5':
                    address_text = value or _address_from_ab_structured(structured)
                    if address_text:
                        _append_contact_value(contacts[rid], 'addresses', address_text, label)
                    continue

                if prop in {'13', '46'}:
                    service = _humanize_contact_service(_ab_entry_first(structured, 'service') or label)
                    username = _ab_entry_first(structured, 'username') or value
                    if username:
                        _append_contact_extra(contacts[rid], service or ('Social' if prop == '13' else 'Messaging'), username)
                    url_value = _ab_entry_first(structured, 'url')
                    if url_value and (not username or url_value.casefold() not in username.casefold()):
                        _append_contact_extra(contacts[rid], f'{service or "Profile"} URL', url_value)
                    continue

                if prop == '16':
                    tone_title, tone_value = _decode_contact_tone_value(value)
                    if tone_title and tone_value:
                        _append_contact_extra(contacts[rid], tone_title, tone_value)
                    continue

                if not value:
                    continue

                kind = _guess_contact_value_kind(prop, label, value)
                if kind:
                    _append_contact_value(contacts[rid], kind, value, label)
                elif prop in {'17'}:
                    _set_contact_note(contacts[rid], value)
                elif prop in {'9'}:
                    contacts[rid]['birthday'] = _format_contact_birthday(value) or contacts[rid].get('birthday', '')
                elif label and not _looks_like_contact_blob_noise(value):
                    _append_contact_extra(contacts[rid], label, value)

    out = []
    for contact in contacts.values():
        for key in ('phones', 'emails', 'urls', 'addresses', 'extras'):
            contact[key] = _dedupe_preserve(contact.get(key, []))
        contact['birthday'] = _format_contact_birthday(contact.get('birthday', ''))
        contact['note'] = _sanitize_contact_field(contact.get('note', ''))
        out.append(contact)
    return out


def _find_contact_record_table(conn: sqlite3.Connection) -> str:
    best_table = ''
    best_score = -1
    for table in _sqlite_user_tables(conn):
        cols = {c.lower() for c in _table_columns(conn, table)}
        score = 0
        if 'z_pk' in cols:
            score += 10
        for col in ('zfirstname', 'zlastname', 'zfullname', 'zorganization', 'znickname', 'znote'):
            if col in cols:
                score += 8
        if 'record' in table.lower() or 'contact' in table.lower() or 'person' in table.lower():
            score += 6
        if score > best_score:
            best_score = score
            best_table = table
    return best_table if best_score >= 18 else ''


def _owner_col_from_cols(cols: set[str]) -> str:
    for needle in ('zowner', 'zrecord', 'zcontact', 'zperson', 'record_id', 'owner_id', 'contact_id'):
        for col in cols:
            if col.lower() == needle:
                return col
    return ''


def _extract_address_from_row(row: sqlite3.Row, cols: set[str]) -> tuple[str, str]:
    label = ''
    label_col = next((c for c in cols if c.lower() in {'zlabel', 'label', 'zlocalizedlabel'}), '')
    if label_col:
        label = _contact_value_from_raw(label_col, row[label_col])

    parts = []
    for col in cols:
        low = col.lower()
        if low in {'zstreet', 'street', 'zaddress', 'address', 'zsubstreet'}:
            parts.append(_contact_value_from_raw(col, row[col]))
    city_line = []
    for col in cols:
        low = col.lower()
        if low in {'zcity', 'city', 'zsubadministrativearea', 'subadministrativearea', 'zdistrict', 'district'}:
            val = _contact_value_from_raw(col, row[col])
            if val:
                city_line.append(val)
    if city_line:
        parts.append(', '.join(city_line))
    region_line = []
    for col in cols:
        low = col.lower()
        if low in {'zstate', 'state', 'zpostalcode', 'postalcode', 'zip', 'zipcode', 'country', 'zcountry'}:
            val = _contact_value_from_raw(col, row[col])
            if val:
                region_line.append(val)
    if region_line:
        parts.append(', '.join(region_line))
    text = '\n'.join(part for part in parts if part)
    return (_sanitize_contact_field(text), label)


def _extract_contacts_coredata(conn: sqlite3.Connection, source_db: str) -> list[dict]:
    record_table = _find_contact_record_table(conn)
    if not record_table:
        return []

    rows = conn.execute(f'SELECT * FROM "{record_table}" ORDER BY ROWID').fetchall()
    contacts: dict[int, dict] = {}
    for row in rows:
        rid_text = _row_pick(row, 'z_pk', 'rowid', 'pk', 'id')
        try:
            rid = int(float(rid_text))
        except Exception:
            continue
        first = _row_pick(row, 'zfirstname', 'firstname', 'first')
        middle = _row_pick(row, 'zmiddlename', 'middlename', 'middle')
        last = _row_pick(row, 'zlastname', 'lastname', 'last')
        full = _row_pick(row, 'zfullname', 'fullname', 'composite_name')
        organization = _row_pick(row, 'zorganization', 'organization', 'company')
        display_name = full or ' '.join(part for part in [first, middle, last] if part).strip() or organization or f'Contact {rid}'
        contact = _contact_blank()
        contact.update({
            'record_id': rid,
            'display_name': display_name,
            'organization': organization,
            'job_title': _row_pick(row, 'zjobtitle', 'jobtitle', 'title'),
            'department': _row_pick(row, 'zdepartmentname', 'department', 'zdepartment'),
            'birthday': _format_contact_birthday(_row_pick(row, 'zbirthday', 'birthday')),
            'source_db': source_db,
        })
        _set_contact_note(contact, _row_pick(row, 'znote', 'note'))
        for col in row.keys():
            low = col.lower()
            value = _contact_value_from_raw(col, row[col])
            if not value:
                continue
            if low in {'zurl', 'url', 'website', 'homepage'}:
                _append_contact_value(contact, 'urls', value)
            elif low in {'zemailaddress', 'email', 'emailaddress'}:
                _append_contact_value(contact, 'emails', value)
            elif low in {'zfullnumber', 'znumber', 'fullnumber', 'number', 'phone', 'telephone'}:
                _append_contact_value(contact, 'phones', value)
            elif low in {'zbirthday', 'birthday'}:
                contact['birthday'] = _format_contact_birthday(value)
            elif low in {'zsocialprofile', 'socialprofile', 'social', 'zusername', 'username'}:
                _append_contact_extra(contact, 'Social', value)
            elif any(tok in low for tok in ('ringtone', 'texttone', 'text_tone')):
                _append_contact_extra(contact, 'Ringtone' if 'ringtone' in low else 'Text Tone', value)
        contacts[rid] = contact

    if not contacts:
        return []

    for table in _sqlite_user_tables(conn):
        if table == record_table:
            continue
        cols = _table_columns(conn, table)
        owner_col = _owner_col_from_cols(cols)
        if not owner_col:
            continue

        low_cols = {c.lower() for c in cols}
        label_col = next((c for c in cols if c.lower() in {'zlabel', 'label', 'zlocalizedlabel'}), '')

        kind = ''
        value_col = ''
        if any(tok in table.lower() for tok in ('phone', 'number')) or {'zfullnumber', 'znumber', 'fullnumber', 'number', 'phone', 'telephone'} & low_cols:
            kind = 'phones'
            value_col = next((c for c in cols if c.lower() in {'zfullnumber', 'znumber', 'fullnumber', 'number', 'phone', 'telephone'}), '')
        elif 'email' in table.lower() or {'zaddress', 'zemailaddress', 'email', 'emailaddress'} & low_cols:
            kind = 'emails'
            value_col = next((c for c in cols if c.lower() in {'zemailaddress', 'email', 'emailaddress', 'zaddress'}), '')
        elif 'url' in table.lower() or {'zurl', 'url', 'website'} & low_cols:
            kind = 'urls'
            value_col = next((c for c in cols if c.lower() in {'zurl', 'url', 'website'}), '')
        elif any(tok in table.lower() for tok in ('postal', 'address', 'location')) or {'zstreet', 'street', 'zcity', 'city', 'zpostalcode', 'postalcode', 'zcountry', 'country'} & low_cols:
            kind = 'addresses'
        elif any(tok in table.lower() for tok in ('social', 'profile', 'username')):
            kind = 'extras'
            value_col = next((c for c in cols if c.lower() in {'zusername', 'username', 'zservice', 'service', 'zdisplayvalue', 'displayvalue', 'zvalue', 'value', 'zname', 'name'}), '')
        elif 'ringtone' in table.lower() or 'texttone' in table.lower() or 'tone' in table.lower():
            kind = 'extras'
            value_col = next((c for c in cols if c.lower() in {'zname', 'name', 'zvalue', 'value', 'zdisplayvalue', 'displayvalue'}), '')

        try:
            rows2 = conn.execute(f'SELECT * FROM "{table}"').fetchall()
        except Exception:
            continue

        for row in rows2:
            try:
                owner = int(float(row[owner_col]))
            except Exception:
                continue
            if owner not in contacts:
                continue
            label = _contact_value_from_raw(label_col, row[label_col]) if label_col else ''
            if kind == 'addresses':
                value, guessed_label = _extract_address_from_row(row, cols)
                if value:
                    _append_contact_value(contacts[owner], 'addresses', value, label or guessed_label)
                continue
            if not value_col:
                continue
            value = _contact_value_from_raw(value_col, row[value_col])
            if not value:
                continue
            if kind == 'extras':
                title = label or _normalize_text(table).replace('ZABCD', '').replace('Z', ' ').strip()
                _append_contact_extra(contacts[owner], title, value)
            else:
                _append_contact_value(contacts[owner], kind, value, label)

    out = []
    for contact in contacts.values():
        for key in ('phones', 'emails', 'urls', 'addresses', 'extras'):
            contact[key] = _dedupe_preserve(contact.get(key, []))
        contact['birthday'] = _format_contact_birthday(contact.get('birthday', ''))
        contact['note'] = _sanitize_contact_field(contact.get('note', ''))
        out.append(contact)
    return out


def _extract_contact_avatars(root_dir: Path, db_candidates: list[Path], contacts: list[dict], out_dir: Path):
    if not contacts:
        return
    out_dir.mkdir(parents=True, exist_ok=True)
    by_id = {}
    for contact in contacts:
        rid = contact.get('record_id')
        if rid is None:
            continue
        try:
            by_id[int(rid)] = contact
        except Exception:
            pass
    if not by_id:
        return

    image_dbs = [p for p in db_candidates if 'image' in p.name.lower() or 'photo' in p.name.lower() or 'addressbook' in p.name.lower()]
    for db_path in image_dbs:
        try:
            conn = sqlite3.connect(str(db_path))
            conn.row_factory = sqlite3.Row
        except Exception:
            continue
        try:
            for table in _sqlite_user_tables(conn):
                cols = _table_columns(conn, table)
                owner_col = _owner_col_from_cols(cols)
                if not owner_col:
                    continue
                blob_cols = [c for c in cols if any(tok in c.lower() for tok in ('data', 'image', 'photo', 'thumbnail', 'jpeg', 'png'))]
                if not blob_cols:
                    continue
                try:
                    rows = conn.execute(f'SELECT * FROM "{table}" LIMIT 500').fetchall()
                except Exception:
                    continue
                for row in rows:
                    try:
                        owner = int(float(row[owner_col]))
                    except Exception:
                        continue
                    contact = by_id.get(owner)
                    if not contact or contact.get('avatar_rel_path'):
                        continue
                    for blob_col in blob_cols:
                        raw = row[blob_col]
                        if isinstance(raw, memoryview):
                            raw = bytes(raw)
                        if not isinstance(raw, (bytes, bytearray)):
                            continue
                        raw = bytes(raw)
                        if len(raw) < 64 or not _contact_is_image_bytes(raw):
                            continue
                        ext = _contact_image_extension(raw)
                        img_name = f'contact_{owner}{ext}'
                        img_path = out_dir / img_name
                        try:
                            img_path.write_bytes(raw)
                        except Exception:
                            continue
                        contact['avatar_rel_path'] = f'avatars/{img_name}'
                        break
        finally:
            conn.close()


def _extract_contacts_from_sqlite(sqlite_path: Path) -> list[dict]:
    try:
        conn = sqlite3.connect(str(sqlite_path))
        conn.row_factory = sqlite3.Row
    except Exception:
        return []
    try:
        contacts = _extract_contacts_abperson(conn, sqlite_path.name)
        if contacts:
            return contacts
        return _extract_contacts_coredata(conn, sqlite_path.name)
    finally:
        conn.close()


def export_contacts_readable(root_dir: str | Path, out_dir: str | Path) -> int:
    root_dir = Path(root_dir)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    db_dir = out_dir / 'db_tables'
    db_dir.mkdir(parents=True, exist_ok=True)
    avatars_dir = out_dir / 'avatars'
    avatars_dir.mkdir(parents=True, exist_ok=True)

    summary_txt_path = out_dir / 'contacts_summary.txt'
    csv_path = out_dir / 'contacts.csv'
    txt_path = out_dir / 'contacts.txt'
    index_html_path = out_dir / 'contacts.html'

    all_files = sorted((p for p in root_dir.rglob('*') if p.is_file()), key=lambda p: str(p).lower())
    db_candidates = [p for p in all_files if p.suffix.lower() in CONTACT_DB_EXTENSIONS]
    primary_db = None
    for path in db_candidates:
        name = path.name.lower()
        if 'addressbookimages' in name:
            continue
        if 'addressbook' in name or 'contact' in name:
            primary_db = path
            break
    if primary_db is None and db_candidates:
        primary_db = db_candidates[0]

    contacts: list[dict] = []
    if primary_db is not None:
        contacts = _extract_contacts_from_sqlite(primary_db)
        _extract_contact_avatars(root_dir, db_candidates, contacts, avatars_dir)

    table_cards: list[dict] = []
    summary_lines: list[str] = []
    written = 0

    contacts_sorted = sorted(contacts, key=lambda c: (c.get('display_name') or '').casefold())
    if contacts_sorted:
        with csv_path.open('w', encoding='utf-8-sig', newline='') as f_csv, txt_path.open('w', encoding='utf-8') as f_txt:
            writer = csv.writer(f_csv)
            writer.writerow(['display_name', 'organization', 'phones', 'emails', 'urls', 'addresses', 'extras', 'job_title', 'department', 'birthday', 'note', 'avatar_rel_path', 'source_db'])
            for contact in contacts_sorted:
                writer.writerow([
                    contact.get('display_name', ''),
                    contact.get('organization', ''),
                    ' | '.join(contact.get('phones', [])),
                    ' | '.join(contact.get('emails', [])),
                    ' | '.join(contact.get('urls', [])),
                    ' | '.join(contact.get('addresses', [])),
                    ' | '.join(contact.get('extras', [])),
                    contact.get('job_title', ''),
                    contact.get('department', ''),
                    contact.get('birthday', ''),
                    contact.get('note', ''),
                    contact.get('avatar_rel_path', ''),
                    contact.get('source_db', ''),
                ])
                f_txt.write((contact.get('display_name') or 'Unnamed Contact') + '\n')
                for label, key in (
                    ('Organization', 'organization'),
                    ('Phone', 'phones'),
                    ('Email', 'emails'),
                    ('URL', 'urls'),
                    ('Home / Address', 'addresses'),
                    ('Job Title', 'job_title'),
                    ('Department', 'department'),
                    ('Birthday', 'birthday'),
                    ('Note', 'note'),
                    ('Avatar', 'avatar_rel_path'),
                    ('Source DB', 'source_db'),
                ):
                    value = contact.get(key, '')
                    if isinstance(value, list):
                        value = '; '.join(value)
                    value = _normalize_text(value)
                    if value:
                        f_txt.write(f'  {label}: {value}\n')
                f_txt.write('\n')
        written += 2
        summary_lines.append(f'Contacts parsed: {len(contacts_sorted)}')
        summary_lines.append('Fields exported: name, avatar, phone, email, url, address/home, extra labeled fields, birthday, note')
    else:
        summary_lines.append('Contacts parsed: 0')

    for db_path in db_candidates[:6]:
        rel_db = str(db_path.relative_to(root_dir)).replace('\\', '/')
        try:
            conn = sqlite3.connect(str(db_path))
            conn.row_factory = sqlite3.Row
        except Exception as exc:
            summary_lines.append(f'DB skipped: {rel_db} ({exc})')
            continue
        try:
            tables = _sqlite_user_tables(conn)
            for table in tables[:12]:
                try:
                    count = conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
                except Exception:
                    count = 0
                try:
                    rows = conn.execute(f'SELECT * FROM "{table}" LIMIT 50').fetchall()
                except Exception as exc:
                    summary_lines.append(f'Table skipped: {rel_db} / {table} ({exc})')
                    continue
                preview_rows: list[dict] = []
                for row in rows:
                    item = {}
                    for key in row.keys():
                        item[key] = _db_value_to_text(key, row[key])
                    preview_rows.append(item)
                safe_name = re.sub(r'[^a-zA-Z0-9._-]+', '_', f'{db_path.stem}_{table}')[:120] + '.html'
                page_path = db_dir / safe_name
                page_path.write_text(
                    _render_contacts_table_preview_page(
                        title=f'{db_path.name} · {table}',
                        subtitle=f'{count} row(s) in {rel_db}',
                        rows=preview_rows,
                        back_href='../contacts.html',
                    ),
                    encoding='utf-8',
                )
                csv_table_path = db_dir / (Path(safe_name).stem + '.csv')
                if preview_rows:
                    keys = list(preview_rows[0].keys())
                    with csv_table_path.open('w', encoding='utf-8-sig', newline='') as f_csv:
                        writer = csv.DictWriter(f_csv, fieldnames=keys)
                        writer.writeheader()
                        writer.writerows(preview_rows)
                table_cards.append(
                    {
                        'href': f'db_tables/{page_path.name}',
                        'label': f'{db_path.name} · {table}',
                        'meta': f'{count} row(s)',
                        'preview': ' | '.join(
                            f'{k}={v}'
                            for k, v in (preview_rows[0].items() if preview_rows else [])
                        )[:280] or 'Open table preview',
                    }
                )
        finally:
            conn.close()

    note = ''
    if contacts_sorted:
        note = 'Long random text is not a real note from the phone. That usually means the database row contained binary / structured payload, and the exporter now skips that noise instead of showing it as Notes.'
    index_html_path.write_text(
        _render_contacts_index_page(
            title='Contacts',
            subtitle=f'{len(contacts_sorted)} contact(s) parsed from the backup',
            contacts=contacts_sorted,
            table_cards=table_cards,
            note=note,
        ),
        encoding='utf-8',
    )
    written += 1
    summary_txt_path.write_text('\n'.join(summary_lines) + '\n', encoding='utf-8')
    written += 1
    return written

def export_voicemail_readable(root_dir: str | Path, out_dir: str | Path) -> int:
    root_dir = Path(root_dir)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    db_dir = out_dir / "db_tables"
    db_dir.mkdir(parents=True, exist_ok=True)

    audio_entries: list[dict] = []
    table_cards: list[dict] = []
    summary_lines: list[str] = []
    written = 0

    audio_csv_path = out_dir / "voicemail_audio.csv"
    audio_txt_path = out_dir / "voicemail_audio.txt"
    summary_txt_path = out_dir / "voicemail_summary.txt"
    index_html_path = out_dir / "voicemail.html"

    all_files = sorted((p for p in root_dir.rglob('*') if p.is_file()), key=lambda p: str(p).lower())
    db_candidates: list[Path] = []

    for path in all_files:
        suffix = path.suffix.lower()
        if suffix in VOICEMAIL_AUDIO_EXTENSIONS:
            rel = path.relative_to(root_dir)
            stat = path.stat()
            modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            audio_entries.append(
                {
                    "name": path.name,
                    "relative_path": str(rel).replace("\\", "/"),
                    "extension": suffix.lstrip(".").upper(),
                    "size_bytes": stat.st_size,
                    "size_text": _human_file_size(stat.st_size),
                    "modified": modified,
                    "uri": path.resolve().as_uri(),
                }
            )
        elif suffix in {".db", ".sqlite", ".sqlite3", ".storedata"}:
            db_candidates.append(path)

    if audio_entries:
        with audio_csv_path.open("w", encoding="utf-8-sig", newline="") as f_csv, audio_txt_path.open("w", encoding="utf-8") as f_txt:
            writer = csv.writer(f_csv)
            writer.writerow(["file_name", "relative_path", "extension", "size_bytes", "modified_time", "audio_uri"])
            for entry in audio_entries:
                writer.writerow([
                    entry["name"],
                    entry["relative_path"],
                    entry["extension"],
                    entry["size_bytes"],
                    entry["modified"],
                    entry["uri"],
                ])
                f_txt.write(f'[{entry["modified"]}] {entry["name"]} | {entry["size_text"]}\n{entry["relative_path"]}\n\n')
        written += 2
        summary_lines.append(f"Audio files: {len(audio_entries)}")
    else:
        summary_lines.append("Audio files: 0")

    for db_path in db_candidates:
        rel_db = str(db_path.relative_to(root_dir)).replace("\\", "/")
        try:
            conn = sqlite3.connect(str(db_path))
            conn.row_factory = sqlite3.Row
        except Exception as exc:
            summary_lines.append(f"DB skipped: {rel_db} ({exc})")
            continue
        try:
            tables = _sqlite_user_tables(conn)
            if not tables:
                summary_lines.append(f"DB: {rel_db} | no user tables")
                continue
            for table in tables:
                try:
                    count = conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
                except Exception:
                    count = 0
                try:
                    rows = conn.execute(f'SELECT * FROM "{table}" LIMIT 250').fetchall()
                except Exception as exc:
                    summary_lines.append(f"DB table skipped: {rel_db} :: {table} ({exc})")
                    continue

                columns = list(rows[0].keys()) if rows else sorted(_table_columns(conn, table))
                normalized_rows: list[dict] = []
                for row in rows:
                    if hasattr(row, 'keys'):
                        row_dict = {col: _db_value_to_text(col, row[col]) for col in row.keys()}
                    else:
                        row_dict = {
                            columns[i] if i < len(columns) else f'col_{i+1}': _db_value_to_text(columns[i] if i < len(columns) else f'col_{i+1}', value)
                            for i, value in enumerate(row)
                        }
                    normalized_rows.append(row_dict)

                slug = _safe_filename(f'{db_path.stem}_{table}', default='voicemail_table')
                csv_path = db_dir / f'{slug}.csv'
                html_path = db_dir / f'{slug}.html'

                with csv_path.open('w', encoding='utf-8-sig', newline='') as f_csv:
                    writer = csv.writer(f_csv)
                    writer.writerow(columns)
                    for row_dict in normalized_rows:
                        writer.writerow([row_dict.get(col, '') for col in columns])
                written += 1

                html_doc = _render_db_table_preview_page(
                    f'{db_path.name} · {table}',
                    f'{count} row(s) in table {table}. Showing up to 250 rows in a readable layout.',
                    normalized_rows,
                    '../voicemail.html',
                )
                html_path.write_text(html_doc, encoding='utf-8')
                written += 1

                preview_bits = []
                if normalized_rows:
                    first_row = normalized_rows[0]
                    for col in columns[:3]:
                        val = first_row.get(col, '')
                        if val:
                            preview_bits.append(f'{col}: {val}')
                    preview = " | ".join(preview_bits) or "Open for table preview"
                else:
                    preview = "No rows in this table"
                table_cards.append(
                    {
                        "label": f'{db_path.name} · {table}',
                        "meta": f'{count} row(s) · {rel_db}',
                        "preview": preview,
                        "href": f'db_tables/{slug}.html',
                    }
                )
                summary_lines.append(f"DB: {rel_db} :: {table} | rows={count}")
        finally:
            conn.close()

    note = "Voicemail export now includes a readable HTML index, audio file list, and preview pages for any SQLite voicemail tables that were found."
    index_html_path.write_text(
        _render_voicemail_index_page(
            "Voicemail",
            "Readable export for voicemail audio and voicemail-related databases. You do not need to browse the raw .db file first.",
            audio_entries,
            table_cards,
            note=note,
        ),
        encoding='utf-8',
    )
    written += 1

    summary_txt_path.write_text("\n".join(summary_lines) + "\n", encoding='utf-8')
    written += 1
    return written



class ValidationWorker(QObject):
    finished = Signal(bool, str)

    def __init__(self, folder: str, password: str):
        super().__init__()
        self.folder = folder
        self.password = password

    def run(self):
        try:
            backup = try_open_backup(self.folder, self.password)
            self.finished.emit(True, f"Opened backup: {backup.path}")
        except Exception as exc:
            self.finished.emit(False, str(exc))


class DecryptWorker(QObject):
    log = Signal(str, str)
    status = Signal(str, str)
    count = Signal(int)
    progress_value = Signal(int)
    progress_busy = Signal(bool)
    finished = Signal(bool, int)

    def __init__(self, folder: str, password: str, output: str, opts: dict[str, bool]):
        super().__init__()
        self.folder = folder
        self.password = password
        self.output = output
        self.opts = opts
        self.done_files = 0
        self.done_steps = 0
        self.total_steps = max(sum(1 for v in opts.values() if v), 1)

    def _tick(self):
        self.done_steps += 1
        pct = int(self.done_steps / self.total_steps * 100)
        self.progress_value.emit(pct)

    def _emit_count(self, amount: int):
        self.done_files += amount
        self.count.emit(self.done_files)

    def _log_result(self, label: str, amount: int):
        if amount > 0:
            self.log.emit(f"{label} — {amount} file(s)", "ok")
        else:
            self.log.emit(f"{label} — no matching files found", "warn")

    def run(self):
        try:
            backup = try_open_backup(self.folder, self.password)
        except Exception as exc:
            self.log.emit(f"Failed to open backup: {exc}", "err")
            self.finished.emit(False, self.done_files)
            return

        all_ok = True

        def exact_one(label: str, candidates: list[tuple[str, str]], out_file: str, postprocess: Callable[[str], int] | None = None):
            nonlocal all_ok
            self.status.emit(f"Extracting {label}…", "normal")
            self.log.emit(f"Extracting {label}…", "dim")
            extracted = 0
            last_error = None
            for domain, rel in candidates:
                try:
                    entry = backup.get_entry_by_domain_and_path(domain, rel)
                    Path(out_file).parent.mkdir(parents=True, exist_ok=True)
                    Path(out_file).write_bytes(entry.read_bytes())
                    extracted = 1
                    break
                except Exception as exc:
                    last_error = exc
            if extracted:
                self._emit_count(extracted)
                self._log_result(label, extracted)
                if postprocess is not None:
                    try:
                        extra_written = int(postprocess(out_file) or 0)
                        if extra_written > 0:
                            self._emit_count(extra_written)
                            self.log.emit(f"{label} readable export — {extra_written} file(s)", "ok")
                    except Exception as exc:
                        self.log.emit(f"{label} readable export — {exc}", "warn")
            else:
                all_ok = False
                self.log.emit(f"{label} — {last_error or 'not found'}", "err")
            self._tick()

        def bulk(label: str, out_dir: str, matcher: Callable, preserve_domain: bool = True, postprocess: Callable[[str], int] | None = None, postprocess_even_if_empty: bool = False):
            nonlocal all_ok
            self.status.emit(f"Extracting {label}…", "normal")
            self.progress_busy.emit(True)
            self.log.emit(f"Extracting {label}…", "dim")
            extracted = 0
            try:
                for entry in backup.iter_files():
                    if matcher(entry):
                        write_entry(entry, out_dir, preserve_domain=preserve_domain)
                        extracted += 1
                self._emit_count(extracted)
                self._log_result(label, extracted)
                if postprocess is not None and (extracted or postprocess_even_if_empty):
                    try:
                        extra_written = int(postprocess(out_dir) or 0)
                        if extra_written > 0:
                            self._emit_count(extra_written)
                            self.log.emit(f"{label} readable export — {extra_written} file(s)", "ok")
                    except Exception as exc:
                        self.log.emit(f"{label} readable export — {exc}", "warn")
            except Exception as exc:
                self.log.emit(f"{label} — {exc}", "err")
                all_ok = False
            finally:
                self.progress_busy.emit(False)
            self._tick()

        if self.opts.get("call"):
            exact_one(
                "Call History",
                CALL_HISTORY_CANDIDATES,
                os.path.join(self.output, "call_history.sqlite"),
                postprocess=lambda p: export_call_history_readable(p, os.path.join(self.output, "call_history_readable")),
            )
        if self.opts.get("sms"):
            exact_one(
                "SMS & iMessage",
                SMS_CANDIDATES,
                os.path.join(self.output, "sms.sqlite"),
                postprocess=lambda p: export_sms_readable(p, os.path.join(self.output, "sms_readable")),
            )
        if self.opts.get("photos"):
            bulk("Photos", os.path.join(self.output, "photos"), classify_photo, preserve_domain=False)
        if self.opts.get("contacts"):
            bulk(
                "Contacts",
                os.path.join(self.output, "contacts"),
                classify_contacts,
                preserve_domain=True,
                postprocess=lambda p: export_contacts_readable(p, os.path.join(self.output, "contacts_readable")),
            )
        if self.opts.get("voicemail"):
            bulk(
                "Voicemail",
                os.path.join(self.output, "voicemail"),
                classify_voicemail,
                preserve_domain=True,
                postprocess=lambda p: export_voicemail_readable(p, os.path.join(self.output, "voicemail_readable")),
            )
        self.finished.emit(all_ok, self.done_files)


class App(QMainWindow):
    CATEGORIES = [
        ("call", "Call History", "📞"),
        ("sms", "SMS & iMessage", "💬"),
        ("photos", "Photos", "🖼"),
        ("contacts", "Contacts", "👤"),
        ("voicemail", "Voicemail", "🎙"),
    ]

    def __init__(self):
        super().__init__()
        self.setWindowTitle("iPhone backup Decryptor by Andy Le 0868231181")
        self.setWindowIcon(app_icon())
        self.setMinimumSize(880, 660)
        self.resize(880, 660)

        self.backup_unlocked = False
        self._real_folder = ""
        self._real_output = ""

        self._validation_thread: QThread | None = None
        self._validation_worker: ValidationWorker | None = None
        self._decrypt_thread: QThread | None = None
        self._decrypt_worker: DecryptWorker | None = None

        self._build_ui()
        self._build_menu()
        self.output_input.setText(str(default_output_dir()))
        self._real_output = str(default_output_dir())
        self._lock_categories()
        self._autofill_backup_path()

    def _build_menu(self):
        menu = self.menuBar()
        menu.setNativeMenuBar(False)
        menu.setStyleSheet(
            f"""
            QMenuBar {{
                background: {GLASS_STR};
                color: {TEXT};
                border-bottom: 1px solid {BORDER};
                padding: 4px 8px;
                spacing: 8px;
            }}
            QMenuBar::item {{
                background: transparent;
                padding: 8px 12px;
                border-radius: 8px;
            }}
            QMenuBar::item:selected {{
                background: {GLASS_HOV};
            }}
            QMenu {{
                background: {GLASS_STR};
                color: {TEXT};
                border: 1px solid {BORDER};
                padding: 6px;
            }}
            QMenu::item {{
                padding: 8px 18px;
                border-radius: 8px;
            }}
            QMenu::item:selected {{
                background: {GLASS_HOV};
            }}
            """
        )

        help_menu = menu.addMenu("&Help")

        guide_action = QAction("Apple Devices Backup Guide", self)
        guide_action.triggered.connect(self._show_help_guide)
        help_menu.addAction(guide_action)

        #store_action = QAction("Open Apple Devices in Microsoft Store", self)
        #store_action.triggered.connect(self._open_apple_devices_store)
        #help_menu.addAction(store_action)

        #help_menu.addSeparator()

        #about_action = QAction("About This App", self)
        #about_action.triggered.connect(self._show_about_dialog)
        #help_menu.addAction(about_action)

    def _build_ui(self):
        root = QWidget()
        root.setStyleSheet(
            f"""
            QWidget {{
                background: {BG};
                color: {TEXT};
                font-family: Segoe UI, Helvetica Neue, Arial;
                font-size: 13px;
            }}
            QLabel#title {{
                font-size: 28px;
                font-weight: 700;
                color: {TEXT};
            }}
            QLabel#subtitle {{
                color: {TEXT_SEC};
                font-size: 12px;
            }}
            QPushButton {{
                background: {GLASS_STR};
                color: {TEXT_SEC};
                border: 1px solid {BORDER};
                border-radius: 12px;
                padding: 10px 16px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background: {GLASS_HOV};
            }}
            QPushButton:disabled {{
                color: {TEXT_DIM};
                border-color: {BORDER};
            }}
            QPushButton#accentBtn {{
                background: {ACCENT};
                color: white;
                border: none;
            }}
            QPushButton#accentBtn:hover {{
                background: #5a52e0;
            }}
            QPushButton#accentBtn:disabled {{
                background: #403d73;
                color: #d8d7f0;
            }}
            QLineEdit {{
                background: #111120;
                color: {TEXT};
                border: 1px solid {BORDER};
                border-radius: 10px;
                padding: 10px 12px;
                selection-background-color: {ACCENT};
            }}
            QLineEdit:focus {{
                border: 1px solid {ACCENT};
            }}
            QCheckBox {{
                color: {TEXT_SEC};
            }}
            QProgressBar {{
                border: none;
                border-radius: 5px;
                background: {BORDER};
                text-align: center;
                color: {TEXT};
                min-height: 10px;
                max-height: 10px;
            }}
            QProgressBar::chunk {{
                background: {ACCENT};
                border-radius: 5px;
            }}
            QTextEdit {{
                background: {GLASS_STR};
                border: 1px solid {BORDER};
                border-radius: 16px;
                color: {TEXT_SEC};
                padding: 8px;
            }}
            QScrollArea {{
                border: none;
                background: {BG};
            }}
            QScrollBar:vertical, QScrollBar:horizontal {{
                width: 0px;
                height: 0px;
                background: transparent;
            }}
            """
        )
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        outer.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)
        main = QVBoxLayout(content)
        main.setContentsMargins(32, 28, 32, 28)
        main.setSpacing(16)

        hdr = QHBoxLayout()
        hdr.setSpacing(12)
        dot = QFrame()
        dot.setFixedSize(6, 6)
        dot.setStyleSheet(f"background: {ACCENT}; border-radius: 3px;")
        hdr.addWidget(dot, 0, Qt.AlignmentFlag.AlignTop)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        title = QLabel("iPhone backup Decryptor")
        title.setObjectName("title")
        subtitle = QLabel("Auto-find local iPhone backup, unlock it, and extract selected data to readable files")
        subtitle.setObjectName("subtitle")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        hdr.addLayout(title_col)
        hdr.addStretch(1)
        main.addLayout(hdr)

        body = QHBoxLayout()
        body.setSpacing(20)
        main.addLayout(body)

        left = QVBoxLayout()
        left.setSpacing(10)
        right = QVBoxLayout()
        right.setSpacing(10)
        body.addLayout(left, 1)
        body.addLayout(right, 1)

        # Backup location
        left.addWidget(section_label("BACKUP LOCATION"))
        card = card_widget()
        left.addWidget(card)
        c = QVBoxLayout(card)
        c.setContentsMargins(18, 18, 18, 18)
        c.setSpacing(12)

        top = QHBoxLayout()
        folder_lbl = QLabel("Folder")
        folder_lbl.setStyleSheet(f"color:{TEXT}; font-size:13px;")
        self.auto_find_btn = QPushButton("Auto Find")
        self.auto_find_btn.clicked.connect(self._autofill_backup_path)
        self.folder_btn = QPushButton("Browse")
        self.folder_btn.clicked.connect(self._choose_folder)
        top.addWidget(folder_lbl)
        top.addStretch(1)
        top.addWidget(self.auto_find_btn)
        top.addWidget(self.folder_btn)
        c.addLayout(top)

        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("No folder selected")
        c.addWidget(self.folder_input)
        c.addWidget(divider())

        pw_row = QHBoxLayout()
        pw_row.addWidget(QLabel("Password"))
        pw_row.addStretch(1)
        self.show_pw = QCheckBox("Show")
        self.show_pw.toggled.connect(self._toggle_pw)
        pw_row.addWidget(self.show_pw)
        c.addLayout(pw_row)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        c.addWidget(self.password_input)
        c.addWidget(divider())

        hint = QLabel(
            "Auto Find scans standard Apple backup folders. You can still browse manually. "
            "If you choose a parent folder, the app will try to resolve the latest valid backup inside it."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color:{TEXT_DIM}; font-size:12px;")
        c.addWidget(hint)

        self.unlock_row = QHBoxLayout()
        self.unlock_btn = QPushButton("Unlock Backup")
        self.unlock_btn.setObjectName("accentBtn")
        self.unlock_btn.clicked.connect(self._unlock_backup)
        self.folder_status = QLabel("")
        self.folder_status.setWordWrap(True)
        self.folder_status.setStyleSheet(f"color:{WARN}; font-size:12px;")
        self.unlock_row.addWidget(self.unlock_btn)
        self.unlock_row.addWidget(self.folder_status, 1)
        c.addLayout(self.unlock_row)

        # Output folder
        left.addWidget(section_label("OUTPUT FOLDER"))
        out_card = card_widget()
        left.addWidget(out_card)
        oc = QVBoxLayout(out_card)
        oc.setContentsMargins(18, 18, 18, 18)
        oc.setSpacing(12)

        out_top = QHBoxLayout()
        out_top.addWidget(QLabel("Save to"))
        out_top.addStretch(1)
        self.output_btn = QPushButton("Browse")
        self.output_btn.clicked.connect(self._choose_output)
        out_top.addWidget(self.output_btn)
        oc.addLayout(out_top)

        self.output_input = QLineEdit()
        self.output_input.setPlaceholderText("No folder selected")
        oc.addWidget(self.output_input)

        # Progress
        left.addWidget(section_label("PROGRESS"))
        prog_card = card_widget()
        left.addWidget(prog_card)
        pc = QVBoxLayout(prog_card)
        pc.setContentsMargins(18, 18, 18, 18)
        pc.setSpacing(8)

        self.status_label = QLabel("Waiting for backup…")
        self.status_label.setStyleSheet(f"color:{TEXT_SEC}; font-size:12px;")
        pc.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        pc.addWidget(self.progress)

        self.count_label = QLabel("")
        self.count_label.setStyleSheet(f"color:{TEXT_DIM}; font-size:12px;")
        pc.addWidget(self.count_label)

        left.addStretch(1)

        # Extract panel
        right.addWidget(section_label("WHAT TO EXTRACT"))
        ext_card = card_widget()
        right.addWidget(ext_card)
        ec = QVBoxLayout(ext_card)
        ec.setContentsMargins(16, 16, 16, 16)
        ec.setSpacing(4)

        self.category_rows: dict[str, RowToggle] = {}
        for i, (key, label, icon) in enumerate(self.CATEGORIES):
            row = RowToggle(key, label, icon)
            self.category_rows[key] = row
            ec.addWidget(row)
            if i < len(self.CATEGORIES) - 1:
                ec.addWidget(divider())

        ec.addWidget(divider())
        self.select_all_row = RowToggle("all", "Select All")
        self.select_all_row.toggled.connect(self._select_all_changed)
        ec.addWidget(self.select_all_row)

        self.extract_hint = QLabel("Unlock your backup first to select files.")
        self.extract_hint.setStyleSheet(f"color:{TEXT_DIM}; font-size:12px;")
        right.addWidget(self.extract_hint)

        self.extract_note = QLabel(
            "SMS, Call History, Contacts, and Voicemail will also generate readable HTML/TXT/CSV reports so you do not have to inspect raw SQLite / plist files directly."
        )
        self.extract_note.setWordWrap(True)
        self.extract_note.setStyleSheet(f"color:{TEXT_DIM}; font-size:12px;")
        right.addWidget(self.extract_note)

        self.run_btn = QPushButton("Extract Selected")
        self.run_btn.setObjectName("accentBtn")
        self.run_btn.clicked.connect(self._run_extract)
        self.run_btn.setEnabled(False)
        right.addWidget(self.run_btn)

        right.addWidget(section_label("LOG"))
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(250)
        right.addWidget(self.log, 1)

    def _message_box_style(self) -> str:
        return f"""
        QMessageBox {{
            background: {BG};
        }}
        QMessageBox QLabel {{
            color: {TEXT};
            min-width: 360px;
            font-size: 13px;
        }}
        QMessageBox QPushButton {{
            background: {GLASS_STR};
            color: {TEXT};
            border: 1px solid {BORDER};
            border-radius: 10px;
            padding: 10px 18px;
            min-width: 120px;
            font-weight: 600;
        }}
        QMessageBox QPushButton:hover {{
            background: {GLASS_HOV};
        }}
        QMessageBox QPushButton#openFolderBtn {{
            background: {ACCENT};
            border: none;
            color: white;
        }}
        QMessageBox QPushButton#openFolderBtn:hover {{
            background: #5a52e0;
        }}
        """

    def _show_message(self, title: str, text: str, icon=QMessageBox.Icon.Information, allow_open_folder: bool = False) -> None:
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setWindowIcon(app_icon())
        box.setIcon(QMessageBox.Icon.NoIcon if allow_open_folder else icon)
        box.setText(text)
        box.setStyleSheet(self._message_box_style())

        open_btn = None
        if allow_open_folder and self._real_output:
            open_btn = box.addButton("Open Folder", QMessageBox.ButtonRole.ActionRole)
            open_btn.setObjectName("openFolderBtn")
        ok_btn = box.addButton(QMessageBox.StandardButton.Ok)
        box.setDefaultButton(ok_btn)
        box.exec()

        if open_btn is not None and box.clickedButton() == open_btn and self._real_output:
            self._open_output_folder()

    def _open_apple_devices_store(self):
        QDesktopServices.openUrl(
            QUrl("https://apps.microsoft.com/detail/9NP83LWLPZ9K?hl=en-us&gl=US&ocid=pdpshare")
        )

    def _show_help_guide(self):
        steps_html = f"""
        <b>How to create and read an iPhone backup</b><br><br>
        <ol style="margin-left:18px;">
            <li>Open <b>Microsoft Store</b>, search for <b>Apple Devices</b>, and install it. You can also open it directly from the Help menu.</li>
            <li>Connect your iPhone to the PC with a USB cable. In the <b>General</b> tab, choose <b>Back up all of the data on your iPhone to this computer</b>. Turn on <b>Encrypt local backup</b> if you want to protect the backup with a password.</li>
            <li>Click <b>Back Up Now</b> and wait until Apple Devices finishes creating the backup.</li>
            <li>Open Andy's app and click <b>Auto Find</b>. The app will locate the newest backup automatically. If the backup is encrypted, enter the same password you used in Apple Devices, then click <b>Unlock Backup</b>.</li>
            <li>Choose your <b>Output Folder</b>. If you do not change it, the extracted files will be saved to your <b>Downloads</b> folder by default.</li>
            <li>In <b>What to extract</b>, select the categories you want to inspect, or choose <b>Select All</b>.</li>
            <li>Click <b>Extract Selected</b> to start the final extraction and decryption process.</li>
        </ol>
        <span style="color:{TEXT_DIM};">Tip: If Auto Find cannot locate a backup, create one first in Apple Devices, then return to this app and try again.</span>
        """

        box = QMessageBox(self)
        box.setWindowTitle("Help — Apple Devices Backup Guide")
        box.setWindowIcon(app_icon())
        box.setIcon(QMessageBox.Icon.NoIcon)
        help_image = resource_file_path("screen.png")
        if help_image is not None:
            pixmap = QPixmap(str(help_image))
            if not pixmap.isNull():
                box.setIconPixmap(
                    pixmap.scaled(
                        360,
                        720,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )
                )
        box.setTextFormat(Qt.TextFormat.RichText)
        box.setText(steps_html)
        box.setStyleSheet(self._message_box_style())

        store_btn = box.addButton("Open Microsoft Store", QMessageBox.ButtonRole.ActionRole)
        ok_btn = box.addButton(QMessageBox.StandardButton.Ok)
        box.setDefaultButton(ok_btn)
        box.exec()

        if box.clickedButton() == store_btn:
            self._open_apple_devices_store()

    def _show_about_dialog(self):
        about_text = (
            "iPhone backup Decryptor by Andy Le 0868231181\n\n"
            "Use Apple Devices to create a local iPhone backup, then unlock it here and extract readable files such as SMS, call history, contacts, photos, and voicemail."
        )
        self._show_message("About This App", about_text)

    def _open_output_folder(self):
        if not self._real_output:
            return
        folder = Path(self._real_output)
        if not folder.exists():
            folder.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder.resolve())))

    def _autofill_backup_path(self):
        found = auto_find_latest_backup()
        if found:
            self.folder_input.setText(str(found))
            self.folder_status.setStyleSheet(f"color:{SUCCESS}; font-size:12px;")
            self.folder_status.setText(f"Auto-found latest backup: {found.name}")
            self._log_write(f"Auto-found backup: {found}", "ok")
        else:
            self.folder_status.setStyleSheet(f"color:{WARN}; font-size:12px;")
            self.folder_status.setText("No local iPhone backup was found in standard Apple backup folders.")
            self._log_write("No auto-found backup. Browse manually or create a backup in Apple Devices/iTunes first.", "warn")

    def _choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose Backup Folder")
        if folder:
            self.folder_input.setText(folder)
            resolved = resolve_backup_folder(folder)
            if resolved and str(resolved) != folder:
                self.folder_status.setStyleSheet(f"color:{SUCCESS}; font-size:12px;")
                self.folder_status.setText(f"Resolved to: {resolved}")
                self._log_write(f"Resolved selected folder to backup: {resolved}", "ok")
            elif resolved:
                self.folder_status.setStyleSheet(f"color:{SUCCESS}; font-size:12px;")
                self.folder_status.setText("Valid backup folder selected.")
            else:
                self.folder_status.setStyleSheet(f"color:{WARN}; font-size:12px;")
                self.folder_status.setText("That folder is not a valid backup yet.")

    def _choose_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose Output Folder", str(default_output_dir()))
        if folder:
            self.output_input.setText(folder)
            self._real_output = folder

    def _toggle_pw(self, checked: bool):
        self.password_input.setEchoMode(
            QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
        )

    def _lock_categories(self):
        for row in self.category_rows.values():
            row.set_enabled(False)
        self.select_all_row.set_enabled(False)
        self.extract_hint.setText("Unlock your backup first to select files.")
        self.run_btn.setEnabled(False)

    def _unlock_categories(self):
        for row in self.category_rows.values():
            row.set_enabled(True)
        self.select_all_row.set_enabled(True)
        self.extract_hint.setText("Select the data you want to extract.")
        self.run_btn.setEnabled(True)

    def _select_all_changed(self, checked: bool):
        if not self.backup_unlocked:
            return
        for row in self.category_rows.values():
            row.set_checked(checked)

    def _log_write(self, msg: str, tag: str = ""):
        colors = {
            "ok": SUCCESS,
            "err": ERROR,
            "warn": WARN,
            "dim": TEXT_DIM,
        }
        color = colors.get(tag, TEXT_SEC)
        safe_msg = html.escape(str(msg)).replace("\n", "<br>")
        self.log.moveCursor(QTextCursor.End)
        self.log.insertHtml(
            f'<span style="color:{QColor(color).name()};">{safe_msg}</span><br>'
        )
        self.log.ensureCursorVisible()

    def _set_busy(self, busy: bool):
        self.unlock_btn.setEnabled(not busy)
        self.folder_btn.setEnabled(not busy)
        self.auto_find_btn.setEnabled(not busy)
        self.output_btn.setEnabled(not busy)
        self.run_btn.setEnabled(not busy and self.backup_unlocked)

    def _unlock_backup(self):
        selected = self.folder_input.text().strip()
        resolved = resolve_backup_folder(selected)
        if not resolved:
            self._show_message(
                "Backup Not Found",
                "Could not find a valid iPhone backup folder. Use Auto Find or browse to a folder that contains Manifest.db.",
                icon=QMessageBox.Icon.Warning,
            )
            self.folder_status.setStyleSheet(f"color:{WARN}; font-size:12px;")
            self.folder_status.setText("No valid backup folder could be resolved.")
            return

        self._real_folder = str(resolved)
        self.folder_input.setText(self._real_folder)

        password = self.password_input.text()
        self._set_busy(True)
        self.status_label.setText("Validating backup…")
        self.progress.setRange(0, 0)
        self.count_label.setText("")
        self._log_write(f"Opening backup: {self._real_folder}", "dim")

        self._validation_thread = QThread()
        self._validation_worker = ValidationWorker(self._real_folder, password)
        self._validation_worker.moveToThread(self._validation_thread)
        self._validation_thread.started.connect(self._validation_worker.run)
        self._validation_worker.finished.connect(self._on_validation_finished)
        self._validation_worker.finished.connect(self._validation_thread.quit)
        self._validation_worker.finished.connect(self._validation_worker.deleteLater)
        self._validation_thread.finished.connect(self._validation_thread.deleteLater)
        self._validation_thread.start()

    def _on_validation_finished(self, ok: bool, message: str):
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self._set_busy(False)

        if ok:
            self.backup_unlocked = True
            self._unlock_categories()
            self.status_label.setText("Backup unlocked.")
            self.folder_status.setStyleSheet(f"color:{SUCCESS}; font-size:12px;")
            self.folder_status.setText(f"Using backup: {Path(self._real_folder).name}")
            self._log_write(message or "Backup unlocked.", "ok")

            current_output = self.output_input.text().strip()
            if not current_output:
                current_output = str(default_output_dir())
                self.output_input.setText(current_output)
            self._real_output = current_output
        else:
            self.backup_unlocked = False
            self._lock_categories()
            self.status_label.setText("Unlock failed.")
            self.folder_status.setStyleSheet(f"color:{ERROR}; font-size:12px;")
            self.folder_status.setText("Could not open backup. Check password or backup integrity.")
            self._log_write(message, "err")

    def _run_extract(self):
        if not self.backup_unlocked:
            self._show_message("Unlock First", "Unlock the backup first.")
            return

        output = self.output_input.text().strip()
        if not output:
            self._show_message("Output Required", "Choose an output folder first.", icon=QMessageBox.Icon.Warning)
            return

        opts = {key: row.is_checked() for key, row in self.category_rows.items()}
        if not any(opts.values()):
            self._show_message("Nothing Selected", "Choose at least one item to extract.", icon=QMessageBox.Icon.Warning)
            return

        Path(output).mkdir(parents=True, exist_ok=True)
        self._real_output = output
        self._set_busy(True)
        self.status_label.setText("Preparing extraction…")
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.count_label.setText("")
        self._log_write(f"Output folder: {output}", "dim")

        self._decrypt_thread = QThread()
        self._decrypt_worker = DecryptWorker(
            self._real_folder,
            self.password_input.text(),
            self._real_output,
            opts,
        )
        self._decrypt_worker.moveToThread(self._decrypt_thread)
        self._decrypt_thread.started.connect(self._decrypt_worker.run)

        self._decrypt_worker.log.connect(self._log_write)
        self._decrypt_worker.status.connect(lambda txt, _mode: self.status_label.setText(txt))
        self._decrypt_worker.count.connect(lambda n: self.count_label.setText(f"{n} file(s) extracted"))
        self._decrypt_worker.progress_value.connect(self.progress.setValue)
        self._decrypt_worker.progress_busy.connect(self._progress_busy)
        self._decrypt_worker.finished.connect(self._on_extract_finished)
        self._decrypt_worker.finished.connect(self._decrypt_thread.quit)
        self._decrypt_worker.finished.connect(self._decrypt_worker.deleteLater)
        self._decrypt_thread.finished.connect(self._decrypt_thread.deleteLater)
        self._decrypt_thread.start()

    def _progress_busy(self, busy: bool):
        if busy:
            self.progress.setRange(0, 0)
        else:
            self.progress.setRange(0, 100)
            self.progress.setValue(min(max(self.progress.value(), 0), 100))

    def _on_extract_finished(self, ok: bool, count: int):
        self._set_busy(False)
        self.progress.setRange(0, 100)
        self.progress.setValue(100 if ok else max(self.progress.value(), 0))
        self.status_label.setText("Done." if ok else "Finished with errors.")
        self.count_label.setText(f"{count} file(s) extracted")
        self._log_write(
            f"Extraction complete — {count} file(s) written to {self._real_output}",
            "ok" if ok else "warn",
        )

        if ok:
            self._show_message(
                "Done",
                f"Extraction finished. Files saved to:\n{self._real_output}",
                allow_open_folder=True,
            )
        else:
            self._show_message(
                "Finished with Errors",
                f"Extraction finished with warnings/errors. Check the log.\n\nOutput: {self._real_output}",
                icon=QMessageBox.Icon.Warning,
                allow_open_folder=True,
            )



if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(app_icon())
    win = App()
    win.show()
    sys.exit(app.exec())
