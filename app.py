import sys
import os
import sqlite3
import smtplib
import json
import re
import html
import ctypes
import faulthandler
import logging
import math
import tempfile
import threading
import traceback
from pathlib import Path
from contextlib import contextmanager
from logging.handlers import RotatingFileHandler
from email.message import EmailMessage
from datetime import datetime, timedelta

from PySide6.QtCore import (
    Qt,
    QTimer,
    QSize,
    QUrl,
    QObject,
    Signal,
    QRunnable,
    QThreadPool,
    QLockFile,
    qInstallMessageHandler,
)
from PySide6.QtGui import QAction, QColor, QDesktopServices, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QFormLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QMessageBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QHBoxLayout,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QSplitter,
    QMenu,
    QFrame,
    QSizePolicy,
    QAbstractItemView,
    QHeaderView,
    QStyledItemDelegate,
    QStyle,
    QScrollArea,
)


# ------------------------
# CONFIG
# ------------------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
SENDER_EMAIL = "intersportreminders@gmail.com"
#APP_PASSWORD = os.getenv("EMAIL_PASS", "").strip()
APP_PASSWORD = "jqfa pqxk rpgy lydx"

APP_NAME = "IntersportApp"
APP_VERSION = "1.1.3.0"
DB_FILENAME = "store.db"
SETTINGS_FILENAME = "settings.json"
INTERSPORT_QSS_FILE = "intersport.qss"
SPORTSTORE_QSS_FILE = "Sportstore.qss"
INTERSPORT_LOGO_FILE = "KOBALT ZONDER PLAATSNAAM.png"
SPORTSTORE_LOGO_FILE = "Sportstore kobalt.png"

REMINDER_AFTER_DAYS = 7
REMINDER_REPEAT_DAYS = 7
AUTO_CHECK_EVERY_MS = 60 * 60 * 1000  # 1 hour
REMINDER_HEARTBEAT_MS = 60 * 1000  # short timer survives sleep/resume more reliably
DB_BUSY_TIMEOUT_MS = 3000
EMAIL_TIMEOUT_SECONDS = 10
EMAIL_QUOTA_COOLDOWN_HOURS = 24
MAX_EMAIL_WORKERS = 3
BACKUP_RETENTION = 7


LOGGER = logging.getLogger(APP_NAME)
_LOG_PATH = None
_FAULT_LOG_STREAM = None
_DATA_DIR_READY = False
_MIGRATION_NOTICES = []
_SETTINGS_LOCK = threading.RLock()


class DataMigrationError(RuntimeError):
    pass


def _configure_diagnostics() -> Path:
    """Create a persistent startup/crash log, with a temp-folder fallback."""
    global _LOG_PATH, _FAULT_LOG_STREAM
    if _LOG_PATH is not None:
        return _LOG_PATH

    candidates = [
        _canonical_appdata_dir() / "logs",
        Path(tempfile.gettempdir()) / APP_NAME / "logs",
    ]
    last_error = None
    for directory in candidates:
        try:
            directory.mkdir(parents=True, exist_ok=True)
            log_path = directory / "IntersportApp.log"
            handler = RotatingFileHandler(
                log_path,
                maxBytes=1_000_000,
                backupCount=3,
                encoding="utf-8",
            )
            handler.setFormatter(
                logging.Formatter("%(asctime)s %(levelname)s %(message)s")
            )
            LOGGER.setLevel(logging.INFO)
            LOGGER.addHandler(handler)
            LOGGER.propagate = False

            fault_path = directory / "IntersportApp-crash.log"
            _FAULT_LOG_STREAM = fault_path.open("a", encoding="utf-8")
            faulthandler.enable(file=_FAULT_LOG_STREAM, all_threads=True)

            _LOG_PATH = log_path
            return log_path
        except (OSError, RuntimeError) as exc:
            last_error = exc

    raise OSError(f"Diagnostisch logboek kon niet worden aangemaakt: {last_error}")


def _qt_message_handler(_message_type, _context, message):
    LOGGER.warning("Qt: %s", message)


def _report_fatal_error(title: str, exc: BaseException):
    LOGGER.critical("%s: %s", title, exc, exc_info=(type(exc), exc, exc.__traceback__))
    log_path = str(_LOG_PATH or "niet beschikbaar")
    message = (
        f"{title}. De applicatie wordt afgesloten.\n\n"
        f"Reden: {exc}\n\n"
        f"Diagnostisch logboek: {log_path}"
    )
    try:
        QMessageBox.critical(None, "IntersportApp kan niet starten", message)
    except Exception:
        pass


def _unhandled_exception_hook(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    LOGGER.critical(
        "Onverwachte fout:\n%s",
        "".join(traceback.format_exception(exc_type, exc_value, exc_traceback)),
    )
    try:
        QMessageBox.critical(
            None,
            "Onverwachte fout",
            f"Er is een onverwachte fout opgetreden.\n\n"
            f"Reden: {exc_value}\n\n"
            f"Diagnostisch logboek: {_LOG_PATH or 'niet beschikbaar'}",
        )
    except Exception:
        sys.__excepthook__(exc_type, exc_value, exc_traceback)


def _windows_package_family_name() -> str:
    """Return the current MSIX package family, or an empty string when unpackaged."""
    if os.name != "nt":
        return ""

    try:
        get_family_name = ctypes.WinDLL("kernel32", use_last_error=True).GetCurrentPackageFamilyName
        get_family_name.argtypes = [ctypes.POINTER(ctypes.c_uint32), ctypes.c_wchar_p]
        get_family_name.restype = ctypes.c_long

        length = ctypes.c_uint32(0)
        result = get_family_name(ctypes.byref(length), None)
        if result == 15700:  # APPMODEL_ERROR_NO_PACKAGE
            return ""
        if result != 122 or not length.value:  # ERROR_INSUFFICIENT_BUFFER
            return ""

        buffer = ctypes.create_unicode_buffer(length.value)
        result = get_family_name(ctypes.byref(length), buffer)
        return buffer.value if result == 0 else ""
    except (AttributeError, OSError):
        return ""


def _canonical_appdata_dir() -> Path:
    """Use MSIX LocalState when packaged and Local AppData otherwise."""
    override = os.getenv("INTERSPORT_DATA_DIR", "").strip()
    if override:
        return Path(override).expanduser()

    local_appdata = Path(os.getenv("LOCALAPPDATA", os.getcwd()))
    package_family = _windows_package_family_name()
    if package_family:
        return local_appdata / "Packages" / package_family / "LocalState" / APP_NAME
    return local_appdata / APP_NAME


def _legacy_data_dirs(target: Path) -> list[Path]:
    """Known locations used by older unpackaged and virtualized MSIX releases."""
    candidates = []
    roaming = os.getenv("APPDATA", "").strip()
    local = os.getenv("LOCALAPPDATA", "").strip()

    if roaming:
        candidates.append(Path(roaming) / APP_NAME)
    if local:
        local_path = Path(local)
        candidates.append(local_path / APP_NAME)
        packages_path = local_path / "Packages"
        if packages_path.is_dir():
            for package_path in packages_path.glob(f"{APP_NAME}_*"):
                candidates.extend(
                    [
                        package_path / "LocalCache" / "Roaming" / APP_NAME,
                        package_path / "LocalCache" / "Local" / APP_NAME,
                        package_path / "RoamingState" / APP_NAME,
                        package_path / "LocalState" / APP_NAME,
                    ]
                )

    target_key = os.path.normcase(os.path.abspath(target))
    unique = []
    seen = {target_key}
    for candidate in candidates:
        key = os.path.normcase(os.path.abspath(candidate))
        if key not in seen:
            seen.add(key)
            unique.append(candidate)
    return unique


def _valid_database(path: Path) -> bool:
    conn = None
    try:
        uri = path.resolve().as_uri() + "?mode=ro"
        conn = sqlite3.connect(uri, uri=True, timeout=5)
        result = conn.execute("PRAGMA quick_check").fetchone()
        return bool(result) and result[0] == "ok"
    except (OSError, sqlite3.Error, TypeError, IndexError):
        return False
    finally:
        if conn is not None:
            conn.close()


def _migrate_database(source: Path, destination: Path):
    temp_path = destination.with_name(f"{destination.name}.migrating-{os.getpid()}")
    source_conn = None
    destination_conn = None
    try:
        source_uri = source.resolve().as_uri() + "?mode=ro"
        source_conn = sqlite3.connect(source_uri, uri=True, timeout=10)
        destination_conn = sqlite3.connect(temp_path, timeout=10)
        source_conn.backup(destination_conn)
        destination_conn.commit()
        destination_conn.close()
        destination_conn = None
        source_conn.close()
        source_conn = None
        if not destination.exists():
            os.replace(temp_path, destination)
    finally:
        if destination_conn is not None:
            destination_conn.close()
        if source_conn is not None:
            source_conn.close()
        if temp_path.exists():
            temp_path.unlink()


def _migrate_legacy_data(target: Path):
    target.mkdir(parents=True, exist_ok=True)
    legacy_dirs = _legacy_data_dirs(target)
    LOGGER.info("Actieve gegevensmap: %s", target)

    destination_db = target / DB_FILENAME
    if not destination_db.exists():
        database_candidates = [
            path / DB_FILENAME
            for path in legacy_dirs
            if (path / DB_FILENAME).is_file()
        ]
        database_candidates.sort(key=lambda path: path.stat().st_mtime, reverse=True)
        LOGGER.info(
            "Databasekandidaten: %s",
            ", ".join(str(path) for path in database_candidates) or "geen",
        )

        migrated = False
        failures = []
        for source_db in database_candidates:
            if not _valid_database(source_db):
                LOGGER.warning("Ongeldige database overgeslagen: %s", source_db)
                failures.append(f"{source_db} (ongeldige database)")
                continue
            try:
                _migrate_database(source_db, destination_db)
                LOGGER.info("Database gemigreerd vanuit %s", source_db)
                _MIGRATION_NOTICES.append(f"Database gemigreerd vanuit {source_db}")
                migrated = True
                break
            except (OSError, sqlite3.Error) as exc:
                LOGGER.exception("Migratie mislukt vanuit %s", source_db)
                failures.append(f"{source_db} ({exc})")

        if database_candidates and not migrated:
            raise DataMigrationError(
                "Bestaande database gevonden, maar migratie is mislukt:\n" + "\n".join(failures)
            )

    destination_settings = target / SETTINGS_FILENAME
    if not destination_settings.exists():
        settings_candidates = [
            path / SETTINGS_FILENAME
            for path in legacy_dirs
            if (path / SETTINGS_FILENAME).is_file()
        ]
        settings_candidates.sort(key=lambda path: path.stat().st_mtime, reverse=True)
        for source_settings in settings_candidates:
            try:
                with source_settings.open("r", encoding="utf-8") as source_file:
                    settings = json.load(source_file)
                _write_json_atomic(destination_settings, settings)
                LOGGER.info("Instellingen gemigreerd vanuit %s", source_settings)
                _MIGRATION_NOTICES.append(f"Instellingen gemigreerd vanuit {source_settings}")
                break
            except (OSError, json.JSONDecodeError, TypeError):
                continue


def _appdata_dir() -> str:
    global _DATA_DIR_READY
    target = _canonical_appdata_dir()
    if not _DATA_DIR_READY:
        _migrate_legacy_data(target)
        _DATA_DIR_READY = True
    else:
        target.mkdir(parents=True, exist_ok=True)
    return str(target)


def _db_path() -> str:
    return os.path.join(_appdata_dir(), DB_FILENAME)


def _settings_path() -> str:
    return os.path.join(_appdata_dir(), SETTINGS_FILENAME)


def resource_path(relative_path: str) -> str:
    """
    Works in dev + PyInstaller (onedir/onefile).
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)


# ------------------------
# VALIDATION: EMAIL / PHONE
# ------------------------
_EMAIL_REGEX = re.compile(
    r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$"
)


def _normalize_phone(raw: str) -> str | None:
    s = (raw or "").strip()
    if not s:
        return None

    s_no_spaces = re.sub(r"\s+", "", s)
    digits = re.sub(r"\D", "", s)

    # Nederlands mobiel: 06XXXXXXXX -> +316XXXXXXXX
    if digits.startswith("06") and len(digits) == 10:
        return "+31" + digits[1:]

    # Nederlands mobiel al in E.164: +316XXXXXXXX
    if s_no_spaces.startswith("+316") and digits.startswith("316") and len(digits) == 11:
        return "+" + digits

    # Nederlands vast nummer: accepteren, maar NIET omzetten naar +31
    # Voorbeelden: 0101234567, 020 123 4567, 0413 121233, 077-1234567
    if digits.startswith("0") and len(digits) == 10 and not digits.startswith("06"):
        return s

    # Internationaal nummer in E.164-formaat: + gevolgd door 8 t/m 15 cijfers
    if s_no_spaces.startswith("+") and 8 <= len(digits) <= 15:
        return "+" + digits

    return None


def validate_email_or_phone(value: str) -> tuple[bool, str, str]:
    v = (value or "").strip()
    if not v:
        return False, "", "Vul een geldig e-mailadres of telefoonnummer in."

    phone_norm = _normalize_phone(v)
    if phone_norm is not None:
        return True, phone_norm, ""

    if _EMAIL_REGEX.match(v):
        return True, v, ""

    return (
        False,
        v,
        "Vul een geldig e-mailadres of telefoonnummer in.\n\n"
        "Toegestaan:\n"
        "• e-mail (bijv. naam@domein.nl)\n"
        "• mobiel: 06XXXXXXXX of +316XXXXXXXX\n"
        "• vast NL-nummer: bijv. 0413121233 of 0413 121233\n"
        "• internationaal: + gevolgd door 8 t/m 15 cijfers"
    )


def is_email(value: str) -> bool:
    v = (value or "").strip()
    return bool(v) and bool(_EMAIL_REGEX.match(v))


# ------------------------
# DROPDOWN/DASHBOARD COLORS
# ------------------------
CONTACTSTATUS_COLORS = {
    "niet bereikbaar/nog bellen": "#7F7F7F",
    "Bericht gestuurd": "#FFEB9C",
    "Gebeld": "#C6EFCE",
    "s.v.p. bellen": "#FFC7CE",
}

BETAALSTATUS_COLORS = {
    "Betaald": "#C6EFCE",
    "Niet betaald!": "#FFC7CE",
    "Op factuur": "#E8EEF9",
}

BESTELSTATUS_COLORS = {
    "Nog opvragen/bestellen": "#FFC7CE",
    "Onderweg/Besteld": "#FFEB9C",
    "Onderweg/besteld": "#FFEB9C",
    "Op locatie": "#92D050",
    "Opgehaald/Afgewikkeld": "#C6EFCE",
}


def _readable_text_color(bg_hex: str) -> QColor:
    h = bg_hex.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return QColor("#0F172A") if lum > 160 else QColor("#FFFFFF")


def _normalize_key(s: str) -> str:
    return (s or "").strip()


def get_status_color(value: str, color_map: dict[str, str]) -> str | None:
    v = _normalize_key(value)
    if not v:
        return None
    if v in color_map:
        return color_map[v]
    vl = v.casefold()
    for k, col in color_map.items():
        if k.casefold() == vl:
            return col
    return None


def apply_combo_colors(combo: QComboBox, color_map: dict[str, str]):
    model = combo.model()
    for i in range(combo.count()):
        text = combo.itemText(i)
        if not text:
            continue
        hex_color = get_status_color(text, color_map)
        if not hex_color:
            continue
        bg = QColor(hex_color)
        fg = _readable_text_color(hex_color)

        idx = model.index(i, 0)
        model.setData(idx, bg, Qt.BackgroundRole)
        model.setData(idx, fg, Qt.ForegroundRole)

    def _update_current_style():
        t = combo.currentText()
        c = get_status_color(t, color_map)
        if c:
            fg = _readable_text_color(c).name()
            combo.setStyleSheet(
                f"QComboBox{{ background:{c}; color:{fg}; border:1px solid #D7DEEA; "
                f"padding:10px 12px; border-radius:12px; }}"
                f"QComboBox::drop-down{{ border:none; width:26px; }}"
            )
        else:
            combo.setStyleSheet("")

    combo.currentTextChanged.connect(lambda _: _update_current_style())
    _update_current_style()


# ------------------------
# TABLE DELEGATE
# ------------------------
class StatusColorDelegate(QStyledItemDelegate):
    def __init__(self, color_map: dict[str, str], parent=None):
        super().__init__(parent)
        self.color_map = color_map

    def paint(self, painter, option, index):
        value = index.data(Qt.DisplayRole) or ""
        hex_color = get_status_color(str(value), self.color_map)

        if hex_color:
            opt = option
            bg = QColor(hex_color)
            fg = _readable_text_color(hex_color)

            painter.save()
            painter.fillRect(opt.rect, bg)

            if opt.state & QStyle.State_Selected:  # type: ignore
                sel = QColor(22, 65, 150, 40)
                painter.fillRect(opt.rect, sel)

            text_rect = opt.rect.adjusted(8, 0, -8, 0)
            painter.setPen(fg)
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, str(value))
            painter.restore()
            return

        super().paint(painter, option, index)


# ------------------------
# DATABASE
# ------------------------
REQUEST_COLUMNS = (
    "id, klantnaam, verkoper, email, opmerking, maat, filiaal, collega, contactstatus, betaalstatus, bestelstatus, "
    "productcode, eancode, adviesprijs, afgerond, created_at, reminder_last_sent_at"
)


def connect():
    conn = sqlite3.connect(_db_path(), timeout=DB_BUSY_TIMEOUT_MS / 1000)
    conn.row_factory = sqlite3.Row
    conn.execute(f"PRAGMA busy_timeout={DB_BUSY_TIMEOUT_MS}")
    return conn


@contextmanager
def db_connection():
    conn = connect()
    try:
        with conn:
            yield conn
    finally:
        conn.close()


def init_db():
    with db_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                klantnaam TEXT NOT NULL,
                verkoper TEXT NOT NULL,
                email TEXT NOT NULL,
                opmerking TEXT NOT NULL,
                maat TEXT,
                filiaal TEXT,
                collega TEXT,
                contactstatus TEXT,
                betaalstatus TEXT,
                bestelstatus TEXT,
                productcode TEXT,
                eancode TEXT,
                adviesprijs REAL,
                afgerond INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        existing_columns = {
            row["name"] if isinstance(row, sqlite3.Row) else row[1]
            for row in conn.execute("PRAGMA table_info(requests)").fetchall()
        }
        if "id" not in existing_columns:
            raise DataMigrationError("De tabel requests heeft geen geldige id-kolom.")

        expected_columns = {
            "klantnaam": "TEXT NOT NULL DEFAULT ''",
            "verkoper": "TEXT NOT NULL DEFAULT ''",
            "email": "TEXT NOT NULL DEFAULT ''",
            "opmerking": "TEXT NOT NULL DEFAULT ''",
            "maat": "TEXT",
            "filiaal": "TEXT",
            "collega": "TEXT",
            "contactstatus": "TEXT",
            "betaalstatus": "TEXT",
            "bestelstatus": "TEXT",
            "productcode": "TEXT",
            "eancode": "TEXT",
            "adviesprijs": "REAL",
            "afgerond": "INTEGER NOT NULL DEFAULT 0",
            "created_at": "TEXT",
            "reminder_last_sent_at": "TEXT",
        }
        for column_name, declaration in expected_columns.items():
            if column_name not in existing_columns:
                LOGGER.info("Databasekolom toevoegen: %s", column_name)
                conn.execute(
                    f'ALTER TABLE requests ADD COLUMN "{column_name}" {declaration}'
                )

        conn.execute(
            "UPDATE requests SET created_at=? WHERE created_at IS NULL OR created_at=''",
            (_dt_str(datetime.now()),),
        )
        conn.execute("UPDATE requests SET afgerond=0 WHERE afgerond IS NULL")

        conn.execute("CREATE INDEX IF NOT EXISTS idx_requests_open_created ON requests(afgerond, created_at)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_requests_open_lastsent ON requests(afgerond, reminder_last_sent_at)")

    with db_connection() as conn:
        integrity = conn.execute("PRAGMA quick_check").fetchone()
        if not integrity or integrity[0] != "ok":
            raise DataMigrationError(
                f"Database-integriteitscontrole mislukt: {integrity[0] if integrity else 'onbekend'}"
            )
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")

    try:
        _create_daily_database_backup()
    except (OSError, sqlite3.Error):
        LOGGER.exception("Dagelijkse databaseback-up kon niet worden gemaakt")


def _create_daily_database_backup():
    source = Path(_db_path())
    if not source.is_file():
        return

    backup_dir = Path(_appdata_dir()) / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    destination = backup_dir / f"store-{datetime.now():%Y%m%d}.db"
    if not destination.exists():
        _migrate_database(source, destination)
        LOGGER.info("Dagelijkse databaseback-up gemaakt: %s", destination)

    backups = sorted(
        backup_dir.glob("store-*.db"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for old_backup in backups[BACKUP_RETENTION:]:
        try:
            old_backup.unlink()
        except OSError:
            LOGGER.warning("Oude databaseback-up kon niet worden verwijderd: %s", old_backup)


def add_request(data):
    with db_connection() as conn:
        cursor = conn.execute(
            """
            INSERT INTO requests (
                klantnaam, verkoper, email, opmerking, maat,
                filiaal, collega, contactstatus, betaalstatus,
                bestelstatus, productcode, eancode, adviesprijs,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            data,
        )
        return cursor.lastrowid


def get_all_requests():
    with db_connection() as conn:
        open_rows = conn.execute(
            f"SELECT {REQUEST_COLUMNS} FROM requests WHERE afgerond=0 ORDER BY id DESC"
        ).fetchall()
        done_rows = conn.execute(
            f"SELECT {REQUEST_COLUMNS} FROM requests WHERE afgerond=1 ORDER BY id DESC LIMIT 25"
        ).fetchall()
        return list(open_rows) + list(done_rows)


def get_request_by_id(row_id: int):
    with db_connection() as conn:
        return conn.execute(
            f"SELECT {REQUEST_COLUMNS} FROM requests WHERE id=?",
            (row_id,),
        ).fetchone()


def set_request_completed(row_id: int, completed: bool) -> bool:
    with db_connection() as conn:
        cursor = conn.execute(
            "UPDATE requests SET afgerond=? WHERE id=?",
            (1 if completed else 0, row_id),
        )
        return cursor.rowcount == 1


def update_request(data) -> bool:
    with db_connection() as conn:
        cursor = conn.execute(
            """UPDATE requests
               SET klantnaam=?, verkoper=?, email=?, opmerking=?, maat=?,
                   filiaal=?, collega=?, contactstatus=?, betaalstatus=?, bestelstatus=?,
                   productcode=?, eancode=?, adviesprijs=?
               WHERE id=?""",
            data,
        )
        return cursor.rowcount == 1


# ------------------------
# SETTINGS (JSON)
# ------------------------
def _write_json_atomic(path: Path, data):
    temp_path = path.with_name(
        f"{path.name}.tmp-{os.getpid()}-{threading.get_ident()}"
    )
    try:
        with temp_path.open("w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def load_settings():
    with _SETTINGS_LOCK:
        path = _settings_path()
        if not os.path.exists(path):
            return {}
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}


def save_settings(data):
    with _SETTINGS_LOCK:
        _write_json_atomic(Path(_settings_path()), data)


def get_setting(key):
    return load_settings().get(key, "")


def set_setting(key, value):
    with _SETTINGS_LOCK:
        settings = load_settings()
        settings[key] = value
        save_settings(settings)


def get_active_theme() -> str:
    theme = str(get_setting("style_theme") or "Intersport").strip()
    return theme if theme in ("Intersport", "Sportstore") else "Intersport"


def get_active_qss_file() -> str:
    return SPORTSTORE_QSS_FILE if get_active_theme() == "Sportstore" else INTERSPORT_QSS_FILE


# ------------------------
# EMAIL WORKER
# ------------------------
class _EmailWorkerSignals(QObject):
    finished = Signal(object)
    error = Signal(str)


class EmailWorker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = _EmailWorkerSignals()

    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
            try:
                self.signals.finished.emit(result)
            except RuntimeError:
                LOGGER.info("Achtergrondtaak afgerond nadat het venster was gesloten")
        except Exception as e:
            try:
                self.signals.error.emit(str(e))
            except RuntimeError:
                LOGGER.warning("Achtergrondtaak mislukte na afsluiten: %s", e)


def _require_email_password():
    if not APP_PASSWORD:
        raise RuntimeError(
            "EMAIL_PASS environment variable ontbreekt.\n\n"
            "Zet deze op je Gmail app-wachtwoord om e-mail te kunnen versturen.\n"
            "Voorbeeld:\n"
            "  Windows (PowerShell):  $env:EMAIL_PASS='xxxx xxxx xxxx xxxx'\n"
        )


class EmailQuotaExceededError(RuntimeError):
    pass


class EmailCooldownError(RuntimeError):
    pass


def _email_quota_cooldown_until():
    raw_value = str(get_setting("email_quota_cooldown_until") or "").strip()
    if not raw_value:
        return None
    try:
        cooldown_until = datetime.fromisoformat(raw_value)
    except ValueError:
        set_setting("email_quota_cooldown_until", "")
        return None
    if cooldown_until <= datetime.now():
        set_setting("email_quota_cooldown_until", "")
        return None
    return cooldown_until


def _format_cooldown(cooldown_until: datetime) -> str:
    return cooldown_until.strftime("%d-%m-%Y om %H:%M")


def _smtp_send_message(msg: EmailMessage):
    """Send once, and stop retrying locally after Gmail reports its daily limit."""
    _require_email_password()
    cooldown_until = _email_quota_cooldown_until()
    if cooldown_until is not None:
        raise EmailCooldownError(
            "E-mailverzending is tijdelijk gepauzeerd tot "
            f"{_format_cooldown(cooldown_until)}, omdat Gmail eerder meldde dat "
            "de dagelijkse verzendlimiet is bereikt. De opgeslagen gegevens blijven behouden."
        )

    try:
        with smtplib.SMTP_SSL(
            SMTP_SERVER, SMTP_PORT, timeout=EMAIL_TIMEOUT_SECONDS
        ) as server:
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(msg)
    except smtplib.SMTPResponseException as exc:
        detail = exc.smtp_error
        if isinstance(detail, bytes):
            detail = detail.decode("utf-8", errors="replace")
        detail = str(detail).replace("\\n", " ").strip()
        is_daily_limit = exc.smtp_code == 550 and (
            "5.4.5" in detail or "daily user sending limit" in detail.lower()
        )
        if is_daily_limit:
            cooldown_until = datetime.now() + timedelta(hours=EMAIL_QUOTA_COOLDOWN_HOURS)
            set_setting(
                "email_quota_cooldown_until",
                cooldown_until.isoformat(timespec="seconds"),
            )
            LOGGER.error(
                "Gmail-verzendlimiet bereikt; verzending gepauzeerd tot %s",
                cooldown_until,
            )
            raise EmailQuotaExceededError(
                "Gmail heeft de dagelijkse verzendlimiet bereikt. Nieuwe e-mails worden "
                f"op deze computer gepauzeerd tot {_format_cooldown(cooldown_until)}. "
                "De bestelling is wel opgeslagen."
            ) from exc
        if isinstance(exc, smtplib.SMTPAuthenticationError):
            raise RuntimeError(
                "Aanmelden bij Gmail is mislukt. Controleer het Gmail-account en het app-wachtwoord."
            ) from exc
        raise RuntimeError(
            f"Gmail heeft de e-mail geweigerd (code {exc.smtp_code}): {detail}"
        ) from exc
    except (TimeoutError, OSError, smtplib.SMTPException) as exc:
        raise RuntimeError(
            "De e-mailserver is niet bereikbaar of reageerde niet op tijd. "
            "Controleer de internetverbinding en probeer het later opnieuw."
        ) from exc


def send_test_email(receiver):
    msg = EmailMessage()
    msg["Subject"] = "Test Email - Intersport Reminder App"
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver
    msg.set_content("Dit is een testmail vanuit de Intersport Reminder App.")

    _smtp_send_message(msg)


def send_customer_arrival_email(request_row):
    _require_email_password()

    def _get_active_brand():
        theme = get_active_theme()
        return "SPORTSTORE" if theme == "Sportstore" else "INTERSPORT"

    brand_name = _get_active_brand()
    receiver_email = (request_row["email"] or "").strip()
    verkoper = (request_row["verkoper"] or "").strip()

    if not receiver_email:
        raise RuntimeError("Geen e-mailadres gevonden voor deze klant.")

    logo_filename = SPORTSTORE_LOGO_FILE if get_active_theme() == "Sportstore" else INTERSPORT_LOGO_FILE

    logo_path = resource_path(logo_filename)
    logo_cid = "brandlogo"

    subject = f"Uw bestelling is binnen bij {brand_name}"

    text_body = (
        f"Beste {request_row['klantnaam'] or 'klant'},\n\n"
        "Goed nieuws! Het artikel dat u bij ons heeft besteld, is zojuist bij ons in de winkel aangekomen.\n"
        "Uw bestelling ligt voor u klaar en kan op elk gewenst moment tijdens onze openingstijden worden opgehaald. Wij helpen u graag verder wanneer u langskomt.\n\n"
        f"Met sportieve groeten,\n{verkoper}"
    )

    klant_naam = html.escape((request_row["klantnaam"] or "klant").strip())
    verkoper_html = html.escape(verkoper)

    html_body = f"""\
<html>
  <body style="margin:0; padding:0; background-color:#f4f6f8; font-family:Arial, Helvetica, sans-serif; color:#1f2937;">
    <div style="max-width:720px; margin:0 auto; padding:32px 20px;">
      <div style="background:#ffffff; border:1px solid #e5e7eb; border-radius:14px; overflow:hidden;">
        <div style="padding:24px 24px 10px 24px; text-align:left;">
          <img src="cid:{logo_cid}" alt="{brand_name} logo" style="max-width:260px; max-height:70px; height:auto; width:auto;">
        </div>

        <div style="padding:8px 24px 28px 24px;">
          <h1 style="margin:0 0 18px 0; font-size:24px; line-height:1.3; color:#163e96;">
            Uw bestelling is binnen
          </h1>

          <p style="margin:0 0 14px 0; font-size:15px; line-height:1.7;">
            Beste {klant_naam},
          </p>

          <p style="margin:0 0 14px 0; font-size:15px; line-height:1.7;">
            Goed nieuws! Het artikel dat u bij ons heeft besteld, is zojuist bij ons in de winkel aangekomen.
          </p>

          <p style="margin:0 0 20px 0; font-size:15px; line-height:1.7;">
            Uw bestelling ligt voor u klaar en kan op elk gewenst moment tijdens onze openingstijden worden opgehaald. Wij helpen u graag verder wanneer u langskomt.
          </p>

          <p style="margin:24px 0 0 0; font-size:15px; line-height:1.7;">
            Met sportieve groeten,<br>
            <strong>{verkoper_html}</strong>
          </p>
        </div>
      </div>
    </div>
  </body>
</html>
"""

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg.set_content(text_body)
    msg.add_alternative(html_body, subtype="html")

    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            logo_data = f.read()

        if logo_path.lower().endswith(".png"):
            maintype, subtype = "image", "png"
        elif logo_path.lower().endswith(".jpg") or logo_path.lower().endswith(".jpeg"):
            maintype, subtype = "image", "jpeg"
        else:
            maintype, subtype = "image", "png"

        msg.get_payload()[-1].add_related(
            logo_data,
            maintype=maintype,
            subtype=subtype,
            cid=f"<{logo_cid}>",
            filename=os.path.basename(logo_path),
        )

    _smtp_send_message(msg)


def send_customer_confirmation_email(request_row):
    _require_email_password()

    def _get_active_brand():
        theme = get_active_theme()
        return "SPORTSTORE" if theme == "Sportstore" else "INTERSPORT"

    brand_name = _get_active_brand()
    receiver_email = (request_row["email"] or "").strip()
    verkoper = (request_row["verkoper"] or "").strip()

    if not receiver_email:
        raise RuntimeError("Geen e-mailadres gevonden voor deze klant.")

    logo_filename = SPORTSTORE_LOGO_FILE if get_active_theme() == "Sportstore" else INTERSPORT_LOGO_FILE

    logo_path = resource_path(logo_filename)
    logo_cid = "brandlogo"

    subject = f"Bedankt voor uw bestelling bij {brand_name}"

    text_body = (
        f"Beste {request_row['klantnaam'] or 'klant'},\n\n"
        "Hartelijk dank voor uw bestelling en voor het steunen van uw lokale ondernemer."
        "Wij gaan direct aan de slag om uw bestelling zo snel mogelijk voor u te verwerken.\n"
        "Zodra uw bestelling bij ons binnen is, ontvangt u direct een bericht van ons.\n\n"
        f"Met sportieve groeten,\n{verkoper}"
    )

    klant_naam = html.escape((request_row["klantnaam"] or "klant").strip())
    verkoper_html = html.escape(verkoper)

    html_body = f"""\
<html>
  <body style="margin:0; padding:0; background-color:#f4f6f8; font-family:Arial, Helvetica, sans-serif; color:#1f2937;">
    <div style="max-width:720px; margin:0 auto; padding:32px 20px;">
      <div style="background:#ffffff; border:1px solid #e5e7eb; border-radius:14px; overflow:hidden;">
        <div style="padding:24px 24px 10px 24px; text-align:left;">
          <img src="cid:{logo_cid}" alt="{brand_name} logo" style="max-width:260px; max-height:70px; height:auto; width:auto;">
        </div>

        <div style="padding:8px 24px 28px 24px;">
          <h1 style="margin:0 0 18px 0; font-size:24px; line-height:1.3; color:#163e96;">
            Bedankt voor uw bestelling
          </h1>

          <p style="margin:0 0 14px 0; font-size:15px; line-height:1.7;">
            Beste {klant_naam},
          </p>

          <p style="margin:0 0 14px 0; font-size:15px; line-height:1.7;">
            Hartelijk dank voor uw bestelling en voor het steunen van uw lokale ondernemer.
          </p>

          <p style="margin:0 0 14px 0; font-size:15px; line-height:1.7;">
            Wij gaan direct aan de slag om uw bestelling zo snel mogelijk voor u te verwerken.
          </p>

          <p style="margin:0 0 20px 0; font-size:15px; line-height:1.7;">
            Zodra uw bestelling bij ons binnen is, ontvangt u direct een bericht van ons.
          </p>

          <p style="margin:24px 0 0 0; font-size:15px; line-height:1.7;">
            Met sportieve groeten,<br>
            <strong>{verkoper_html}</strong>
          </p>
        </div>
      </div>
    </div>
  </body>
</html>
"""

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg.set_content(text_body)
    msg.add_alternative(html_body, subtype="html")

    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            logo_data = f.read()

        if logo_path.lower().endswith(".png"):
            maintype, subtype = "image", "png"
        elif logo_path.lower().endswith(".jpg") or logo_path.lower().endswith(".jpeg"):
            maintype, subtype = "image", "jpeg"
        else:
            maintype, subtype = "image", "png"

        msg.get_payload()[-1].add_related(
            logo_data,
            maintype=maintype,
            subtype=subtype,
            cid=f"<{logo_cid}>",
            filename=os.path.basename(logo_path),
        )

    _smtp_send_message(msg)


def _parse_dt(value):
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


def _dt_str(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%d %H:%M:%S")


# =========================================================
# MAIN WINDOW
# =========================================================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.threadpool = QThreadPool(self)
        self.threadpool.setMaxThreadCount(MAX_EMAIL_WORKERS)
        self._closing = False
        self._active_email_workers = []
        self._manual_reminders_in_progress = False
        self._auto_reminders_in_progress = False
        self._next_auto_reminder_check_at = None

        self.setWindowTitle("INTERSPORT • Reminders")
        self.resize(1360, 720)

        root = QVBoxLayout()
        root.setContentsMargins(7, 7, 7, 7)
        root.setSpacing(12)

        root.addWidget(self._build_header())

        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_form_tab(), "Invoeren")
        self.tabs.addTab(self.create_dashboard_tab(), "Dashboard")
        self.tabs.addTab(self.create_settings_tab(), "Instellingen")

        root.addWidget(self.tabs, 1)
        self.setLayout(root)

        self.load_styles()
        self.refresh_table(raise_errors=True)

        self.startup_reminder_timer = QTimer(self)
        self.startup_reminder_timer.setSingleShot(True)
        self.startup_reminder_timer.setInterval(3000)
        self.startup_reminder_timer.timeout.connect(self.run_auto_reminder_check_once)
        self.startup_reminder_timer.start()

        self.reminder_timer = QTimer(self)
        self.reminder_timer.setInterval(REMINDER_HEARTBEAT_MS)
        self.reminder_timer.timeout.connect(self._run_reminder_heartbeat)
        self.reminder_timer.start()

    def _track_email_worker(self, worker: EmailWorker):
        self._active_email_workers.append(worker)
        return worker

    def _release_email_worker(self, worker: EmailWorker):
        try:
            self._active_email_workers.remove(worker)
        except ValueError:
            pass

    def closeEvent(self, event):
        self._closing = True
        self.startup_reminder_timer.stop()
        self.reminder_timer.stop()
        self.search_filter_timer.stop()
        self.threadpool.clear()
        if self.threadpool.activeThreadCount():
            LOGGER.info(
                "Applicatie sluit terwijl %s achtergrondtaak/taken nog afronden",
                self.threadpool.activeThreadCount(),
            )
        event.accept()

    def _build_header(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("TopBar")
        bar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        bar.setMinimumHeight(62)

        lay = QHBoxLayout(bar)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(12)

        self.brand_logo = QLabel()
        self.brand_logo.setObjectName("BrandBadge")
        self.brand_logo.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.brand_logo.setMinimumHeight(40)
        self.brand_logo.setMinimumWidth(170)

        title = QLabel("Reminders Dashboard")
        title.setObjectName("HeaderTitle")

        subtitle = QLabel("Openstaande bestellingen beheren • automatische reminders na 7 dagen")
        subtitle.setObjectName("HeaderSubtitle")

        text_col = QVBoxLayout()
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.setSpacing(2)
        text_col.addWidget(title)
        text_col.addWidget(subtitle)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        quick_refresh = QPushButton("Ververs")
        quick_refresh.setObjectName("BtnSecondary")
        quick_refresh.setMinimumHeight(34)
        quick_refresh.clicked.connect(self.refresh_table)

        lay.addWidget(self.brand_logo)
        lay.addLayout(text_col)
        lay.addWidget(spacer)
        lay.addWidget(quick_refresh)
        return bar

    def _get_active_logo_file(self) -> str | None:
        theme = get_active_theme()
        candidates = []
        if theme == "Sportstore":
            candidates = [SPORTSTORE_LOGO_FILE, INTERSPORT_LOGO_FILE]
        else:
            candidates = [INTERSPORT_LOGO_FILE, SPORTSTORE_LOGO_FILE]

        for filename in candidates:
            full_path = resource_path(filename)
            if os.path.exists(full_path):
                return full_path
        return None

    def update_branding(self):
        logo_path = self._get_active_logo_file()
        theme = get_active_theme()

        if logo_path:
            pixmap = QPixmap(logo_path)
            if not pixmap.isNull():
                scaled = pixmap.scaledToHeight(40, Qt.SmoothTransformation)
                self.brand_logo.setPixmap(scaled)
                self.brand_logo.setText("")
                return

        self.brand_logo.setPixmap(QPixmap())
        self.brand_logo.setText(theme.upper())

    # =========================================================
    # TAB 1: FORM
    # =========================================================
    def create_form_tab(self):
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        widget = QWidget()
        outer = QVBoxLayout(widget)
        outer.setContentsMargins(18, 18, 18, 18)
        outer.setSpacing(12)

        card = QFrame()
        card.setObjectName("Card")
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(18, 18, 18, 18)
        card_lay.setSpacing(12)

        header = QLabel("Nieuwe bestelling / verzoek")
        header.setObjectName("CardTitle")
        hint = QLabel("Velden met * zijn verplicht.")
        hint.setObjectName("CardHint")

        card_lay.addWidget(header)
        card_lay.addWidget(hint)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)
        form.setFormAlignment(Qt.AlignTop)
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(10)

        self.inputs = {}

        self.inputs["klant"] = QLineEdit()
        self.inputs["verkoper"] = QLineEdit()
        self.inputs["email"] = QLineEdit()
        self.inputs["opmerking"] = QLineEdit()
        self.inputs["maat"] = QLineEdit()
        self.inputs["product"] = QLineEdit()
        self.inputs["ean"] = QLineEdit()
        self.inputs["prijs"] = QLineEdit()

        self.inputs["opmerking"].setPlaceholderText("Bijv. maat / kleur / levertijd / actie...")
        self.inputs["maat"].setPlaceholderText("bijv. 42 / M / 140")
        self.inputs["email"].setPlaceholderText("bijv. naam@domein.nl, 0612345678, 0413 121233 of +48123456789")
        self.inputs["prijs"].setPlaceholderText("bijv. 129,95")
        self.inputs["ean"].setPlaceholderText("bijv. 8712345678901")

        form.addRow("Klantnaam *", self.inputs["klant"])
        form.addRow("Verkoper *", self.inputs["verkoper"])
        form.addRow("E-mail of Telefoonnummer *", self.inputs["email"])
        form.addRow("bestelling *", self.inputs["opmerking"])
        form.addRow("Maat", self.inputs["maat"])

        product_ean_widget = QWidget()
        product_ean_layout = QHBoxLayout(product_ean_widget)
        product_ean_layout.setContentsMargins(0, 0, 0, 0)
        product_ean_layout.setSpacing(10)
        product_ean_layout.addWidget(self.inputs["product"])
        product_ean_layout.addWidget(self.inputs["ean"])
        form.addRow("Artikelnummer * / EAN code *", product_ean_widget)

        form.addRow("Adviesprijs", self.inputs["prijs"])

        self.inputs["filiaal"] = QComboBox()
        self.inputs["filiaal"].addItem("")
        self.inputs["filiaal"].addItems(["1) Helmond", "2) Veghel", "3) Venray", "4) Venlo", "5) Breda", "x) Leverancier"])
        self.inputs["collega"] = QLineEdit()

        filiaal_collega_widget = QWidget()
        filiaal_collega_layout = QHBoxLayout(filiaal_collega_widget)
        filiaal_collega_layout.setContentsMargins(0, 0, 0, 0)
        filiaal_collega_layout.setSpacing(10)
        filiaal_collega_layout.addWidget(self.inputs["filiaal"])
        filiaal_collega_layout.addWidget(self.inputs["collega"])
        form.addRow("Aangesproken filiaal / collega", filiaal_collega_widget)

        self.inputs["contactstatus"] = QComboBox()
        self.inputs["contactstatus"].addItems(["s.v.p. bellen", "Gebeld", "Bericht gestuurd", "niet bereikbaar/nog bellen"])
        form.addRow("Contactstatus", self.inputs["contactstatus"])
        apply_combo_colors(self.inputs["contactstatus"], CONTACTSTATUS_COLORS)

        self.inputs["betaalstatus"] = QComboBox()
        self.inputs["betaalstatus"].addItems(["Niet betaald!", "Betaald", "Op factuur"])
        form.addRow("Betaalstatus", self.inputs["betaalstatus"])
        apply_combo_colors(self.inputs["betaalstatus"], BETAALSTATUS_COLORS)

        self.inputs["bestelstatus"] = QComboBox()
        self.inputs["bestelstatus"].addItems(["Nog opvragen/bestellen", "Onderweg/Besteld", "Op locatie", "Opgehaald/Afgewikkeld"])
        form.addRow("Bestelstatus", self.inputs["bestelstatus"])
        apply_combo_colors(self.inputs["bestelstatus"], BESTELSTATUS_COLORS)

        card_lay.addLayout(form)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)

        add_btn = QPushButton("Opslaan")
        add_btn.setObjectName("BtnPrimary")
        add_btn.setMinimumHeight(38)
        add_btn.clicked.connect(self.handle_add)

        btn_row.addWidget(add_btn)
        card_lay.addLayout(btn_row)

        outer.addWidget(card)
        outer.addStretch()

        scroll.setWidget(widget)
        container_layout.addWidget(scroll)

        return container

    # =========================================================
    # TAB 2: DASHBOARD
    # =========================================================
    def create_dashboard_tab(self):
        widget = QWidget()
        main_layout = QVBoxLayout(widget)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(12)

        top_row = QHBoxLayout()
        top_row.setSpacing(10)

        self.search = QLineEdit()
        self.search.setObjectName("SearchField")
        self.search.setPlaceholderText("Zoeken (klant, product, EAN, maat, bestelling, filiaal, verkoper...)")
        self.search_filter_timer = QTimer(self)
        self.search_filter_timer.setSingleShot(True)
        self.search_filter_timer.setInterval(200)
        self.search_filter_timer.timeout.connect(
            lambda: self.filter_tables(self.search.text())
        )
        self.search.textChanged.connect(lambda _text: self.search_filter_timer.start())

        top_row.addWidget(QLabel("Zoek:"))
        top_row.addWidget(self.search, 1)

        self.manual_send_btn = QPushButton("Verstuur reminders nu")
        self.manual_send_btn.setObjectName("BtnDanger")
        self.manual_send_btn.setMinimumHeight(36)
        self.manual_send_btn.clicked.connect(self.send_all_reminders_manual)

        top_row.addWidget(self.manual_send_btn)
        main_layout.addLayout(top_row)

        splitter = QSplitter(Qt.Vertical)

        open_widget = QFrame()
        open_widget.setObjectName("Card")
        open_layout = QVBoxLayout(open_widget)
        open_layout.setContentsMargins(7, 7, 7, 7)
        open_layout.setSpacing(10)

        open_label = QLabel("Openstaande bestellingen")
        open_label.setObjectName("SectionTitle")

        self.table_open = QTableWidget()
        self._setup_table(self.table_open)
        self.table_open.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_open.customContextMenuRequested.connect(lambda pos: self.show_context_menu(self.table_open, pos))
        self.table_open.itemDoubleClicked.connect(lambda item: self.edit_row(self.table_open, item.row()))

        open_layout.addWidget(open_label)
        open_layout.addWidget(self.table_open)

        done_widget = QFrame()
        done_widget.setObjectName("Card")
        done_layout = QVBoxLayout(done_widget)
        done_layout.setContentsMargins(7, 7, 7, 7)
        done_layout.setSpacing(10)

        done_label = QLabel("Afgeronde bestellingen")
        done_label.setObjectName("SectionTitle")

        self.table_done = QTableWidget()
        self._setup_table(self.table_done)
        self.table_done.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_done.customContextMenuRequested.connect(lambda pos: self.show_context_menu(self.table_done, pos))
        self.table_done.itemDoubleClicked.connect(lambda item: self.edit_row(self.table_done, item.row()))

        done_layout.addWidget(done_label)
        done_layout.addWidget(self.table_done)

        splitter.addWidget(open_widget)
        splitter.addWidget(done_widget)
        splitter.setSizes([int(self.height() * 0.72), int(self.height() * 0.28)])

        main_layout.addWidget(splitter, 1)

        note_label = QLabel("Tip: rechtermuisknop op een bestelling voor acties (bewerken/afronden).")
        note_label.setObjectName("HintText")
        main_layout.addWidget(note_label)
        return widget

    def _setup_table(self, table: QTableWidget):
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setWordWrap(True)
        table.setSortingEnabled(True)
        table.verticalHeader().setVisible(False)
        table.setShowGrid(False)
        table.setFocusPolicy(Qt.NoFocus)
        table.setIconSize(QSize(16, 16))

        hh = table.horizontalHeader()
        hh.setStretchLastSection(True)
        hh.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)

    def _install_status_delegates(self, table: QTableWidget):
        def find_col(name: str) -> int:
            for c in range(table.columnCount()):
                hdr = table.horizontalHeaderItem(c)
                if hdr and hdr.text().strip().casefold() == name.casefold():
                    return c
            return -1

        c_contact = find_col("Contact")
        c_betaal = find_col("Betaal")
        c_bestel = find_col("Bestel")

        if c_contact >= 0:
            table.setItemDelegateForColumn(c_contact, StatusColorDelegate(CONTACTSTATUS_COLORS, table))
        if c_betaal >= 0:
            table.setItemDelegateForColumn(c_betaal, StatusColorDelegate(BETAALSTATUS_COLORS, table))
        if c_bestel >= 0:
            table.setItemDelegateForColumn(c_bestel, StatusColorDelegate(BESTELSTATUS_COLORS, table))

    # =========================================================
    # TAB 3: SETTINGS
    # =========================================================
    def create_settings_tab(self):
        widget = QWidget()
        outer = QVBoxLayout(widget)
        outer.setContentsMargins(18, 18, 18, 18)
        outer.setSpacing(12)

        card = QFrame()
        card.setObjectName("Card")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(18, 18, 18, 18)
        lay.setSpacing(12)

        title = QLabel("E-mail instellingen")
        title.setObjectName("CardTitle")
        lay.addWidget(title)

        form = QFormLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(10)

        self.reminder_email_input = QLineEdit()
        self.reminder_email_input.setPlaceholderText("bijv. filiaal@intersport.nl")
        self.reminder_email_input.setText(get_setting("reminder_email"))
        form.addRow("Email voor reminders", self.reminder_email_input)

        self.test_email_input = QLineEdit()
        self.test_email_input.setPlaceholderText("Bijv. jouw mailadres om te testen")
        self.test_email_input.setText(get_setting("reminder_email"))
        form.addRow("Test e-mail naar", self.test_email_input)

        self.style_theme_input = QComboBox()
        self.style_theme_input.addItems(["Intersport", "Sportstore"])
        current_theme = self.style_theme_input.findText(get_active_theme())
        self.style_theme_input.setCurrentIndex(current_theme if current_theme >= 0 else 0)
        form.addRow("Opmaak versie", self.style_theme_input)

        lay.addLayout(form)

        btns = QHBoxLayout()
        btns.addStretch(1)

        self.test_btn = QPushButton("Verstuur test e-mail")
        self.test_btn.setObjectName("BtnSecondary")
        self.test_btn.setMinimumHeight(36)
        self.test_btn.clicked.connect(self.handle_send_test_email)

        save_btn = QPushButton("Opslaan")
        save_btn.setObjectName("BtnPrimary")
        save_btn.setMinimumHeight(36)
        save_btn.clicked.connect(self.save_settings)

        btns.addWidget(self.test_btn)
        btns.addWidget(save_btn)
        lay.addLayout(btns)

        info = QLabel("De automatische check draait bij opstarten en daarna ieder uur.")
        info.setObjectName("HintText")
        lay.addWidget(info)

        credit = QLabel(f"© 2026 Alexander Brinkman. Alle rechten voorbehouden. V{APP_VERSION}")
        credit.setObjectName("FooterCredit")
        credit.setAlignment(Qt.AlignRight)
        lay.addWidget(credit)

        outer.addWidget(card)

        storage_card = QFrame()
        storage_card.setObjectName("Card")
        storage_layout = QVBoxLayout(storage_card)
        storage_layout.setContentsMargins(18, 18, 18, 18)
        storage_layout.setSpacing(12)

        storage_title = QLabel("Gegevensopslag")
        storage_title.setObjectName("CardTitle")
        storage_layout.addWidget(storage_title)

        storage_form = QFormLayout()
        storage_form.setHorizontalSpacing(14)
        self.data_path_input = QLineEdit(_appdata_dir())
        self.data_path_input.setReadOnly(True)
        self.data_path_input.setToolTip(_appdata_dir())
        storage_form.addRow("Actieve gegevensmap", self.data_path_input)
        storage_layout.addLayout(storage_form)

        storage_buttons = QHBoxLayout()
        storage_buttons.addStretch(1)
        open_storage_btn = QPushButton("Open gegevensmap")
        open_storage_btn.setObjectName("BtnSecondary")
        open_storage_btn.setMinimumHeight(36)
        open_storage_btn.clicked.connect(self.open_data_folder)
        storage_buttons.addWidget(open_storage_btn)
        storage_layout.addLayout(storage_buttons)

        storage_info = QLabel(
            f"Database: {DB_FILENAME}  •  Instellingen: {SETTINGS_FILENAME}\n"
            f"Diagnostisch logboek: {_LOG_PATH or 'niet actief'}"
        )
        storage_info.setObjectName("HintText")
        storage_info.setTextInteractionFlags(Qt.TextSelectableByMouse)
        storage_info.setWordWrap(True)
        storage_layout.addWidget(storage_info)

        if _MIGRATION_NOTICES:
            migration_info = QLabel("\n".join(_MIGRATION_NOTICES))
            migration_info.setObjectName("HintText")
            migration_info.setWordWrap(True)
            storage_layout.addWidget(migration_info)

        outer.addWidget(storage_card)
        outer.addStretch(1)
        return widget

    def open_data_folder(self):
        data_dir = _appdata_dir()
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(data_dir)):
            QMessageBox.warning(
                self,
                "Map openen mislukt",
                f"De gegevensmap kon niet worden geopend.\n\n{data_dir}",
            )

    def save_settings(self):
        email = self.reminder_email_input.text().strip()
        theme = self.style_theme_input.currentText().strip() or "Intersport"

        try:
            settings = load_settings()
            settings["reminder_email"] = email
            settings["style_theme"] = theme
            save_settings(settings)
        except (OSError, TypeError) as exc:
            QMessageBox.critical(
                self,
                "Opslaan mislukt",
                f"De instellingen konden niet worden opgeslagen.\n\n"
                f"Map: {_appdata_dir()}\nReden: {exc}",
            )
            return

        if not self.test_email_input.text().strip():
            self.test_email_input.setText(email)

        self.load_styles()
        QMessageBox.information(self, "Succes", "Instellingen opgeslagen!")

    def handle_send_test_email(self):
        receiver = self.test_email_input.text().strip() or self.reminder_email_input.text().strip()
        if not receiver:
            QMessageBox.warning(self, "Fout", "Vul een test e-mailadres in.")
            return

        self.test_btn.setEnabled(False)
        self.test_btn.setText("Versturen...")

        worker = self._track_email_worker(EmailWorker(send_test_email, receiver))

        def _ok(_):
            self._release_email_worker(worker)
            if self._closing:
                return
            self.test_btn.setEnabled(True)
            self.test_btn.setText("Verstuur test e-mail")
            QMessageBox.information(self, "Gelukt", f"Test e-mail succesvol verzonden naar:\n{receiver}")

        def _err(msg: str):
            self._release_email_worker(worker)
            LOGGER.warning("Test-e-mail mislukt: %s", msg)
            if self._closing:
                return
            self.test_btn.setEnabled(True)
            self.test_btn.setText("Verstuur test e-mail")
            QMessageBox.critical(self, "E-mail fout", f"Test e-mail kon niet worden verzonden.\n\nReden:\n{msg}")

        worker.signals.finished.connect(_ok)
        worker.signals.error.connect(_err)
        self.threadpool.start(worker)

    # =========================================================
    # ADD REQUEST
    # =========================================================
    def handle_add(self):
        klant = self.inputs["klant"].text().strip()
        verkoper = self.inputs["verkoper"].text().strip()
        contact = self.inputs["email"].text().strip()
        opmerking = self.inputs["opmerking"].text().strip()
        maat = self.inputs["maat"].text().strip()
        productcode = self.inputs["product"].text().strip()
        eancode = self.inputs["ean"].text().strip()

        if not klant or not verkoper or not contact or not opmerking or not productcode or not eancode:
            QMessageBox.warning(
                self,
                "Fout",
                "Klantnaam, Verkoper, E-mail of Telefoonnummer, bestelling, Artikelnummer en EAN code zijn verplicht."
            )
            return

        ok, normalized_contact, err = validate_email_or_phone(self.inputs["email"].text())
        if not ok:
            QMessageBox.warning(self, "Fout", err)
            return

        prijs_text = self.inputs["prijs"].text().strip()
        prijs = None
        if prijs_text:
            try:
                prijs = float(prijs_text.replace(",", "."))
                if not math.isfinite(prijs):
                    raise ValueError
            except ValueError:
                QMessageBox.warning(self, "Fout", "Adviesprijs moet een getal zijn (bijv. 123,45).")
                return

        data = (
            klant,
            verkoper,
            normalized_contact,
            opmerking,
            maat,
            self.inputs["filiaal"].currentText(),
            self.inputs["collega"].text().strip(),
            self.inputs["contactstatus"].currentText(),
            self.inputs["betaalstatus"].currentText(),
            self.inputs["bestelstatus"].currentText(),
            productcode,
            eancode,
            prijs,
            _dt_str(datetime.now()),
        )

        try:
            new_row_id = add_request(data)
            LOGGER.info("Bestelling opgeslagen met id %s", new_row_id)
        except (OSError, sqlite3.Error) as exc:
            LOGGER.exception("Bestelling opslaan mislukt")
            QMessageBox.critical(
                self,
                "Bestelling opslaan mislukt",
                f"De bestelling kon niet in de database worden opgeslagen.\n\n"
                f"Database: {_db_path()}\nReden: {exc}",
            )
            return

        try:
            new_row = get_request_by_id(new_row_id)
        except (OSError, sqlite3.Error):
            LOGGER.exception(
                "Opgeslagen bestelling %s kon niet opnieuw worden geladen; bevestigingsmail overgeslagen",
                new_row_id,
            )
            new_row = None

        for _, w in self.inputs.items():
            if isinstance(w, QLineEdit):
                w.clear()
            elif isinstance(w, QComboBox):
                w.setCurrentIndex(0)

        self.refresh_table()
        self.tabs.setCurrentIndex(1)

        if new_row and is_email(normalized_contact):
            worker = self._track_email_worker(EmailWorker(send_customer_confirmation_email, new_row))

            def _cleanup():
                self._release_email_worker(worker)

            def _ok(_):
                _cleanup()
                if self._closing:
                    return
                QMessageBox.information(
                    self,
                    "Succes",
                    f"Opgeslagen!\n\nAutomatische e-mail is verzonden naar:\n{normalized_contact}"
                )

            def _err(msg: str):
                _cleanup()
                LOGGER.warning("Bevestigingsmail mislukt voor bestelling %s: %s", new_row_id, msg)
                if self._closing:
                    return
                QMessageBox.warning(
                    self,
                    "Opgeslagen, maar e-mail mislukt",
                    f"De entry is opgeslagen, maar de automatische e-mail kon niet worden verzonden.\n\nReden:\n{msg}"
                )

            worker.signals.finished.connect(_ok)
            worker.signals.error.connect(_err)
            self.threadpool.start(worker)
        else:
            QMessageBox.information(self, "Succes", "Opgeslagen!")

    # =========================================================
    # REMINDERS
    # =========================================================
    def _get_reminder_candidates(self):
        now = datetime.now()
        cutoff_open_str = _dt_str(now - timedelta(days=REMINDER_AFTER_DAYS))
        repeat_cutoff_str = _dt_str(now - timedelta(days=REMINDER_REPEAT_DAYS))

        with db_connection() as conn:
            return conn.execute(
                f"""
                SELECT {REQUEST_COLUMNS}
                FROM requests
                WHERE afgerond=0
                  AND created_at <= ?
                  AND (
                        reminder_last_sent_at IS NULL
                        OR reminder_last_sent_at = ''
                        OR reminder_last_sent_at <= ?
                  )
                ORDER BY created_at ASC
                """,
                (cutoff_open_str, repeat_cutoff_str),
            ).fetchall()

    def _send_reminder_digest(self, request_rows):
        receiver = get_setting("reminder_email")
        if not receiver:
            raise RuntimeError("Geen reminder e-mail ingesteld in Instellingen.")

        rows = list(request_rows)
        if not rows:
            return 0

        text_entries = []
        html_entries = []
        for index, request_row in enumerate(rows, start=1):
            values = {
                "klant": (request_row["klantnaam"] or "").strip(),
                "verkoper": (request_row["verkoper"] or "").strip(),
                "contact": (request_row["email"] or "").strip(),
                "bestelling": (request_row["opmerking"] or "").strip(),
                "productcode": (request_row["productcode"] or "").strip(),
                "eancode": (request_row["eancode"] or "").strip(),
                "created": (request_row["created_at"] or "").strip(),
            }
            try:
                values["created"] = datetime.fromisoformat(values["created"]).strftime(
                    "%d-%m-%Y %H:%M"
                )
            except (TypeError, ValueError):
                pass

            text_entries.append(
                f"{index}. {values['klant']}\n"
                f"   Verkoper: {values['verkoper']}\n"
                f"   Contact: {values['contact']}\n"
                f"   Bestelling: {values['bestelling']}\n"
                f"   Productcode: {values['productcode']}\n"
                f"   EAN-code: {values['eancode']}\n"
                f"   Aangemaakt: {values['created']}"
            )
            safe = {key: html.escape(value) for key, value in values.items()}
            html_entries.append(
                "<tr>"
                f"<td>{index}</td><td>{safe['klant']}</td><td>{safe['verkoper']}</td>"
                f"<td>{safe['contact']}</td><td>{safe['bestelling']}</td>"
                f"<td>{safe['productcode']}</td><td>{safe['eancode']}</td>"
                f"<td>{safe['created']}</td>"
                "</tr>"
            )

        count = len(rows)
        subject = f"Reminder: {count} openstaande bestelling{'en' if count != 1 else ''}"
        text_body = (
            "Beste collega,\n\n"
            f"Er {'staan' if count != 1 else 'staat'} {count} bestelling"
            f"{'en' if count != 1 else ''} langer dan 7 dagen open.\n\n"
            + "\n\n".join(text_entries)
            + "\n\nWil je deze bestellingen controleren en indien mogelijk afronden?"
        )
        html_rows = "".join(html_entries)
        html_body = f"""
        <html>
        <body style="margin:0; padding:0; background-color:#f4f6f8; font-family:Arial, Helvetica, sans-serif; color:#1f2937;">
            <div style="max-width:1100px; margin:0 auto; padding:32px 20px;">
            <div style="background:#ffffff; border-radius:14px; overflow:hidden; border:1px solid #e5e7eb;">
                <div style="background:#1e3a8a; padding:20px 24px;">
                <h1 style="margin:0; font-size:22px; line-height:1.3; color:#ffffff;">{count} openstaande bestelling{'en' if count != 1 else ''}</h1>
                </div>

                <div style="padding:24px;">
                <p style="margin:0 0 16px 0; font-size:15px;">Beste collega,</p>
                <p style="margin:0 0 20px 0; font-size:15px; line-height:1.6;">
                    Hieronder staan de bestellingen die langer dan 7 dagen openstaan.
                </p>

                <table cellpadding="8" cellspacing="0" border="1" width="100%" style="border-collapse:collapse; margin:0 0 20px 0; border-color:#e5e7eb; font-size:13px;">
                    <thead style="background:#f9fafb;"><tr><th>#</th><th>Klant</th><th>Verkoper</th><th>Contact</th><th>Bestelling</th><th>Productcode</th><th>EAN</th><th>Aangemaakt</th></tr></thead>
                    <tbody>{html_rows}</tbody>
                </table>

                <p style="margin:0 0 18px 0; font-size:15px; line-height:1.6;">
                    Wil je deze bestelling controleren en indien mogelijk afronden?
                </p>
                </div>
            </div>
            </div>
        </body>
        </html>
        """

        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = SENDER_EMAIL
        msg["To"] = receiver
        msg.set_content(text_body)
        msg.add_alternative(html_body, subtype="html")

        _smtp_send_message(msg)
        return count

    def _mark_reminders_sent(self, row_ids):
        sent_at = _dt_str(datetime.now())
        with db_connection() as conn:
            conn.executemany(
                "UPDATE requests SET reminder_last_sent_at=? WHERE id=?",
                ((sent_at, row_id) for row_id in row_ids),
            )

    def _run_reminders(self, auto: bool):
        if self._closing or self._auto_reminders_in_progress or self._manual_reminders_in_progress:
            return

        reminder_receiver = str(get_setting("reminder_email") or "").strip()
        if not reminder_receiver:
            LOGGER.info("Remindercontrole overgeslagen: geen reminderadres ingesteld")
            if not auto:
                QMessageBox.warning(
                    self,
                    "Reminders",
                    "Stel eerst een reminder-e-mailadres in bij Instellingen.",
                )
            return

        cooldown_until = _email_quota_cooldown_until()
        if cooldown_until is not None:
            message = (
                "E-mailverzending is tijdelijk gepauzeerd tot "
                f"{_format_cooldown(cooldown_until)}, omdat de dagelijkse Gmail-limiet "
                "is bereikt. De openstaande bestellingen blijven bewaard."
            )
            LOGGER.info("Remindercontrole overgeslagen: %s", message)
            if not auto:
                QMessageBox.warning(self, "Reminders gepauzeerd", message)
            return

        if auto:
            self._auto_reminders_in_progress = True
        else:
            self._manual_reminders_in_progress = True
            self.manual_send_btn.setEnabled(False)
            self.manual_send_btn.setText("Bezig...")

        try:
            candidates = self._get_reminder_candidates()
        except (OSError, sqlite3.Error) as exc:
            LOGGER.exception("Reminder-kandidaten konden niet worden geladen")
            self._auto_reminders_in_progress = False
            self._manual_reminders_in_progress = False
            if not auto and not self._closing:
                self.manual_send_btn.setEnabled(True)
                self.manual_send_btn.setText("Verstuur reminders nu")
                QMessageBox.critical(
                    self,
                    "Reminders mislukt",
                    f"De database kon niet worden gelezen.\n\n{exc}",
                )
            return

        if not candidates:
            if auto:
                self._auto_reminders_in_progress = False
            else:
                self._manual_reminders_in_progress = False
                self.manual_send_btn.setEnabled(True)
                self.manual_send_btn.setText("Verstuur reminders nu")
                QMessageBox.information(
                    self,
                    "Reminders",
                    "Geen openstaande bestellingen ouder dan 7 dagen die een reminder nodig hebben.",
                )
            return

        row_ids = [row["id"] for row in candidates]
        worker = self._track_email_worker(
            EmailWorker(self._send_reminder_digest, candidates)
        )

        def _restore_controls():
            if auto:
                self._auto_reminders_in_progress = False
            else:
                self._manual_reminders_in_progress = False
                if not self._closing:
                    self.manual_send_btn.setEnabled(True)
                    self.manual_send_btn.setText("Verstuur reminders nu")

        def _ok(sent_count, current_worker=worker):
            self._release_email_worker(current_worker)
            try:
                self._mark_reminders_sent(row_ids)
            except Exception as exc:
                LOGGER.exception("Reminderstatus kon niet worden opgeslagen")
                _restore_controls()
                if not auto and not self._closing:
                    QMessageBox.warning(
                        self,
                        "Reminder verzonden, status niet opgeslagen",
                        "De reminder is verzonden, maar de database kon niet worden bijgewerkt. "
                        f"De reminder kan later nogmaals verschijnen.\n\n{exc}",
                    )
                return
            _restore_controls()
            if not auto and not self._closing:
                QMessageBox.information(
                    self,
                    "Reminders",
                    f"Eén overzichtsmail verzonden voor {sent_count} openstaande "
                    f"bestelling{'en' if sent_count != 1 else ''}.",
                )

        def _err(message: str, current_worker=worker):
            self._release_email_worker(current_worker)
            LOGGER.warning("Reminder-overzicht mislukt: %s", message)
            _restore_controls()
            if not auto and not self._closing:
                QMessageBox.warning(self, "Reminders mislukt", message)

        worker.signals.finished.connect(_ok)
        worker.signals.error.connect(_err)
        self.threadpool.start(worker)

    def run_auto_reminder_check_once(self):
        self._next_auto_reminder_check_at = datetime.now() + timedelta(
            milliseconds=AUTO_CHECK_EVERY_MS
        )
        LOGGER.info(
            "Automatische remindercontrole gestart; volgende controle rond %s",
            _dt_str(self._next_auto_reminder_check_at),
        )
        self._run_reminders(auto=True)

    def _run_reminder_heartbeat(self):
        """Use wall-clock scheduling so a delayed Qt timer catches up after sleep/resume."""
        now = datetime.now()
        if (
            self._next_auto_reminder_check_at is None
            or now >= self._next_auto_reminder_check_at
        ):
            self.run_auto_reminder_check_once()

    def send_all_reminders_manual(self):
        self._run_reminders(auto=False)

    # =========================================================
    # DASHBOARD FUNCTIES
    # =========================================================
    def populate_table(self, table, rows):
        headers = [
            "Klant",
            "Verkoper",
            "E-mail/Telefoon",
            "Opmerking",
            "Maat",
            "Filiaal",
            "Collega",
            "Contact",
            "Betaal",
            "Bestel",
            "Product",
            "EAN",
            "Prijs",
            "Aangemaakt",
            "ID",
        ]

        table.setSortingEnabled(False)
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        self._install_status_delegates(table)

        table.setRowCount(len(rows))

        for r, row in enumerate(rows):
            ordered_row = [
                row["klantnaam"],
                row["verkoper"],
                row["email"],
                row["opmerking"],
                row["maat"],
                row["filiaal"],
                row["collega"],
                row["contactstatus"],
                row["betaalstatus"],
                row["bestelstatus"],
                row["productcode"],
                row["eancode"],
                row["adviesprijs"],
                row["created_at"],
                row["id"],
            ]

            for c, value in enumerate(ordered_row):
                text = "" if value is None else str(value)
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                item.setToolTip(text)
                table.setItem(r, c, item)

        hh = table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)

        widths = {
            0: 140, 1: 120, 2: 220, 3: 320, 4: 90, 5: 160, 6: 160,
            7: 150, 8: 120, 9: 170, 10: 130, 11: 170, 12: 110, 13: 150, 14: 70,
        }
        for col, w in widths.items():
            if col < table.columnCount():
                table.setColumnWidth(col, w)

        table.verticalHeader().setDefaultSectionSize(48)

        table.setSortingEnabled(True)

    def refresh_table(self, _checked=False, raise_errors=False):
        try:
            self.rows = get_all_requests()
        except (OSError, sqlite3.Error) as exc:
            LOGGER.exception("Dashboard kon niet uit de database worden geladen")
            if raise_errors:
                raise
            QMessageBox.critical(
                self,
                "Database lezen mislukt",
                f"Het dashboard kon niet worden vernieuwd.\n\n"
                f"Database: {_db_path()}\nReden: {exc}",
            )
            return False

        rows_open = [r for r in self.rows if r["afgerond"] == 0]
        rows_done = [r for r in self.rows if r["afgerond"] == 1]

        self.populate_table(self.table_open, rows_open)
        self.populate_table(self.table_done, rows_done)

        if hasattr(self, "search") and self.search.text().strip():
            self.filter_tables(self.search.text())
        return True

    def filter_tables(self, text):
        text = text.lower().strip()
        for table in (self.table_open, self.table_done):
            for r in range(table.rowCount()):
                match = False
                if not text:
                    match = True
                else:
                    for c in range(table.columnCount()):
                        item = table.item(r, c)
                        if item and text in item.text().lower():
                            match = True
                            break
                table.setRowHidden(r, not match)

    def _get_selected_row_id_from_table(self, table: QTableWidget, table_row: int) -> int | None:
        id_item = table.item(table_row, table.columnCount() - 1)
        if not id_item:
            return None
        try:
            return int(id_item.text())
        except Exception:
            return None

    def mark_complete(self, row_id):
        row_data = next((r for r in self.rows if r["id"] == row_id), None)
        if row_data is None:
            QMessageBox.warning(self, "Niet gevonden", "Deze bestelling bestaat niet meer.")
            self.refresh_table()
            return
        reply = QMessageBox.question(
            self,
            "Bevestigen",
            "Weet je zeker dat je deze bestelling wilt afronden?\n\n"
            f"Klant: {row_data['klantnaam']}\n"
            f"bestelling: {row_data['opmerking']}",
            QMessageBox.Yes | QMessageBox.No,
        )

        if reply == QMessageBox.Yes:
            try:
                if not set_request_completed(row_id, True):
                    raise LookupError("De bestelling bestaat niet meer.")
            except (OSError, sqlite3.Error, LookupError) as exc:
                LOGGER.exception("Bestelling %s kon niet worden afgerond", row_id)
                QMessageBox.critical(self, "Afronden mislukt", str(exc))
                return
            self.refresh_table()

    # =========================================================
    # CONTEXT MENU + ACTIONS
    # =========================================================
    def send_auto_message_to_customer(self, row_id: int):
        try:
            row = get_request_by_id(row_id)
        except (OSError, sqlite3.Error) as exc:
            LOGGER.exception("Bestelling %s kon niet worden geladen", row_id)
            QMessageBox.critical(self, "Database lezen mislukt", str(exc))
            return
        if not row:
            QMessageBox.warning(self, "Fout", "Deze entry kon niet worden gevonden in de database.")
            return

        contact_value = (row["email"] or "").strip()
        if not contact_value:
            QMessageBox.information(
                self,
                "Geen contactgegevens",
                "Er staat geen e-mailadres of telefoonnummer bij deze entry.\n\nEr kan geen bericht gestuurd worden.",
            )
            return

        if is_email(contact_value):
            worker = self._track_email_worker(EmailWorker(send_customer_arrival_email, row))

            def _cleanup():
                self._release_email_worker(worker)

            def _ok(_):
                _cleanup()
                if self._closing:
                    return
                QMessageBox.information(
                    self,
                    "Bericht verzonden",
                    f"Automatische e-mail is verzonden naar:\n{contact_value}"
                )

            def _err(msg: str):
                _cleanup()
                LOGGER.warning("Aankomstmail mislukt voor bestelling %s: %s", row_id, msg)
                if self._closing:
                    return
                QMessageBox.critical(
                    self,
                    "E-mail fout",
                    f"De e-mail kon niet worden verzonden.\n\nReden:\n{msg}"
                )

            worker.signals.finished.connect(_ok)
            worker.signals.error.connect(_err)
            self.threadpool.start(worker)
            return

        QMessageBox.information(
            self,
            "Telefoonnummer gevonden",
            "Er is een telefoonnummer gevonden bij deze entry.\n\n"
            "WhatsApp-automatisering is nog niet gekoppeld; er is nu geen bericht verstuurd.",
        )

    def show_context_menu(self, table, position):
        row = table.rowAt(position.y())
        if row < 0:
            return

        row_id = self._get_selected_row_id_from_table(table, row)
        if row_id is None:
            return

        menu = QMenu(self)

        auto_msg_action = QAction("Stuur automatisch bericht naar klant", self)
        auto_msg_action.triggered.connect(lambda: self.send_auto_message_to_customer(row_id))
        menu.addAction(auto_msg_action)

        menu.addSeparator()

        if table == self.table_open:
            bewerk_action = QAction("Bewerken", self)
            bewerk_action.triggered.connect(lambda: self.edit_row(table, row))
            menu.addAction(bewerk_action)

            menu.addSeparator()

            afronden_action = QAction("Afronden", self)
            afronden_action.triggered.connect(lambda: self.context_mark_complete_open(row))
            menu.addAction(afronden_action)
        else:
            niet_afgerond_action = QAction("Niet afgerond", self)
            niet_afgerond_action.triggered.connect(lambda: self.context_mark_not_complete(row))
            menu.addAction(niet_afgerond_action)

        menu.exec(table.viewport().mapToGlobal(position))

    def context_mark_complete_open(self, table_row):
        id_item = self.table_open.item(table_row, self.table_open.columnCount() - 1)
        if not id_item:
            return
        self.mark_complete(int(id_item.text()))

    def context_mark_not_complete(self, table_row):
        id_item = self.table_done.item(table_row, self.table_done.columnCount() - 1)
        if not id_item:
            return
        try:
            row_id = int(id_item.text())
            if not set_request_completed(row_id, False):
                raise LookupError("De bestelling bestaat niet meer.")
        except (ValueError, OSError, sqlite3.Error, LookupError) as exc:
            LOGGER.exception("Bestelling kon niet worden heropend")
            QMessageBox.critical(self, "Heropenen mislukt", str(exc))
            return
        self.refresh_table()

    # =========================================================
    # EDIT
    # =========================================================
    def edit_row(self, table, row):
        row_data = [
            table.item(row, c).text() if table.item(row, c) else ""
            for c in range(table.columnCount())
        ]

        dialog = QDialog(self)
        dialog.setWindowTitle("Bestelling bewerken")
        dialog.setObjectName("Dialog")
        layout = QFormLayout(dialog)
        layout.setHorizontalSpacing(14)
        layout.setVerticalSpacing(10)

        inputs = {}

        for key, col_index, label in [
            ("klant", 0, "Klantnaam"),
            ("verkoper", 1, "Verkoper"),
            ("email", 2, "E-mail of Telefoonnummer"),
            ("opmerking", 3, "Bestelling"),
            ("maat", 4, "Maat"),
            ("prijs", 12, "Adviesprijs"),
        ]:
            line = QLineEdit(row_data[col_index])
            if key == "opmerking":
                line.setMaximumWidth(560)
            if key == "email":
                line.setPlaceholderText("bijv. naam@domein.nl, 0612345678, 0413 121233 of +48123456789")
            layout.addRow(label, line)
            inputs[key] = line

        inputs["product"] = QLineEdit(row_data[10])
        inputs["ean"] = QLineEdit(row_data[11])
        inputs["ean"].setPlaceholderText("bijv. 8712345678901")

        product_ean_widget = QWidget()
        product_ean_layout = QHBoxLayout(product_ean_widget)
        product_ean_layout.setContentsMargins(0, 0, 0, 0)
        product_ean_layout.setSpacing(10)
        product_ean_layout.addWidget(inputs["product"])
        product_ean_layout.addWidget(inputs["ean"])
        layout.addRow("Artikelnummer / EAN code", product_ean_widget)

        inputs["filiaal"] = QComboBox()
        inputs["filiaal"].addItem("")
        inputs["filiaal"].addItems(["1) Helmond", "2) Veghel", "3) Venray", "4) Venlo", "5) Breda", "x) Leverancier"])
        index = inputs["filiaal"].findText(row_data[5])
        inputs["filiaal"].setCurrentIndex(index if index >= 0 else 0)
        inputs["collega"] = QLineEdit(row_data[6])

        filiaal_collega_widget = QWidget()
        filiaal_collega_layout = QHBoxLayout(filiaal_collega_widget)
        filiaal_collega_layout.setContentsMargins(0, 0, 0, 0)
        filiaal_collega_layout.setSpacing(10)
        filiaal_collega_layout.addWidget(inputs["filiaal"])
        filiaal_collega_layout.addWidget(inputs["collega"])
        layout.addRow("Aangesproken filiaal / collega", filiaal_collega_widget)

        inputs["contactstatus"] = QComboBox()
        inputs["contactstatus"].addItems(["s.v.p. bellen", "Gebeld", "Bericht gestuurd", "niet bereikbaar/nog bellen"])
        idx = inputs["contactstatus"].findText(row_data[7])
        inputs["contactstatus"].setCurrentIndex(idx if idx >= 0 else 0)
        layout.addRow("Contactstatus", inputs["contactstatus"])
        apply_combo_colors(inputs["contactstatus"], CONTACTSTATUS_COLORS)

        inputs["betaalstatus"] = QComboBox()
        inputs["betaalstatus"].addItems(["Niet betaald!", "Betaald", "Op factuur"])
        idx = inputs["betaalstatus"].findText(row_data[8])
        inputs["betaalstatus"].setCurrentIndex(idx if idx >= 0 else 0)
        layout.addRow("Betaalstatus", inputs["betaalstatus"])
        apply_combo_colors(inputs["betaalstatus"], BETAALSTATUS_COLORS)

        inputs["bestelstatus"] = QComboBox()
        inputs["bestelstatus"].addItems(["Nog opvragen/bestellen", "Onderweg/Besteld", "Op locatie", "Opgehaald/Afgewikkeld"])
        idx = inputs["bestelstatus"].findText(row_data[9])
        inputs["bestelstatus"].setCurrentIndex(idx if idx >= 0 else 0)
        layout.addRow("Bestelstatus", inputs["bestelstatus"])
        apply_combo_colors(inputs["bestelstatus"], BESTELSTATUS_COLORS)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Save).setText("Opslaan")
        buttons.button(QDialogButtonBox.Cancel).setText("Annuleren")
        buttons.setObjectName("DialogButtons")
        layout.addRow(buttons)

        def save_changes():
            if not inputs["klant"].text().strip():
                QMessageBox.warning(dialog, "Fout", "Klantnaam is verplicht.")
                return
            if not inputs["verkoper"].text().strip():
                QMessageBox.warning(dialog, "Fout", "Verkoper is verplicht.")
                return
            if not inputs["email"].text().strip():
                QMessageBox.warning(dialog, "Fout", "E-mail of Telefoonnummer is verplicht.")
                return
            if not inputs["opmerking"].text().strip():
                QMessageBox.warning(dialog, "Fout", "bestelling is verplicht.")
                return
            if not inputs["product"].text().strip():
                QMessageBox.warning(dialog, "Fout", "Artikelnummer is verplicht.")
                return
            if not inputs["ean"].text().strip():
                QMessageBox.warning(dialog, "Fout", "EAN code is verplicht.")
                return

            ok, normalized_contact, err = validate_email_or_phone(inputs["email"].text())
            if not ok:
                QMessageBox.warning(dialog, "Fout", err)
                return

            prijs_text = inputs["prijs"].text().strip()
            if prijs_text:
                try:
                    prijs = float(prijs_text.replace(",", "."))
                    if not math.isfinite(prijs):
                        raise ValueError
                except ValueError:
                    QMessageBox.warning(dialog, "Fout", "Adviesprijs moet een getal zijn (bijv. 123,45).")
                    return
            else:
                prijs = None

            data = (
                inputs["klant"].text().strip(),
                inputs["verkoper"].text().strip(),
                normalized_contact,
                inputs["opmerking"].text().strip(),
                inputs["maat"].text().strip(),
                inputs["filiaal"].currentText(),
                inputs["collega"].text().strip(),
                inputs["contactstatus"].currentText(),
                inputs["betaalstatus"].currentText(),
                inputs["bestelstatus"].currentText(),
                inputs["product"].text().strip(),
                inputs["ean"].text().strip(),
                prijs,
                int(row_data[14]),
            )

            try:
                if not update_request(data):
                    raise LookupError("De bestelling bestaat niet meer.")
            except (OSError, sqlite3.Error, LookupError) as exc:
                LOGGER.exception("Bestelling %s kon niet worden bijgewerkt", row_data[14])
                QMessageBox.critical(
                    dialog,
                    "Wijzigen mislukt",
                    f"De bestelling kon niet worden opgeslagen.\n\n{exc}",
                )
                return

            dialog.accept()
            self.refresh_table()

        buttons.accepted.connect(save_changes)
        buttons.rejected.connect(dialog.reject)
        dialog.exec()

    # =========================================================
    # STYLING
    # =========================================================
    def load_styles(self):
        self.update_branding()
        qss_path = resource_path(get_active_qss_file())
        if os.path.exists(qss_path):
            try:
                with open(qss_path, "r", encoding="utf-8") as f:
                    self.setStyleSheet(f.read())
                return
            except Exception as e:
                LOGGER.warning("Opmaakbestand kon niet worden geladen (%s): %s", qss_path, e)

        fallback_qss_path = resource_path(INTERSPORT_QSS_FILE)
        if os.path.exists(fallback_qss_path):
            try:
                with open(fallback_qss_path, "r", encoding="utf-8") as f:
                    self.setStyleSheet(f.read())
                return
            except Exception as e:
                LOGGER.warning(
                    "Terugval-opmaakbestand kon niet worden geladen (%s): %s",
                    fallback_qss_path,
                    e,
                )

        # fallback
        self.setStyleSheet(
            """
            QWidget { font-family: "Segoe UI", Arial; font-size: 13px; background: #F5F7FB; color: #0F172A; }
            QLabel { background: transparent; }
            QPushButton { background:#164196; color:white; padding:10px 14px; border-radius:10px; font-weight:600; }
            QPushButton:hover { background:#082D78; }
            QLineEdit, QComboBox { background:white; padding:10px; border-radius:10px; border:1px solid #D7DEEA; }
            QTableWidget { background:white; border:1px solid #E6EAF2; border-radius:12px; }
            QHeaderView::section { background:#082D78; color:white; padding:10px; border:none; font-weight:600; }
            """
        )


# ------------------------
# RUN
# ------------------------
if __name__ == "__main__":
    global_instance_lock = None
    diagnostics_error = None
    try:
        _configure_diagnostics()
    except Exception as exc:
        diagnostics_error = exc

    qInstallMessageHandler(_qt_message_handler)
    sys.excepthook = _unhandled_exception_hook

    try:
        app = QApplication(sys.argv)
        app.setApplicationName(APP_NAME)
    except Exception as exc:
        _report_fatal_error("Qt kon niet worden gestart", exc)
        sys.exit(1)

    if diagnostics_error is not None:
        QMessageBox.warning(
            None,
            "Diagnostiek niet beschikbaar",
            f"Het diagnostische logboek kon niet worden aangemaakt.\n\n{diagnostics_error}",
        )

    LOGGER.info("IntersportApp %s wordt gestart", APP_VERSION)
    LOGGER.info("MSIX package family: %s", _windows_package_family_name() or "niet verpakt")
    LOGGER.info("Beoogde gegevensmap: %s", _canonical_appdata_dir())

    try:
        lock_directory = _canonical_appdata_dir()
        lock_directory.mkdir(parents=True, exist_ok=True)
        global_instance_lock = QLockFile(str(lock_directory / "IntersportApp.lock"))
        global_instance_lock.setStaleLockTime(30_000)
        if not global_instance_lock.tryLock(250):
            LOGGER.warning("Tweede applicatie-instantie geblokkeerd")
            QMessageBox.information(
                None,
                "IntersportApp is al geopend",
                "Er draait al een exemplaar van IntersportApp voor deze gebruiker.\n\n"
                "Sluit het andere venster voordat je de app opnieuw opent.",
            )
            sys.exit(0)
    except (OSError, RuntimeError) as exc:
        _report_fatal_error("Applicatievergrendeling aanmaken mislukt", exc)
        sys.exit(1)

    try:
        init_db()
        LOGGER.info("Database succesvol geïnitialiseerd: %s", _db_path())
    except Exception as exc:
        _report_fatal_error("Database initialiseren mislukt", exc)
        sys.exit(1)

    try:
        window = MainWindow()
        window.show()
        LOGGER.info("Hoofdvenster succesvol geopend")
    except Exception as exc:
        _report_fatal_error("Hoofdvenster openen mislukt", exc)
        sys.exit(1)

    exit_code = app.exec()
    LOGGER.info("Applicatie afgesloten met code %s", exit_code)
    sys.exit(exit_code)
