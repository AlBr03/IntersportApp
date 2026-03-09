import sys
import os
import sqlite3
import smtplib
import json
import re
import html
from email.message import EmailMessage
from datetime import datetime, timedelta

from PySide6.QtCore import (
    Qt,
    QTimer,
    QSize,
    QObject,
    Signal,
    QRunnable,
    QThreadPool,
    QStandardPaths,
)
from PySide6.QtGui import QAction, QColor, QPixmap
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
DB_FILENAME = "store.db"
SETTINGS_FILENAME = "settings.json"
INTERSPORT_QSS_FILE = "intersport.qss"
SPORTSTORE_QSS_FILE = "Sportstore.qss"
INTERSPORT_LOGO_FILE = "KOBALT ZONDER PLAATSNAAM.png"
SPORTSTORE_LOGO_FILE = "Sportstore Kobalt blauw.png"

REMINDER_AFTER_DAYS = 7
REMINDER_REPEAT_DAYS = 7
AUTO_CHECK_EVERY_MS = 60 * 60 * 1000  # 1 hour


def _appdata_dir() -> str:
    """
    Writable dir (works under MSIX too).
    """
    base = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
    if not base:
        base = os.path.join(os.getenv("LOCALAPPDATA", os.getcwd()), APP_NAME)
    os.makedirs(base, exist_ok=True)
    return base


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
        "• leeg\n"
        "• e-mail (bijv. naam@domein.nl)\n"
        "• 06XXXXXXXX (10 cijfers)\n"
        "• +316XXXXXXXX (bijv. +31612345678)"
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
    "id, klantnaam, verkoper, email, opmerking, filiaal, collega, contactstatus, betaalstatus, bestelstatus, "
    "productcode, eancode, adviesprijs, afgerond, created_at, reminder_last_sent_at"
)


def connect():
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    return conn


def _column_exists(conn, table_name: str, column_name: str) -> bool:
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    return any((r["name"] if isinstance(r, sqlite3.Row) else r[1]) == column_name for r in rows)


def init_db():
    with connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                klantnaam TEXT NOT NULL,
                verkoper TEXT NOT NULL,
                email TEXT NOT NULL,
                opmerking TEXT NOT NULL,
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

        if not _column_exists(conn, "requests", "eancode"):
            conn.execute("ALTER TABLE requests ADD COLUMN eancode TEXT")

        if not _column_exists(conn, "requests", "collega"):
            conn.execute("ALTER TABLE requests ADD COLUMN collega TEXT")

        if not _column_exists(conn, "requests", "reminder_last_sent_at"):
            conn.execute("ALTER TABLE requests ADD COLUMN reminder_last_sent_at TEXT")

        conn.execute("CREATE INDEX IF NOT EXISTS idx_requests_open_created ON requests(afgerond, created_at)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_requests_open_lastsent ON requests(afgerond, reminder_last_sent_at)")


def add_request(data):
    with connect() as conn:
        cursor = conn.execute(
            """
            INSERT INTO requests (
                klantnaam, verkoper, email, opmerking,
                filiaal, collega, contactstatus, betaalstatus,
                bestelstatus, productcode, eancode, adviesprijs,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            data,
        )
        return cursor.lastrowid


def get_all_requests():
    with connect() as conn:
        open_rows = conn.execute(
            f"SELECT {REQUEST_COLUMNS} FROM requests WHERE afgerond=0 ORDER BY id DESC"
        ).fetchall()
        done_rows = conn.execute(
            f"SELECT {REQUEST_COLUMNS} FROM requests WHERE afgerond=1 ORDER BY id DESC LIMIT 25"
        ).fetchall()
        return list(open_rows) + list(done_rows)


def get_request_by_id(row_id: int):
    with connect() as conn:
        return conn.execute(
            f"SELECT {REQUEST_COLUMNS} FROM requests WHERE id=?",
            (row_id,),
        ).fetchone()


# ------------------------
# SETTINGS (JSON)
# ------------------------
def load_settings():
    path = _settings_path()
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_settings(data):
    path = _settings_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


def get_setting(key):
    return load_settings().get(key, "")


def set_setting(key, value):
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
            self.signals.finished.emit(result)
        except Exception as e:
            self.signals.error.emit(str(e))


def _require_email_password():
    if not APP_PASSWORD:
        raise RuntimeError(
            "EMAIL_PASS environment variable ontbreekt.\n\n"
            "Zet deze op je Gmail app-wachtwoord om e-mail te kunnen versturen.\n"
            "Voorbeeld:\n"
            "  Windows (PowerShell):  $env:EMAIL_PASS='xxxx xxxx xxxx xxxx'\n"
        )


def send_test_email(receiver):
    _require_email_password()

    msg = EmailMessage()
    msg["Subject"] = "Test Email - Intersport Reminder App"
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver
    msg.set_content("Dit is een testmail vanuit de Intersport Reminder App.")

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)


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
        "Het artikel dat u bij ons in de winkel besteld hebt, is zojuist bij ons aangekomen.\n"
        "U kunt het op elk gewenst moment komen ophalen bij ons in de winkel.\n\n"
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
            Het artikel dat u bij ons in de winkel besteld hebt, is zojuist bij ons aangekomen.
          </p>

          <p style="margin:0 0 20px 0; font-size:15px; line-height:1.7;">
            U kunt het op elk gewenst moment komen ophalen bij ons in de winkel.
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

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)

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
        "Wij zijn erg dankbaar dat u hebt gekozen voor het winkelen bij uw lokale ondernemer. "
        "Wij gaan direct aan de slag om uw bestelling zo spoedig en goed mogelijk af te ronden. "
        "Wij laten u direct iets weten zodra uw bestelling bij ons binnen is!\n\n"
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
            Wij zijn erg dankbaar dat u hebt gekozen voor het winkelen bij uw lokale ondernemer.
          </p>

          <p style="margin:0 0 14px 0; font-size:15px; line-height:1.7;">
            Wij gaan direct aan de slag om uw bestelling zo spoedig en goed mogelijk af te ronden.
          </p>

          <p style="margin:0 0 20px 0; font-size:15px; line-height:1.7;">
            Wij laten u direct iets weten zodra uw bestelling bij ons binnen is!
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

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)


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

        self.threadpool = QThreadPool.globalInstance()
        self._active_email_workers = []
        self._manual_reminders_in_progress = False
        self._auto_reminders_in_progress = False

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
        self.refresh_table()

        self.run_auto_reminder_check_once()
        self.reminder_timer = QTimer(self)
        self.reminder_timer.setInterval(AUTO_CHECK_EVERY_MS)
        self.reminder_timer.timeout.connect(self.run_auto_reminder_check_once)
        self.reminder_timer.start()

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
        self.inputs["product"] = QLineEdit()
        self.inputs["ean"] = QLineEdit()
        self.inputs["prijs"] = QLineEdit()

        self.inputs["opmerking"].setPlaceholderText("Bijv. maat / kleur / levertijd / actie...")
        self.inputs["email"].setPlaceholderText("bijv. naam@domein.nl, 0612345678 of +48123456789")
        self.inputs["prijs"].setPlaceholderText("bijv. 129,95")
        self.inputs["ean"].setPlaceholderText("bijv. 8712345678901")

        form.addRow("Klantnaam *", self.inputs["klant"])
        form.addRow("Verkoper *", self.inputs["verkoper"])
        form.addRow("E-mail of Telefoonnummer *", self.inputs["email"])
        form.addRow("bestelling *", self.inputs["opmerking"])

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
        self.inputs["filiaal"].addItems(["1) Helmond", "2) Veghel", "3) Venray", "4) Venlo", "5) Breda"])
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
        self.inputs["bestelstatus"].addItems(["Nog opvragen/bestellen", "Onderweg/Besteld", "Op locatie"])
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
        self.search.setPlaceholderText("Zoeken (klant, product, EAN, bestelling, filiaal, verkoper...)")
        self.search.textChanged.connect(self.filter_tables)

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

        credit = QLabel("© 2026 Alexander Brinkman. Alle rechten voorbehouden. V1.0.2.0")
        credit.setObjectName("FooterCredit")
        credit.setAlignment(Qt.AlignRight)
        lay.addWidget(credit)

        outer.addWidget(card)
        outer.addStretch(1)
        return widget

    def save_settings(self):
        email = self.reminder_email_input.text().strip()
        theme = self.style_theme_input.currentText().strip() or "Intersport"

        set_setting("reminder_email", email)
        set_setting("style_theme", theme)

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

        worker = EmailWorker(send_test_email, receiver)

        def _ok(_):
            self.test_btn.setEnabled(True)
            self.test_btn.setText("Verstuur test e-mail")
            QMessageBox.information(self, "Gelukt", f"Test e-mail succesvol verzonden naar:\n{receiver}")

        def _err(msg: str):
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
            except ValueError:
                QMessageBox.warning(self, "Fout", "Adviesprijs moet een getal zijn (bijv. 123,45).")
                return

        data = (
            klant,
            verkoper,
            normalized_contact,
            opmerking,
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

        new_row_id = add_request(data)
        new_row = get_request_by_id(new_row_id)

        for _, w in self.inputs.items():
            if isinstance(w, QLineEdit):
                w.clear()
            elif isinstance(w, QComboBox):
                w.setCurrentIndex(0)

        self.refresh_table()
        self.tabs.setCurrentIndex(1)

        if new_row and is_email(normalized_contact):
            worker = EmailWorker(send_customer_confirmation_email, new_row)
            self._active_email_workers.append(worker)

            def _cleanup():
                try:
                    self._active_email_workers.remove(worker)
                except ValueError:
                    pass

            def _ok(_):
                _cleanup()
                QMessageBox.information(
                    self,
                    "Succes",
                    f"Opgeslagen!\n\nAutomatische e-mail is verzonden naar:\n{normalized_contact}"
                )

            def _err(msg: str):
                _cleanup()
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

        with connect() as conn:
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

    def _send_reminder_email_for_row(self, request_row):
        receiver = get_setting("reminder_email")
        if not receiver:
            raise RuntimeError("Geen reminder e-mail ingesteld in Instellingen.")

        klant = (request_row["klantnaam"] or "").strip()
        verkoper = (request_row["verkoper"] or "").strip()
        contact = (request_row["email"] or "").strip()
        bestelling = (request_row["opmerking"] or "").strip()
        productcode = (request_row["productcode"] or "").strip()
        eancode = (request_row["eancode"] or "").strip()
        created_at = (request_row["created_at"] or "").strip()
        row_id = request_row["id"]

        # Mooier weergeven van datum/tijd als dat lukt
        created_display = created_at
        try:
            created_display = datetime.fromisoformat(created_at).strftime("%d-%m-%Y %H:%M")
        except Exception:
            pass

        subject = f"Reminder: bestelling van {klant} staat nog open"

        text_body = (
            "Beste collega,\n\n"
            "Er staat een bestelling/verzoek langer dan 7 dagen open.\n\n"
            "Overzicht van de bestelling:\n"
            f"- Klantnaam: {klant}\n"
            f"- Verkoper: {verkoper}\n"
            f"- E-mail of telefoonnummer: {contact}\n"
            f"- Bestelling / opmerking: {bestelling}\n"
            f"- Productcode: {productcode}\n"
            f"- EAN-code: {eancode}\n"
            f"- Aangemaakt op: {created_display}\n\n"
            "Wil je deze bestelling controleren en indien mogelijk afronden?"
        )

        html_body = f"""
        <html>
        <body style="margin:0; padding:0; background-color:#f4f6f8; font-family:Arial, Helvetica, sans-serif; color:#1f2937;">
            <div style="max-width:720px; margin:0 auto; padding:32px 20px;">
            <div style="background:#ffffff; border-radius:14px; overflow:hidden; border:1px solid #e5e7eb;">
                <div style="background:#1e3a8a; padding:20px 24px;">
                <h1 style="margin:0; font-size:22px; line-height:1.3; color:#ffffff;">Reminder openstaande bestelling</h1>
                </div>

                <div style="padding:24px;">
                <p style="margin:0 0 16px 0; font-size:15px;">Beste collega,</p>
                <p style="margin:0 0 20px 0; font-size:15px; line-height:1.6;">
                    Er staat een bestelling/verzoek langer dan 7 dagen open. Hieronder vind je een overzicht van alle verplichte velden van deze entry.
                </p>

                <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 20px 0;">
                    <tr>
                    <td style="padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; font-weight:bold;">Klantnaam</td>
                    <td style="padding:12px 14px; background:#ffffff; border:1px solid #e5e7eb;">{klant}</td>
                    </tr>
                    <tr>
                    <td style="padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; font-weight:bold;">Verkoper</td>
                    <td style="padding:12px 14px; background:#ffffff; border:1px solid #e5e7eb;">{verkoper}</td>
                    </tr>
                    <tr>
                    <td style="padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; font-weight:bold;">E-mail of telefoonnummer</td>
                    <td style="padding:12px 14px; background:#ffffff; border:1px solid #e5e7eb;">{contact}</td>
                    </tr>
                    <tr>
                    <td style="padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; font-weight:bold;">Bestelling / opmerking</td>
                    <td style="padding:12px 14px; background:#ffffff; border:1px solid #e5e7eb;">{bestelling}</td>
                    </tr>
                    <tr>
                    <td style="padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; font-weight:bold;">Productcode</td>
                    <td style="padding:12px 14px; background:#ffffff; border:1px solid #e5e7eb;">{productcode}</td>
                    </tr>
                    <tr>
                    <td style="padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; font-weight:bold;">EAN-code</td>
                    <td style="padding:12px 14px; background:#ffffff; border:1px solid #e5e7eb;">{eancode}</td>
                    </tr>
                    <tr>
                    <td style="padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; font-weight:bold;">Aangemaakt op</td>
                    <td style="padding:12px 14px; background:#ffffff; border:1px solid #e5e7eb;">{created_display}</td>
                    </tr>
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

        _require_email_password()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(msg)

    def _mark_reminder_sent(self, row_id: int):
        with connect() as conn:
            conn.execute(
                "UPDATE requests SET reminder_last_sent_at=? WHERE id=?",
                (_dt_str(datetime.now()), row_id),
            )

    def _run_reminders(self, auto: bool):
        if auto:
            if self._auto_reminders_in_progress:
                return
            self._auto_reminders_in_progress = True
        else:
            if self._manual_reminders_in_progress:
                return
            self._manual_reminders_in_progress = True
            self.manual_send_btn.setEnabled(False)
            self.manual_send_btn.setText("Bezig...")

        candidates = self._get_reminder_candidates()

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

        state = {"auto": auto, "pending": len(candidates), "sent": 0, "failed": 0}

        def _finalize():
            if state["auto"]:
                self._auto_reminders_in_progress = False
            else:
                self._manual_reminders_in_progress = False
                self.manual_send_btn.setEnabled(True)
                self.manual_send_btn.setText("Verstuur reminders nu")

                if state["failed"] == 0:
                    QMessageBox.information(self, "Reminders", f"Reminders verzonden: {state['sent']}")
                else:
                    QMessageBox.warning(
                        self,
                        "Reminders",
                        f"Reminders verzonden: {state['sent']}\nMislukt: {state['failed']}",
                    )

        for row in candidates:
            row_id = row["id"]
            worker = EmailWorker(self._send_reminder_email_for_row, row)

            def _ok(_result=None, rid=row_id):
                try:
                    self._mark_reminder_sent(rid)
                    state["sent"] += 1
                except Exception:
                    state["failed"] += 1
                finally:
                    state["pending"] -= 1
                    if state["pending"] <= 0:
                        _finalize()

            def _err(_msg: str, rid=row_id):
                state["failed"] += 1
                state["pending"] -= 1
                if state["pending"] <= 0:
                    _finalize()

            worker.signals.finished.connect(_ok)
            worker.signals.error.connect(_err)
            self.threadpool.start(worker)

    def run_auto_reminder_check_once(self):
        self._run_reminders(auto=True)

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
            0: 140, 1: 120, 2: 220, 3: 420, 4: 160, 5: 160,
            6: 150, 7: 120, 8: 160, 9: 130, 10: 170, 11: 110, 12: 150, 13: 70,
        }
        for col, w in widths.items():
            if col < table.columnCount():
                table.setColumnWidth(col, w)

        table.resizeRowsToContents()
        max_row_height = 56
        for row_idx in range(table.rowCount()):
            if table.rowHeight(row_idx) > max_row_height:
                table.setRowHeight(row_idx, max_row_height)

        table.setSortingEnabled(True)

    def refresh_table(self):
        self.rows = get_all_requests()

        rows_open = [r for r in self.rows if r["afgerond"] == 0]
        rows_done = [r for r in self.rows if r["afgerond"] == 1]

        self.populate_table(self.table_open, rows_open)
        self.populate_table(self.table_done, rows_done)

        if hasattr(self, "search") and self.search.text().strip():
            self.filter_tables(self.search.text())

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
        row_data = [r for r in self.rows if r["id"] == row_id][0]
        reply = QMessageBox.question(
            self,
            "Bevestigen",
            "Weet je zeker dat je deze bestelling wilt afronden?\n\n"
            f"Klant: {row_data['klantnaam']}\n"
            f"bestelling: {row_data['opmerking']}",
            QMessageBox.Yes | QMessageBox.No,
        )

        if reply == QMessageBox.Yes:
            with connect() as conn:
                conn.execute("UPDATE requests SET afgerond=1 WHERE id=?", (row_id,))
            self.refresh_table()

    # =========================================================
    # CONTEXT MENU + ACTIONS
    # =========================================================
    def send_auto_message_to_customer(self, row_id: int):
        row = get_request_by_id(row_id)
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
            worker = EmailWorker(send_customer_arrival_email, row)
            self._active_email_workers.append(worker)

            def _cleanup():
                try:
                    self._active_email_workers.remove(worker)
                except ValueError:
                    pass

            def _ok(_):
                _cleanup()
                QMessageBox.information(
                    self,
                    "Bericht verzonden",
                    f"Automatische e-mail is verzonden naar:\n{contact_value}"
                )

            def _err(msg: str):
                _cleanup()
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
        row_id = int(id_item.text())
        with connect() as conn:
            conn.execute("UPDATE requests SET afgerond=0 WHERE id=?", (row_id,))
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
            ("prijs", 11, "Adviesprijs"),
        ]:
            line = QLineEdit(row_data[col_index])
            if key == "opmerking":
                line.setMaximumWidth(560)
            if key == "email":
                line.setPlaceholderText("bijv. naam@domein.nl, 0612345678 of +48123456789")
            layout.addRow(label, line)
            inputs[key] = line

        inputs["product"] = QLineEdit(row_data[9])
        inputs["ean"] = QLineEdit(row_data[10])
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
        inputs["filiaal"].addItems(["1) Helmond", "2) Veghel", "3) Venray", "4) Venlo", "5) Breda"])
        index = inputs["filiaal"].findText(row_data[4])
        inputs["filiaal"].setCurrentIndex(index if index >= 0 else 0)
        inputs["collega"] = QLineEdit(row_data[5])

        filiaal_collega_widget = QWidget()
        filiaal_collega_layout = QHBoxLayout(filiaal_collega_widget)
        filiaal_collega_layout.setContentsMargins(0, 0, 0, 0)
        filiaal_collega_layout.setSpacing(10)
        filiaal_collega_layout.addWidget(inputs["filiaal"])
        filiaal_collega_layout.addWidget(inputs["collega"])
        layout.addRow("Aangesproken filiaal / collega", filiaal_collega_widget)

        inputs["contactstatus"] = QComboBox()
        inputs["contactstatus"].addItems(["s.v.p. bellen", "Gebeld", "Bericht gestuurd", "niet bereikbaar/nog bellen"])
        idx = inputs["contactstatus"].findText(row_data[6])
        inputs["contactstatus"].setCurrentIndex(idx if idx >= 0 else 0)
        layout.addRow("Contactstatus", inputs["contactstatus"])
        apply_combo_colors(inputs["contactstatus"], CONTACTSTATUS_COLORS)

        inputs["betaalstatus"] = QComboBox()
        inputs["betaalstatus"].addItems(["Niet betaald!", "Betaald", "Op factuur"])
        idx = inputs["betaalstatus"].findText(row_data[7])
        inputs["betaalstatus"].setCurrentIndex(idx if idx >= 0 else 0)
        layout.addRow("Betaalstatus", inputs["betaalstatus"])
        apply_combo_colors(inputs["betaalstatus"], BETAALSTATUS_COLORS)

        inputs["bestelstatus"] = QComboBox()
        inputs["bestelstatus"].addItems(["Nog opvragen/bestellen", "Onderweg/Besteld", "Op locatie"])
        idx = inputs["bestelstatus"].findText(row_data[8])
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
                inputs["filiaal"].currentText(),
                inputs["collega"].text().strip(),
                inputs["contactstatus"].currentText(),
                inputs["betaalstatus"].currentText(),
                inputs["bestelstatus"].currentText(),
                inputs["product"].text().strip(),
                inputs["ean"].text().strip(),
                prijs,
                int(row_data[13]),  # ID
            )

            with connect() as conn:
                conn.execute(
                    """UPDATE requests
                       SET klantnaam=?, verkoper=?, email=?, opmerking=?,
                           filiaal=?, collega=?, contactstatus=?, betaalstatus=?, bestelstatus=?,
                           productcode=?, eancode=?, adviesprijs=?
                       WHERE id=?""",
                    data,
                )

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
                print(f"[STYLE] Could not load {qss_path}: {e}")

        fallback_qss_path = resource_path(INTERSPORT_QSS_FILE)
        if os.path.exists(fallback_qss_path):
            try:
                with open(fallback_qss_path, "r", encoding="utf-8") as f:
                    self.setStyleSheet(f.read())
                return
            except Exception as e:
                print(f"[STYLE] Could not load fallback {fallback_qss_path}: {e}")

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
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)

    init_db()

    window = MainWindow()
    window.show()
    sys.exit(app.exec())