import json
import os
import smtplib
import sqlite3
import tempfile
import unittest
from pathlib import Path
from email.message import EmailMessage
from datetime import datetime, timedelta
from unittest.mock import MagicMock, patch
from contextlib import closing

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import app


OLD_SCHEMA = """
CREATE TABLE requests (
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
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    reminder_last_sent_at TEXT
)
"""


class StorageMigrationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.roaming = self.root / "Roaming"
        self.local = self.root / "Local"
        self.target = self.root / "Canonical"
        self.roaming.mkdir()
        self.local.mkdir()

        self.environment = patch.dict(
            os.environ,
            {
                "APPDATA": str(self.roaming),
                "LOCALAPPDATA": str(self.local),
                "INTERSPORT_DATA_DIR": str(self.target),
            },
        )
        self.environment.start()
        app._DATA_DIR_READY = False
        app._MIGRATION_NOTICES.clear()

    def tearDown(self):
        self.environment.stop()
        app._DATA_DIR_READY = False
        app._MIGRATION_NOTICES.clear()
        self.temp_dir.cleanup()

    def _create_legacy_database(self, directory: Path, customer="Bestaande klant") -> Path:
        directory.mkdir(parents=True, exist_ok=True)
        database = directory / app.DB_FILENAME
        with closing(sqlite3.connect(database)) as conn:
            with conn:
                conn.execute(OLD_SCHEMA)
                conn.execute(
                    """INSERT INTO requests (
                        klantnaam, verkoper, email, opmerking, productcode, eancode
                    ) VALUES (?, ?, ?, ?, ?, ?)""",
                    (customer, "Verkoper", "0612345678", "Test", "ART-1", "1234567890123"),
                )
        return database

    def _new_order(self, customer="Nieuwe klant"):
        return (
            customer,
            "Verkoper",
            "+31612345678",
            "Nieuwe bestelling",
            "M",
            "1) Helmond",
            "Collega",
            "Gebeld",
            "Betaald",
            "Onderweg/Besteld",
            "ART-2",
            "9876543210123",
            99.95,
            "2026-09-30 12:00:00",
        )

    def test_legacy_database_and_settings_are_migrated(self):
        legacy = self.roaming / app.APP_NAME
        self._create_legacy_database(legacy)
        legacy.mkdir(parents=True, exist_ok=True)
        (legacy / app.SETTINGS_FILENAME).write_text(
            json.dumps({"style_theme": "Sportstore"}), encoding="utf-8"
        )

        app.init_db()

        self.assertEqual(Path(app._appdata_dir()), self.target)
        self.assertTrue((self.target / app.DB_FILENAME).is_file())
        self.assertTrue((self.target / app.SETTINGS_FILENAME).is_file())
        with closing(sqlite3.connect(self.target / app.DB_FILENAME)) as conn:
            columns = {row[1] for row in conn.execute("PRAGMA table_info(requests)")}
            customer = conn.execute("SELECT klantnaam FROM requests").fetchone()[0]
        self.assertIn("maat", columns)
        self.assertEqual(customer, "Bestaande klant")
        self.assertEqual(app.get_setting("style_theme"), "Sportstore")
        self.assertEqual(len(list((self.target / "backups").glob("store-*.db"))), 1)

    def test_invalid_newer_database_falls_back_to_valid_copy(self):
        valid = self._create_legacy_database(self.roaming / app.APP_NAME, "Geldige klant")
        invalid_dir = self.local / "Packages" / f"{app.APP_NAME}_old" / "LocalCache" / "Roaming" / app.APP_NAME
        invalid_dir.mkdir(parents=True)
        invalid = invalid_dir / app.DB_FILENAME
        invalid.write_bytes(b"not a sqlite database")
        newer_time = valid.stat().st_mtime + 10
        os.utime(invalid, (newer_time, newer_time))

        app.init_db()

        with closing(sqlite3.connect(self.target / app.DB_FILENAME)) as conn:
            customer = conn.execute("SELECT klantnaam FROM requests").fetchone()[0]
        self.assertEqual(customer, "Geldige klant")

    def test_new_order_can_be_inserted_after_migration(self):
        self._create_legacy_database(self.roaming / app.APP_NAME)
        app.init_db()

        row_id = app.add_request(self._new_order())

        self.assertIsNotNone(row_id)
        self.assertEqual(app.get_request_by_id(row_id)["maat"], "M")

    def test_update_and_completion_operations_are_committed(self):
        app.init_db()
        row_id = app.add_request(self._new_order())
        updated = list(self._new_order("Gewijzigde klant")[:-1])
        updated[-1] = 109.95
        updated.append(row_id)

        self.assertTrue(app.update_request(tuple(updated)))
        self.assertEqual(app.get_request_by_id(row_id)["klantnaam"], "Gewijzigde klant")
        self.assertTrue(app.set_request_completed(row_id, True))
        self.assertEqual(app.get_request_by_id(row_id)["afgerond"], 1)
        self.assertTrue(app.set_request_completed(row_id, False))
        self.assertEqual(app.get_request_by_id(row_id)["afgerond"], 0)
        self.assertFalse(app.set_request_completed(999999, True))

    def test_locked_database_fails_quickly_without_losing_data(self):
        app.init_db()
        locker = sqlite3.connect(self.target / app.DB_FILENAME)
        try:
            locker.execute("BEGIN IMMEDIATE")
            with patch.object(app, "DB_BUSY_TIMEOUT_MS", 100):
                with self.assertRaises(sqlite3.OperationalError):
                    app.add_request(self._new_order())
        finally:
            locker.rollback()
            locker.close()

        self.assertEqual(len(app.get_all_requests()), 0)

    def test_gmail_daily_limit_starts_cooldown_and_blocks_retries(self):
        app.init_db()
        message = EmailMessage()
        message["From"] = app.SENDER_EMAIL
        message["To"] = "test@example.com"
        message.set_content("test")

        smtp = MagicMock()
        smtp.__enter__.return_value.send_message.side_effect = smtplib.SMTPDataError(
            550, b"5.4.5 Daily user sending limit exceeded"
        )
        with patch.object(app.smtplib, "SMTP_SSL", return_value=smtp):
            with self.assertRaises(app.EmailQuotaExceededError):
                app._smtp_send_message(message)

        self.assertTrue(app.get_setting("email_quota_cooldown_until"))
        with patch.object(app.smtplib, "SMTP_SSL") as smtp_factory:
            with self.assertRaises(app.EmailCooldownError):
                app._smtp_send_message(message)
        smtp_factory.assert_not_called()


class CanonicalPathTests(unittest.TestCase):
    def test_msix_uses_physical_local_state_path(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.dict(os.environ, {"LOCALAPPDATA": temp_dir}, clear=False):
                with patch.object(app, "_windows_package_family_name", return_value="IntersportApp_test"):
                    with patch.dict(os.environ, {"INTERSPORT_DATA_DIR": ""}, clear=False):
                        expected = Path(temp_dir) / "Packages" / "IntersportApp_test" / "LocalState" / app.APP_NAME
                        self.assertEqual(app._canonical_appdata_dir(), expected)


class OrderFormTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt_app = app.QApplication.instance() or app.QApplication([])
        cls.qt_app.setApplicationName(app.APP_NAME)

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.environment = patch.dict(
            os.environ,
            {
                "APPDATA": str(self.root / "Roaming"),
                "LOCALAPPDATA": str(self.root / "Local"),
                "INTERSPORT_DATA_DIR": str(self.root / "Data"),
            },
        )
        self.environment.start()
        app._DATA_DIR_READY = False
        app._MIGRATION_NOTICES.clear()
        app.init_db()
        self.window = app.MainWindow()

    def tearDown(self):
        self.window.close()
        self.window.deleteLater()
        self.qt_app.processEvents()
        self.environment.stop()
        app._DATA_DIR_READY = False
        app._MIGRATION_NOTICES.clear()
        self.temp_dir.cleanup()

    def _fill_valid_order(self):
        self.window.inputs["klant"].setText("Formulierklant")
        self.window.inputs["verkoper"].setText("Verkoper")
        self.window.inputs["email"].setText("0612345678")
        self.window.inputs["opmerking"].setText("Bestelling")
        self.window.inputs["product"].setText("ART-UI")
        self.window.inputs["ean"].setText("1234567890123")

    def test_save_button_path_inserts_and_refreshes_order(self):
        self._fill_valid_order()
        with patch.object(app.QMessageBox, "information") as information:
            self.window.handle_add()

        rows = app.get_all_requests()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["klantnaam"], "Formulierklant")
        information.assert_called_once()

    def test_save_button_reports_database_lock_without_clearing_form(self):
        self._fill_valid_order()
        locker = sqlite3.connect(Path(app._db_path()))
        try:
            locker.execute("BEGIN IMMEDIATE")
            with patch.object(app, "DB_BUSY_TIMEOUT_MS", 100):
                with patch.object(app.QMessageBox, "critical") as critical:
                    self.window.handle_add()
            critical.assert_called_once()
        finally:
            locker.rollback()
            locker.close()

        self.assertEqual(self.window.inputs["klant"].text(), "Formulierklant")
        self.assertEqual(len(app.get_all_requests()), 0)

    def test_reminders_are_combined_into_one_digest(self):
        app.set_setting("reminder_email", "winkel@example.com")
        def order(customer):
            return (
                customer, "Verkoper", "klant@example.com", "Schoenen", "M",
                "1) Helmond", "Collega", "Gebeld", "Betaald",
                "Onderweg/Besteld", "ART-2", "9876543210123", 99.95,
                "2026-09-20 12:00:00",
            )

        app.add_request(order("Klant één"))
        app.add_request(order("Klant twee"))
        rows = app.get_all_requests()

        with patch.object(app, "_smtp_send_message") as send:
            count = self.window._send_reminder_digest(rows)

        self.assertEqual(count, 2)
        send.assert_called_once()
        message = send.call_args.args[0]
        self.assertIn("2 openstaande bestellingen", message["Subject"])
        self.assertIn("Klant één", message.get_body(preferencelist=("plain",)).get_content())

        self.window._mark_reminders_sent([row["id"] for row in rows])
        self.assertTrue(all(row["reminder_last_sent_at"] for row in app.get_all_requests()))

    def test_continuous_reminder_scheduler_catches_up_when_due(self):
        self.assertTrue(self.window.reminder_timer.isActive())
        self.assertEqual(
            self.window.reminder_timer.interval(), app.REMINDER_HEARTBEAT_MS
        )

        with patch.object(self.window, "run_auto_reminder_check_once") as run_check:
            self.window._next_auto_reminder_check_at = datetime.now() + timedelta(minutes=5)
            self.window._run_reminder_heartbeat()
            run_check.assert_not_called()

            self.window._next_auto_reminder_check_at = datetime.now() - timedelta(seconds=1)
            self.window._run_reminder_heartbeat()
            run_check.assert_called_once()

    def test_auto_check_always_schedules_the_next_hour(self):
        before = datetime.now() + timedelta(milliseconds=app.AUTO_CHECK_EVERY_MS - 1000)
        with patch.object(self.window, "_run_reminders") as run_reminders:
            self.window.run_auto_reminder_check_once()

        run_reminders.assert_called_once_with(auto=True)
        self.assertGreater(self.window._next_auto_reminder_check_at, before)


if __name__ == "__main__":
    unittest.main()
