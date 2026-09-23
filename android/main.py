"""On-device DCT template validator. Input spreadsheets stay on this device."""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import threading
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from androidstorage4kivy import Chooser, ShareSheet, SharedStorage
from kivy.app import App
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput

import dct_all_template_validator as validator


SUPPORTED = {".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".txt"}
DEFAULT_SETTINGS = dict(validator.DEFAULT_RULE_SETTINGS)
SETTING_LABELS = [
    ("negative_sign_convention_share", "Negative sign threshold (%)", 100),
    ("date_tolerance_days", "Date tolerance (days)", 1),
    ("installment_amount_tolerance_pct", "Installment tolerance (%)", 100),
    ("deposit_amount_tolerance_pct", "Maturity amount tolerance (%)", 100),
    ("rd_total_paid_tolerance_pct", "RD paid total tolerance (%)", 100),
]


def timestamp():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


class DCTMobile(App):
    title = "DCT Validator"

    def build(self):
        self.storage = SharedStorage()
        self.chooser = Chooser(self._picked_shared_files)
        self.selected_inputs = []
        self.latest_report = None
        self.latest_archive = None
        self.latest_record = None
        self.busy = False
        self.settings = self._load_settings()
        validator.RULE_SETTINGS.update(self.settings)

        scroll = ScrollView(do_scroll_x=False, bar_width=dp(4))
        root = BoxLayout(orientation="vertical", padding=(dp(16), dp(14)), spacing=dp(10),
                         size_hint_y=None)
        root.bind(minimum_height=root.setter("height"))
        scroll.add_widget(root)
        root.add_widget(Label(text="DCT Validator", size_hint_y=None, height=dp(48),
                              font_size="22sp", bold=True))
        root.add_widget(Label(
            text="Validate DCT spreadsheets on your phone. Files stay on this device.",
            size_hint_y=None, height=dp(52), halign="left", valign="middle"))
        self.pick_button = Button(text="Choose DCT files", size_hint_y=None, height=dp(52),
                                  on_release=lambda *_: self._choose_files())
        root.add_widget(self.pick_button)
        self.file_label = Label(text="No files selected", size_hint_y=None, height=dp(36),
                                halign="left", valign="middle", font_size="14sp")
        root.add_widget(self.file_label)
        self.run_button = Button(text="Validate selected files", size_hint_y=None, height=dp(52),
                                 disabled=True, on_release=lambda *_: self._start_validation())
        root.add_widget(self.run_button)
        self.settings_button = Button(text="Validation settings", size_hint_y=None, height=dp(46),
                                      on_release=lambda *_: self._settings_popup())
        root.add_widget(self.settings_button)
        self.share_button = Button(text="Share validation report", size_hint_y=None, height=dp(50),
                                   disabled=True, on_release=lambda *_: self._share_report())
        root.add_widget(self.share_button)
        self.review_button = Button(text="Mark report reviewed", size_hint_y=None, height=dp(46),
                                    disabled=True, on_release=lambda *_: self._mark_reviewed())
        root.add_widget(self.review_button)
        self.status = Label(text="Ready", size_hint_y=None, height=dp(44), halign="left", valign="middle")
        root.add_widget(self.status)
        self.log = TextInput(text="", readonly=True, multiline=True, size_hint_y=None,
                             height=dp(190), font_size="13sp", padding=(dp(10), dp(10)))
        root.add_widget(self.log)
        return root

    def _private_root(self):
        path = Path(self.user_data_dir)
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _load_settings(self):
        path = self._private_root() / "dct_validator_settings.json"
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
            values = dict(DEFAULT_SETTINGS)
            for key in values:
                if key in loaded:
                    values[key] = float(loaded[key])
            return values
        except (OSError, ValueError, TypeError):
            return dict(DEFAULT_SETTINGS)

    def _save_settings(self):
        path = self._private_root() / "dct_validator_settings.json"
        path.write_text(json.dumps(self.settings, indent=2), encoding="utf-8")
        validator.RULE_SETTINGS.update(self.settings)

    def _append(self, message):
        Clock.schedule_once(lambda _dt: self._append_ui(message), 0)

    def _append_ui(self, message):
        self.log.text = (self.log.text + "\n" + str(message)).strip()
        lines = self.log.text.splitlines()
        if lines:
            self.log.cursor = (len(lines[-1]), len(lines) - 1)

    def _choose_files(self):
        self._append("Choose one or more DCT spreadsheet files.")
        self.chooser.choose_content("*/*", multiple=True)

    def _picked_shared_files(self, shared_files):
        if not shared_files:
            return
        self.busy = True
        self.status.text = "Copying selected files into app-private storage…"
        self.pick_button.disabled = True
        threading.Thread(target=self._copy_selected, args=(shared_files,), daemon=True).start()

    def _copy_selected(self, shared_files):
        folder = self._private_root() / "selected_files"
        folder.mkdir(parents=True, exist_ok=True)
        picked = []
        try:
            for uri in shared_files:
                private_path = self.storage.copy_from_shared(uri)
                if not private_path:
                    continue
                source = Path(str(private_path))
                if source.suffix.lower() not in SUPPORTED or not source.is_file():
                    continue
                target = folder / source.name
                if target.exists():
                    target = folder / f"{source.stem}_{len(picked)+1}{source.suffix}"
                shutil.copy2(source, target)
                picked.append(target)
        except Exception as exc:
            Clock.schedule_once(lambda _dt, message=str(exc): self._selection_failed(message), 0)
            return
        Clock.schedule_once(lambda _dt: self._selection_ready(picked), 0)

    def _selection_failed(self, message):
        self.busy = False
        self.pick_button.disabled = False
        self.status.text = "Could not read the selected files."
        self._append(message)

    def _selection_ready(self, files):
        self.busy = False
        self.pick_button.disabled = False
        self.selected_inputs = files
        if files:
            self.file_label.text = f"{len(files)} file(s) selected"
            for path in files:
                self._append(path.name)
            self.run_button.disabled = False
            self.status.text = f"{len(files)} file(s) ready to validate."
        else:
            self.file_label.text = "No supported spreadsheet files selected"
            self.run_button.disabled = True
            self.status.text = "Select .xls, .xlsx, .xlsm, .csv, .tsv, or .txt files."

    def _start_validation(self):
        if self.busy or not self.selected_inputs:
            return
        self.busy = True
        self.run_button.disabled = True
        self.pick_button.disabled = True
        self.settings_button.disabled = True
        self.share_button.disabled = True
        self.review_button.disabled = True
        self.status.text = "Validating on device…"
        self._append("Starting DCT validation. Large folders may take a while.")
        threading.Thread(target=self._validate_worker, daemon=True).start()

    def _validate_worker(self):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = self._private_root() / "runs" / stamp
        input_dir = base / "input"
        input_dir.mkdir(parents=True, exist_ok=True)
        inputs = []
        try:
            for source in self.selected_inputs:
                dest = input_dir / source.name
                shutil.copy2(source, dest)
                inputs.append({"name": source.name, "size": dest.stat().st_size, "sha256": sha256(dest)})
            validator.RULE_SETTINGS.update(self.settings)
            report_dir, results, failures = validator.run_folder(input_dir, recursive=False, log=self._append)
            archive = base / f"DCT_Validation_{stamp}.zip"
            with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zf:
                for path in sorted(report_dir.rglob("*")):
                    if path.is_file():
                        zf.write(path, path.relative_to(report_dir.parent))
            outputs = [{"name": str(path.relative_to(report_dir)), "size": path.stat().st_size,
                        "sha256": sha256(path)} for path in sorted(report_dir.rglob("*")) if path.is_file()]
            record = {
                "id": stamp, "type": "validation", "started_at": datetime.fromtimestamp(
                    base.stat().st_ctime, timezone.utc).isoformat(timespec="seconds"),
                "finished_at": timestamp(), "inputs": inputs, "settings": dict(self.settings),
                "report_files": outputs, "failures": [{"name": f.name, "message": f.failure} for f in failures],
                "report_archive": archive.name,
            }
            history_path = self._private_root() / "run_history.json"
            try:
                history = json.loads(history_path.read_text(encoding="utf-8")) if history_path.exists() else []
            except (OSError, ValueError):
                history = []
            history.append(record)
            history_path.write_text(json.dumps(history, indent=2), encoding="utf-8")
            clean = sum(getattr(r, "n_clean", 0) for r in results)
            rejected = sum(getattr(r, "n_rejected", 0) for r in results)
            review = sum(getattr(r, "n_review", 0) for r in results)
            Clock.schedule_once(lambda _dt: self._validation_finished(
                report_dir, archive, record, clean, rejected, review, len(failures)), 0)
        except Exception as exc:
            Clock.schedule_once(lambda _dt, message=str(exc): self._validation_failed(message), 0)

    def _validation_failed(self, message):
        self.busy = False
        self.run_button.disabled = not bool(self.selected_inputs)
        self.pick_button.disabled = False
        self.settings_button.disabled = False
        self.status.text = "Validation could not be completed."
        self._append("Error: " + message)

    def _validation_finished(self, report, archive, record, clean, rejected, review, failures):
        self.busy = False
        self.latest_report = report
        self.latest_archive = archive
        self.latest_record = record
        self.run_button.disabled = False
        self.pick_button.disabled = False
        self.settings_button.disabled = False
        self.share_button.disabled = False
        self.review_button.disabled = False
        self.status.text = "Validation complete. Review the workbook before sharing or upload."
        self._append(f"Clean rows: {clean} | Rejected rows: {rejected} | Review items: {review} | Files unread: {failures}")
        self._append(f"Private report archive ready: {archive.name}")

    def _share_report(self):
        if not self.latest_archive or not self.latest_archive.exists():
            return
        try:
            shared = self.storage.copy_to_shared(str(self.latest_archive))
            if not shared:
                raise RuntimeError("Could not copy the report to Android shared storage.")
            ShareSheet().share_file(shared)
            self._append("Android share menu opened. Choose where to save or send the report ZIP.")
        except Exception as exc:
            self._append("Could not share the report: " + str(exc))

    def _mark_reviewed(self):
        if not self.latest_record:
            return
        self.latest_record["reviewed_at"] = timestamp()
        history_path = self._private_root() / "run_history.json"
        try:
            records = json.loads(history_path.read_text(encoding="utf-8"))
            for record in reversed(records):
                if record.get("id") == self.latest_record["id"]:
                    record["reviewed_at"] = self.latest_record["reviewed_at"]
                    break
            history_path.write_text(json.dumps(records, indent=2), encoding="utf-8")
            self._append("Report review recorded on this device.")
        except (OSError, ValueError) as exc:
            self._append("Could not save review record: " + str(exc))

    def _settings_popup(self):
        form = BoxLayout(orientation="vertical", spacing=dp(8), padding=dp(12), size_hint_y=None)
        form.bind(minimum_height=form.setter("height"))
        fields = {}
        for key, label, scale in SETTING_LABELS:
            row = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(72), spacing=dp(4))
            row.add_widget(Label(text=label, halign="left", size_hint_y=None, height=dp(30)))
            value = self.settings.get(key, DEFAULT_SETTINGS[key]) * scale
            field = TextInput(text=str(value), multiline=False, input_filter="float", size_hint_y=None,
                              height=dp(38), font_size="16sp")
            fields[key] = (field, scale)
            row.add_widget(field)
            form.add_widget(row)
        buttons = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(8))
        popup = Popup(title="Validation settings", size_hint=(0.94, 0.92))
        layout = BoxLayout(orientation="vertical", spacing=dp(8))
        settings_scroll = ScrollView(do_scroll_x=False, bar_width=dp(4))
        settings_scroll.add_widget(form)
        layout.add_widget(settings_scroll)
        def save(*_):
            try:
                values = {}
                for key, (field, scale) in fields.items():
                    raw = float(field.text)
                    limit = 366 if scale == 1 else 100
                    if not 0 <= raw <= limit:
                        raise ValueError(f"{SETTING_LABELS[list(fields).index(key)][1]} must be between 0 and {limit}.")
                    values[key] = raw / scale
            except ValueError as exc:
                self._append("Settings not saved: " + str(exc))
                return
            self.settings = values
            self._save_settings()
            self._append("Settings saved locally for the next validation run.")
            popup.dismiss()
        buttons.add_widget(Button(text="Cancel", on_release=popup.dismiss))
        buttons.add_widget(Button(text="Save", on_release=save))
        layout.add_widget(buttons)
        popup.content = layout
        popup.open()


if __name__ == "__main__":
    DCTMobile().run()
