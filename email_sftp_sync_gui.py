import csv
import os
import shutil
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional
import zipfile

import tkinter as tk
from tkinter import messagebox, ttk

try:
    import pythoncom
    import win32timezone
    import win32com.client
except Exception:
    pythoncom = None
    win32timezone = None
    win32com = None

try:
    import paramiko
except Exception:
    paramiko = None

try:
    import pyzipper
except Exception:
    pyzipper = None


def get_runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(os.getenv("APPDATA", str(Path.home()))) / "EmailSFTPSync"
    else:
        base = Path(__file__).resolve().parent
    base.mkdir(parents=True, exist_ok=True)
    return base


RUNTIME_DIR = get_runtime_dir()
STATE_FILE = RUNTIME_DIR / "processed_emails.csv"
LOG_FILE = RUNTIME_DIR / "sync_log.txt"


@dataclass
class SyncConfig:
    sftp_host: str
    sftp_port: int
    sftp_username: str
    sftp_password: str
    sftp_remote_dir: str
    zip_password: str
    frequency_value: int
    frequency_unit: str  # "minute" or "hour"
    subject_keyword: str
    recipient_keyword: str


class EmailSFTPSyncApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Exchange Email to SFTP Sync")
        self.root.geometry("760x620")

        self.running = False
        self.worker_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.processed_ids = self.load_processed_ids()

        self.build_ui()

    def build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=14)
        main.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(main, text="Exchange 信件篩選並上傳 SFTP", font=("Microsoft JhengHei UI", 14, "bold"))
        title.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        # SFTP settings
        ttk.Label(main, text="SFTP Host").grid(row=1, column=0, sticky="w", pady=4)
        self.host_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.host_var, width=44).grid(row=1, column=1, columnspan=3, sticky="we", pady=4)

        ttk.Label(main, text="SFTP Port").grid(row=2, column=0, sticky="w", pady=4)
        self.port_var = tk.StringVar(value="22")
        ttk.Entry(main, textvariable=self.port_var, width=12).grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(main, text="SFTP 帳號").grid(row=3, column=0, sticky="w", pady=4)
        self.user_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.user_var, width=30).grid(row=3, column=1, sticky="w", pady=4)

        ttk.Label(main, text="SFTP 密碼").grid(row=4, column=0, sticky="w", pady=4)
        self.password_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.password_var, width=30, show="*").grid(row=4, column=1, sticky="w", pady=4)

        ttk.Label(main, text="SFTP 目標路徑").grid(row=5, column=0, sticky="w", pady=4)
        self.remote_dir_var = tk.StringVar(value="/upload")
        ttk.Entry(main, textvariable=self.remote_dir_var, width=44).grid(row=5, column=1, columnspan=3, sticky="we", pady=4)

        ttk.Label(main, text="ZIP 解壓密碼").grid(row=6, column=0, sticky="w", pady=4)
        self.zip_password_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.zip_password_var, width=30, show="*").grid(row=6, column=1, sticky="w", pady=4)

        # Filter settings
        sep1 = ttk.Separator(main, orient=tk.HORIZONTAL)
        sep1.grid(row=7, column=0, columnspan=4, sticky="we", pady=12)

        ttk.Label(main, text="篩選條件（可留空）", font=("Microsoft JhengHei UI", 10, "bold")).grid(
            row=8, column=0, columnspan=4, sticky="w", pady=(0, 8)
        )

        ttk.Label(main, text="主旨包含").grid(row=9, column=0, sticky="w", pady=4)
        self.subject_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.subject_var, width=44).grid(row=9, column=1, columnspan=3, sticky="we", pady=4)

        ttk.Label(main, text="收件人包含").grid(row=10, column=0, sticky="w", pady=4)
        self.recipient_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.recipient_var, width=44).grid(row=10, column=1, columnspan=3, sticky="we", pady=4)

        # Frequency settings
        sep2 = ttk.Separator(main, orient=tk.HORIZONTAL)
        sep2.grid(row=11, column=0, columnspan=4, sticky="we", pady=12)

        ttk.Label(main, text="觸發頻率", font=("Microsoft JhengHei UI", 10, "bold")).grid(
            row=12, column=0, columnspan=4, sticky="w", pady=(0, 8)
        )

        ttk.Label(main, text="每").grid(row=13, column=0, sticky="w", pady=4)
        self.frequency_value_var = tk.StringVar(value="1")
        ttk.Entry(main, textvariable=self.frequency_value_var, width=8).grid(row=13, column=1, sticky="w", pady=4)

        self.frequency_unit_var = tk.StringVar(value="minute")
        frequency_unit_combo = ttk.Combobox(
            main,
            textvariable=self.frequency_unit_var,
            values=["minute", "hour"],
            state="readonly",
            width=12,
        )
        frequency_unit_combo.grid(row=13, column=2, sticky="w", pady=4)

        # Buttons and status
        sep3 = ttk.Separator(main, orient=tk.HORIZONTAL)
        sep3.grid(row=14, column=0, columnspan=4, sticky="we", pady=12)

        self.start_btn = ttk.Button(main, text="啟動持續同步", command=self.start_sync)
        self.start_btn.grid(row=15, column=0, pady=4, sticky="w")

        self.stop_btn = ttk.Button(main, text="停止", command=self.stop_sync, state=tk.DISABLED)
        self.stop_btn.grid(row=15, column=1, pady=4, sticky="w")

        self.test_btn = ttk.Button(main, text="測試 SFTP 連線", command=self.test_sftp)
        self.test_btn.grid(row=15, column=2, pady=4, sticky="w")

        self.status_var = tk.StringVar(value="狀態: 待機")
        ttk.Label(main, textvariable=self.status_var, foreground="#1e3a8a").grid(row=16, column=0, columnspan=4, sticky="w", pady=(8, 6))

        ttk.Label(main, text="執行紀錄", font=("Microsoft JhengHei UI", 10, "bold")).grid(row=17, column=0, sticky="w", pady=(8, 4))

        self.log_text = tk.Text(main, height=14, wrap=tk.WORD)
        self.log_text.grid(row=18, column=0, columnspan=4, sticky="nsew")

        scroll = ttk.Scrollbar(main, orient=tk.VERTICAL, command=self.log_text.yview)
        scroll.grid(row=18, column=4, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        main.columnconfigure(1, weight=1)
        main.columnconfigure(3, weight=1)
        main.rowconfigure(18, weight=1)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def log(self, message: str) -> None:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {message}\n"
        self.log_text.insert(tk.END, line)
        self.log_text.see(tk.END)
        self.log_text.update_idletasks()

        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(line)

    def set_status(self, text: str) -> None:
        self.status_var.set(f"狀態: {text}")

    def load_processed_ids(self) -> set[str]:
        ids: set[str] = set()
        if not STATE_FILE.exists():
            return ids

        try:
            with STATE_FILE.open("r", newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                for row in reader:
                    if row:
                        ids.add(row[0])
        except Exception:
            pass
        return ids

    def append_processed_id(self, item_id: str) -> None:
        self.processed_ids.add(item_id)
        with STATE_FILE.open("a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([item_id])

    def parse_config(self) -> SyncConfig:
        host = self.host_var.get().strip()
        port_text = self.port_var.get().strip()
        username = self.user_var.get().strip()
        password = self.password_var.get()
        remote_dir = self.remote_dir_var.get().strip()
        zip_password = self.zip_password_var.get()
        freq_text = self.frequency_value_var.get().strip()
        freq_unit = self.frequency_unit_var.get().strip()
        subject = self.subject_var.get().strip()
        recipient = self.recipient_var.get().strip()

        if not host:
            raise ValueError("請輸入 SFTP Host")
        if not username:
            raise ValueError("請輸入 SFTP 帳號")
        if not password:
            raise ValueError("請輸入 SFTP 密碼")
        if not remote_dir:
            raise ValueError("請輸入 SFTP 目標路徑")

        try:
            port = int(port_text)
        except ValueError:
            raise ValueError("SFTP Port 必須是數字")

        if port <= 0 or port > 65535:
            raise ValueError("SFTP Port 範圍需為 1-65535")

        try:
            frequency_value = int(freq_text)
        except ValueError:
            raise ValueError("觸發頻率數值必須是整數")

        if frequency_value <= 0:
            raise ValueError("觸發頻率數值必須大於 0")

        if freq_unit not in {"minute", "hour"}:
            raise ValueError("觸發單位必須是 minute 或 hour")

        return SyncConfig(
            sftp_host=host,
            sftp_port=port,
            sftp_username=username,
            sftp_password=password,
            sftp_remote_dir=remote_dir,
            zip_password=zip_password,
            frequency_value=frequency_value,
            frequency_unit=freq_unit,
            subject_keyword=subject,
            recipient_keyword=recipient,
        )

    def test_sftp(self) -> None:
        try:
            config = self.parse_config()
        except ValueError as exc:
            messagebox.showerror("設定錯誤", str(exc))
            return

        if paramiko is None:
            messagebox.showerror("缺少套件", "缺少 paramiko，請先安裝: pip install paramiko")
            return

        self.set_status("測試 SFTP 中")
        self.log("測試 SFTP 連線...")

        try:
            transport = paramiko.Transport((config.sftp_host, config.sftp_port))
            transport.connect(username=config.sftp_username, password=config.sftp_password)
            sftp = paramiko.SFTPClient.from_transport(transport)
            sftp.listdir(config.sftp_remote_dir)
            sftp.close()
            transport.close()
            self.log("SFTP 測試成功")
            self.set_status("SFTP 測試成功")
            messagebox.showinfo("成功", "SFTP 連線與目標路徑測試成功")
        except Exception as exc:
            self.log(f"SFTP 測試失敗: {exc}")
            self.set_status("SFTP 測試失敗")
            messagebox.showerror("失敗", f"SFTP 測試失敗:\n{exc}")

    def start_sync(self) -> None:
        try:
            config = self.parse_config()
        except ValueError as exc:
            messagebox.showerror("設定錯誤", str(exc))
            return

        if paramiko is None:
            messagebox.showerror("缺少套件", "缺少 paramiko，請先安裝: pip install paramiko")
            return

        if pyzipper is None:
            messagebox.showerror("缺少套件", "缺少 pyzipper，請重新取得最新版 EXE")
            return

        if pythoncom is None or win32com is None:
            messagebox.showerror("缺少套件", "缺少 pywin32，請先安裝: pip install pywin32")
            return

        self.running = True
        self.stop_event.clear()
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)

        self.set_status("同步中")
        self.log("啟動持續同步")

        self.worker_thread = threading.Thread(target=self.sync_loop, args=(config,), daemon=True)
        self.worker_thread.start()

    def stop_sync(self) -> None:
        if not self.running:
            return
        self.stop_event.set()
        self.running = False
        self.start_btn.configure(state=tk.NORMAL)
        self.stop_btn.configure(state=tk.DISABLED)
        self.set_status("已停止")
        self.log("同步已停止")

    def sync_loop(self, config: SyncConfig) -> None:
        seconds = config.frequency_value * 60
        if config.frequency_unit == "hour":
            seconds = config.frequency_value * 3600

        while not self.stop_event.is_set():
            try:
                self.process_once(config)
            except Exception as exc:
                self.log(f"同步處理發生錯誤: {exc}")

            if self.stop_event.wait(seconds):
                break

    def process_once(self, config: SyncConfig) -> None:
        self.log("開始掃描信箱...")
        mails = self.fetch_matching_emails(config)
        if not mails:
            self.log("沒有符合條件且尚未處理的郵件")
            return

        self.log(f"找到 {len(mails)} 封待上傳郵件")

        transport = paramiko.Transport((config.sftp_host, config.sftp_port))
        transport.connect(username=config.sftp_username, password=config.sftp_password)
        sftp = paramiko.SFTPClient.from_transport(transport)

        try:
            self.ensure_remote_dir(sftp, config.sftp_remote_dir)

            for mail in mails:
                item_id = mail["item_id"]
                safe_subject = mail["safe_subject"]
                temp_dir = RUNTIME_DIR / f"work_{item_id[:20]}_{safe_subject}"
                try:
                    self.export_mail_bundle(mail, config, sftp, temp_dir)
                    self.append_processed_id(item_id)
                    self.log(f"上傳成功: {safe_subject}")
                finally:
                    if temp_dir.exists():
                        shutil.rmtree(temp_dir, ignore_errors=True)

        finally:
            sftp.close()
            transport.close()

    def export_mail_bundle(self, mail: dict, config: SyncConfig, sftp: "paramiko.SFTPClient", temp_dir: Path) -> None:
        temp_dir.mkdir(parents=True, exist_ok=True)

        mail_item = self.get_mail_item(mail["item_id"])
        mail_base_name = mail["base_name"]
        remote_mail_dir = f"{config.sftp_remote_dir.rstrip('/')}/{mail_base_name}"
        self.ensure_remote_dir(sftp, remote_mail_dir)

        msg_file = temp_dir / f"{mail_base_name}.msg"
        mail_item.SaveAs(str(msg_file), 3)
        self.upload_file(sftp, msg_file, f"{remote_mail_dir}/{msg_file.name}")

        attachments = getattr(mail_item, "Attachments", None)
        if attachments is None:
            return

        for index in range(1, attachments.Count + 1):
            attachment = attachments.Item(index)
            attachment_name = self.make_safe_name(str(getattr(attachment, "FileName", "attachment")))
            if not attachment_name.lower().endswith(".zip"):
                continue

            zip_path = temp_dir / attachment_name
            extract_dir = temp_dir / f"extract_{Path(attachment_name).stem}"
            extract_dir.mkdir(parents=True, exist_ok=True)

            attachment.SaveAsFile(str(zip_path))
            extracted = self.extract_zip(zip_path, extract_dir, config.zip_password)
            if not extracted:
                self.log(f"ZIP 無可上傳內容: {attachment_name}")
                continue

            remote_zip_dir = f"{remote_mail_dir}/{Path(attachment_name).stem}"
            self.ensure_remote_dir(sftp, remote_zip_dir)
            for extracted_file in extracted:
                relative_path = extracted_file.relative_to(extract_dir).as_posix()
                remote_path = f"{remote_zip_dir}/{relative_path}"
                self.ensure_remote_dir(sftp, str(Path(remote_path).parent).replace('\\', '/'))
                self.upload_file(sftp, extracted_file, remote_path)

    def upload_file(self, sftp: "paramiko.SFTPClient", local_path: Path, remote_path: str) -> None:
        sftp.put(str(local_path), remote_path)
        self.log(f"上傳成功: {remote_path}")

    def get_mail_item(self, item_id: str):
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return namespace.GetItemFromID(item_id)

    def extract_zip(self, zip_path: Path, extract_dir: Path, password: str) -> List[Path]:
        try:
            with pyzipper.AESZipFile(zip_path) as zip_file:
                if password:
                    zip_file.setpassword(password.encode("utf-8"))
                zip_file.extractall(extract_dir)
        except RuntimeError as exc:
            raise RuntimeError(f"ZIP 解壓失敗，請確認密碼是否正確: {zip_path.name} ({exc})") from exc
        except zipfile.BadZipFile as exc:
            raise RuntimeError(f"ZIP 檔案損毀或格式不支援: {zip_path.name}") from exc

        return [path for path in extract_dir.rglob("*") if path.is_file()]

    def make_safe_name(self, value: str) -> str:
        safe_value = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in value).strip("._")
        return safe_value[:80] or "mail"

    def ensure_remote_dir(self, sftp: "paramiko.SFTPClient", remote_dir: str) -> None:
        parts = [p for p in remote_dir.strip("/").split("/") if p]
        if remote_dir.startswith("/"):
            current = "/"
        else:
            current = "."

        for part in parts:
            if current in ("/", "."):
                next_dir = f"{current.rstrip('/')}/{part}" if current == "/" else part
            else:
                next_dir = f"{current}/{part}"

            try:
                sftp.listdir(next_dir)
            except Exception:
                sftp.mkdir(next_dir)

            current = next_dir

    def fetch_matching_emails(self, config: SyncConfig) -> List[dict]:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)

        result: List[dict] = []
        now = datetime.now().strftime("%Y%m%d_%H%M%S")

        for item in items:
            if getattr(item, "Class", None) != 43:
                continue

            item_id = str(getattr(item, "EntryID", "")).strip()
            if not item_id or item_id in self.processed_ids:
                continue

            subject = str(getattr(item, "Subject", "") or "")
            to_value = str(getattr(item, "To", "") or "")

            if config.subject_keyword and config.subject_keyword.lower() not in subject.lower():
                continue

            if config.recipient_keyword and config.recipient_keyword.lower() not in to_value.lower():
                continue

            safe_subject = self.make_safe_name(subject) or "no_subject"
            base_name = f"mail_{now}_{safe_subject}_{len(result)+1}"

            result.append(
                {
                    "item_id": item_id,
                    "safe_subject": safe_subject,
                    "base_name": base_name,
                }
            )

        return result

    def on_close(self) -> None:
        if self.running:
            if not messagebox.askyesno("確認", "同步仍在執行中，是否要停止並關閉？"):
                return
            self.stop_sync()

        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    app = EmailSFTPSyncApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
