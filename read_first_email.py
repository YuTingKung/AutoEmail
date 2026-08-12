import sys
def read_latest_email_from_outlook() -> int:
    print("=== 讀取信箱最新一封信件 (Outlook/Exchange) ===")

    try:
        import pythoncom
        import win32com.client
    except Exception:
        print("錯誤: 缺少 pywin32 套件，請先安裝: pip install pywin32")
        return 1

    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)  # 最新到最舊

        first_mail = None
        for item in items:
            if getattr(item, "Class", None) == 43:  # 43 = olMail
                first_mail = item
                break

        if first_mail is None:
            print("收件匣沒有可讀取的郵件")
            return 0

        body = (first_mail.Body or "").strip()
        print("\n=== 最新一封信件 ===")
        print(f"寄件者: {first_mail.SenderName}")
        print(f"時間: {first_mail.ReceivedTime}")
        print(f"主旨: {first_mail.Subject}")
        print("\n--- 內容 ---")
        print(body[:5000] if body else "(空白內容)")
        return 0

    except Exception as exc:
        print(f"發生未預期錯誤: {exc}")
        print("提示: 請確認 Outlook 已安裝，且可正常開啟公司 Exchange 信箱。")
        return 1


def main() -> int:
    return read_latest_email_from_outlook()


if __name__ == "__main__":
    sys.exit(main())
