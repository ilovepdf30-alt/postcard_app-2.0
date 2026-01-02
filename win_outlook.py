import os
from config import WIN
from utils import norm_str

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None

def outlook_list_accounts() -> list[str]:
    if not WIN or win32com is None:
        return []
    app = win32com.client.Dispatch("Outlook.Application")
    session = app.Session
    accs = []
    try:
        for acc in session.Accounts:
            try:
                accs.append(acc.SmtpAddress)
            except Exception:
                accs.append(str(acc.DisplayName))
    except Exception:
        pass
    return accs

def outlook_send_mail(from_account_smtp: str, to_email: str, subject: str, body: str, attachment_path: str) -> None:
    if not WIN or win32com is None:
        raise RuntimeError("Отправка через Outlook доступна только на Windows (Outlook + pywin32).")

    outlook = win32com.client.Dispatch("Outlook.Application")
    session = outlook.Session

    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body or ""
    mail.To = to_email

    if attachment_path:
        mail.Attachments.Add(os.path.abspath(attachment_path))

    chosen = None
    from_account_smtp = norm_str(from_account_smtp)
    for acc in session.Accounts:
        try:
            smtp = acc.SmtpAddress
        except Exception:
            smtp = str(acc.DisplayName)
        if smtp and from_account_smtp and smtp.lower() == from_account_smtp.lower():
            chosen = acc
            break

    if chosen is not None:
        try:
            mail.SendUsingAccount = chosen
        except Exception:
            pass
    else:
        try:
            if from_account_smtp:
                mail.SentOnBehalfOfName = from_account_smtp
        except Exception:
            pass

    mail.Send()