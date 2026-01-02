import os
import re
import platform
import subprocess
from tkinter import messagebox

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", re.UNICODE)
RE_ILLEGAL_FS = re.compile(r'[<>:"/\\|?*\x00-\x1F]')

def norm_str(s) -> str:
    if s is None:
        return ""
    s = str(s).replace("\u00A0", " ")
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("Ё", "Е").replace("ё", "е")
    return s

def sanitize_filename(name: str) -> str:
    name = norm_str(name)
    name = RE_ILLEGAL_FS.sub("", name)
    name = name.strip(" .")
    return name if name else "Без_имени"

def is_email_like(email: str) -> bool:
    email = norm_str(email)
    if not email or " " in email:
        return False
    if email.count("@") != 1:
        return False
    local, dom = email.split("@")
    if not local or not dom or "." not in dom:
        return False
    if any(ch in email for ch in [",", ";", "(", ")", ":", "<", ">"]):
        return False
    return True

def detect_gender_by_patronymic(p: str) -> str:
    p = norm_str(p).lower().rstrip(".")
    if not p:
        return ""
    if p.endswith(("овна", "евна", "ична")):
        return "Жен"
    if p.endswith(("ович", "евич", "ич")):
        return "Муж"
    return ""

def toggle_gender(cur: str) -> str:
    cur = norm_str(cur)
    if cur == "Жен":
        return "Муж"
    return "Жен" if cur == "Муж" else "Муж"

def build_obrashenie(first: str, patr: str, gender: str) -> str:
    first = norm_str(first)
    patr = norm_str(patr)
    prefix = "Уважаемая" if gender == "Жен" else "Уважаемый"
    parts = [prefix, first]
    if patr:
        parts.append(patr)
    return " ".join(parts).strip()

def open_path(path: str):
    path = os.path.abspath(path)
    try:
        if platform.system().lower().startswith("win"):
            os.startfile(path)  # noqa
        elif platform.system() == "Darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception:
        messagebox.showinfo("Путь", path)