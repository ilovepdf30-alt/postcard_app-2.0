from dataclasses import dataclass
from urllib.parse import urlencode
import requests
from bs4 import BeautifulSoup

from config import BASE_URL, HEADERS
from utils import EMAIL_RE, norm_str, is_email_like

def fio_for_search_row(row) -> str:
    parts = []
    for col in ["Фамилия", "Имя", "Отчество"]:
        v = norm_str(row.get(col, ""))
        if v:
            parts.append(v)
    return " ".join(parts).strip()

def search_person_url(session: requests.Session, fio: str) -> str | None:
    url = f"{BASE_URL}/search/?{urlencode({'search_text': fio})}"
    try:
        resp = session.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
    except Exception:
        return None

    soup = BeautifulSoup(resp.text, "html.parser")
    container = soup.find("div", id="container")
    if not container:
        return None

    for link in container.find_all("a", href=True):
        tag_div = link.find("div", class_="grey tag")
        if tag_div and "Кто есть кто" in tag_div.get_text(strip=True):
            href = link["href"]
            return BASE_URL + href if href.startswith("/") else href

    for link in container.find_all("a", href=True):
        href = link["href"]
        if "/person/" in href:
            return BASE_URL + href if href.startswith("/") else href

    return None

def parse_person_page(session: requests.Session, url: str) -> tuple[str | None, str | None]:
    try:
        resp = session.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
    except Exception:
        return None, None

    soup = BeautifulSoup(resp.text, "html.parser")

    dob = None
    dob_span = soup.find("span", class_="span-bold", string=lambda t: t and "Дата рождения" in t)
    if dob_span:
        text = dob_span.parent.get_text(" ", strip=True)
        dob = text.replace("Дата рождения:", "").strip()

    email = None
    mail_span = soup.find("span", class_="span-bold", string=lambda t: t and "Электронная почта" in t)
    if mail_span:
        block = mail_span.parent
        a = block.find("a", href=True)
        if a and "mailto:" in a["href"]:
            email = a["href"].replace("mailto:", "").strip()
        else:
            m = EMAIL_RE.search(block.get_text(" ", strip=True))
            if m:
                email = m.group(0)

    if email and not is_email_like(email):
        email = None
    return email, dob

@dataclass
class TatcenterResult:
    email: str = ""
    url: str = ""
    dob: str = ""