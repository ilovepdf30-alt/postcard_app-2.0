import os
import time
import shutil
from datetime import datetime

import pandas as pd
import requests
from docx import Document

from model import DataModel
from utils import norm_str, is_email_like, build_obrashenie, open_path
from tatcenter import fio_for_search_row, search_person_url, parse_person_page
from docx_render import replace_placeholders_docx
from win_word_pdf import word_export_pdf_batch
from win_outlook import outlook_send_mail, outlook_list_accounts
from config import WIN

class AppController:
    """
    Бизнес-логика. UI сюда "делегирует" действия.
    Все методы принимают callbacks для прогресса (чтобы UI был тонким).
    """
    def __init__(self, model: DataModel):
        self.m = model

    # ---- project / file system ----
    def open_result(self):
        if not self.m.state.project_dir:
            raise RuntimeError("Выберите папку проекта.")
        open_path(self.m.result_dir())

    def outlook_accounts(self) -> list[str]:
        return outlook_list_accounts()

    # ---- excel / template ----
    def load_excel(self, path: str):
        df = pd.read_excel(path)
        self.m.df = self.m.ensure_columns(df)
        self.m.state.excel_path = path
        self.m.apply_auto_gender()

    def load_template(self, path: str):
        self.m.state.template_path = path

    def set_project_dir(self, d: str):
        self.m.state.project_dir = d
        self.m.ensure_result_dirs()

    # ---- tatcenter ----
    def tatcenter_fetch(
        self,
        only_indices: list[int] | None,
        progress_cb,
        message_cb,
        pause: float = 1.0,
    ) -> dict:
        if self.m.df is None:
            raise RuntimeError("Сначала загрузите Excel.")

        df = self.m.df

        def needs_tc(row):
            main = norm_str(row.get("E-mail", ""))
            tc = norm_str(row.get("E-mail_Татцентр", ""))
            return (not tc) and (not main or not is_email_like(main))

        if only_indices:
            targets = df.loc[only_indices]
            targets = targets[targets.apply(needs_tc, axis=1)]
            scope_text = "выделенным строкам"
        else:
            targets = df[df.apply(needs_tc, axis=1)]
            scope_text = "всем строкам"

        if targets.empty:
            return {"scope": scope_text, "found": 0, "not_found": 0, "errors": 0, "total": 0}

        session = requests.Session()
        found = not_found = errors = 0
        total = len(targets)

        for n, (idx, row) in enumerate(targets.iterrows(), start=1):
            fio = fio_for_search_row(row)
            message_cb(f"[{n}/{total}] {fio}")
            progress_cb(n, total)

            try:
                url = search_person_url(session, fio)
                time.sleep(pause)
                if not url:
                    not_found += 1
                    continue

                email, dob = parse_person_page(session, url)
                time.sleep(pause)

                if email and is_email_like(email):
                    df.at[idx, "E-mail_Татцентр"] = norm_str(email)
                    df.at[idx, "URL Tatcenter"] = url
                    found += 1
                else:
                    not_found += 1

                if dob:
                    df.at[idx, "Дата рождения (Татцентр)"] = norm_str(dob)

            except Exception:
                errors += 1

        return {"scope": scope_text, "found": found, "not_found": not_found, "errors": errors, "total": total}

    def apply_tatcenter_to_main_email(self) -> int:
        if self.m.df is None:
            return 0
        df = self.m.df

        def can_apply(row):
            main = norm_str(row.get("E-mail", ""))
            tc = norm_str(row.get("E-mail_Татцентр", ""))
            return (not is_email_like(main)) and is_email_like(tc)

        mask = df.apply(can_apply, axis=1)
        cnt = int(mask.sum())
        if cnt:
            df.loc[mask, "E-mail"] = df.loc[mask, "E-mail_Татцентр"]
        return cnt

    # ---- docx/pdf ----
    def generate_docx(self, common_text: str, progress_cb, message_cb):
        if self.m.df is None:
            raise RuntimeError("Сначала загрузите Excel.")
        if not self.m.state.project_dir:
            raise RuntimeError("Выберите папку проекта.")
        if not self.m.state.template_path:
            raise RuntimeError("Выберите шаблон DOCX.")

        df = self.m.df
        docx_dir = self.m.result_dir("DOCX")

        total = len(df)
        for n, (idx, row) in enumerate(df.iterrows(), start=1):
            message_cb(f"[{n}/{total}] {row['Фамилия']} {row['Имя']}")
            progress_cb(n, total)

            doc = Document(self.m.state.template_path)
            mapping = {
                "<<OBRASHENIE>>": build_obrashenie(row["Имя"], row["Отчество"], row["Пол (итог)"]),
                "<<TEXT>>": (common_text or "").rstrip("\n"),
            }
            replace_placeholders_docx(doc, mapping)
            out_path = self.m.docx_path_for_idx(idx)
            doc.save(out_path)

        return docx_dir

    def generate_pdf(self):
        if self.m.df is None:
            raise RuntimeError("Сначала загрузите Excel.")
        if not self.m.state.project_dir:
            raise RuntimeError("Выберите папку проекта.")
        if not WIN:
            raise RuntimeError("Сборка PDF доступна только на Windows (Word + pywin32).")

        df = self.m.df
        docx_dir = self.m.result_dir("DOCX")
        pdf_dir = self.m.result_dir("PDF")

        if not os.path.isdir(docx_dir) or not os.listdir(docx_dir):
            raise RuntimeError("Сначала собери DOCX (кнопка «Собрать DOCX»).")

        docx_paths = []
        pdf_paths = []
        for idx, _row in df.iterrows():
            docx_paths.append(self.m.docx_path_for_idx(idx))
            pdf_paths.append(self.m.pdf_path_for_idx(idx))

        word_export_pdf_batch(docx_paths, pdf_paths)
        return pdf_dir

    def export_pdf_files(self, dest_dir: str) -> dict:
        if self.m.df is None:
            raise RuntimeError("Сначала загрузите Excel.")
        if not self.m.state.project_dir:
            raise RuntimeError("Выберите папку проекта.")

        pdf_dir = self.m.result_dir("PDF")
        if not os.path.isdir(pdf_dir):
            raise RuntimeError("Папка RESULT/PDF не найдена. Сначала собери PDF.")

        copied = missing = errors = 0
        for idx, _row in self.m.df.iterrows():
            pdf_path = self.m.pdf_path_for_idx(idx)
            if not os.path.exists(pdf_path):
                missing += 1
                continue
            try:
                shutil.copy2(pdf_path, os.path.join(dest_dir, os.path.basename(pdf_path)))
                copied += 1
            except Exception:
                errors += 1

        return {"copied": copied, "missing": missing, "errors": errors, "dest": dest_dir}

    # ---- outlook ----
    def send_test_one(self, sender: str, subject: str, idx: int):
        if not WIN:
            raise RuntimeError("Отправка доступна только на Windows (Outlook + pywin32).")
        if self.m.df is None:
            raise RuntimeError("Нет данных.")

        pdf_path = self.m.pdf_path_for_idx(idx)
        if not os.path.exists(pdf_path):
            raise RuntimeError("PDF не найден. Сначала собери PDF.")

        sender = norm_str(sender) or self.m.state.sender_email
        subject = norm_str(subject) or "Поздравление"
        outlook_send_mail(sender, sender, subject, "", pdf_path)
        return sender, os.path.basename(pdf_path)

    def send_mails(self, sender: str, subject: str, only_checked: bool) -> str:
        if not WIN:
            raise RuntimeError("Отправка доступна только на Windows (Outlook + pywin32).")
        if self.m.df is None:
            raise RuntimeError("Нет данных.")
        if not self.m.state.project_dir:
            raise RuntimeError("Выберите папку проекта.")

        sender = norm_str(sender) or self.m.state.sender_email
        subject = norm_str(subject) or "Поздравление"

        report = []
        sent = errors = 0

        for idx, row in self.m.df.iterrows():
            if only_checked and not bool(row.get("Отправлять", True)):
                continue

            to = norm_str(row.get("E-mail", ""))
            if not is_email_like(to):
                continue

            pdf_path = self.m.pdf_path_for_idx(idx)
            if not os.path.exists(pdf_path):
                continue

            try:
                outlook_send_mail(sender, to, subject, "", pdf_path)
                sent += 1
                report.append([to, os.path.basename(pdf_path), "SENT", ""])
            except Exception as e:
                errors += 1
                report.append([to, os.path.basename(pdf_path), "ERROR", str(e)])

        out_csv = self.m.result_dir(f"send_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        pd.DataFrame(report, columns=["To", "PDF", "Status", "Reason"]).to_csv(out_csv, index=False, encoding="utf-8-sig")
        return out_csv