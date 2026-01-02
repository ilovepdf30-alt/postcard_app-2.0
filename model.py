from dataclasses import dataclass
import os
import pandas as pd

from utils import norm_str, is_email_like, sanitize_filename, detect_gender_by_patronymic

@dataclass
class AppState:
    excel_path: str = ""
    template_path: str = ""
    project_dir: str = ""
    sender_email: str = "Mon.OrgOtdel@tatar.ru"
    subject: str = "Поздравление"

class DataModel:
    """
    Хранит df + правила колонок/статусов + нейминг файлов.
    UI/Controller не должны напрямую "придумывать" колонки.
    """
    def __init__(self):
        self.state = AppState()
        self.df: pd.DataFrame | None = None

    # ---- columns ----
    def ensure_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        df.columns = [norm_str(c) for c in df.columns]
        required = ["Фамилия", "Имя", "Отчество"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Не найдены обязательные колонки: {missing}")

        for col, default in [
            ("E-mail", ""),
            ("Пол (итог)", ""),
            ("Пол (авто)", ""),
            ("E-mail_Татцентр", ""),
            ("URL Tatcenter", ""),
            ("Дата рождения (Татцентр)", ""),
        ]:
            if col not in df.columns:
                df[col] = default

        if "Отправлять" not in df.columns:
            df["Отправлять"] = True

        for c in [
            "Фамилия", "Имя", "Отчество", "E-mail", "Пол (итог)", "Пол (авто)",
            "E-mail_Татцентр", "URL Tatcenter", "Дата рождения (Татцентр)"
        ]:
            df[c] = df[c].apply(norm_str)

        df["Отправлять"] = df["Отправлять"].fillna(True).astype(bool)
        return df

    def apply_auto_gender(self):
        if self.df is None:
            return
        self.df["Пол (авто)"] = self.df["Отчество"].apply(detect_gender_by_patronymic)
        mask = self.df["Пол (итог)"].eq("") & self.df["Пол (авто)"].ne("")
        self.df.loc[mask, "Пол (итог)"] = self.df.loc[mask, "Пол (авто)"]

    # ---- status ----
    def compute_status_row(self, row: pd.Series) -> tuple[bool, bool, str]:
        g_ok = norm_str(row.get("Пол (итог)", "")) in ("Муж", "Жен")
        e_ok = is_email_like(norm_str(row.get("E-mail", "")))
        parts = []
        if not g_ok:
            parts.append("нет пола")
        if not e_ok:
            parts.append("e-mail пуст/битый")
        if not parts:
            return True, True, "ОК"
        return g_ok, e_ok, "Проблема: " + ", ".join(parts)

    # ---- result dirs ----
    def ensure_result_dirs(self):
        if not self.state.project_dir:
            return
        os.makedirs(os.path.join(self.state.project_dir, "RESULT", "DOCX"), exist_ok=True)
        os.makedirs(os.path.join(self.state.project_dir, "RESULT", "PDF"), exist_ok=True)
        os.makedirs(os.path.join(self.state.project_dir, "RESULT", "PREVIEW"), exist_ok=True)

    def result_dir(self, *parts) -> str:
        if not self.state.project_dir:
            raise RuntimeError("Сначала выберите папку проекта.")
        self.ensure_result_dirs()
        return os.path.join(self.state.project_dir, "RESULT", *parts)

    # ---- filename policy ----
    def base_name_for_row(self, row: pd.Series) -> str:
        fam = norm_str(row.get("Фамилия", ""))
        im = norm_str(row.get("Имя", ""))
        ot = norm_str(row.get("Отчество", ""))
        initials = (im[0] + "." if im else "") + (ot[0] + "." if ot else "")
        return sanitize_filename(f"{fam} {initials}".strip())

    def pdf_path_for_idx(self, idx: int) -> str:
        if self.df is None:
            raise RuntimeError("Нет данных.")
        row = self.df.loc[idx]
        base = self.base_name_for_row(row)
        return self.result_dir("PDF", base + ".pdf")

    def docx_path_for_idx(self, idx: int) -> str:
        if self.df is None:
            raise RuntimeError("Нет данных.")
        row = self.df.loc[idx]
        base = self.base_name_for_row(row)
        return self.result_dir("DOCX", base + ".docx")