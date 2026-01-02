import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

from config import WIN
from model import DataModel
from controller import AppController
from utils import norm_str, is_email_like, toggle_gender
from preview import render_pdf_page_to_photoimage


PAD = 12


class ProgressDialog(tk.Toplevel):
    def __init__(self, master, title: str):
        super().__init__(master)
        self.title(title)
        self.geometry("520x150")
        self.resizable(False, False)

        self.lbl = ttk.Label(self, text="Старт...", padding=10)
        self.lbl.pack(fill="x")

        self.pbar = ttk.Progressbar(self, length=480, mode="determinate")
        self.pbar.pack(pady=8)

        self.update()

    def set_total(self, total: int):
        self.pbar["maximum"] = max(1, total)
        self.pbar["value"] = 0
        self.update()

    def set_progress(self, n: int, total: int):
        self.pbar["maximum"] = max(1, total)
        self.pbar["value"] = n
        self.update()

    def set_text(self, text: str):
        self.lbl.config(text=text)
        self.update()


class PostcardApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # macOS Retina: чтобы не было мелко/“дешево”
        if not WIN:
            try:
                self.tk.call("tk", "scaling", 1.25)
            except Exception:
                pass

        self.title("Открытки — Excel → Пол/E-mail → DOCX/PDF → Outlook")
        self.geometry("1500x900")
        self.minsize(1180, 720)

        self.model = DataModel()
        self.ctrl = AppController(self.model)

        self.view_idx: list[int] = []
        self._edit_widget = None
        self.pdf_cache_imgtk = None
        self._preview_after_id = None

        self._build_styles()
        self._build_ui()
        self._refresh_everything()
        self.refresh_table()

    # -------------------------
    # Styles
    # -------------------------
    def _build_styles(self):
        style = ttk.Style(self)

        # На macOS стили ttk реально “ложатся” в clam. aqua — почти не кастомится.
        preferred = ["vista", "xpnative", "clam"] if WIN else ["clam", "aqua"]
        for t in preferred:
            try:
                style.theme_use(t)
                break
            except Exception:
                continue

        # Базовые цвета
        self.UI_BG = "#f6f7fb"
        self.CARD_BG = "#ffffff"
        self.TEXT = "#111827"
        self.SUBTLE = "#6b7280"
        self.OK = "#16a34a"
        self.WARN = "#b45309"
        self.BAD = "#dc2626"
        self.BORDER = "#d1d5db"

        # Нормализуем дефолтные шрифты Tk (чтобы не было “разнобоя”)
        try:
            f = tkfont.nametofont("TkDefaultFont")
            f.configure(size=12)
            self.option_add("*Font", f)
        except Exception:
            pass

        default_font = ("Arial", 12)
        title_font = ("Arial", 13, "bold")

        try:
            self.configure(bg=self.UI_BG)
        except Exception:
            pass

        # Фреймы: карточки БЕЗ рамок (рамку дадим только таблице/preview)
        style.configure("TFrame", background=self.UI_BG)
        style.configure("Card.TFrame", background=self.CARD_BG, relief="flat", borderwidth=0)

        style.configure("CardTitle.TLabel", background=self.CARD_BG, foreground=self.TEXT, font=title_font)
        style.configure("CardSub.TLabel", background=self.CARD_BG, foreground=self.SUBTLE, font=("Arial", 11))
        style.configure("TLabel", background=self.UI_BG, foreground=self.TEXT, font=default_font)

        style.configure("StatusOK.TLabel", foreground=self.OK, font=("Arial", 12, "bold"))
        style.configure("StatusBad.TLabel", foreground=self.BAD, font=("Arial", 12, "bold"))
        style.configure("StatusWarn.TLabel", foreground=self.WARN, font=("Arial", 12, "bold"))

        style.configure("TButton", padding=(10, 7), font=default_font)
        style.configure("Big.TButton", padding=(14, 10), font=("Arial", 12, "bold"))
        style.map("Big.TButton", foreground=[("disabled", "#9ca3af")])

        # Чипы фильтра
        style.configure("Chip.TButton", padding=(10, 6), font=("Arial", 11, "bold"))
        style.configure("ChipActive.TButton", padding=(10, 6), font=("Arial", 11, "bold"))
        # В clam background работает через map
        style.map(
            "ChipActive.TButton",
            background=[("!disabled", "#111827")],
            foreground=[("!disabled", "#ffffff")],
        )

        style.configure(
            "Treeview",
            rowheight=32,
            font=("Arial", 12),
            fieldbackground=self.CARD_BG,
            background=self.CARD_BG,
        )
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"), padding=(10, 10))

        try:
            self.option_add("*Canvas.Background", "#eef1f6")
        except Exception:
            pass

    # -------------------------
    # UI layout
    # -------------------------
    def _build_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- TOP
        top = ttk.Frame(self, style="Card.TFrame", padding=14)
        top.grid(row=0, column=0, sticky="ew", padx=PAD, pady=(PAD, 10))
        top.grid_columnconfigure(1, weight=1)

        ttk.Label(top, text="Проект", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(top, text="Excel / шаблон DOCX / папка RESULT", style="CardSub.TLabel").grid(
            row=0, column=1, sticky="w", padx=(10, 0)
        )

        row1 = ttk.Frame(top, style="Card.TFrame")
        row1.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        row1.grid_columnconfigure(1, weight=1)
        row1.grid_columnconfigure(3, weight=1)
        row1.grid_columnconfigure(5, weight=1)

        self.excel_var = tk.StringVar(value="")
        self.template_var = tk.StringVar(value="")
        self.project_var = tk.StringVar(value="")

        ttk.Button(row1, text="📄 Excel", command=self.load_excel).grid(row=0, column=0, sticky="w")
        ttk.Entry(row1, textvariable=self.excel_var, state="readonly").grid(
            row=0, column=1, sticky="ew", padx=(8, 14)
        )

        ttk.Button(row1, text="🧩 Шаблон DOCX", command=self.load_template).grid(row=0, column=2, sticky="w")
        ttk.Entry(row1, textvariable=self.template_var, state="readonly").grid(
            row=0, column=3, sticky="ew", padx=(8, 14)
        )

        ttk.Button(row1, text="📁 Папка проекта", command=self.choose_project_dir).grid(row=0, column=4, sticky="w")
        ttk.Entry(row1, textvariable=self.project_var, state="readonly").grid(
            row=0, column=5, sticky="ew", padx=(8, 0)
        )

        row2 = ttk.Frame(top, style="Card.TFrame")
        row2.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))

        self.st_data = ttk.Label(row2, text="Данные: нет", style="StatusBad.TLabel")
        self.st_data.grid(row=0, column=0, sticky="w", padx=(0, 18))
        self.st_gender = ttk.Label(row2, text="Пол: —", style="StatusWarn.TLabel")
        self.st_gender.grid(row=0, column=1, sticky="w", padx=(0, 18))
        self.st_email = ttk.Label(row2, text="E-mail: —", style="StatusWarn.TLabel")
        self.st_email.grid(row=0, column=2, sticky="w", padx=(0, 18))
        self.st_pdf = ttk.Label(row2, text="PDF: —", style="StatusWarn.TLabel")
        self.st_pdf.grid(row=0, column=3, sticky="w")

        ttk.Separator(self, orient="horizontal").grid(row=0, column=0, sticky="ew", padx=PAD, pady=(0, 0))

        # --- MIDDLE split
        mid = ttk.Frame(self, style="TFrame")
        mid.grid(row=1, column=0, sticky="nsew", padx=PAD)
        mid.grid_rowconfigure(0, weight=1)
        mid.grid_columnconfigure(0, weight=1)

        paned = ttk.Panedwindow(mid, orient="horizontal")
        paned.grid(row=0, column=0, sticky="nsew")

        left = ttk.Frame(paned, style="Card.TFrame", padding=12)
        right = ttk.Frame(paned, style="Card.TFrame", padding=12)
        paned.add(left, weight=3)
        paned.add(right, weight=2)

        # LEFT: filters/search
        left.grid_rowconfigure(2, weight=1)
        left.grid_columnconfigure(0, weight=1)

        ttk.Label(left, text="Адресаты", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        filt = ttk.Frame(left, style="Card.TFrame")
        filt.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        filt.grid_columnconfigure(10, weight=1)

        self.filter_var = tk.StringVar(value="all")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh_table())

        # Чипы фильтра
        self._chip_buttons = {}

        def add_chip(text, val, col):
            b = ttk.Button(
                filt,
                text=text,
                style="Chip.TButton",
                command=lambda v=val: (self.filter_var.set(v), self._update_filter_chips(), self.refresh_table()),
            )
            b.grid(row=0, column=col, padx=(0, 6), sticky="w")
            self._chip_buttons[val] = b

        add_chip("Все", "all", 0)
        add_chip("Проблемные", "problems", 1)
        add_chip("Без пола", "no_gender", 2)
        add_chip("Без e-mail", "no_email", 3)
        add_chip("Отмеченные", "checked", 4)

        ttk.Label(filt, text="Поиск фамилии:").grid(row=0, column=8, sticky="e", padx=(16, 6))
        ttk.Entry(filt, textvariable=self.search_var, width=24).grid(row=0, column=9, sticky="e")

        right_tools = ttk.Frame(filt, style="Card.TFrame")
        right_tools.grid(row=0, column=10, sticky="e", padx=(8, 0))

        ttk.Button(
            right_tools,
            text="✕",
            width=3,
            command=lambda: (self.search_var.set(""), self.refresh_table()),
        ).pack(side="left")

        ttk.Button(
            right_tools,
            text="Автоширина",
            width=11,
            command=self._autofit_columns,
        ).pack(side="left", padx=(6, 0))

        self._update_filter_chips()

        # LEFT: table (одна аккуратная рамка)
        table_border = tk.Frame(left, bg=self.BORDER)
        table_border.grid(row=2, column=0, sticky="nsew")
        table_border.grid_rowconfigure(0, weight=1)
        table_border.grid_columnconfigure(0, weight=1)

        table_wrap = tk.Frame(table_border, bg=self.CARD_BG, highlightthickness=1, highlightbackground=self.BORDER)
        table_wrap.grid(row=0, column=0, sticky="nsew")
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(1, weight=0)
        table_wrap.grid_columnconfigure(0, weight=1)

        cols = ["✓", "Фамилия", "Имя", "Отчество", "Пол", "E-mail", "Статус"]
        self.cols = cols
        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings", selectmode="extended")

        # разумные дефолтные ширины (потом можно “Автоширина”)
        col_widths = {"✓": 55, "Фамилия": 180, "Имя": 160, "Отчество": 190, "Пол": 80, "E-mail": 260, "Статус": 360}

        for c in cols:
            self.tree.heading(c, text=c, anchor="center")
            self.tree.column(c, width=col_widths.get(c, 140), anchor="center")

        self.tree.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew")

        # Зебра + мягкие статусы
        self.tree.tag_configure("zebra0", background="#ffffff")
        self.tree.tag_configure("zebra1", background="#f8fafc")
        self.tree.tag_configure("ok", background="#f0fdf4")
        self.tree.tag_configure("bad_gender", background="#fff5f5")
        self.tree.tag_configure("bad_email", background="#fffbeb")

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.refresh_preview())

        # RIGHT: preview
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(2, weight=1)

        ttk.Label(right, text="Предпросмотр", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        self.preview_title = ttk.Label(right, text="—", style="CardSub.TLabel")
        self.preview_title.grid(row=1, column=0, sticky="w", pady=(8, 6))

        preview_wrap = tk.Frame(right, bg=self.CARD_BG, highlightthickness=1, highlightbackground=self.BORDER)
        preview_wrap.grid(row=2, column=0, sticky="nsew")
        preview_wrap.grid_rowconfigure(0, weight=1)
        preview_wrap.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(preview_wrap, bg="#eef1f6", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # --- BOTTOM actions
        bottom = ttk.Frame(self, style="TFrame")
        bottom.grid(row=2, column=0, sticky="ew", padx=PAD, pady=(10, PAD))
        bottom.grid_columnconfigure(0, weight=1)

        actions = ttk.Frame(bottom, style="Card.TFrame", padding=12)
        actions.grid(row=0, column=0, sticky="ew")
        actions.grid_columnconfigure(0, weight=1)

        ttk.Label(actions, text="Действия", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        btnrow = ttk.Frame(actions, style="Card.TFrame")
        btnrow.grid(row=1, column=0, sticky="ew", pady=(10, 0))

        self.tc_only_selected_var = tk.BooleanVar(value=False)

        self.btn_tc = ttk.Button(
            btnrow,
            text="🔎 Tatcenter: найти e-mail (пустые/битые)",
            style="Big.TButton",
            command=self.tatcenter_fetch,
        )
        self.btn_tc.grid(row=0, column=0, padx=(0, 10), pady=(0, 8), sticky="w")

        ttk.Checkbutton(btnrow, text="только выделенные", variable=self.tc_only_selected_var).grid(
            row=0, column=1, padx=(0, 16), pady=(0, 8), sticky="w"
        )

        self.btn_apply_tc = ttk.Button(
            btnrow, text="⬅ Заполнить E-mail из Tatcenter", style="Big.TButton", command=self.apply_tatcenter
        )
        self.btn_apply_tc.grid(row=0, column=2, padx=(0, 10), pady=(0, 8), sticky="w")

        self.btn_docx = ttk.Button(btnrow, text="🧾 Собрать DOCX", style="Big.TButton", command=self.generate_docx)
        self.btn_docx.grid(row=0, column=3, padx=(0, 10), pady=(0, 8), sticky="w")

        self.btn_pdf = ttk.Button(btnrow, text="🖨️ Собрать PDF", style="Big.TButton", command=self.generate_pdf)
        self.btn_pdf.grid(row=0, column=4, padx=(0, 10), pady=(0, 8), sticky="w")

        self.btn_open = ttk.Button(btnrow, text="📂 Открыть RESULT", style="Big.TButton", command=self.open_result)
        self.btn_open.grid(row=0, column=5, padx=(0, 10), pady=(0, 8), sticky="w")

        self.btn_export = ttk.Button(btnrow, text="⬇ Выгрузить PDF…", style="Big.TButton", command=self.export_pdf)
        self.btn_export.grid(row=0, column=6, pady=(0, 8), sticky="w")

        out = ttk.Frame(bottom, style="Card.TFrame", padding=12)
        out.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        out.grid_columnconfigure(3, weight=1)

        ttk.Label(out, text="Outlook", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", columnspan=7)

        ttk.Label(out, text="Отправлять с:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.sender_var = tk.StringVar(value=self.model.state.sender_email)
        self.sender_combo = ttk.Combobox(out, textvariable=self.sender_var, width=45, state="normal")
        self.sender_combo.grid(row=1, column=1, sticky="w", padx=(8, 14), pady=(10, 0))

        ttk.Label(out, text="Тема:").grid(row=1, column=2, sticky="w", pady=(10, 0))
        self.subject_var = tk.StringVar(value=self.model.state.subject)
        ttk.Entry(out, textvariable=self.subject_var, width=26).grid(
            row=1, column=3, sticky="ew", padx=(8, 14), pady=(10, 0)
        )

        self.btn_test = ttk.Button(out, text="🧪 ТЕСТ: выбранного себе", style="Big.TButton", command=self.send_test_one)
        self.btn_test.grid(row=1, column=4, padx=(0, 10), pady=(10, 0), sticky="w")

        self.btn_send_checked = ttk.Button(
            out, text="📨 Отправить отмеченным", style="Big.TButton", command=lambda: self.send_mails(True)
        )
        self.btn_send_checked.grid(row=1, column=5, padx=(0, 10), pady=(10, 0), sticky="w")

        self.btn_send_all = ttk.Button(out, text="📤 Отправить всем", style="Big.TButton", command=lambda: self.send_mails(False))
        self.btn_send_all.grid(row=1, column=6, pady=(10, 0), sticky="w")

        textbox = ttk.LabelFrame(
            bottom,
            text="Текст (для <<TEXT>> в DOCX). Для e-mail тело будет пустым.",
            padding=12,
        )
        textbox.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        textbox.grid_columnconfigure(0, weight=1)

        self.common_text = tk.Text(textbox, height=4, wrap="word")
        self.common_text.grid(row=0, column=0, sticky="ew")

        self.status_bar = ttk.Label(
            self,
            text="Клик по ✓ — отметить; клик по Пол — переключить; двойной клик по E-mail — редактировать",
            style="CardSub.TLabel",
        )
        self.status_bar.grid(row=3, column=0, sticky="ew", padx=PAD, pady=(0, 10))

    def _update_filter_chips(self):
        cur = self.filter_var.get()
        for val, btn in getattr(self, "_chip_buttons", {}).items():
            btn.configure(style=("ChipActive.TButton" if val == cur else "Chip.TButton"))

    # -------------------------
    # Load/select
    # -------------------------
    def load_excel(self):
        path = filedialog.askopenfilename(title="Выберите Excel (.xlsx)", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            self.ctrl.load_excel(path)
            self.excel_var.set(path)
            self.refresh_table()
            self._refresh_everything()
            # один раз после загрузки — ок
            self._autofit_columns()
        except Exception as e:
            messagebox.showerror("Excel", str(e))

    def load_template(self):
        path = filedialog.askopenfilename(title="Выберите шаблон DOCX", filetypes=[("Word", "*.docx")])
        if not path:
            return
        self.ctrl.load_template(path)
        self.template_var.set(path)
        self._refresh_everything()

    def choose_project_dir(self):
        d = filedialog.askdirectory(title="Выберите папку проекта")
        if not d:
            return
        self.ctrl.set_project_dir(d)
        self.project_var.set(d)
        self._refresh_everything()
        self.refresh_preview()

    # -------------------------
    # Table
    # -------------------------
    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        if self.model.df is None:
            self.view_idx = []
            return

        df = self.model.df.copy()

        f = self.filter_var.get()
        if f == "problems":
            df = df[df.apply(lambda r: self.model.compute_status_row(r)[2] != "ОК", axis=1)]
        elif f == "no_gender":
            df = df[df["Пол (итог)"].apply(norm_str) == ""]
        elif f == "no_email":
            df = df[~df["E-mail"].apply(lambda x: is_email_like(norm_str(x)))]
        elif f == "checked":
            df = df[df["Отправлять"] == True]  # noqa

        q = norm_str(self.search_var.get()).lower()
        if q:
            df = df[df["Фамилия"].apply(lambda x: norm_str(x).lower().startswith(q))]

        self.view_idx = list(df.index)

        for i, idx in enumerate(self.view_idx):
            row = self.model.df.loc[idx]
            g_ok, e_ok, status = self.model.compute_status_row(row)

            values = [
                "✓" if bool(row.get("Отправлять", True)) else "",
                row["Фамилия"], row["Имя"], row["Отчество"],
                row["Пол (итог)"], row["E-mail"], status,
            ]

            zebra = "zebra0" if (i % 2 == 0) else "zebra1"
            tags = [zebra]

            if status == "ОК":
                tags.append("ok")
            else:
                if not g_ok:
                    tags.append("bad_gender")
                if not e_ok:
                    tags.append("bad_email")

            self.tree.insert("", "end", iid=str(idx), values=values, tags=tuple(tags))

        items = self.tree.get_children()
        if items:
            self.tree.selection_set(items[0])
            self.tree.focus(items[0])
            self.tree.see(items[0])

        self.refresh_preview()
        self._refresh_everything()

    def _autofit_columns(self, sample_rows: int = 200):
        if not hasattr(self, "tree"):
            return
        try:
            font = tkfont.Font(font=self.tree.cget("font"))
        except Exception:
            font = tkfont.nametofont("TkDefaultFont")

        items = self.tree.get_children()
        sample = items[:sample_rows]
        min_w = 55
        pad = 26
        max_w = {"Статус": 620, "E-mail": 520, "✓": 60}
        default_max = 520

        for col in self.cols:
            maxlen = font.measure(str(col))
            for iid in sample:
                val = str(self.tree.set(iid, col))
                maxlen = max(maxlen, font.measure(val))
            width = max(min_w, maxlen + pad)
            width = min(width, max_w.get(col, default_max))
            self.tree.column(col, width=width)

    # -------------------------
    # Inline editing
    # -------------------------
    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        rowid = self.tree.identify_row(event.y)
        if not rowid or self.model.df is None:
            return

        col_name = self.cols[int(col.replace("#", "")) - 1]
        idx = int(rowid)

        if col_name == "✓":
            self.model.df.at[idx, "Отправлять"] = not bool(self.model.df.at[idx, "Отправлять"])
            self.refresh_table()
            self.tree.selection_set(str(idx))
            return

        if col_name == "Пол":
            self.model.df.at[idx, "Пол (итог)"] = toggle_gender(self.model.df.at[idx, "Пол (итог)"])
            self.refresh_table()
            self.tree.selection_set(str(idx))
            return

    def on_tree_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        rowid = self.tree.identify_row(event.y)
        if not rowid or self.model.df is None:
            return

        col_name = self.cols[int(col.replace("#", "")) - 1]
        if col_name == "E-mail":
            self.start_cell_edit(int(rowid))

    def start_cell_edit(self, idx: int):
        if self._edit_widget is not None:
            try:
                self._edit_widget.destroy()
            except Exception:
                pass
            self._edit_widget = None

        col_id = "#6"  # E-mail
        bbox = self.tree.bbox(str(idx), col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        var = tk.StringVar(value=self.model.df.at[idx, "E-mail"])
        ent = ttk.Entry(self.tree, textvariable=var)
        ent.place(x=x, y=y, width=w, height=h)
        ent.focus_set()

        def commit(*_):
            self.model.df.at[idx, "E-mail"] = norm_str(var.get())
            ent.destroy()
            self._edit_widget = None
            self.refresh_table()
            self.tree.selection_set(str(idx))

        ent.bind("<Return>", commit)
        ent.bind("<FocusOut>", commit)
        self._edit_widget = ent

    # -------------------------
    # Preview
    # -------------------------
    def _on_canvas_configure(self, _event):
        if self._preview_after_id is not None:
            try:
                self.after_cancel(self._preview_after_id)
            except Exception:
                pass
        self._preview_after_id = self.after(120, self.refresh_preview)

    def refresh_preview(self):
        self._preview_after_id = None
        self.canvas.delete("all")
        self.pdf_cache_imgtk = None
        self.preview_title.configure(text="—")

        if self.model.df is None:
            self._draw_empty_preview("Загрузите Excel.")
            return

        items = list(self.tree.get_children())
        if not items:
            self._draw_empty_preview("Нет строк.")
            return

        sel = self.tree.selection()
        iid = sel[0] if sel else items[0]
        idx = int(iid)

        try:
            pdf_path = self.model.pdf_path_for_idx(idx)
        except Exception:
            self._draw_empty_preview("Выберите папку проекта.")
            return

        self.preview_title.configure(text=os.path.basename(pdf_path))

        if not os.path.exists(pdf_path):
            self._draw_empty_preview("PDF не найден.\nСоберите PDF (Windows) или откройте RESULT/PDF.")
            return

        try:
            cw = max(500, self.canvas.winfo_width())
            ch = max(500, self.canvas.winfo_height())
            self.pdf_cache_imgtk = render_pdf_page_to_photoimage(pdf_path, 0, cw, ch)
            self.canvas.create_image(cw // 2, ch // 2, image=self.pdf_cache_imgtk, anchor="center")
        except Exception as e:
            self._draw_empty_preview(f"Не удалось отрисовать PDF:\n{e}")

    def _draw_empty_preview(self, text: str):
        cw = max(500, self.canvas.winfo_width())
        ch = max(500, self.canvas.winfo_height())
        self.canvas.create_text(
            cw // 2,
            ch // 2,
            text=text,
            fill="#6b7280",
            font=("Arial", 13),
            justify="center",
            anchor="center",
        )

    # -------------------------
    # Actions
    # -------------------------
    def tatcenter_fetch(self):
        if self.model.df is None:
            messagebox.showwarning("Tatcenter", "Сначала загрузите Excel.")
            return

        only_sel = bool(self.tc_only_selected_var.get())
        indices = None
        if only_sel:
            sel = list(self.tree.selection())
            if not sel:
                messagebox.showinfo("Tatcenter", "Включено «только выделенные» — выдели строки и повтори.")
                return
            indices = []
            for iid in sel:
                try:
                    indices.append(int(iid))
                except Exception:
                    pass

        if not messagebox.askyesno("Tatcenter", "Запускаем поиск на tatcenter.ru?"):
            return

        prog = ProgressDialog(self, "Tatcenter: поиск")

        def p(n, total):
            prog.set_progress(n, total)

        def msg(t):
            prog.set_text(t)

        try:
            res = self.ctrl.tatcenter_fetch(indices, p, msg, pause=1.0)
        finally:
            try:
                prog.destroy()
            except Exception:
                pass

        self.refresh_table()

        msg_txt = (
            f"Поиск завершён ({res['scope']}).\n"
            f"Найдено e-mail: {res['found']}\n"
            f"Не найдено: {res['not_found']}\n"
            f"Ошибки: {res['errors']}\n"
        )
        messagebox.showinfo("Tatcenter", msg_txt)

    def apply_tatcenter(self):
        try:
            moved = self.ctrl.apply_tatcenter_to_main_email()
            self.refresh_table()
            messagebox.showinfo("Tatcenter → E-mail", f"Перенесено строк: {moved}")
        except Exception as e:
            messagebox.showerror("Tatcenter → E-mail", str(e))

    def generate_docx(self):
        prog = None
        try:
            text = self.common_text.get("1.0", "end").rstrip("\n")
            prog = ProgressDialog(self, "Генерация DOCX")
            total = len(self.model.df) if self.model.df is not None else 1
            prog.set_total(total)

            def p(n, total):
                prog.set_progress(n, total)

            def msg(t):
                prog.set_text(t)

            out_dir = self.ctrl.generate_docx(text, p, msg)
            prog.destroy()
            messagebox.showinfo("DOCX", f"Готово. DOCX сохранены в:\n{out_dir}")
            self._refresh_everything()
        except Exception as e:
            try:
                if prog is not None:
                    prog.destroy()
            except Exception:
                pass
            messagebox.showerror("DOCX", str(e))

    def generate_pdf(self):
        try:
            out_dir = self.ctrl.generate_pdf()
            messagebox.showinfo("PDF", f"Готово. PDF сохранены в:\n{out_dir}")
            self._refresh_everything()
            self.refresh_preview()
        except Exception as e:
            messagebox.showerror("PDF", str(e))

    def open_result(self):
        try:
            self.ctrl.open_result()
        except Exception as e:
            messagebox.showerror("RESULT", str(e))

    def export_pdf(self):
        try:
            dest = filedialog.askdirectory(title="Куда выгрузить PDF?")
            if not dest:
                return
            res = self.ctrl.export_pdf_files(dest)
            messagebox.showinfo(
                "Экспорт PDF",
                f"Скопировано: {res['copied']}\nНе найдено: {res['missing']}\nОшибки: {res['errors']}\n\nПапка:\n{res['dest']}",
            )
        except Exception as e:
            messagebox.showerror("Экспорт PDF", str(e))

    def send_test_one(self):
        if self.model.df is None:
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Тест", "Выберите строку.")
            return
        idx = int(sel[0])
        try:
            sender = norm_str(self.sender_var.get()) or self.model.state.sender_email
            subject = norm_str(self.subject_var.get()) or "Поздравление"
            to, pdfname = self.ctrl.send_test_one(sender, subject, idx)
            messagebox.showinfo("Тест", f"Отправлено на {to}:\n{pdfname}")
        except Exception as e:
            messagebox.showerror("Outlook", str(e))

    def send_mails(self, only_checked: bool):
        try:
            sender = norm_str(self.sender_var.get()) or self.model.state.sender_email
            subject = norm_str(self.subject_var.get()) or "Поздравление"

            if not messagebox.askyesno("Подтверждение", f"От: {sender}\nТема: {subject}\nТело: пустое\n\nОтправляем?"):
                return

            out_csv = self.ctrl.send_mails(sender, subject, only_checked)
            messagebox.showinfo("Outlook", f"Готово.\nОтчет:\n{out_csv}")
        except Exception as e:
            messagebox.showerror("Outlook", str(e))

    # -------------------------
    # Status + buttons gating
    # -------------------------
    def _refresh_everything(self):
        df = self.model.df
        if df is None:
            self.st_data.configure(text="Данные: нет", style="StatusBad.TLabel")
            self.st_gender.configure(text="Пол: —", style="StatusWarn.TLabel")
            self.st_email.configure(text="E-mail: —", style="StatusWarn.TLabel")
            self.st_pdf.configure(text="PDF: —", style="StatusWarn.TLabel")
            self._set_buttons_enabled(False)
            self._refresh_accounts()
            return

        total = len(df)
        g_empty = int((df["Пол (итог)"].apply(norm_str) == "").sum())
        e_bad = int((~df["E-mail"].apply(lambda x: is_email_like(norm_str(x)))).sum())

        self.st_data.configure(text=f"Данные: {total}", style="StatusOK.TLabel")
        self.st_gender.configure(
            text=("Пол: OK" if g_empty == 0 else f"Пол: проблем {g_empty}"),
            style=("StatusOK.TLabel" if g_empty == 0 else "StatusBad.TLabel"),
        )
        self.st_email.configure(
            text=("E-mail: OK" if e_bad == 0 else f"E-mail: проблем {e_bad}"),
            style=("StatusOK.TLabel" if e_bad == 0 else "StatusWarn.TLabel"),
        )

        pdf_count = 0
        if self.model.state.project_dir:
            try:
                pdf_dir = self.model.result_dir("PDF")
                if os.path.isdir(pdf_dir):
                    pdf_count = len([f for f in os.listdir(pdf_dir) if f.lower().endswith(".pdf")])
            except Exception:
                pdf_count = 0

        self.st_pdf.configure(
            text=(f"PDF: {pdf_count}" if pdf_count else "PDF: нет"),
            style=("StatusOK.TLabel" if pdf_count else "StatusWarn.TLabel"),
        )

        self._refresh_accounts()
        self._set_buttons_enabled(True)

    def _refresh_accounts(self):
        accs = self.ctrl.outlook_accounts() if WIN else []
        cur = norm_str(self.sender_var.get())
        base_default = self.model.state.sender_email

        if base_default and base_default not in accs:
            accs = [base_default] + accs
        if cur and cur not in accs:
            accs = [cur] + accs
        if not accs:
            accs = [base_default] if base_default else []

        self.sender_combo["values"] = accs
        if not cur and accs:
            self.sender_combo.set(accs[0])

    def _set_buttons_enabled(self, has_data: bool):
        has_excel = has_data
        has_template = bool(self.model.state.template_path)
        has_project = bool(self.model.state.project_dir)

        self.btn_tc.configure(state=("normal" if has_excel else "disabled"))
        self.btn_apply_tc.configure(state=("normal" if has_excel else "disabled"))
        self.btn_docx.configure(state=("normal" if (has_excel and has_template and has_project) else "disabled"))

        if WIN and has_excel and has_project:
            self.btn_pdf.configure(state="normal")
            self.btn_test.configure(state="normal")
            self.btn_send_checked.configure(state="normal")
            self.btn_send_all.configure(state="normal")
        else:
            self.btn_pdf.configure(state="disabled")
            self.btn_test.configure(state="disabled")
            self.btn_send_checked.configure(state="disabled")
            self.btn_send_all.configure(state="disabled")

        self.btn_open.configure(state=("normal" if has_project else "disabled"))
        self.btn_export.configure(state=("normal" if has_project else "disabled"))
