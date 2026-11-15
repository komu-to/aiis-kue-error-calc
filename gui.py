"""
Графический интерфейс (GUI) для Windows 10/11
Генератор таблиц погрешностей измерительных каналов АИИС КУЭ
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import List, Dict, Any, Tuple, Optional

from core import (
    TT_PARAMS,
    TH_PARAMS,
    METER_BASE_PROFILES,
    PRESET_PROJECT_GROUPS,
    generate_group_table,
)
from exporters import (
    HEADERS_SUMMARY,
    HEADERS_FULL,
    get_tsv_text,
    export_to_excel,
    export_to_docx,
    make_export_filename,
    export_table51_to_docx,
    export_table51_to_excel,
    get_table51_tsv_text,
)

TEMP_PRESETS = [
    ("0 ... +40 °C (Паспорт-формуляр ФО, Δt=20°C)", 20.0),
    ("+10 ... +35 °C (Отапливаемый шкаф, Δt=15°C)", 15.0),
    ("-3 ... +43 °C (Отчет об испытаниях, Δt=23°C)", 23.0),
    ("-40 ... +55 °C (Наружная установка, Δt=60°C)", 60.0),
    ("Пользовательское значение...", None),
]

MAG_PRESETS = [
    ("По ГОСТ 31819 (0,5% / 1,0% / 2,0% при Iном, cos φ=1) [По умолчанию]", "gost"),
    ("Во всех точках диапазона (по паспорту счетчика, напр. РИМ)", "all_points"),
    ("Пользовательское значение (ввод в %)...", "custom"),
    ("Не учитывать (н/н во всех точках)", "none"),
]


class ErrorCalcApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Генератор таблиц погрешностей ИК АИИС КУЭ (ГОСТ Р 8.596 / МИ 3000)")
        self.root.geometry("1180x740")
        self.root.minsize(960, 600)

        # Текущие рассчитанные строки
        self.current_rows: List[Dict[str, Any]] = []

        # Настройка стиля ttk
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.setup_ui()

        # Первичный расчет по умолчанию
        self.on_calculate()

    def setup_ui(self):
        # Главный контейнер
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Верхняя панель заголовка ---
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        title_label = ttk.Label(
            header_frame,
            text="Расчет пределов погрешностей измерительных каналов АИИС КУЭ",
            font=("Segoe UI", 13, "bold"),
            foreground="#1F497D"
        )
        title_label.pack(side=tk.LEFT)

        subtitle_label = ttk.Label(
            header_frame,
            text="ГОСТ 7746-2015 • ГОСТ 1983-2015 • МИ 3000-2006 • ГОСТ Р 8.596-2002",
            font=("Segoe UI", 9, "italic"),
            foreground="#555555"
        )
        subtitle_label.pack(side=tk.RIGHT, pady=(4, 0))

        # --- Панель настроек (Controls) ---
        controls_group = ttk.LabelFrame(main_frame, text=" Параметры измерительного канала, климатических условий и влияющих величин ", padding="10")
        controls_group.pack(fill=tk.X, pady=(0, 10))

        # Сетка параметров: 3 колонки
        grid_frame = ttk.Frame(controls_group)
        grid_frame.pack(fill=tk.X)

        # --- Колонка 1: ТТ и ТН ---
        c1 = ttk.Frame(grid_frame)
        c1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        ttk.Label(c1, text="Класс точности ТТ:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=2)
        self.tt_var = tk.StringVar(value="0.5S")
        self.tt_combo = ttk.Combobox(c1, textvariable=self.tt_var, values=list(TT_PARAMS.keys()), state="readonly", width=16)
        self.tt_combo.grid(row=0, column=1, sticky=tk.W, pady=2, padx=5)
        self.tt_combo.bind("<<ComboboxSelected>>", self.on_tt_changed)

        ttk.Label(c1, text="Класс точности ТН:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=tk.W, pady=2)
        self.th_var = tk.StringVar(value="0.5")
        self.th_combo = ttk.Combobox(c1, textvariable=self.th_var, values=list(TH_PARAMS.keys()), state="readonly", width=16)
        self.th_combo.grid(row=1, column=1, sticky=tk.W, pady=2, padx=5)
        self.th_combo.bind("<<ComboboxSelected>>", self.on_th_changed)

        ttk.Label(c1, text="Потери в линии ТН (δL, %):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.delta_L_var = tk.StringVar(value="0.25")
        self.delta_L_entry = ttk.Entry(c1, textvariable=self.delta_L_var, width=18)
        self.delta_L_entry.grid(row=2, column=1, sticky=tk.W, pady=2, padx=5)

        # --- Колонка 2: Счетчик и номер ИК ---
        c2 = ttk.Frame(grid_frame)
        c2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        ttk.Label(c2, text="Класс счетчика (акт/реакт):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=2)
        self.meter_var = tk.StringVar(value="0.2S / 0.5")
        self.meter_combo = ttk.Combobox(c2, textvariable=self.meter_var, values=list(METER_BASE_PROFILES.keys()), state="readonly", width=18)
        self.meter_combo.grid(row=0, column=1, sticky=tk.W, pady=2, padx=5)
        self.meter_combo.bind("<<ComboboxSelected>>", lambda e: self.on_calculate())

        ttk.Label(c2, text="Номер / обозначение ИК:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.ik_name_var = tk.StringVar(value="1, 2, 15, 16")
        self.ik_name_entry = ttk.Entry(c2, textvariable=self.ik_name_var, width=20)
        self.ik_name_entry.grid(row=1, column=1, sticky=tk.W, pady=2, padx=5)

        ttk.Label(c2, text="Магнитная индукция (δми):", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky=tk.W, pady=2)
        self.mag_preset_var = tk.StringVar(value=MAG_PRESETS[0][0])
        self.mag_combo = ttk.Combobox(
            c2,
            textvariable=self.mag_preset_var,
            values=[title for title, _ in MAG_PRESETS],
            state="readonly",
            width=28
        )
        self.mag_combo.grid(row=2, column=1, sticky=tk.W, pady=2, padx=5)
        self.mag_combo.bind("<<ComboboxSelected>>", self.on_mag_mode_changed)

        # --- Колонка 3: Температура и свои значения ---
        c3 = ttk.Frame(grid_frame)
        c3.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        ttk.Label(c3, text="Температурные условия:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=2)
        self.temp_preset_var = tk.StringVar(value=TEMP_PRESETS[0][0])
        self.temp_combo = ttk.Combobox(
            c3,
            textvariable=self.temp_preset_var,
            values=[name for name, _ in TEMP_PRESETS],
            state="readonly",
            width=32
        )
        self.temp_combo.grid(row=0, column=1, sticky=tk.W, pady=2, padx=5)
        self.temp_combo.bind("<<ComboboxSelected>>", self.on_temp_preset_changed)

        ttk.Label(c3, text="Отклонение темп. (Δt, °C):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.delta_t_var = tk.StringVar(value="20.0")
        self.delta_t_entry = ttk.Entry(c3, textvariable=self.delta_t_var, width=12)
        self.delta_t_entry.grid(row=1, column=1, sticky=tk.W, pady=2, padx=5)

        ttk.Label(c3, text="Свое δми (% при необходимости):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.custom_dmi_var = tk.StringVar(value="0.5")
        self.custom_dmi_entry = ttk.Entry(c3, textvariable=self.custom_dmi_var, width=12, state=tk.DISABLED)
        self.custom_dmi_entry.grid(row=2, column=1, sticky=tk.W, pady=2, padx=5)

        # --- Панель кнопок и переключателя вида ---
        btn_bar = ttk.Frame(controls_group)
        btn_bar.pack(fill=tk.X, pady=(8, 0))

        # Переключатель формата таблицы
        view_frame = ttk.LabelFrame(btn_bar, text=" Формат отображения ", padding="3")
        view_frame.pack(side=tk.LEFT, padx=(0, 15))

        self.mode_var = tk.StringVar(value="full")  # По умолчанию открываем полную таблицу с 29 графами, чтобы δми сразу было видно!
        ttk.Radiobutton(
            view_frame,
            text="Полная таблица (29 граф по МИ 3000)",
            variable=self.mode_var,
            value="full",
            command=self.refresh_table_view
        ).pack(side=tk.LEFT, padx=5)

        ttk.Radiobutton(
            view_frame,
            text="Сводная таблица (15 граф)",
            variable=self.mode_var,
            value="summary",
            command=self.refresh_table_view
        ).pack(side=tk.LEFT, padx=5)

        # Кнопки действий
        btn_action_frame = ttk.Frame(btn_bar)
        btn_action_frame.pack(side=tk.RIGHT)

        ttk.Button(btn_action_frame, text="🚀 Рассчитать вариант", command=self.on_calculate).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_action_frame, text="📑 Все группы проекта", command=self.on_calculate_all_groups).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_action_frame, text="📋 В буфер", command=self.on_copy_clipboard).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_action_frame, text="📊 Excel", command=self.on_export_excel).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_action_frame, text="📝 Word", command=self.on_export_docx).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_action_frame, text="📑 Таблица 5.1 (Word)", command=self.on_export_table51_docx).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_action_frame, text="📊 Таблица 5.1 (Excel)", command=self.on_export_table51_excel).pack(side=tk.LEFT, padx=3)

        # --- Область отображения таблицы (Treeview со скроллбарами) ---
        table_container = ttk.Frame(main_frame)
        table_container.pack(fill=tk.BOTH, expand=True, pady=(5, 5))

        self.tree_scroll_y = ttk.Scrollbar(table_container, orient=tk.VERTICAL)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_scroll_x = ttk.Scrollbar(table_container, orient=tk.HORIZONTAL)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            table_container,
            yscrollcommand=self.tree_scroll_y.set,
            xscrollcommand=self.tree_scroll_x.set,
            selectmode="extended"
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        # Контекстное меню по правому клику
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Копировать выделенное (TSV)", command=self.copy_selection)
        self.context_menu.add_command(label="Копировать всю таблицу", command=self.on_copy_clipboard)
        self.context_menu.add_command(label="Копировать как Таблицу 5.1 (δW осн)", command=self.on_copy_table51)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Экспорт в Excel (.xlsx)", command=self.on_export_excel)
        self.context_menu.add_command(label="Экспорт в Word (.docx)", command=self.on_export_docx)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Экспорт в Таблицу 5.1 (Word, δW осн)", command=self.on_export_table51_docx)
        self.context_menu.add_command(label="Экспорт в Таблицу 5.1 (Excel, δW осн)", command=self.on_export_table51_excel)
        self.tree.bind("<Button-3>", self.show_context_menu)

        # Привязка клавиши Ctrl+C
        self.root.bind("<Control-c>", lambda e: self.copy_selection())

        # --- Статусная строка ---
        self.status_var = tk.StringVar(value="Готово к расчету.")
        status_bar = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding=(5, 3),
            font=("Segoe UI", 9)
        )
        status_bar.pack(fill=tk.X, pady=(2, 0))

    def on_tt_changed(self, event=None):
        tt = self.tt_var.get()
        if tt == "Без ТТ":
            if self.th_var.get() != "Без ТН":
                self.th_var.set("Без ТН")
                self.delta_L_var.set("0.00")
        self.on_calculate()

    def on_th_changed(self, event=None):
        th = self.th_var.get()
        if th == "Без ТН":
            self.delta_L_var.set("0.00")
        else:
            if float(self.delta_L_var.get().replace(",", ".") or "0") == 0.0:
                self.delta_L_var.set("0.25")
        self.on_calculate()

    def on_temp_preset_changed(self, event=None):
        selected_title = self.temp_preset_var.get()
        for name, val in TEMP_PRESETS:
            if name == selected_title:
                if val is not None:
                    self.delta_t_var.set(str(val))
                    self.on_calculate()
                break

    def on_mag_mode_changed(self, event=None):
        selected_title = self.mag_preset_var.get()
        is_custom = ("custom" in selected_title.lower() or "пользовательское" in selected_title.lower())
        self.custom_dmi_entry.config(state=tk.NORMAL if is_custom else tk.DISABLED)
        self.on_calculate()

    def get_parsed_params(self) -> Tuple[str, str, str, str, float, float, str, float]:
        ik_num = self.ik_name_var.get().strip() or "ИК"
        tt_class = self.tt_var.get()
        th_class = self.th_var.get()
        meter_class = self.meter_var.get()

        try:
            delta_t = float(self.delta_t_var.get().replace(",", "."))
        except ValueError:
            delta_t = 20.0
            self.delta_t_var.set("20.0")

        try:
            delta_L = float(self.delta_L_var.get().replace(",", "."))
        except ValueError:
            delta_L = 0.25
            self.delta_L_var.set("0.25")

        # Режим влияния магнитной индукции
        selected_mag = self.mag_preset_var.get()
        mag_mode = "gost"
        for title, mode_key in MAG_PRESETS:
            if title == selected_mag:
                mag_mode = mode_key
                break

        try:
            custom_dmi = float(self.custom_dmi_var.get().replace(",", "."))
        except ValueError:
            custom_dmi = 0.5
            self.custom_dmi_var.set("0.5")

        return ik_num, tt_class, th_class, meter_class, delta_t, delta_L, mag_mode, custom_dmi

    def on_calculate(self):
        """Расчет одного выбранного варианта."""
        ik_num, tt_class, th_class, meter_class, delta_t, delta_L, mag_mode, custom_dmi = self.get_parsed_params()
        self.is_all_groups = False
        self.current_calc_params = {
            "ik": ik_num,
            "tt": tt_class,
            "th": th_class,
            "meter": meter_class,
            "delta_t": delta_t
        }
        self.current_rows = generate_group_table(
            ik_num=ik_num,
            tt_class=tt_class,
            th_class=th_class,
            meter_class=meter_class,
            delta_t=delta_t,
            delta_L=delta_L,
            magnetic_mode=mag_mode,
            custom_dmi_val=custom_dmi
        )
        self.refresh_table_view()
        self.status_var.set(f"Рассчитан канал [{ik_num}]: ТТ {tt_class}, ТН {th_class}, Счетчик {meter_class}, Δt={delta_t}°C, Магн. индукция={mag_mode}. Строк: {len(self.current_rows)}")

    def on_calculate_all_groups(self):
        """Пакетный расчет всех типовых групп каналов проекта."""
        _, _, _, _, delta_t, _, mag_mode, custom_dmi = self.get_parsed_params()
        self.is_all_groups = True
        self.current_calc_params = {
            "ik": "Все_ИК_1-23",
            "tt": "все",
            "th": "все",
            "meter": "все",
            "delta_t": delta_t
        }
        all_rows = []
        for g in PRESET_PROJECT_GROUPS:
            rows = generate_group_table(
                ik_num=g["ik"],
                tt_class=g["tt"],
                th_class=g["th"],
                meter_class=g["meter"],
                delta_t=delta_t,
                delta_L=g["delta_L"],
                magnetic_mode=mag_mode,
                custom_dmi_val=custom_dmi
            )
            all_rows.extend(rows)

        self.current_rows = all_rows
        self.refresh_table_view()
        self.status_var.set(f"Пакетный расчет завершен: рассчитано {len(PRESET_PROJECT_GROUPS)} групп ({len(all_rows)} строк)")

    def refresh_table_view(self):
        """Обновление колонок и данных в Treeview."""
        mode = self.mode_var.get()
        headers_spec = HEADERS_SUMMARY if mode == "summary" else HEADERS_FULL

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = [col_id for col_id, _ in headers_spec]
        self.tree["show"] = "headings"

        # Настройка заголовков колонок
        for col_id, label in headers_spec:
            self.tree.heading(col_id, text=label, anchor=tk.CENTER)
            min_width = max(len(label) * 8 + 14, 60)
            self.tree.column(col_id, width=min_width, anchor=tk.CENTER, stretch=False)

        # Заполнение данными
        for r_idx, r in enumerate(self.current_rows):
            vals = [r.get(col_id, "") for col_id, _ in headers_spec]
            tag = "even" if r_idx % 2 == 0 else "odd"
            self.tree.insert("", tk.END, values=vals, tags=(tag,))

        self.tree.tag_configure("odd", background="#F5F8FA")
        self.tree.tag_configure("even", background="#FFFFFF")

    def show_context_menu(self, event):
        self.context_menu.post(event.x_root, event.y_root)

    def copy_selection(self):
        selected_items = self.tree.selection()
        if not selected_items:
            self.on_copy_clipboard()
            return

        mode = self.mode_var.get()
        headers_spec = HEADERS_SUMMARY if mode == "summary" else HEADERS_FULL
        header_line = "\t".join(label for _, label in headers_spec)

        lines = [header_line]
        for item_id in selected_items:
            vals = self.tree.item(item_id, "values")
            lines.append("\t".join(str(v) for v in vals))

        text_to_copy = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text_to_copy)
        self.status_var.set(f"Скопировано выделенных строк в буфер обмена: {len(selected_items)}")

    def on_copy_clipboard(self):
        if not self.current_rows:
            messagebox.showwarning("Предупреждение", "Нет данных для копирования.")
            return

        mode = self.mode_var.get()
        tsv_text = get_tsv_text(self.current_rows, mode=mode)
        self.root.clipboard_clear()
        self.root.clipboard_append(tsv_text)
        self.status_var.set(f"Таблица ({len(self.current_rows)} строк) скопирована в буфер обмена (TSV для Excel/Word).")
        messagebox.showinfo("Буфер обмена", "Таблица успешно скопирована!\nТеперь вы можете нажать Ctrl+V в Excel или Word.")

    def on_export_excel(self):
        if not self.current_rows:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта.")
            return

        mode = self.mode_var.get()
        params = getattr(self, "current_calc_params", {})
        def_name = make_export_filename(
            ik_str=params.get("ik", self.ik_var.get()),
            tt_str=params.get("tt", self.tt_var.get()),
            th_str=params.get("th", self.th_var.get()),
            meter_str=params.get("meter", self.meter_var.get()),
            delta_t=params.get("delta_t", self.delta_t_var.get()),
            ext="xlsx",
            is_all_groups=getattr(self, "is_all_groups", False)
        )
        filepath = filedialog.asksaveasfilename(
            title="Сохранить таблицу в Excel",
            defaultextension=".xlsx",
            filetypes=[("Файлы Excel (*.xlsx)", "*.xlsx"), ("Все файлы (*.*)", "*.*")],
            initialfile=def_name
        )
        if not filepath:
            return

        try:
            if getattr(self, "is_all_groups", False):
                title = f"Сводная таблица погрешностей всех групп ИК АИИС КУЭ (Δt={params.get('delta_t', self.delta_t_var.get())}°C)"
            else:
                title = (
                    f"Таблица погрешностей ИК {params.get('ik', self.ik_var.get())} "
                    f"(ТТ: {params.get('tt', self.tt_var.get())}, "
                    f"ТН: {params.get('th', self.th_var.get())}, "
                    f"Счетчик: {params.get('meter', self.meter_var.get())}, "
                    f"Δt={params.get('delta_t', self.delta_t_var.get())}°C)"
                )
            export_to_excel(self.current_rows, filepath, mode=mode, title=title)
            self.status_var.set(f"Успешно экспортировано в Excel: {os.path.basename(filepath)}")
            messagebox.showinfo("Экспорт завершен", f"Таблица успешно сохранена в файл:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Ошибка экспорта", f"Не удалось сохранить файл Excel:\n{e}")

    def on_export_docx(self):
        if not self.current_rows:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта.")
            return

        mode = self.mode_var.get()
        params = getattr(self, "current_calc_params", {})
        def_name = make_export_filename(
            ik_str=params.get("ik", self.ik_var.get()),
            tt_str=params.get("tt", self.tt_var.get()),
            th_str=params.get("th", self.th_var.get()),
            meter_str=params.get("meter", self.meter_var.get()),
            delta_t=params.get("delta_t", self.delta_t_var.get()),
            ext="docx",
            is_all_groups=getattr(self, "is_all_groups", False)
        )
        filepath = filedialog.asksaveasfilename(
            title="Сохранить таблицу в Word",
            defaultextension=".docx",
            filetypes=[("Файлы Word (*.docx)", "*.docx"), ("Все файлы (*.*)", "*.*")],
            initialfile=def_name
        )
        if not filepath:
            return

        try:
            if getattr(self, "is_all_groups", False):
                title = f"Таблица пределов погрешностей измерений всех каналов АИИС КУЭ (Δt={params.get('delta_t', self.delta_t_var.get())}°C)"
            else:
                title = (
                    f"Таблица погрешностей ИК {params.get('ik', self.ik_var.get())} "
                    f"(ТТ: {params.get('tt', self.tt_var.get())}, "
                    f"ТН: {params.get('th', self.th_var.get())}, "
                    f"Счетчик: {params.get('meter', self.meter_var.get())}, "
                    f"Δt={params.get('delta_t', self.delta_t_var.get())}°C)"
                )
            export_to_docx(self.current_rows, filepath, mode=mode, title=title)
            self.status_var.set(f"Успешно экспортировано в Word: {os.path.basename(filepath)}")
            messagebox.showinfo("Экспорт завершен", f"Таблица успешно сохранена в файл Word:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Ошибка экспорта", f"Не удалось сохранить файл Word:\n{e}")

    def on_copy_table51(self):
        """Копирование данных в формате Таблицы 5.1 (δW осн) в буфер обмена."""
        if not self.current_rows:
            messagebox.showwarning("Предупреждение", "Нет данных для копирования.")
            return

        tsv_text = get_table51_tsv_text(self.current_rows, val_type="osn")
        self.root.clipboard_clear()
        self.root.clipboard_append(tsv_text)
        self.status_var.set("Матрица Таблицы 5.1 (δW осн) скопирована в буфер обмена (TSV).")
        messagebox.showinfo("Буфер обмена", "Матрица Таблицы 5.1 (δW осн: акт и реакт) скопирована в буфер обмена!\nТеперь вы можете нажать Ctrl+V в Excel или Word.")

    def on_export_table51_docx(self):
        """Экспорт в формат Таблицы 5.1 проекта (Word, δW осн)."""
        if not self.current_rows:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта.")
            return

        params = getattr(self, "current_calc_params", {})
        if getattr(self, "is_all_groups", False):
            def_name = f"Таблица_5.1_δW_осн-Все_ИК_1-23.docx"
            title = f"Таблица 5.1 — Пределы допускаемой основной погрешности измерений ИК АИИС КУЭ, ± δWосн (%)"
        else:
            base_fn = make_export_filename(
                ik_str=params.get("ik", self.ik_var.get()),
                tt_str=params.get("tt", self.tt_var.get()),
                th_str=params.get("th", self.th_var.get()),
                meter_str=params.get("meter", self.meter_var.get()),
                delta_t=params.get("delta_t", self.delta_t_var.get()),
                ext="docx"
            )
            def_name = f"Таблица_5.1_δW_осн-{base_fn}"
            title = f"Таблица 5.1 — Пределы допускаемой основной погрешности ИК {params.get('ik', self.ik_var.get())}, ± δWосн (%)"

        filepath = filedialog.asksaveasfilename(
            title="Сохранить в формате Таблицы 5.1 (Word)",
            defaultextension=".docx",
            filetypes=[("Файлы Word (*.docx)", "*.docx"), ("Все файлы (*.*)", "*.*")],
            initialfile=def_name
        )
        if not filepath:
            return

        try:
            export_table51_to_docx(self.current_rows, filepath, val_type="osn", title=title)
            self.status_var.set(f"Таблица 5.1 (δW осн) успешно экспортирована в Word: {os.path.basename(filepath)}")
            messagebox.showinfo("Экспорт завершен", f"Таблица 5.1 (δW осн) успешно сохранена в файл Word:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Ошибка экспорта", f"Не удалось сохранить Таблицу 5.1 в Word:\n{e}")

    def on_export_table51_excel(self):
        """Экспорт в формат Таблицы 5.1 проекта (Excel, δW осн)."""
        if not self.current_rows:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта.")
            return

        params = getattr(self, "current_calc_params", {})
        if getattr(self, "is_all_groups", False):
            def_name = f"Таблица_5.1_δW_осн-Все_ИК_1-23.xlsx"
            title = f"Таблица 5.1 — Пределы допускаемой основной погрешности измерений ИК АИИС КУЭ, ± δWосн (%)"
        else:
            base_fn = make_export_filename(
                ik_str=params.get("ik", self.ik_var.get()),
                tt_str=params.get("tt", self.tt_var.get()),
                th_str=params.get("th", self.th_var.get()),
                meter_str=params.get("meter", self.meter_var.get()),
                delta_t=params.get("delta_t", self.delta_t_var.get()),
                ext="xlsx"
            )
            def_name = f"Таблица_5.1_δW_осн-{base_fn}"
            title = f"Таблица 5.1 — Пределы допускаемой основной погрешности ИК {params.get('ik', self.ik_var.get())}, ± δWосн (%)"

        filepath = filedialog.asksaveasfilename(
            title="Сохранить в формате Таблицы 5.1 (Excel)",
            defaultextension=".xlsx",
            filetypes=[("Файлы Excel (*.xlsx)", "*.xlsx"), ("Все файлы (*.*)", "*.*")],
            initialfile=def_name
        )
        if not filepath:
            return

        try:
            export_table51_to_excel(self.current_rows, filepath, val_type="osn", title=title)
            self.status_var.set(f"Таблица 5.1 (δW осн) успешно экспортирована в Excel: {os.path.basename(filepath)}")
            messagebox.showinfo("Экспорт завершен", f"Таблица 5.1 (δW осн) успешно сохранена в файл Excel:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Ошибка экспорта", f"Не удалось сохранить Таблицу 5.1 в Excel:\n{e}")


def main():
    root = tk.Tk()
    app = ErrorCalcApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
