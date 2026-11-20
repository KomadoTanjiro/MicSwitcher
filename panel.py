"""Панель настроек: одно окно, которое не закрывается после каждого действия.

Окно живёт в потоке tkinter приложения (PopupManager). Панель ничего не знает
о звуке: состояние берёт снимком actions.state() (словарь из panel_state()),
а действия вызывает в фоновом потоке и после них перерисовывается.

actions - объект с методами приложения:
    state(), call_ui(fn, *args), select_device(kind, slot, id), activate_slot(kind, slot),
    toggle(kind), apply_profile(name), save_profile_named(name), delete_profile(name),
    set_profile_hotkey(name, combo), set_eq(kind, preset), set_rnnoise_threshold(value),
    set_hotkeys(values), set_autorun(enabled), set_auto_refresh(enabled),
    record_hotkey(), open_target(target), install_equalizer_apo(), install_noise_suppression(),
    open_config(), open_log(), reload_config().
"""
import threading
import tkinter as tk
from tkinter import ttk

BG = "#111827"
CARD = "#1f2937"
FIELD = "#374151"
FG = "#f9fafb"
MUTED = "#9ca3af"
ACCENT = "#3c78dc"
FONT = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 12, "bold")

HOTKEY_FIELDS = (
    ("hotkey", "Переключить микрофон"),
    ("hotkey_out", "Переключить вывод"),
    ("hotkey_profile", "Следующий профиль"),
)


def setup_style(root):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(".", background=BG, foreground=FG, fieldbackground=FIELD, font=FONT)
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD)
    style.configure("TLabel", background=BG, foreground=FG)
    style.configure("Card.TLabel", background=CARD, foreground=FG)
    style.configure("Muted.TLabel", background=CARD, foreground=MUTED)
    style.configure("Title.TLabel", background=CARD, foreground=FG, font=FONT_BOLD)
    style.configure("Header.TLabel", background=BG, foreground=FG, font=FONT_TITLE)
    style.configure("HeaderMuted.TLabel", background=BG, foreground=MUTED)
    style.configure("TButton", background=FIELD, foreground=FG, borderwidth=0, padding=(10, 4))
    style.map("TButton", background=[("active", "#4b5563"), ("disabled", CARD)], foreground=[("disabled", MUTED)])
    style.configure("Accent.TButton", background=ACCENT, foreground="white")
    style.map("Accent.TButton", background=[("active", "#2f63bd")])
    style.configure("TCombobox", fieldbackground=FIELD, background=FIELD, foreground=FG, arrowcolor=FG)
    style.map("TCombobox", fieldbackground=[("readonly", FIELD)], foreground=[("readonly", FG)])
    style.configure("TEntry", fieldbackground=FIELD, foreground=FG, insertcolor=FG)
    style.configure("TCheckbutton", background=CARD, foreground=FG)
    style.map("TCheckbutton", background=[("active", CARD)])
    style.configure("TNotebook", background=BG, borderwidth=0)
    style.configure("TNotebook.Tab", background=CARD, foreground=MUTED, padding=(14, 6))
    style.map("TNotebook.Tab", background=[("selected", FIELD)], foreground=[("selected", FG)])
    root.option_add("*TCombobox*Listbox.background", FIELD)
    root.option_add("*TCombobox*Listbox.foreground", FG)
    root.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
    root.option_add("*TCombobox*Listbox.font", FONT)


class Choice:
    """Combobox «подпись ≥ значение»: показывает подписи, отдаёт значения."""

    def __init__(self, parent, on_change, width=38):
        self.values = []
        self.on_change = on_change
        self.box = ttk.Combobox(parent, state="readonly", width=width)
        self.box.bind("<<ComboboxSelected>>", self._selected)

    def set(self, options, current, missing_label="не подключен"):
        """options: [(значение, подпись)]."""
        self.values = [value for value, _ in options]
        self.box["values"] = [label for _, label in options]
        if current in self.values:
            self.box.current(self.values.index(current))
        else:
            self.box.set(missing_label if current else "не выбран")

    def _selected(self, _event=None):
        index = self.box.current()
        if 0 <= index < len(self.values):
            self.on_change(self.values[index])


class Panel:
    REFRESH_MS = 3000
    SLOW_REFRESH_MS = 10000

    def __init__(self, root, actions):
        self.root = root
        self.actions = actions
        self.state = None
        self._refreshing = False
        self._closed = False

        setup_style(root)
        self.win = tk.Toplevel(root)
        self.win.title("MicSwitcher")
        self.win.configure(bg=BG)
        self.win.minsize(560, 460)
        self.win.protocol("WM_DELETE_WINDOW", self.close)

        self._build()
        self.request_refresh()
        self.win.after(self.REFRESH_MS, self._tick)

    # --- жизненный цикл -------------------------------------------------

    def show(self):
        self.win.deiconify()
        self.win.lift()
        self.win.focus_force()
        self.request_refresh()

    def close(self):
        self._closed = True
        self.win.destroy()

    @property
    def closed(self):
        return self._closed

    def _tick(self):
        if self._closed:
            return
        self.request_refresh()
        slow = self.state is not None and not self.state.get("watchable", True)
        self.win.after(self.SLOW_REFRESH_MS if slow else self.REFRESH_MS, self._tick)

    def request_refresh(self):
        """Снимок состояния - в фоне (там вызовы звука), отрисовка - в потоке tkinter."""
        if self._refreshing or self._closed:
            return
        self._refreshing = True

        def work():
            try:
                state = self.actions.state()
            except Exception:
                state = None
            self.actions.call_ui(self._apply, state)

        threading.Thread(target=work, daemon=True).start()

    def do(self, fn, *args):
        """Действие приложения в фоне, затем обновление панели."""

        def work():
            try:
                fn(*args)
            finally:
                self.actions.call_ui(self.request_refresh)

        threading.Thread(target=work, daemon=True).start()

    # --- построение -----------------------------------------------------

    def _card(self, parent, title):
        frame = ttk.Frame(parent, style="Card.TFrame", padding=12)
        frame.pack(fill="x", padx=12, pady=(12, 0))
        ttk.Label(frame, text=title, style="Title.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))
        frame.columnconfigure(1, weight=1)
        return frame

    def _build(self):
        header = ttk.Frame(self.win, padding=(16, 12, 16, 4))
        header.pack(fill="x")
        ttk.Label(header, text="MicSwitcher", style="Header.TLabel").pack(side="left")
        self.backend_label = ttk.Label(header, text="", style="HeaderMuted.TLabel")
        self.backend_label.pack(side="right")

        notebook = ttk.Notebook(self.win)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.notebook = notebook

        self._build_devices(self._tab(notebook, "Устройства"))
        self._build_profiles(self._tab(notebook, "Профили"))
        self._build_eq(self._tab(notebook, "Эквалайзер"))
        self._build_settings(self._tab(notebook, "Настройки"))

    @staticmethod
    def _tab(notebook, title):
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=title)
        return frame

    def _build_devices(self, tab):
        self.kind_widgets = {}
        for key, title in (("mic", "Микрофон"), ("out", "Вывод")):
            card = self._card(tab, title)
            current = ttk.Label(card, text="", style="Muted.TLabel")
            current.grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 6))
            slots = {}
            for slot in (1, 2):
                ttk.Label(card, text=f"{slot}:", style="Card.TLabel").grid(row=1 + slot, column=0, sticky="w", padx=(0, 8))
                choice = Choice(card, lambda value, k=key, s=slot: self.do(self.actions.select_device, k, s, value))
                choice.box.grid(row=1 + slot, column=1, sticky="ew", pady=2)
                button = ttk.Button(card, text="Включить", command=lambda k=key, s=slot: self.do(self.actions.activate_slot, k, s))
                button.grid(row=1 + slot, column=2, padx=(8, 0), pady=2)
                slots[slot] = choice
            toggle = ttk.Button(card, text="Переключить", style="Accent.TButton", command=lambda k=key: self.do(self.actions.toggle, k))
            toggle.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))
            self.kind_widgets[key] = {"current": current, "slots": slots, "toggle": toggle}

    def _build_profiles(self, tab):
        card = self._card(tab, "Профили")
        self.profile_list = tk.Listbox(
            card,
            height=7,
            bg=FIELD,
            fg=FG,
            selectbackground=ACCENT,
            highlightthickness=0,
            borderwidth=0,
            font=FONT,
            activestyle="none",
            exportselection=False,
        )
        self.profile_list.grid(row=1, column=0, columnspan=4, sticky="ew")
        self.profile_list.bind("<<ListboxSelect>>", lambda _e: self._profile_selected())
        self.profile_list.bind("<Double-Button-1>", lambda _e: self._profile_apply())
        self.profile_names = []

        buttons = ttk.Frame(card, style="Card.TFrame")
        buttons.grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))
        ttk.Button(buttons, text="Применить", style="Accent.TButton", command=self._profile_apply).pack(side="left")
        ttk.Button(buttons, text="Удалить", command=self._profile_delete).pack(side="left", padx=(8, 0))

        ttk.Label(card, text="Хоткей профиля:", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.profile_hotkey = ttk.Entry(card, width=24)
        self.profile_hotkey.grid(row=3, column=1, sticky="w", pady=(10, 0))
        self.profile_hotkey_record = ttk.Button(card, text="Записать", command=lambda: self._record(self.profile_hotkey, self.profile_hotkey_record))
        self.profile_hotkey_record.grid(row=3, column=2, padx=(8, 0), pady=(10, 0))
        ttk.Button(card, text="Назначить", command=self._profile_set_hotkey).grid(row=3, column=3, padx=(8, 0), pady=(10, 0))

        save = self._card(tab, "Сохранить текущие устройства и эквалайзер как профиль")
        ttk.Label(save, text="Название:", style="Card.TLabel").grid(row=1, column=0, sticky="w")
        self.profile_name = ttk.Entry(save)
        self.profile_name.grid(row=1, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(save, text="Сохранить", style="Accent.TButton", command=self._profile_save).grid(row=1, column=2, padx=(8, 0))

    def _build_eq(self, tab):
        status = self._card(tab, "Equalizer APO")
        self.eq_status = ttk.Label(status, text="", style="Muted.TLabel", wraplength=480, justify="left")
        self.eq_status.grid(row=1, column=0, columnspan=3, sticky="w")
        self.eq_button = ttk.Button(status, text="", command=lambda: self._open("eq"))
        self.eq_button.grid(row=2, column=0, sticky="w", pady=(8, 0))

        presets = self._card(tab, "Пресеты")
        self.eq_choices = {}
        for row, (key, title) in enumerate((("out", "Вывод"), ("mic", "Микрофон")), start=1):
            ttk.Label(presets, text=f"{title}:", style="Card.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 8))
            choice = Choice(presets, lambda value, k=key: self.do(self.actions.set_eq, k, value), width=30)
            choice.box.grid(row=row, column=1, sticky="w", pady=2)
            self.eq_choices[key] = choice

        noise = self._card(tab, "Шумодав RNNoise")
        self.rnnoise_status = ttk.Label(noise, text="", style="Muted.TLabel", wraplength=480, justify="left")
        self.rnnoise_status.grid(row=1, column=0, columnspan=3, sticky="w")
        ttk.Label(noise, text="Порог голоса:", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0), padx=(0, 8))
        self.threshold = Choice(noise, lambda value: self.do(self.actions.set_rnnoise_threshold, value), width=30)
        self.threshold.box.grid(row=2, column=1, sticky="w", pady=(8, 0))
        self.rnnoise_button = ttk.Button(
            noise, text="Установить шумодав", style="Accent.TButton", command=lambda: self.do(self.actions.install_noise_suppression)
        )
        self.rnnoise_button.grid(row=3, column=0, sticky="w", pady=(8, 0))

    def _build_settings(self, tab):
        keys = self._card(tab, "Горячие клавиши")
        self.hotkey_entries = {}
        for row, (key, title) in enumerate(HOTKEY_FIELDS, start=1):
            ttk.Label(keys, text=title, style="Card.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=2)
            entry = ttk.Entry(keys, width=24)
            entry.grid(row=row, column=1, sticky="w", pady=2)
            record = ttk.Button(keys, text="Записать")
            record.configure(command=lambda e=entry, b=record: self._record(e, b))
            record.grid(row=row, column=2, padx=(8, 0), pady=2)
            self.hotkey_entries[key] = entry
        ttk.Label(keys, text="Пусто - хоткей выключен.", style="Muted.TLabel").grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(keys, text="Сохранить хоткеи", style="Accent.TButton", command=self._save_hotkeys).grid(row=4, column=2, pady=(6, 0))

        general = self._card(tab, "Общее")
        self.autorun_var = tk.BooleanVar()
        ttk.Checkbutton(
            general,
            text="Запускать вместе с Windows",
            variable=self.autorun_var,
            command=lambda: self.do(self.actions.set_autorun, self.autorun_var.get()),
        ).grid(row=1, column=0, columnspan=3, sticky="w")
        self.watch_var = tk.BooleanVar()
        ttk.Checkbutton(
            general,
            text="Следить за подключением устройств",
            variable=self.watch_var,
            command=lambda: self.do(self.actions.set_auto_refresh, self.watch_var.get()),
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))

        files = ttk.Frame(general, style="Card.TFrame")
        files.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Button(files, text="Открыть config.json", command=lambda: self.do(self.actions.open_config)).pack(side="left")
        ttk.Button(files, text="Перечитать config.json", command=lambda: self.do(self.actions.reload_config)).pack(side="left", padx=(8, 0))
        ttk.Button(files, text="Открыть лог", command=lambda: self.do(self.actions.open_log)).pack(side="left", padx=(8, 0))

    # --- действия -------------------------------------------------------

    def _selected_profile(self):
        selection = self.profile_list.curselection()
        if not selection or selection[0] >= len(self.profile_names):
            return None
        return self.profile_names[selection[0]]

    def _profile_selected(self):
        name = self._selected_profile()
        profile = self._profile(name)
        if profile is not None:
            self._set_entry(self.profile_hotkey, profile["hotkey"], force=True)

    def _profile(self, name):
        if not self.state or name is None:
            return None
        return next((p for p in self.state["profiles"] if p["name"] == name), None)

    def _profile_apply(self):
        name = self._selected_profile()
        if name:
            self.do(self.actions.apply_profile, name)

    def _profile_delete(self):
        name = self._selected_profile()
        if name:
            self.do(self.actions.delete_profile, name)

    def _profile_set_hotkey(self):
        name = self._selected_profile()
        if name:
            self.do(self.actions.set_profile_hotkey, name, self.profile_hotkey.get().strip())

    def _profile_save(self):
        name = self.profile_name.get().strip()
        if name:
            self.do(self.actions.save_profile_named, name)
            self.profile_name.delete(0, "end")

    def _save_hotkeys(self):
        values = {key: entry.get().strip() for key, entry in self.hotkey_entries.items()}
        self.do(self.actions.set_hotkeys, values)

    def _record(self, entry, button):
        """Ждём сочетание в фоне: keyboard.read_hotkey блокирует поток."""
        button.configure(text="Нажми...", state="disabled")

        def work():
            try:
                combo = self.actions.record_hotkey()
            except Exception:
                combo = None
            self.actions.call_ui(done, combo)

        def done(combo):
            if self._closed:
                return
            button.configure(text="Записать", state="normal")
            if combo:
                self._set_entry(entry, combo, force=True)

        threading.Thread(target=work, daemon=True).start()

    def _open(self, what):
        if not self.state:
            return
        eq = self.state["eq"]
        if eq["found"]:
            self.do(self.actions.open_target, eq["dir"])
        else:
            self.do(self.actions.install_equalizer_apo)

    # --- отрисовка ------------------------------------------------------

    def _set_entry(self, entry, value, force=False):
        """Не перетираем поле, в котором пользователь сейчас печатает."""
        if not force and self.win.focus_get() is entry:
            return
        entry.delete(0, "end")
        entry.insert(0, value or "")

    def _apply(self, state):
        self._refreshing = False
        if self._closed or state is None:
            return
        previous = self.state
        first = previous is None
        self.state = state

        self.backend_label.configure(text=f"Звук: {state['backend']}")

        for kind in state["kinds"]:
            widgets = self.kind_widgets[kind["key"]]
            widgets["current"].configure(text=f"Сейчас: {kind['current']}")
            options = [(d["id"], d["label"]) for d in kind["devices"]]
            for slot, choice in widgets["slots"].items():
                choice.set(options, kind["slots"].get(slot, ""))
            hotkey = kind["hotkey"]
            widgets["toggle"].configure(text=f"Переключить  ({hotkey})" if hotkey else "Переключить")

        selected = self._selected_profile()
        self.profile_names = [p["name"] for p in state["profiles"]]
        self.profile_list.delete(0, "end")
        for profile in state["profiles"]:
            mark = "• " if profile["name"] == state["active_profile"] else "   "
            hotkey = f"   ({profile['hotkey']})" if profile["hotkey"] else ""
            self.profile_list.insert("end", f"{mark}{profile['name']}{hotkey}")
        if selected in self.profile_names:
            self.profile_list.selection_set(self.profile_names.index(selected))
        if first or not self.profile_name.get():
            self._set_entry(self.profile_name, state["profile_name_suggestion"])

        eq = state["eq"]
        if eq["found"]:
            self.eq_status.configure(text=f"Установлен: {eq['dir']}")
            self.eq_button.configure(text="Открыть папку Equalizer APO")
        else:
            self.eq_status.configure(
                text="Не установлен. Без него пресеты и шумодав не работают. Установщик скачается и запустится сам: "
                "в конце отметь в нём наушники и микрофоны и перезагрузи компьютер."
            )
            self.eq_button.configure(text="Установить Equalizer APO")
        for key, choice in self.eq_choices.items():
            info = eq["kinds"][key]
            choice.set(info["presets"], info["current"])
        if eq["rnnoise"]:
            self.rnnoise_status.configure(text=f"Плагин: {eq['rnnoise']}\nВключается пресетом микрофона «Шумодав» или «Голос + шумодав».")
            self.rnnoise_button.grid_remove()
        elif eq["found"]:
            self.rnnoise_status.configure(text="Плагин не установлен. Поставится сам при выборе пресета «Шумодав» или кнопкой.")
            self.rnnoise_button.grid()
        else:
            self.rnnoise_status.configure(text="Нужен Equalizer APO.")
            self.rnnoise_button.grid_remove()
        self.threshold.set(eq["thresholds"], eq["threshold"])

        if first or previous["hotkeys"] != state["hotkeys"]:
            for key, entry in self.hotkey_entries.items():
                self._set_entry(entry, state["hotkeys"].get(key, ""))
        self.autorun_var.set(state["autorun"])
        self.watch_var.set(state["auto_refresh"])
