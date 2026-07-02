import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from pathlib import Path
import threading
import traceback

from core.debug_log import get_logger
from core.template_manager import TemplateManager, TemplateNotFoundError

log = get_logger("gui")

WHITE     = "#FFFFFF"
BG        = "#F3F4F6"
PANEL     = "#FFFFFF"
DIVIDER   = "#E5E7EB"
TEXT      = "#111827"
TEXT_SEC  = "#6B7280"
TEXT_MUTE = "#9CA3AF"
ACCENT    = "#2563EB"
ACCENT_H  = "#1D4ED8"
ERR_RED   = "#DC2626"
WARN_AMB  = "#D97706"
INFO_BLU  = "#2563EB"
OK_GRN    = "#059669"


class ThesisFormatterGUI(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("論文格式修正工具")
        self.geometry("880x680")
        self.minsize(760, 560)
        self.configure(bg=BG)
        self._dragging = False
        self._status_open = False
        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background=BG, foreground=TEXT,
                        font=("Microsoft JhengHei", 10))
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("Primary.TButton",
                        background=ACCENT, foreground=WHITE,
                        font=("Microsoft JhengHei", 11, "bold"),
                        padding=(20, 8), relief="flat")
        style.map("Primary.TButton",
                  background=[("active", ACCENT_H), ("disabled", "#C7C7CC")])
        style.configure("Ghost.TButton",
                        background=WHITE, foreground=TEXT,
                        font=("Microsoft JhengHei", 10),
                        padding=(12, 5), relief="solid", borderwidth=1)
        style.map("Ghost.TButton", background=[("active", BG)])
        style.configure("Treeview",
                        background=WHITE, foreground=TEXT,
                        font=("Microsoft JhengHei", 9),
                        rowheight=28, fieldbackground=WHITE)
        style.configure("Treeview.Heading",
                        font=("Microsoft JhengHei", 10, "bold"),
                        background=BG, foreground=TEXT, relief="flat")
        style.map("Treeview.Heading", background=[("active", "#E5E7EB")])
        style.map("Treeview",
                  background=[("selected", "#DBEAFE")],
                  foreground=[("selected", ACCENT)])
        style.configure("Horizontal.TProgressbar",
                        background=ACCENT, troughcolor="#E5E7EB")

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = ttk.Frame(self, padding="24 0 24 24")
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)

        self._build_header(root)
        self._build_cards(root)
        self._build_action(root)
        self._build_results(root)
        self._build_status(root)

    # ── Header ──────────────────────────────────────────────
    def _build_header(self, parent):
        h = tk.Frame(parent, bg=BG, highlightthickness=0)
        h.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        h.columnconfigure(0, weight=1)
        tk.Label(h, text="論文格式修正工具", bg=BG, fg=TEXT,
                 font=("Microsoft JhengHei", 22, "bold"),
                 anchor="w").grid(row=0, column=0, sticky="w")
        tk.Label(h, text="自動檢查並修正 Word 論文格式問題",
                 bg=BG, fg=TEXT_SEC, font=("Microsoft JhengHei", 10),
                 anchor="w").grid(row=1, column=0, sticky="w", pady=(2, 0))
        tk.Frame(h, bg=DIVIDER, height=1, bd=0,
                 highlightthickness=0).grid(row=2, column=0, sticky="ew", pady=(14, 0))

    # ── Cards ───────────────────────────────────────────────
    def _build_cards(self, parent):
        row = ttk.Frame(parent)
        row.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        row.columnconfigure(0, weight=3)
        row.columnconfigure(1, weight=2)
        self._build_file_card(row)
        self._build_template_card(row)

    def _build_file_card(self, parent):
        card = tk.Frame(parent, bg=PANEL, bd=0, highlightthickness=0)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        card.columnconfigure(0, weight=1)

        tk.Label(card, text="選擇論文檔案", bg=PANEL, fg=TEXT,
                 font=("Microsoft JhengHei", 13, "bold"),
                 anchor="w").grid(row=0, column=0, sticky="w", padx=16, pady=(16, 2))
        tk.Label(card, text="載入需要修正格式的 Word 文件", bg=PANEL, fg=TEXT_SEC,
                 font=("Microsoft JhengHei", 9),
                 anchor="w").grid(row=1, column=0, sticky="w", padx=16, pady=(0, 12))

        self.drop_frame = tk.Frame(card, bg="#FAFAFA",
                                   highlightbackground=DIVIDER,
                                   highlightthickness=2, bd=0, height=60)
        self.drop_frame.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 12))
        self.drop_frame.grid_propagate(False)
        self.drop_frame.columnconfigure(0, weight=1)
        self.drop_label = tk.Label(self.drop_frame,
                text="將 .docx 檔案拖曳至此處，或點擊下方按鈕選擇",
                bg="#FAFAFA", fg=TEXT_MUTE,
                font=("Microsoft JhengHei", 10), justify="center")
        self.drop_label.place(relx=0.5, rely=0.5, anchor="center")
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_frame.dnd_bind("<<DragEnter>>", self._on_drag_enter)
        self.drop_frame.dnd_bind("<<DragLeave>>", self._on_drag_leave)
        self.drop_frame.dnd_bind("<<DragPos>>", self._on_drag_pos)

        erow = tk.Frame(card, bg=PANEL)
        erow.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 16))
        erow.columnconfigure(0, weight=1)
        self.file_entry = tk.Entry(
            erow, font=("Consolas", 10), bd=1, relief="solid",
            highlightbackground=DIVIDER, bg=WHITE, fg=TEXT, insertbackground=ACCENT)
        self.file_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 8))
        ttk.Button(erow, text="瀏覽…", style="Ghost.TButton",
                   command=self.select_file).pack(side="right")

    def _build_template_card(self, parent):
        card = tk.Frame(parent, bg=PANEL, bd=0, highlightthickness=0)
        card.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        card.columnconfigure(0, weight=1)
        tk.Label(card, text="格式模板", bg=PANEL, fg=TEXT,
                 font=("Microsoft JhengHei", 13, "bold"),
                 anchor="w").grid(row=0, column=0, sticky="w", padx=16, pady=(16, 2))
        tk.Label(card, text="選擇適用的論文格式規範", bg=PANEL, fg=TEXT_SEC,
                 font=("Microsoft JhengHei", 9),
                 anchor="w").grid(row=1, column=0, sticky="w", padx=16, pady=(0, 12))

        self.template_var = tk.StringVar()
        self.template_combo = ttk.Combobox(
            card, textvariable=self.template_var,
            state="readonly", font=("Microsoft JhengHei", 10))
        self.template_combo.grid(row=2, column=0, sticky="ew", padx=16, ipady=5)
        self._refresh_template_list()
        self.template_combo.bind("<<ComboboxSelected>>", self._on_template_change)

        self._template_info_label = tk.Label(card, text="", bg=PANEL, fg=TEXT_SEC,
            font=("Microsoft JhengHei", 8), anchor="w", wraplength=220)
        self._template_info_label.grid(row=3, column=0, sticky="ew", padx=16, pady=(4, 0))

        ttk.Button(card, text="管理模板…", style="Ghost.TButton",
                   command=self._open_template_manager).grid(
            row=4, column=0, sticky="ew", padx=16, pady=(10, 16))

    # ── Action ──────────────────────────────────────────────
    def _build_action(self, parent):
        row = ttk.Frame(parent)
        row.grid(row=2, column=0, sticky="ew", pady=(0, 16))
        row.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(row, mode="indeterminate",
                                        length=200, style="Horizontal.TProgressbar")
        self.progress.grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.progress.grid_remove()
        ttk.Button(row, text="清除結果", style="Ghost.TButton",
                   command=self.clear_results).grid(row=0, column=1, padx=(0, 8))
        self.run_btn = ttk.Button(row, text="開始格式修正",
                                  style="Primary.TButton",
                                  command=self.run_formatter)
        self.run_btn.grid(row=0, column=2)

    # ── Results ─────────────────────────────────────────────
    def _build_results(self, parent):
        card = tk.Frame(parent, bg=PANEL, bd=0, highlightthickness=0)
        card.grid(row=3, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)

        self._output_label = tk.Label(card, text="", bg=PANEL, fg=TEXT_SEC,
            font=("Microsoft JhengHei", 9), anchor="w")
        self._output_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=18, pady=(10, 0))

        columns = ("severity", "type", "message")
        self.tree = ttk.Treeview(card, columns=columns, show="headings",
                                 height=14, selectmode="browse")
        self.tree.heading("severity", text="等級", anchor="center")
        self.tree.heading("type", text="類型", anchor="w")
        self.tree.heading("message", text="說明", anchor="w")
        self.tree.column("severity", width=80, anchor="center", minwidth=70)
        self.tree.column("type", width=150, minwidth=90)
        self.tree.column("message", width=600, minwidth=200)
        self.tree.grid(row=1, column=0, sticky="nsew", padx=16, pady=16)
        sb = ttk.Scrollbar(card, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=1, column=1, sticky="ns", pady=16)
        self.tree.tag_configure("error", foreground=ERR_RED)
        self.tree.tag_configure("warning", foreground=WARN_AMB)
        self.tree.tag_configure("info", foreground=INFO_BLU)
        self.tree.tag_configure("saved", foreground=OK_GRN)
        self.tree.tag_configure("even", background="#FAFAFA")
        self.tree.tag_configure("odd", background=WHITE)

    # ── Collapsible Status ──────────────────────────────────
    def _build_status(self, parent):
        self._status_frame = tk.Frame(parent, bg=BG, highlightthickness=0)
        self._status_frame.columnconfigure(0, weight=1)
        self._status_frame.columnconfigure(1, weight=0)

        self._status_toggle = tk.Label(
            self._status_frame, text="▶", bg=BG, fg=TEXT_SEC,
            font=("Microsoft JhengHei", 9), cursor="hand2", padx=4)
        self._status_toggle.grid(row=0, column=0, sticky="w", pady=(6, 0))
        self._status_toggle.bind("<Button-1>", self._toggle_status)
        self._status_toggle.grid_remove()

        self.status_label = tk.Label(
            self._status_frame, text="", bg=BG, fg=TEXT_SEC,
            font=("Microsoft JhengHei", 9), anchor="w")
        self.status_label.grid(row=0, column=0, sticky="w", pady=(6, 0))
        self.status_label.grid_remove()

    def _toggle_status(self, event=None):
        if self._status_open:
            self.status_label.grid_remove()
            self._status_open = False
            self._status_toggle.config(text="▶")
        else:
            self.status_label.grid()
            self._status_open = True
            self._status_toggle.config(text="▼")

    def _show_status(self, text, fg=TEXT_SEC):
        self.status_label.config(text=text, fg=fg)
        if text:
            self._status_frame.grid(row=4, column=0, sticky="ew", pady=(0, 0))
            self._status_toggle.grid()
            if not self._status_open:
                self.status_label.grid()
                self._status_open = True
                self._status_toggle.config(text="▼")
        else:
            self._status_frame.grid_remove()
            self._status_open = False

    # ── Drag feedback ───────────────────────────────────────
    def _on_drag_enter(self, event):
        self._dragging = True
        self.drop_frame.configure(bg="#DBEAFE", highlightbackground=ACCENT)
        self.drop_label.configure(bg="#DBEAFE", fg=ACCENT, text="放開以選取此檔案")

    def _on_drag_leave(self, event):
        self._dragging = False
        self.drop_frame.configure(bg="#FAFAFA", highlightbackground=DIVIDER)
        self.drop_label.configure(bg="#FAFAFA", fg=TEXT_MUTE,
                                  text="將 .docx 檔案拖曳至此處，或點擊下方按鈕選擇")

    def _on_drag_pos(self, event):
        if not self._dragging:
            self._on_drag_enter(event)

    def _on_drop(self, event):
        self._on_drag_leave(event)
        files = self.tk.splitlist(event.data)
        if files:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, files[0])

    # ── Template ────────────────────────────────────────────
    def _refresh_template_list(self):
        templates = TemplateManager.get_all_templates()
        self._template_map = {}
        names = []
        for t in templates:
            display = t["name"]
            if t.get("builtin"):
                display += " (內建)"
            else:
                display += " (自訂)"
            names.append(display)
            self._template_map[display] = t["key"]
        self.template_combo["values"] = names
        if not self.template_var.get() or self.template_var.get() not in names:
            self.template_var.set(names[0] if names else "")

    def _on_template_change(self, event=None):
        display = self.template_var.get()
        if not display:
            self._template_info_label.config(text="")
            return
        template_key = self._template_map.get(display, "")
        try:
            info = TemplateManager.get_template_info(template_key)
            desc = info.get('description', '')
            src = "內建" if info.get('builtin') else "自訂"
            text = f"✓ {info['name']} ({src})"
            if desc:
                text += f" — {desc}"
            self._template_info_label.config(text=text, fg=OK_GRN)
        except Exception:
            self._template_info_label.config(text=f"選取: {display}", fg=TEXT_SEC)

    def _open_template_manager(self):
        from gui.template_dialog import TemplateManagerDialog
        dlg = TemplateManagerDialog(self)
        self.wait_window(dlg)
        self._refresh_template_list()

    # ── Handlers ────────────────────────────────────────────
    def select_file(self):
        path = filedialog.askopenfilename(
            title="選擇論文檔案",
            filetypes=[("Word 文件", "*.docx"), ("所有檔案", "*.*")])
        if path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)

    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._show_status("")

    def run_formatter(self):
        input_path = self.file_entry.get().strip()
        if not input_path or not Path(input_path).exists():
            messagebox.showerror("錯誤", "請選擇有效的 .docx 檔案")
            return
        self.clear_results()
        self.run_btn.config(state="disabled")
        self.progress.grid()
        self.progress.start(12)
        self._show_status("正在分析與修正中…", fg=TEXT_SEC)
        template_key = self.template_var.get()
        thread = threading.Thread(target=self._process_document,
                                  args=(input_path, template_key))
        thread.daemon = True
        thread.start()

    def _process_document(self, input_path, template_display):
        log.info("=== _process_document ===")
        from core.rule_engine import RuleEngine
        from core.analyzer import FormatAnalyzer
        from core.fixer import FormatFixer
        input_path = Path(input_path)
        output_path = input_path.parent / f"{input_path.stem}_格式修正{input_path.suffix}"
        try:
            template_key = self._template_map.get(template_display, template_display)
            log.info("template_display=%s, resolved key=%s", template_display, template_key)

            rule_path = TemplateManager.get_template_path(template_key)
            rule_file = rule_path.name
            log.info("Template rule file: %s", rule_file)

            analyzer = FormatAnalyzer(rule_file)
            issues = analyzer.analyze(str(input_path))
            log.info("Analyzer found %d issues", len(issues))

            fixer = FormatFixer(rule_file)
            fixes = fixer.fix_document(str(input_path), str(output_path))
            self.after(0, self._update_results, issues, fixes, output_path.name)
        except Exception as e:
            log.error(traceback.format_exc())
            self.after(0, lambda e=e, td=template_display: messagebox.showerror(
                "處理錯誤", f"模板「{td}」處理失敗:\n{e}"))
        finally:
            self.after(0, self._reset_ui)

    def _reset_ui(self):
        self.progress.stop()
        self.progress.grid_remove()
        self.run_btn.config(state="normal")

    def _update_results(self, issues, fixes, output_name):
        for item in self.tree.get_children():
            self.tree.delete(item)
        template_display = self.template_var.get() or "預設"
        self._output_label.config(
            text=f"模板: {template_display}  |  輸出: {output_name}")
        idx = 0
        for r in issues:
            sev = r.get("severity", "info")
            tag_base = sev if sev in ("error", "warning", "info") else "info"
            row_tag = "even" if idx % 2 == 0 else "odd"
            sev_display = {"error": "錯誤", "warning": "警告", "info": "提示"}.get(sev, sev)
            self.tree.insert("", tk.END,
                values=(sev_display, r.get("type", ""), r.get("message", "")),
                tags=(tag_base, row_tag))
            idx += 1
        for r in fixes:
            row_tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert("", tk.END,
                values=("修正", r.get("type", ""), r.get("message", "")),
                tags=("saved", row_tag))
            idx += 1
        self._show_status(f"完成！已儲存至 {output_name}", fg=OK_GRN)


if __name__ == "__main__":
    app = ThesisFormatterGUI()
    app.mainloop()
