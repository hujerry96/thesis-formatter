import copy
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path

from core.template_manager import TemplateManager, TemplateExistsError, TemplateNotFoundError, BuiltinTemplateError

BG        = "#F3F4F6"
CARD_BG   = "#FFFFFF"
ACCENT    = "#2563EB"
TEXT      = "#111827"
TEXT_SEC  = "#6B7280"
BORDER    = "#E5E7EB"
OK_GRN    = "#059669"


class ScrollableTabFrame(tk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=BG, **kw)
        self._canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        vs = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self._canvas.yview)
        self.inner = tk.Frame(self._canvas, bg=BG)
        self.inner.bind("<Configure>", lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox("all")))
        self._win = self._canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self._canvas.bind("<Configure>", lambda e: self._canvas.itemconfigure(
            self._win, width=e.width))
        self._canvas.configure(yscrollcommand=vs.set)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vs.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.bind("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, e):
        self._canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")


class TemplateEditorDialog(tk.Toplevel):
    def __init__(self, parent, template_key: str, template_data: dict, is_new: bool = False):
        super().__init__(parent)
        self.withdraw()
        self.template_key = template_key
        self.template_data = copy.deepcopy(template_data)
        self.is_new = is_new
        self.result = None
        self.title("新增模板" if is_new else "編輯模板")
        self.geometry("680x720")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.transient(parent)
        self._fonts = self._get_font_list()
        try:
            self._build_ui()
            self._load_data()
        except Exception as e:
            import traceback as _tb
            _tb.print_exc()
            messagebox.showerror("錯誤", f"無法載入編輯器:\n{e}", parent=self)
            self.destroy()
            return
        self.grab_set()
        self.deiconify()

    def _get_font_list(self):
        import tkinter.font as tkfont
        return sorted(tkfont.families())

    def _field(self, parent, label, row, width=16):
        tk.Label(parent, text=label, bg=CARD_BG, fg=TEXT,
                 font=("Microsoft JhengHei UI", 9), width=14, anchor="w").grid(
            row=row, column=0, sticky="w", pady=2)
        e = tk.Entry(parent, font=("Consolas", 10), bd=1, relief="solid", width=width)
        e.grid(row=row, column=1, sticky="ew", padx=(8, 0), ipady=3, pady=2)
        return e

    def _combo(self, parent, label, row, values, width=20):
        tk.Label(parent, text=label, bg=CARD_BG, fg=TEXT,
                 font=("Microsoft JhengHei UI", 9), width=14, anchor="w").grid(
            row=row, column=0, sticky="w", pady=2)
        var = tk.StringVar()
        cb = ttk.Combobox(parent, textvariable=var, values=values, state="readonly", width=width)
        cb.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=2)
        return var

    def _check(self, parent, label, row):
        var = tk.BooleanVar()
        tk.Checkbutton(parent, text=label, variable=var, bg=CARD_BG, fg=TEXT,
                       font=("Microsoft JhengHei UI", 9), selectcolor=CARD_BG).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=1)
        return var

    def _card(self, parent, title, row):
        c = tk.LabelFrame(parent, text=f" {title} ", bg=CARD_BG, fg=TEXT,
                          font=("Microsoft JhengHei UI", 10, "bold"),
                          highlightbackground=BORDER, highlightthickness=1, bd=0,
                          padx=12, pady=8)
        c.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        c.columnconfigure(1, weight=1)
        return c

    def _build_ui(self):
        # ── 頂部：模板名稱 + 描述 ──
        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", padx=12, pady=(8, 4))
        top.columnconfigure(1, weight=1)
        tk.Label(top, text="模板名稱", bg=BG, fg=TEXT, font=("Microsoft JhengHei UI", 9)).grid(row=0, column=0, sticky="w")
        self.entry_name = tk.Entry(top, font=("Microsoft JhengHei UI", 10), bd=1, relief="solid")
        self.entry_name.grid(row=0, column=1, sticky="ew", padx=(8, 0), ipady=3, pady=(0, 4))
        tk.Label(top, text="描述", bg=BG, fg=TEXT, font=("Microsoft JhengHei UI", 9)).grid(row=1, column=0, sticky="w")
        self.entry_desc = tk.Entry(top, font=("Microsoft JhengHei UI", 10), bd=1, relief="solid")
        self.entry_desc.grid(row=1, column=1, sticky="ew", padx=(8, 0), ipady=3)

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=4)

        # ── Tab 1: 紙張與頁面 ──
        tab1 = ScrollableTabFrame(notebook)
        notebook.add(tab1, text=" 紙張與頁面 ")
        f1 = tab1.inner
        f1.columnconfigure(1, weight=1)

        c_paper = self._card(f1, "紙張", 0)
        self.paper_size = self._combo(c_paper, "尺寸", 0, ["A4", "A3", "Letter", "B5"])
        self.paper_orient = self._combo(c_paper, "方向", 1, ["portrait", "landscape"])

        c_body = self._card(f1, "內文邊距 (cm)", 1)
        self.page_top = self._field(c_body, "上", 0)
        self.page_bottom = self._field(c_body, "下", 1)
        self.page_left = self._field(c_body, "左", 2)
        self.page_right = self._field(c_body, "右", 3)

        c_cover = self._card(f1, "封面邊距 (cm)", 2)
        self.cover_top = self._field(c_cover, "上", 0)
        self.cover_bottom = self._field(c_cover, "下", 1)
        self.cover_left = self._field(c_cover, "左", 2)
        self.cover_right = self._field(c_cover, "右", 3)

        c_bind = self._card(f1, "裝訂", 3)
        self.bind_extra = self._field(c_bind, "額外左邊距 (cm)", 0)

        # ── Tab 2: 行距與段落 ──
        tab3 = ScrollableTabFrame(notebook)
        notebook.add(tab3, text=" 行距與段落 ")
        f3 = tab3.inner
        f3.columnconfigure(1, weight=1)

        c_ls = self._card(f3, "行距", 0)
        self.ls_type = self._combo(c_ls, "類型", 0, ["multiple", "exact", "at_least"])
        self.ls_value = self._field(c_ls, "數值", 1)

        c_ps = self._card(f3, "段落間距 (pt)", 1)
        self.ps_before = self._field(c_ps, "段前", 0)
        self.ps_after = self._field(c_ps, "段後", 1)

        c_indent = self._card(f3, "首行縮排", 2)
        self.indent_enabled = self._check(c_indent, "啟用", 0)
        self.indent_cm = self._field(c_indent, "數值 (cm)", 1)
        self.indent_chars = self._field(c_indent, "字元數", 2)

        # ── Tab 4: 標題樣式（使用表格佈局）──
        tab4 = ScrollableTabFrame(notebook)
        notebook.add(tab4, text=" 標題樣式 ")
        f4 = tab4.inner
        f4.columnconfigure(0, weight=1)

        heading_labels = [("摘要/致謝/目錄", "title"), ("標題 1", "heading1"),
                          ("標題 2", "heading2"), ("標題 3", "heading3"), ("標題 4", "heading4")]

        self.align_map_display = {'center': '置中', 'left': '靠左', 'right': '靠右', 'justify': '左右對齊'}

        self.heading_entries = {}
        for hi, (lbl, key) in enumerate(heading_labels):
            c_h = self._card(f4, lbl, hi)
            entries = {}
            row = 0

            # 第一行：中文字體 + 英文字體
            frow = tk.Frame(c_h, bg=CARD_BG)
            frow.grid(row=row, column=0, columnspan=2, sticky="ew", pady=1)
            frow.columnconfigure(1, weight=1)
            frow.columnconfigure(3, weight=1)
            tk.Label(frow, text="中文字體", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 8), anchor="w").grid(row=0, column=0)
            w = ttk.Combobox(frow, values=self._fonts, state="readonly", width=16)
            w.grid(row=0, column=1, sticky="ew", padx=(0, 8))
            entries['zh_font'] = w
            tk.Label(frow, text="英文字體", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 8), anchor="w").grid(row=0, column=2)
            ew = ttk.Combobox(frow, values=self._fonts, state="readonly", width=16)
            ew.grid(row=0, column=3, sticky="ew")
            entries['en_font'] = ew

            # 第二行：大小 + 粗體 + 對齊
            arow = tk.Frame(c_h, bg=CARD_BG)
            arow.grid(row=row+1, column=0, columnspan=2, sticky="ew", pady=1)
            arow.columnconfigure(1, weight=1)
            arow.columnconfigure(3, weight=1)
            arow.columnconfigure(5, weight=1)
            tk.Label(arow, text="大小", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 8), width=4, anchor="w").grid(row=0, column=0)
            e = tk.Entry(arow, font=("Consolas", 9), bd=1, relief="solid", width=6)
            e.grid(row=0, column=1, sticky="ew", padx=(0, 8), ipady=2)
            entries['font_size'] = e
            tk.Label(arow, text="粗體", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 8), width=4, anchor="w").grid(row=0, column=2)
            bv = ttk.Combobox(arow, values=["是", "否"], state="readonly", width=5)
            bv.grid(row=0, column=3, sticky="ew", padx=(0, 8))
            entries['bold'] = bv
            tk.Label(arow, text="對齊", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 8), width=4, anchor="w").grid(row=0, column=4)
            av = ttk.Combobox(arow, values=["center", "left", "right", "justify"], state="readonly", width=10)
            av.grid(row=0, column=5, sticky="ew")
            entries['align'] = av

            # 第三行：段前 + 段後
            brow = tk.Frame(c_h, bg=CARD_BG)
            brow.grid(row=row+2, column=0, columnspan=2, sticky="ew", pady=1)
            brow.columnconfigure(1, weight=1)
            brow.columnconfigure(3, weight=1)
            tk.Label(brow, text="段前 (pt)", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 8), anchor="w").grid(row=0, column=0)
            e1 = tk.Entry(brow, font=("Consolas", 9), bd=1, relief="solid", width=6)
            e1.grid(row=0, column=1, sticky="w", padx=(0, 16), ipady=2)
            entries['before'] = e1
            tk.Label(brow, text="段後 (pt)", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 8), anchor="w").grid(row=0, column=2)
            e2 = tk.Entry(brow, font=("Consolas", 9), bd=1, relief="solid", width=6)
            e2.grid(row=0, column=3, sticky="w", ipady=2)
            entries['after'] = e2

            # 第四行：段前分頁
            entries['page_break'] = self._check(c_h, "段前分頁", row+3)

            self.heading_entries[key] = entries

        # ── Tab 5: 內文·頁碼·圖表 ──
        tab5 = ScrollableTabFrame(notebook)
        notebook.add(tab5, text=" 內文·頁碼·圖表 ")
        f5 = tab5.inner
        f5.columnconfigure(1, weight=1)

        c_body = self._card(f5, "內文樣式", 0)
        self.body_zh = self._combo(c_body, "中文字體", 0, self._fonts)
        self.body_en = self._combo(c_body, "英文字體", 1, self._fonts)
        self.body_size = self._field(c_body, "大小", 2)
        self.body_align = self._combo(c_body, "對齊", 3, ["left", "center", "right", "justify"])

        c_fm = self._card(f5, "前置頁頁碼", 1)
        self.fm_style = self._combo(c_fm, "格式", 0, ["roman_lower", "roman_upper", "arabic", "none"])
        self.fm_pos = self._combo(c_fm, "位置", 1, ["bottom", "top"])
        self.fm_align = self._combo(c_fm, "對齊", 2, ["center", "left", "right"])

        c_bm = self._card(f5, "正文頁碼", 2)
        self.bm_style = self._combo(c_bm, "格式", 0, ["arabic", "roman_lower", "roman_upper", "none"])
        self.bm_pos = self._combo(c_bm, "位置", 1, ["bottom", "top"])
        self.bm_align = self._combo(c_bm, "對齊", 2, ["center", "left", "right"])
        self.bm_start = self._field(c_bm, "起始頁碼", 3)

        c_fig = self._card(f5, "圖片格式", 3)
        self.fig_num = self._combo(c_fig, "編號格式", 0, ["dot", "dash", "underscore"])
        self.fig_caption = self._check(c_fig, "標題在下方", 1)

        c_tab = self._card(f5, "表格格式", 4)
        self.tab_num = self._combo(c_tab, "編號格式", 0, ["dot", "dash", "underscore"])
        self.tab_caption = self._check(c_tab, "標題在上方", 1)

        # ── 按鈕 ──
        btn = tk.Frame(self, bg=BG)
        btn.pack(fill="x", padx=8, pady=(0, 8))
        tk.Button(btn, text="取消", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 10),
                  bd=1, relief="solid", padx=20, pady=4, command=self.destroy).pack(side="right", padx=(8, 0))
        tk.Button(btn, text="儲存", bg=ACCENT, fg="white", font=("Microsoft JhengHei UI", 10, "bold"),
                  bd=0, padx=20, pady=4, command=self._save).pack(side="right")

    def _load_data(self):
        d = self.template_data
        tpl = d.get('template', {})
        if self.entry_name:
            self.entry_name.delete(0, tk.END)
            self.entry_name.insert(0, tpl.get('name', ''))
        if self.entry_desc:
            self.entry_desc.delete(0, tk.END)
            self.entry_desc.insert(0, tpl.get('description', ''))

        paper = d.get('paper', {})
        self.paper_size.set(paper.get('size', 'A4'))
        self.paper_orient.set(paper.get('orientation', 'portrait'))

        bm = d.get('page_margins', {}).get('body', {})
        self.page_top.insert(0, str(bm.get('top', 2.5)))
        self.page_bottom.insert(0, str(bm.get('bottom', 2.5)))
        self.page_left.insert(0, str(bm.get('left', 2.7)))
        self.page_right.insert(0, str(bm.get('right', 2.7)))

        cm = d.get('page_margins', {}).get('cover', {})
        self.cover_top.insert(0, str(cm.get('top', 2.54)))
        self.cover_bottom.insert(0, str(cm.get('bottom', 2.54)))
        self.cover_left.insert(0, str(cm.get('left', 3.18)))
        self.cover_right.insert(0, str(cm.get('right', 3.18)))

        bind = d.get('binding', {})
        self.bind_extra.insert(0, str(bind.get('extra_left', 0)))

        # fonts section kept for compatibility but no longer shown in UI

        ls = d.get('line_spacing', {})
        self.ls_type.set(ls.get('type', 'multiple'))
        self.ls_value.insert(0, str(ls.get('value', 1.5)))

        ps = d.get('paragraph_spacing', {})
        self.ps_before.insert(0, str(ps.get('before', 0)))
        self.ps_after.insert(0, str(ps.get('after', 0)))

        fi = d.get('first_line_indent', {})
        self.indent_enabled.set(fi.get('enabled', True))
        self.indent_cm.insert(0, str(fi.get('value_cm', 0.75)))
        self.indent_chars.insert(0, str(fi.get('value_chars', 2)))

        hs = d.get('heading_styles', {})
        for hkey, entries in self.heading_entries.items():
            cfg = hs.get(hkey, {})
            entries['zh_font'].set(cfg.get('chinese_font', cfg.get('font_name', '')))
            entries['en_font'].set(cfg.get('english_font', 'Times New Roman'))
            entries['font_size'].insert(0, str(cfg.get('font_size', 12)))
            entries['bold'].set("是" if cfg.get('bold', True) else "否")
            entries['align'].set(cfg.get('align', 'left'))
            entries['before'].insert(0, str(cfg.get('before', 0)))
            entries['after'].insert(0, str(cfg.get('after', 0)))
            entries['page_break'].set(cfg.get('page_break_before', False))

        bs = d.get('body_style', {})
        self.body_zh.set(bs.get('font_name', ''))
        self.body_en.set(bs.get('english_font', 'Times New Roman'))
        self.body_size.insert(0, str(bs.get('font_size', 12)))
        self.body_align.set(bs.get('align', 'left'))

        pn = d.get('page_number', {})
        fm = pn.get('front_matter', {})
        self.fm_style.set(fm.get('style', 'roman_lower'))
        self.fm_pos.set(fm.get('position', 'bottom'))
        self.fm_align.set(fm.get('align', 'center'))
        bd = pn.get('body', {})
        self.bm_style.set(bd.get('style', 'arabic'))
        self.bm_pos.set(bd.get('position', 'bottom'))
        self.bm_align.set(bd.get('align', 'center'))
        self.bm_start.insert(0, str(bd.get('start_from', 1)))

        fig = d.get('figures', {})
        self.fig_num.set(fig.get('numbering', 'dot'))
        self.fig_caption.set(fig.get('caption_below', True))

        tab = d.get('tables', {})
        self.tab_num.set(tab.get('numbering', 'dot'))
        self.tab_caption.set(tab.get('caption_above', True))

    def _find_entry(self, label):
        for w in self.winfo_toplevel().winfo_children():
            for child in self._find_all_entries(w):
                try:
                    if child.cget("text") == label:
                        grid = child.grid_info()
                        row = int(grid["row"])
                        col = int(grid["column"])
                        parent = child.master
                        for sibling in parent.winfo_children():
                            sg = sibling.grid_info()
                            if sg and int(sg.get("row", -1)) == row and int(sg.get("column", -1)) == col + 1:
                                if isinstance(sibling, tk.Entry):
                                    return sibling
                except: pass
        return None

    def _find_all_entries(self, widget):
        entries = []
        try:
            for child in widget.winfo_children():
                if isinstance(child, tk.Entry):
                    entries.append(child)
                entries.extend(self._find_all_entries(child))
        except: pass
        return entries

    def _fi(self, entry, d=0):
        try: return int(entry.get().strip() or d)
        except ValueError: return d

    def _ff(self, entry, d=0.0):
        try: return float(entry.get().strip() or d)
        except ValueError: return d

    def _save(self):
        name = self.entry_name.get().strip() if self.entry_name else ''
        if not name:
            messagebox.showerror("驗證錯誤", "模板名稱不能為空白", parent=self)
            return

        desc = self.entry_desc.get().strip() if self.entry_desc else ''
        d = self.template_data

        d['template'] = {'name': name, 'description': desc}

        d['paper'] = {
            'size': self.paper_size.get(),
            'width_cm': 21.0 if self.paper_size.get() == 'A4' else 29.7,
            'height_cm': 29.7 if self.paper_size.get() == 'A4' else 21.0,
            'orientation': self.paper_orient.get(),
        }

        d.setdefault('page_margins', {})
        d['page_margins']['body'] = {
            'top': self._ff(self.page_top, 2.5), 'bottom': self._ff(self.page_bottom, 2.5),
            'left': self._ff(self.page_left, 2.7), 'right': self._ff(self.page_right, 2.7)}
        d['page_margins']['cover'] = {
            'top': self._ff(self.cover_top, 2.54), 'bottom': self._ff(self.cover_bottom, 2.54),
            'left': self._ff(self.cover_left, 3.18), 'right': self._ff(self.cover_right, 3.18)}

        d['binding'] = {'extra_left': self._ff(self.bind_extra, 0)}

        # 保留 fonts 區塊（相容性，不再顯示於 UI）
        d['fonts'] = d.get('fonts', {
            'chinese': {'name': '新細明體', 'size': 12},
            'chinese_title': {'name': '標楷體', 'size': 12},
            'english': {'name': 'Times New Roman', 'size': 12},
            'heading_number': {'name': '標楷體'},
        })

        d['line_spacing'] = {'type': self.ls_type.get(), 'value': self._ff(self.ls_value, 1.5)}
        d['paragraph_spacing'] = {'before': self._fi(self.ps_before, 0), 'after': self._fi(self.ps_after, 0)}
        d['first_line_indent'] = {
            'enabled': self.indent_enabled.get(),
            'value_cm': self._ff(self.indent_cm, 0.75),
            'value_chars': self._fi(self.indent_chars, 2),
        }

        old_hs = d.get('heading_styles', {})
        d['heading_styles'] = {}
        for hkey, entries in self.heading_entries.items():
            old = old_hs.get(hkey, {})
            zh = entries['zh_font'].get()
            en = entries['en_font'].get()
            d['heading_styles'][hkey] = {
                'font_size': self._fi(entries['font_size'], 12),
                'bold': entries['bold'].get() == "是",
                'align': entries['align'].get(),
                'font_name': zh,
                'chinese_font': zh,
                'english_font': en,
                'before': self._fi(entries['before'], 0),
                'after': self._fi(entries['after'], 0),
                'keep_with_next': old.get('keep_with_next', True),
                'page_break_before': entries['page_break'].get(),
                'line_spacing': old.get('line_spacing', 1.5),
                'numbering': old.get('numbering', {}),
            }

        d['body_style'] = {
            'font_name': self.body_zh.get(),
            'english_font': self.body_en.get(),
            'font_size': self._fi(self.body_size, 12),
            'align': self.body_align.get(),
            'first_line_indent_chars': d.get('first_line_indent', {}).get('value_chars', 2),
            'line_spacing': d.get('line_spacing', {}).get('value', 1.5),
        }

        d.setdefault('page_number', {})
        d['page_number']['front_matter'] = {
            'style': self.fm_style.get(), 'position': self.fm_pos.get(),
            'align': self.fm_align.get(), 'font_size': 12}
        d['page_number']['body'] = {
            'style': self.bm_style.get(), 'position': self.bm_pos.get(),
            'align': self.bm_align.get(), 'start_from': self._fi(self.bm_start, 1), 'font_size': 12}

        d['figures'] = {
            'numbering': self.fig_num.get(), 'example': d.get('figures', {}).get('example', ''),
            'caption_below': self.fig_caption.get(),
            'caption_bold': d.get('figures', {}).get('caption_bold', False),
            'caption_align': d.get('figures', {}).get('caption_align', 'justify'),
            'caption_font_size': d.get('figures', {}).get('caption_font_size', 12),
            'source_font_size': d.get('figures', {}).get('source_font_size', 10),
        }
        d['tables'] = {
            'numbering': self.tab_num.get(), 'example': d.get('tables', {}).get('example', ''),
            'caption_above': self.tab_caption.get(),
            'caption_bold': d.get('tables', {}).get('caption_bold', True),
            'caption_align': d.get('tables', {}).get('caption_align', 'justify'),
            'caption_font_size': d.get('tables', {}).get('caption_font_size', 12),
            'source_font_size': d.get('tables', {}).get('source_font_size', 10),
        }

        self.result = d
        self.destroy()


class TemplateManagerDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("模板管理")
        self.geometry("700x500")
        self.minsize(600, 400)
        self.configure(bg=BG)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self._build_ui()
        self._refresh_list()

    def _build_ui(self):
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True, padx=16, pady=16)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        tk.Label(outer, text="管理模板", bg=BG, fg=TEXT,
                 font=("Microsoft JhengHei UI", 14, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 8))

        columns = ("name", "key", "source")
        self.tree = ttk.Treeview(outer, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("name", text="模板名稱", anchor="w")
        self.tree.heading("key", text="識別碼", anchor="w")
        self.tree.heading("source", text="來源", anchor="center")
        self.tree.column("name", width=250, minwidth=150)
        self.tree.column("key", width=120, minwidth=80)
        self.tree.column("source", width=80, minwidth=60, anchor="center")
        self.tree.grid(row=1, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=1, column=1, sticky="ns")

        self.tree.bind("<Double-1>", lambda e: self._edit_selected())
        self.tree.bind("<Return>", lambda e: self._edit_selected())

        btn_frame = tk.Frame(outer, bg=BG)
        btn_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))

        tk.Button(btn_frame, text="新增", bg="#059669", fg="white", font=("Microsoft JhengHei UI", 10),
                  bd=0, padx=16, pady=4, command=self._add_new).pack(side="left", padx=(0, 6))
        self.btn_edit = tk.Button(btn_frame, text="編輯", bg=ACCENT, fg="white", font=("Microsoft JhengHei UI", 10),
                                  bd=0, padx=16, pady=4, command=self._edit_selected)
        self.btn_edit.pack(side="left", padx=(0, 6))
        self.btn_del = tk.Button(btn_frame, text="刪除", bg="#DC2626", fg="white", font=("Microsoft JhengHei UI", 10),
                                 bd=0, padx=16, pady=4, command=self._delete_selected)
        self.btn_del.pack(side="left", padx=(0, 6))
        tk.Button(btn_frame, text="匯入", bg="#D97706", fg="white", font=("Microsoft JhengHei UI", 10),
                  bd=0, padx=16, pady=4, command=self._import_template).pack(side="left", padx=(0, 6))
        self.btn_export = tk.Button(btn_frame, text="匯出", bg=TEXT_SEC, fg="white", font=("Microsoft JhengHei UI", 10),
                                    bd=0, padx=16, pady=4, command=self._export_selected)
        self.btn_export.pack(side="left", padx=(0, 6))
        tk.Button(btn_frame, text="複製", bg="#7C3AED", fg="white", font=("Microsoft JhengHei UI", 10),
                  bd=0, padx=16, pady=4, command=self._duplicate_selected).pack(side="left", padx=(0, 6))
        tk.Button(btn_frame, text="關閉", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 10),
                  bd=1, relief="solid", padx=20, pady=4, command=self.destroy).pack(side="right")

        self.status_label = tk.Label(outer, text="", bg=BG, fg=OK_GRN, font=("Microsoft JhengHei UI", 9))
        self.status_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

    def _refresh_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.templates = TemplateManager.get_all_templates()
        for t in self.templates:
            src = "內建" if t.get('builtin') else "自訂"
            self.tree.insert("", tk.END, iid=t['key'], values=(t['name'], t['key'], src))

    def _get_selected_key(self) -> str:
        sel = self.tree.selection()
        return sel[0] if sel else None

    def _get_selected_info(self) -> dict:
        key = self._get_selected_key()
        if not key:
            return None
        for t in self.templates:
            if t['key'] == key:
                return t
        return None

    def _add_new(self):
        sel_key = self._get_selected_key() or "zh"
        base_info = TemplateManager.get_template_info(sel_key)
        dialog = tk.Toplevel(self)
        dialog.title("新增自訂模板")
        dialog.geometry("400x220")
        dialog.configure(bg=BG)
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        f = tk.Frame(dialog, bg=BG)
        f.pack(fill="both", expand=True, padx=16, pady=16)
        tk.Label(f, text="以「{}」為基礎".format(base_info['name']), bg=BG, fg=TEXT,
                 font=("Microsoft JhengHei UI", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))
        tk.Label(f, text="識別碼", bg=BG, fg=TEXT, font=("Microsoft JhengHei UI", 9)).grid(row=1, column=0, sticky="w")
        entry_key = tk.Entry(f, font=("Consolas", 10), bd=1, relief="solid")
        entry_key.grid(row=1, column=1, sticky="ew", padx=(8, 0), ipady=3)
        f.columnconfigure(1, weight=1)
        tk.Label(f, text="模板名稱", bg=BG, fg=TEXT, font=("Microsoft JhengHei UI", 9)).grid(row=2, column=0, sticky="w", pady=(8, 0))
        entry_name = tk.Entry(f, font=("Microsoft JhengHei UI", 10), bd=1, relief="solid")
        entry_name.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(8, 0), ipady=3)

        def do_create():
            key = entry_key.get().strip()
            name = entry_name.get().strip()
            if not key: messagebox.showerror("錯誤", "請輸入識別碼", parent=dialog); return
            if not name: messagebox.showerror("錯誤", "請輸入模板名稱", parent=dialog); return
            try:
                TemplateManager.create_template(key, name, base_on=sel_key)
                self._refresh_list()
                self.status_label.config(text=f"已建立自訂模板: {name}", fg=OK_GRN)
                dialog.destroy()
            except (TemplateExistsError, Exception) as e:
                messagebox.showerror("錯誤", str(e), parent=dialog)

        btn_f = tk.Frame(f, bg=BG)
        btn_f.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        tk.Button(btn_f, text="取消", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 10),
                  bd=1, relief="solid", padx=16, pady=4, command=dialog.destroy).pack(side="right", padx=(8, 0))
        tk.Button(btn_f, text="建立", bg=ACCENT, fg="white", font=("Microsoft JhengHei UI", 10, "bold"),
                  bd=0, padx=16, pady=4, command=do_create).pack(side="right")
        entry_key.focus_set()

    def _edit_selected(self):
        info = self._get_selected_info()
        if not info: messagebox.showinfo("提示", "請先選擇一個模板", parent=self); return
        if info.get('builtin'): messagebox.showinfo("提示", "內建模板無法編輯，請複製為自訂模板後再編輯", parent=self); return
        try: data = TemplateManager.load_template_data(info['key'])
        except Exception as e: messagebox.showerror("錯誤", f"無法載入模板: {e}", parent=self); return
        dlg = TemplateEditorDialog(self, info['key'], data, is_new=False)
        if dlg.winfo_exists():
            self.wait_window(dlg)
            if dlg.result is not None:
                try:
                    TemplateManager.save_template(info['key'], dlg.result)
                    self._refresh_list()
                    self.status_label.config(text=f"已儲存模板: {dlg.result.get('template', {}).get('name', '')}", fg=OK_GRN)
                except Exception as e: messagebox.showerror("錯誤", str(e), parent=self)

    def _delete_selected(self):
        info = self._get_selected_info()
        if not info: messagebox.showinfo("提示", "請先選擇一個模板", parent=self); return
        if info.get('builtin'): messagebox.showinfo("提示", "內建模板無法刪除", parent=self); return
        if not messagebox.askyesno("確認刪除", f"確定刪除自訂模板「{info['name']}」？", parent=self): return
        try:
            TemplateManager.delete_template(info['key'])
            self._refresh_list()
            self.status_label.config(text=f"已刪除模板: {info['name']}", fg=OK_GRN)
        except Exception as e: messagebox.showerror("錯誤", str(e), parent=self)

    def _import_template(self):
        path = filedialog.askopenfilename(title="選擇要匯入的模板",
                                          filetypes=[("模板檔案", "*.yaml *.yml *.json"), ("YAML 檔案", "*.yaml"),
                                                     ("JSON 檔案", "*.json"), ("所有檔案", "*.*")], parent=self)
        if not path: return
        try:
            key = TemplateManager.import_template(path)
            info = TemplateManager.get_template_info(key)
            self._refresh_list()
            self.status_label.config(text=f"已匯入模板: {info['name']}", fg=OK_GRN)
        except (TemplateExistsError, Exception) as e:
            if isinstance(e, TemplateExistsError):
                if messagebox.askyesno("模板已存在", "自訂模板已存在，是否覆蓋？", parent=self):
                    try:
                        TemplateManager.delete_template(key)
                        TemplateManager.import_template(path)
                        self._refresh_list()
                        self.status_label.config(text="已覆蓋匯入模板", fg=OK_GRN)
                    except Exception as e2: messagebox.showerror("錯誤", str(e2), parent=self)
            else: messagebox.showerror("錯誤", str(e), parent=self)

    def _export_selected(self):
        info = self._get_selected_info()
        if not info: messagebox.showinfo("提示", "請先選擇一個模板", parent=self); return
        path = filedialog.asksaveasfilename(title="匯出模板", defaultextension=".yaml",
            initialfile=f"{info['key']}.yaml",
            filetypes=[("YAML 檔案", "*.yaml"), ("JSON 檔案", "*.json"), ("所有檔案", "*.*")], parent=self)
        if not path: return
        try:
            TemplateManager.export_template(info['key'], path)
            self.status_label.config(text=f"已匯出至: {path}", fg=OK_GRN)
        except Exception as e: messagebox.showerror("錯誤", str(e), parent=self)

    def _duplicate_selected(self):
        info = self._get_selected_info()
        if not info: messagebox.showinfo("提示", "請先選擇一個模板", parent=self); return
        dialog = tk.Toplevel(self)
        dialog.title("複製模板")
        dialog.geometry("400x220")
        dialog.configure(bg=BG)
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        f = tk.Frame(dialog, bg=BG)
        f.pack(fill="both", expand=True, padx=16, pady=16)
        f.columnconfigure(1, weight=1)
        tk.Label(f, text="以「{}」為基礎".format(info['name']), bg=BG, fg=TEXT,
                 font=("Microsoft JhengHei UI", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))
        tk.Label(f, text="識別碼", bg=BG, fg=TEXT, font=("Microsoft JhengHei UI", 9)).grid(row=1, column=0, sticky="w")
        entry_key = tk.Entry(f, font=("Consolas", 10), bd=1, relief="solid")
        entry_key.grid(row=1, column=1, sticky="ew", padx=(8, 0), ipady=3)
        tk.Label(f, text="模板名稱", bg=BG, fg=TEXT, font=("Microsoft JhengHei UI", 9)).grid(row=2, column=0, sticky="w", pady=(8, 0))
        entry_name = tk.Entry(f, font=("Microsoft JhengHei UI", 10), bd=1, relief="solid")
        entry_name.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(8, 0), ipady=3)

        def do_duplicate():
            new_key = entry_key.get().strip()
            new_name = entry_name.get().strip()
            if not new_key: messagebox.showerror("錯誤", "請輸入識別碼", parent=dialog); return
            if not new_name: messagebox.showerror("錯誤", "請輸入模板名稱", parent=dialog); return
            try:
                TemplateManager.duplicate_template(info['key'], new_key, new_name)
                self._refresh_list()
                self.status_label.config(text=f"已複製模板: {new_name}", fg=OK_GRN)
                dialog.destroy()
            except (TemplateExistsError, BuiltinTemplateError, Exception) as e:
                messagebox.showerror("錯誤", str(e), parent=dialog)

        btn_f = tk.Frame(f, bg=BG)
        btn_f.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        tk.Button(btn_f, text="取消", bg=CARD_BG, fg=TEXT, font=("Microsoft JhengHei UI", 10),
                  bd=1, relief="solid", padx=16, pady=4, command=dialog.destroy).pack(side="right", padx=(8, 0))
        tk.Button(btn_f, text="複製", bg="#7C3AED", fg="white", font=("Microsoft JhengHei UI", 10, "bold"),
                  bd=0, padx=16, pady=4, command=do_duplicate).pack(side="right")
        entry_key.focus_set()
