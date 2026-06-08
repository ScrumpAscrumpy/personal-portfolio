#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DocFiller — Word表格智能填写工具
上传任意 Word 表格 → 自动识别填写字段 → 输入信息 → 生成填好的文档

依赖: pip install python-docx
运行: python docfiller.py
"""
import json
import os
import re
import subprocess
import sys
import uuid
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document

# ═══════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════

HISTORY_FILE = str(Path.home() / '.docfiller_history.json')
CONFIG_FILE  = str(Path.home() / '.docfiller_config.json')

C = {
    'bg':      '#F0F2F5',
    'surface': '#FFFFFF',
    'primary': '#2563EB',
    'pri_d':   '#1D4ED8',
    'success': '#16A34A',
    'danger':  '#DC2626',
    'text':    '#111827',
    'muted':   '#6B7280',
    'border':  '#E5E7EB',
    'header':  '#1E3A5F',
    'entry':   '#F9FAFB',
}


def font(size=10, bold=False):
    fam = 'Microsoft YaHei,PingFang SC,Noto Sans CJK SC,WenQuanYi Micro Hei,SimHei,TkDefaultFont'
    return (fam.split(',')[0], size, 'bold' if bold else 'normal')


# ═══════════════════════════════════════════════
#  DOCUMENT ANALYSIS
# ═══════════════════════════════════════════════

class FormField:
    __slots__ = ('label', 'source', 'hint', 'table_idx', 'row_idx', 'cell_pos')

    def __init__(self, label, source, hint='', table_idx=0, row_idx=0, cell_pos=0):
        self.label     = label       # display / match key
        self.source    = source      # 'table' | 'paragraph'
        self.hint      = hint        # placeholder hint text e.g. "年   月"
        self.table_idx = table_idx
        self.row_idx   = row_idx
        self.cell_pos  = cell_pos    # position index in unique-cell list


def _unique_cells(row):
    seen, result = set(), []
    for cell in row.cells:
        cid = id(cell)
        if cid not in seen:
            seen.add(cid)
            result.append(cell)
    return result


_DATE_LOC_RE = re.compile(
    r'^[\s　\xa0]*'
    r'(?:年[\s　\xa0年月日]*月|省[\s　\xa0省市县区]*市|'
    r'县[\s　\xa0]*区|[\s　\xa0]+)'
    r'[\s　\xa0]*$'
)


def _is_value_cell(text):
    """True if a cell is empty or holds only a format hint."""
    t = text.strip()
    if not t:
        return True
    if len(t) > 60:          # Long text = real content, not a blank
        return False
    if '□' in t or '■' in t:  # Checkbox options — not a blank value cell
        return False
    if set(t) <= set(' 　\xa0\t'):
        return True
    if _DATE_LOC_RE.match(t):
        return True
    if '省' in t and '市' in t and len(t) < 25:
        return True
    if '年' in t and '月' in t and len(t) < 15:
        return True
    return False


def _is_label(text, col_pos):
    t = text.strip()
    if not t:
        return False
    if len(t) > 16:
        return False
    if '□' in t or '■' in t or '●' in t:
        return False
    if col_pos == 0 and len(t) > 8:   # first col is usually a section header
        return False
    if re.search(r'\d{5,}', t):
        return False
    return True


_PARA_RE = re.compile(r'([^\s：:□■\n]{1,15})\s*[：:]\s{3,}')


def analyze_document(path: str):
    """Parse a Word document; return list of FormField objects."""
    doc  = Document(path)
    fields: list[FormField] = []
    seen = set()

    def add(f: FormField):
        key = f.label
        if key not in seen:
            seen.add(key)
            fields.append(f)

    # ── table fields ──────────────────────────
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            ucs = _unique_cells(row)
            for i, cell in enumerate(ucs):
                if not _is_label(cell.text, i):
                    continue
                if i + 1 >= len(ucs):
                    continue
                nxt = ucs[i + 1]
                if _is_value_cell(nxt.text):
                    add(FormField(
                        label     = cell.text.strip().replace('\n', ' '),
                        source    = 'table',
                        hint      = nxt.text.strip(),
                        table_idx = t_idx,
                        row_idx   = r_idx,
                        cell_pos  = i + 1,       # value cell position
                    ))

    # ── paragraph fields ──────────────────────
    for para in doc.paragraphs:
        for m in _PARA_RE.finditer(para.text):
            lbl = m.group(1).strip()
            add(FormField(label=lbl, source='paragraph'))

    return fields


# ═══════════════════════════════════════════════
#  DOCUMENT FILLING
# ═══════════════════════════════════════════════

def _write_cell(cell, value: str):
    """Write text into a table cell, preserving the first run's font."""
    paras = cell.paragraphs
    if not paras:
        return
    first = paras[0]
    runs  = first.runs
    if runs:
        runs[0].text = value
        for run in runs[1:]:
            run.text = ''
    else:
        first.add_run(value)
    for para in paras[1:]:
        for run in para.runs:
            run.text = ''


def _write_paragraph(para, label: str, value: str):
    """Replace the blank spaces after a label in a paragraph run."""
    pattern = re.compile(
        rf'({re.escape(label)}\s*[：:]\s+)'
    )
    full = para.text
    if not pattern.search(full):
        return
    new_text = pattern.sub(rf'\g<1>{value}  ', full, count=1)
    runs = para.runs
    if runs:
        runs[0].text = new_text
        for run in runs[1:]:
            run.text = ''
    else:
        para.add_run(new_text)


def fill_document(src_path: str, label_values: dict, output_path: str):
    """
    Load the document, fill fields by label, save to output_path.
    label_values: {label_str: value_str}
    """
    doc = Document(src_path)
    applied = set()

    # ── table fields ──────────────────────────
    for table in doc.tables:
        for row in table.rows:
            ucs = _unique_cells(row)
            for i, cell in enumerate(ucs):
                if not _is_label(cell.text, i):
                    continue
                if i + 1 >= len(ucs):
                    continue
                label = cell.text.strip().replace('\n', ' ')
                value = label_values.get(label, '')
                if value and label not in applied:
                    _write_cell(ucs[i + 1], value)
                    applied.add(label)

    # ── paragraph fields ──────────────────────
    for para in doc.paragraphs:
        for m in _PARA_RE.finditer(para.text):
            label = m.group(1).strip()
            value = label_values.get(label, '')
            if value and label not in applied:
                _write_paragraph(para, label, value)
                applied.add(label)

    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    doc.save(output_path)


# ═══════════════════════════════════════════════
#  HISTORY / CONFIG
# ═══════════════════════════════════════════════

def load_config() -> dict:
    try:
        with open(CONFIG_FILE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg: dict):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_history() -> list:
    try:
        with open(HISTORY_FILE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def append_history(record: dict):
    hist = load_history()
    hist.insert(0, record)
    hist = hist[:100]
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(hist, f, ensure_ascii=False, indent=2)


# ═══════════════════════════════════════════════
#  UI STYLES
# ═══════════════════════════════════════════════

def apply_styles():
    s = ttk.Style()
    try:
        s.theme_use('clam')
    except Exception:
        pass

    s.configure('TFrame',          background=C['bg'])
    s.configure('Card.TFrame',     background=C['surface'])
    s.configure('TLabel',          background=C['bg'],      foreground=C['text'], font=font(10))
    s.configure('Card.TLabel',     background=C['surface'], foreground=C['text'], font=font(10))
    s.configure('Hint.TLabel',     background=C['surface'], foreground=C['muted'], font=font(9))
    s.configure('Title.TLabel',    background=C['surface'], foreground=C['text'], font=font(13, True))
    s.configure('Section.TLabel',  background=C['bg'],      foreground=C['muted'], font=font(9, True))
    s.configure('Sep.TFrame',      background=C['border'])

    for name, bg, fg, hbg in [
        ('Prim.TButton',  C['primary'], 'white', C['pri_d']),
        ('Succ.TButton',  C['success'], 'white', '#15803D'),
        ('Dang.TButton',  C['danger'],  'white', '#B91C1C'),
        ('Grey.TButton',  C['border'],  C['text'], '#D1D5DB'),
    ]:
        s.configure(name, background=bg, foreground=fg,
                    font=font(10, True), borderwidth=0, relief='flat', padding=(12, 6))
        s.map(name, background=[('active', hbg)], foreground=[('active', fg)])

    s.configure('TEntry', fieldbackground=C['entry'], font=font(10),
                borderwidth=1, relief='flat')
    s.configure('Treeview', font=font(10), background=C['surface'],
                fieldbackground=C['surface'], rowheight=28)
    s.configure('Treeview.Heading', font=font(10, True), background=C['bg'])
    s.map('Treeview', background=[('selected', C['primary'])],
                      foreground=[('selected', 'white')])
    s.configure('TNotebook',     background=C['bg'], borderwidth=0)
    s.configure('TNotebook.Tab', font=font(10), padding=(14, 7),
                background=C['bg'], foreground=C['muted'])
    s.map('TNotebook.Tab',
          background=[('selected', C['surface'])],
          foreground=[('selected', C['primary'])])
    s.configure('TScrollbar', background=C['border'], troughcolor=C['bg'])


# ═══════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════

class DocFillerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('DocFiller · Word表格智能填写')
        self.root.geometry('980x740')
        self.root.minsize(800, 580)
        self.root.configure(bg=C['bg'])

        self.cfg       = load_config()
        self.doc_path  = None
        self.fields    = []
        self._vars     = {}   # label -> StringVar

        apply_styles()
        self._build_ui()

    # ──────────────── BUILD UI ────────────────

    def _build_ui(self):
        # ── header ──
        hdr = tk.Frame(self.root, bg=C['header'], height=52)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text='DocFiller', bg=C['header'], fg='white',
                 font=font(14, True)).pack(side=tk.LEFT, padx=20, pady=10)
        tk.Label(hdr, text='Word表格智能填写工具', bg=C['header'],
                 fg='#93C5FD', font=font(10)).pack(side=tk.LEFT, pady=10)
        tk.Button(hdr, text='⚙ 设置', bg=C['header'], fg='#93C5FD',
                  relief='flat', font=font(10), cursor='hand2',
                  activebackground='#2D4A6E', activeforeground='white',
                  bd=0, padx=10, command=self._open_settings
                  ).pack(side=tk.RIGHT, padx=12, pady=10)

        # ── notebook ──
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        self._build_fill_tab()
        self._build_history_tab()

        self.nb.add(self._fill_tab,    text='  ✏  填写表单  ')
        self.nb.add(self._history_tab, text='  🕐  历史记录  ')

        # ── status bar ──
        sep = tk.Frame(self.root, bg=C['border'], height=1)
        sep.pack(fill=tk.X)
        sb = tk.Frame(self.root, bg=C['bg'], height=26)
        sb.pack(fill=tk.X)
        sb.pack_propagate(False)
        self._status_var = tk.StringVar(value='就绪  —  请打开一个 Word 文档')
        tk.Label(sb, textvariable=self._status_var, bg=C['bg'],
                 fg=C['muted'], font=font(9)).pack(side=tk.LEFT, padx=12)

        self.nb.bind('<<NotebookTabChanged>>', lambda e: self._on_tab())

    # ──────────────── FILL TAB ────────────────

    def _build_fill_tab(self):
        self._fill_tab = ttk.Frame(self.nb)

        # Top: open file area
        top = ttk.Frame(self._fill_tab, style='Card.TFrame', padding=(16, 12))
        top.pack(fill=tk.X, padx=12, pady=(12, 0))

        ttk.Label(top, text='Word 文档', style='Card.TLabel',
                  font=font(10, True)).pack(side=tk.LEFT)
        self._path_var = tk.StringVar(value='尚未选择文档')
        ttk.Label(top, textvariable=self._path_var,
                  style='Hint.TLabel').pack(side=tk.LEFT, padx=12)
        ttk.Button(top, text='📂  打开文档', style='Prim.TButton',
                   command=self._open_doc).pack(side=tk.RIGHT)

        # Scrollable form area
        mid = ttk.Frame(self._fill_tab)
        mid.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        self._canvas = tk.Canvas(mid, bg=C['bg'], highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(mid, orient='vertical', command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._form_frame = ttk.Frame(self._canvas, style='Card.TFrame', padding=(24, 16))
        self._cwin = self._canvas.create_window((0, 0), window=self._form_frame, anchor='nw')

        self._form_frame.bind('<Configure>',
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox('all')))
        self._canvas.bind('<Configure>',
            lambda e: self._canvas.itemconfig(self._cwin, width=e.width))
        self._canvas.bind_all('<MouseWheel>',
            lambda e: self._canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        self._canvas.bind_all('<Button-4>',
            lambda e: self._canvas.yview_scroll(-1, 'units'))
        self._canvas.bind_all('<Button-5>',
            lambda e: self._canvas.yview_scroll(1, 'units'))

        # Empty state
        self._empty_lbl = ttk.Label(self._form_frame,
            text='点击「打开文档」选择一个 Word 表格文件\n程序会自动识别可填写的字段',
            style='Hint.TLabel', font=font(12), justify='center')
        self._empty_lbl.pack(pady=60)

        # Bottom bar
        btm = ttk.Frame(self._fill_tab, style='Card.TFrame', padding=(16, 10))
        btm.pack(fill=tk.X, padx=12, pady=(0, 12))
        ttk.Frame(btm, style='Sep.TFrame', height=1).pack(fill=tk.X, pady=(0, 10))

        out_row = ttk.Frame(btm, style='Card.TFrame')
        out_row.pack(fill=tk.X)
        ttk.Label(out_row, text='保存到：', style='Card.TLabel').pack(side=tk.LEFT)
        self._out_var = tk.StringVar(value=self.cfg.get('output_dir', ''))
        self._out_lbl = ttk.Label(out_row, textvariable=self._out_var,
                                   style='Hint.TLabel', cursor='hand2')
        self._out_lbl.pack(side=tk.LEFT, padx=6)
        self._out_lbl.bind('<Button-1>', lambda e: self._pick_out_dir())
        ttk.Button(out_row, text='更改', style='Grey.TButton',
                   command=self._pick_out_dir).pack(side=tk.RIGHT)

        act_row = ttk.Frame(btm, style='Card.TFrame')
        act_row.pack(fill=tk.X, pady=(10, 0))
        self._gen_btn = ttk.Button(act_row, text='生成文档', style='Succ.TButton',
                                    command=self._generate, width=18)
        self._gen_btn.pack(side=tk.LEFT)
        ttk.Button(act_row, text='清空表单', style='Grey.TButton',
                   command=self._clear_form).pack(side=tk.LEFT, padx=8)
        self._result_lbl = ttk.Label(act_row, text='', style='Card.TLabel',
                                      foreground=C['success'])
        self._result_lbl.pack(side=tk.LEFT, padx=8)

    # ──────────────── HISTORY TAB ────────────────

    def _build_history_tab(self):
        self._history_tab = ttk.Frame(self.nb)

        top = ttk.Frame(self._history_tab, style='Card.TFrame', padding=(16, 10))
        top.pack(fill=tk.X, padx=12, pady=(12, 0))
        ttk.Label(top, text='历史记录', style='Title.TLabel').pack(side=tk.LEFT)
        ttk.Button(top, text='刷新', style='Grey.TButton',
                   command=self._refresh_history).pack(side=tk.RIGHT)

        tf = ttk.Frame(self._history_tab)
        tf.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        cols = ('time', 'doc', 'fields', 'output')
        self._hist_tree = ttk.Treeview(tf, columns=cols, show='headings', selectmode='browse')
        self._hist_tree.heading('time',   text='时间')
        self._hist_tree.heading('doc',    text='原文档')
        self._hist_tree.heading('fields', text='字段数')
        self._hist_tree.heading('output', text='输出路径')
        self._hist_tree.column('time',   width=155)
        self._hist_tree.column('doc',    width=160)
        self._hist_tree.column('fields', width=65, anchor='center')
        self._hist_tree.column('output', width=380)

        vsb = ttk.Scrollbar(tf, orient='vertical', command=self._hist_tree.yview)
        self._hist_tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._hist_tree.pack(fill=tk.BOTH, expand=True)
        self._hist_tree.bind('<Double-1>', lambda e: self._refill_from_history())

        bf = ttk.Frame(self._history_tab, style='Card.TFrame', padding=(12, 8))
        bf.pack(fill=tk.X, padx=12, pady=(0, 12))
        ttk.Frame(bf, style='Sep.TFrame', height=1).pack(fill=tk.X, pady=(0, 8))
        ttk.Button(bf, text='打开输出文件夹', style='Prim.TButton',
                   command=self._open_hist_folder).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(bf, text='重新填写', style='Grey.TButton',
                   command=self._refill_from_history).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(bf, text='双击行可重新填写', style='Hint.TLabel').pack(side=tk.RIGHT)

        self._history_records = []
        self._refresh_history()

    # ──────────────── OPEN DOC ────────────────

    def _open_doc(self):
        path = filedialog.askopenfilename(
            title='选择 Word 表格文件',
            filetypes=[('Word 文档', '*.docx'), ('所有文件', '*.*')],
            initialdir=self.cfg.get('last_dir', str(Path.home()))
        )
        if not path:
            return

        self.cfg['last_dir'] = str(Path(path).parent)
        save_config(self.cfg)

        self.doc_path = path
        self._path_var.set(os.path.basename(path))
        self._status_var.set(f'正在分析文档：{os.path.basename(path)} ...')
        self.root.update()

        try:
            self.fields = analyze_document(path)
        except Exception as e:
            messagebox.showerror('分析失败', str(e), parent=self.root)
            return

        if not self.fields:
            messagebox.showinfo('提示', '未在该文档中检测到可填写字段。\n'
                                        '请确认文档包含"标签名：____"格式的空白填写区域。',
                                parent=self.root)
            return

        self._render_form(self.fields)
        self._gen_btn.config(text=f'生成文档  ({len(self.fields)} 个字段)')
        if not self._out_var.get():
            self._out_var.set(str(Path(path).parent))
        self._status_var.set(
            f'已识别 {len(self.fields)} 个字段  ·  {os.path.basename(path)}')

    # ──────────────── RENDER FORM ────────────────

    def _render_form(self, fields, prefill=None):
        for w in self._form_frame.winfo_children():
            w.destroy()
        self._vars.clear()

        if not fields:
            ttk.Label(self._form_frame, text='未检测到字段', style='Hint.TLabel').pack(pady=40)
            return

        ttk.Label(self._form_frame, text='请填写以下信息',
                  style='Title.TLabel').grid(row=0, column=0, columnspan=2,
                                              sticky='w', pady=(0, 16))
        ttk.Label(self._form_frame,
                  text='标有 * 的为从文档中检测到的字段，填写后将写入对应位置',
                  style='Hint.TLabel').grid(row=1, column=0, columnspan=2,
                                             sticky='w', pady=(0, 12))

        # Group by section (table_idx)
        # Sort by (table_idx, row_idx) for table fields; paragraph fields at end
        table_f  = [f for f in fields if f.source == 'table']
        para_f   = [f for f in fields if f.source == 'paragraph']
        sorted_f = sorted(table_f, key=lambda f: (f.table_idx, f.row_idx)) + para_f

        base_row = 2
        for i, field in enumerate(sorted_f):
            row = base_row + i
            label_txt = f'{field.label}  *'
            lbl = ttk.Label(self._form_frame, text=label_txt, style='Card.TLabel',
                             font=font(10, True))
            lbl.grid(row=row, column=0, sticky='ne', padx=(0, 14), pady=7)

            var = tk.StringVar()
            self._vars[field.label] = var

            if prefill and field.label in prefill:
                var.set(prefill[field.label])

            entry = ttk.Entry(self._form_frame, textvariable=var, width=48,
                               font=font(10))
            entry.grid(row=row, column=1, sticky='ew', pady=7)

            if field.hint:
                # Show hint as placeholder-like text below entry
                hint_text = f'示例格式：{field.hint}'
                ttk.Label(self._form_frame, text=hint_text,
                           style='Hint.TLabel').grid(
                    row=base_row + len(sorted_f) + 10 + i,
                    column=1, sticky='w', pady=(0, 2))

        # Paragraph fields section header
        if para_f and table_f:
            sep_row = base_row + len(table_f)
            ttk.Frame(self._form_frame, style='Sep.TFrame', height=1).grid(
                row=sep_row, column=0, columnspan=2, sticky='ew', pady=8)

        self._form_frame.columnconfigure(1, weight=1)

    # ──────────────── GENERATE ────────────────

    def _generate(self):
        if not self.doc_path:
            messagebox.showwarning('提示', '请先打开一个 Word 文档', parent=self.root)
            return
        if not self.fields:
            messagebox.showwarning('提示', '未检测到任何字段', parent=self.root)
            return

        values = {lbl: var.get().strip() for lbl, var in self._vars.items()}
        filled = {k: v for k, v in values.items() if v}

        if not filled:
            messagebox.showwarning('提示', '请至少填写一个字段', parent=self.root)
            return

        out_dir = self._out_var.get().strip() or str(Path(self.doc_path).parent)
        ts      = datetime.now().strftime('%Y%m%d_%H%M%S')
        stem    = Path(self.doc_path).stem
        out_name = f'{stem}_已填写_{ts}.docx'
        out_path = os.path.join(out_dir, out_name)

        try:
            fill_document(self.doc_path, filled, out_path)
        except Exception as e:
            messagebox.showerror('生成失败', str(e), parent=self.root)
            return

        record = {
            'id':         str(uuid.uuid4()),
            'timestamp':  datetime.now().isoformat(),
            'doc_path':   self.doc_path,
            'doc_name':   os.path.basename(self.doc_path),
            'values':     filled,
            'output':     out_path,
            'field_count': len(filled),
        }
        append_history(record)
        self._refresh_history()

        self._result_lbl.config(text=f'✓ 已生成：{out_name}')
        self.root.after(5000, lambda: self._result_lbl.config(text=''))
        self._status_var.set(f'文档已生成 → {out_path}')

        self._open_path(out_dir)

    # ──────────────── HISTORY ACTIONS ────────────────

    def _on_tab(self):
        if self.nb.index('current') == 1:
            self._refresh_history()

    def _refresh_history(self):
        self._history_records = load_history()
        self._hist_tree.delete(*self._hist_tree.get_children())
        for rec in self._history_records:
            ts  = rec.get('timestamp', '')[:19].replace('T', ' ')
            doc = rec.get('doc_name', '—')
            cnt = str(rec.get('field_count', 0))
            out = rec.get('output', '—')
            self._hist_tree.insert('', tk.END, iid=rec['id'],
                                    values=(ts, doc, cnt, out))

    def _selected_history(self):
        sel = self._hist_tree.selection()
        if not sel:
            messagebox.showwarning('提示', '请先选择一条记录', parent=self.root)
            return None
        iid = sel[0]
        return next((r for r in self._history_records if r['id'] == iid), None)

    def _open_hist_folder(self):
        rec = self._selected_history()
        if not rec:
            return
        folder = str(Path(rec.get('output', '')).parent)
        if not os.path.exists(folder):
            messagebox.showwarning('提示', f'文件夹不存在：{folder}', parent=self.root)
            return
        self._open_path(folder)

    def _refill_from_history(self):
        rec = self._selected_history()
        if not rec:
            return
        doc_path = rec.get('doc_path', '')
        if not os.path.exists(doc_path):
            messagebox.showwarning('提示',
                f'原文档不存在，请重新打开：\n{doc_path}', parent=self.root)
            return
        self.doc_path = doc_path
        self._path_var.set(os.path.basename(doc_path))
        try:
            self.fields = analyze_document(doc_path)
        except Exception as e:
            messagebox.showerror('分析失败', str(e), parent=self.root)
            return
        self._render_form(self.fields, prefill=rec.get('values', {}))
        self._gen_btn.config(text=f'生成文档  ({len(self.fields)} 个字段)')
        self.nb.select(0)
        self._status_var.set(f'已从历史加载：{os.path.basename(doc_path)}')

    # ──────────────── SETTINGS ────────────────

    def _open_settings(self):
        win = tk.Toplevel(self.root)
        win.title('设置')
        win.geometry('480x180')
        win.resizable(False, False)
        win.configure(bg=C['bg'])
        win.grab_set()
        win.transient(self.root)

        f = ttk.Frame(win, style='Card.TFrame', padding=20)
        f.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        ttk.Label(f, text='默认输出目录', style='Card.TLabel').grid(
            row=0, column=0, sticky='w', pady=6)
        od = tk.StringVar(value=self.cfg.get('output_dir', ''))
        ttk.Entry(f, textvariable=od, width=38).grid(
            row=0, column=1, padx=8, pady=6, sticky='ew')
        ttk.Button(f, text='浏览', style='Grey.TButton',
                   command=lambda: od.set(
                       filedialog.askdirectory(
                           initialdir=od.get() or str(Path.home()),
                           title='选择默认输出目录') or od.get())
                   ).grid(row=0, column=2, padx=4)

        ttk.Label(f, text='留空则保存到原文档所在目录', style='Hint.TLabel').grid(
            row=1, column=1, sticky='w', padx=8)

        f.columnconfigure(1, weight=1)

        def save():
            self.cfg['output_dir'] = od.get().strip()
            self._out_var.set(self.cfg['output_dir'])
            save_config(self.cfg)
            messagebox.showinfo('已保存', '设置已保存', parent=win)
            win.destroy()

        bf = ttk.Frame(f)
        bf.grid(row=2, column=0, columnspan=3, pady=12, sticky='e')
        ttk.Button(bf, text='取消', style='Grey.TButton',
                   command=win.destroy).pack(side=tk.LEFT, padx=4)
        ttk.Button(bf, text='保存', style='Prim.TButton',
                   command=save).pack(side=tk.LEFT, padx=4)

    # ──────────────── HELPERS ────────────────

    def _pick_out_dir(self):
        d = filedialog.askdirectory(
            initialdir=self._out_var.get() or str(Path.home()),
            title='选择输出目录')
        if d:
            self._out_var.set(d)

    def _clear_form(self):
        for var in self._vars.values():
            var.set('')

    def _open_path(self, path):
        if not os.path.exists(path):
            return
        try:
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.run(['open', path])
            else:
                subprocess.run(['xdg-open', path])
        except Exception:
            pass

    def run(self):
        self.root.mainloop()


# ═══════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════

if __name__ == '__main__':
    DocFillerApp().run()
