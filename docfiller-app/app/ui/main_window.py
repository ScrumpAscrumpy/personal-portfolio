import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys

from ..core.storage import Storage
from ..core.template_manager import TemplateManager


# Chinese-friendly font stack
def _cjk_font(size=10, weight='normal'):
    families = ['Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC',
                'WenQuanYi Micro Hei', 'SimHei', 'TkDefaultFont']
    return (families[0], size, weight)


COLORS = {
    'bg': '#F0F2F5',
    'surface': '#FFFFFF',
    'primary': '#2563EB',
    'primary_dark': '#1D4ED8',
    'danger': '#DC2626',
    'danger_dark': '#B91C1C',
    'success': '#16A34A',
    'text': '#1F2937',
    'text_muted': '#6B7280',
    'border': '#E5E7EB',
    'header_bg': '#1E3A5F',
    'tab_active': '#2563EB',
}


def setup_styles():
    style = ttk.Style()
    try:
        style.theme_use('clam')
    except Exception:
        pass

    font_normal = _cjk_font(10)
    font_bold = _cjk_font(10, 'bold')
    font_small = _cjk_font(9)
    font_large = _cjk_font(12, 'bold')

    style.configure('TFrame', background=COLORS['bg'])
    style.configure('Surface.TFrame', background=COLORS['surface'])
    style.configure('TLabel', background=COLORS['bg'], foreground=COLORS['text'], font=font_normal)
    style.configure('Surface.TLabel', background=COLORS['surface'], foreground=COLORS['text'], font=font_normal)
    style.configure('Muted.TLabel', background=COLORS['surface'], foreground=COLORS['text_muted'], font=font_small)
    style.configure('Title.TLabel', background=COLORS['surface'], foreground=COLORS['text'], font=font_large)
    style.configure('Header.TLabel', background=COLORS['header_bg'], foreground='white', font=font_bold)

    style.configure('TEntry', font=font_normal, fieldbackground='white',
                    borderwidth=1, relief='flat')
    style.map('TEntry', fieldbackground=[('focus', '#EFF6FF')])

    style.configure('Primary.TButton', background=COLORS['primary'], foreground='white',
                    font=font_bold, borderwidth=0, relief='flat', padding=(12, 6))
    style.map('Primary.TButton',
              background=[('active', COLORS['primary_dark']), ('pressed', COLORS['primary_dark'])],
              foreground=[('active', 'white')])

    style.configure('Danger.TButton', background=COLORS['danger'], foreground='white',
                    font=font_normal, borderwidth=0, relief='flat', padding=(8, 4))
    style.map('Danger.TButton',
              background=[('active', COLORS['danger_dark'])],
              foreground=[('active', 'white')])

    style.configure('Secondary.TButton', background='#E5E7EB', foreground=COLORS['text'],
                    font=font_normal, borderwidth=0, relief='flat', padding=(8, 4))
    style.map('Secondary.TButton',
              background=[('active', '#D1D5DB')],
              foreground=[('active', COLORS['text'])])

    style.configure('TCombobox', font=font_normal, fieldbackground='white')
    style.configure('TScrollbar', background=COLORS['border'], troughcolor=COLORS['bg'])

    style.configure('TNotebook', background=COLORS['bg'], borderwidth=0)
    style.configure('TNotebook.Tab', font=font_normal, padding=(16, 8),
                    background=COLORS['bg'], foreground=COLORS['text_muted'])
    style.map('TNotebook.Tab',
              background=[('selected', COLORS['surface'])],
              foreground=[('selected', COLORS['primary'])],
              expand=[('selected', [1, 1, 1, 0])])

    style.configure('Treeview', font=font_normal, background=COLORS['surface'],
                    fieldbackground=COLORS['surface'], foreground=COLORS['text'],
                    rowheight=28)
    style.configure('Treeview.Heading', font=font_bold, background=COLORS['bg'],
                    foreground=COLORS['text'])
    style.map('Treeview', background=[('selected', COLORS['primary'])],
              foreground=[('selected', 'white')])

    style.configure('Separator.TFrame', background=COLORS['border'])


class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('DocFiller · 批量文档填写系统')
        self.root.geometry('960x720')
        self.root.minsize(800, 600)
        self.root.configure(bg=COLORS['bg'])

        self.storage = Storage()
        self.template_manager = TemplateManager(self.storage)

        setup_styles()
        self._build_ui()

    def _build_ui(self):
        # Header bar
        header = tk.Frame(self.root, bg=COLORS['header_bg'], height=52)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        tk.Label(header, text='DocFiller', bg=COLORS['header_bg'],
                 fg='white', font=_cjk_font(14, 'bold')).pack(side=tk.LEFT, padx=20, pady=10)
        tk.Label(header, text='批量文档填写系统', bg=COLORS['header_bg'],
                 fg='#93C5FD', font=_cjk_font(10)).pack(side=tk.LEFT, pady=10)

        # Settings button in header
        settings_btn = tk.Button(header, text='⚙ 设置', bg=COLORS['header_bg'],
                                  fg='#93C5FD', relief='flat', font=_cjk_font(10),
                                  cursor='hand2', activebackground='#2D4A6E',
                                  activeforeground='white', bd=0, padx=12,
                                  command=self._open_settings)
        settings_btn.pack(side=tk.RIGHT, pady=8, padx=8)

        # Notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        from .form_tab import FormTab
        from .template_tab import TemplateTab
        from .history_tab import HistoryTab

        self.form_tab = FormTab(self.notebook, self)
        self.template_tab = TemplateTab(self.notebook, self)
        self.history_tab = HistoryTab(self.notebook, self)

        self.notebook.add(self.form_tab.frame, text='  ✏  填写表单  ')
        self.notebook.add(self.template_tab.frame, text='  📋  模板管理  ')
        self.notebook.add(self.history_tab.frame, text='  🕐  历史记录  ')

        # Status bar
        statusbar = tk.Frame(self.root, bg=COLORS['border'], height=1)
        statusbar.pack(fill=tk.X)
        status_frame = tk.Frame(self.root, bg=COLORS['bg'], height=30)
        status_frame.pack(fill=tk.X)
        status_frame.pack_propagate(False)

        self.status_var = tk.StringVar(value=f'数据目录: {self.storage.data_dir}')
        tk.Label(status_frame, textvariable=self.status_var, bg=COLORS['bg'],
                 fg=COLORS['text_muted'], font=_cjk_font(9)).pack(side=tk.LEFT, padx=12)

        # Notify tabs when switching
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_change)

    def _on_tab_change(self, event):
        tab_idx = self.notebook.index('current')
        if tab_idx == 0:
            self.form_tab.on_show()
        elif tab_idx == 1:
            self.template_tab.on_show()
        elif tab_idx == 2:
            self.history_tab.on_show()

    def _open_settings(self):
        win = tk.Toplevel(self.root)
        win.title('设置')
        win.geometry('520x220')
        win.resizable(False, False)
        win.configure(bg=COLORS['bg'])
        win.grab_set()
        win.transient(self.root)

        pad = {'padx': 16, 'pady': 6}

        f = ttk.Frame(win, style='Surface.TFrame', padding=20)
        f.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

        ttk.Label(f, text='数据目录', style='Surface.TLabel',
                  font=_cjk_font(9)).grid(row=0, column=0, sticky='w', pady=4)
        data_var = tk.StringVar(value=self.storage.data_dir)
        ttk.Entry(f, textvariable=data_var, width=40).grid(row=0, column=1, padx=8, pady=4)
        ttk.Button(f, text='浏览', style='Secondary.TButton',
                   command=lambda: data_var.set(filedialog.askdirectory(
                       initialdir=data_var.get(), title='选择数据目录') or data_var.get())
                   ).grid(row=0, column=2, padx=4)

        ttk.Label(f, text='输出目录', style='Surface.TLabel',
                  font=_cjk_font(9)).grid(row=1, column=0, sticky='w', pady=4)
        out_var = tk.StringVar(value=self.storage.output_dir)
        ttk.Entry(f, textvariable=out_var, width=40).grid(row=1, column=1, padx=8, pady=4)
        ttk.Button(f, text='浏览', style='Secondary.TButton',
                   command=lambda: out_var.set(filedialog.askdirectory(
                       initialdir=out_var.get() or self.storage.data_dir,
                       title='选择输出目录') or out_var.get())
                   ).grid(row=1, column=2, padx=4)

        ttk.Label(f, text='留空则使用: 数据目录/output', style='Muted.TLabel').grid(
            row=2, column=1, sticky='w', padx=8)

        def save():
            self.storage.data_dir = data_var.get().strip() or self.storage.data_dir
            self.storage.output_dir = out_var.get().strip()
            self.storage.save_config()
            self.status_var.set(f'数据目录: {self.storage.data_dir}')
            messagebox.showinfo('已保存', '设置已保存。', parent=win)
            win.destroy()

        btn_frame = ttk.Frame(f)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=12, sticky='e')
        ttk.Button(btn_frame, text='取消', style='Secondary.TButton',
                   command=win.destroy).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text='保存', style='Primary.TButton',
                   command=save).pack(side=tk.LEFT, padx=4)

    def goto_form_with_data(self, field_values: dict, group_id: str):
        """Pre-fill the form tab with given values and switch to it."""
        self.form_tab.load_data(group_id, field_values)
        self.notebook.select(0)

    def refresh_all(self):
        self.template_manager.reload()
        self.form_tab.on_show()
        self.template_tab.on_show()
        self.history_tab.on_show()

    def run(self):
        self.root.mainloop()
