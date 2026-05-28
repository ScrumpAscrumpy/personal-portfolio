import tkinter as tk
from tkinter import ttk, messagebox
import os
import subprocess
import sys

from .main_window import COLORS, _cjk_font


class HistoryTab:
    def __init__(self, parent, app):
        self.app = app
        self.frame = ttk.Frame(parent)
        self._records = []

        self._build()

    def _build(self):
        # Title row
        top = ttk.Frame(self.frame, style='Surface.TFrame', padding=(16, 10))
        top.pack(fill=tk.X, padx=12, pady=(12, 0))
        ttk.Label(top, text='历史记录', style='Title.TLabel').pack(side=tk.LEFT)
        self._count_label = ttk.Label(top, text='', style='Muted.TLabel')
        self._count_label.pack(side=tk.LEFT, padx=8, pady=2)
        ttk.Button(top, text='刷新', style='Secondary.TButton',
                   command=self.on_show).pack(side=tk.RIGHT)

        # Treeview
        tree_frame = ttk.Frame(self.frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        cols = ('timestamp', 'group', 'fields_preview', 'count', 'folder')
        self._tree = ttk.Treeview(tree_frame, columns=cols, show='headings',
                                   selectmode='browse')
        self._tree.heading('timestamp', text='时间')
        self._tree.heading('group', text='模板组')
        self._tree.heading('fields_preview', text='内容预览')
        self._tree.heading('count', text='文档数')
        self._tree.heading('folder', text='输出路径')

        self._tree.column('timestamp', width=150, minwidth=130)
        self._tree.column('group', width=130, minwidth=100)
        self._tree.column('fields_preview', width=220, minwidth=150)
        self._tree.column('count', width=60, anchor='center')
        self._tree.column('folder', width=260, minwidth=150)

        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self._tree.pack(fill=tk.BOTH, expand=True)

        self._tree.bind('<Double-1>', lambda e: self._refill())

        # Bottom buttons
        btn_frame = ttk.Frame(self.frame, style='Surface.TFrame', padding=(12, 8))
        btn_frame.pack(fill=tk.X, padx=12, pady=(0, 12))

        sep = ttk.Frame(btn_frame, style='Separator.TFrame', height=1)
        sep.pack(fill=tk.X, pady=(0, 8))

        ttk.Button(btn_frame, text='打开输出文件夹', style='Primary.TButton',
                   command=self._open_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_frame, text='重新填写此记录', style='Secondary.TButton',
                   command=self._refill).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_frame, text='删除记录', style='Danger.TButton',
                   command=self._delete_record).pack(side=tk.LEFT)

        hint = ttk.Label(btn_frame, text='双击行可重新填写', style='Muted.TLabel')
        hint.pack(side=tk.RIGHT, padx=4)

    def on_show(self):
        self._records = self.app.storage.load_history()
        self._tree.delete(*self._tree.get_children())

        for rec in self._records:
            ts_raw = rec.get('timestamp', '')
            ts = ts_raw[:19].replace('T', ' ') if ts_raw else '—'
            group = rec.get('group_name', '—')
            values = rec.get('field_values', {})
            preview = '  '.join(f'{k}: {v}' for k, v in list(values.items())[:3] if v)
            if len(values) > 3:
                preview += '…'
            count = str(rec.get('filled_count', 0))
            folder = rec.get('output_folder', '—')

            self._tree.insert('', tk.END, iid=rec.get('id', ts_raw),
                               values=(ts, group, preview, count, folder))

        self._count_label.config(text=f'共 {len(self._records)} 条')

    def _selected_record(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning('提示', '请先选择一条记录', parent=self.app.root)
            return None
        iid = sel[0]
        return next((r for r in self._records
                     if r.get('id', r.get('timestamp', '')) == iid), None)

    def _open_folder(self):
        rec = self._selected_record()
        if not rec:
            return
        folder = rec.get('output_folder', '')
        if not folder or not os.path.exists(folder):
            messagebox.showwarning('提示', f'文件夹不存在：{folder}', parent=self.app.root)
            return
        try:
            if sys.platform == 'win32':
                os.startfile(folder)
            elif sys.platform == 'darwin':
                subprocess.run(['open', folder])
            else:
                subprocess.run(['xdg-open', folder])
        except Exception as e:
            messagebox.showerror('错误', str(e), parent=self.app.root)

    def _refill(self):
        rec = self._selected_record()
        if not rec:
            return
        group_id = rec.get('group_id', '')
        field_values = rec.get('field_values', {})
        if not group_id:
            messagebox.showwarning('提示', '记录中没有模板组信息', parent=self.app.root)
            return
        group = self.app.template_manager.get_group(group_id)
        if not group:
            messagebox.showwarning('提示', f'原模板组已被删除（ID: {group_id}）',
                                    parent=self.app.root)
            return
        self.app.goto_form_with_data(field_values, group_id)

    def _delete_record(self):
        rec = self._selected_record()
        if not rec:
            return
        if not messagebox.askyesno('确认删除', '删除该历史记录？（不会删除已生成的文档）',
                                    parent=self.app.root):
            return
        rec_id = rec.get('id', '')
        if rec_id:
            self.app.storage.delete_history_record(rec_id)
        self.on_show()
