import tkinter as tk
from tkinter import ttk, messagebox
import os
import subprocess
import sys
from datetime import datetime

from ..core import doc_filler
from .main_window import COLORS, _cjk_font


class FormTab:
    def __init__(self, parent, app):
        self.app = app
        self.frame = ttk.Frame(parent)
        self._field_vars = {}   # key -> StringVar
        self._current_group_id = None

        self._build()

    def _build(self):
        # Top: group selector
        top = ttk.Frame(self.frame, style='Surface.TFrame', padding=(16, 12))
        top.pack(fill=tk.X, padx=12, pady=(12, 0))

        ttk.Label(top, text='选择模板组：', style='Surface.TLabel',
                  font=_cjk_font(10, 'bold')).grid(row=0, column=0, sticky='w')

        self._group_var = tk.StringVar()
        self._group_combo = ttk.Combobox(top, textvariable=self._group_var,
                                          state='readonly', width=36,
                                          font=_cjk_font(10))
        self._group_combo.grid(row=0, column=1, padx=(8, 0), sticky='w')
        self._group_combo.bind('<<ComboboxSelected>>', self._on_group_selected)

        self._desc_label = ttk.Label(top, text='', style='Muted.TLabel')
        self._desc_label.grid(row=1, column=1, padx=(8, 0), sticky='w', pady=(2, 0))

        # Scrollable form area
        form_outer = ttk.Frame(self.frame)
        form_outer.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        self._canvas = tk.Canvas(form_outer, bg=COLORS['bg'],
                                  highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(form_outer, orient='vertical',
                                   command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._form_frame = ttk.Frame(self._canvas, style='Surface.TFrame', padding=(20, 12))
        self._canvas_window = self._canvas.create_window(
            (0, 0), window=self._form_frame, anchor='nw')

        self._form_frame.bind('<Configure>', self._on_frame_configure)
        self._canvas.bind('<Configure>', self._on_canvas_configure)
        self._canvas.bind_all('<MouseWheel>', self._on_mousewheel)
        self._canvas.bind_all('<Button-4>', self._on_mousewheel)
        self._canvas.bind_all('<Button-5>', self._on_mousewheel)

        # Empty state label
        self._empty_label = ttk.Label(self._form_frame,
                                       text='请先选择模板组，或前往「模板管理」创建模板组',
                                       style='Muted.TLabel', font=_cjk_font(11))
        self._empty_label.grid(row=0, column=0, pady=40, padx=20)

        # Bottom: output path + generate button
        bottom = ttk.Frame(self.frame, style='Surface.TFrame', padding=(16, 12))
        bottom.pack(fill=tk.X, padx=12, pady=(0, 12))
        sep = ttk.Frame(bottom, style='Separator.TFrame', height=1)
        sep.pack(fill=tk.X, pady=(0, 10))

        path_row = ttk.Frame(bottom, style='Surface.TFrame')
        path_row.pack(fill=tk.X)

        ttk.Label(path_row, text='输出到：', style='Surface.TLabel').pack(side=tk.LEFT)
        self._out_label = ttk.Label(path_row,
                                     text=self.app.storage.effective_output_dir,
                                     style='Muted.TLabel', cursor='hand2')
        self._out_label.pack(side=tk.LEFT, padx=4)
        self._out_label.bind('<Button-1>', lambda e: self._open_folder(
            self.app.storage.effective_output_dir))

        btn_row = ttk.Frame(bottom, style='Surface.TFrame')
        btn_row.pack(fill=tk.X, pady=(10, 0))

        self._gen_btn = ttk.Button(btn_row, text='生成文档',
                                    style='Primary.TButton', width=20,
                                    command=self._generate)
        self._gen_btn.pack(side=tk.LEFT)

        self._status_label = ttk.Label(btn_row, text='', style='Surface.TLabel',
                                        foreground=COLORS['success'])
        self._status_label.pack(side=tk.LEFT, padx=12)

    # ---------- Scroll helpers ----------

    def _on_frame_configure(self, event):
        self._canvas.configure(scrollregion=self._canvas.bbox('all'))

    def _on_canvas_configure(self, event):
        self._canvas.itemconfig(self._canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        if event.num == 4:
            self._canvas.yview_scroll(-1, 'units')
        elif event.num == 5:
            self._canvas.yview_scroll(1, 'units')
        else:
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    # ---------- Group selection ----------

    def on_show(self):
        groups = self.app.template_manager.groups
        names = [g['name'] for g in groups]
        self._group_combo['values'] = names
        self._out_label.config(text=self.app.storage.effective_output_dir)

        # Restore selection if still valid
        if self._current_group_id:
            group = self.app.template_manager.get_group(self._current_group_id)
            if group:
                self._group_var.set(group['name'])
                return

        if names and not self._group_var.get():
            self._group_var.set(names[0])
            self._on_group_selected(None)

    def _on_group_selected(self, event):
        groups = self.app.template_manager.groups
        name = self._group_var.get()
        group = next((g for g in groups if g['name'] == name), None)
        if not group:
            return
        self._current_group_id = group['id']
        self._desc_label.config(text=group.get('description', ''))
        self._render_fields(group)

    def _render_fields(self, group: dict):
        for w in self._form_frame.winfo_children():
            w.destroy()
        self._field_vars.clear()

        fields = group.get('fields', [])
        if not fields:
            ttk.Label(self._form_frame,
                      text='该模板组暂无字段定义。请前往「模板管理」添加字段。',
                      style='Muted.TLabel', font=_cjk_font(11)).grid(
                row=0, column=0, pady=40, padx=20)
            return

        ttk.Label(self._form_frame, text='请填写以下信息',
                  style='Title.TLabel').grid(row=0, column=0, columnspan=2,
                                              sticky='w', pady=(4, 16))

        for i, field in enumerate(fields):
            row = i + 1
            key = field['key']
            label = field.get('label', key)
            required = field.get('required', True)
            hint = field.get('hint', '')
            ftype = field.get('field_type', 'text')

            # Label
            label_text = f'{label}{"  *" if required else ""}'
            lbl = ttk.Label(self._form_frame, text=label_text, style='Surface.TLabel',
                             font=_cjk_font(10, 'bold' if required else 'normal'))
            lbl.grid(row=row, column=0, sticky='ne', padx=(0, 12), pady=6)

            var = tk.StringVar()
            self._field_vars[key] = var

            if ftype == 'multiline':
                txt_frame = ttk.Frame(self._form_frame, style='Surface.TFrame')
                txt_frame.grid(row=row, column=1, sticky='ew', pady=6)
                txt = tk.Text(txt_frame, height=3, width=40, font=_cjk_font(10),
                               relief='solid', bd=1, wrap='word',
                               bg='white', fg=COLORS['text'])
                txt.pack(fill=tk.X)
                # Bind Text widget to StringVar via trace
                def make_sync(widget, v):
                    def sync_to_var(*args):
                        v.set(widget.get('1.0', 'end-1c'))
                    def sync_to_widget(*args):
                        content = v.get()
                        widget.delete('1.0', 'end')
                        widget.insert('1.0', content)
                    widget.bind('<KeyRelease>', sync_to_var)
                    return sync_to_widget
                cb = make_sync(txt, var)
                var.trace_add('write', lambda *a, cb=cb: cb())
                self._field_vars[f'__widget_{key}'] = txt
            else:
                entry = ttk.Entry(self._form_frame, textvariable=var, width=42,
                                   font=_cjk_font(10))
                entry.grid(row=row, column=1, sticky='ew', pady=6)

            if hint:
                ttk.Label(self._form_frame, text=f'  {hint}',
                           style='Muted.TLabel').grid(
                    row=1000 + i, column=1, sticky='w', pady=(0, 4))

        self._form_frame.columnconfigure(1, weight=1)
        self._gen_btn.config(
            text=f'生成文档  ({len(group.get("templates", []))} 个文档)')

    def load_data(self, group_id: str, field_values: dict):
        """Pre-fill form with given values (called from history tab)."""
        groups = self.app.template_manager.groups
        group = next((g for g in groups if g['id'] == group_id), None)
        if not group:
            return
        self._current_group_id = group_id
        self._group_var.set(group['name'])
        self._desc_label.config(text=group.get('description', ''))
        self._render_fields(group)
        for key, val in field_values.items():
            if key in self._field_vars:
                self._field_vars[key].set(val)

    # ---------- Generate ----------

    def _generate(self):
        if not self._current_group_id:
            messagebox.showwarning('提示', '请先选择模板组', parent=self.app.root)
            return

        group = self.app.template_manager.get_group(self._current_group_id)
        if not group:
            return

        templates = group.get('templates', [])
        if not templates:
            messagebox.showwarning('提示', '该模板组没有添加任何 Word 模板文件', parent=self.app.root)
            return

        # Collect values
        values = {}
        for field in group.get('fields', []):
            key = field['key']
            val = self._field_vars.get(key, tk.StringVar()).get()
            if field.get('required') and not val.strip():
                messagebox.showwarning('提示', f'请填写必填项：{field.get("label", key)}',
                                        parent=self.app.root)
                return
            values[key] = val

        # Build output folder
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_gname = ''.join(c for c in group['name'] if c.isalnum() or c in '._- ').strip()
        out_folder = os.path.join(self.app.storage.effective_output_dir,
                                   f'{ts}_{safe_gname}')
        os.makedirs(out_folder, exist_ok=True)

        errors = []
        filled = []
        for tmpl in templates:
            tpath = tmpl.get('path', '')
            tname = tmpl.get('name', os.path.basename(tpath))
            if not os.path.exists(tpath):
                errors.append(f'找不到模板文件：{tname}（{tpath}）')
                continue
            base = os.path.splitext(os.path.basename(tpath))[0]
            out_path = os.path.join(out_folder, f'{base}.docx')
            try:
                doc_filler.fill_document(tpath, values, out_path)
                filled.append(tname)
            except Exception as e:
                errors.append(f'{tname}: {e}')

        # Save history
        record = {
            'group_id': group['id'],
            'group_name': group['name'],
            'timestamp': datetime.now().isoformat(),
            'field_values': values,
            'output_folder': out_folder,
            'filled_count': len(filled),
            'errors': errors,
        }
        self.app.storage.save_history(record)
        self.app.history_tab.on_show()

        if errors:
            msg = f'成功生成 {len(filled)} 个文档，{len(errors)} 个错误：\n' + '\n'.join(errors)
            messagebox.showwarning('部分成功', msg, parent=self.app.root)
        else:
            self._status_label.config(text=f'✓ 已生成 {len(filled)} 个文档')
            self.app.root.after(4000, lambda: self._status_label.config(text=''))

        self._open_folder(out_folder)

    def _open_folder(self, path: str):
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
