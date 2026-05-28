import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import copy

from ..core import doc_filler
from ..core.template_manager import new_group, new_field, new_template_file
from .main_window import COLORS, _cjk_font


FIELD_TYPES = [('text', '单行文本'), ('multiline', '多行文本'),
               ('date', '日期'), ('number', '数字')]
FIELD_TYPE_LABELS = {k: v for k, v in FIELD_TYPES}


class TemplateTab:
    def __init__(self, parent, app):
        self.app = app
        self.frame = ttk.Frame(parent)
        self._editing_group = None   # working copy of the group being edited

        self._build()

    def _build(self):
        pane = ttk.PanedWindow(self.frame, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        # ---- Left: group list ----
        left = ttk.Frame(pane, style='Surface.TFrame', padding=8)
        pane.add(left, weight=1)

        ttk.Label(left, text='模板组列表', style='Title.TLabel').pack(anchor='w', pady=(0, 8))

        list_frame = ttk.Frame(left)
        list_frame.pack(fill=tk.BOTH, expand=True)

        self._group_list = tk.Listbox(list_frame, font=_cjk_font(10),
                                       selectmode=tk.SINGLE, activestyle='none',
                                       bg=COLORS['bg'], fg=COLORS['text'],
                                       selectbackground=COLORS['primary'],
                                       selectforeground='white',
                                       relief='flat', bd=0, highlightthickness=0)
        sb = ttk.Scrollbar(list_frame, orient='vertical',
                            command=self._group_list.yview)
        self._group_list.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._group_list.pack(fill=tk.BOTH, expand=True)
        self._group_list.bind('<<ListboxSelect>>', self._on_group_select)

        btn_row = ttk.Frame(left)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text='+ 新建', style='Primary.TButton',
                   command=self._new_group).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_row, text='删除', style='Danger.TButton',
                   command=self._delete_group).pack(side=tk.LEFT)

        # ---- Right: edit panel ----
        right = ttk.Frame(pane, style='Surface.TFrame', padding=16)
        pane.add(right, weight=3)

        self._right_frame = right
        self._build_edit_panel(right)

    def _build_edit_panel(self, parent):
        self._edit_container = ttk.Frame(parent, style='Surface.TFrame')
        self._edit_container.pack(fill=tk.BOTH, expand=True)
        self._show_placeholder()

    def _show_placeholder(self):
        for w in self._edit_container.winfo_children():
            w.destroy()
        ttk.Label(self._edit_container,
                  text='← 选择或新建一个模板组',
                  style='Muted.TLabel', font=_cjk_font(12)).pack(expand=True)

    def _render_edit(self, group: dict):
        for w in self._edit_container.winfo_children():
            w.destroy()

        c = self._edit_container

        # Name / desc row
        info_frame = ttk.Frame(c, style='Surface.TFrame')
        info_frame.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(info_frame, text='组名称：', style='Surface.TLabel',
                  font=_cjk_font(10, 'bold')).grid(row=0, column=0, sticky='w', pady=4)
        self._name_var = tk.StringVar(value=group.get('name', ''))
        ttk.Entry(info_frame, textvariable=self._name_var, width=30,
                   font=_cjk_font(10)).grid(row=0, column=1, sticky='ew', padx=8, pady=4)

        ttk.Label(info_frame, text='说明：', style='Surface.TLabel').grid(
            row=1, column=0, sticky='w', pady=4)
        self._desc_var = tk.StringVar(value=group.get('description', ''))
        ttk.Entry(info_frame, textvariable=self._desc_var, width=42,
                   font=_cjk_font(10)).grid(row=1, column=1, sticky='ew', padx=8, pady=4)
        info_frame.columnconfigure(1, weight=1)

        sep = ttk.Frame(c, style='Separator.TFrame', height=1)
        sep.pack(fill=tk.X, pady=8)

        # Fields section
        fields_label_row = ttk.Frame(c, style='Surface.TFrame')
        fields_label_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(fields_label_row, text='字段定义', style='Surface.TLabel',
                  font=_cjk_font(10, 'bold')).pack(side=tk.LEFT)
        ttk.Button(fields_label_row, text='+ 添加字段', style='Secondary.TButton',
                   command=self._add_field).pack(side=tk.RIGHT)

        fields_frame = ttk.Frame(c, style='Surface.TFrame')
        fields_frame.pack(fill=tk.X, pady=(0, 4))

        self._fields_tree = ttk.Treeview(
            fields_frame,
            columns=('key', 'label', 'type', 'required'),
            show='headings', height=6, selectmode='browse'
        )
        self._fields_tree.heading('key', text='占位符键')
        self._fields_tree.heading('label', text='显示标签')
        self._fields_tree.heading('type', text='类型')
        self._fields_tree.heading('required', text='必填')
        self._fields_tree.column('key', width=120)
        self._fields_tree.column('label', width=120)
        self._fields_tree.column('type', width=80)
        self._fields_tree.column('required', width=50)

        fsb = ttk.Scrollbar(fields_frame, orient='vertical',
                             command=self._fields_tree.yview)
        self._fields_tree.configure(yscrollcommand=fsb.set)
        fsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._fields_tree.pack(fill=tk.X, expand=True)
        self._fields_tree.bind('<Double-1>', lambda e: self._edit_field())

        self._refresh_fields_tree()

        fields_btn_row = ttk.Frame(c, style='Surface.TFrame')
        fields_btn_row.pack(fill=tk.X, pady=4)
        ttk.Button(fields_btn_row, text='编辑', style='Secondary.TButton',
                   command=self._edit_field).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(fields_btn_row, text='上移', style='Secondary.TButton',
                   command=lambda: self._move_field(-1)).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(fields_btn_row, text='下移', style='Secondary.TButton',
                   command=lambda: self._move_field(1)).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(fields_btn_row, text='删除字段', style='Danger.TButton',
                   command=self._delete_field).pack(side=tk.LEFT, padx=(0, 4))

        sep2 = ttk.Frame(c, style='Separator.TFrame', height=1)
        sep2.pack(fill=tk.X, pady=8)

        # Templates section
        tmpl_label_row = ttk.Frame(c, style='Surface.TFrame')
        tmpl_label_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(tmpl_label_row, text='Word 模板文件', style='Surface.TLabel',
                  font=_cjk_font(10, 'bold')).pack(side=tk.LEFT)
        ttk.Button(tmpl_label_row, text='+ 添加 .docx', style='Secondary.TButton',
                   command=self._add_template).pack(side=tk.RIGHT)

        tmpl_frame = ttk.Frame(c, style='Surface.TFrame')
        tmpl_frame.pack(fill=tk.X, pady=(0, 4))

        self._tmpl_tree = ttk.Treeview(
            tmpl_frame,
            columns=('name', 'path'),
            show='headings', height=5, selectmode='browse'
        )
        self._tmpl_tree.heading('name', text='名称')
        self._tmpl_tree.heading('path', text='文件路径')
        self._tmpl_tree.column('name', width=160)
        self._tmpl_tree.column('path', width=300)

        tsb = ttk.Scrollbar(tmpl_frame, orient='vertical',
                             command=self._tmpl_tree.yview)
        self._tmpl_tree.configure(yscrollcommand=tsb.set)
        tsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._tmpl_tree.pack(fill=tk.X, expand=True)

        self._refresh_tmpl_tree()

        tmpl_btn_row = ttk.Frame(c, style='Surface.TFrame')
        tmpl_btn_row.pack(fill=tk.X, pady=4)
        ttk.Button(tmpl_btn_row, text='重命名', style='Secondary.TButton',
                   command=self._rename_template).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(tmpl_btn_row, text='扫描占位符 →自动添加字段', style='Secondary.TButton',
                   command=self._scan_placeholders).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(tmpl_btn_row, text='移除', style='Danger.TButton',
                   command=self._remove_template).pack(side=tk.LEFT)

        # Save button
        sep3 = ttk.Frame(c, style='Separator.TFrame', height=1)
        sep3.pack(fill=tk.X, pady=8)
        save_row = ttk.Frame(c, style='Surface.TFrame')
        save_row.pack(fill=tk.X)
        ttk.Button(save_row, text='保存模板组', style='Primary.TButton',
                   command=self._save_group).pack(side=tk.LEFT)
        self._save_status = ttk.Label(save_row, text='', style='Surface.TLabel',
                                       foreground=COLORS['success'])
        self._save_status.pack(side=tk.LEFT, padx=12)

    # ---------- Group list ----------

    def on_show(self):
        self.app.template_manager.reload()
        self._refresh_group_list()

    def _refresh_group_list(self):
        self._group_list.delete(0, tk.END)
        for g in self.app.template_manager.groups:
            self._group_list.insert(tk.END, f"  {g['name']}")

    def _on_group_select(self, event):
        sel = self._group_list.curselection()
        if not sel:
            return
        idx = sel[0]
        groups = self.app.template_manager.groups
        if idx >= len(groups):
            return
        self._editing_group = copy.deepcopy(groups[idx])
        self._render_edit(self._editing_group)

    def _new_group(self):
        name = _ask_string(self.app.root, '新建模板组', '请输入模板组名称：')
        if not name:
            return
        g = new_group(name)
        self._editing_group = g
        self._refresh_group_list()
        self._render_edit(g)

    def _delete_group(self):
        sel = self._group_list.curselection()
        if not sel:
            messagebox.showwarning('提示', '请先选择要删除的模板组', parent=self.app.root)
            return
        idx = sel[0]
        groups = self.app.template_manager.groups
        if idx >= len(groups):
            return
        g = groups[idx]
        if not messagebox.askyesno('确认删除', f'删除模板组「{g["name"]}」？此操作不可恢复。',
                                    parent=self.app.root):
            return
        self.app.template_manager.delete_group(g['id'])
        self._editing_group = None
        self._refresh_group_list()
        self._show_placeholder()

    # ---------- Fields ----------

    def _refresh_fields_tree(self):
        self._fields_tree.delete(*self._fields_tree.get_children())
        for f in self._editing_group.get('fields', []):
            ftype_label = FIELD_TYPE_LABELS.get(f.get('field_type', 'text'), f.get('field_type', ''))
            req = '✓' if f.get('required', True) else ''
            self._fields_tree.insert('', tk.END, iid=f['key'],
                                      values=(f['key'], f.get('label', f['key']),
                                              ftype_label, req))

    def _add_field(self):
        self._open_field_dialog(None)

    def _edit_field(self):
        sel = self._fields_tree.selection()
        if not sel:
            messagebox.showwarning('提示', '请先选择字段', parent=self.app.root)
            return
        key = sel[0]
        fields = self._editing_group.get('fields', [])
        field = next((f for f in fields if f['key'] == key), None)
        if field:
            self._open_field_dialog(field)

    def _delete_field(self):
        sel = self._fields_tree.selection()
        if not sel:
            return
        key = sel[0]
        fields = self._editing_group.get('fields', [])
        self._editing_group['fields'] = [f for f in fields if f['key'] != key]
        self._refresh_fields_tree()

    def _move_field(self, direction: int):
        sel = self._fields_tree.selection()
        if not sel:
            return
        key = sel[0]
        fields = self._editing_group.get('fields', [])
        idx = next((i for i, f in enumerate(fields) if f['key'] == key), None)
        if idx is None:
            return
        new_idx = idx + direction
        if 0 <= new_idx < len(fields):
            fields[idx], fields[new_idx] = fields[new_idx], fields[idx]
            self._editing_group['fields'] = fields
            self._refresh_fields_tree()
            self._fields_tree.selection_set(key)

    def _open_field_dialog(self, existing_field):
        win = tk.Toplevel(self.app.root)
        win.title('编辑字段' if existing_field else '添加字段')
        win.geometry('400x300')
        win.resizable(False, False)
        win.configure(bg=COLORS['bg'])
        win.grab_set()
        win.transient(self.app.root)

        f = ttk.Frame(win, style='Surface.TFrame', padding=20)
        f.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        ttk.Label(f, text='占位符键 (如: 姓名)', style='Surface.TLabel').grid(
            row=0, column=0, sticky='w', pady=6)
        key_var = tk.StringVar(value=existing_field['key'] if existing_field else '')
        key_entry = ttk.Entry(f, textvariable=key_var, width=24, font=_cjk_font(10))
        key_entry.grid(row=0, column=1, sticky='ew', padx=8, pady=6)

        ttk.Label(f, text='显示标签', style='Surface.TLabel').grid(
            row=1, column=0, sticky='w', pady=6)
        label_var = tk.StringVar(value=existing_field.get('label', '') if existing_field else '')
        ttk.Entry(f, textvariable=label_var, width=24, font=_cjk_font(10)).grid(
            row=1, column=1, sticky='ew', padx=8, pady=6)

        ttk.Label(f, text='字段类型', style='Surface.TLabel').grid(
            row=2, column=0, sticky='w', pady=6)
        type_var = tk.StringVar(value=existing_field.get('field_type', 'text') if existing_field else 'text')
        type_combo = ttk.Combobox(f, textvariable=type_var, state='readonly', width=14)
        type_combo['values'] = [v for _, v in FIELD_TYPES]
        current_label = FIELD_TYPE_LABELS.get(type_var.get(), '单行文本')
        type_var.set(current_label)
        type_combo.grid(row=2, column=1, sticky='w', padx=8, pady=6)

        ttk.Label(f, text='提示文字', style='Surface.TLabel').grid(
            row=3, column=0, sticky='w', pady=6)
        hint_var = tk.StringVar(value=existing_field.get('hint', '') if existing_field else '')
        ttk.Entry(f, textvariable=hint_var, width=24, font=_cjk_font(10)).grid(
            row=3, column=1, sticky='ew', padx=8, pady=6)

        req_var = tk.BooleanVar(value=existing_field.get('required', True) if existing_field else True)
        ttk.Checkbutton(f, text='必填', variable=req_var,
                         style='TCheckbutton').grid(
            row=4, column=1, sticky='w', padx=8, pady=6)

        f.columnconfigure(1, weight=1)

        type_key_map = {v: k for k, v in FIELD_TYPES}

        def save():
            k = key_var.get().strip()
            lbl = label_var.get().strip()
            if not k:
                messagebox.showwarning('提示', '占位符键不能为空', parent=win)
                return
            ftype = type_key_map.get(type_var.get(), 'text')

            fields = self._editing_group.get('fields', [])
            if existing_field:
                # Update existing
                old_key = existing_field['key']
                for field in fields:
                    if field['key'] == old_key:
                        field.update({'key': k, 'label': lbl or k,
                                      'field_type': ftype, 'required': req_var.get(),
                                      'hint': hint_var.get().strip()})
                        break
            else:
                if any(f['key'] == k for f in fields):
                    messagebox.showwarning('提示', f'键值「{k}」已存在', parent=win)
                    return
                fields.append(new_field(k, lbl or k, ftype, req_var.get(),
                                         hint_var.get().strip()))
            self._editing_group['fields'] = fields
            self._refresh_fields_tree()
            win.destroy()

        btn_row = ttk.Frame(f)
        btn_row.grid(row=5, column=0, columnspan=2, pady=(12, 0), sticky='e')
        ttk.Button(btn_row, text='取消', style='Secondary.TButton',
                   command=win.destroy).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text='确认', style='Primary.TButton',
                   command=save).pack(side=tk.LEFT, padx=4)

        key_entry.focus_set()

    # ---------- Templates ----------

    def _refresh_tmpl_tree(self):
        self._tmpl_tree.delete(*self._tmpl_tree.get_children())
        for i, t in enumerate(self._editing_group.get('templates', [])):
            self._tmpl_tree.insert('', tk.END, iid=str(i),
                                    values=(t['name'], t['path']))

    def _add_template(self):
        paths = filedialog.askopenfilenames(
            title='选择 Word 模板文件',
            filetypes=[('Word 文档', '*.docx'), ('所有文件', '*.*')],
            parent=self.app.root
        )
        if not paths:
            return
        templates = self._editing_group.get('templates', [])
        for p in paths:
            name = os.path.splitext(os.path.basename(p))[0]
            if not any(t['path'] == p for t in templates):
                templates.append(new_template_file(name, p))
        self._editing_group['templates'] = templates
        self._refresh_tmpl_tree()

    def _rename_template(self):
        sel = self._tmpl_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        templates = self._editing_group.get('templates', [])
        if idx >= len(templates):
            return
        old_name = templates[idx]['name']
        new_name = _ask_string(self.app.root, '重命名', '请输入新名称：', initial=old_name)
        if new_name:
            templates[idx]['name'] = new_name
            self._editing_group['templates'] = templates
            self._refresh_tmpl_tree()

    def _remove_template(self):
        sel = self._tmpl_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        templates = self._editing_group.get('templates', [])
        if idx < len(templates):
            del templates[idx]
            self._editing_group['templates'] = templates
            self._refresh_tmpl_tree()

    def _scan_placeholders(self):
        sel = self._tmpl_tree.selection()
        if not sel:
            messagebox.showwarning('提示', '请先选择要扫描的模板文件', parent=self.app.root)
            return
        idx = int(sel[0])
        templates = self._editing_group.get('templates', [])
        if idx >= len(templates):
            return
        tpath = templates[idx]['path']
        if not os.path.exists(tpath):
            messagebox.showerror('错误', f'文件不存在：{tpath}', parent=self.app.root)
            return
        try:
            keys = doc_filler.extract_placeholders(tpath)
        except Exception as e:
            messagebox.showerror('扫描失败', str(e), parent=self.app.root)
            return

        if not keys:
            messagebox.showinfo('扫描结果', '该文档中未发现 {{占位符}}', parent=self.app.root)
            return

        existing_keys = {f['key'] for f in self._editing_group.get('fields', [])}
        added = []
        for k in keys:
            if k not in existing_keys:
                self._editing_group.setdefault('fields', []).append(
                    new_field(k, k))
                added.append(k)

        self._refresh_fields_tree()
        if added:
            messagebox.showinfo('扫描结果',
                                f'发现 {len(keys)} 个占位符，新增 {len(added)} 个字段：\n' +
                                '、'.join(added),
                                parent=self.app.root)
        else:
            messagebox.showinfo('扫描结果',
                                f'发现 {len(keys)} 个占位符，全部已存在于字段列表中',
                                parent=self.app.root)

    # ---------- Save ----------

    def _save_group(self):
        if not self._editing_group:
            return
        name = self._name_var.get().strip()
        if not name:
            messagebox.showwarning('提示', '请输入模板组名称', parent=self.app.root)
            return
        self._editing_group['name'] = name
        self._editing_group['description'] = self._desc_var.get().strip()
        self.app.template_manager.save_group(self._editing_group)
        self._refresh_group_list()
        self._save_status.config(text='✓ 已保存')
        self.app.root.after(3000, lambda: self._save_status.config(text=''))
        # Sync form tab group list
        self.app.form_tab.on_show()


def _ask_string(parent, title: str, prompt: str, initial: str = '') -> str:
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry('360x130')
    win.resizable(False, False)
    win.configure(bg=COLORS['bg'])
    win.grab_set()
    win.transient(parent)

    result = tk.StringVar()

    f = ttk.Frame(win, style='Surface.TFrame', padding=16)
    f.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    ttk.Label(f, text=prompt, style='Surface.TLabel').pack(anchor='w')
    entry = ttk.Entry(f, font=_cjk_font(10), width=36)
    entry.insert(0, initial)
    entry.pack(fill=tk.X, pady=8)
    entry.select_range(0, tk.END)
    entry.focus_set()

    btn_row = ttk.Frame(f, style='Surface.TFrame')
    btn_row.pack(anchor='e')

    def ok():
        result.set(entry.get().strip())
        win.destroy()

    entry.bind('<Return>', lambda e: ok())
    ttk.Button(btn_row, text='取消', style='Secondary.TButton',
               command=win.destroy).pack(side=tk.LEFT, padx=4)
    ttk.Button(btn_row, text='确认', style='Primary.TButton',
               command=ok).pack(side=tk.LEFT, padx=4)

    win.wait_window()
    return result.get()
