import json
import os
import uuid
from pathlib import Path
from datetime import datetime
from typing import List

CONFIG_FILE = os.path.join(str(Path.home()), '.docfiller_config.json')
DEFAULT_DATA_DIR = str(Path.home() / 'DocFiller')


class Storage:
    def __init__(self):
        self.data_dir = DEFAULT_DATA_DIR
        self.output_dir = ''
        self._load_config()
        self._ensure_dirs()

    @property
    def effective_output_dir(self):
        return self.output_dir or os.path.join(self.data_dir, 'output')

    @property
    def word_templates_dir(self):
        return os.path.join(self.data_dir, 'word_templates')

    @property
    def groups_dir(self):
        return os.path.join(self.data_dir, 'template_groups')

    @property
    def history_dir(self):
        return os.path.join(self.data_dir, 'history')

    def _ensure_dirs(self):
        for d in [self.data_dir, self.effective_output_dir,
                  self.word_templates_dir, self.groups_dir, self.history_dir]:
            os.makedirs(d, exist_ok=True)

    def _load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                self.data_dir = cfg.get('data_dir', DEFAULT_DATA_DIR)
                self.output_dir = cfg.get('output_dir', '')
            except Exception:
                pass

    def save_config(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(
                {'data_dir': self.data_dir, 'output_dir': self.output_dir},
                f, ensure_ascii=False, indent=2
            )
        self._ensure_dirs()

    # --- Template Groups ---

    def save_group(self, group: dict):
        path = os.path.join(self.groups_dir, f"{group['id']}.json")
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(group, f, ensure_ascii=False, indent=2)

    def load_groups(self) -> List[dict]:
        groups = []
        if not os.path.exists(self.groups_dir):
            return groups
        for fname in os.listdir(self.groups_dir):
            if fname.endswith('.json'):
                try:
                    with open(os.path.join(self.groups_dir, fname), 'r', encoding='utf-8') as f:
                        groups.append(json.load(f))
                except Exception:
                    pass
        return sorted(groups, key=lambda g: g.get('name', ''))

    def delete_group(self, group_id: str):
        path = os.path.join(self.groups_dir, f"{group_id}.json")
        if os.path.exists(path):
            os.remove(path)

    # --- History ---

    def save_history(self, record: dict) -> str:
        if 'id' not in record:
            record['id'] = str(uuid.uuid4())
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        group_name = record.get('group_name', '未知')
        safe = ''.join(c for c in group_name if c.isalnum() or c in '._- ').strip()[:20]
        fname = f"{ts}_{safe}.json"
        path = os.path.join(self.history_dir, fname)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
        return path

    def load_history(self) -> List[dict]:
        records = []
        if not os.path.exists(self.history_dir):
            return records
        fnames = sorted(os.listdir(self.history_dir), reverse=True)
        for fname in fnames[:200]:
            if fname.endswith('.json'):
                try:
                    with open(os.path.join(self.history_dir, fname), 'r', encoding='utf-8') as f:
                        records.append(json.load(f))
                except Exception:
                    pass
        return records

    def delete_history_record(self, record_id: str):
        for fname in os.listdir(self.history_dir):
            if fname.endswith('.json'):
                path = os.path.join(self.history_dir, fname)
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    if data.get('id') == record_id:
                        os.remove(path)
                        return
                except Exception:
                    pass
