import uuid
from typing import List, Optional
from .storage import Storage


def new_group(name: str, description: str = '') -> dict:
    return {
        'id': str(uuid.uuid4())[:8],
        'name': name,
        'description': description,
        'fields': [],       # list of {key, label, field_type, required, hint}
        'templates': [],    # list of {name, path}
    }


def new_field(key: str, label: str = '', field_type: str = 'text',
              required: bool = True, hint: str = '') -> dict:
    return {
        'key': key,
        'label': label or key,
        'field_type': field_type,
        'required': required,
        'hint': hint,
    }


def new_template_file(name: str, path: str) -> dict:
    return {'name': name, 'path': path}


class TemplateManager:
    def __init__(self, storage: Storage):
        self._storage = storage
        self._groups: List[dict] = []
        self.reload()

    def reload(self):
        self._groups = self._storage.load_groups()

    @property
    def groups(self) -> List[dict]:
        return self._groups

    def get_group(self, group_id: str) -> Optional[dict]:
        return next((g for g in self._groups if g['id'] == group_id), None)

    def save_group(self, group: dict):
        self._storage.save_group(group)
        self.reload()

    def delete_group(self, group_id: str):
        self._storage.delete_group(group_id)
        self.reload()

    def group_names(self) -> List[str]:
        return [g['name'] for g in self._groups]
