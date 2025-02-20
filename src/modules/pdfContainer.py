import os
import shutil
import json
import zipfile


class PDFContainer:
    def __init__(self):
        self.container_file = "kvalSpec.dbf"
        self.deletion_file = "__marked_for_deletion.json"

        if not os.path.exists(self.container_file):
            with zipfile.ZipFile(self.container_file, 'w', compression=zipfile.ZIP_STORED) as zf:
                zf.writestr(self.deletion_file, json.dumps([]))

    def _load_deletion_marks(self):
        with zipfile.ZipFile(self.container_file, 'r') as zf:
            if self.deletion_file in zf.namelist():
                return set(json.loads(zf.read(self.deletion_file).decode()))
        return set()

    def _save_deletion_marks(self, marks):
        temp_path = f"{self.container_file}.tmp"
        with zipfile.ZipFile(self.container_file, 'r') as zf:
            files = {name: zf.read(name) for name in zf.namelist() if name != self.deletion_file}

        with zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_STORED) as temp_zf:
            for name, content in files.items():
                temp_zf.writestr(name, content)
            temp_zf.writestr(self.deletion_file, json.dumps(list(marks)))

        shutil.move(temp_path, self.container_file)

    def add_file(self, file_path: str, prikaz_dt: str = "2023-12-01", prikaz_nom: str = "123") -> str:
        if not os.path.isfile(file_path) or not file_path.endswith('.pdf'):
            raise ValueError("Файл должен быть PDF.")

        all_files = self.list_files()
        file_name = self._generate_unique_name(prikaz_dt, prikaz_nom, all_files)

        with zipfile.ZipFile(self.container_file, 'a', compression=zipfile.ZIP_STORED) as zf:
            zf.write(file_path, file_name)

        return file_name

    def _generate_unique_name(self, prikaz_dt: str, prikaz_nom: str, all_files: list) -> str:
        base_name = f"Приказ от {prikaz_dt} N {prikaz_nom}.pdf"
        if base_name not in all_files:
            return base_name

        counter = 1
        while True:
            unique_name = f"Приказ от {prikaz_dt} N {prikaz_nom}-({counter}).pdf"
            if unique_name not in all_files:
                return unique_name
            counter += 1

    def mark_for_deletion(self, file_name: str):
        marks = self._load_deletion_marks()
        marks.add(file_name)
        self._save_deletion_marks(marks)

    def unmark_for_deletion(self, file_name: str):
        marks = self._load_deletion_marks()
        if file_name in marks:
            marks.remove(file_name)
            self._save_deletion_marks(marks)

    def is_marked_for_deletion(self, file_name: str) -> bool:
        marks = self._load_deletion_marks()
        return file_name in marks

    def pack_container(self):
        marks = self._load_deletion_marks()
        temp_path = f"{self.container_file}.tmp"
        with zipfile.ZipFile(self.container_file, 'r') as zf:
            files = {name: zf.read(name) for name in zf.namelist() if name not in marks and name != self.deletion_file}

        with zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_STORED) as temp_zf:
            for name, content in files.items():
                temp_zf.writestr(name, content)
            temp_zf.writestr(self.deletion_file, json.dumps([]))

        shutil.move(temp_path, self.container_file)

    def extract_file(self, file_name: str, destination: str):
        with zipfile.ZipFile(self.container_file, 'r') as zf:
            if file_name not in zf.namelist():
                raise FileNotFoundError(f"Файл {file_name} не найден в контейнере.")
            with open(destination, 'wb') as f:
                f.write(zf.read(file_name))

    def list_files(self):
        with zipfile.ZipFile(self.container_file, 'r') as zf:
            return [f for f in zf.namelist() if f != self.deletion_file]