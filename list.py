"""Дискавери и фильтрация файлов: сканер-потоки (shallow/mid/deep), реестр
найденных файлов, natural-sort и фоновый воркер поиска/фильтрации."""
import os
import re
from pathlib import Path

from PySide6.QtCore import QThread, Signal

try:
    from rapidfuzz import fuzz
except ImportError:
    fuzz = None

IGNORED_DIRS = {".git", ".svn", ".idea", ".vscode", "__pycache__", "$recycle.bin", "system volume information"}


class FileRegistry:
    def __init__(self):
        self.files: list[tuple[Path, str]] = []
        self.extensions: set[str] = set()
        self.ext_counts: dict[str, int] = {}

    def clear(self):
        self.files.clear()
        self.extensions.clear()
        self.ext_counts.clear()

    def add_files(self, new_files: list[tuple[Path, str]]) -> bool:
        new_ext_found = False
        for full_path, rel_path in new_files:
            self.files.append((full_path, rel_path))
            ext = Path(rel_path).suffix.lower()
            if not ext:
                continue
            self.ext_counts[ext] = self.ext_counts.get(ext, 0) + 1
            if ext not in self.extensions:
                self.extensions.add(ext)
                new_ext_found = True
        return new_ext_found

    def remove_files(self, full_paths_to_remove: set[str]):
        """Убирает файлы из реестра после удаления с диска (Delete),
        не трогая живущий scanner-поток — только in-memory состояние."""
        remaining = []
        for full_path, rel_path in self.files:
            if str(full_path) in full_paths_to_remove:
                ext = Path(rel_path).suffix.lower()
                if ext in self.ext_counts:
                    self.ext_counts[ext] -= 1
                    if self.ext_counts[ext] <= 0:
                        del self.ext_counts[ext]
                        self.extensions.discard(ext)
            else:
                remaining.append((full_path, rel_path))
        self.files = remaining


class ShallowScannerThread(QThread):
    files_found = Signal(list)

    def __init__(self, folder: str, is_recursive: bool = False, max_depth: int = 2):
        super().__init__()
        self.folder = folder
        self.is_recursive = is_recursive
        self.max_depth = max_depth
        self._is_cancelled = False

    def stop(self):
        self._is_cancelled = True

    def run(self):
        batch = []
        try:
            folder_path = Path(self.folder)
            for root, dirs, files in os.walk(self.folder):
                if self._is_cancelled:
                    return

                rel = Path(root).relative_to(folder_path)
                depth = len(rel.parts)

                if not self.is_recursive:
                    if depth > 0:
                        dirs.clear()
                        continue
                else:
                    dirs[:] = [d for d in dirs if d.lower() not in IGNORED_DIRS]
                    if depth > self.max_depth:
                        dirs.clear()
                        continue

                rel_dir = os.path.relpath(root, self.folder)
                rel_prefix = "" if rel_dir == "." else rel_dir
                for file_name in files:
                    if self._is_cancelled:
                        return
                    full_path = Path(root) / file_name
                    rel_path = os.path.join(rel_prefix, file_name)
                    batch.append((full_path, rel_path))

                    if len(batch) >= 100:
                        self.files_found.emit(batch)
                        batch = []

            if batch and not self._is_cancelled:
                self.files_found.emit(batch)
        except OSError:
            pass


class MidScannerThread(QThread):
    files_found = Signal(list)

    def __init__(self, folder: str, min_depth: int = 3, max_depth: int = 4):
        super().__init__()
        self.folder = folder
        self.min_depth = min_depth
        self.max_depth = max_depth
        self._is_cancelled = False

    def stop(self):
        self._is_cancelled = True

    def run(self):
        batch = []
        try:
            folder_path = Path(self.folder)
            for root, dirs, files in os.walk(self.folder):
                if self._is_cancelled:
                    return
                dirs[:] = [d for d in dirs if d.lower() not in IGNORED_DIRS]
                rel = Path(root).relative_to(folder_path)
                depth = len(rel.parts)

                if depth > self.max_depth:
                    dirs.clear()
                    continue
                if depth < self.min_depth:
                    continue

                rel_dir = os.path.relpath(root, self.folder)
                rel_prefix = "" if rel_dir == "." else rel_dir
                for file_name in files:
                    if self._is_cancelled:
                        return
                    full_path = Path(root) / file_name
                    rel_path = os.path.join(rel_prefix, file_name)
                    batch.append((full_path, rel_path))

                    if len(batch) >= 150:
                        self.files_found.emit(batch)
                        batch = []

            if batch and not self._is_cancelled:
                self.files_found.emit(batch)
        except OSError:
            pass


class DeepScannerThread(QThread):
    files_found = Signal(list)

    def __init__(self, folder: str, min_depth: int = 5):
        super().__init__()
        self.folder = folder
        self.min_depth = min_depth
        self._is_cancelled = False

    def stop(self):
        self._is_cancelled = True

    def run(self):
        batch = []
        try:
            folder_path = Path(self.folder)
            for root, dirs, files in os.walk(self.folder):
                if self._is_cancelled:
                    return
                dirs[:] = [d for d in dirs if d.lower() not in IGNORED_DIRS]
                rel = Path(root).relative_to(folder_path)
                depth = len(rel.parts)

                if depth < self.min_depth:
                    continue

                rel_dir = os.path.relpath(root, self.folder)
                rel_prefix = "" if rel_dir == "." else rel_dir
                for file_name in files:
                    if self._is_cancelled:
                        return
                    full_path = Path(root) / file_name
                    rel_path = os.path.join(rel_prefix, file_name)
                    batch.append((full_path, rel_path))

                    if len(batch) >= 250:
                        self.files_found.emit(batch)
                        batch = []

            if batch and not self._is_cancelled:
                self.files_found.emit(batch)
        except OSError:
            pass


_NATURAL_SORT_RE = re.compile(r'(\d+)')


def natural_sort_key(text: str):
    """'photo_2' должен идти перед 'photo_1000' — обычная строковая
    сортировка сравнивает посимвольно ('1' < '2' до дальнейших цифр),
    поэтому режем строку на числовые/нечисловые куски и сравниваем числа
    как числа."""
    return [int(chunk) if chunk.isdigit() else chunk.lower() for chunk in _NATURAL_SORT_RE.split(text)]


def fuzzy_match(tokens: list[str], target_str: str) -> bool:
    target_str = target_str.lower()
    for token in tokens:
        if not token:
            continue
        if token in target_str:
            continue
        if fuzz and fuzz.partial_ratio(token, target_str) > 75:
            continue
        return False
    return True


class FilterWorkerThread(QThread):
    """Считает набор видимых элементов (fuzzy-match + фильтр по расширениям)
    в фоновом потоке. Никаких Qt-виджетов тут не создаётся — только plain
    python данные, поэтому поток безопасен."""

    filtered_ready = Signal(list, list, int)

    def __init__(
        self,
        files: list[tuple[Path, str]],
        tokens: list[str],
        active_extensions: set[str],
        max_items: int,
        generation: int,
    ):
        super().__init__()
        self.files = files
        self.tokens = tokens
        self.active_extensions = active_extensions
        self.max_items = max_items
        self.generation = generation
        self._is_cancelled = False

    def stop(self):
        self._is_cancelled = True

    def run(self):
        # Собираем ВСЕ совпадения (без ранней остановки на max_items) — это
        # необходимо, чтобы natural-sort ниже был честным: нельзя обрезать
        # список до сортировки, иначе в видимых 500 может оказаться
        # случайный срез в порядке сканирования диска, а не первые по имени.
        matches = []  # (dir_part, file_name, full_path)
        if self.active_extensions:
            for full_path_obj, rel_path in self.files:
                if self._is_cancelled:
                    return
                file_name = os.path.basename(rel_path)
                extension = Path(file_name).suffix.lower()
                if extension not in self.active_extensions:
                    continue
                if self.tokens and not fuzzy_match(self.tokens, rel_path):
                    continue
                dir_part = os.path.dirname(rel_path)
                matches.append((dir_part, file_name, str(full_path_obj)))

        if self._is_cancelled:
            return

        matches.sort(key=lambda m: (natural_sort_key(m[0]), natural_sort_key(m[1])))

        if self._is_cancelled:
            return

        all_matched_paths = [full_path for _dir_part, _file_name, full_path in matches]
        visible_results = [
            (full_path, file_name, dir_part)
            for dir_part, file_name, full_path in matches[: self.max_items]
        ]

        if not self._is_cancelled:
            self.filtered_ready.emit(visible_results, all_matched_paths, self.generation)