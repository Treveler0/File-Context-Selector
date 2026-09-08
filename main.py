import csv
import ctypes
import os
import re
import subprocess
import sys
import time
import traceback
import winreg
from ctypes import wintypes
from pathlib import Path
import win32com.client

from PySide6.QtCore import (
    QAbstractNativeEventFilter,
    QEvent,
    QFileInfo,
    QModelIndex,
    QRect,
    QSize,
    Qt,
    QThread,
    QTimer,
    Signal,
)
from PySide6.QtGui import (
    QAction,
    QColor,
    QCursor,
    QIcon,
    QKeySequence,
    QPainter,
    QShortcut,
)
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import (
    QApplication,
    QFileIconProvider,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QPushButton,
    QSizePolicy,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QSystemTrayIcon,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


try:
    from rapidfuzz import fuzz
except ImportError:
    fuzz = None

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(APP_DIR)

APP_VERSION = 1.50
APP_ID = "FileContextSelector.SingleInstance"
HOTKEY_ID = 1
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
VK_SPACE = 0x20
REG_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_NAME = "FileContextSelector"
IGNORED_DIRS = {".git", ".svn", ".idea", ".vscode", "__pycache__", "$recycle.bin", "system volume information"}
MAX_VISIBLE_ITEMS = 500
FILTER_DEBOUNCE_MS = 80
EXT_GRID_COLUMNS = 6
EXT_ROWS_THRESHOLD = 4
EXT_DROPDOWN_THRESHOLD = EXT_ROWS_THRESHOLD * EXT_GRID_COLUMNS  # > 24 расширений

icon_provider = QFileIconProvider()
icon_cache: dict[str, QIcon] = {}


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", POINT),
    ]


class NativeHotkeyFilter(QAbstractNativeEventFilter):
    WM_HOTKEY = 0x0312

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def nativeEventFilter(self, eventType, message):
        if eventType in (b"windows_generic_MSG", b"windows_dispatcher_MSG", "windows_generic_MSG", "windows_dispatcher_MSG"):
            msg = ctypes.cast(int(message), ctypes.POINTER(MSG)).contents
            if msg.message == self.WM_HOTKEY and msg.wParam == HOTKEY_ID:
                self.callback()
                return True, 0
        return False, 0


def get_desktop_path() -> str:
    registry_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path) as key:
            value, _ = winreg.QueryValueEx(key, "Desktop")
            expanded_path = os.path.expandvars(value)
            if os.path.isdir(expanded_path):
                return expanded_path
    except OSError:
        pass
    return str(Path.home() / "Desktop")


def get_file_icon(file_path: str) -> QIcon:
    ext = Path(file_path).suffix.lower()
    if ext in (".exe", ".lnk", ".url"):
        if file_path not in icon_cache:
            icon = None
            if ext == ".url":
                try:
                    import configparser
                    config = configparser.ConfigParser(interpolation=None)
                    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                        config.read_file(f)
                    if "InternetShortcut" in config:
                        icon_file = config["InternetShortcut"].get("IconFile")
                        if icon_file:
                            clean_icon_path = icon_file.strip('"').strip()
                            if os.path.exists(clean_icon_path):
                                icon_info = QFileInfo(clean_icon_path)
                                icon = icon_provider.icon(icon_info)
                except Exception:
                    pass
            if icon is None or icon.isNull():
                file_info = QFileInfo(file_path)
                icon = icon_provider.icon(file_info)
            icon_cache[file_path] = icon
        return icon_cache[file_path]
    if ext not in icon_cache:
        file_info = QFileInfo(file_path)
        icon_cache[ext] = icon_provider.icon(file_info)
    return icon_cache[ext]


def get_active_explorer_path() -> str | None:
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None
    class_name_buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_name_buf, 256)
    class_name = class_name_buf.value
    if class_name in ("Progman", "WorkerW"):
        return get_desktop_path()
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        for window in shell.Windows():
            try:
                if int(window.HWND) != int(hwnd):
                    continue
                full_name = str(window.FullName).lower()
                if not full_name.endswith("\\explorer.exe"):
                    continue
                path = str(window.Document.Folder.Self.Path)
                if path and os.path.isdir(path):
                    return path
            except Exception:
                continue
    except Exception:
        pass

    return None


class FileItemDelegate(QStyledItemDelegate):
    """Делегат для быстрой отрисовки имени и относительного пути без QWidget контейнеров"""
    def __init__(self, parent=None):
        super().__init__(parent)

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex):
        painter.save()
        
        # 1. Отрисовка фона (Выделение / Фокус клавиатуры / Наведение мыши)
        is_selected = bool(option.state & QStyle.State_Selected)
        is_focused = bool(option.state & QStyle.State_HasFocus)
        is_hovered = bool(option.state & QStyle.State_MouseOver)

        if is_selected:
            painter.fillRect(option.rect, QColor("#1976d2"))
        elif is_hovered:
            painter.fillRect(option.rect, QColor("#2b2c30"))
        else:
            painter.fillRect(option.rect, QColor("#202124"))

        # Отрисовка рамки фокуса для визуализации текущей позиции курсора
        if is_focused and not is_selected:
            painter.setPen(QColor("#4a90e2"))
            painter.drawRect(option.rect.adjusted(0, 0, -1, -1))

        rect = option.rect.adjusted(8, 0, -10, 0)
        
        # 2. Иконка
        icon = index.data(Qt.DecorationRole)
        if icon and not icon.isNull():
            icon_rect = QRect(rect.left(), rect.top() + (rect.height() - 22) // 2, 22, 22)
            icon.paint(painter, icon_rect)
            rect.setLeft(rect.left() + 30)

        # Данные из элементов
        file_name = index.data(Qt.DisplayRole) or ""
        dir_part = index.data(Qt.UserRole + 1) or ""

        # 3. Имя файла
        painter.setFont(option.font)
        if is_selected:
            painter.setPen(QColor("#ffffff"))
        else:
            painter.setPen(QColor("#eeeeee"))
            
        metrics = painter.fontMetrics()
        name_width = metrics.horizontalAdvance(file_name)
        name_rect = QRect(rect.left(), rect.top(), name_width, rect.height())
        painter.drawText(name_rect, Qt.AlignVCenter | Qt.AlignLeft, file_name)

        # 4. Путь к папке (справа)
        if dir_part:
            dir_font = option.font
            dir_font.setPointSize(8)
            painter.setFont(dir_font)
            painter.setPen(QColor("#9aa0a6"))
            
            dir_rect = QRect(rect.left() + name_width + 10, rect.top(), rect.width() - name_width - 10, rect.height())
            painter.drawText(dir_rect, Qt.AlignVCenter | Qt.AlignRight, dir_part)

        painter.restore()

    def sizeHint(self, option: QStyleOptionViewItem, index: QModelIndex) -> QSize:
        return QSize(option.rect.width(), 34)


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


def find_associated_executable(extension: str) -> str | None:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            rf"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{extension}\UserChoice",
        ) as key:
            prog_id, _ = winreg.QueryValueEx(key, "ProgId")
    except OSError:
        prog_id = None
    candidates = []
    if prog_id:
        candidates.append((winreg.HKEY_CURRENT_USER, rf"Software\Classes\{prog_id}\shell\open\command"))
        candidates.append((winreg.HKEY_LOCAL_MACHINE, rf"Software\Classes\{prog_id}\shell\open\command"))
    candidates += [
        (winreg.HKEY_CURRENT_USER, rf"Software\Classes\{extension}\shell\open\command"),
        (winreg.HKEY_LOCAL_MACHINE, rf"Software\Classes\{extension}\shell\open\command"),
    ]
    command = None
    for hive, key_path in candidates:
        try:
            with winreg.OpenKey(hive, key_path) as key:
                command, _ = winreg.QueryValueEx(key, None)
                if command:
                    break
        except OSError:
            continue
    if not command:
        return None
    match = re.match(r'\s*"([^"]+\.exe)"', command, re.I)
    if not match:
        match = re.match(r'\s*([^\s]+\.exe)', command, re.I)
    return match.group(1) if match else None


IMAGE_EXTENSIONS_FOR_PHOTO_VIEWER = {
    ".bmp", ".dib", ".gif", ".ico", ".jfif", ".jpe", ".jpeg", ".jpg",
    ".png", ".tif", ".tiff", ".wdp", ".webp",
}


def _find_photo_viewer_dll() -> str | None:
    """Ищет классический Windows Photo Viewer (PhotoViewer.dll). Он физически
    остаётся в системе на Windows 10/11 (просто не зарегистрирован как
    приложение по умолчанию), поэтому проверяем оба варианта Program Files."""
    for env_var in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env_var)
        if not base:
            continue
        candidate = os.path.join(base, "Windows Photo Viewer", "PhotoViewer.dll")
        if os.path.isfile(candidate):
            return candidate
    return None


PHOTO_VIEWER_DLL_PATH = _find_photo_viewer_dll()


def mask_header(text: str) -> str:
    """Буква -> 'A', цифра -> '9', остальное (пробелы, знаки) — как есть."""
    return "".join("9" if ch.isdigit() else ("A" if ch.isalpha() else ch) for ch in text)


def export_selected_to_csv(paths: list[str], folder: str, include_paths: bool) -> str | None:
    """Экспортирует paths в <folder>/selected_list.csv.
    Разделитель — ';' (дефолтный список-разделитель Excel на ru-RU локали),
    квотирование — штатное через модуль csv (RFC 4180): поле, где встретится
    ';' или кавычка, автоматически оборачивается в кавычки — так что запятая
    ИЛИ ';' в имени файла ничего не сломает. Первая строка — не читаемый
    заголовок, а его маска (буква -> A, цифра -> 9)."""
    if not paths or not folder:
        return None

    if include_paths:
        header = "Путь"
        rows = [[path] for path in paths]
    else:
        header = "Имя файла"
        rows = [[os.path.basename(path)] for path in paths]

    output_path = os.path.join(folder, "selected_list.csv")
    try:
        with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";", quoting=csv.QUOTE_MINIMAL)
            writer.writerow([mask_header(header)])
            writer.writerows(rows)
        return output_path
    except OSError as e:
        print(f"Ошибка экспорта CSV: {e}")
        return None


def reveal_files_in_explorer(directory: str, file_paths: list[str]):
    """Открывает окно Проводника в directory с выделением ВСЕХ file_paths разом
    (все должны лежать непосредственно в этой директории — группировку по
    папкам делает вызывающий код). Это единственный способ гарантированно
    получить рабочую навигацию вперёд/назад в Photos/Photo Viewer — они её
    показывают только когда открыты настоящим двойным кликом из Explorer, а
    не запущены сторонним процессом.

    Основной путь — SHOpenFolderAndSelectItems (COM shell API, тот же
    механизм, которым пользуется сам Explorer для "выделить и показать").
    При любой проблеме с ним — откат на простой `explorer /select,`,
    который гарантированно работает, но выделяет только первый файл."""
    try:
        from win32com.shell import shell

        folder_pidl, _ = shell.SHParseDisplayName(directory, 0, None)
        desktop = shell.SHGetDesktopFolder()
        shell_folder = desktop.BindToObject(folder_pidl, None, shell.IID_IShellFolder)

        item_pidls = []
        for path in file_paths:
            name = os.path.basename(path)
            try:
                item_pidls.append(shell_folder.ParseDisplayName(0, None, name)[1])
            except Exception:
                print(f"Не удалось найти {name} в {directory}:")
                traceback.print_exc()

        if item_pidls:
            shell.SHOpenFolderAndSelectItems(folder_pidl, item_pidls, 0)
            return
    except Exception:
        # Тут была слишком общая обработка: один except на весь блок не
        # позволяет понять, какая именно COM-строка упала (в pywin32
        # win32com.shell сигнатуры плохо документированы, и предыдущие два
        # захода на угадывание индекса аргумента дали противоречивые ошибки
        # — значит, падали РАЗНЫЕ строки, а не одна и та же). Печатаем полный
        # traceback с номером строки, чтобы в следующий раз чинить прицельно,
        # а не гадать ещё раз.
        print(f"reveal_files_in_explorer: сбой для {directory}")
        traceback.print_exc()

    # --- Fallback: гарантированно рабочий, но только на первый файл группы ---
    if file_paths:
        try:
            subprocess.Popen(["explorer.exe", "/select,", file_paths[0]], close_fds=True)
        except OSError as e:
            print(f"Ошибка открытия папки {directory}: {e}")
    else:
        os.startfile(directory)


def open_selected_files(paths: list[str]):
    if not paths:
        return
    extension = Path(paths[0]).suffix.lower()
    if extension == ".url":
        for path in paths:
            try:
                os.startfile(path)
                time.sleep(0.05)
            except OSError as e:
                print(f"Ошибка открытия URL-файла {path}: {e}")
        return

    # --- Fallback для изображений (по той же логике, что и notepad-fallback ниже) ---
    # Приложение "Фотографии" в Windows 10/11 не показывает стрелки next/prev,
    # если запущено не из Explorer, а из стороннего процесса — это ограничение
    # самого приложения (подтверждённое сообществом), а не наш баг. Обходим:
    # для стандартных растровых форматов открываем классический Windows Photo
    # Viewer напрямую через его DLL — он сканирует папку сам и умеет листать
    # соседние файлы вне зависимости от того, кто его вызвал. Ассоциации по
    # умолчанию в системе при этом не трогаем — вызывается сама библиотека,
    # а не ProgId по умолчанию, поэтому не требует прав администратора и
    # изменений в реестре.
    if extension in IMAGE_EXTENSIONS_FOR_PHOTO_VIEWER and PHOTO_VIEWER_DLL_PATH:
        # Каждый процесс rundll32+PhotoViewer.dll — довольно тяжёлый (не только
        # по памяти, похоже, ещё и по GDI/USER-хендлам), и уже 2-3 одновременных
        # экземпляра могут упереться в "Не хватает оперативной памяти" даже при
        # свободной физической памяти. Поэтому принципиально открываем только
        # ОДИН экземпляр — на первый выбранный файл, а не по процессу на каждый.
        path = paths[0]
        try:
            # Внимание: без пробела перед "ImageView_Fullscreen" и без
            # кавычек вокруг самого пути — это единственный подтверждённый
            # рабочий синтаксис вызова данной точки входа DLL.
            command_line = f'rundll32.exe "{PHOTO_VIEWER_DLL_PATH}",ImageView_Fullscreen {path}'
            subprocess.Popen(command_line, close_fds=True)
        except OSError as e:
            print(f"Ошибка открытия изображения {path}: {e}")
        return

    exe = find_associated_executable(extension)
    is_notepad = exe and "notepad.exe" in exe.lower()

    if is_notepad:
        for path in paths:
            subprocess.Popen([exe, path], close_fds=True)
            time.sleep(0.05)
    elif exe and os.path.isfile(exe):
        subprocess.Popen([exe, *paths], close_fds=True)
    else:
        for path in paths:
            try:
                os.startfile(path)
                time.sleep(0.05)
            except OSError as e:
                print(f"Ошибка открытия файла {path}: {e}")


def is_autostart_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_RUN_KEY, 0, winreg.KEY_READ) as key:
            winreg.QueryValueEx(key, APP_NAME)
            return True
    except OSError:
        return False


def set_autostart(enable: bool):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
            if enable:
                if getattr(sys, "frozen", False):
                    app_path = f'"{os.path.abspath(sys.executable)}"'
                else:
                    python_exe = sys.executable
                    if python_exe.endswith("python.exe"):
                        python_exe = python_exe.replace("python.exe", "pythonw.exe")
                    app_path = f'"{python_exe}" "{os.path.abspath(sys.argv[0])}"'

                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, app_path)
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except OSError:
                    pass
    except OSError as e:
        print(f"Failed to set autostart: {e}", file=sys.stderr)


class SearchLineEdit(QLineEdit):
    def __init__(self, target_list: QListWidget, parent=None):
        super().__init__(parent)
        self.target_list = target_list

    def keyPressEvent(self, event):
        key = event.key()
        if key in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Down):
            if self.target_list.count() > 0:
                self.target_list.setFocus()
                if self.target_list.currentRow() < 0:
                    self.target_list.setCurrentRow(0)
            return
        if key == Qt.Key_Tab:
            self.target_list.setFocus()
            if self.target_list.count() > 0 and self.target_list.currentRow() < 0:
                self.target_list.setCurrentRow(0)
            return
        super().keyPressEvent(event)


class FileListWidget(QListWidget):
    def keyPressEvent(self, event):
        key = event.key()
        if key in (Qt.Key_Left, Qt.Key_Right):
            event.ignore()
            return
        if key == Qt.Key_Up and self.currentRow() <= 0:
            window = self.window()
            if hasattr(window, "search"):
                window.search.setFocus()
            return
        if key == Qt.Key_Tab or (key == Qt.Key_Backtab and event.modifiers() & Qt.ShiftModifier):
            window = self.window()
            if hasattr(window, "search"):
                window.search.setFocus()
            return

        super().keyPressEvent(event)


class CheckableMenu(QMenu):
    """QMenu, который не закрывается при клике по чекаемому пункту —
    так в дропдауне расширений можно отметить сразу несколько без
    повторного открытия меню."""

    def mouseReleaseEvent(self, event):
        action = self.activeAction()
        if action is not None and action.isCheckable():
            action.trigger()
            event.accept()
            return
        super().mouseReleaseEvent(event)


class SelectorWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.folder = None
        self.active_extensions = set()
        self.selected_paths = set()
        self.available_extensions = []
        self.global_preferred_extensions = None
        self.is_recursive = False

        self.registry = FileRegistry()
        self.shallow_scanner = None
        self.mid_scanner = None
        self.deep_scanner = None

        self.ui_update_timer = QTimer(self)
        self.ui_update_timer.setSingleShot(True)
        self.ui_update_timer.setInterval(100)
        self.ui_update_timer.timeout.connect(self._on_ui_update_tick)
        self.pending_ui_refresh = False

        # --- Асинхронная фильтрация списка ---
        self.filter_worker: FilterWorkerThread | None = None
        self.filter_generation = 0
        self.all_matched_paths: list[str] = []
        self.filter_debounce_timer = QTimer(self)
        self.filter_debounce_timer.setSingleShot(True)
        self.filter_debounce_timer.setInterval(FILTER_DEBOUNCE_MS)
        self.filter_debounce_timer.timeout.connect(self._run_filter_now)

        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        self.setMinimumSize(700, 520)
        self.resize(760, 620)
        self.build_ui()
        self.setup_shortcuts()

    def setup_shortcuts(self):
        QShortcut(QKeySequence("Return"), self, activated=self.open_selected)
        QShortcut(QKeySequence("Enter"), self, activated=self.open_selected)
        QShortcut(QKeySequence("Shift+Return"), self, activated=self.open_folder_for_selected)
        QShortcut(QKeySequence("Shift+Enter"), self, activated=self.open_folder_for_selected)
        QShortcut(QKeySequence("Escape"), self, activated=self.hide)
        QShortcut(QKeySequence("Ctrl+A"), self, activated=self.select_all)
        QShortcut(QKeySequence("Alt+R"), self, activated=self.toggle_recursive)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Left:
            self.switch_tab_relative(-1)
            return
        elif event.key() == Qt.Key_Right:
            self.switch_tab_relative(1)
            return
        super().keyPressEvent(event)

    def changeEvent(self, event):
        if event.type() == QEvent.ActivationChange and not self.isActiveWindow():
            self.hide()
        super().changeEvent(event)

    def stop_scanners(self):
        self.ui_update_timer.stop()
        self.filter_debounce_timer.stop()
        if self.filter_worker and self.filter_worker.isRunning():
            self.filter_worker.stop()
            self.filter_worker.wait(200)
        self.filter_worker = None
        for scanner_attr in ("shallow_scanner", "mid_scanner", "deep_scanner"):
            scanner = getattr(self, scanner_attr)
            if scanner and scanner.isRunning():
                scanner.stop()
                scanner.wait(200)
                setattr(self, scanner_attr, None)

    def build_ui(self):
        self.setStyleSheet("""
            QWidget {
                background: #202124;
                color: #eeeeee;
                font-family: "Segoe UI";
                font-size: 10pt;
            }
            QLineEdit {
                background: #303134;
                border: 1px solid #4a4b4e;
                border-radius: 8px;
                padding: 9px 12px;
            }
            QListWidget {
                background: #202124;
                border: none;
                outline: none;
            }
            QPushButton {
                background: #303134;
                border: 1px solid #4a4b4e;
                border-radius: 7px;
                padding: 7px 14px;
            }
            QPushButton:hover {
                background: #3a3b3f;
            }
        """)
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 14)
        root.setSpacing(10)
        self.title = QLabel("File Selector")
        self.title.setStyleSheet("font-size: 12pt; font-weight: 600;")
        path_row = QHBoxLayout()
        self.path_label = QLabel("")
        self.path_label.setStyleSheet("color: #a9aaad;")
        self.recursive_button = QPushButton("Recursive (Alt+R)")
        self.recursive_button.setCheckable(True)
        self.recursive_button.setFocusPolicy(Qt.NoFocus)
        self.recursive_button.setToolTip("Toggle Recursive Search (Alt+R)")
        self.recursive_button.clicked.connect(self.toggle_recursive)
        path_row.addWidget(self.path_label, 1)
        path_row.addWidget(self.recursive_button, 0, Qt.AlignRight)
        self.ext_container = QVBoxLayout()
        self.ext_container.setSpacing(6)
        
        self.list = FileListWidget()
        self.list.setItemDelegate(FileItemDelegate(self.list))
        self.list.setSelectionMode(QListWidget.MultiSelection)
        self.list.setFocusPolicy(Qt.StrongFocus)
        self.list.itemDoubleClicked.connect(lambda _: self.open_selected())
        
        self.search = SearchLineEdit(self.list)
        self.search.setPlaceholderText("Fuzzy search files or directories (tokens: space, comma, dot)...")
        self.search.textChanged.connect(self.on_search_text_changed)
        bottom = QHBoxLayout()
        self.count_label = QLabel("0 selected")
        self.count_label.setStyleSheet("color: #a9aaad;")
        self.open_folder_button = QPushButton("Open Folder")
        self.open_folder_button.setFocusPolicy(Qt.NoFocus)
        self.open_folder_button.setToolTip("Shift+Enter")
        self.open_folder_button.clicked.connect(self.open_folder_for_selected)
        self.open_button = QPushButton("Open")
        self.open_button.setFocusPolicy(Qt.NoFocus)
        self.open_button.clicked.connect(self.open_selected)
        self.export_list_button = QPushButton("Export List")
        self.export_list_button.setFocusPolicy(Qt.NoFocus)
        self.export_list_button.setToolTip("Экспорт имени+расширения выбранных файлов в selected_list.csv")
        self.export_list_button.clicked.connect(lambda: self.export_list_for_selected(include_paths=False))
        self.export_paths_button = QPushButton("Export with Paths")
        self.export_paths_button.setFocusPolicy(Qt.NoFocus)
        self.export_paths_button.setToolTip("Экспорт полного пути выбранных файлов в selected_list.csv")
        self.export_paths_button.clicked.connect(lambda: self.export_list_for_selected(include_paths=True))
        bottom.addWidget(self.count_label)
        bottom.addStretch()
        bottom.addWidget(self.export_list_button)
        bottom.addWidget(self.export_paths_button)
        bottom.addStretch()
        bottom.addWidget(self.open_folder_button)
        bottom.addWidget(self.open_button)

        self.open_folder_hint = QLabel(
            "Для нескольких файлов будут открыты их директории (в каждой — со своим выделением)"
        )
        self.open_folder_hint.setStyleSheet("""
            QLabel {
                color: #d4d4d4;
                background: #3a3b3f;
                border: 1px solid #55565a;
                border-radius: 5px;
                padding: 4px 10px;
                font-size: 11px;
            }
        """)
        self.open_folder_hint.setVisible(False)
        hint_row = QHBoxLayout()
        hint_row.addStretch()
        hint_row.addWidget(self.open_folder_hint)

        root.addWidget(self.title)
        root.addLayout(path_row)
        root.addLayout(self.ext_container)
        root.addWidget(self.search)
        root.addWidget(self.list, 1)
        root.addLayout(hint_row)
        root.addLayout(bottom)
        self.list.itemSelectionChanged.connect(self.update_count)

    def show_for_folder(self, folder: str):
        self.stop_scanners()
        self.folder = folder
        self.path_label.setText(folder)
        self.selected_paths.clear()
        self.search.clear()
        self.registry.clear()
        self.available_extensions = []

        self.show()
        self.raise_()
        self.activateWindow()
        self.search.setFocus()
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(
            screen.center().x() - self.width() // 2,
            screen.center().y() - self.height() // 2,
        )

        self.start_scanners()

    def start_scanners(self):
        self.shallow_scanner = ShallowScannerThread(self.folder, is_recursive=self.is_recursive, max_depth=2)
        self.shallow_scanner.files_found.connect(self.on_files_discovered)
        self.shallow_scanner.start()

        if self.is_recursive:
            self.mid_scanner = MidScannerThread(self.folder, min_depth=3, max_depth=4)
            self.mid_scanner.files_found.connect(self.on_files_discovered)
            self.mid_scanner.start()

            self.deep_scanner = DeepScannerThread(self.folder, min_depth=5)
            self.deep_scanner.files_found.connect(self.on_files_discovered)
            self.deep_scanner.start()

    def on_files_discovered(self, new_files: list):
        new_ext = self.registry.add_files(new_files)
        if new_ext:
            self.available_extensions = sorted(list(self.registry.extensions))
            if self.global_preferred_extensions is not None:
                intersection = self.global_preferred_extensions.intersection(self.available_extensions)
                self.active_extensions = set(intersection) if intersection else set(self.available_extensions)
            else:
                self.active_extensions = set(self.available_extensions)
            self.rebuild_extension_buttons()

        self.schedule_ui_update()

    def schedule_ui_update(self):
        self.pending_ui_refresh = True
        if not self.ui_update_timer.isActive():
            self.ui_update_timer.start()

    def _on_ui_update_tick(self):
        if self.pending_ui_refresh:
            self.pending_ui_refresh = False
            self.request_filter()

    def toggle_recursive(self):
        self.stop_scanners()
        self.is_recursive = not self.is_recursive
        self.recursive_button.setChecked(self.is_recursive)
        active_style = """
            QPushButton {
                background: #1976d2;
                border: 1px solid #2196f3;
                border-radius: 7px;
                padding: 7px 14px;
                font-weight: bold;
            }
        """
        self.recursive_button.setStyleSheet(active_style if self.is_recursive else "")
        if self.folder:
            self.registry.clear()
            self.available_extensions = []
            self.start_scanners()

    def switch_tab_relative(self, delta: int):
        num_exts = len(self.available_extensions)
        if num_exts == 0:
            return
        available_set = set(self.available_extensions)
        if not self.active_extensions:
            current_state = -1
        elif self.active_extensions == available_set:
            current_state = 0
        elif len(self.active_extensions) == 1:
            ext = next(iter(self.active_extensions))
            if ext in self.available_extensions:
                current_state = self.available_extensions.index(ext) + 1
            else:
                current_state = 0
        else:
            current_state = 0
        total_states = num_exts + 2
        normalized_current = current_state + 1
        new_normalized = (normalized_current + delta) % total_states
        new_state = new_normalized - 1
        if new_state == -1:
            self.active_extensions.clear()
            self.global_preferred_extensions = set()
        elif new_state == 0:
            self.active_extensions = set(available_set)
            self.global_preferred_extensions = None
        else:
            target_ext = self.available_extensions[new_state - 1]
            self.active_extensions = {target_ext}
            self.global_preferred_extensions = set(self.active_extensions)
        self.update_extension_controls()
        self.request_filter()

    def rebuild_extension_buttons(self):
        # --- 1. Безопасная очистка старого содержимого (кнопки или дропдаун) ---
        while self.ext_container.count():
            item = self.ext_container.takeAt(0)
            if item.widget():
                item.widget().hide()
                item.widget().deleteLater()
            elif item.layout():
                layout = item.layout()
                while layout.count():
                    child = layout.takeAt(0)
                    if child.widget():
                        child.widget().hide()
                        child.widget().deleteLater()

        self.ext_actions = {}
        self.ext_all_action = None
        if hasattr(self, "ext_dropdown_button"):
            del self.ext_dropdown_button

        # --- 2. Выбор представления: сетка кнопок или выпадающий список ---
        # При большом числе расширений сетка кнопок (6 колонок) превращается
        # в нечитаемую простыню, поэтому после EXT_ROWS_THRESHOLD строк
        # переключаемся на компактный dropdown с чекбоксами.
        use_dropdown = len(self.available_extensions) > EXT_DROPDOWN_THRESHOLD
        self.ext_display_mode = "dropdown" if use_dropdown else "buttons"

        if use_dropdown:
            self._build_extension_dropdown()
        else:
            self._build_extension_grid()

        self.update_extension_controls()

    def _build_extension_grid(self):
        buttons_to_add = []

        all_button = QPushButton()
        all_button.setCheckable(True)
        all_button.setFocusPolicy(Qt.NoFocus)
        all_button.setProperty("is_all_button", True)
        all_button.clicked.connect(self.toggle_all_extensions)
        buttons_to_add.append(all_button)

        for ext in self.available_extensions:
            button = QPushButton()
            button.setCheckable(True)
            button.setFocusPolicy(Qt.NoFocus)
            button.setProperty("extension", ext)
            button.clicked.connect(lambda _, e=ext: self.toggle_extension(e))
            buttons_to_add.append(button)

        # --- Компоновка в сетку на 6 колонок ---
        grid_layout = QGridLayout()
        grid_layout.setSpacing(6)
        grid_layout.setContentsMargins(0, 0, 0, 0)

        for index, btn in enumerate(buttons_to_add):
            row = index // EXT_GRID_COLUMNS
            col = index % EXT_GRID_COLUMNS
            # Policy 'Expanding' разрешает кнопкам растягиваться по горизонтали
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            grid_layout.addWidget(btn, row, col)

        # Равномерно распределяем ширину между всеми колонками
        for col in range(EXT_GRID_COLUMNS):
            grid_layout.setColumnStretch(col, 1)

        self.ext_container.addLayout(grid_layout)

    def _build_extension_dropdown(self):
        self.ext_dropdown_button = QToolButton()
        self.ext_dropdown_button.setPopupMode(QToolButton.InstantPopup)
        self.ext_dropdown_button.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.ext_dropdown_button.setFocusPolicy(Qt.NoFocus)
        self.ext_dropdown_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.ext_dropdown_button.setStyleSheet("""
            QToolButton {
                background: #303134;
                border: 1px solid #4a4b4e;
                border-radius: 7px;
                padding: 7px 14px;
                text-align: left;
            }
            QToolButton:hover {
                background: #3a3b3f;
            }
            QToolButton::menu-indicator {
                subcontrol-position: right center;
                subcontrol-origin: padding;
                right: 8px;
            }
        """)

        menu = CheckableMenu(self.ext_dropdown_button)
        menu.setStyleSheet("""
            QMenu {
                background: #2b2c30;
                color: #eeeeee;
                border: 1px solid #4a4b4e;
            }
            QMenu::item {
                padding: 6px 24px 6px 20px;
            }
            QMenu::item:selected {
                background: #1976d2;
            }
        """)

        self.ext_all_action = QAction("All", menu)
        self.ext_all_action.setCheckable(True)
        self.ext_all_action.triggered.connect(lambda: self.toggle_all_extensions())
        menu.addAction(self.ext_all_action)
        menu.addSeparator()

        self.ext_actions: dict[str, QAction] = {}
        for ext in self.available_extensions:
            action = QAction(ext, menu)
            action.setCheckable(True)
            action.setProperty("extension", ext)
            action.triggered.connect(lambda _checked=False, e=ext: self.toggle_extension(e))
            menu.addAction(action)
            self.ext_actions[ext] = action

        self.ext_dropdown_button.setMenu(menu)

        row = QHBoxLayout()
        row.addWidget(self.ext_dropdown_button)
        self.ext_container.addLayout(row)

    def get_extension_buttons(self) -> list[QPushButton]:
        buttons = []
        for i in range(self.ext_container.count()):
            layout = self.ext_container.itemAt(i).layout()
            if not layout:
                continue
            for j in range(layout.count()):
                widget = layout.itemAt(j).widget()
                if isinstance(widget, QPushButton):
                    buttons.append(widget)
        return buttons

    def get_all_button(self) -> QPushButton | None:
        for btn in self.get_extension_buttons():
            if btn.property("is_all_button"):
                return btn
        return None

    def toggle_extension(self, extension: str):
        available_set = set(self.available_extensions)
        is_all_active = bool(available_set) and self.active_extensions == available_set
        if is_all_active:
            self.active_extensions = {extension}
        else:
            if extension in self.active_extensions:
                self.active_extensions.remove(extension)
            else:
                self.active_extensions.add(extension)
        self.global_preferred_extensions = set(self.active_extensions)
        self.update_extension_controls()
        self.request_filter()

    def toggle_all_extensions(self):
        available_set = set(self.available_extensions)
        if self.active_extensions == available_set and available_set:
            self.active_extensions.clear()
            self.global_preferred_extensions = set()
        else:
            self.active_extensions = set(available_set)
            self.global_preferred_extensions = None
        self.update_extension_controls()
        self.request_filter()

    def update_extension_controls(self):
        if getattr(self, "ext_display_mode", "buttons") == "dropdown":
            self.update_extension_dropdown()
        else:
            self.update_extension_buttons()

    def update_extension_dropdown(self):
        if not hasattr(self, "ext_dropdown_button"):
            return
        available_set = set(self.available_extensions)
        all_active = bool(available_set) and self.active_extensions == available_set
        total_files_count = len(self.registry.files)
        total_selected_count = len(self.selected_paths)

        if self.ext_all_action is not None:
            self.ext_all_action.blockSignals(True)
            self.ext_all_action.setChecked(all_active)
            self.ext_all_action.blockSignals(False)

        for ext, action in self.ext_actions.items():
            ext_found_count = self.registry.ext_counts.get(ext, 0)
            selected_count = sum(1 for path in self.selected_paths if Path(path).suffix.lower() == ext)
            is_active = ext in self.active_extensions
            action.blockSignals(True)
            action.setChecked(is_active)
            action.setText(f"{ext}  ({selected_count}/{ext_found_count})")
            action.blockSignals(False)

        active_count = len(self.active_extensions)
        total_ext_count = len(self.available_extensions)
        if all_active:
            summary = f"All types  ({total_selected_count}/{total_files_count})"
        elif active_count == 0:
            summary = "Types: none selected"
        elif active_count == 1:
            only_ext = next(iter(self.active_extensions))
            summary = f"{only_ext}  ({total_selected_count}/{total_files_count})"
        else:
            summary = f"Types: {active_count}/{total_ext_count}  ({total_selected_count}/{total_files_count})"
        self.ext_dropdown_button.setText(summary)

    def update_extension_buttons(self):
        all_button = self.get_all_button()
        all_active = bool(self.available_extensions) and self.active_extensions == set(self.available_extensions)
        active_style = """
            QPushButton {
                background: #1976d2;
                border: 1px solid #2196f3;
                border-radius: 7px;
                padding: 7px 14px;
            }
            QPushButton:hover {
                background: #2196f3;
            }
        """
        total_files_count = len(self.registry.files)
        total_selected_count = len(self.selected_paths)
        if all_button is not None:
            all_button.blockSignals(True)
            all_button.setChecked(all_active)
            all_button.setText(f"All ({total_selected_count}/{total_files_count})")
            all_button.setStyleSheet(active_style if all_active else "")
            all_button.blockSignals(False)
        for widget in self.get_extension_buttons():
            if widget.property("is_all_button"):
                continue
            extension = widget.property("extension")
            if not extension:
                continue
            ext_found_count = self.registry.ext_counts.get(extension, 0)
            selected_count = sum(1 for path in self.selected_paths if Path(path).suffix.lower() == extension)
            is_active = extension in self.active_extensions
            widget.blockSignals(True)
            widget.setChecked(is_active)
            widget.setText(f"{extension} ({selected_count}/{ext_found_count})")
            widget.setStyleSheet(active_style if is_active else "")
            widget.blockSignals(False)

    def on_search_text_changed(self, _text: str):
        # Не гоняем воркер на каждое нажатие клавиши — ждём короткую паузу.
        self.filter_debounce_timer.start()

    def request_filter(self):
        """Немедленно (без debounce) пересчитать список — для кликов по
        кнопкам расширений, переключения вкладок и т.п."""
        self.filter_debounce_timer.stop()
        self._run_filter_now()

    def _run_filter_now(self):
        text = self.search.text().lower().strip()
        tokens = [t for t in re.split(r'[\s,\.]+', text) if t]

        if self.filter_worker and self.filter_worker.isRunning():
            self.filter_worker.stop()
            self.filter_worker.wait(50)

        self.filter_generation += 1
        generation = self.filter_generation

        if not self.active_extensions:
            self.filter_worker = None
            self.all_matched_paths = []
            self._apply_filtered_results([], generation)
            return

        worker = FilterWorkerThread(
            list(self.registry.files),
            tokens,
            set(self.active_extensions),
            MAX_VISIBLE_ITEMS,
            generation,
        )
        worker.filtered_ready.connect(self._on_filtered_ready)
        self.filter_worker = worker
        worker.start()

    def _on_filtered_ready(self, results: list, all_matched_paths: list, generation: int):
        # Отбрасываем устаревший результат, если пользователь уже успел
        # ввести что-то ещё / переключить фильтр, пока воркер считал.
        if generation != self.filter_generation:
            return
        self.all_matched_paths = all_matched_paths
        self._apply_filtered_results(results, generation)

    def _apply_filtered_results(self, results: list, generation: int):
        self.list.blockSignals(True)
        try:
            self.list.clear()
            for full_path, file_name, dir_part in results:
                item = QListWidgetItem()
                item.setData(Qt.UserRole, full_path)
                item.setData(Qt.DisplayRole, file_name)
                item.setData(Qt.UserRole + 1, dir_part)
                item.setIcon(get_file_icon(full_path))

                self.list.addItem(item)
                if full_path in self.selected_paths:
                    item.setSelected(True)
        finally:
            self.list.blockSignals(False)
        self.update_count()

    def _sync_visible_selection(self):
        for i in range(self.list.count()):
            item = self.list.item(i)
            path = item.data(Qt.UserRole)
            if not path:
                continue
            if item.isSelected():
                self.selected_paths.add(path)
            else:
                self.selected_paths.discard(path)

    def select_all(self):
        # Ctrl+A должен выделять ВСЕ файлы, подходящие под текущий фильтр
        # (расширения + поиск), а не только видимый (обрезанный до
        # MAX_VISIBLE_ITEMS) кусок списка. self.list.selectAll() подсвечивает
        # видимые строки для наглядности, а реальный набор для Open берётся
        # из self.selected_paths, куда добавляем весь all_matched_paths.
        if self.all_matched_paths:
            self.selected_paths.update(self.all_matched_paths)
        self.list.selectAll()
        self.update_count()

    def update_count(self):
        self._sync_visible_selection()
        count = len(self.selected_paths)
        self.count_label.setText(f"{count} selected")
        self.open_button.setEnabled(count > 0 or self.list.currentRow() >= 0)
        self.open_folder_hint.setVisible(count > 1)
        self.update_extension_controls()

    def open_selected(self):
        if self.search.hasFocus():
            if self.list.count() > 0:
                self.list.setFocus()
                if self.list.currentRow() < 0:
                    self.list.setCurrentRow(0)
            return
        self._sync_visible_selection()
        paths = list(self.selected_paths)
        if not paths and self.list.currentRow() >= 0:
            current_item = self.list.currentItem()
            if current_item:
                path = current_item.data(Qt.UserRole)
                if path:
                    paths = [path]
        if paths:
            open_selected_files(paths)
        self.stop_scanners()
        self.selected_paths.clear()
        self.search.clear()
        self.hide()

    def open_folder_for_selected(self):
        if self.search.hasFocus():
            return
        self._sync_visible_selection()
        paths = list(self.selected_paths)
        if not paths and self.list.currentRow() >= 0:
            current_item = self.list.currentItem()
            if current_item:
                path = current_item.data(Qt.UserRole)
                if path:
                    paths = [path]
        if not paths:
            return

        # Группируем по директории — на каждую уникальную папку открываем
        # одно окно Explorer с выделением всех файлов из неё разом.
        groups: dict[str, list[str]] = {}
        for path in paths:
            groups.setdefault(os.path.dirname(path), []).append(path)

        for directory, group_paths in groups.items():
            try:
                reveal_files_in_explorer(directory, group_paths)
            except Exception as e:
                print(f"Не удалось открыть {directory} в Проводнике: {e}")

        self.stop_scanners()
        self.selected_paths.clear()
        self.search.clear()
        self.hide()

    def export_list_for_selected(self, include_paths: bool):
        self._sync_visible_selection()
        paths = list(self.selected_paths)
        if not paths and self.list.currentRow() >= 0:
            current_item = self.list.currentItem()
            if current_item:
                path = current_item.data(Qt.UserRole)
                if path:
                    paths = [path]
        if not paths or not self.folder:
            return
        # Экспорт — вспомогательное действие, а не завершение выбора файла,
        # поэтому (в отличие от Open/Open Folder) не закрываем окно и не
        # сбрасываем текущее выделение.
        output_path = export_selected_to_csv(paths, self.folder, include_paths)
        if output_path:
            print(f"Экспортировано {len(paths)} файлов в {output_path}")


class App:
    def __init__(self):
        self.qt_app = QApplication(sys.argv)
        self.qt_app.setApplicationName("File Context Selector")
        self.qt_app.setQuitOnLastWindowClosed(False)

        app_icon = self.qt_app.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
        self.qt_app.setWindowIcon(app_icon)

        if not getattr(sys, "frozen", False):
            try:
                app_icon.pixmap(256, 256).save("app_icon.ico", "ICO")
            except Exception:
                pass
        QLocalServer.removeServer(APP_ID)
        self.server = QLocalServer()
        if not self.server.listen(APP_ID):
            socket = QLocalSocket()
            socket.connectToServer(APP_ID)
            if socket.waitForConnected(500):
                self._send_to_existing_instance()
                sys.exit(0)
            else:
                QLocalServer.removeServer(APP_ID)
                self.server.listen(APP_ID)
        self.server.newConnection.connect(self._consume_activation)
        self.window = SelectorWindow()
        self.window.setWindowIcon(app_icon)
        self.hotkey_filter = NativeHotkeyFilter(self.toggle)
        self.qt_app.installNativeEventFilter(self.hotkey_filter)
        user32 = ctypes.windll.user32
        self.hotkey_registered = bool(
            user32.RegisterHotKey(None, HOTKEY_ID, MOD_CONTROL | MOD_ALT, VK_SPACE)
        )
        if not self.hotkey_registered:
            print("Could not register Ctrl+Alt+Space hotkey", file=sys.stderr)
        self.tray_icon = QSystemTrayIcon(self.qt_app)
        self.tray_icon.setIcon(app_icon)
        self.tray_icon.setToolTip("File Context Selector\nCtrl+Alt+Space")
        self.build_tray_menu()
        self.tray_icon.show()

    def build_tray_menu(self):
        self.tray_menu = QMenu()
        self.autostart_action = QAction("Run at Windows startup", self.tray_menu)
        self.autostart_action.setCheckable(True)
        self.autostart_action.setChecked(is_autostart_enabled())
        self.autostart_action.toggled.connect(set_autostart)
        self.tray_menu.addAction(self.autostart_action)
        self.tray_menu.addSeparator()
        self.exit_action = self.tray_menu.addAction("Exit")
        self.exit_action.triggered.connect(self.exit_app)
        self.tray_icon.setContextMenu(self.tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)

    def on_tray_activated(self, reason: QSystemTrayIcon.ActivationReason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.Context):
            self.tray_menu.exec(QCursor.pos())

    def _send_to_existing_instance(self):
        socket = QLocalSocket()
        socket.connectToServer(APP_ID)
        if socket.waitForConnected(300):
            socket.write(b"activate")
            socket.flush()
            socket.waitForBytesWritten(300)
            socket.disconnectFromServer()

    def _consume_activation(self):
        connection = self.server.nextPendingConnection()
        if connection:
            connection.readyRead.connect(connection.deleteLater)
            connection.disconnectFromServer()
        self.toggle()

    def show_selector(self):
        folder = get_active_explorer_path()
        if folder:
            self.window.show_for_folder(folder)

    def toggle(self):
        if self.window.isVisible():
            self.window.hide()
            return
        self.show_selector()

    def exit_app(self):
        user32 = ctypes.windll.user32
        if self.hotkey_registered:
            user32.UnregisterHotKey(None, HOTKEY_ID)
            self.hotkey_registered = False
        if self.server:
            self.server.close()
            QLocalServer.removeServer(APP_ID)
        if self.tray_icon:
            self.tray_icon.hide()
        if self.window:
            self.window.close()
        self.qt_app.quit()

    def run(self) -> int:
        return self.qt_app.exec()


if __name__ == "__main__":
    app = App()
    sys.exit(app.run())