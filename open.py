"""Открытие файлов: разрешение ассоциации по умолчанию, fallback на
Windows Photo Viewer для изображений, и раскрытие файлов в Проводнике
(Open Folder) через SHOpenFolderAndSelectItems с fallback на explorer /select,."""
import os
import re
import subprocess
import time
import traceback
import winreg
from pathlib import Path


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


