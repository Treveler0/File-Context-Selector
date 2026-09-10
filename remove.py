"""Удаление файлов: в Корзину (через SHFileOperationW) или безвозвратно."""
import ctypes
import os
from ctypes import wintypes


class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("wFunc", wintypes.UINT),
        ("pFrom", wintypes.LPCWSTR),
        ("pTo", wintypes.LPCWSTR),
        ("fFlags", ctypes.c_ushort),
        ("fAnyOperationsAborted", wintypes.BOOL),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", wintypes.LPCWSTR),
    ]


FO_DELETE = 0x0003
FOF_ALLOWUNDO = 0x0040       # можно восстановить из Корзины
FOF_NOCONFIRMATION = 0x0010  # без диалога "вы уверены?"
FOF_SILENT = 0x0004          # без прогресс-бара


def delete_files_to_recycle_bin(paths: list[str]) -> bool:
    """Отправляет файлы в Корзину через SHFileOperationW — тот же API,
    которым пользуется сам Explorer при Del, поэтому файлы восстановимы
    штатными средствами Windows."""
    if not paths:
        return True
    # pFrom — пути, разделённые '\0', с ДВОЙНЫМ '\0' в конце (так требует API).
    p_from = "\0".join(paths) + "\0\0"
    op = SHFILEOPSTRUCTW()
    op.hwnd = 0
    op.wFunc = FO_DELETE
    op.pFrom = p_from
    op.pTo = None
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
    result = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return result == 0 and not op.fAnyOperationsAborted


def delete_files_permanently(paths: list[str]) -> list[str]:
    """Удаляет файлы безвозвратно (мимо Корзины). Не бросает исключение на
    первом же неудачном файле — копит неудачные пути и возвращает их."""
    failed = []
    for path in paths:
        try:
            os.remove(path)
        except OSError as e:
            print(f"Не удалось удалить {path}: {e}")
            failed.append(path)
    return failed

