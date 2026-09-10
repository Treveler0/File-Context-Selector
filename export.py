"""Экспорт выбранных файлов в CSV (Export List)."""
import csv
import os


_SIZE_UNITS = ["b", "K", "M", "G", "T"]


def format_size(num_bytes: int) -> str:
    """Читаемый размер в английской локализации (b/K/M/G/T), кратность 1024."""
    size = float(num_bytes)
    for unit in _SIZE_UNITS:
        if unit == "b":
            if size < 1024:
                return f"{int(size)} b"
        elif size < 1024 or unit == _SIZE_UNITS[-1]:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{int(num_bytes)} b"


def export_selected_to_csv(paths: list[str], folder: str) -> str | None:
    """Экспортирует paths в <folder>/selected_list.csv — обычная читаемая
    таблица (без маскировки заголовков): Path, Filename (W ext),
    Filename (W/O ext), Byte-Size, Size. Разделитель — ';' (дефолтный
    список-разделитель Excel на ru-RU локали), квотирование — штатное через
    модуль csv (RFC 4180): поле, где встретится ';' или кавычка, автоматически
    оборачивается в кавычки — запятая/';' в имени файла ничего не сломает."""
    if not paths or not folder:
        return None

    rows = []
    for path in paths:
        filename_with_ext = os.path.basename(path)
        filename_without_ext, _ext = os.path.splitext(filename_with_ext)
        try:
            byte_size = os.path.getsize(path)
        except OSError:
            byte_size = 0
        rows.append([
            path,
            filename_with_ext,
            filename_without_ext,
            byte_size,
            format_size(byte_size),
        ])

    output_path = os.path.join(folder, "selected_list.csv")
    try:
        with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";", quoting=csv.QUOTE_MINIMAL)
            writer.writerow(["Path", "Filename (W ext)", "Filename (W/O ext)", "Byte-Size", "Size"])
            writer.writerows(rows)
        return output_path
    except OSError as e:
        print(f"Ошибка экспорта CSV: {e}")
        return None

