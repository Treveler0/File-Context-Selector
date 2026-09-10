"""Точка входа. Вся логика — в app.py и соседних модулях (list.py, open.py,
remove.py, export.py)."""
import sys

from app import App

if __name__ == "__main__":
    app = App()
    sys.exit(app.run())
