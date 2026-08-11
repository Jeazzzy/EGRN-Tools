"""Ключи сортировки для имён и путей с числовыми индексами."""

from __future__ import annotations

from pathlib import Path
import re
import unicodedata


def natural_text_key(value: str) -> tuple[tuple[int, object], ...]:
    """Сортирует ``И2`` раньше ``И10`` без зависимости от регистра."""

    normalised = unicodedata.normalize("NFKC", str(value)).casefold()
    return tuple(
        (0, int(part)) if part.isdigit() else (1, part)
        for part in re.split(r"(\d+)", normalised)
        if part
    )


def natural_path_key(path: str | Path):
    """Возвращает natural-sort ключ отдельно для каждой части пути."""

    return tuple(natural_text_key(part) for part in Path(path).parts)
