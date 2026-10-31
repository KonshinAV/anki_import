#!/usr/bin/env python3
"""
Импорт заметок из Excel в Anki (через AnkiConnect) + озвучка (gTTS / edge-tts).
Все настройки лежат в config.yaml рядом со скриптом.
Язык сообщений в консоли (ru / de) задаётся в config.yaml (ui.language) или флагом --lang.

Excel-Import nach Anki (über AnkiConnect) + Vertonung (gTTS / edge-tts).
Die Sprache der Konsolenausgabe (ru / de) wird in config.yaml (ui.language) oder mit --lang gewählt.

Примеры запуска / Beispiele:
    python3 anki_import.py                      # импорт + озвучка
    python3 anki_import.py --lang de            # сообщения на немецком
    python3 anki_import.py --dry-run            # ничего не менять, только показать план
    python3 anki_import.py --engine edge        # озвучка через edge-tts
    python3 anki_import.py --no-audio           # только импорт
    python3 anki_import.py --audio-only         # только озвучка пустых аудио-полей
    python3 anki_import.py --config my.yaml     # другой файл настроек
    python3 anki_import.py -h                   # все параметры

При запуске скрипт проверяет, что колода и тип заметки из config.yaml существуют в Anki,
и создаёт недостающее (тип заметки - по шаблонам из секции note_type).
С --dry-run все проверки выполняются, но ничего не создаётся и не меняется.

Если Excel-файл из конфига не существует, скрипт создаст его (рядом со скриптом)
с шапкой из полей и завершится: заполните файл и запустите снова.
"""
from __future__ import annotations

import argparse
import asyncio
import base64
import hashlib
import html
import io
import logging
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from contextlib import nullcontext
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import requests

try:
    import yaml
except ImportError:
    sys.exit("Нужен пакет PyYAML / PyYAML-Paket erforderlich: pip install pyyaml")

try:                                    # полоса загрузки (необязательная зависимость)
    from tqdm import tqdm
    from tqdm.contrib.logging import logging_redirect_tqdm
except ImportError:
    tqdm = None
    logging_redirect_tqdm = None

log = logging.getLogger("anki_import")

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = SCRIPT_DIR / "config.yaml"


# --------------------------------------------------------------------------- #
# Локализация сообщений (ru / de)
# --------------------------------------------------------------------------- #
SUPPORTED_LANGS = ("ru", "de")
DEFAULT_LANG = "ru"
_LANG = DEFAULT_LANG

# ключ: (русский текст, немецкий текст); подстановки {name} - через str.format
MESSAGES: dict[str, tuple[str, str]] = {
    # --- конфигурация ---
    "cfg_not_found": ("Файл конфигурации не найден: {path}",
                      "Konfigurationsdatei nicht gefunden: {path}"),
    "cfg_empty": ("Файл конфигурации пуст или имеет неверный формат: {path}",
                  "Konfigurationsdatei ist leer oder hat ein ungültiges Format: {path}"),
    "cfg_lang": ("ui.language должен быть 'ru' или 'de' (сейчас: {value!r})",
                 "ui.language muss 'ru' oder 'de' sein (aktuell: {value!r})"),
    "cfg_missing": ("В {name} не хватает параметров: {items}",
                    "In {name} fehlen Parameter: {items}"),
    "cfg_engine": ("audio.engine должен быть 'gtts' или 'edge'",
                   "audio.engine muss 'gtts' oder 'edge' sein"),
    "cfg_cards": ("note_type.cards должен быть непустым списком карточек",
                  "note_type.cards muss eine nicht leere Liste von Karten sein"),
    "cfg_card_fields": ("note_type.cards[{idx}]: нужны непустые name, front, back",
                        "note_type.cards[{idx}]: name, front und back dürfen nicht leer sein"),
    # --- повторы / операции ---
    "retry_fail": ("{what}: попытка {n}/{attempts} не удалась ({err}), повтор...",
                   "{what}: Versuch {n}/{attempts} fehlgeschlagen ({err}), neuer Versuch..."),
    "op_default": ("операция", "Vorgang"),
    "op_add": ("добавление '{key}'", "Hinzufügen von '{key}'"),
    "op_upload": ("загрузка {name}", "Hochladen von {name}"),
    "op_update_note": ("обновление заметки {id}", "Aktualisieren der Notiz {id}"),
    # --- полосы загрузки ---
    "unit_default": ("шт", "Stk."),
    "bar_tts": ("Озвучка", "Vertonung"),
    "unit_text": ("текст", "Text"),
    "bar_upload": ("Загрузка аудио в Anki", "Audio-Upload nach Anki"),
    "unit_file": ("файл", "Datei"),
    "bar_update": ("Обновление заметок", "Notizen werden aktualisiert"),
    "unit_note": ("заметка", "Notiz"),
    # --- Excel ---
    "sheet_not_found": ("Лист '{sheet}' не найден. Доступные листы: {sheets}",
                        "Tabellenblatt '{sheet}' nicht gefunden. Verfügbare Blätter: {sheets}"),
    "excel_created": ("Excel-файл не найден, создан новый: {path} (лист '{sheet}', колонок: {cols})",
                      "Excel-Datei nicht gefunden, neue Datei erstellt: {path} (Blatt '{sheet}', Spalten: {cols})"),
    "excel_fill": ("Заполните его данными и запустите скрипт снова.",
                   "Bitte mit Daten füllen und das Skript erneut starten."),
    "excel_missing_dry": ("[dry-run] Excel-файл {path} не найден, будет создан с шапкой "
                          "(импорт в этот запуск не попадёт).",
                          "[dry-run] Excel-Datei {path} nicht gefunden, sie wird mit Kopfzeile erstellt "
                          "(ein Import findet in diesem Lauf nicht statt)."),
    "rows_read": ("Прочитано строк из Excel: {n} ({name})",
                  "Zeilen aus Excel gelesen: {n} ({name})"),
    # --- AnkiConnect ---
    "anki_conn": ("Не удалось подключиться к AnkiConnect ({url}). "
                  "Anki запущен и установлен дополнительный модуль AnkiConnect?",
                  "Verbindung zu AnkiConnect ({url}) fehlgeschlagen. "
                  "Läuft Anki und ist das AnkiConnect-Add-on installiert?"),
    "anki_error": ("AnkiConnect ({action}): {error}", "AnkiConnect ({action}): {error}"),
    # --- проверка имён ---
    "deck_empty": ("anki.deck: имя колоды не задано", "anki.deck: Stapelname nicht angegeben"),
    "deck_ctrl": ("anki.deck: в имени колоды недопустимые управляющие символы: {name!r}",
                  "anki.deck: Stapelname enthält ungültige Steuerzeichen: {name!r}"),
    "deck_part": ("anki.deck: пустая часть в имени колоды (лишние '::' или пробелы): {name!r}",
                  "anki.deck: leerer Teil im Stapelnamen (überflüssige '::' oder Leerzeichen): {name!r}"),
    "model_empty": ("anki.model: имя типа заметки не задано",
                    "anki.model: Name des Notiztyps nicht angegeben"),
    "model_ctrl": ("anki.model: в имени недопустимые управляющие символы: {name!r}",
                   "anki.model: Name enthält ungültige Steuerzeichen: {name!r}"),
    # --- колода ---
    "deck_case": ("Колода уже есть с другим регистром ('{found}'), используем её.",
                  "Stapel existiert bereits mit anderer Groß-/Kleinschreibung ('{found}'), wird verwendet."),
    "deck_exists": ("Колода '{name}' существует.", "Stapel '{name}' existiert."),
    "deck_dry": ("[dry-run] колода '{name}' не найдена, будет создана.",
                 "[dry-run] Stapel '{name}' nicht gefunden, wird erstellt."),
    "deck_created": ("Колода '{name}' создана.", "Stapel '{name}' erstellt."),
    # --- тип заметки ---
    "model_exists": ("Тип заметки '{name}' существует.", "Notiztyp '{name}' existiert."),
    "model_missing_fields": ("В существующем типе '{name}' нет полей из конфига (они не будут заполнены): "
                             "{fields}. Тип не изменяется.",
                             "Im vorhandenen Notiztyp '{name}' fehlen Felder aus der Konfiguration "
                             "(sie werden nicht befüllt): {fields}. Der Typ wird nicht geändert."),
    "model_unknown_fields": ("Шаблоны note_type ссылаются на поля, которых нет в типе: {unknown}. "
                             "Поля типа: {fields}",
                             "Die note_type-Vorlagen verweisen auf Felder, die im Typ fehlen: {unknown}. "
                             "Felder des Typs: {fields}"),
    "model_dry": ("[dry-run] тип заметки '{name}' не найден, будет создан: полей {fields}, "
                  "карточек {cards} ({names}).",
                  "[dry-run] Notiztyp '{name}' nicht gefunden, wird erstellt: Felder {fields}, "
                  "Karten {cards} ({names})."),
    "model_dry_fields": ("[dry-run] поля: {fields}", "[dry-run] Felder: {fields}"),
    "model_created": ("Тип заметки '{name}' создан: полей {fields}, карточек {cards}.",
                      "Notiztyp '{name}' erstellt: Felder {fields}, Karten {cards}."),
    "media_missing": ("Файл '{name}' из шаблона карточки не найден в медиа Anki "
                      "(положите его в collection.media, иначе звук не проиграется).",
                      "Datei '{name}' aus der Kartenvorlage fehlt in den Anki-Medien "
                      "(in collection.media ablegen, sonst wird kein Ton abgespielt)."),
    "media_check_fail": ("Не удалось проверить медиафайл {name}: {err}",
                         "Mediendatei {name} konnte nicht geprüft werden: {err}"),
    # --- импорт ---
    "no_key_in_model": ("В модели '{model}' нет поля '{key}'. Поля: {fields}",
                        "Im Notiztyp '{model}' gibt es kein Feld '{key}'. Felder: {fields}"),
    "fields_not_in_model": ("Этих полей нет в модели Anki, они не будут заполнены: {fields}",
                            "Diese Felder fehlen im Anki-Notiztyp und werden nicht befüllt: {fields}"),
    "no_key_in_excel": ("В Excel нет колонки '{key}'. Колонки: {cols}",
                        "In Excel gibt es keine Spalte '{key}'. Spalten: {cols}"),
    "cols_missing": ("Этих колонок нет в Excel, поля будут пустыми: {cols}",
                     "Diese Spalten fehlen in Excel, die Felder bleiben leer: {cols}"),
    "row_no_key": ("Строка без значения в '{key}' пропущена: {row}",
                   "Zeile ohne Wert in '{key}' übersprungen: {row}"),
    "already_exists": ("Уже существует: {key}", "Existiert bereits: {key}"),
    "dry_will_add": ("[dry-run] будет добавлено: {key}", "[dry-run] wird hinzugefügt: {key}"),
    "add_failed": ("Не удалось добавить '{key}': {err}",
                   "'{key}' konnte nicht hinzugefügt werden: {err}"),
    "note_added": ("Добавлена заметка {id}: {key}", "Notiz hinzugefügt {id}: {key}"),
    "import_summary": ("Импорт{dry}: добавлено {added}, пропущено {skipped}, ошибок {failed}",
                       "Import{dry}: hinzugefügt {added}, übersprungen {skipped}, Fehler {failed}"),
    "plan_audio": ("[dry-run] для новых заметок будет озвучено: полей {fields}, уникальных текстов {texts}",
                   "[dry-run] Für neue Notizen werden vertont: Felder {fields}, eindeutige Texte {texts}"),
    # --- TTS ---
    "gtts_failed": ("gTTS не смог озвучить '{text}': {err}",
                    "gTTS konnte '{text}' nicht vertonen: {err}"),
    "edge_empty": ("пустой ответ", "leere Antwort"),
    "edge_failed": ("edge-tts не смог озвучить '{text}': {err}",
                    "edge-tts konnte '{text}' nicht vertonen: {err}"),
    "tts_generating": ("Генерация озвучки: {n} уникальных текстов ({engine})",
                       "Vertonung wird erzeugt: {n} eindeutige Texte ({engine})"),
    "audio_nothing": ("Озвучка: нечего делать (нет объектов для озвучки в очереди).",
                      "Vertonung: nichts zu tun (keine Elemente in der Warteschlange)."),
    "audio_plan": ("Озвучка: полей {fields}, уникальных текстов {texts}",
                   "Vertonung: Felder {fields}, eindeutige Texte {texts}"),
    "audio_dry_job": ("[dry-run] заметка {id} -> {field}: {text}",
                      "[dry-run] Notiz {id} -> {field}: {text}"),
    "update_failed": ("Не удалось обновить заметку {id}: {err}",
                      "Notiz {id} konnte nicht aktualisiert werden: {err}"),
    "audio_done": ("Озвучка: обновлено заметок {notes}, файлов загружено {files}",
                   "Vertonung: Notizen aktualisiert {notes}, Dateien hochgeladen {files}"),
    "cache_remove_failed": ("Не удалось удалить {path}: {err}",
                            "{path} konnte nicht gelöscht werden: {err}"),
    "cache_removed": ("Удалено временных .mp3: {n}", "Temporäre .mp3-Dateien gelöscht: {n}"),
    # --- main ---
    "dry_mode": ("[dry-run] Режим проверки: в Anki и в файлах ничего меняться не будет.",
                 "[dry-run] Prüfmodus: In Anki und in den Dateien wird nichts geändert."),
    "tqdm_tip": ("Совет: pip install tqdm — появится полоса загрузки при озвучке.",
                 "Tipp: pip install tqdm — zeigt beim Vertonen einen Fortschrittsbalken an."),
    "dry_done": ("[dry-run] Завершено: никаких изменений внесено не было.",
                 "[dry-run] Fertig: Es wurden keine Änderungen vorgenommen."),
}


def set_lang(code: str) -> None:
    """Устанавливает язык сообщений ('ru' или 'de')."""
    global _LANG
    if code not in SUPPORTED_LANGS:
        raise ValueError(t("cfg_lang", value=code))
    _LANG = code


def t(msg_id: str, /, **kw) -> str:
    """Возвращает сообщение на текущем языке с подстановкой параметров.
    msg_id - только позиционный параметр, чтобы он не конфликтовал с подстановкой {key} и т.п."""
    msg = MESSAGES[msg_id][SUPPORTED_LANGS.index(_LANG)]
    return msg.format(**kw) if kw else msg


# --------------------------------------------------------------------------- #
# Конфигурация
# --------------------------------------------------------------------------- #
REQUIRED_KEYS = {
    "anki": ["url", "deck", "model", "timeout"],
    "excel": ["file", "sheet"],
    "fields": ["key", "base", "sentences_count", "sentence_fields"],
    "audio": ["engine", "gtts_lang", "edge_voice", "workers", "cache_dir", "keep_files",
              "main_source", "main_dest", "sentence_source", "sentence_dest"],
    "note_type": ["css", "cards"],
    "retry": ["attempts", "delay"],
}

# параметры повторов; заполняются из config.yaml при загрузке
RETRY = {"attempts": 3, "delay": 2.0}


class Config:
    """Обёртка над данными из YAML + производные значения."""

    def __init__(self, raw: dict, base_dir: Path):
        self.raw = raw
        self.base_dir = base_dir

        self.anki = raw["anki"]
        self.excel = raw["excel"]
        self.fields = raw["fields"]
        self.audio = raw["audio"]
        self.note_type = raw["note_type"]

        self.xlsx_path = self.resolve(self.excel["file"])
        self.sheet = self.excel["sheet"]
        self.cache_dir = self.resolve(self.audio["cache_dir"])
        self.key_field = self.fields["key"]

        n = int(self.fields["sentences_count"])
        self.text_fields: list[str] = list(self.fields["base"]) + [
            tpl.format(n=i) for i in range(1, n + 1) for tpl in self.fields["sentence_fields"]
        ]
        self.audio_map: dict[str, str] = {self.audio["main_source"]: self.audio["main_dest"]}
        for i in range(1, n + 1):
            self.audio_map[self.audio["sentence_source"].format(n=i)] = self.audio["sentence_dest"].format(n=i)

        # все поля типа заметки в Anki: текстовые + аудио (в логичном порядке)
        model_fields: list[str] = list(self.fields["base"]) + [self.audio["main_dest"]]
        for i in range(1, n + 1):
            model_fields += [tpl.format(n=i) for tpl in self.fields["sentence_fields"]]
            model_fields.append(self.audio["sentence_dest"].format(n=i))
        self.model_fields: list[str] = list(dict.fromkeys(model_fields))   # без дублей, порядок сохранён

    def resolve(self, value: str) -> Path:
        """Относительные пути считаем от папки скрипта."""
        p = Path(str(value)).expanduser()
        return p if p.is_absolute() else self.base_dir / p


def load_config(path: Path, overrides: dict[tuple[str, str], object] | None = None) -> Config:
    if not path.is_file():
        raise FileNotFoundError(t("cfg_not_found", path=path))
    with open(path, encoding="utf-8") as f:
        raw = yaml.safe_load(f)
    if not isinstance(raw, dict):
        raise ValueError(t("cfg_empty", path=path))

    # значения из командной строки перекрывают yaml
    for (section, key), value in (overrides or {}).items():
        if value is not None:
            raw.setdefault(section, {})[key] = value

    # язык интерфейса (необязательный параметр, по умолчанию ru) - задаём как можно раньше,
    # чтобы дальнейшие сообщения об ошибках уже были на выбранном языке
    ui = raw.get("ui")
    lang = ui.get("language", DEFAULT_LANG) if isinstance(ui, dict) else DEFAULT_LANG
    lang = str(lang).strip().lower()
    if lang not in SUPPORTED_LANGS:
        raise ValueError(t("cfg_lang", value=lang))
    set_lang(lang)

    missing = [f"{sec}.{key}" for sec, keys in REQUIRED_KEYS.items() for key in keys
               if not isinstance(raw.get(sec), dict) or key not in raw[sec]]
    if missing:
        raise ValueError(t("cfg_missing", name=path.name, items=", ".join(missing)))
    if raw["audio"]["engine"] not in ("gtts", "edge"):
        raise ValueError(t("cfg_engine"))
    cards = raw["note_type"]["cards"]
    if not isinstance(cards, list) or not cards:
        raise ValueError(t("cfg_cards"))
    for idx, card in enumerate(cards, 1):
        if not isinstance(card, dict) or not all(str(card.get(k, "")).strip() for k in ("name", "front", "back")):
            raise ValueError(t("cfg_card_fields", idx=idx))

    RETRY["attempts"] = int(raw["retry"]["attempts"])
    RETRY["delay"] = float(raw["retry"]["delay"])
    return Config(raw, path.resolve().parent)


# --------------------------------------------------------------------------- #
# Вспомогательные функции
# --------------------------------------------------------------------------- #
def retry(func, what: str | None = None):
    """Вызывает func(); при ошибке повторяет с растущей паузой (параметры из конфига)."""
    what = what or t("op_default")
    attempts = RETRY["attempts"]
    for n in range(1, attempts + 1):
        try:
            return func()
        except Exception as e:  # noqa: BLE001
            if n == attempts:
                raise
            log.warning(t("retry_fail", what=what, n=n, attempts=attempts, err=e))
            time.sleep(RETRY["delay"] * n)


def clean_text(value: str) -> str:
    """Убирает HTML, [sound:...] и лишние пробелы (для озвучки и сравнения)."""
    value = re.sub(r"\[sound:[^\]]*\]", " ", value or "")
    value = re.sub(r"<[^>]+>", " ", value)
    value = html.unescape(value)
    return re.sub(r"\s+", " ", value).strip()


class _NoBar:
    """Заглушка, если tqdm не установлен: ничего не рисует."""
    def update(self, n: int = 1) -> None: pass
    def close(self) -> None: pass
    def __enter__(self): return self
    def __exit__(self, *exc): return False


def progress(total: int, desc: str, unit: str | None = None):
    """Полоса загрузки в stderr (потокобезопасная). Без tqdm — пустая заглушка."""
    if tqdm is None or total <= 0:
        return _NoBar()
    return tqdm(total=total, desc=desc, unit=unit or t("unit_default"),
                file=sys.stderr, dynamic_ncols=True)


def log_with_progress():
    """Лог-сообщения не ломают полосу загрузки."""
    return logging_redirect_tqdm() if tqdm is not None else nullcontext()


def cell_to_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
def create_excel_template(cfg: Config) -> None:
    """Создаёт Excel-файл с шапкой (поля из конфига) на листе cfg.sheet."""
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    cfg.xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = cfg.sheet
    ws.append(cfg.text_fields)
    for idx, name in enumerate(cfg.text_fields, 1):
        ws.cell(row=1, column=idx).font = Font(bold=True)
        ws.column_dimensions[get_column_letter(idx)].width = max(14, min(len(name) + 4, 30))
    ws.freeze_panes = "A2"
    wb.save(cfg.xlsx_path)
    wb.close()


def read_xlsx_file(filename: Path, sheet_name: str) -> list[dict[str, str]]:
    """Читает лист в список словарей {колонка: строка}. Пустые строки пропускает."""
    workbook = openpyxl.load_workbook(filename, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(t("sheet_not_found", sheet=sheet_name, sheets=workbook.sheetnames))
        rows = workbook[sheet_name].iter_rows(values_only=True)
        try:
            header = [cell_to_str(c) for c in next(rows)]
        except StopIteration:
            return []

        data = []
        for row in rows:
            if not any(c is not None and cell_to_str(c) != "" for c in row):
                continue
            data.append({k: cell_to_str(v) for k, v in zip(header, row) if k})
        return data
    finally:
        workbook.close()


# --------------------------------------------------------------------------- #
# AnkiConnect
# --------------------------------------------------------------------------- #
class AnkiClient:
    def __init__(self, url: str, deck_name: str, model_name: str, timeout: int = 60):
        self.url = url
        self.deck_name = deck_name
        self.model_name = model_name
        self.timeout = timeout
        self.session = requests.Session()

    def invoke(self, action: str, **params):
        payload = {"action": action, "version": 6, "params": params}
        try:
            resp = self.session.post(self.url, json=payload, timeout=self.timeout)
            resp.raise_for_status()
            data = resp.json()
        except requests.ConnectionError as e:
            raise RuntimeError(t("anki_conn", url=self.url)) from e
        if data.get("error") is not None:
            raise RuntimeError(t("anki_error", action=action, error=data["error"]))
        return data["result"]

    # ---- чтение ----
    def deck_notes(self) -> list[dict]:
        """Актуальная информация о всех заметках колоды."""
        deck = self.deck_name.replace('"', '\\"')
        ids = self.invoke("findNotes", query=f'deck:"{deck}"')
        notes: list[dict] = []
        for i in range(0, len(ids), 200):
            notes.extend(self.invoke("notesInfo", notes=ids[i:i + 200]))
        return notes

    def model_fields(self) -> list[str]:
        return self.invoke("modelFieldNames", modelName=self.model_name)

    def deck_names(self) -> list[str]:
        return self.invoke("deckNames")

    def model_names(self) -> list[str]:
        return self.invoke("modelNames")

    def media_exists(self, filename: str) -> bool:
        return bool(self.invoke("getMediaFilesNames", pattern=filename))

    # ---- запись ----
    def add_note(self, fields: dict[str, str], tags: list[str] | None = None) -> int:
        note = {
            "deckName": self.deck_name,
            "modelName": self.model_name,
            "fields": fields,
            # дубликаты мы проверяем сами (по key_field)
            "options": {"allowDuplicate": True},
            "tags": tags or [],
        }
        return self.invoke("addNote", note=note)

    def update_fields(self, note_id: int, fields: dict[str, str]) -> None:
        self.invoke("updateNoteFields", note={"id": note_id, "fields": fields})

    def create_deck(self, name: str) -> int:
        return self.invoke("createDeck", deck=name)

    def create_model(self, name: str, fields: list[str], css: str, cards: list[dict]) -> dict:
        return self.invoke(
            "createModel", modelName=name, inOrderFields=fields, css=css, isCloze=False,
            cardTemplates=[{"Name": str(c["name"]), "Front": str(c["front"]), "Back": str(c["back"])}
                           for c in cards])

    def store_media(self, filename: str, data: bytes) -> None:
        self.invoke("storeMediaFile", filename=filename,
                    data=base64.b64encode(data).decode("ascii"))


# --------------------------------------------------------------------------- #
# Подготовка Anki: колода и тип заметки
# --------------------------------------------------------------------------- #
TEMPLATE_TOKEN = re.compile(r"\{\{([^{}]+)\}\}")
SPECIAL_TOKENS = {"frontside", "tags", "type", "deck", "subdeck", "card", "cardflag"}
SOUND_TOKEN = re.compile(r"\[sound:([^\]]+)\]")


def validate_deck_name(name) -> str:
    """Проверяет и нормализует имя колоды (вложенность через '::')."""
    if not isinstance(name, str) or not name.strip():
        raise ValueError(t("deck_empty"))
    if re.search(r"[\x00-\x1f]", name):
        raise ValueError(t("deck_ctrl", name=name))
    parts = [p.strip() for p in name.split("::")]
    if any(not p for p in parts):
        raise ValueError(t("deck_part", name=name))
    return "::".join(parts)


def validate_model_name(name) -> str:
    if not isinstance(name, str) or not name.strip():
        raise ValueError(t("model_empty"))
    if re.search(r"[\x00-\x1f]", name):
        raise ValueError(t("model_ctrl", name=name))
    return name.strip()


def template_fields(template: str) -> set[str]:
    """Имена полей, на которые ссылается шаблон карточки ({{hint:x}}, {{#x}} и т.п. учитываются)."""
    names = set()
    for token in TEMPLATE_TOKEN.findall(template):
        token = token.strip()
        if token.startswith(("#", "^", "/")):
            token = token[1:].strip()
        name = token.split(":")[-1].strip()
        if name and name.casefold() not in SPECIAL_TOKENS:
            names.add(name)
    return names


def ensure_deck(anki: AnkiClient, dry_run: bool = False) -> None:
    """Проверяет колоду; если её нет - создаёт (в dry-run только сообщает)."""
    name = validate_deck_name(anki.deck_name)
    existing = {d.casefold(): d for d in anki.deck_names()}   # Anki считает имена колод регистронезависимыми
    found = existing.get(name.casefold())
    if found:
        if found != name:
            log.warning(t("deck_case", found=found))
        anki.deck_name = found
        log.info(t("deck_exists", name=found))
        return
    anki.deck_name = name
    if dry_run:
        log.info(t("deck_dry", name=name))
        return
    anki.create_deck(name)
    log.info(t("deck_created", name=name))


def ensure_model(anki: AnkiClient, cfg: Config, dry_run: bool = False) -> list[str]:
    """Проверяет тип заметки; если его нет - создаёт по config.yaml.
    Возвращает список полей (реальных или планируемых в dry-run)."""
    name = validate_model_name(anki.model_name)
    existing = {m.casefold(): m for m in anki.model_names()}
    found = existing.get(name.casefold())
    if found:
        anki.model_name = found
        fields = anki.model_fields()
        log.info(t("model_exists", name=found))
        missing = [f for f in cfg.model_fields if f not in fields]
        if missing:
            log.warning(t("model_missing_fields", name=found, fields=missing))
        return fields

    anki.model_name = name
    fields = cfg.model_fields
    cards = cfg.note_type["cards"]

    # все поля из шаблонов должны существовать в типе, иначе карточки в Anki будут с ошибкой
    unknown = {}
    for card in cards:
        used = template_fields(str(card["front"])) | template_fields(str(card["back"]))
        bad = sorted(used - set(fields))
        if bad:
            unknown[card["name"]] = bad
    if unknown:
        raise ValueError(t("model_unknown_fields", unknown=unknown, fields=fields))

    if dry_run:
        log.info(t("model_dry", name=name, fields=len(fields), cards=len(cards),
                   names=", ".join(str(c["name"]) for c in cards)))
        log.info(t("model_dry_fields", fields=", ".join(fields)))
        return fields

    anki.create_model(name, fields, str(cfg.note_type["css"]), cards)
    log.info(t("model_created", name=name, fields=len(fields), cards=len(cards)))
    return fields


def check_template_media(anki: AnkiClient, cfg: Config) -> None:
    """Предупреждает, если файлы из [sound:...] в шаблонах отсутствуют в медиа-папке Anki."""
    files = set()
    for card in cfg.note_type["cards"]:
        files |= set(SOUND_TOKEN.findall(str(card["front"]) + str(card["back"])))
    for fname in sorted(files):
        try:
            if not anki.media_exists(fname):
                log.warning(t("media_missing", name=fname))
        except Exception as e:  # noqa: BLE001
            log.debug(t("media_check_fail", name=fname, err=e))


# --------------------------------------------------------------------------- #
# Импорт заметок
# --------------------------------------------------------------------------- #
def import_notes(anki: AnkiClient, cfg: Config, rows: list[dict[str, str]],
                 model_fields: list[str],
                 dry_run: bool = False) -> tuple[list[int], list[dict[str, str]]]:
    """Добавляет новые заметки. Возвращает (ID добавленных, строки, которые добавились бы в dry-run)."""
    key_field = cfg.key_field
    if key_field not in model_fields:
        raise RuntimeError(t("no_key_in_model", model=anki.model_name, key=key_field, fields=model_fields))

    send_fields = [f for f in cfg.text_fields if f in model_fields]
    skipped_fields = [f for f in cfg.text_fields if f not in model_fields]
    if skipped_fields:
        log.warning(t("fields_not_in_model", fields=skipped_fields))

    if rows:
        missing_cols = [f for f in send_fields if f not in rows[0]]
        if key_field in missing_cols:
            raise RuntimeError(t("no_key_in_excel", key=key_field, cols=list(rows[0])))
        if missing_cols:
            log.warning(t("cols_missing", cols=missing_cols))

    existing = {
        clean_text(n["fields"][key_field]["value"]).casefold()
        for n in anki.deck_notes() if key_field in n["fields"]
    }

    new_ids: list[int] = []
    planned: list[dict[str, str]] = []
    added = skipped = failed = 0
    for row in rows:
        key_raw = row.get(key_field, "")
        key = clean_text(key_raw).casefold()
        if not key:
            log.warning(t("row_no_key", key=key_field, row=row))
            skipped += 1
            continue
        if key in existing:
            log.info(t("already_exists", key=key_raw))
            skipped += 1
            continue

        fields = {f: row.get(f, "") for f in send_fields}
        if dry_run:
            log.info(t("dry_will_add", key=key_raw))
            existing.add(key)
            planned.append(row)
            added += 1
            continue
        try:
            note_id = retry(lambda: anki.add_note(fields), what=t("op_add", key=key_raw))
        except Exception as e:  # noqa: BLE001
            log.error(t("add_failed", key=key_raw, err=e))
            failed += 1
            continue
        existing.add(key)
        new_ids.append(note_id)
        added += 1
        log.info(t("note_added", id=note_id, key=key_raw))

    log.info(t("import_summary", dry=" (dry-run)" if dry_run else "",
               added=added, skipped=skipped, failed=failed))
    return new_ids, planned


def plan_audio_for_rows(cfg: Config, rows: list[dict[str, str]]) -> None:
    """dry-run: сколько текстов будет озвучено для заметок, которые ещё не добавлены."""
    texts = set()
    fields = 0
    for row in rows:
        for src in cfg.audio_map:
            text = clean_text(row.get(src, ""))
            if text:
                fields += 1
                texts.add(text)
    if fields:
        log.info(t("plan_audio", fields=fields, texts=len(texts)))


# --------------------------------------------------------------------------- #
# TTS-движки: общий интерфейс synthesize_many(texts) -> {text: bytes | None}
# --------------------------------------------------------------------------- #
class GoogleTTS:
    name = "gtts"

    def __init__(self, lang: str, workers: int):
        from gtts import gTTS
        self._gTTS = gTTS
        self.lang = lang
        self.workers = workers
        self.voice = lang

    def _one(self, text: str) -> bytes:
        def run():
            buf = io.BytesIO()
            self._gTTS(text=text, lang=self.lang).write_to_fp(buf)
            return buf.getvalue()
        return retry(run, what=f"gTTS '{text[:30]}'")

    def _safe(self, text: str):
        try:
            return text, self._one(text)
        except Exception as e:  # noqa: BLE001
            log.error(t("gtts_failed", text=text[:50], err=e))
            return text, None

    def synthesize_many(self, texts: list[str], on_done=None) -> dict[str, bytes | None]:
        def task(text: str):
            result = self._safe(text)
            if on_done:
                on_done()           # текст обработан (успешно или с ошибкой)
            return result

        with ThreadPoolExecutor(max_workers=self.workers) as pool:
            return dict(pool.map(task, texts))


class EdgeTTS:
    name = "edge"

    def __init__(self, voice: str, workers: int):
        import edge_tts
        self._edge = edge_tts
        self.voice = voice
        self.workers = workers

    async def _one(self, text: str, sem: asyncio.Semaphore) -> tuple[str, bytes | None]:
        attempts = RETRY["attempts"]
        async with sem:
            for attempt in range(1, attempts + 1):
                try:
                    chunks = []
                    async for chunk in self._edge.Communicate(text, self.voice).stream():
                        if chunk["type"] == "audio":
                            chunks.append(chunk["data"])
                    data = b"".join(chunks)
                    if not data:
                        raise RuntimeError(t("edge_empty"))
                    return text, data
                except Exception as e:  # noqa: BLE001
                    if attempt == attempts:
                        log.error(t("edge_failed", text=text[:50], err=e))
                        return text, None
                    await asyncio.sleep(RETRY["delay"] * attempt)
        return text, None

    async def _many(self, texts: list[str], on_done=None) -> dict[str, bytes | None]:
        sem = asyncio.Semaphore(self.workers)

        async def tracked(text: str):
            result = await self._one(text, sem)
            if on_done:
                on_done()
            return result

        return dict(await asyncio.gather(*(tracked(x) for x in texts)))

    def synthesize_many(self, texts: list[str], on_done=None) -> dict[str, bytes | None]:
        return asyncio.run(self._many(texts, on_done))   # один event loop на весь пакет


def make_engine(cfg: Config):
    a = cfg.audio
    workers = int(a["workers"])
    if a["engine"] == "edge":
        return EdgeTTS(voice=a["edge_voice"], workers=workers)
    return GoogleTTS(lang=a["gtts_lang"], workers=workers)


# --------------------------------------------------------------------------- #
# Озвучка заметок
# --------------------------------------------------------------------------- #
@dataclass
class AudioJob:
    note_id: int
    dest_field: str
    text: str


def collect_audio_jobs(notes: list[dict], audio_map: dict[str, str],
                       only_ids: set[int] | None = None,
                       overwrite: bool = False) -> list[AudioJob]:
    jobs = []
    for note in notes:
        if only_ids is not None and note["noteId"] not in only_ids:
            continue
        fields = note["fields"]
        for src, dst in audio_map.items():
            if src not in fields or dst not in fields:
                continue
            text = clean_text(fields[src]["value"])
            if not text:
                continue
            if fields[dst]["value"].strip() and not overwrite:
                continue
            jobs.append(AudioJob(note["noteId"], dst, text))
    return jobs


def audio_filename(engine, text: str) -> str:
    digest = hashlib.sha1(f"{engine.name}|{engine.voice}|{text}".encode("utf-8")).hexdigest()[:20]
    return f"tts-{engine.name}-{digest}.mp3"


def build_audio(engine, texts: list[str], cache_dir: Path) -> dict[str, tuple[str, bytes]]:
    """Возвращает {text: (filename, bytes)}; использует дисковый кеш."""
    cache_dir.mkdir(parents=True, exist_ok=True)
    result: dict[str, tuple[str, bytes]] = {}
    missing: list[str] = []
    for text in texts:
        name = audio_filename(engine, text)
        path = cache_dir / name
        if path.exists() and path.stat().st_size > 0:
            result[text] = (name, path.read_bytes())
        else:
            missing.append(text)

    if missing:
        log.info(t("tts_generating", n=len(missing), engine=engine.name))
        with progress(len(missing), t("bar_tts"), t("unit_text")) as bar:
            generated = engine.synthesize_many(missing, on_done=bar.update)
        for text, data in generated.items():
            if data:
                name = audio_filename(engine, text)
                (cache_dir / name).write_bytes(data)
                result[text] = (name, data)
    return result


def make_audio_records(anki: AnkiClient, cfg: Config, engine,
                       only_ids: set[int] | None = None,
                       overwrite: bool = False, dry_run: bool = False) -> None:
    notes = anki.deck_notes()
    jobs = collect_audio_jobs(notes, cfg.audio_map, only_ids, overwrite)
    if not jobs:
        log.info(t("audio_nothing"))
        return

    texts = sorted({j.text for j in jobs})
    log.info(t("audio_plan", fields=len(jobs), texts=len(texts)))
    if dry_run:
        for j in jobs:
            log.info(t("audio_dry_job", id=j.note_id, field=j.dest_field, text=j.text[:60]))
        return

    audio = build_audio(engine, texts, cfg.cache_dir)

    to_upload: dict[str, bytes] = {}
    updates: dict[int, dict[str, str]] = {}
    for job in jobs:
        if job.text not in audio:
            continue            # озвучить не удалось, при следующем запуске повторится
        name, data = audio[job.text]
        to_upload[name] = data
        updates.setdefault(job.note_id, {})[job.dest_field] = f"[sound:{name}]"

    with progress(len(to_upload), t("bar_upload"), t("unit_file")) as bar:
        for name, data in to_upload.items():
            retry(lambda: anki.store_media(name, data), what=t("op_upload", name=name))
            bar.update()

    # одно обновление на заметку, сразу со всеми аудио-полями
    with progress(len(updates), t("bar_update"), t("unit_note")) as bar:
        for note_id, fields in updates.items():
            try:
                retry(lambda: anki.update_fields(note_id, fields), what=t("op_update_note", id=note_id))
            except Exception as e:  # noqa: BLE001
                log.error(t("update_failed", id=note_id, err=e))
            bar.update()
    log.info(t("audio_done", notes=len(updates), files=len(to_upload)))


def cleanup_cache(cache_dir: Path) -> None:
    """Удаляет сгенерированные .mp3 из временной папки (и саму папку, если пуста)."""
    if not cache_dir.is_dir():
        return
    removed = 0
    for f in cache_dir.glob("tts-*.mp3"):
        try:
            f.unlink()
            removed += 1
        except OSError as e:
            log.warning(t("cache_remove_failed", path=f, err=e))
    try:
        cache_dir.rmdir()       # удалится только если папка пуста
    except OSError:
        pass
    if removed:
        log.info(t("cache_removed", n=removed))


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="Импорт из Excel в Anki + озвучка. Настройки: config.yaml / "
                    "Excel-Import nach Anki + Vertonung. Einstellungen: config.yaml")
    p.add_argument("--config", default=str(DEFAULT_CONFIG),
                   help="Файл конфигурации (по умолчанию config.yaml рядом со скриптом)")
    # переопределения значений из конфига
    p.add_argument("--lang", choices=list(SUPPORTED_LANGS),
                   help="Язык сообщений / Sprache der Meldungen (перекрывает ui.language)")
    p.add_argument("--file", help="Excel-файл (перекрывает excel.file)")
    p.add_argument("--sheet", help="Имя листа (перекрывает excel.sheet)")
    p.add_argument("--deck", help="Колода (перекрывает anki.deck)")
    p.add_argument("--model", help="Тип заметки (перекрывает anki.model)")
    p.add_argument("--url", help="URL AnkiConnect (перекрывает anki.url)")
    p.add_argument("--engine", choices=["gtts", "edge"], help="Движок озвучки")
    p.add_argument("--voice", help="Голос для edge-tts")
    p.add_argument("--workers", type=int, help="Параллельных запросов к TTS")
    p.add_argument("--cache-dir", help="Временная папка для .mp3")
    p.add_argument("--keep-audio", action="store_true", default=None,
                   help="Не удалять .mp3 после выполнения (использовать как кеш)")
    # режимы
    p.add_argument("--dry-run", action="store_true", help="Ничего не менять, только показать план")
    p.add_argument("--no-audio", action="store_true", help="Не озвучивать")
    p.add_argument("--audio-only", action="store_true", help="Только озвучка, без импорта")
    p.add_argument("--new-only", action="store_true",
                   help="Озвучивать только заметки, добавленные в этом запуске")
    p.add_argument("--overwrite-audio", action="store_true",
                   help="Перезаписывать уже заполненные аудио-поля")
    p.add_argument("-v", "--verbose", action="store_true", help="Подробный лог")
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s: %(message)s")

    if args.lang:                       # чтобы и ошибки чтения конфига были на нужном языке
        set_lang(args.lang)

    cfg = load_config(Path(args.config).expanduser(), overrides={
        ("ui", "language"): args.lang,
        ("excel", "file"): args.file,
        ("excel", "sheet"): args.sheet,
        ("anki", "deck"): args.deck,
        ("anki", "model"): args.model,
        ("anki", "url"): args.url,
        ("audio", "engine"): args.engine,
        ("audio", "edge_voice"): args.voice,
        ("audio", "workers"): args.workers,
        ("audio", "cache_dir"): args.cache_dir,
        ("audio", "keep_files"): True if args.keep_audio else None,
    })

    dry = args.dry_run
    excel_missing = not args.audio_only and not cfg.xlsx_path.is_file()

    # Excel нет: создаём шапку и останавливаемся (Anki для этого не нужен)
    if excel_missing and not dry:
        create_excel_template(cfg)
        log.warning(t("excel_created", path=cfg.xlsx_path, sheet=cfg.sheet, cols=len(cfg.text_fields)))
        log.warning(t("excel_fill"))
        return 0

    if dry:
        log.info(t("dry_mode"))

    anki = AnkiClient(cfg.anki["url"], cfg.anki["deck"], cfg.anki["model"],
                      timeout=int(cfg.anki["timeout"]))
    anki.invoke("version")      # быстрая проверка соединения

    # колода и тип заметки: проверить, при необходимости создать
    ensure_deck(anki, dry_run=dry)
    model_fields = ensure_model(anki, cfg, dry_run=dry)
    check_template_media(anki, cfg)

    new_ids: list[int] = []
    planned_rows: list[dict[str, str]] = []
    if not args.audio_only:
        if excel_missing:           # только dry-run
            log.info(t("excel_missing_dry", path=cfg.xlsx_path))
        else:
            rows = read_xlsx_file(cfg.xlsx_path, cfg.sheet)
            log.info(t("rows_read", n=len(rows), name=cfg.xlsx_path.name))
            new_ids, planned_rows = import_notes(anki, cfg, rows, model_fields, dry_run=dry)

    if not args.no_audio:
        engine = make_engine(cfg)
        only_ids = set(new_ids) if args.new_only else None
        if tqdm is None and not dry:
            log.info(t("tqdm_tip"))
        try:
            with log_with_progress():
                if dry and planned_rows:
                    plan_audio_for_rows(cfg, planned_rows)
                make_audio_records(anki, cfg, engine, only_ids=only_ids,
                                   overwrite=args.overwrite_audio, dry_run=dry)
        finally:
            if not cfg.audio["keep_files"] and not dry:
                cleanup_cache(cfg.cache_dir)

    if dry:
        log.info(t("dry_done"))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (RuntimeError, ValueError, FileNotFoundError, yaml.YAMLError) as e:
        log.error("%s", e)
        sys.exit(1)