#!/usr/bin/env python3
"""
Импорт заметок из Excel в Anki (через AnkiConnect) + озвучка (gTTS / edge-tts).
Все настройки лежат в config.yaml рядом со скриптом.

Примеры запуска:
    python3 anki_import.py                      # импорт + озвучка
    python3 anki_import.py --dry-run            # ничего не менять, только показать план
    python3 anki_import.py --engine edge        # озвучка через edge-tts
    python3 anki_import.py --no-audio           # только импорт
    python3 anki_import.py --audio-only         # только озвучка пустых аудио-полей
    python3 anki_import.py --translate          # заполнить пустые base_ru переводом
    python3 anki_import.py --config my.yaml     # другой файл настроек
    python3 anki_import.py -h                   # все параметры

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
    sys.exit("Нужен пакет PyYAML: pip install pyyaml")

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
# Конфигурация
# --------------------------------------------------------------------------- #
REQUIRED_KEYS = {
    "anki": ["url", "deck", "model", "timeout"],
    "excel": ["file", "sheet"],
    "fields": ["key", "base", "sentences_count", "sentence_fields"],
    "audio": ["engine", "gtts_lang", "edge_voice", "workers", "cache_dir", "keep_files",
              "main_source", "main_dest", "sentence_source", "sentence_dest"],
    "translate": ["source_lang", "target_lang", "source_field", "target_field"],
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
        self.translate = raw["translate"]

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

    def resolve(self, value: str) -> Path:
        """Относительные пути считаем от папки скрипта."""
        p = Path(str(value)).expanduser()
        return p if p.is_absolute() else self.base_dir / p


def load_config(path: Path, overrides: dict[tuple[str, str], object] | None = None) -> Config:
    if not path.is_file():
        raise FileNotFoundError(f"Файл конфигурации не найден: {path}")
    with open(path, encoding="utf-8") as f:
        raw = yaml.safe_load(f)
    if not isinstance(raw, dict):
        raise ValueError(f"Файл конфигурации пуст или имеет неверный формат: {path}")

    # значения из командной строки перекрывают yaml
    for (section, key), value in (overrides or {}).items():
        if value is not None:
            raw.setdefault(section, {})[key] = value

    missing = [f"{sec}.{key}" for sec, keys in REQUIRED_KEYS.items() for key in keys
               if not isinstance(raw.get(sec), dict) or key not in raw[sec]]
    if missing:
        raise ValueError(f"В {path.name} не хватает параметров: {', '.join(missing)}")
    if raw["audio"]["engine"] not in ("gtts", "edge"):
        raise ValueError("audio.engine должен быть 'gtts' или 'edge'")

    RETRY["attempts"] = int(raw["retry"]["attempts"])
    RETRY["delay"] = float(raw["retry"]["delay"])
    return Config(raw, path.resolve().parent)


# --------------------------------------------------------------------------- #
# Вспомогательные функции
# --------------------------------------------------------------------------- #
def retry(func, what: str = "операция"):
    """Вызывает func(); при ошибке повторяет с растущей паузой (параметры из конфига)."""
    attempts = RETRY["attempts"]
    for n in range(1, attempts + 1):
        try:
            return func()
        except Exception as e:  # noqa: BLE001
            if n == attempts:
                raise
            log.warning("%s: попытка %d/%d не удалась (%s), повтор...", what, n, attempts, e)
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


def progress(total: int, desc: str, unit: str = "шт"):
    """Полоса загрузки в stderr (потокобезопасная). Без tqdm — пустая заглушка."""
    if tqdm is None or total <= 0:
        return _NoBar()
    return tqdm(total=total, desc=desc, unit=unit, file=sys.stderr, dynamic_ncols=True)


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
            raise ValueError(
                f"Лист '{sheet_name}' не найден. Доступные листы: {workbook.sheetnames}"
            )
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
            raise RuntimeError(
                f"Не удалось подключиться к AnkiConnect ({self.url}). "
                "Anki запущен и установлен дополнительный модуль AnkiConnect?"
            ) from e
        if data.get("error") is not None:
            raise RuntimeError(f"AnkiConnect ({action}): {data['error']}")
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

    def store_media(self, filename: str, data: bytes) -> None:
        self.invoke("storeMediaFile", filename=filename,
                    data=base64.b64encode(data).decode("ascii"))


# --------------------------------------------------------------------------- #
# Импорт заметок
# --------------------------------------------------------------------------- #
def import_notes(anki: AnkiClient, cfg: Config, rows: list[dict[str, str]],
                 dry_run: bool = False) -> list[int]:
    """Добавляет новые заметки. Возвращает ID добавленных."""
    key_field = cfg.key_field
    model_fields = anki.model_fields()
    if key_field not in model_fields:
        raise RuntimeError(f"В модели '{anki.model_name}' нет поля '{key_field}'. Поля: {model_fields}")

    send_fields = [f for f in cfg.text_fields if f in model_fields]
    skipped_fields = [f for f in cfg.text_fields if f not in model_fields]
    if skipped_fields:
        log.warning("Этих полей нет в модели Anki, они не будут заполнены: %s", skipped_fields)

    if rows:
        missing_cols = [f for f in send_fields if f not in rows[0]]
        if key_field in missing_cols:
            raise RuntimeError(f"В Excel нет колонки '{key_field}'. Колонки: {list(rows[0])}")
        if missing_cols:
            log.warning("Этих колонок нет в Excel, поля будут пустыми: %s", missing_cols)

    existing = {
        clean_text(n["fields"][key_field]["value"]).casefold()
        for n in anki.deck_notes() if key_field in n["fields"]
    }

    new_ids: list[int] = []
    added = skipped = failed = 0
    for row in rows:
        key_raw = row.get(key_field, "")
        key = clean_text(key_raw).casefold()
        if not key:
            log.warning("Строка без значения в '%s' пропущена: %s", key_field, row)
            skipped += 1
            continue
        if key in existing:
            log.info("Уже существует: %s", key_raw)
            skipped += 1
            continue

        fields = {f: row.get(f, "") for f in send_fields}
        if dry_run:
            log.info("[dry-run] будет добавлено: %s", key_raw)
            existing.add(key)
            added += 1
            continue
        try:
            note_id = retry(lambda: anki.add_note(fields), what=f"добавление '{key_raw}'")
        except Exception as e:  # noqa: BLE001
            log.error("Не удалось добавить '%s': %s", key_raw, e)
            failed += 1
            continue
        existing.add(key)
        new_ids.append(note_id)
        added += 1
        log.info("Добавлена заметка %s: %s", note_id, key_raw)

    log.info("Импорт: добавлено %d, пропущено %d, ошибок %d", added, skipped, failed)
    return new_ids


# --------------------------------------------------------------------------- #
# Перевод (необязательно)
# --------------------------------------------------------------------------- #
def translate_empty(anki: AnkiClient, cfg: Config, dry_run: bool = False) -> None:
    try:
        from deep_translator import GoogleTranslator
    except ImportError:
        raise RuntimeError("Для --translate нужен пакет: pip install deep-translator")
    t = cfg.translate
    translator = GoogleTranslator(source=t["source_lang"], target=t["target_lang"])
    src_field, dst_field = t["source_field"], t["target_field"]

    notes = [n for n in anki.deck_notes() if src_field in n["fields"] and dst_field in n["fields"]]
    for i, note in enumerate(notes, 1):
        src = clean_text(note["fields"][src_field]["value"])
        dst = note["fields"][dst_field]["value"].strip()
        if not src or dst:
            continue
        log.info("[%d/%d] перевод: %s", i, len(notes), src)
        if dry_run:
            continue
        try:
            translated = retry(lambda: translator.translate(src), what=f"перевод '{src}'")
            anki.update_fields(note["noteId"], {dst_field: translated})
        except Exception as e:  # noqa: BLE001
            log.error("Не удалось перевести '%s': %s", src, e)


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
            log.error("gTTS не смог озвучить '%s': %s", text[:50], e)
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
                        raise RuntimeError("пустой ответ")
                    return text, data
                except Exception as e:  # noqa: BLE001
                    if attempt == attempts:
                        log.error("edge-tts не смог озвучить '%s': %s", text[:50], e)
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

        return dict(await asyncio.gather(*(tracked(t) for t in texts)))

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
        log.info("Генерация озвучки: %d уникальных текстов (%s)", len(missing), engine.name)
        with progress(len(missing), "Озвучка", "текст") as bar:
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
        log.info("Озвучка: нечего делать (нет объектов для озвучки в очереди).")
        return

    texts = sorted({j.text for j in jobs})
    log.info("Озвучка: полей %d, уникальных текстов %d", len(jobs), len(texts))
    if dry_run:
        for j in jobs:
            log.info("[dry-run] заметка %s -> %s: %s", j.note_id, j.dest_field, j.text[:60])
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

    with progress(len(to_upload), "Загрузка аудио в Anki", "файл") as bar:
        for name, data in to_upload.items():
            retry(lambda: anki.store_media(name, data), what=f"загрузка {name}")
            bar.update()

    # одно обновление на заметку, сразу со всеми аудио-полями
    with progress(len(updates), "Обновление заметок", "заметка") as bar:
        for note_id, fields in updates.items():
            try:
                retry(lambda: anki.update_fields(note_id, fields), what=f"обновление заметки {note_id}")
            except Exception as e:  # noqa: BLE001
                log.error("Не удалось обновить заметку %s: %s", note_id, e)
            bar.update()
    log.info("Озвучка: обновлено заметок %d, файлов загружено %d", len(updates), len(to_upload))


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
            log.warning("Не удалось удалить %s: %s", f, e)
    try:
        cache_dir.rmdir()       # удалится только если папка пуста
    except OSError:
        pass
    if removed:
        log.info("Удалено временных .mp3: %d", removed)


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="Импорт из Excel в Anki + озвучка. Настройки: config.yaml")
    p.add_argument("--config", default=str(DEFAULT_CONFIG),
                   help="Файл конфигурации (по умолчанию config.yaml рядом со скриптом)")
    # переопределения значений из конфига
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
    p.add_argument("--translate", action="store_true",
                   help="Заполнить пустые поля перевода (нужен deep-translator)")
    p.add_argument("-v", "--verbose", action="store_true", help="Подробный лог")
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s: %(message)s")

    cfg = load_config(Path(args.config).expanduser(), overrides={
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

    new_ids: list[int] = []
    if not args.audio_only:
        # Excel-файл должен существовать; если нет, создаём шапку и останавливаемся
        if not cfg.xlsx_path.is_file():
            if args.dry_run:
                log.info("[dry-run] Excel-файл %s не найден, будет создан с шапкой.", cfg.xlsx_path)
            else:
                create_excel_template(cfg)
                log.warning("Excel-файл не найден, создан новый: %s (лист '%s', колонок: %d)",
                            cfg.xlsx_path, cfg.sheet, len(cfg.text_fields))
                log.warning("Заполните его данными и запустите скрипт снова.")
            return 0

    anki = AnkiClient(cfg.anki["url"], cfg.anki["deck"], cfg.anki["model"],
                      timeout=int(cfg.anki["timeout"]))
    anki.invoke("version")      # быстрая проверка соединения

    if not args.audio_only:
        rows = read_xlsx_file(cfg.xlsx_path, cfg.sheet)
        log.info("Прочитано строк из Excel: %d (%s)", len(rows), cfg.xlsx_path.name)
        new_ids = import_notes(anki, cfg, rows, dry_run=args.dry_run)

    if args.translate:
        translate_empty(anki, cfg, dry_run=args.dry_run)

    if not args.no_audio:
        engine = make_engine(cfg)
        only_ids = set(new_ids) if args.new_only else None
        if tqdm is None and not args.dry_run:
            log.info("Совет: pip install tqdm — появится полоса загрузки при озвучке.")
        try:
            with log_with_progress():
                make_audio_records(anki, cfg, engine, only_ids=only_ids,
                                   overwrite=args.overwrite_audio, dry_run=args.dry_run)
        finally:
            if not cfg.audio["keep_files"]:
                cleanup_cache(cfg.cache_dir)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (RuntimeError, ValueError, FileNotFoundError, yaml.YAMLError) as e:
        log.error("%s", e)
        sys.exit(1)