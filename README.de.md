# anki_import

[🇷🇺 Русский](README.md) · 🇩🇪 Deutsch

Das Skript importiert Notizen aus Excel nach Anki (über [AnkiConnect](https://ankiweb.net/shared/info/2055492159)) und vertont ausgewählte Felder mit **gTTS** oder **edge-tts**. Es eignet sich zum Vokabellernen mit Beispielsätzen (standardmäßig für Deutsch → Russisch konfiguriert).

## Funktionen

- Import von Excel-Zeilen in einen Anki-Stapel; Duplikate werden anhand eines Schlüsselfelds erkannt und übersprungen.
- Automatisches Anlegen von Stapel und Notiztyp (Felder, Kartenvorlagen und CSS stammen aus `config.yaml`), falls sie noch nicht existieren.
- Vertonung leerer Audiofelder: Die Dateien werden in die Anki-Medien hochgeladen, im Feld steht anschließend `[sound:...]`.
- Zwei Sprachausgabe-Engines: `gtts` (Standard) und `edge` (Microsoft-Edge-Stimmen).
- Parallele Vertonung, Wiederholungsversuche bei Netzwerkfehlern, `.mp3`-Cache auf der Festplatte, Fortschrittsbalken (wenn `tqdm` installiert ist).
- Modus `--dry-run`: Alle Prüfungen laufen, aber nichts wird erstellt oder geändert.
- Konsolenausgabe wahlweise auf Russisch oder Deutsch.

## Voraussetzungen

- Python 3.8+
- Laufendes Anki mit dem Add-on **AnkiConnect** (Code `2055492159`)
- Internetverbindung (für die Vertonung)

## Installation

```bash
git clone <Repository-URL>
cd <Repository-Ordner>
pip install -r requirements.txt
```

Abhängigkeiten:

| Paket | Zweck | Pflicht |
|---|---|---|
| `openpyxl` | Excel lesen/erstellen | ja |
| `requests` | Anfragen an AnkiConnect | ja |
| `PyYAML` | `config.yaml` lesen | ja |
| `gTTS` | Vertonung (Standard-Engine) | ja |
| `edge-tts` | Vertonung (`audio.engine: edge`) | nein |
| `tqdm` | Fortschrittsbalken | nein |

## Schnellstart

1. Anki starten.
2. Einstellungen in `config.yaml` prüfen (Stapel, Notiztyp, Name der Excel-Datei).
3. Skript starten:
   ```bash
   python3 anki_import.py
   ```
   Existiert die Excel-Datei aus der Konfiguration nicht, erstellt das Skript sie neben sich mit einer Kopfzeile aus den Feldern und beendet sich. Datei füllen und das Skript erneut starten.
4. Empfehlung: zuerst den Plan ohne Änderungen ansehen:
   ```bash
   python3 anki_import.py --dry-run
   ```

## Beispiele

```bash
python3 anki_import.py                      # Import + Vertonung
python3 anki_import.py --lang de            # Meldungen auf Deutsch
python3 anki_import.py --dry-run            # nichts ändern, nur den Plan zeigen
python3 anki_import.py --engine edge        # Vertonung mit edge-tts
python3 anki_import.py --no-audio           # nur Import
python3 anki_import.py --audio-only         # nur leere Audiofelder vertonen
python3 anki_import.py --config my.yaml     # andere Konfigurationsdatei
python3 anki_import.py -h                   # alle Parameter
```

## Kommandozeilenparameter

Parameter überschreiben die Werte aus `config.yaml`.

| Parameter | Beschreibung |
|---|---|
| `--config` | Konfigurationsdatei (Standard: `config.yaml` neben dem Skript) |
| `--lang {ru,de}` | Sprache der Meldungen (`ui.language`) |
| `--file` | Excel-Datei (`excel.file`) |
| `--sheet` | Name des Tabellenblatts (`excel.sheet`) |
| `--deck` | Stapel (`anki.deck`) |
| `--model` | Notiztyp (`anki.model`) |
| `--url` | AnkiConnect-URL (`anki.url`) |
| `--engine {gtts,edge}` | Vertonungs-Engine (`audio.engine`) |
| `--voice` | Stimme für edge-tts (`audio.edge_voice`) |
| `--workers` | Anzahl paralleler TTS-Anfragen (`audio.workers`) |
| `--cache-dir` | temporärer Ordner für `.mp3` (`audio.cache_dir`) |
| `--keep-audio` | `.mp3` nach dem Lauf nicht löschen (Cache) |
| `--dry-run` | nichts ändern, nur den Plan zeigen |
| `--no-audio` | nicht vertonen |
| `--audio-only` | nur vertonen, kein Import |
| `--new-only` | nur in diesem Lauf hinzugefügte Notizen vertonen |
| `--overwrite-audio` | bereits gefüllte Audiofelder überschreiben |
| `-v`, `--verbose` | ausführliches Log |

## Funktionsweise

1. **Anki prüfen.** Das Skript prüft die Verbindung sowie das Vorhandensein von Stapel und Notiztyp. Fehlendes wird angelegt (mit `--dry-run` nur gemeldet). Namen werden ohne Beachtung der Groß-/Kleinschreibung verglichen. Ein bereits vorhandener Notiztyp wird **nicht verändert**; das Skript warnt lediglich vor Feldern aus der Konfiguration, die ihm fehlen.
2. **Import.** Die Excel-Zeilen werden aus dem angegebenen Blatt gelesen; leere Zeilen werden übersprungen. Eine Notiz wird nur hinzugefügt, wenn der Wert des Schlüsselfelds (`fields.key`) im Stapel noch nicht vorkommt (Vergleich ohne HTML, `[sound:...]` und Groß-/Kleinschreibung).
3. **Vertonung.** Für Notizen des Stapels, deren Audiofeld leer und deren Quelltextfeld gefüllt ist, wird eine mp3 erzeugt. Gleiche Texte werden nur einmal vertont. Die Datei wird als `tts-<Engine>-<Hash>.mp3` in die Anki-Medien hochgeladen und im Feld als `[sound:...]` eingetragen. Texte, die nicht vertont werden konnten, werden beim nächsten Lauf erneut versucht.
4. **Aufräumen.** Der temporäre `.mp3`-Ordner wird gelöscht, sofern nicht `audio.keep_files: true` oder `--keep-audio` gesetzt ist.

## Excel-Format

- Die erste Zeile ist die Kopfzeile: Spaltennamen müssen den Feldnamen entsprechen (siehe `fields.base` und `fields.sentences`).
- Spalten für Audiofelder werden nicht benötigt, sie werden automatisch befüllt.
- Eine leere Zelle in der Schlüsselspalte: Die Zeile wird übersprungen.

Mit der Standardkonfiguration lauten die Spalten:

```
base_de | full_de | audio_text_de | base_tr | s1_de | s1_tr | ... | s9_de | s9_tr
```

## Konfiguration (`config.yaml`)

Relative Pfade beziehen sich auf den Ordner des Skripts.

| Abschnitt | Parameter |
|---|---|
| `ui` | `language` — Sprache der Konsole: `ru` oder `de` (Standard `ru`) |
| `anki` | `url`, `deck` (`::` = Verschachtelung), `model`, `timeout` |
| `excel` | `file`, `sheet` |
| `fields` | `key` — Feld zur Duplikaterkennung; `base` — normale Felder (zugleich Excel-Spalten); `sentences` — Liste der Sätze |
| `note_type` | `css` und `cards` (Name, Vorder- und Rückseite der Vorlage); wird **nur beim Erstellen** des Typs verwendet |
| `audio` | `engine`, `gtts_lang`, `edge_voice`, `workers`, `cache_dir`, `keep_files`, `main_source`, `main_dest` |
| `retry` | `attempts`, `delay` — Wiederholungen bei Netzwerkfehlern |

### Sätze

Jeder Satz wird explizit angegeben, die Feldnamen sind frei wählbar:

```yaml
sentences:
  - { de: s1_de, tr: s1_tr, audio: s1_audio }
```

- `de` — Feld mit dem Satz (Pflicht)
- `tr` — Feld mit der Übersetzung (Pflicht)
- `audio` — Feld für `[sound:...]` (optional; ohne dieses Feld wird der Satz nicht vertont)

Die Anzahl der Sätze entspricht der Anzahl der Listeneinträge.

### Vertonung des Hauptworts

`audio.main_source` ist das Feld, dessen Text vertont wird, `audio.main_dest` das Feld, in das der Ton eingetragen wird. Standard: `audio_text_de` → `base_audio`.

### Notiztyp

Die Felder des Typs werden automatisch gebildet: `fields.base` + `audio.main_dest` + je Satz `de`, `tr` und `audio` (falls angegeben). Alle `{{Felder}}` in den Vorlagen müssen darunter sein, sonst bricht das Skript vor dem Erstellen des Typs ab.


## Häufige Probleme

- **Verbindung zu AnkiConnect fehlgeschlagen** — prüfen, ob Anki läuft, AnkiConnect installiert ist und `anki.url` stimmt (Standard `http://localhost:8765`).
- **Im Notiztyp fehlen Felder aus der Konfiguration** — ein vorhandener Typ wird nicht geändert; Felder in Anki manuell ergänzen oder einen neuen Typ mit anderem Namen in `anki.model` anlegen lassen.
- **Vertonung schlägt fehl** — Internetverbindung prüfen; die Zahl paralleler Anfragen mit `--workers` verringern oder `--engine edge` versuchen.
- **Neu vertonen** — `--overwrite-audio` verwenden.

## Dateien

| Datei | Zweck |
|---|---|
| `anki_import.py` | Hauptskript |
| `config.yaml` | Einstellungen |
| `requirements.txt` | Abhängigkeiten |
| `.gitignore` | schließt `*.mp3` und `*.xlsx` aus |
