# PythonWeather

Ruft aktuelle Wetterdaten und Vorhersagen für beliebige Städte ab und exportiert sie als formatierte Excel-Datei.

## Features

- Aktuelles Wetter und mehrtägige Tagesvorhersage
- Stündliche Vorhersage
- Excel-Export mit mehreren Tabellenblättern
- Diagramme: Temperaturverlauf (täglich) und Diagramm Temperatur & Niederschlag (stündlich)
- Conditional formatting für Temperaturspalten
- Mehrere Städte gleichzeitig möglich

## Installation

```bash
pip install -r requirements.txt
```

## Verwendung

```bash
python main.py <Stadt1> [Stadt2] [Stadt3] ...
```

**Beispiele:**

```bash
python main.py Basel
python main.py Basel Bern Zürich
```

Die Excel-Dateien werden im Verzeichnis `./out/` gespeichert:

```
out/Wetter_Basel_2026-03-15_14-30.xlsx
```

## Abhängigkeiten

- `python-weather` – Wetterdaten via wttr.in
- `openpyxl` – Excel-Erstellung und Diagramme
