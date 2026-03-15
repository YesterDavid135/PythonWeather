# PythonWeather

Fetches current weather data and forecasts for any city and exports them as a formatted Excel file.

## Features

- Current weather and multi-day daily forecast
- Hourly forecast
- Excel export with multiple sheets
- Charts: temperature trend (daily) and combined temperature & precipitation chart (hourly)
- Conditional formatting for temperature columns
- Multiple cities can be processed at once

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python main.py <city1> [city2] [city3] ...
```

**Examples:**

```bash
python main.py Basel
python main.py Basel Bern Zurich
```

Excel files are saved in the `./out/` directory:

```
out/Wetter_Basel_2026-03-15_14-30.xlsx
```

## Dependencies

- `python-weather` – weather data via wttr.in
- `openpyxl` – Excel creation and charts

---

# PythonWeather

Ruft aktuelle Wetterdaten und Vorhersagen für beliebige Städte ab und exportiert sie als formatierte Excel-Datei.

## Features

- Aktuelles Wetter und mehrtägige Tagesvorhersage
- Stündliche Vorhersage
- Excel-Export mit mehreren Tabellenblättern
- Diagramme: Temperaturverlauf (täglich) und Diagramm Temperatur & Niederschlag (stündlich)
- Conditional formatting für Temperaturspalten
- Mehrere Städte gleichzeitig möglich
- Automatisches Erstellen des Output-Folders

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
