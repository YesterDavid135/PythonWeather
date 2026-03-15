"""Hauptprogramm: Wetterdaten abrufen und als Excel-Datei exportieren.

Verwendung:
    python main.py <Stadt1> [Stadt2] [Stadt3] ...

Beispiel:
    python main.py Berlin
    python main.py Berlin Hamburg Muenchen
"""

from __future__ import annotations

import argparse
import asyncio
import sys

from weather_fetcher import WeatherDataFetcher
from data_processor import DataProcessor
from excel_exporter import ExcelExporter


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Wetterdaten abrufen und als Excel-Datei exportieren."
    )
    parser.add_argument(
        "cities",
        nargs="+",
        help="Ein oder mehrere Staedtenamen (z.B. Basel Zürich)",
    )
    return parser.parse_args()


async def main() -> None:
    args = parse_args()

    fetcher = WeatherDataFetcher()
    processor = DataProcessor()
    exporter = ExcelExporter()

    print(f"Rufe Wetterdaten ab für: {', '.join(args.cities)}\n")

    results = await fetcher.fetch_multiple(args.cities)

    if not results:
        print("\nKeine Wetterdaten konnten abgerufen werden.")
        sys.exit(1)

    print()

    for city, forecast in results.items():
        current_data = processor.process_current_weather(forecast)
        daily_data = processor.process_daily_forecasts(forecast)
        hourly_data = processor.process_hourly_forecasts(forecast)

        try:
            filepath = exporter.export(city, current_data, daily_data, hourly_data)
            print(f"Excel-Datei erstellt: {filepath}")
        except PermissionError as e:
            print(f"Fehler: {e}")

    print("\nFertig.")


if __name__ == "__main__":
    asyncio.run(main())
