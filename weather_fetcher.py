from __future__ import annotations

import python_weather
from aiohttp import ClientError


class WeatherFetchError(Exception):
    pass


class WeatherDataFetcher:
    async def fetch(self, city: str) -> python_weather.forecast.Forecast:
        """Ruft Wetterdaten für eine einzelne Stadt ab."""
        if not city or not city.strip():
            raise WeatherFetchError("Ungültige Eingabe: Stadtname darf nicht leer sein.")

        try:
            async with python_weather.Client(
                unit=python_weather.METRIC,
                locale=python_weather.Locale.GERMAN
            ) as client:
                weather = await client.get(city.strip())

            if not weather.daily_forecasts:
                raise WeatherFetchError(
                    f"Keine Wetterdaten für '{city}' gefunden."
                )

            return weather

        except ClientError:
            raise WeatherFetchError(
                "Keine Internetverbindung. Bitte Netzwerk prüfen."
            )
        except WeatherFetchError:
            raise
        except Exception as e:
            raise WeatherFetchError(
                f"Fehler beim Abruf der Wetterdaten für '{city}': {e}"
            )

    async def fetch_multiple(self, cities: list[str]) -> dict[str, python_weather.forecast.Forecast]:
        """Ruft Wetterdaten für mehrere Städte ab."""
        results = {}
        for city in cities:
            try:
                forecast = await self.fetch(city)
                results[city.strip()] = forecast
                print(f"Wetterdaten für '{city}' erfolgreich abgerufen.")
            except WeatherFetchError as e:
                print(f"Warnung: {e}")
        return results
