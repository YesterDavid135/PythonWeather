from __future__ import annotations

import python_weather.forecast


class DataProcessor:
    """Bereitet Wetterdaten für die Excel-Ausgabe auf."""

    def process_current_weather(self, forecast: python_weather.forecast.Forecast) -> dict:
        return {
            "Ort": forecast.location + ", " + forecast.country,
            "Datum/Uhrzeit": forecast.datetime.strftime("%d.%m.%Y %H:%M"),
            "Temperatur (°C)": forecast.temperature,
            "Gefühlt (°C)": forecast.feels_like,
            "Luftfeuchtigkeit (%)": forecast.humidity,
            "Windgeschwindigkeit (km/h)": forecast.wind_speed,
            "Windrichtung": str(forecast.wind_direction) + " " + forecast.wind_direction.emoji,
            "Luftdruck (hPa)": forecast.pressure,
            "Sichtweite (km)": forecast.visibility,
            "Beschreibung": forecast.description,
            "UV-Index": str(forecast.ultraviolet),
            "Niederschlag (mm)": forecast.precipitation,
        }

    def process_daily_forecasts(self, forecast: python_weather.forecast.Forecast) -> list[dict]:
        daily_data = []
        for daily in forecast:
            sunrise_str = daily.sunrise.strftime("%H:%M") if daily.sunrise else "N/A"
            sunset_str = daily.sunset.strftime("%H:%M") if daily.sunset else "N/A"

            daily_data.append({
                "Datum": daily.date.strftime("%d.%m.%Y"),
                "Temperatur (°C)": daily.temperature,
                "Höchsttemperatur (°C)": daily.highest_temperature,
                "Tiefsttemperatur (°C)": daily.lowest_temperature,
                "Sonnenstunden": daily.sunlight,
                "Sonnenaufgang": sunrise_str,
                "Sonnenuntergang": sunset_str,
                "Mondphase": str(daily.moon_phase) + " " + daily.moon_phase.emoji,
            })
        return daily_data

    def process_hourly_forecasts(self, forecast: python_weather.forecast.Forecast) -> list[dict]:
        hourly_data = []
        for daily in forecast:
            for hourly in daily:
                hourly_data.append({
                    "Datum": daily.date.strftime("%d.%m.%Y"),
                    "Uhrzeit": hourly.time.strftime("%H:%M"),
                    "Emoji": hourly.kind.emoji,
                    "Temperatur (°C)": hourly.temperature,
                    "Gefühlt (°C)": hourly.feels_like,
                    "Luftfeuchtigkeit (%)": hourly.humidity,
                    "Windgeschwindigkeit (km/h)": hourly.wind_speed,
                    "Beschreibung": hourly.description,
                    "Niederschlag (mm)": hourly.precipitation,
                    "Regenwahrscheinlichkeit (%)": hourly.chances_of_rain,
                    "Bewölkung (%)": hourly.cloud_cover,
                })
        return hourly_data
