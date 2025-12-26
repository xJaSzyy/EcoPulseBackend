using System.Net;
using System.Security.Authentication;
using System.Text.Json;
using EcoPulseBackend.Models.Weather;
using Microsoft.AspNetCore.Mvc;

namespace EcoPulseBackend.Controllers;

public class WeatherController : ControllerBase
{
    private readonly IHttpClientFactory _httpClientFactory;

    /// <summary>
    /// Конструктор
    /// </summary>
    /// <param name="httpClientFactory">IHttpClientFactory</param>
    public WeatherController(IHttpClientFactory httpClientFactory)
    {
        _httpClientFactory = httpClientFactory;
    }

    /// <summary>
    /// Метод получения текущей погоды
    /// </summary>
    /// <param name="city">Название города</param>
    /// <returns></returns>
    [HttpGet("weather/current")]
    [ProducesResponseType(typeof(WeatherViewModel), (int)HttpStatusCode.OK)]
    [ProducesResponseType(typeof(string), (int)HttpStatusCode.InternalServerError)]
    public async Task<IActionResult> GetCurrentWeather([FromQuery] string city)
    {
        var httpClient = _httpClientFactory.CreateClient("weather");

        var weatherUrl =
            "https://api.open-meteo.com/v1/forecast?latitude=55.355198&longitude=86.086847&current_weather=true";

        try
        {
            var response = await httpClient.GetAsync(weatherUrl);

            if (!response.IsSuccessStatusCode)
            {
                return StatusCode((int)response.StatusCode, "Ошибка при запросе к API погоды.");
            }

            var responseContent = await response.Content.ReadAsStringAsync();
            var weatherResponse = JsonSerializer.Deserialize<OpenMeteoResponse>(responseContent);

            var result = new WeatherViewModel
            {
                Date = DateTime.UtcNow.Date,
                Temperature = (float)weatherResponse!.CurrentWeather.Temperature,
                WindSpeed = (float)weatherResponse.CurrentWeather.WindSpeed,
                WindDirection = (int)(weatherResponse.CurrentWeather.WindDirection + 180) % 360,
                IconClass = GetWeatherInfo(weatherResponse.CurrentWeather.WeatherCode,
                    weatherResponse.CurrentWeather.IsDay == 1).IconClass
            };

            return Ok(result);
        }
        catch (Exception ex)
        {
            return Ok(ex.ToString());
        }
    }

    private (string Description, string IconClass) GetWeatherInfo(int weatherCode, bool isDay)
    {
        return weatherCode switch
        {
            0 => isDay
                ? ("Ясное небо", "wi-day-sunny")
                : ("Ясная ночь", "wi-night-clear"),

            1 => isDay
                ? ("Преимущественно ясно", "wi-day-cloudy")
                : ("Преимущественно ясно", "wi-night-alt-cloudy"),

            2 => ("Переменная облачность", "wi-cloudy"),

            3 => ("Пасмурно", "wi-cloudy"),

            45 or 48 => ("Туман", "wi-fog"),

            51 or 53 or 55 => ("Морось", "wi-sprinkle"),

            56 or 57 => ("Ледяная морось", "wi-rain-mix"),

            61 or 63 or 65 => ("Дождь", "wi-rain"),

            66 or 67 => ("Ледяной дождь", "wi-rain-mix"),

            71 or 73 or 75 => ("Снег", "wi-snow"),

            77 => ("Снежные зёрна", "wi-snow"),

            80 or 81 or 82 => ("Ливень", "wi-showers"),

            85 or 86 => ("Снегопад", "wi-snow-wind"),

            95 => ("Гроза", "wi-thunderstorm"),

            96 or 99 => ("Гроза с градом", "wi-storm-showers"),

            _ => ("Неизвестно", "wi-na")
        };
    }
}