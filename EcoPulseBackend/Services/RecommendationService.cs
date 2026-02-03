using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models.Recommendation;

namespace EcoPulseBackend.Services;

public class RecommendationService : IRecommendationService
{
    private readonly string _ollamaUrl;
    private const string ModelName = "gpt-oss:120b-cloud";
    
    public RecommendationService(IConfiguration configuration)
    {
        _ollamaUrl = configuration["OllamaUrl"] ?? "http://localhost:11434/api/chat";
    }

    public async Task<RecommendationResult> GetRecommendation(GetRecommendationModel model)
    {
        var result = new RecommendationResult();

        var systemPrompt = "Ты - специалист по рекомендациям в сфере экологии. " +
                           "Твоя задача четко и кратко формулировать советы и рекомендации для жителей города. " +
                           "Ты должен выделить 3 рекомендации, не больше, не меньше. " +
                           "Эти рекомендации должны быть сформулированы четко и ясно, максимально кратко, но со смыслом." +
                           "Советы должны быть для обычных жителей города." +
                           "Никаких форматирований (без жирного текста)" +
                           "Как могут выглядеть рекомендации: " +
                           "1. Желательно сократить время пребывания на улице" +
                           "2. По возможности, не открывайте сегодня окна" +
                           "3. Больше выходите на прогулки - воздух чистый" +
                           "4. Наслаждайтесь активным отдыхом на улице без каких-либо опасений за здоровье." +
                           "5. Качество воздуха приемлемо, но чувствительные люди могут испытывать легкий дискомфорт." +
                           "6. Ограничьте активный отдых на улице." +
                           "7. Избегайте активного отдыха на улице." +
                           "8. Пользуйтесь общественным транспортом вместо личных автомобилей, чтобы уменьшить общие выбросы в атмосферу.";

        using var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var messages = new List<object>
        {
            new { Role = "system", Content = systemPrompt.Trim() },
            new { Role = "user", Content = model.Context.Trim() }
        };

        try
        {
            var requestBody = new
            {
                Model = ModelName,
                Messages = messages,
                Stream = false
            };

            var json = JsonSerializer.Serialize(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await httpClient.PostAsync(_ollamaUrl, content);
            response.EnsureSuccessStatusCode();

            var responseJson = await response.Content.ReadAsStringAsync();
            var chatResponse = JsonSerializer.Deserialize<JsonElement>(responseJson);

            if (chatResponse.TryGetProperty("message", out var message) &&
                message.TryGetProperty("content", out var contentElement))
            {
                var assistantReply = contentElement.GetString() ?? "";

                result = new RecommendationResult  
                {
                    Context = model.Context,
                    Recommendation = assistantReply
                };
                
                return result;  
            }
            else
            {
                result.Recommendation = "Ошибка: неожиданный формат ответа от сервера.";
                return result;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка в GetRecommendation: {ex}");
            result.Recommendation = $"Ошибка: {ex.Message}";
            return result;
        }
    }
}