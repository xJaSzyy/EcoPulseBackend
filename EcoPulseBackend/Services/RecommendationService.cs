using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models.Recommendation;

namespace EcoPulseBackend.Services;

public class RecommendationService : IRecommendationService
{
    private readonly string _ollamaUrl;
    private const string ModelName = "llama3.2";
    
    public RecommendationService(IConfiguration configuration)
    {
        _ollamaUrl = configuration["OllamaUrl"] ?? "http://localhost:11434/api/chat";
    }

    public async Task<RecommendationResult> GetRecommendation(GetRecommendationModel model)
    {
        var result = new RecommendationResult();

        var systemPrompt = "Ты - помощник по рекомендациям в сфере экологии.";

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