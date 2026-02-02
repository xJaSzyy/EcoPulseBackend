using EcoPulseBackend.Models.Recommendation;

namespace EcoPulseBackend.Interfaces;

public interface IRecommendationService
{
    Task<RecommendationResult> GetRecommendation(GetRecommendationModel model);
}