using EcoPulseBackend.Models;

namespace EcoPulseBackend.Extensions;

public static class GeoUtils
{
    private const double EarthRadiusMeters = 6371000; 

    public static double DistanceMeters(Coordinates a, Coordinates b)
    {
        var lat1 = DegreesToRadians(a.Lat);
        var lon1 = DegreesToRadians(a.Lon);
        var lat2 = DegreesToRadians(b.Lat);
        var lon2 = DegreesToRadians(b.Lon);

        var dLat = lat2 - lat1;
        var dLon = lon2 - lon1;

        var sinLat = Math.Sin(dLat / 2);
        var sinLon = Math.Sin(dLon / 2);

        var h = sinLat * sinLat +
                Math.Cos(lat1) * Math.Cos(lat2) *
                sinLon * sinLon;

        var c = 2 * Math.Atan2(Math.Sqrt(h), Math.Sqrt(1 - h));

        return EarthRadiusMeters * c;
    }

    private static double DegreesToRadians(double deg) => deg * Math.PI / 180.0;
}