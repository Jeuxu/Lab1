namespace Domain
{
    public class CityAirQuality
    {
        public int Rank { get; set; }
        public string CityCountry { get; set; } = "";
        public int AverageAQI { get; set; }  

        public List<MonthlyAQI> MonthlyData { get; set; } = new List<MonthlyAQI>();
    }
}
