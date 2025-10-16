using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using Domain;

namespace Data.Providers
{
    public class CsvProvider : IProvider
    {
        private readonly string[] _months =
        {
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        };

        public List<CityAirQuality> Read(string filePath)
        {
            var cities = new List<CityAirQuality>();
            if (!File.Exists(filePath))
                return cities;

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                Delimiter = ",",
                IgnoreBlankLines = true,
                TrimOptions = TrimOptions.Trim,
                BadDataFound = null
            };

            try
            {
                using var reader = new StreamReader(filePath);
                using var csv = new CsvReader(reader, config);


                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                        int rank = csv.GetField<int>("rank");
                        string cityCountry = csv.GetField<string>("city");
                        int avg = csv.GetField<int>("avg");

                        var city = new CityAirQuality
                        {
                            Rank = rank,
                            CityCountry = cityCountry,
                            AverageAQI = avg,
                            MonthlyData = new List<MonthlyAQI>()
                        };

                        foreach (var m in _months)
                        {
                            int val = 0;
                            csv.TryGetField<int>(m.ToLower(), out val);
                            city.MonthlyData.Add(new MonthlyAQI { Month = m, Value = val });
                        }

                        cities.Add(city);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error reading: ", ex);
            }

            return cities;
        }

        public void Write(string filePath, List<CityAirQuality> data)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                Delimiter = ",",
            };

            using var writer = new StreamWriter(filePath);
            using var csv = new CsvWriter(writer, config);

            csv.WriteField("rank");
            csv.WriteField("city");
            csv.WriteField("avg");
            foreach (var m in _months)
                csv.WriteField(m.ToLower());
            csv.NextRecord();

            foreach (var city in data)
            {
                csv.WriteField(city.Rank);
                csv.WriteField(city.CityCountry);
                csv.WriteField(city.AverageAQI);

                foreach (var m in _months)
                {
                    var record = city.MonthlyData.FirstOrDefault(x => x.Month == m);
                    csv.WriteField(record?.Value ?? 0);
                }

                csv.NextRecord();
            }
        }
    }
}