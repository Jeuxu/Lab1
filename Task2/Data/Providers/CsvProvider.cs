using Data.Exceptions;
using Domain;
using Microsoft.VisualBasic.FileIO;

namespace Data.Providers
{
    public class CsvProvider: IProvider
    {
        public List<CityAirQuality> Read(string filePath)
        {
            var cities = new List<CityAirQuality>();
            if (!File.Exists(filePath)) return cities;

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            using (var parser = new TextFieldParser(filePath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;


                while (!parser.EndOfData)
                {
                    try
                    {
                        string[] parts = parser.ReadFields();

                        if (parts == null) continue;

                        if (parts.Length < 15) continue;

                        if (!int.TryParse(parts[0].Trim(), out int rank))
                        {
                            continue;
                        }

                        var cityField = parts[1].Trim();

                        int avg = 0;
                        int.TryParse(parts[2].Trim(), out avg);

                        var city = new CityAirQuality
                        {
                            Rank = rank,
                            CityCountry = cityField,
                            AverageAQI = avg,
                            MonthlyData = new List<MonthlyAQI>()
                        };

                        for (int m = 0; m < 12; m++)
                        {
                            int idx = 3 + m;
                            int value = 0;
                            if (idx < parts.Length)
                            {
                                var cell = parts[idx].Trim();
                                int.TryParse(cell, out value);
                            }

                            city.MonthlyData.Add(new MonthlyAQI
                            {
                                Month = months[m],
                                Value = value
                            });
                        }

                        cities.Add(city);
                    }
                    catch (FormatException fe)
                    {
                        throw new WrongProductFileFormatException("File format is incorrect", fe);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error reading file", ex);
                    }
                }
            }

            return cities;
        }

        public void Write(string filePath, List<CityAirQuality> data)
        {
            using (var writer = new StreamWriter(filePath))
            {
                writer.WriteLine("rank,city,avg,jan,feb,mar,apr,may,jun,jul,aug,sep,oct,nov,dec");

                foreach (var city in data)
                {
                    var row = new List<string>
                {
                    city.Rank.ToString(),
                    city.CityCountry,
                    city.AverageAQI.ToString()
                };

                    string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
                    foreach (var m in months)
                    {
                        var record = city.MonthlyData.FirstOrDefault(x => x.Month == m);
                        row.Add(record != null ? record.Value.ToString() : "0");
                    }

                    writer.WriteLine(string.Join(",", row));
                }
            }
        }
    }
}
