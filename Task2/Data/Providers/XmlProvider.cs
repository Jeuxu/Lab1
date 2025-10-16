using Domain;
using System.Xml.Linq;

namespace Data.Providers
{
    public class XmlProvider : IProvider
    {
        public List<CityAirQuality> Read(string filePath)
        {
            try
            {
                var doc = XDocument.Load(filePath);
                return doc.Root!
                    .Elements("CityAirQuality")
                    .Select(city => new CityAirQuality
                    {
                        Rank = (int)city.Element("Rank")!,
                        CityCountry = (string)city.Element("CityCountry")!,
                        AverageAQI = (int)city.Element("AverageAQI")!,
                        MonthlyData = city.Element("MonthlyData")!
                            .Elements("MonthlyAQI")
                            .Select(m => new MonthlyAQI
                            {
                                Month = (string)m.Element("Month")!,
                                Value = (int)m.Element("Value")!
                            })
                            .ToList()
                    })
                    .ToList();
            }
            catch (Exception ex)
            {
                throw new Exception("Error reading file", ex);
            }
        }

        public void Write(string filePath, List<CityAirQuality> data)
        {
            var doc = new XDocument(
                new XElement("Cities",
                    data.Select(city =>
                        new XElement("CityAirQuality",
                            new XElement("Rank", city.Rank),
                            new XElement("CityCountry", city.CityCountry),
                            new XElement("AverageAQI", city.AverageAQI),
                            new XElement("MonthlyData",
                                city.MonthlyData.Select(m =>
                                    new XElement("MonthlyAQI",
                                        new XElement("Month", m.Month),
                                        new XElement("Value", m.Value)
                                    )
                                )
                            )
                        )
                    )
                )
            );

            doc.Save(filePath);
        }
    }
}