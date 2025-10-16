using Domain;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Data.Providers
{
    public class JsonProvider: IProvider
    {
        public List<CityAirQuality> Read(string filePath)
        {
            try
            {
                string json = File.ReadAllText(filePath);
                var data = JsonSerializer.Deserialize<List<CityAirQuality>>(json);
                if (data != null)
                {
                    return data;
                }
                else
                {
                    return new List<CityAirQuality>();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error reading file", ex);
            }
        }
        public void Write(string filePath, List<CityAirQuality> data)
        {
            string json = JsonSerializer.Serialize(data.ToArray());
            File.WriteAllText(filePath, json);
        }
    }
}
