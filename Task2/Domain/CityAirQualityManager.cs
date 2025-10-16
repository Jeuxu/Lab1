using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Domain
{
    public class CityAirQualityManager
    {
        public List<CityAirQuality> cityAirQualities;
        
        public CityAirQualityManager() {
            cityAirQualities = new List<CityAirQuality>();
        }

        private void sortList()
        {
            foreach (var city in cityAirQualities)
            {
                if (city.MonthlyData != null && city.MonthlyData.Count > 0)
                {
                    city.AverageAQI = (int)Math.Round(city.MonthlyData.Average(m => m.Value));
                }
                else
                {
                    city.AverageAQI = 0;
                }
            }

            cityAirQualities = cityAirQualities
                .OrderByDescending(c => c.AverageAQI)
                .ToList();

            for (int i = 0; i < cityAirQualities.Count; i++)
            {
                cityAirQualities[i].Rank = i + 1;
            }
        }

        public void addCity(CityAirQuality city)
        {
            cityAirQualities.Add(city);
            sortList();
        }

        public void editCity(CityAirQuality city)
        {
            CityAirQuality cityForEdit = cityAirQualities.First(c => c.CityCountry == city.CityCountry);
            cityForEdit.AverageAQI = city.AverageAQI;
            cityForEdit.MonthlyData = city.MonthlyData;
            sortList();
        }

        public void removeCity(string cityCountryName) 
        {
            CityAirQuality cityForRemove = cityAirQualities.First(c => c.CityCountry == cityCountryName);
            cityAirQualities.Remove(cityForRemove);
            sortList();
        }
    }
}
