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

        public void removeCity(string cityCountryName) 
        {
            CityAirQuality cityForRemove = cityAirQualities.First(c => c.CityCountry == cityCountryName);
            cityAirQualities.Remove(cityForRemove);
        }
    }
}
