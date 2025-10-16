using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Domain;

namespace Data.Providers
{
    public interface IProvider
    {
        List<CityAirQuality> Read(string filePath);
        void Write(string filePath, List<CityAirQuality> data);
    }
}
