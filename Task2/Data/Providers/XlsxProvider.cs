using Data.Providers;
using Domain;
using OfficeOpenXml;
using System.Diagnostics;
using System.Globalization;

public class XlsxProvider: IProvider
{

    private readonly string[] _months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

    public XlsxProvider()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Illia");
    }

    public List<CityAirQuality> Read(string filePath)
    {
        try
        {
            var result = new List<CityAirQuality>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int row = 2;
    
                while (worksheet.Cells[row, 1].Value != null)
                {
                    var cityAQ = new CityAirQuality
                    {
                        Rank = Convert.ToInt32(worksheet.Cells[row, 1].Value),
                        CityCountry = worksheet.Cells[row, 2].Text,
                        AverageAQI = Convert.ToInt32(worksheet.Cells[row, 3].Value)
                    };
    
                    for (int i = 0; i < _months.Length; i++)
                    {
                        var value = Convert.ToInt32(worksheet.Cells[row, 4 + i].Value);
                        cityAQ.MonthlyData.Add(new MonthlyAQI
                        {
                            Month = _months[i],
                            Value = value
                        });
                    }

                    result.Add(cityAQ);
                    row++;
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            throw new Exception("Error reading file", ex);
        }
    }

    public void Write(string filePath, List<CityAirQuality> data)
    {
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("AirQuality");


            worksheet.Cells[1, 1].Value = "Rank";
            worksheet.Cells[1, 2].Value = "City";
            worksheet.Cells[1, 3].Value = "Avg";


            for (int i = 0; i < _months.Length; i++)
            {
                worksheet.Cells[1, 4 + i].Value = _months[i];
            }

            for (int i = 0; i < data.Count; i++)
            {
                var row = i + 2;
                var city = data[i];

                worksheet.Cells[row, 1].Value = city.Rank;
                worksheet.Cells[row, 2].Value = city.CityCountry;
                worksheet.Cells[row, 3].Value = city.AverageAQI;

                for (int j = 0; j < _months.Length; j++)
                {
                    worksheet.Cells[row, 4 + j].Value = city.MonthlyData[j].Value;
                }
            }

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            package.SaveAs(new FileInfo(filePath));
        }
    }
}