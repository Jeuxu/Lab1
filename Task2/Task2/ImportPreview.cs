using Domain;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UI
{
    public partial class ImportPreview : Form
    {
        public ImportPreview(List<CityAirQuality> data)
        {
            InitializeComponent();

            int rowCount = data.Count;
            int colCount = 3 + 12;
            bool hasMissing = data.Any(c =>
                string.IsNullOrEmpty(c.CityCountry) ||
                c.MonthlyData.Any(m => m.Value == null));

            string summary = $"Рядків: {rowCount}\r\n" +
                             $"Стовпців: {colCount}\r\n" +
                             $"Типи полів: Rank(int), CityCountry(string), AverageAQI(int), MonthlyData(int)\r\n" +
                             $"Пропуски: {(hasMissing ? "є" : "немає")}";

            var preview = string.Join("\r\n", data.Take(5).Select(c =>
                $"{c.Rank}, {c.CityCountry}, {c.AverageAQI}, " +
                string.Join(", ", c.MonthlyData.Select(m => m.Value.ToString() ?? "-"))
            ));
            textBox1.Text = summary + "\r\n\r\nПопередній перегляд (перші " + 5 + " рядків):\r\n" + preview;
        }
    }
}
