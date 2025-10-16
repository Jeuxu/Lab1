using Data;
using Data.Providers;
using Data.Reports;
using DocumentFormat.OpenXml.Bibliography;
using Domain;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Windows.Forms;
using System.Xml.Linq;
using UI;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Lab1._2
{
    public partial class MainForm : Form
    {
        DocxReportService docxReportService = new DocxReportService();
        CityAirQualityManager cityAirQualityManager = new CityAirQualityManager();
        private IProvider _m;
        XlsxReportService xlsxReportService = new XlsxReportService();
        private readonly string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
        private string _path = "";

        public MainForm()
        {
            InitializeComponent();
            LogMessage("Програма запущена");
        }

        private void LogMessage(string message)
        {
            if (textBoxLogs.InvokeRequired)
            {
                textBoxLogs.Invoke(new Action(() => LogMessage(message)));
            }
            else
            {
                textBoxLogs.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
                textBoxLogs.SelectionStart = textBoxLogs.Text.Length;
                textBoxLogs.ScrollToCaret();
            }
        }

        private void updateDataGridViewData(List<CityAirQuality> data)
        {
            LogMessage("Оновлення таблиці даних...");
            dataGridViewData.Columns.Clear();
            dataGridViewData.Rows.Clear();

            dataGridViewData.Columns.Add("Rank", "Rank");
            dataGridViewData.Columns.Add("CityCountry", "City / Country");
            dataGridViewData.Columns.Add("AverageAQI", "Average AQI");

            foreach (var month in months)
            {
                dataGridViewData.Columns.Add(month, month);
            }

            foreach (var city in data)
            {
                var row = new List<object>
                {
                    city.Rank,
                    city.CityCountry,
                    city.AverageAQI
                };

                foreach (var month in months)
                {
                    var record = city.MonthlyData.FirstOrDefault(m => m.Month == month);
                    row.Add(record != null ? record.Value : (object)"-");
                }

                dataGridViewData.Rows.Add(row.ToArray());
            }

            dataGridViewData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridViewData.AllowUserToAddRows = false;
            dataGridViewData.ReadOnly = false;

            LogMessage($"Таблиця оновлена: {data.Count} рядків");
        }

        private void buttonApply_Click(object sender, EventArgs e)
        {
            LogMessage("Користувач застосував фільтр");
            List<CityAirQuality> data = cityAirQualityManager.cityAirQualities;
            if (data == null || data.Count == 0)
            {
                LogMessage("Спроба застосувати фільтр без даних");
                return;
            }

            string field = comboBoxField.SelectedItem?.ToString();
            string condition = comboBoxCondition.SelectedItem?.ToString();
            string input = textBoxValue.Text.Trim();

            if (string.IsNullOrEmpty(field) || string.IsNullOrEmpty(input))
            {
                MessageBox.Show("Будь ласка, виберіть поле та введіть значення.");
                LogMessage("Спроба застосувати фільтр без введених даних");
                return;
            }

            bool isNumber = int.TryParse(input, out int numericValue);

            if (field == "City")
            {
                data = data.Where(c => c.CityCountry.Split(", ")[0] == input).ToList();
            }
            else if (field == "Country")
            {
                data = data.Where(c => c.CityCountry.Split(", ")[1] == input).ToList();
            }
            else
            {
                if (!isNumber)
                {
                    MessageBox.Show("Для цього поля потрібно ввести числове значення.");
                    LogMessage("Помилка: некоректне числове значення у фільтрі");
                    return;
                }
                if (string.IsNullOrEmpty(condition))
                {
                    MessageBox.Show("Не обрано умову!");
                    LogMessage("Помилка: необрано умову у фільтрі");
                    return;
                }
                switch (field)
                {
                    case "Rank":
                        switch (condition)
                        {
                            case ">":
                                data = data.Where(c => c.Rank > numericValue).ToList();
                                break;
                            case "<":
                                data = data.Where(c => c.Rank < numericValue).ToList();
                                break;
                            case "=":
                                data = data.Where(c => c.Rank == numericValue).ToList();
                                break;
                            case "!=":
                                data = data.Where(c => c.Rank != numericValue).ToList();
                                break;
                            case ">=":
                                data = data.Where(c => c.Rank >= numericValue).ToList();
                                break;
                            case "<=":
                                data = data.Where(c => c.Rank <= numericValue).ToList();
                                break;
                        }
                        break;
                    case "AverageAQI":
                        switch (condition)
                        {
                            case ">":
                                data = data.Where(c => c.AverageAQI > numericValue).ToList();
                                break;
                            case "<":
                                data = data.Where(c => c.AverageAQI < numericValue).ToList();
                                break;
                            case "=":
                                data = data.Where(c => c.AverageAQI == numericValue).ToList();
                                break;
                            case "!=":
                                data = data.Where(c => c.AverageAQI != numericValue).ToList();
                                break;
                            case ">=":
                                data = data.Where(c => c.AverageAQI >= numericValue).ToList();
                                break;
                            case "<=":
                                data = data.Where(c => c.AverageAQI <= numericValue).ToList();
                                break;
                        }
                        break;
                }
            }

            updateDataGridViewData(data.ToList());
            LogMessage($"Фільтр застосовано: поле={field}, умова={condition}, значення={input}. Рядків після фільтрації: {data.Count}");
        }

        private void comboBoxField_SelectedIndexChanged(object sender, EventArgs e)
        {
            string field = comboBoxField.SelectedItem?.ToString();
            comboBoxCondition.Enabled = field != "City" & field != "Country";
            LogMessage($"Користувач вибрав поле для фільтрації: {field}");
        }

        private void відкритиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LogMessage("Користувач відкрив діалог вибору файлу");
            var dialog = new OpenFileDialog();
            dialog.Filter = $"CSV files (*.csv)|*.csv|JSON files (*.json)|*.json|XML files (*.xml)|*.xml|XLSX files (*.xlsx)|*.xlsx";

            var result = dialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                _path = dialog.FileName;
                string extension = Path.GetExtension(_path).ToLower().TrimStart('.');

                setManager(extension);
                List<CityAirQuality> data = _m.Read(_path);

                if (data == null || data.Count == 0)
                {
                    MessageBox.Show("Файл пустий або дані не завантажені.");
                    LogMessage("Спроба показати підсумок для пустих даних");
                    return;
                }
                ImportPreview importPreview = new ImportPreview(data);
                DialogResult dialog1 = importPreview.ShowDialog();
                if (dialog1 == DialogResult.OK)
                {
                    cityAirQualityManager.cityAirQualities = data;
                    updateDataGridViewData(cityAirQualityManager.cityAirQualities);
                    LogMessage($"Файл завантажено: {_path} ({extension.ToUpper()})");
                }
                else
                {
                    LogMessage($"Файл не було завантажено: {_path} ({extension.ToUpper()})");
                }
            }
            else
            {
                LogMessage("Вибір файлу скасовано користувачем");
            }
        }

        private void вихідToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LogMessage("Користувач намагається вийти з програми");
            var result = MessageBox.Show("Ви впевнені, що хочете вийти?", "Підтвердження", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                LogMessage("Програма завершила роботу");
                Application.Exit();
            }
            else
            {
                LogMessage("Вихід з програми скасовано користувачем");
            }
        }

        private void buttonToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var extension = (sender as ToolStripMenuItem).Tag.ToString();
            LogMessage($"Користувач зберігає файл у форматі {extension.ToUpper()}");

            var dialog = new SaveFileDialog();
            dialog.Filter = $"{char.ToUpper(extension[0]) + extension.Substring(1).ToLower()} files (*.{extension})|*.{extension}";

            var result = dialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                string path = dialog.FileName;
                setManager(extension);
                _m.Write(path, cityAirQualityManager.cityAirQualities);
                LogMessage($"Файл збережено: {path} ({extension.ToUpper()})");
            }
            else
            {
                LogMessage($"Збереження файлу у форматі {extension.ToUpper()} скасовано користувачем");
            }
        }

        private void setManager(string extension)
        {
            switch (extension)
            {
                case "csv":
                    _m = new CsvProvider();
                    break;
                case "json":
                    _m = new JsonProvider();
                    break;
                case "xml":
                    _m = new XmlProvider();
                    break;
                case "xlsx":
                    _m = new XlsxProvider();
                    break;
            }
            LogMessage($"Менеджер даних встановлено: {extension.ToUpper()}");
        }

        private void editCity(object sender, DataGridViewCellEventArgs e)
        {
            LogMessage("Користувач зберігає зміни у файлі");

            foreach (DataGridViewRow row in dataGridViewData.Rows)
            {
                if (row.IsNewRow) continue;
                try
                {
                    var city = new CityAirQuality
                    {
                        Rank = int.TryParse(row.Cells["Rank"].Value?.ToString(), out int rank) ? rank : 0,
                        CityCountry = $"{row.Cells["CityCountry"].Value?.ToString()}",
                        AverageAQI = int.TryParse(row.Cells["AverageAQI"].Value?.ToString(), out int avg) ? avg : 0
                    };

                    foreach (var m in months)
                    {
                        if (dataGridViewData.Columns.Contains(m))
                        {
                            int val = int.TryParse(row.Cells[m].Value?.ToString(), out int v) ? v : 0;
                            city.MonthlyData.Add(new MonthlyAQI { Month = m, Value = val });
                        }
                    }

                    cityAirQualityManager.editCity(city);
                }
                catch (Exception ex)
                {
                    LogMessage("Помилка при редагуванні файлу: " + ex.Message);
                    throw new Exception($"Error reading data from table: {ex.Message}");
                }
            }
        }


        private void buttonApplyChart_Click(object sender, EventArgs e)
        {
            LogMessage("Користувач будує діаграму");
            int chartType = cmbChartType.SelectedIndex;
            string searchedText = textBoxCityCountry.Text?.Trim();
            int chartVolume = comboBoxVolume.SelectedIndex;
            if (chartType == -1)
            {
                MessageBox.Show("Виберіть тип діаграми.");
                LogMessage("Помилка: не вибрано тип діаграми");
                return;
            }

            PlotModel model = null;

            var city = cityAirQualityManager.cityAirQualities.FirstOrDefault(c => c.CityCountry == searchedText);

            switch (chartType)
            {
                case 0:
                    if (city == null)
                    {
                        MessageBox.Show($"Місто '{searchedText}' не знайдено.");
                        LogMessage("Помилка: місто не знайдено");
                        return;
                    }
                    model = CreateLineChart(city);
                    break;
                case 1:
                    if (city == null)
                    {
                        MessageBox.Show($"Місто '{searchedText}' не знайдено.");
                        LogMessage("Помилка: місто не знайдено");
                        return;
                    }
                    model = CreateBarChart(city);
                    break;
                case 2:
                    if (chartVolume == -1)
                    {
                        MessageBox.Show("Виберіть об'єм діаграми.");
                        LogMessage("Помилка: не вибрано об'єм діаграми");
                        return;
                    }
                    if (chartVolume == 0)
                    {
                        model = CreatePieChart("");
                    }
                    if (chartVolume == 1)
                    {
                        model = CreatePieChart(searchedText);
                    }
                    break;
            }

            plotView1.Model = model;
            LogMessage($"Діаграму побудовано. Тип={chartType}");
        }

        private PlotModel CreateLineChart(CityAirQuality city)
        {
            LogMessage($"Створення лінійної діаграми для {city.CityCountry}");
            var model = new PlotModel();

            var categoryAxis = new CategoryAxis { Position = AxisPosition.Bottom };
            categoryAxis.Labels.AddRange(city.MonthlyData.Select(m => m.Month));
            model.Axes.Add(categoryAxis);

            model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "AQI" });

            var series = new LineSeries { MarkerType = MarkerType.Circle };
            foreach (var m in city.MonthlyData)
                series.Points.Add(new DataPoint(categoryAxis.Labels.IndexOf(m.Month), m.Value));

            model.Series.Add(series);
            return model;
        }

        private PlotModel CreateBarChart(CityAirQuality city)
        {
            LogMessage($"Створення стовпчикової діаграми для {city.CityCountry}");
            var model = new PlotModel();

            var categoryAxis = new CategoryAxis { Position = AxisPosition.Left };
            categoryAxis.Labels.AddRange(city.MonthlyData.Select(m => m.Month));
            model.Axes.Add(categoryAxis);

            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "AQI" });

            var series = new BarSeries
            {
                LabelPlacement = LabelPlacement.Inside,
                LabelFormatString = "{0}"
            };

            foreach (var m in city.MonthlyData)
            {
                series.Items.Add(new BarItem { Value = m.Value });
            }

            model.Series.Add(series);
            return model;
        }

        private PlotModel CreatePieChart(string city)
        {
            LogMessage("Створення кругової діаграми");
            var model = new PlotModel
            {
                Background = OxyColors.White,             
                PlotAreaBorderColor = OxyColors.Gray     
            };

            var pieSeries = new PieSeries
            {
                StrokeThickness = 0.5,
                InsideLabelPosition = 0.5,
                AngleSpan = 360,
                StartAngle = 0,   
                OutsideLabelFormat = "{1}",
                InsideLabelFormat = "{2:0.0}%"
            };
            Dictionary<string, double> AirQualities;

            if (city == "")
            {
                AirQualities = countAverageAQIForCountries(cityAirQualityManager.cityAirQualities);
            }
            else
            {
                AirQualities = countAverageAQIForCountry(cityAirQualityManager.cityAirQualities, city);
            }


            foreach (var keyValue in AirQualities)
            {
                pieSeries.Slices.Add(new PieSlice(keyValue.Key, keyValue.Value) { Fill = OxyColor.FromRgb(255, 200, 100) });
            }

            model.Series.Add(pieSeries);
            return model;
        }

        private Dictionary<string, double> countAverageAQIForCountries(List<CityAirQuality> data)
        {
            Dictionary<string, double> countriesAirQualities = new Dictionary<string, double>();
            Dictionary<string, int> countriesCounts = new Dictionary<string, int>();

            foreach (var cityAirQuality in data)
            {
                string[] parts = cityAirQuality.CityCountry.Split(", ");
                string country = parts.Length > 1 ? parts[1].Trim() : parts[0].Trim();


                if (countriesAirQualities.ContainsKey(country))
                {
                    countriesAirQualities[country] += cityAirQuality.AverageAQI;
                    countriesCounts[country]++;
                }
                else
                {
                    countriesAirQualities[country] = cityAirQuality.AverageAQI;
                    countriesCounts[country] = 1;
                }
            }

            foreach (var country in countriesAirQualities.Keys.ToList())
            {
                countriesAirQualities[country] /= countriesCounts[country];
            }
            return countriesAirQualities;
        }

        private Dictionary<string, double> countAverageAQIForCountry(List<CityAirQuality> data, string countryName)
        {
            Dictionary<string, double> countryAirQualities = new Dictionary<string, double>();
            Dictionary<string, int> countryCounts = new Dictionary<string, int>();

            foreach (var cityAirQuality in data)
            {
                string[] parts = cityAirQuality.CityCountry.Split(", ");
                string country = parts[1].Trim();
                string city = parts[0].Trim();

                if (country == countryName & countryAirQualities.ContainsKey(country))
                {
                    countryAirQualities[city] += cityAirQuality.AverageAQI;
                    countryCounts[city]++;
                }
                else if (country == countryName)
                {
                    countryAirQualities[city] = cityAirQuality.AverageAQI;
                    countryCounts[city] = 1;
                }
            }

            foreach (var city in countryAirQualities.Keys.ToList())
            {
                countryAirQualities[city] /= countryCounts[city];
            }
            return countryAirQualities;
        }

        private void cmbChartType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbChartType.SelectedIndex == 2)
            {
                textBoxCityCountry.Enabled = false;
                comboBoxVolume.Enabled = true;
            }
            else
            {
                textBoxCityCountry.Enabled = true;
                textBoxCityCountry.PlaceholderText = "Місто, Країна";
                comboBoxVolume.Enabled = false;
            }
            LogMessage($"Вибрано тип діаграми: {cmbChartType.SelectedIndex}");
        }


        private void comboBoxVolume_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxVolume.SelectedIndex == 1)
            {
                textBoxCityCountry.PlaceholderText = "Країна";
                textBoxCityCountry.Enabled = true;
            }
            else
            {
                textBoxCityCountry.Enabled = false;
            }
        }
        private void buttonChartExport_Click(object sender, EventArgs e)
        {
            LogMessage("Користувач експортує діаграму");
            if (plotView1.Model == null)
            {
                MessageBox.Show("Немає діаграми для експорту.");
                LogMessage("Помилка: спроба експорту без діаграми");
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "PNG Image (*.png)|*.png";
                dialog.Title = "Зберегти діаграму як PNG";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = dialog.FileName;
                        plotView1.Model.Background = OxyColors.White;
                        var pngExporter = new PngExporter { Width = 800, Height = 600 };
                        pngExporter.ExportToFile(plotView1.Model, filePath);

                        MessageBox.Show($"Діаграму збережено: {filePath}");
                        LogMessage($"Діаграму експортовано у PNG: {filePath}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка при збереженні: " + ex.Message);
                        LogMessage("Помилка експорту діаграми: " + ex.Message);
                    }
                }
                else
                {
                    LogMessage("Експорт діаграми скасовано користувачем");
                }
            }
        }

        private void згенеруватиВToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (cityAirQualityManager.cityAirQualities == null || cityAirQualityManager.cityAirQualities.Count == 0)
            {
                MessageBox.Show("Дані відсутні. Завантажте файл перед генерацією звіту.");
                LogMessage("Спроба згенерувати XLSX-звіт без даних");
                return;
            }

            LogMessage("Користувач почав генерацію XLSX-звіту");

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Excel файли (*.xlsx)|*.xlsx";
                dialog.Title = "Зберегти XLSX звіт";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = dialog.FileName;

                        plotView1.Model = CreatePieChart("");
                        string chartPath = $"../../../../Data/Outputs/Chart_{DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")}.png";
                        plotView1.Model.Background = OxyColors.White;
                        var pngExporter = new PngExporter { Width = 800, Height = 600 };
                        pngExporter.ExportToFile(plotView1.Model, chartPath);
                        plotView1.Model = null;

                        xlsxReportService.GenerateReport(
                            cityAirQualityManager.cityAirQualities,
                            filePath,
                            chartPath
                        );

                        MessageBox.Show($"XLSX-звіт успішно збережено: {filePath}");
                        LogMessage($"XLSX звіт згенеровано: {filePath}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка при генерації XLSX звіту: " + ex.Message);
                        LogMessage("Помилка генерації XLSX звіту: " + ex.Message);
                    }
                }
                else
                {
                    LogMessage("Генерацію XLSX-звіту скасовано користувачем");
                }
            }
        }

        private void згенеруватиDOCXзвітToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (cityAirQualityManager.cityAirQualities == null || cityAirQualityManager.cityAirQualities.Count == 0)
            {
                MessageBox.Show("Дані відсутні. Завантажте файл перед генерацією звіту.");
                LogMessage("Спроба згенерувати DOCX-звіт без даних");
                return;
            }

            LogMessage("Користувач почав генерацію DOCX-звіту");

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Word документи (*.docx)|*.docx";
                dialog.Title = "Зберегти DOCX звіт";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = dialog.FileName;

                        
                        plotView1.Model = CreatePieChart("");
                        string chartPath = $"../../../../Data/Outputs/Chart_{DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")}.png";
                        plotView1.Model.Background = OxyColors.White;
                        var pngExporter = new PngExporter { Width = 800, Height = 600 };
                        pngExporter.ExportToFile(plotView1.Model, chartPath);
                        plotView1.Model = null;


                        docxReportService.GenerateReport(
                            filePath,
                            cityAirQualityManager.cityAirQualities,
                            chartPath
                        );

                        MessageBox.Show($"DOCX-звіт успішно збережено: {filePath}");
                        LogMessage($"DOCX звіт згенеровано: {filePath}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка при генерації DOCX звіту: " + ex.Message);
                        LogMessage("Помилка генерації DOCX звіту: " + ex.Message);
                    }
                }
                else
                {
                    LogMessage("Генерацію DOCX-звіту скасовано користувачем");
                }
            }
        }

        private void buttonRemove_Click(object sender, EventArgs e)
        {
            LogMessage("Користувач намагається видалити місто");
            if (dataGridViewData.SelectedRows.Count == 0)
            {
                MessageBox.Show("Оберіть місто для видалення!");
                LogMessage("Помилка: спроба видалити місто без вибору");
                return;
            }
            string cityCountry = dataGridViewData.SelectedRows[0].Cells[1].Value.ToString();
            cityAirQualityManager.removeCity(cityCountry);
            updateDataGridViewData(cityAirQualityManager.cityAirQualities);
            MessageBox.Show("Місто було видалено!");
            LogMessage($"{cityCountry} видалено");
        }

        private void buttonLogExport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBoxLogs.Text))
            {
                MessageBox.Show("Логи порожні, нічого експортувати.");
                LogMessage("Спроба експортувати пусті логи");
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Text files (*.txt)|*.txt";
                dialog.Title = "Зберегти логи у TXT";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = dialog.FileName;
                        var textProvider = new TextProvider();
                        textProvider.Write(filePath, textBoxLogs.Text);

                        MessageBox.Show($"Логи успішно збережено: {filePath}");
                        LogMessage($"Логи експортовано у файл: {filePath}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка при експорті логів: " + ex.Message);
                        LogMessage("Помилка експорту логів: " + ex.Message);
                    }
                }
                else
                {
                    LogMessage("Експорт логів скасовано користувачем");
                }
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            AddCityForm addCityForm = new AddCityForm();
            addCityForm.ShowDialog();
            LogMessage("Відкриття форми для додавання міста");
            if (addCityForm.DialogResult == DialogResult.OK)
            {
                CityAirQuality newCity = new CityAirQuality
                {
                    CityCountry = addCityForm.CityName,
                    MonthlyData = addCityForm.MonthlyData,
                    AverageAQI = (int)Math.Round(addCityForm.MonthlyData.Average(m => m.Value))
                };

                cityAirQualityManager.addCity(newCity);
                updateDataGridViewData(cityAirQualityManager.cityAirQualities);
                LogMessage("Додавання нового міста.");
            }
            else
            {
                LogMessage("Додавання нового міста відхилено.");
            }
        }
    }
}

