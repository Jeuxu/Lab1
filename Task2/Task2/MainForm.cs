using Data;
using Data.Exceptions;
using Data.Providers;
using Domain;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using System.IO;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Data.Reports;

namespace Lab1._2
{
    public partial class MainForm : Form
    {
        CityAirQualityManager cityAirQualityManager = new CityAirQualityManager();
        private IProvider _m = new CsvProvider();
        XlsxReportService xlsxReportService = new XlsxReportService();
        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
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

        private void ShowFileSummary(List<CityAirQuality> data, int previewRows = 5)
        {
            if (data == null || data.Count == 0)
            {
                MessageBox.Show("Файл пустий або дані не завантажені.");
                LogMessage("Спроба показати підсумок для пустих даних");
                return;
            }

            int rowCount = data.Count;
            int colCount = 3 + 12;
            bool hasMissing = data.Any(c =>
                string.IsNullOrEmpty(c.CityCountry) ||
                c.MonthlyData.Any(m => m.Value == null));

            string summary = $"Рядків: {rowCount}\n" +
                             $"Стовпців: {colCount}\n" +
                             $"Типи полів: Rank(int), CityCountry(string), AverageAQI(int), MonthlyData(int?)\n" +
                             $"Пропуски: {(hasMissing ? "є" : "немає")}";

            var preview = string.Join("\n", data.Take(previewRows).Select(c =>
                $"{c.Rank}, {c.CityCountry}, {c.AverageAQI}, " +
                string.Join(", ", c.MonthlyData.Select(m => m.Value.ToString() ?? "-"))
            ));

            MessageBox.Show(summary + "\n\nПопередній перегляд (перші " + previewRows + " рядків):\n" + preview);
            LogMessage($"Показано підсумок: {rowCount} рядків, {colCount} стовпців. {(hasMissing ? "Є пропуски" : "Без пропусків")}");
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
                switch (field)
                {
                    case "Rank":
                        data = applyCondition(condition, data, numericValue, c => c.Rank);
                        break;
                    case "AverageAQI":
                        data = applyCondition(condition, data, numericValue, c => c.AverageAQI);
                        break;
                }
            }

            updateDataGridViewData(data.ToList());
            LogMessage($"Фільтр застосовано: поле={field}, умова={condition}, значення={input}. Рядків після фільтрації: {data.Count}");
        }

        private List<CityAirQuality> applyCondition(string condition, List<CityAirQuality> data, int value, Func<CityAirQuality, int> selector)
        {
            return condition switch
            {
                ">" => data.Where(c => selector(c) > value).ToList(),
                "<" => data.Where(c => selector(c) < value).ToList(),
                "=" => data.Where(c => selector(c) == value).ToList(),
                "!=" => data.Where(c => selector(c) != value).ToList(),
                ">=" => data.Where(c => selector(c) >= value).ToList(),
                "<=" => data.Where(c => selector(c) <= value).ToList(),
                _ => data,
            };
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
                cityAirQualityManager.cityAirQualities = _m.Read(_path);
                updateDataGridViewData(cityAirQualityManager.cityAirQualities);
                ShowFileSummary(cityAirQualityManager.cityAirQualities);
                buttonEdit.Enabled = true;
                MessageBox.Show($"File loaded: {extension.ToUpper()}");
                LogMessage($"Файл завантажено: {_path} ({extension.ToUpper()})");
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

        private void buttonEdit_Click(object sender, EventArgs e)
        {
            LogMessage("Користувач зберігає зміни у файлі");
            try
            {
                List<CityAirQuality> cities = ReadAirQualityFromGrid();
                string extension = Path.GetExtension(_path).ToLower().TrimStart('.');
                setManager(extension);
                cityAirQualityManager.cityAirQualities = cities;
                _m.Write(_path, cityAirQualityManager.cityAirQualities);
                LogMessage($"Файл оновлено: {_path}");
            }
            catch (FormatException fe)
            {
                LogMessage("Помилка: некоректний формат файлу при редагуванні");
                throw new WrongProductFileFormatException("File format is incorrect", fe);
            }
            catch (Exception ex)
            {
                LogMessage("Помилка при редагуванні файлу: " + ex.Message);
                throw new Exception("Error reading file", ex);
            }
        }

        public List<CityAirQuality> ReadAirQualityFromGrid()
        {
            LogMessage("Читання даних з таблиці");
            var result = new List<CityAirQuality>();

            foreach (DataGridViewRow row in dataGridViewData.Rows)
            {
                if (row.IsNewRow) continue;

                var city = new CityAirQuality
                {
                    Rank = int.TryParse(row.Cells["Rank"].Value?.ToString(), out int rank) ? rank : 0,
                    CityCountry = $"\"{row.Cells["CityCountry"].Value?.ToString()}\"",
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

                result.Add(city);
            }

            LogMessage($"З таблиці прочитано {result.Count} рядків");
            return result;
        }

        private void buttonApplyChart_Click(object sender, EventArgs e)
        {
            LogMessage("Користувач будує графік");
            int chartType = cmbChartType.SelectedIndex;
            string searchedText = textBoxCityCountry.Text?.Trim();

            if (string.IsNullOrEmpty(searchedText))
            {
                MessageBox.Show("Введіть назву міста та країни або лише країни для кругового графіка.");
                LogMessage("Помилка: не введено назву міста/країни для графіка");
                return;
            }
            if (chartType == -1)
            {
                MessageBox.Show("Виберіть ти графіка.");
                LogMessage("Помилка: не вибрано тип графіка");
                return;
            }

            PlotModel model = null;

            switch (chartType)
            {
                case 0:
                    var cityLine = cityAirQualityManager.cityAirQualities.FirstOrDefault(c => c.CityCountry == searchedText);
                    if (cityLine == null)
                    {
                        MessageBox.Show($"Місто, Країну '{searchedText}' не знайдено.");
                        LogMessage("Помилка: місто не знайдено");
                        break;
                    }
                    model = CreateLineChart(cityLine);
                    break;

                case 1:
                    var cityBar = cityAirQualityManager.cityAirQualities.FirstOrDefault(c => c.CityCountry == searchedText);
                    if (cityBar == null)
                    {
                        MessageBox.Show($"Місто, Країну '{searchedText}' не знайдено.");
                        LogMessage("Помилка: місто не знайдено");
                        break;
                    }
                    model = CreateBarChart(cityBar);
                    break;

                case 2:
                    List<CityAirQuality> data = cityAirQualityManager.cityAirQualities.Where(c => c.CityCountry.Split(", ")[1] == searchedText).ToList();
                    if (data.Count() == 0)
                    {
                        MessageBox.Show($"Країну '{searchedText}' не знайдено.");
                        LogMessage("Помилка: країну не знайдено");
                        break;
                    }
                    model = CreatePieChart(data);
                    break;
            }

            plotView1.Model = model;
            LogMessage($"Графік побудовано. Тип={chartType}, параметр={searchedText}");
        }

        private PlotModel CreateLineChart(CityAirQuality city)
        {
            LogMessage($"Створення лінійного графіка для {city.CityCountry}");
            var model = new PlotModel { Title = $"Monthly AQI - {city.CityCountry}" };

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
            LogMessage($"Створення стовпчикового графіка для {city.CityCountry}");
            var model = new PlotModel { Title = $"Monthly AQI (Bar horizontal) - {city.CityCountry}" };

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

        private PlotModel CreatePieChart(List<CityAirQuality> data)
        {
            LogMessage("Створення кругової діаграми для країни");
            var model = new PlotModel { Title = "AQI Share by City" };

            var pieSeries = new PieSeries
            {
                StrokeThickness = 0.5,
                InsideLabelPosition = 0.5,
                AngleSpan = 360,
                StartAngle = 0,
                OutsideLabelFormat = "{0}: {1} ({2:0.0}%)",
                InsideLabelFormat = ""
            };
            double total = data.Sum(c => c.AverageAQI);

            foreach (var c in data)
            {
                string label = $"{c.CityCountry}\n{(c.AverageAQI / total * 100):0.0}%";
                pieSeries.Slices.Add(new PieSlice(label, c.AverageAQI));
            }

            model.Series.Add(pieSeries);
            return model;
        }

        private void cmbChartType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbChartType.SelectedIndex == 2)
            {
                textBoxCityCountry.PlaceholderText = "Країна";
            }
            else
            {
                textBoxCityCountry.PlaceholderText = "Місто, Країна";
            }
            LogMessage($"Вибрано тип графіка: {cmbChartType.SelectedIndex}");
        }

        private void buttonChartExport_Click(object sender, EventArgs e)
        {
            LogMessage("Користувач експортує графік");
            if (plotView1.Model == null)
            {
                MessageBox.Show("Немає графіка для експорту.");
                LogMessage("Помилка: спроба експорту без графіка");
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "PNG Image (*.png)|*.png";
                dialog.Title = "Зберегти графік як PNG";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = dialog.FileName;
                        plotView1.Model.Background = OxyColors.White;
                        var pngExporter = new PngExporter { Width = 800, Height = 600 };
                        pngExporter.ExportToFile(plotView1.Model, filePath);

                        MessageBox.Show($"Графік збережено: {filePath}");
                        LogMessage($"Графік експортовано у PNG: {filePath}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка при збереженні: " + ex.Message);
                        LogMessage("Помилка експорту графіка: " + ex.Message);
                    }
                }
                else
                {
                    LogMessage("Експорт графіка скасовано користувачем");
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

                        xlsxReportService.GenerateReport(
                            cityAirQualityManager.cityAirQualities,
                            filePath
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

                        var docxReportService = new DocxReportService();
                        docxReportService.GenerateReport(
                            cityAirQualityManager.cityAirQualities,
                            filePath
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
            string extension = Path.GetExtension(_path).ToLower().TrimStart('.');
            setManager(extension);
            _m.Write(_path, cityAirQualityManager.cityAirQualities);
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
    }
}

