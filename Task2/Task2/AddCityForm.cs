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
    public partial class AddCityForm : Form
    {
        public AddCityForm()
        {
            InitializeComponent();

        }


        public string CityName
        {
            get { return textBoxCityName.Text; }
        }

        public List<MonthlyAQI> MonthlyData
        {
            get
            {
                return new List<MonthlyAQI>
                {
                    new MonthlyAQI { Month = "Jan", Value = (int)numericUpDownJan.Value },
                    new MonthlyAQI { Month = "Feb", Value = (int)numericUpDownFeb.Value },
                    new MonthlyAQI { Month = "Mar", Value = (int)numericUpDownMar.Value },
                    new MonthlyAQI { Month = "Apr", Value = (int)numericUpDownApr.Value },
                    new MonthlyAQI { Month = "May", Value = (int)numericUpDownMay.Value },
                    new MonthlyAQI { Month = "Jun", Value = (int)numericUpDownJun.Value },
                    new MonthlyAQI { Month = "Jul", Value = (int)numericUpDownJul.Value },
                    new MonthlyAQI { Month = "Aug", Value = (int)numericUpDownAug.Value },
                    new MonthlyAQI { Month = "Sep", Value = (int)numericUpDownSep.Value },
                    new MonthlyAQI { Month = "Oct", Value = (int)numericUpDownOct.Value },
                    new MonthlyAQI { Month = "Nov", Value = (int)numericUpDownNov.Value },
                    new MonthlyAQI { Month = "Dec", Value = (int)numericUpDownDec.Value }
                };
            }
        }

        public int AverageAQI
        {
            get
            {
                if (MonthlyData.Count == 0) return 0;
                return (int)Math.Round(MonthlyData.Average(m => m.Value));
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CityName))
            {
                MessageBox.Show("Будь ласка, введіть назву міста та країни.");
                return;
            }

            var parts = CityName.Split(',');

            if (parts.Length != 2 ||
                string.IsNullOrWhiteSpace(parts[0]) ||
                string.IsNullOrWhiteSpace(parts[1]))
            {
                MessageBox.Show("Будь ласка, введіть назву у форматі: 'Місто, Країна'.");
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
