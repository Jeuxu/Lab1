namespace Lab1._2
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            toolStrip1 = new ToolStrip();
            toolStripDropDownButton1 = new ToolStripDropDownButton();
            відкритиToolStripMenuItem = new ToolStripMenuItem();
            зберегтиЯкToolStripMenuItem = new ToolStripMenuItem();
            cSVToolStripMenuItem = new ToolStripMenuItem();
            jSONToolStripMenuItem = new ToolStripMenuItem();
            xMLToolStripMenuItem = new ToolStripMenuItem();
            xLSXToolStripMenuItem = new ToolStripMenuItem();
            звітиToolStripMenuItem = new ToolStripMenuItem();
            згенеруватиВToolStripMenuItem = new ToolStripMenuItem();
            згенеруватиDOCXзвітToolStripMenuItem = new ToolStripMenuItem();
            вихідToolStripMenuItem = new ToolStripMenuItem();
            tabPage4 = new TabPage();
            textBoxLogs = new TextBox();
            groupBox3 = new GroupBox();
            buttonLogExport = new Button();
            tabPage2 = new TabPage();
            plotView1 = new OxyPlot.WindowsForms.PlotView();
            groupBox2 = new GroupBox();
            label6 = new Label();
            comboBoxVolume = new ComboBox();
            buttonChartExport = new Button();
            textBoxCityCountry = new TextBox();
            label4 = new Label();
            buttonApplyChart = new Button();
            cmbChartType = new ComboBox();
            tabPage1 = new TabPage();
            dataGridViewData = new DataGridView();
            groupBox1 = new GroupBox();
            buttonAdd = new Button();
            buttonRemove = new Button();
            textBoxValue = new TextBox();
            label3 = new Label();
            label2 = new Label();
            label1 = new Label();
            comboBoxCondition = new ComboBox();
            comboBoxField = new ComboBox();
            buttonApply = new Button();
            tabControl1 = new TabControl();
            toolStrip1.SuspendLayout();
            tabPage4.SuspendLayout();
            groupBox3.SuspendLayout();
            tabPage2.SuspendLayout();
            groupBox2.SuspendLayout();
            tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewData).BeginInit();
            groupBox1.SuspendLayout();
            tabControl1.SuspendLayout();
            SuspendLayout();
            // 
            // toolStrip1
            // 
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripDropDownButton1 });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(800, 25);
            toolStrip1.TabIndex = 0;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripDropDownButton1
            // 
            toolStripDropDownButton1.DisplayStyle = ToolStripItemDisplayStyle.Text;
            toolStripDropDownButton1.DropDownItems.AddRange(new ToolStripItem[] { відкритиToolStripMenuItem, зберегтиЯкToolStripMenuItem, звітиToolStripMenuItem, вихідToolStripMenuItem });
            toolStripDropDownButton1.Image = (Image)resources.GetObject("toolStripDropDownButton1.Image");
            toolStripDropDownButton1.ImageTransparentColor = Color.Magenta;
            toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            toolStripDropDownButton1.Size = new Size(49, 22);
            toolStripDropDownButton1.Text = "Файл";
            // 
            // відкритиToolStripMenuItem
            // 
            відкритиToolStripMenuItem.Name = "відкритиToolStripMenuItem";
            відкритиToolStripMenuItem.Size = new Size(139, 22);
            відкритиToolStripMenuItem.Text = "Відкрити";
            відкритиToolStripMenuItem.Click += відкритиToolStripMenuItem_Click;
            // 
            // зберегтиЯкToolStripMenuItem
            // 
            зберегтиЯкToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { cSVToolStripMenuItem, jSONToolStripMenuItem, xMLToolStripMenuItem, xLSXToolStripMenuItem });
            зберегтиЯкToolStripMenuItem.Name = "зберегтиЯкToolStripMenuItem";
            зберегтиЯкToolStripMenuItem.Size = new Size(139, 22);
            зберегтиЯкToolStripMenuItem.Text = "Зберегти як";
            // 
            // cSVToolStripMenuItem
            // 
            cSVToolStripMenuItem.Name = "cSVToolStripMenuItem";
            cSVToolStripMenuItem.Size = new Size(102, 22);
            cSVToolStripMenuItem.Tag = "csv";
            cSVToolStripMenuItem.Text = "CSV";
            cSVToolStripMenuItem.Click += buttonToolStripMenuItem_Click;
            // 
            // jSONToolStripMenuItem
            // 
            jSONToolStripMenuItem.Name = "jSONToolStripMenuItem";
            jSONToolStripMenuItem.Size = new Size(102, 22);
            jSONToolStripMenuItem.Tag = "json";
            jSONToolStripMenuItem.Text = "JSON";
            jSONToolStripMenuItem.Click += buttonToolStripMenuItem_Click;
            // 
            // xMLToolStripMenuItem
            // 
            xMLToolStripMenuItem.Name = "xMLToolStripMenuItem";
            xMLToolStripMenuItem.Size = new Size(102, 22);
            xMLToolStripMenuItem.Tag = "xml";
            xMLToolStripMenuItem.Text = "XML";
            xMLToolStripMenuItem.Click += buttonToolStripMenuItem_Click;
            // 
            // xLSXToolStripMenuItem
            // 
            xLSXToolStripMenuItem.Name = "xLSXToolStripMenuItem";
            xLSXToolStripMenuItem.Size = new Size(102, 22);
            xLSXToolStripMenuItem.Tag = "xlsx";
            xLSXToolStripMenuItem.Text = "XLSX";
            xLSXToolStripMenuItem.Click += buttonToolStripMenuItem_Click;
            // 
            // звітиToolStripMenuItem
            // 
            звітиToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { згенеруватиВToolStripMenuItem, згенеруватиDOCXзвітToolStripMenuItem });
            звітиToolStripMenuItem.Name = "звітиToolStripMenuItem";
            звітиToolStripMenuItem.Size = new Size(139, 22);
            звітиToolStripMenuItem.Text = "Звіти";
            // 
            // згенеруватиВToolStripMenuItem
            // 
            згенеруватиВToolStripMenuItem.Name = "згенеруватиВToolStripMenuItem";
            згенеруватиВToolStripMenuItem.Size = new Size(201, 22);
            згенеруватиВToolStripMenuItem.Text = "Згенерувати XLSX-звіт";
            згенеруватиВToolStripMenuItem.Click += згенеруватиВToolStripMenuItem_Click;
            // 
            // згенеруватиDOCXзвітToolStripMenuItem
            // 
            згенеруватиDOCXзвітToolStripMenuItem.Name = "згенеруватиDOCXзвітToolStripMenuItem";
            згенеруватиDOCXзвітToolStripMenuItem.Size = new Size(201, 22);
            згенеруватиDOCXзвітToolStripMenuItem.Text = "Згенерувати DOCX-звіт";
            згенеруватиDOCXзвітToolStripMenuItem.Click += згенеруватиDOCXзвітToolStripMenuItem_Click;
            // 
            // вихідToolStripMenuItem
            // 
            вихідToolStripMenuItem.Name = "вихідToolStripMenuItem";
            вихідToolStripMenuItem.Size = new Size(139, 22);
            вихідToolStripMenuItem.Text = "Вихід";
            вихідToolStripMenuItem.Click += вихідToolStripMenuItem_Click;
            // 
            // tabPage4
            // 
            tabPage4.Controls.Add(textBoxLogs);
            tabPage4.Controls.Add(groupBox3);
            tabPage4.Location = new Point(4, 24);
            tabPage4.Name = "tabPage4";
            tabPage4.Padding = new Padding(3);
            tabPage4.Size = new Size(792, 397);
            tabPage4.TabIndex = 3;
            tabPage4.Text = "Логи";
            tabPage4.UseVisualStyleBackColor = true;
            // 
            // textBoxLogs
            // 
            textBoxLogs.Dock = DockStyle.Fill;
            textBoxLogs.Location = new Point(3, 3);
            textBoxLogs.Multiline = true;
            textBoxLogs.Name = "textBoxLogs";
            textBoxLogs.ReadOnly = true;
            textBoxLogs.ScrollBars = ScrollBars.Vertical;
            textBoxLogs.Size = new Size(583, 391);
            textBoxLogs.TabIndex = 2;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(buttonLogExport);
            groupBox3.Dock = DockStyle.Right;
            groupBox3.Location = new Point(586, 3);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(203, 391);
            groupBox3.TabIndex = 1;
            groupBox3.TabStop = false;
            // 
            // buttonLogExport
            // 
            buttonLogExport.Location = new Point(6, 171);
            buttonLogExport.Name = "buttonLogExport";
            buttonLogExport.Size = new Size(189, 47);
            buttonLogExport.TabIndex = 0;
            buttonLogExport.Text = "Експортувати";
            buttonLogExport.UseVisualStyleBackColor = true;
            buttonLogExport.Click += buttonLogExport_Click;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(plotView1);
            tabPage2.Controls.Add(groupBox2);
            tabPage2.Location = new Point(4, 24);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(792, 397);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "Графіки ";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // plotView1
            // 
            plotView1.Dock = DockStyle.Fill;
            plotView1.Location = new Point(3, 3);
            plotView1.Name = "plotView1";
            plotView1.PanCursor = Cursors.Hand;
            plotView1.Size = new Size(586, 391);
            plotView1.TabIndex = 4;
            plotView1.Text = "plotViewMain";
            plotView1.ZoomHorizontalCursor = Cursors.SizeWE;
            plotView1.ZoomRectangleCursor = Cursors.SizeNWSE;
            plotView1.ZoomVerticalCursor = Cursors.SizeNS;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(label6);
            groupBox2.Controls.Add(comboBoxVolume);
            groupBox2.Controls.Add(buttonChartExport);
            groupBox2.Controls.Add(textBoxCityCountry);
            groupBox2.Controls.Add(label4);
            groupBox2.Controls.Add(buttonApplyChart);
            groupBox2.Controls.Add(cmbChartType);
            groupBox2.Dock = DockStyle.Right;
            groupBox2.Location = new Point(589, 3);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(200, 391);
            groupBox2.TabIndex = 3;
            groupBox2.TabStop = false;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(6, 61);
            label6.Name = "label6";
            label6.Size = new Size(86, 15);
            label6.TabIndex = 11;
            label6.Text = "Об'єм графіка";
            // 
            // comboBoxVolume
            // 
            comboBoxVolume.Enabled = false;
            comboBoxVolume.FormattingEnabled = true;
            comboBoxVolume.Items.AddRange(new object[] { "Все", "Країна" });
            comboBoxVolume.Location = new Point(5, 79);
            comboBoxVolume.Name = "comboBoxVolume";
            comboBoxVolume.Size = new Size(189, 23);
            comboBoxVolume.TabIndex = 10;
            comboBoxVolume.SelectedIndexChanged += comboBoxVolume_SelectedIndexChanged;
            // 
            // buttonChartExport
            // 
            buttonChartExport.Location = new Point(5, 350);
            buttonChartExport.Name = "buttonChartExport";
            buttonChartExport.Size = new Size(189, 35);
            buttonChartExport.TabIndex = 9;
            buttonChartExport.Text = "Експорт у PNG";
            buttonChartExport.UseVisualStyleBackColor = true;
            buttonChartExport.Click += buttonChartExport_Click;
            // 
            // textBoxCityCountry
            // 
            textBoxCityCountry.Enabled = false;
            textBoxCityCountry.Location = new Point(6, 108);
            textBoxCityCountry.Name = "textBoxCityCountry";
            textBoxCityCountry.PlaceholderText = "Місто, Країна";
            textBoxCityCountry.Size = new Size(188, 23);
            textBoxCityCountry.TabIndex = 8;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(5, 17);
            label4.Name = "label4";
            label4.Size = new Size(73, 15);
            label4.TabIndex = 6;
            label4.Text = "Тип графіка";
            // 
            // buttonApplyChart
            // 
            buttonApplyChart.Location = new Point(5, 137);
            buttonApplyChart.Name = "buttonApplyChart";
            buttonApplyChart.Size = new Size(189, 35);
            buttonApplyChart.TabIndex = 5;
            buttonApplyChart.Text = "Обрати";
            buttonApplyChart.UseVisualStyleBackColor = true;
            buttonApplyChart.Click += buttonApplyChart_Click;
            // 
            // cmbChartType
            // 
            cmbChartType.FormattingEnabled = true;
            cmbChartType.Items.AddRange(new object[] { "Лінійний", "Стовпчиковий", "Круговий" });
            cmbChartType.Location = new Point(5, 35);
            cmbChartType.Name = "cmbChartType";
            cmbChartType.Size = new Size(189, 23);
            cmbChartType.TabIndex = 3;
            cmbChartType.SelectedIndexChanged += cmbChartType_SelectedIndexChanged;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(dataGridViewData);
            tabPage1.Controls.Add(groupBox1);
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(792, 397);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "Головна";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // dataGridViewData
            // 
            dataGridViewData.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewData.Dock = DockStyle.Fill;
            dataGridViewData.Location = new Point(3, 3);
            dataGridViewData.Name = "dataGridViewData";
            dataGridViewData.Size = new Size(547, 391);
            dataGridViewData.TabIndex = 2;
            dataGridViewData.CellValueChanged += editCity;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(buttonAdd);
            groupBox1.Controls.Add(buttonRemove);
            groupBox1.Controls.Add(textBoxValue);
            groupBox1.Controls.Add(label3);
            groupBox1.Controls.Add(label2);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(comboBoxCondition);
            groupBox1.Controls.Add(comboBoxField);
            groupBox1.Controls.Add(buttonApply);
            groupBox1.Dock = DockStyle.Right;
            groupBox1.Location = new Point(550, 3);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(239, 391);
            groupBox1.TabIndex = 1;
            groupBox1.TabStop = false;
            // 
            // buttonAdd
            // 
            buttonAdd.Location = new Point(6, 315);
            buttonAdd.Name = "buttonAdd";
            buttonAdd.Size = new Size(227, 32);
            buttonAdd.TabIndex = 10;
            buttonAdd.Text = "Додати";
            buttonAdd.UseVisualStyleBackColor = true;
            buttonAdd.Click += buttonAdd_Click;
            // 
            // buttonRemove
            // 
            buttonRemove.Location = new Point(6, 353);
            buttonRemove.Name = "buttonRemove";
            buttonRemove.Size = new Size(227, 32);
            buttonRemove.TabIndex = 9;
            buttonRemove.Text = "Видалити";
            buttonRemove.UseVisualStyleBackColor = true;
            buttonRemove.Click += buttonRemove_Click;
            // 
            // textBoxValue
            // 
            textBoxValue.Location = new Point(6, 119);
            textBoxValue.Name = "textBoxValue";
            textBoxValue.Size = new Size(227, 23);
            textBoxValue.TabIndex = 7;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(6, 101);
            label3.Name = "label3";
            label3.Size = new Size(60, 15);
            label3.TabIndex = 6;
            label3.Text = "Значення";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(6, 60);
            label2.Name = "label2";
            label2.Size = new Size(42, 15);
            label2.TabIndex = 5;
            label2.Text = "Умова";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(6, 16);
            label1.Name = "label1";
            label1.Size = new Size(36, 15);
            label1.TabIndex = 4;
            label1.Text = "Поле";
            // 
            // comboBoxCondition
            // 
            comboBoxCondition.Enabled = false;
            comboBoxCondition.FormattingEnabled = true;
            comboBoxCondition.Items.AddRange(new object[] { ">", "<", "=", ">=", "<=", "!=" });
            comboBoxCondition.Location = new Point(6, 75);
            comboBoxCondition.Name = "comboBoxCondition";
            comboBoxCondition.Size = new Size(227, 23);
            comboBoxCondition.TabIndex = 2;
            // 
            // comboBoxField
            // 
            comboBoxField.FormattingEnabled = true;
            comboBoxField.Items.AddRange(new object[] { "City", "Country", "Rank", "AverageAQI" });
            comboBoxField.Location = new Point(6, 34);
            comboBoxField.Name = "comboBoxField";
            comboBoxField.Size = new Size(227, 23);
            comboBoxField.TabIndex = 1;
            comboBoxField.SelectedIndexChanged += comboBoxField_SelectedIndexChanged;
            // 
            // buttonApply
            // 
            buttonApply.Location = new Point(6, 148);
            buttonApply.Name = "buttonApply";
            buttonApply.Size = new Size(227, 32);
            buttonApply.TabIndex = 0;
            buttonApply.Text = "Застосувати";
            buttonApply.UseVisualStyleBackColor = true;
            buttonApply.Click += buttonApply_Click;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPage4);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 25);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(800, 425);
            tabControl1.TabIndex = 1;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(tabControl1);
            Controls.Add(toolStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "MainForm";
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            tabPage4.ResumeLayout(false);
            tabPage4.PerformLayout();
            groupBox3.ResumeLayout(false);
            tabPage2.ResumeLayout(false);
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridViewData).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            tabControl1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ToolStrip toolStrip1;
        private ToolStripDropDownButton toolStripDropDownButton1;
        private ToolStripMenuItem відкритиToolStripMenuItem;
        private ToolStripMenuItem зберегтиЯкToolStripMenuItem;
        private ToolStripMenuItem cSVToolStripMenuItem;
        private ToolStripMenuItem jSONToolStripMenuItem;
        private ToolStripMenuItem xMLToolStripMenuItem;
        private ToolStripMenuItem xLSXToolStripMenuItem;
        private ToolStripMenuItem вихідToolStripMenuItem;
        private ToolStripMenuItem звітиToolStripMenuItem;
        private ToolStripMenuItem згенеруватиВToolStripMenuItem;
        private ToolStripMenuItem згенеруватиDOCXзвітToolStripMenuItem;
        private TabPage tabPage4;
        private TabPage tabPage2;
        private OxyPlot.WindowsForms.PlotView plotView1;
        private GroupBox groupBox2;
        private Button buttonChartExport;
        private TextBox textBoxCityCountry;
        private Label label4;
        private Button buttonApplyChart;
        private ComboBox cmbChartType;
        private TabPage tabPage1;
        private DataGridView dataGridViewData;
        private GroupBox groupBox1;
        private Button buttonRemove;
        private TextBox textBoxValue;
        private Label label3;
        private Label label2;
        private Label label1;
        private ComboBox comboBoxCondition;
        private ComboBox comboBoxField;
        private Button buttonApply;
        private TabControl tabControl1;
        private TextBox textBoxLogs;
        private GroupBox groupBox3;
        private Button buttonLogExport;
        private Label label6;
        private ComboBox comboBoxVolume;
        private Button buttonAdd;
    }
}
