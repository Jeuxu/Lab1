namespace Lab1
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
            txtFolderA = new TextBox();
            txtLog = new TextBox();
            btnSync = new Button();
            txtFolderB = new TextBox();
            btnSelectA = new Button();
            btnSelectB = new Button();
            cmbDirection = new ComboBox();
            SuspendLayout();
            // 
            // txtFolderA
            // 
            txtFolderA.Enabled = false;
            txtFolderA.Location = new Point(376, 12);
            txtFolderA.Name = "txtFolderA";
            txtFolderA.PlaceholderText = "Шлях до папки 1";
            txtFolderA.Size = new Size(261, 23);
            txtFolderA.TabIndex = 6;
            // 
            // txtLog
            // 
            txtLog.Dock = DockStyle.Left;
            txtLog.Enabled = false;
            txtLog.Location = new Point(0, 0);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Size = new Size(370, 511);
            txtLog.TabIndex = 3;
            // 
            // btnSync
            // 
            btnSync.Font = new Font("Segoe UI", 11F);
            btnSync.Location = new Point(376, 448);
            btnSync.Name = "btnSync";
            btnSync.Size = new Size(261, 51);
            btnSync.TabIndex = 5;
            btnSync.Text = "Cинхронізація";
            btnSync.UseVisualStyleBackColor = true;
            btnSync.Click += btnSync_Click;
            // 
            // txtFolderB
            // 
            txtFolderB.Enabled = false;
            txtFolderB.Location = new Point(376, 98);
            txtFolderB.Name = "txtFolderB";
            txtFolderB.PlaceholderText = "Шлях до папки 2";
            txtFolderB.Size = new Size(261, 23);
            txtFolderB.TabIndex = 7;
            // 
            // btnSelectA
            // 
            btnSelectA.Font = new Font("Segoe UI", 11F);
            btnSelectA.Location = new Point(376, 41);
            btnSelectA.Name = "btnSelectA";
            btnSelectA.Size = new Size(261, 51);
            btnSelectA.TabIndex = 8;
            btnSelectA.Text = "Вибрати папку 1";
            btnSelectA.UseVisualStyleBackColor = true;
            btnSelectA.Click += btnSelectA_Click;
            // 
            // btnSelectB
            // 
            btnSelectB.Font = new Font("Segoe UI", 11F);
            btnSelectB.Location = new Point(376, 127);
            btnSelectB.Name = "btnSelectB";
            btnSelectB.Size = new Size(261, 51);
            btnSelectB.TabIndex = 9;
            btnSelectB.Text = "Вибрати папку 2";
            btnSelectB.UseVisualStyleBackColor = true;
            btnSelectB.Click += btnSelectB_Click;
            // 
            // cmbDirection
            // 
            cmbDirection.FormattingEnabled = true;
            cmbDirection.Items.AddRange(new object[] { "Ліва -> Права", "Права -> Ліва", "Двостороння" });
            cmbDirection.Location = new Point(376, 419);
            cmbDirection.Name = "cmbDirection";
            cmbDirection.Size = new Size(261, 23);
            cmbDirection.TabIndex = 10;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(648, 511);
            Controls.Add(cmbDirection);
            Controls.Add(btnSelectB);
            Controls.Add(btnSelectA);
            Controls.Add(txtFolderB);
            Controls.Add(txtFolderA);
            Controls.Add(btnSync);
            Controls.Add(txtLog);
            Name = "MainForm";
            Text = "Синхронізація папок";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private TextBox txtLog;
        private Button btnSync;
        private TextBox txtFolderA;
        private TextBox txtFolderB;
        private Button btnSelectA;
        private Button btnSelectB;
        private ComboBox cmbDirection;
    }
}
