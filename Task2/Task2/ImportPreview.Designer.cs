namespace UI
{
    partial class ImportPreview
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImportPreview));
            textBox1 = new TextBox();
            buttonLoad = new Button();
            buttonCancel = new Button();
            SuspendLayout();
            // 
            // textBox1
            // 
            textBox1.Enabled = false;
            textBox1.Location = new Point(12, 12);
            textBox1.Multiline = true;
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(380, 393);
            textBox1.TabIndex = 0;
            // 
            // buttonLoad
            // 
            buttonLoad.DialogResult = DialogResult.OK;
            buttonLoad.Location = new Point(12, 415);
            buttonLoad.Name = "buttonLoad";
            buttonLoad.Size = new Size(189, 23);
            buttonLoad.TabIndex = 1;
            buttonLoad.Text = "Завантажити";
            buttonLoad.UseVisualStyleBackColor = true;
            // 
            // buttonCancel
            // 
            buttonCancel.DialogResult = DialogResult.Cancel;
            buttonCancel.Location = new Point(207, 415);
            buttonCancel.Name = "buttonCancel";
            buttonCancel.Size = new Size(189, 23);
            buttonCancel.TabIndex = 2;
            buttonCancel.Text = "Відхилити";
            buttonCancel.UseVisualStyleBackColor = true;
            // 
            // ImportPreview
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(404, 450);
            Controls.Add(buttonCancel);
            Controls.Add(buttonLoad);
            Controls.Add(textBox1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "ImportPreview";
            Text = "Попередній перегляд";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox textBox1;
        private Button buttonLoad;
        private Button buttonCancel;
    }
}