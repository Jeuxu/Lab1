namespace Lab1
{
    public partial class MainForm : Form
    {
        private string _folderA;
        private string _folderB;

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnSelectA_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    _folderA = fbd.SelectedPath;
                    txtFolderA.Text = _folderA;
                }
            }
        }

        private void btnSelectB_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    _folderB = fbd.SelectedPath;
                    txtFolderB.Text = _folderB;
                }
            }
        }

        private void btnSync_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_folderA) || string.IsNullOrEmpty(_folderB))
            {
                MessageBox.Show("Оберіть обидві папки!");
                return;
            }

            var sync = new FolderSyncService();
            int direction = cmbDirection.SelectedIndex;

            sync.LogGenerated += (msg) => txtLog.AppendText(msg + Environment.NewLine);

            if (direction == 0)
                sync.Sync(_folderA, _folderB);
            else if (direction == 1)
                sync.Sync(_folderB, _folderA);
            else if (direction == 2)
            {
                sync.Sync(_folderA, _folderB);
                sync.Sync(_folderB, _folderA);
            }
            else
            {
                MessageBox.Show("Виберіть тип синхронізації!");
            }
        }
    }
}
