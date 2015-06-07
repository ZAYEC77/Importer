using Importer.Converters;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Importer
{
    public partial class MainForm : Form
    {
        SettingsForm settingsForm = new SettingsForm();

        public MainForm()
        {
            InitializeComponent();
            settingsForm.MdiParent = this;
        }

        private void відкритиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFile();
        }

        private void OpenFile()
        {
            var od = new OpenFileDialog();
            if (od.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var form = new ProgressForm() { MdiParent = this, FileName = od.FileName};
                form.Show();
            }
        }

        private void налаштуванняToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (settingsForm.MdiParent == null)
            {
                settingsForm = new SettingsForm();
                settingsForm.MdiParent = this;
            }
            settingsForm.Show();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            App.Save();
        }

        private void вихідToolStripMenuItem_Click(object sender, EventArgs e)
        {
            App.Save();
            Application.Exit();
        }
    }
}
