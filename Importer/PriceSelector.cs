using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Importer
{
    public partial class PriceSelector : UserControl
    {
        public string FilePath { get { return filePath; } }

        public Price Config { get { return config; } }

        public bool HasFile { get { return filePath != null; } }

        public string Label
        {
            get
            {
                return label.Text;
            }
            set
            {
                label.Text = value;
            }
        }

        protected string filePath = null;
        protected Price config = null;

        public PriceSelector(Price config)
        {
            InitializeComponent();
            this.label.Text = config.Title;
            this.config = config;
        }

        private void button_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var path = openFileDialog.FileName;
                if (File.Exists(path))
                {
                    fileLabel.Text = Path.GetFileName(path);
                    filePath = path;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog.Reset();
            fileLabel.Text = "Не вибрано";
            filePath = null;
        }
    }
}
