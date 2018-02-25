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

namespace Importer.Cross
{
    public partial class CrossSheet : UserControl
    {
        protected Config config;
        protected Cross.Converter converter;

        public CrossSheet(int index, Cross.Converter converter)
        {
            this.converter = converter;
            config = new Config() { SheetNumber = index };
            InitializeComponent();

            bindingSource.DataSource = config;
            this.numericUpDown1.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bindingSource, "CodeCol", true));
            this.numericUpDown2.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bindingSource, "BrandCol", true));
            this.numericUpDown4.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bindingSource, "DestCodeCol", true));
            this.numericUpDown5.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bindingSource, "DestBrandCol", true));
            dataGridView.DataSource = converter.SheetTables[index];
        }

        private void CrossSheet_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sd = new SaveFileDialog() { FileName = Path.ChangeExtension(converter.FileName, ".xls") };
            if (sd.ShowDialog() == DialogResult.OK)
            {
                converter.Convert(config, sd.FileName);
                MessageBox.Show("Готово");
            }
        }
    }
}
