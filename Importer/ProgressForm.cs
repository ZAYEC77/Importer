using Importer.Converters;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Importer
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();
            comboBox1.DataSource = App.Instance.Prices;
        }

        private void ProgressForm_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > -1)
            {
                var convertForSkyAuto = checkBox1.Checked;
                var convertForAvtoPro = checkBox2.Checked;

                if (!(convertForSkyAuto || convertForAvtoPro))
                {
                    MessageBox.Show("Оберіть тип прайсу");
                    return;
                }
                var price = App.Instance.Prices[comboBox1.SelectedIndex];

                var converter = Converter.CreateConverter(price, FileName);
                var data = converter.LoadData();
                var sd = new SaveFileDialog();
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    if (convertForSkyAuto)
                    {
                        converter.UseCoeficients = false;
                        converter.SaveCsv(data, sd.FileName);
                    }
                    if (convertForAvtoPro)
                    {
                        converter.UseCoeficients = true;
                        converter.SaveExcel(data, sd.FileName);
                    }
                    MessageBox.Show("Збережено");
                }
                else
                {
                    MessageBox.Show("Файл для збереження не вибрано");
                }
            }
        }
    }
}
