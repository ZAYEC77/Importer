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
                var price = App.Instance.Prices[comboBox1.SelectedIndex];

                var converter = Converter.CreateConverter(price, FileName);
                converter.UseCoeficients = checkBox1.Checked;
                converter.Convert();
                MessageBox.Show("Збережено");
            }
        }

        private void uploadFile(Converter converter)
        {
            if (MessageBox.Show("Завантажити прайс?", "", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    converter.Upload();
                    MessageBox.Show("Файл завантажено");
                }
                catch (Exception exeption)
                {
                    MessageBox.Show(String.Format("Помилка: {0}", exeption.Message));
                }
            }
        }
    }
}
