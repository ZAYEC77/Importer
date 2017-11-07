using Importer.Converters;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Importer
{
    public partial class MultiPriceForm : Form
    {
        public MultiPriceForm()
        {
            InitializeComponent();
        }

        private void MultiPriceForm_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            var prices = App.Instance.Prices;
            for (int i = 0; i < prices.Count; i++)
            {
                var price = prices[i];
                var control = new PriceSelector(price);
                control.Top = i * control.Height;
                panel1.Controls.Add(control);
            }
        }

        private void MultipleConverterLog(string line)
        {
            logBox.AppendText(String.Format("{0}: {1}\n", DateTime.Now.ToString(), line));
        }


        private void button1_Click(object sender, EventArgs e)
        {
            logBox.Clear();
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();

            if (folderDialog.ShowDialog() == DialogResult.OK && (!string.IsNullOrWhiteSpace(folderDialog.SelectedPath)))
            {
                var convertForSkyAuto = checkBox1.Checked;
                var convertForAvtoPro = checkBox2.Checked;

                var path = folderDialog.SelectedPath;
                var sitePath = Path.Combine(path, "sky_auto");
                Directory.CreateDirectory(sitePath);
                var avtoProPath = Path.Combine(path, "avto_pro");
                Directory.CreateDirectory(avtoProPath);

                var priceSelectors = panel1.Controls.Cast<PriceSelector>();

                MultipleConverterLog("Початок конвертації");
                MultipleConverterLog(String.Format("Вихідна папка: {0}", path));

                foreach (var priceSelector in priceSelectors)
                {
                    if (priceSelector.HasFile)
                    {
                        var config = priceSelector.Config;

                        var converter = Converter.CreateConverter(config, priceSelector.FilePath);
                        if (convertForSkyAuto)
                        {
                            try
                            {
                                MultipleConverterLog(String.Format("Початок конвертації прайсу {0} в CSV", config.Title));
                                converter.UseCoeficients = false;
                                converter.SaveCsv(converter.LoadData(), sitePath);
                                MultipleConverterLog(String.Format("Завершення конвертації прайсу {0} в CSV", config.Title));
                            }
                            catch (Exception exception)
                            {
                                MultipleConverterLog(String.Format("Помилка при обробці прайсу {0} в CSV", config.Title));
                                MultipleConverterLog(exception.Message);
                                MultipleConverterLog("=========");
                            }
                        }
                        if (convertForAvtoPro)
                        {
                            try
                            {
                                MultipleConverterLog(String.Format("Початок конвертації прайсу {0} в Excel", config.Title));
                                converter.UseCoeficients = true;
                                converter.SaveExcel(converter.LoadData(), avtoProPath);
                                MultipleConverterLog(String.Format("Завершення конвертації прайсу {0} в Excel", config.Title));
                            }
                            catch (Exception exception)
                            {
                                MultipleConverterLog(String.Format("Помилка при конвертації прайсу {0} в Excel", config.Title));
                                MultipleConverterLog(exception.Message);
                                MultipleConverterLog("=========");
                            }
                        }
                    }
                }

                MultipleConverterLog("Конвертація завершена");
                MessageBox.Show("Конвертація завершена");
            }
        }
    }
}
