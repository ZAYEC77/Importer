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

        private void button1_Click(object sender, EventArgs e)
        {
            var priceSelectors = panel1.Controls.Cast<PriceSelector>();
            List<String[]> data = new List<String[]>();

            foreach (var priceSelector in priceSelectors)
            {
                if (priceSelector.HasFile)
                {
                    var config = priceSelector.Config;

                    var converter = Converter.CreateConverter(config, priceSelector.FilePath);
                    converter.LoadData();
                    converter.ProcessData();
                    data.AddRange(converter.Data);
                }

            }
            Converter.SaveResultFile(data);
            MessageBox.Show("OK");
        }
    }
}
