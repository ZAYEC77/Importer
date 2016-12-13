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
    public partial class CrossConverter : Form
    {
        public Cross.Converter converter = new Cross.Converter();

        public CrossConverter(string FileName)
        {
            InitializeComponent();
            converter.LoadFile(FileName);
        }

        private void CrossConverter_Load(object sender, EventArgs e)
        {
            var length = converter.SheetTables.Count;
            for (int i = 0; i < length; i++)
            {
                var tabPage = new TabPage("Лист " + (i + 1));
                var crossWidget = new Cross.CrossSheet(i, converter);
                crossWidget.Dock = DockStyle.Fill;
                tabPage.Controls.Add(crossWidget);
                tabControl.TabPages.Add(tabPage);
            }
        }
    }
}
