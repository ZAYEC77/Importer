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
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
            bindingSource.DataSource = App.Instance;
            bindingSourceApp.DataSource = App.Instance;
            bindingSource.AddingNew += bindingSource_AddingNew;
            comboBox1.DataSource = Enum.GetValues(typeof(PriceFormat));
        }

        void bindingSource_AddingNew(object sender, AddingNewEventArgs e)
        {
            e.NewObject = new Price() { Title = "Новий прайс" };
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //this.textBox10.Enabled = !this.checkBox1.Checked;
        }

    }
}
