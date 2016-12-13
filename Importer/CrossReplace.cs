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
    public partial class CrossReplace : Form
    {
        public CrossReplace()
        {
            InitializeComponent();
        }

        private void CrossReplace_Load(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = App.Instance.CrossReplace;
            bindingNavigator1.BindingSource = bs;
            dataGridView1.DataSource = bs;
        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }
    }
}
