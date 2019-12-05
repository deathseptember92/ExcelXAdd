using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XAdd.Properties;

namespace XAdd
{
    public partial class AppendSheetsForm : Form
    {

        public AppendSheetsForm()
        {
            InitializeComponent();
        }


        private void AppendSheetsForm_Deactivate(object sender, EventArgs e)
        {
            Hide();
        }


        private void AppendSheetsCancel_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void AppendSheetsForm_Load(object sender, EventArgs e)
        {
            checkBox2.Checked = true;
        }

        private void AppendSheetsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Hide();
            e.Cancel=true;
        }
    }
}
