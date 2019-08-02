using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XAdd
{
    public partial class SheetsManagerForm : Form
    {

        public event Action SheetsManagerClickNode;
        public event Action SheetsManagerDoubleClickNode;

        public SheetsManagerForm()
        {
            InitializeComponent();
        }

        private void SheetsManagerForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Hide();
            e.Cancel = true;
        }

        private void TreeView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            SheetsManagerDoubleClickNode?.Invoke();
        }

        private void TreeView1_Click(object sender, EventArgs e)
        {
            SheetsManagerClickNode?.Invoke();
        }
    }
}
