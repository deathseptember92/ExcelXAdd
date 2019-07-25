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
        public event Action SelectedNodesToFinalList;
        public event Action RemoveNodesFromFinalList;
        public event Action AppendSheetsClicked;
        public AppendSheetsForm()
        {
            InitializeComponent();
        }

        private void SelectedNodesToFinal_Click(object sender, EventArgs e)
        {
            SelectedNodesToFinalList?.Invoke();
        }

        private void AppendSheetsForm_Deactivate(object sender, EventArgs e)
        {
            Hide();
        }

        private void RemoveNodesFromFinal_Click(object sender, EventArgs e)
        {
            RemoveNodesFromFinalList?.Invoke();
        }

        private void AppendSheetsCancel_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void AppendSheetsOK_Click(object sender, EventArgs e)
        {
            AppendSheetsClicked?.Invoke();
        }

        private void TreeView1_DoubleClick(object sender, EventArgs e)
        {
            SelectedNodesToFinalList?.Invoke();
        }

        private void TreeView2_DoubleClick(object sender, EventArgs e)
        {
            RemoveNodesFromFinalList?.Invoke();
        }

        private void AppendSheetsForm_Load(object sender, EventArgs e)
        {
         
        }
    }
}
