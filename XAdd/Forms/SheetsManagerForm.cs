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

        private void TreeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            treeView1.SelectedNode = e.Node;
            SheetsManagerClickNode?.Invoke();
        }

        private void TreeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            treeView1.SelectedNode = e.Node;
            SheetsManagerDoubleClickNode?.Invoke();
        }
    }
}
