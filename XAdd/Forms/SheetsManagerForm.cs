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
        public event Action SheetsManagerOpenClicked;
        public event Action SheetsManagerRenameClicked;
        public event Action SheetsManagerRemoveClicked;

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

        private void SheetsManagerForm_Activated(object sender, EventArgs e)
        {
            pictureBox1.Image = null;

        }

        private void Panel1_MouseEnter(object sender, EventArgs e)
        {
            panel1.Focus();
        }

        private void PictureBox1_MouseEnter(object sender, EventArgs e)
        {
            panel1.Focus();
        }

        private void SheetsManagerForm_Deactivate(object sender, EventArgs e)
        {

        }

        private void OpenButton_Click(object sender, EventArgs e)
        {
            SheetsManagerOpenClicked?.Invoke();
        }

        private void RenameButton_Click(object sender, EventArgs e)
        {
            SheetsManagerRenameClicked?.Invoke();
        }

        private void RemoveButton_Click(object sender, EventArgs e)
        {
            SheetsManagerRemoveClicked?.Invoke();
        }
    }
}
