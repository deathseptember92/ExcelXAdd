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
    public partial class AppendWorkbooksForm : Form
    {
        public event Action AppendWorkbooksButtonClicked;

        public AppendWorkbooksForm()
        {
            InitializeComponent();
        }

        private void AppendWorkbooksForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Hide();
            
            e.Cancel = true;
        }

        private void AppendWorkbooksForm_Load(object sender, EventArgs e)
        {
           
        }

        private void buttonAppend_Click(object sender, EventArgs e)
        {
            AppendWorkbooksButtonClicked?.Invoke();
        }

        private void buttonFileDialog_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.InitialDirectory = "C:\\";
                ofd.Filter = "(Excel files *.xls,*.xlsx,*.xlsm|*.xls;*.xlsx;*.xlsm)";
                ofd.RestoreDirectory = true;
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    listView1.Items.Clear();
                    foreach (string file in ofd.FileNames)
                    {
                        listView1.Items.Add(file);
                    }
                }
            }
        }

        private void buttonExclude_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.SelectedItems)
            {
                item.Remove();
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.InitialDirectory = "C:\\";
                ofd.Filter = "(Excel files *.xls,*.xlsx,*.xlsm|*.xls;*.xlsx;*.xlsm)";
                ofd.RestoreDirectory = true;
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    foreach (string file in ofd.FileNames)
                    {
                        ListViewItem searchItem = listView1.FindItemWithText(file);
                        if (searchItem == null)
                        {
                            listView1.Items.Add(file);
                        }

                    }
                }
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            Hide();
        }
    }
}
