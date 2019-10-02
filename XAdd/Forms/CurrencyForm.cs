using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace XAdd
{
    public partial class CurrencyForm : Form
    {
        public DateTime dateSelected { get; set; }
        public CurrencyForm()
        {
            InitializeComponent();
        }

        private void CurrencyForm_Load(object sender, EventArgs e)
        {


        }

        private void CurrencyForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            dateSelected = monthCalendar1.SelectionStart;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s = dateSelected.ToString("dd'/'MM'/'yyyy");
            XmlReader myReader = XmlReader.Create($"http://www.cbr.ru/scripts/XML_daily.asp?date_req=29/09/2019");
        }
    }
}
