using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace XAdd
{
    public partial class CurrencyForm : Form
    {
        HttpWebResponse response;
        public string dateSelected { get; set; }
        public CurrencyForm()
        {
            InitializeComponent();
        }

        private void CurrencyForm_Load(object sender, EventArgs e)
        {
            //HttpWebRequest request = (HttpWebRequest)WebRequest.Create(@"http://www.cbr.ru/");

            //try
            //{
            //    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            //}
            //catch (Exception)
            //{
                
            //}
            
            

        }

        private void CurrencyForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }


        private async void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            
            dataGridView1.Rows.Clear();
            dataGridView1.Rows.Add();
            if (monthCalendar1.SelectionStart > DateTime.Now)
            {
                monthCalendar1.SelectionStart = DateTime.Now.AddDays(1);
                monthCalendar1.SelectionEnd = monthCalendar1.SelectionStart;
            }
            label1.Text = "Курс на " + monthCalendar1.SelectionStart.ToString("dd'/'MM'/'yyyy");
            dateSelected = monthCalendar1.SelectionStart.ToString("dd'/'MM'/'yyyy");
            XmlDocument xDoc = new XmlDocument();

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(string.Format(@"http://www.cbr.ru/scripts/XML_daily.asp?date_req={0}", dateSelected));
            request.AllowAutoRedirect = true;
            request.MaximumAutomaticRedirections = 9999;
            try
            {
                response = (HttpWebResponse)await request.GetResponseAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"XAdd",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            

            using (Stream xmlStream = response.GetResponseStream())
            {
                xDoc.Load(xmlStream);
            }

            //try
            //{
            //    xDoc.Load(string.Format(@"http://www.cbr.ru/scripts/XML_daily.asp?date_req={0}", dateSelected));
            //}
            //catch (Exception ex)
            //{

            //    MessageBox.Show("Нет доступа к интернету или источник недоступен! "+ex.Message, "XAdd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            XmlElement xRoot = xDoc.DocumentElement;
            response.Close();

            foreach (XmlNode xnode in xRoot)
            {
                if (xnode.Attributes.Count > 0)
                {
                    XmlNode attr = xnode.Attributes.GetNamedItem("ID");

                    if (attr != null)
                    {
                        if (attr.Value == "R01235")
                        {
                            foreach (XmlNode childNode in xnode.ChildNodes)
                            {
                                if (childNode.Name == "Value")
                                {
                                    dataGridView1.Rows[0].Cells["USD"].Value = childNode.InnerText;
                                }
                            }
                        }
                        if (attr.Value == "R01239")
                        {
                            foreach (XmlNode childNode in xnode.ChildNodes)
                            {
                                if (childNode.Name == "Value")
                                {
                                    dataGridView1.Rows[0].Cells["EUR"].Value = childNode.InnerText;
                                }
                            }
                        }
                        if (attr.Value == "R01035")
                        {
                            foreach (XmlNode childNode in xnode.ChildNodes)
                            {
                                if (childNode.Name == "Value")
                                {
                                    dataGridView1.Rows[0].Cells["GBP"].Value = childNode.InnerText;
                                }
                            }
                        }


                    }
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.cbr.ru");
        }


    }
}
